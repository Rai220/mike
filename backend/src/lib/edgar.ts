import { convert } from "html-to-text";

// SEC EDGAR is a free public API. It requires a descriptive User-Agent that
// identifies the calling application and asks clients to stay well under
// 10 requests/second (https://www.sec.gov/os/accessing-edgar-data).
const EDGAR_FTS_BASE = "https://efts.sec.gov/LATEST/search-index";
const EDGAR_DATA_BASE = "https://data.sec.gov";
const EDGAR_WWW_BASE = "https://www.sec.gov";

const REQUEST_GAP_MS = 150;
const MAX_FILING_BYTES = 30 * 1024 * 1024;
const MAX_FILING_TEXT_CHARS = 2_000_000;
const TICKER_CACHE_TTL_MS = 24 * 60 * 60 * 1000;

type JsonRecord = Record<string, unknown>;

export function edgarUserAgent(): string {
    // SEC rejects requests whose User-Agent lacks a contact in the
    // "Company Name email@domain" shape ("Undeclared Automated Tool" 403),
    // so the fallback must carry an email-format placeholder. Operators
    // should still set a real contact via EDGAR_USER_AGENT.
    return (
        process.env.EDGAR_USER_AGENT?.trim() ||
        "MikeOSS self-hosted mike-admin@example.com"
    );
}

export function padCik(cik: number | string): string {
    const digits = String(cik).replace(/\D/g, "");
    return digits.padStart(10, "0");
}

export function accessionWithDashes(accession: string): string {
    const digits = accession.replace(/\D/g, "");
    if (digits.length !== 18) return accession.trim();
    return `${digits.slice(0, 10)}-${digits.slice(10, 12)}-${digits.slice(12)}`;
}

export function accessionNoDashes(accession: string): string {
    return accession.replace(/\D/g, "");
}

export function edgarFilingIndexUrl(
    cik: number | string,
    accession: string,
): string {
    const cikDigits = String(Number.parseInt(padCik(cik), 10));
    return `${EDGAR_WWW_BASE}/Archives/edgar/data/${cikDigits}/${accessionNoDashes(accession)}`;
}

export function edgarFilingDocumentUrl(
    cik: number | string,
    accession: string,
    document: string,
): string {
    return `${edgarFilingIndexUrl(cik, accession)}/${document}`;
}

// Serialize requests with a small gap so a burst of tool calls cannot push
// the deployment over the SEC's fair-access rate threshold.
let throttleTail: Promise<void> = Promise.resolve();
function throttled<T>(run: () => Promise<T>): Promise<T> {
    const result = throttleTail.then(run);
    throttleTail = result.then(
        () => new Promise((resolve) => setTimeout(resolve, REQUEST_GAP_MS)),
        () => new Promise((resolve) => setTimeout(resolve, REQUEST_GAP_MS)),
    );
    return result;
}

function edgarError(status: number, context: string): Error {
    if (status === 429 || status === 403) {
        return new Error(
            `SEC EDGAR rate limit or access threshold reached (${status}) while ${context}. Stop EDGAR calls for this turn.`,
        );
    }
    return new Error(`SEC EDGAR error (${status}) while ${context}.`);
}

async function edgarFetch(url: string, context: string): Promise<Response> {
    return throttled(async () => {
        const response = await fetch(url, {
            headers: {
                "User-Agent": edgarUserAgent(),
                Accept: "application/json, text/html;q=0.9, */*;q=0.8",
            },
        });
        if (!response.ok) {
            response.body?.cancel().catch(() => {});
            throw edgarError(response.status, context);
        }
        return response;
    });
}

async function edgarFetchJson<T>(url: string, context: string): Promise<T> {
    const response = await edgarFetch(url, context);
    return (await response.json()) as T;
}

function recordField(value: unknown): JsonRecord | null {
    return value && typeof value === "object" && !Array.isArray(value)
        ? (value as JsonRecord)
        : null;
}

function stringItem(value: unknown): string | null {
    return typeof value === "string" && value.trim() ? value.trim() : null;
}

// --- Company lookup -------------------------------------------------------

export type EdgarCompany = {
    cik: string;
    ticker: string | null;
    name: string;
};

let tickerCache: { loadedAt: number; companies: EdgarCompany[] } | null = null;

export function parseCompanyTickers(payload: unknown): EdgarCompany[] {
    const record = recordField(payload);
    if (!record) return [];
    const companies: EdgarCompany[] = [];
    for (const value of Object.values(record)) {
        const row = recordField(value);
        if (!row) continue;
        const cik =
            typeof row.cik_str === "number" || typeof row.cik_str === "string"
                ? padCik(row.cik_str)
                : null;
        const name = stringItem(row.title);
        if (!cik || !name) continue;
        companies.push({ cik, ticker: stringItem(row.ticker), name });
    }
    return companies;
}

export function matchEdgarCompanies(
    companies: EdgarCompany[],
    query: string,
    limit: number,
): EdgarCompany[] {
    const trimmed = query.trim();
    if (!trimmed) return [];
    const upper = trimmed.toUpperCase();
    const digits = trimmed.replace(/\D/g, "");
    const byCik =
        digits && digits === trimmed.replace(/\s/g, "")
            ? companies.filter((company) => company.cik === padCik(digits))
            : [];
    const byTicker = companies.filter((company) => company.ticker === upper);
    const lower = trimmed.toLowerCase();
    const byName = companies.filter((company) =>
        company.name.toLowerCase().includes(lower),
    );
    const seen = new Set<string>();
    const merged: EdgarCompany[] = [];
    for (const company of [...byCik, ...byTicker, ...byName]) {
        const key = `${company.cik}:${company.ticker ?? ""}`;
        if (seen.has(key)) continue;
        seen.add(key);
        merged.push(company);
        if (merged.length >= limit) break;
    }
    return merged;
}

export async function findEdgarCompanies(
    query: string,
    limit = 8,
): Promise<EdgarCompany[]> {
    if (!tickerCache || Date.now() - tickerCache.loadedAt > TICKER_CACHE_TTL_MS) {
        const payload = await edgarFetchJson<unknown>(
            `${EDGAR_WWW_BASE}/files/company_tickers.json`,
            "loading the EDGAR company/ticker directory",
        );
        tickerCache = {
            loadedAt: Date.now(),
            companies: parseCompanyTickers(payload),
        };
    }
    return matchEdgarCompanies(tickerCache.companies, query, limit);
}

// --- Filing history (submissions) -----------------------------------------

export type EdgarFilingRef = {
    accession_number: string;
    form: string;
    filed: string | null;
    report_date: string | null;
    primary_document: string | null;
    primary_doc_description: string | null;
    url: string | null;
};

export function parseSubmissions(
    payload: unknown,
    cik: string,
): { name: string | null; filings: EdgarFilingRef[] } {
    const record = recordField(payload);
    const recent = recordField(recordField(record?.filings)?.recent);
    const name = stringItem(record?.name);
    if (!recent) return { name, filings: [] };
    const accession = Array.isArray(recent.accessionNumber)
        ? recent.accessionNumber
        : [];
    const form = Array.isArray(recent.form) ? recent.form : [];
    const filingDate = Array.isArray(recent.filingDate) ? recent.filingDate : [];
    const reportDate = Array.isArray(recent.reportDate) ? recent.reportDate : [];
    const primaryDocument = Array.isArray(recent.primaryDocument)
        ? recent.primaryDocument
        : [];
    const primaryDocDescription = Array.isArray(recent.primaryDocDescription)
        ? recent.primaryDocDescription
        : [];
    const filings: EdgarFilingRef[] = [];
    for (let i = 0; i < accession.length; i++) {
        const accessionNumber = stringItem(accession[i]);
        const formType = stringItem(form[i]);
        if (!accessionNumber || !formType) continue;
        const primaryDoc = stringItem(primaryDocument[i]);
        filings.push({
            accession_number: accessionNumber,
            form: formType,
            filed: stringItem(filingDate[i]),
            report_date: stringItem(reportDate[i]),
            primary_document: primaryDoc,
            primary_doc_description: stringItem(primaryDocDescription[i]),
            url: primaryDoc
                ? edgarFilingDocumentUrl(cik, accessionNumber, primaryDoc)
                : null,
        });
    }
    return { name, filings };
}

export async function getEdgarFilings(args: {
    cik: number | string;
    forms?: string[];
    filedAfter?: string;
    filedBefore?: string;
    limit?: number;
}): Promise<{ cik: string; name: string | null; filings: EdgarFilingRef[] }> {
    const cik = padCik(args.cik);
    const payload = await edgarFetchJson<unknown>(
        `${EDGAR_DATA_BASE}/submissions/CIK${cik}.json`,
        `loading the filing history for CIK ${cik}`,
    );
    const { name, filings } = parseSubmissions(payload, cik);
    const wantedForms = (args.forms ?? [])
        .map((form) => form.trim().toUpperCase())
        .filter(Boolean);
    const filtered = filings.filter((filing) => {
        if (
            wantedForms.length &&
            !wantedForms.includes(filing.form.toUpperCase())
        ) {
            return false;
        }
        if (args.filedAfter && filing.filed && filing.filed < args.filedAfter) {
            return false;
        }
        if (args.filedBefore && filing.filed && filing.filed > args.filedBefore) {
            return false;
        }
        return true;
    });
    return {
        cik,
        name,
        filings: filtered.slice(0, Math.max(1, Math.min(args.limit ?? 20, 50))),
    };
}

// --- Full-text search ------------------------------------------------------

export type EdgarSearchHit = {
    accession_number: string;
    document: string | null;
    form: string | null;
    filed: string | null;
    companies: { cik: string | null; name: string | null }[];
    file_description: string | null;
    url: string | null;
};

export function parseFullTextSearch(payload: unknown): {
    total: number;
    hits: EdgarSearchHit[];
} {
    const hitsRecord = recordField(recordField(payload)?.hits);
    const totalRecord = recordField(hitsRecord?.total);
    const total =
        typeof totalRecord?.value === "number" ? totalRecord.value : 0;
    const rows = Array.isArray(hitsRecord?.hits) ? hitsRecord.hits : [];
    const hits: EdgarSearchHit[] = [];
    for (const row of rows) {
        const record = recordField(row);
        const source = recordField(record?._source);
        const id = stringItem(record?._id);
        const adsh = stringItem(source?.adsh) ?? id?.split(":")[0] ?? null;
        if (!adsh) continue;
        const document = id?.includes(":")
            ? (stringItem(id.slice(id.indexOf(":") + 1)) ?? null)
            : null;
        const ciks = Array.isArray(source?.ciks) ? source.ciks : [];
        const names = Array.isArray(source?.display_names)
            ? source.display_names
            : [];
        const companies = ciks.length
            ? ciks.map((cikValue, index) => ({
                  cik: stringItem(cikValue) ? padCik(String(cikValue)) : null,
                  name:
                      stringItem(names[index])?.replace(
                          /\s*\(CIK \d+\)\s*$/,
                          "",
                      ) ?? null,
              }))
            : [];
        const firstCik = companies.find((company) => company.cik)?.cik ?? null;
        hits.push({
            accession_number: accessionWithDashes(adsh),
            document,
            form: stringItem(source?.form) ?? stringItem(source?.file_type),
            filed: stringItem(source?.file_date),
            companies,
            file_description: stringItem(source?.file_description),
            url:
                firstCik && document
                    ? edgarFilingDocumentUrl(firstCik, adsh, document)
                    : null,
        });
    }
    return { total, hits };
}

export async function searchEdgarFilings(args: {
    query: string;
    forms?: string[];
    cik?: number | string;
    filedAfter?: string;
    filedBefore?: string;
    limit?: number;
}): Promise<{ total: number; hits: EdgarSearchHit[] }> {
    const params = new URLSearchParams();
    params.set("q", args.query);
    const forms = (args.forms ?? [])
        .map((form) => form.trim().toUpperCase())
        .filter(Boolean);
    if (forms.length) params.set("forms", forms.join(","));
    if (args.cik !== undefined && args.cik !== null && String(args.cik).trim()) {
        params.set("ciks", padCik(args.cik));
    }
    if (args.filedAfter || args.filedBefore) {
        params.set("dateRange", "custom");
        if (args.filedAfter) params.set("startdt", args.filedAfter);
        if (args.filedBefore) params.set("enddt", args.filedBefore);
    }
    const payload = await edgarFetchJson<unknown>(
        `${EDGAR_FTS_BASE}?${params.toString()}`,
        "running EDGAR full-text search",
    );
    const { total, hits } = parseFullTextSearch(payload);
    return { total, hits: hits.slice(0, Math.max(1, Math.min(args.limit ?? 10, 25))) };
}

// --- Filing documents -------------------------------------------------------

function htmlToPlainText(html: string): string {
    return convert(html, {
        wordwrap: false,
        selectors: [
            { selector: "a", options: { ignoreHref: true } },
            { selector: "img", format: "skip" },
        ],
    })
        .replace(/\n{3,}/g, "\n\n")
        .trim();
}

async function resolvePrimaryDocument(
    cik: string,
    accession: string,
): Promise<string> {
    const payload = await edgarFetchJson<unknown>(
        `${edgarFilingIndexUrl(cik, accession)}/index.json`,
        `listing documents for filing ${accessionWithDashes(accession)}`,
    );
    const directory = recordField(recordField(payload)?.directory);
    const items = Array.isArray(directory?.item) ? directory.item : [];
    const names = items
        .map((item) => stringItem(recordField(item)?.name))
        .filter((name): name is string => !!name);
    const primary =
        names.find(
            (name) => /\.htm(l)?$/i.test(name) && !/^r\d+\.htm/i.test(name),
        ) ?? names.find((name) => /\.txt$/i.test(name));
    if (!primary) {
        throw new Error(
            `No readable HTML/text document found in filing ${accessionWithDashes(accession)}.`,
        );
    }
    return primary;
}

export type EdgarFilingText = {
    cik: string;
    accession_number: string;
    document: string;
    url: string;
    text: string;
    truncated: boolean;
};

export async function fetchEdgarFilingText(args: {
    cik: number | string;
    accessionNumber: string;
    document?: string | null;
}): Promise<EdgarFilingText> {
    const cik = padCik(args.cik);
    const accession = accessionWithDashes(args.accessionNumber);
    let document = args.document?.trim() || null;
    if (!document || /\.pdf$/i.test(document)) {
        document = await resolvePrimaryDocument(cik, accession);
    }
    const url = edgarFilingDocumentUrl(cik, accession, document);
    const response = await edgarFetch(
        url,
        `downloading filing document ${document}`,
    );
    const declaredLength = Number(
        response.headers.get("content-length") ?? "0",
    );
    if (declaredLength > MAX_FILING_BYTES) {
        response.body?.cancel().catch(() => {});
        throw new Error(
            `Filing document ${document} is too large to read (${Math.round(declaredLength / 1024 / 1024)} MB).`,
        );
    }
    const raw = await response.text();
    if (raw.length > MAX_FILING_BYTES) {
        throw new Error(
            `Filing document ${document} is too large to read.`,
        );
    }
    const text = /\.htm(l)?$/i.test(document) ? htmlToPlainText(raw) : raw.trim();
    const truncated = text.length > MAX_FILING_TEXT_CHARS;
    return {
        cik,
        accession_number: accession,
        document,
        url,
        text: truncated ? text.slice(0, MAX_FILING_TEXT_CHARS) : text,
        truncated,
    };
}

// --- XBRL company facts -----------------------------------------------------

const DEFAULT_FACT_CONCEPTS = [
    "Revenues",
    "RevenueFromContractWithCustomerExcludingAssessedTax",
    "NetIncomeLoss",
    "Assets",
    "Liabilities",
    "StockholdersEquity",
    "CashAndCashEquivalentsAtCarryingValue",
    "EarningsPerShareDiluted",
];

export type EdgarFactPoint = {
    end: string | null;
    value: number;
    fy: number | null;
    fp: string | null;
    form: string | null;
};

export type EdgarFactsSummary = {
    cik: string;
    entity_name: string | null;
    facts: {
        concept: string;
        label: string | null;
        unit: string;
        points: EdgarFactPoint[];
    }[];
    missing_concepts: string[];
};

export function summarizeCompanyFacts(
    payload: unknown,
    cik: string,
    concepts: string[],
    pointsPerConcept = 6,
): EdgarFactsSummary {
    const record = recordField(payload);
    const entityName = stringItem(record?.entityName);
    const gaap = recordField(recordField(record?.facts)?.["us-gaap"]);
    const facts: EdgarFactsSummary["facts"] = [];
    const missing: string[] = [];
    for (const concept of concepts) {
        const conceptRecord = recordField(gaap?.[concept]);
        const units = recordField(conceptRecord?.units);
        if (!units) {
            missing.push(concept);
            continue;
        }
        const [unit, rawPoints] = Object.entries(units)[0] ?? [null, null];
        if (!unit || !Array.isArray(rawPoints)) {
            missing.push(concept);
            continue;
        }
        const points = rawPoints
            .map((point): EdgarFactPoint | null => {
                const row = recordField(point);
                if (!row || typeof row.val !== "number") return null;
                return {
                    end: stringItem(row.end),
                    value: row.val,
                    fy: typeof row.fy === "number" ? row.fy : null,
                    fp: stringItem(row.fp),
                    form: stringItem(row.form),
                };
            })
            .filter((point): point is EdgarFactPoint => !!point)
            .filter((point) => point.form === "10-K" || point.form === "10-Q")
            .sort((a, b) => (b.end ?? "").localeCompare(a.end ?? ""));
        const deduped: EdgarFactPoint[] = [];
        const seenEnds = new Set<string>();
        for (const point of points) {
            const key = `${point.end}:${point.fp}`;
            if (seenEnds.has(key)) continue;
            seenEnds.add(key);
            deduped.push(point);
            if (deduped.length >= pointsPerConcept) break;
        }
        if (!deduped.length) {
            missing.push(concept);
            continue;
        }
        facts.push({
            concept,
            label: stringItem(conceptRecord?.label),
            unit,
            points: deduped,
        });
    }
    return { cik, entity_name: entityName, facts, missing_concepts: missing };
}

export async function getEdgarCompanyFacts(args: {
    cik: number | string;
    concepts?: string[];
}): Promise<EdgarFactsSummary> {
    const cik = padCik(args.cik);
    const payload = await edgarFetchJson<unknown>(
        `${EDGAR_DATA_BASE}/api/xbrl/companyfacts/CIK${cik}.json`,
        `loading XBRL company facts for CIK ${cik}`,
    );
    const concepts = (args.concepts ?? [])
        .map((concept) => concept.trim())
        .filter(Boolean);
    return summarizeCompanyFacts(
        payload,
        cik,
        concepts.length ? concepts : DEFAULT_FACT_CONCEPTS,
    );
}
