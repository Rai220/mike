export type EdgarToolEvent =
    | {
          type: "edgar_find_company";
          query: string;
          result_count: number;
          error?: string;
      }
    | {
          type: "edgar_list_filings";
          cik: string | null;
          company?: string | null;
          filing_count: number;
          error?: string;
      }
    | {
          type: "edgar_search_filings";
          query: string;
          result_count: number;
          total: number;
          error?: string;
      }
    | {
          type: "edgar_read_filing";
          accession_number: string | null;
          document?: string | null;
          chars_returned: number;
          total_chars: number;
          error?: string;
      }
    | {
          type: "edgar_find_in_filing";
          accession_number: string | null;
          query: string;
          total_matches: number;
          error?: string;
      }
    | {
          type: "edgar_company_facts";
          cik: string | null;
          concept_count: number;
          error?: string;
      };

export const EDGAR_TOOL_NAMES = {
    findCompany: "edgar_find_company",
    listFilings: "edgar_list_filings",
    searchFilings: "edgar_search_filings",
    readFiling: "edgar_read_filing",
    findInFiling: "edgar_find_in_filing",
    companyFacts: "edgar_company_facts",
} as const;

export const EDGAR_SYSTEM_PROMPT = `SEC EDGAR RESEARCH:
Use the edgar_* tools when the question involves a US public company's SEC filings: 10-K/10-Q/8-K disclosures, proxy statements, registration statements, or material agreements filed as exhibits.

Workflow:
1. Resolve the company first with edgar_find_company (ticker, name, or CIK) unless a CIK is already known in this conversation.
2. Locate filings with edgar_list_filings (a company's filing history, filterable by form) or edgar_search_filings (full-text search across filings).
3. Get cite-worthy passages with edgar_find_in_filing using short 1-3 word searches; use at most 3 find calls per assistant turn.
4. Only if snippets are insufficient, page through the document with edgar_read_filing (it returns one bounded chunk per call; use offset to continue). Never try to read a full 10-K end to end.
5. For headline financial figures (revenue, net income, assets), prefer edgar_company_facts (audited XBRL data) over quoting prose.

Citation rules for EDGAR material:
- Base every factual claim about a filing on text or XBRL data retrieved in this turn, never on memory of a filing.
- Cite a filing in prose with a clickable markdown link using the url returned by the tools, naming the company, form, and filing date (for example: [Apple Inc. 10-K (2025-11-01)](https://www.sec.gov/...)). Link each filing the first time you rely on it.
- Do not put EDGAR quotes in the <CITATIONS> block; that block is only for uploaded/generated documents and CourtListener cases.

Limits:
- If any EDGAR call returns a rate-limit or access-threshold error, stop all EDGAR calls for that turn and answer with what is already available.`;

export const EDGAR_TOOLS = [
    {
        type: "function",
        function: {
            name: EDGAR_TOOL_NAMES.findCompany,
            description:
                "Resolve a US public company in SEC EDGAR by ticker symbol, company name, or CIK number. Returns matching companies with their CIK, which the other edgar_* tools require.",
            parameters: {
                type: "object",
                properties: {
                    query: {
                        type: "string",
                        description:
                            "Ticker symbol (e.g. AAPL), company name (e.g. Apple), or CIK number.",
                    },
                    limit: {
                        type: "integer",
                        description: "Maximum matches to return. Default 8.",
                    },
                },
                required: ["query"],
            },
        },
    },
    {
        type: "function",
        function: {
            name: EDGAR_TOOL_NAMES.listFilings,
            description:
                "List a company's recent SEC filings from EDGAR by CIK, optionally filtered by form type and filing date. Returns filing metadata (form, dates, accession number, primary document) — not document text.",
            parameters: {
                type: "object",
                properties: {
                    cik: {
                        type: "string",
                        description:
                            "Company CIK from edgar_find_company or prior results.",
                    },
                    forms: {
                        type: "array",
                        items: { type: "string" },
                        description:
                            'Optional form types to include, e.g. ["10-K"], ["10-Q","8-K"], ["DEF 14A"].',
                    },
                    filedAfter: {
                        type: "string",
                        description: "Optional ISO date lower bound (YYYY-MM-DD).",
                    },
                    filedBefore: {
                        type: "string",
                        description: "Optional ISO date upper bound (YYYY-MM-DD).",
                    },
                    limit: {
                        type: "integer",
                        description: "Maximum filings to return. Default 20, max 50.",
                    },
                },
                required: ["cik"],
            },
        },
    },
    {
        type: "function",
        function: {
            name: EDGAR_TOOL_NAMES.searchFilings,
            description:
                "Full-text search across SEC EDGAR filings (2001-present). Use for finding filings that discuss a topic or contain specific language, optionally narrowed to one company, form types, or a date range. Quote multi-word phrases for exact-phrase matching. Returns filing metadata, not document text.",
            parameters: {
                type: "object",
                properties: {
                    query: {
                        type: "string",
                        description:
                            'Search terms. Use double quotes for exact phrases, e.g. "material adverse effect".',
                    },
                    forms: {
                        type: "array",
                        items: { type: "string" },
                        description: 'Optional form types, e.g. ["10-K","8-K"].',
                    },
                    cik: {
                        type: "string",
                        description: "Optional CIK to limit results to one company.",
                    },
                    filedAfter: {
                        type: "string",
                        description: "Optional ISO date lower bound (YYYY-MM-DD).",
                    },
                    filedBefore: {
                        type: "string",
                        description: "Optional ISO date upper bound (YYYY-MM-DD).",
                    },
                    limit: {
                        type: "integer",
                        description: "Maximum results to return. Default 10, max 25.",
                    },
                },
                required: ["query"],
            },
        },
    },
    {
        type: "function",
        function: {
            name: EDGAR_TOOL_NAMES.findInFiling,
            description:
                "Search within one EDGAR filing document for specific keyword(s) or phrases. Downloads and caches the filing for this turn on first use, then returns matches with surrounding context. Prefer this over edgar_read_filing for targeted lookups. Use no more than 3 calls per assistant turn.",
            parameters: {
                type: "object",
                properties: {
                    cik: {
                        type: "string",
                        description: "Company CIK the filing belongs to.",
                    },
                    accession_number: {
                        type: "string",
                        description:
                            "Filing accession number from edgar_list_filings or edgar_search_filings, e.g. 0000320193-25-000073.",
                    },
                    document: {
                        type: "string",
                        description:
                            "Optional document filename within the filing (primary_document from filing metadata). Defaults to the filing's primary HTML document.",
                    },
                    query: {
                        type: "string",
                        description:
                            "Short term to search for, 1-3 words likely to appear verbatim. Matching is case-insensitive and collapses whitespace.",
                    },
                    max_results: {
                        type: "integer",
                        description: "Maximum matches to return. Default 20.",
                    },
                    context_chars: {
                        type: "integer",
                        description:
                            "Characters of context on each side of each match. Default 160.",
                    },
                },
                required: ["cik", "accession_number", "query"],
            },
        },
    },
    {
        type: "function",
        function: {
            name: EDGAR_TOOL_NAMES.readFiling,
            description:
                "Read one bounded chunk of an EDGAR filing document's plain text. Downloads and caches the filing for this turn on first use. Returns at most max_chars characters starting at offset, plus the document's total length so you can page. Use edgar_find_in_filing first for targeted passages; never page through an entire large filing.",
            parameters: {
                type: "object",
                properties: {
                    cik: {
                        type: "string",
                        description: "Company CIK the filing belongs to.",
                    },
                    accession_number: {
                        type: "string",
                        description:
                            "Filing accession number, e.g. 0000320193-25-000073.",
                    },
                    document: {
                        type: "string",
                        description:
                            "Optional document filename within the filing. Defaults to the filing's primary HTML document.",
                    },
                    offset: {
                        type: "integer",
                        description:
                            "Character offset to start reading from. Default 0.",
                    },
                    max_chars: {
                        type: "integer",
                        description:
                            "Maximum characters to return. Default 20000, max 60000.",
                    },
                },
                required: ["cik", "accession_number"],
            },
        },
    },
    {
        type: "function",
        function: {
            name: EDGAR_TOOL_NAMES.companyFacts,
            description:
                "Get a company's key financial figures from SEC XBRL data (audited values as filed in 10-K/10-Q): revenue, net income, assets, liabilities, equity, cash, diluted EPS by default, or specific us-gaap concepts on request. Prefer this over quoting financial statements as prose.",
            parameters: {
                type: "object",
                properties: {
                    cik: {
                        type: "string",
                        description: "Company CIK from edgar_find_company.",
                    },
                    concepts: {
                        type: "array",
                        items: { type: "string" },
                        description:
                            'Optional specific us-gaap concept tags, e.g. ["OperatingIncomeLoss","LongTermDebt"]. Omit for the default key-figure set.',
                    },
                },
                required: ["cik"],
            },
        },
    },
];
