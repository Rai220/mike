import { describe, expect, it } from "vitest";
import {
    accessionNoDashes,
    accessionWithDashes,
    edgarFilingDocumentUrl,
    matchEdgarCompanies,
    padCik,
    parseCompanyTickers,
    parseFullTextSearch,
    parseSubmissions,
    summarizeCompanyFacts,
} from "./edgar";

describe("padCik", () => {
    it("pads numeric CIKs to 10 digits", () => {
        expect(padCik(320193)).toBe("0000320193");
        expect(padCik("320193")).toBe("0000320193");
        expect(padCik("CIK0000320193")).toBe("0000320193");
    });
});

describe("accession number formatting", () => {
    it("round-trips dashed and dashless forms", () => {
        expect(accessionWithDashes("0000320193-25-000073")).toBe(
            "0000320193-25-000073",
        );
        expect(accessionWithDashes("000032019325000073")).toBe(
            "0000320193-25-000073",
        );
        expect(accessionNoDashes("0000320193-25-000073")).toBe(
            "000032019325000073",
        );
    });
});

describe("edgarFilingDocumentUrl", () => {
    it("builds Archives URLs with unpadded CIK and dashless accession", () => {
        expect(
            edgarFilingDocumentUrl(
                "0000320193",
                "0000320193-25-000073",
                "aapl-20250927.htm",
            ),
        ).toBe(
            "https://www.sec.gov/Archives/edgar/data/320193/000032019325000073/aapl-20250927.htm",
        );
    });
});

describe("company ticker directory", () => {
    const companies = parseCompanyTickers({
        "0": { cik_str: 320193, ticker: "AAPL", title: "Apple Inc." },
        "1": { cik_str: 789019, ticker: "MSFT", title: "Microsoft Corp" },
        "2": { cik_str: 1018724, ticker: "AMZN", title: "Amazon com Inc" },
    });

    it("parses the directory shape", () => {
        expect(companies).toHaveLength(3);
        expect(companies[0]).toEqual({
            cik: "0000320193",
            ticker: "AAPL",
            name: "Apple Inc.",
        });
    });

    it("matches by exact ticker first", () => {
        const matches = matchEdgarCompanies(companies, "aapl", 5);
        expect(matches[0]?.cik).toBe("0000320193");
    });

    it("matches by name substring", () => {
        const matches = matchEdgarCompanies(companies, "micro", 5);
        expect(matches).toHaveLength(1);
        expect(matches[0]?.ticker).toBe("MSFT");
    });

    it("matches by CIK digits", () => {
        const matches = matchEdgarCompanies(companies, "1018724", 5);
        expect(matches[0]?.ticker).toBe("AMZN");
    });

    it("returns nothing for an empty query", () => {
        expect(matchEdgarCompanies(companies, "  ", 5)).toEqual([]);
    });
});

describe("parseSubmissions", () => {
    it("zips the parallel recent-filing arrays", () => {
        const { name, filings } = parseSubmissions(
            {
                name: "Apple Inc.",
                filings: {
                    recent: {
                        accessionNumber: [
                            "0000320193-25-000073",
                            "0000320193-25-000001",
                        ],
                        form: ["10-K", "8-K"],
                        filingDate: ["2025-11-01", "2025-01-02"],
                        reportDate: ["2025-09-27", ""],
                        primaryDocument: ["aapl-20250927.htm", ""],
                        primaryDocDescription: ["10-K", ""],
                    },
                },
            },
            "0000320193",
        );
        expect(name).toBe("Apple Inc.");
        expect(filings).toHaveLength(2);
        expect(filings[0]).toMatchObject({
            accession_number: "0000320193-25-000073",
            form: "10-K",
            filed: "2025-11-01",
            primary_document: "aapl-20250927.htm",
        });
        expect(filings[0].url).toContain("/320193/000032019325000073/");
        expect(filings[1].primary_document).toBeNull();
        expect(filings[1].url).toBeNull();
    });

    it("tolerates a malformed payload", () => {
        expect(parseSubmissions(null, "0000320193").filings).toEqual([]);
        expect(parseSubmissions({ filings: {} }, "0000320193").filings).toEqual(
            [],
        );
    });
});

describe("parseFullTextSearch", () => {
    it("parses hits with accession, document, and companies", () => {
        const { total, hits } = parseFullTextSearch({
            hits: {
                total: { value: 1454 },
                hits: [
                    {
                        _id: "0000815097-23-000012:ccl-20221130.htm",
                        _source: {
                            ciks: ["0000815097", "0001125259"],
                            display_names: [
                                "CARNIVAL CORP  (CCL)  (CIK 0000815097)",
                                "CARNIVAL PLC  (CUK, CUKPF)  (CIK 0001125259)",
                            ],
                            form: "10-K",
                            adsh: "0000815097-23-000012",
                            file_date: "2023-01-27",
                            file_type: "10-K",
                            file_description: "10-K",
                        },
                    },
                ],
            },
        });
        expect(total).toBe(1454);
        expect(hits).toHaveLength(1);
        expect(hits[0]).toMatchObject({
            accession_number: "0000815097-23-000012",
            document: "ccl-20221130.htm",
            form: "10-K",
            filed: "2023-01-27",
        });
        expect(hits[0].companies[0]).toEqual({
            cik: "0000815097",
            name: "CARNIVAL CORP  (CCL)",
        });
        expect(hits[0].url).toContain(
            "/815097/000081509723000012/ccl-20221130.htm",
        );
    });

    it("tolerates a malformed payload", () => {
        expect(parseFullTextSearch(null)).toEqual({ total: 0, hits: [] });
        expect(parseFullTextSearch({ hits: { hits: [{}] } }).hits).toEqual([]);
    });
});

describe("summarizeCompanyFacts", () => {
    const payload = {
        entityName: "Apple Inc.",
        facts: {
            "us-gaap": {
                NetIncomeLoss: {
                    label: "Net Income (Loss)",
                    units: {
                        USD: [
                            {
                                end: "2024-09-28",
                                val: 93736000000,
                                fy: 2024,
                                fp: "FY",
                                form: "10-K",
                            },
                            {
                                end: "2025-09-27",
                                val: 96150000000,
                                fy: 2025,
                                fp: "FY",
                                form: "10-K",
                            },
                            {
                                end: "2025-09-27",
                                val: 96150000000,
                                fy: 2025,
                                fp: "FY",
                                form: "S-8",
                            },
                        ],
                    },
                },
            },
        },
    };

    it("keeps only 10-K/10-Q points, newest first, deduped", () => {
        const summary = summarizeCompanyFacts(payload, "0000320193", [
            "NetIncomeLoss",
            "Revenues",
        ]);
        expect(summary.entity_name).toBe("Apple Inc.");
        expect(summary.facts).toHaveLength(1);
        expect(summary.facts[0].unit).toBe("USD");
        expect(summary.facts[0].points[0]).toMatchObject({
            end: "2025-09-27",
            value: 96150000000,
            form: "10-K",
        });
        expect(summary.facts[0].points).toHaveLength(2);
        expect(summary.missing_concepts).toEqual(["Revenues"]);
    });
});
