# SEC EDGAR integration

Mike's assistant can research US public-company SEC filings through the free
EDGAR APIs: company lookup, filing history, full-text search across filings,
targeted search and bounded reading inside a filing document, and key
financial figures from audited XBRL data.

The `edgar_*` tools require no API key or database tables. In web and project
chat they are offered to the model only when the composer's **EDGAR** toggle
is on (off by default; the choice is remembered per browser and sent as
`use_edgar` on each request — an absent field keeps the tools enabled for
plain API callers). They also sit behind the same `includeResearchTools` gate
as CourtListener, so surfaces without research tools never see them.

## Identify your deployment

SEC's [fair-access policy](https://www.sec.gov/os/accessing-edgar-data) asks
API clients to send a descriptive `User-Agent` with contact information. Set
it in `backend/.env`:

```bash
EDGAR_USER_AGENT="MyFirm Mike (admin@myfirm.com)"
```

Without it, a generic MikeOSS identifier is sent. Requests are serialized with
a small delay so one backend stays far below SEC's 10 requests/second
threshold; when EDGAR still returns a rate-limit response, the assistant is
instructed to stop EDGAR calls for the rest of the turn.

## Behavior notes

- Filing documents are downloaded once per assistant turn and cached in
  memory; `edgar_find_in_filing` and `edgar_read_filing` work against that
  per-turn cache. Reading is bounded and paged so a large 10-K never enters
  the model context wholesale.
- Filing text is third-party content and is wrapped in the standard
  untrusted-content fence before it reaches the model.
- The assistant cites filings with direct `sec.gov` links in prose. EDGAR
  quotes do not participate in the `<CITATIONS>` verification block, which
  remains reserved for uploaded/generated documents and CourtListener cases.
