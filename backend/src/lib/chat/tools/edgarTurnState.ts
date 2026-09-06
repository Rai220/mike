import {
  accessionWithDashes,
  fetchEdgarFilingText,
  padCik,
  type EdgarFilingText,
} from "../../edgar";

// Filing text fetched from EDGAR is cached per assistant turn so the model
// can search and page through a large filing without re-downloading it (or
// re-reading it into context wholesale). Mirrors CourtListener's per-turn
// opinion cache.
export type EdgarTurnState = {
  filingsByKey: Map<string, EdgarFilingText>;
};

export function createEdgarTurnState(): EdgarTurnState {
  return { filingsByKey: new Map() };
}

function filingKey(cik: string, accession: string): string {
  return `${cik}:${accession}`;
}

export async function getOrFetchEdgarFiling(
  state: EdgarTurnState,
  args: {
    cik: number | string;
    accessionNumber: string;
    document?: string | null;
  },
): Promise<EdgarFilingText> {
  const cik = padCik(args.cik);
  const accession = accessionWithDashes(args.accessionNumber);
  const key = filingKey(cik, accession);
  const cached = state.filingsByKey.get(key);
  // A cached entry is reused even when a different document name is asked
  // for: within one turn the model works with the document resolved first.
  if (cached && (!args.document || cached.document === args.document)) {
    return cached;
  }
  const filing = await fetchEdgarFilingText(args);
  state.filingsByKey.set(key, filing);
  return filing;
}
