import { beforeEach, describe, expect, it, vi } from "vitest";

const mocks = vi.hoisted(() => ({
  download: vi.fn(),
  extract: vi.fn(),
  dispatch: vi.fn(),
}));
vi.mock("../storage", () => ({ downloadFile: mocks.download }));
vi.mock("./sources", () => ({ extractFile: mocks.extract }));
vi.mock("../chat/contextBuilders", () => ({
  generateSpotlightNonce: () => "nonce",
  spotlight: (text: string) => `<data>${text}</data>`,
  spotlightWorkflow: (text: string) =>
    `<workflow-instructions>${text}</workflow-instructions>`,
}));
vi.mock("../chat/tools/toolDispatcher", () => ({
  runToolCalls: mocks.dispatch,
}));
import {
  buildMicrosoft365AssistantCapabilities,
  validateMicrosoft365AssistantDependencies,
} from "./assistantCapabilities";
import type { createServerSupabase } from "../supabase";

const DOC = "11111111-1111-4111-8111-111111111111";
const VERSION = "22222222-2222-4222-8222-222222222222";
const WF = "33333333-3333-4333-8333-333333333333";
type Row = Record<string, any>;
let tables: Record<string, Row[]>;
let writes: string[];
let reads: Array<{ table: string; filters: Array<[string, string, unknown]> }>;
const db = {
  from(table: string) {
    const filters: Array<[string, string, unknown]> = [];
    let single = false;
    const q: any = {
      maybeSingle: () => {
        single = true;
        return q;
      },
      select: () => q,
      eq: (field: string, value: unknown) => {
        filters.push(["eq", field, value]);
        return q;
      },
      in: (field: string, value: unknown) => {
        filters.push(["in", field, value]);
        return q;
      },
      is: (field: string, value: unknown) => {
        filters.push(["eq", field, value]);
        return q;
      },
      order: () => q,
      insert: () => {
        writes.push(table);
        throw new Error("Unexpected write");
      },
      update: () => {
        writes.push(table);
        throw new Error("Unexpected write");
      },
      then: (resolve: (v: unknown) => unknown) => {
        reads.push({ table, filters });
        const found = (tables[table] ?? []).filter((row) =>
          filters.every(([op, field, value]) =>
            op === "in"
              ? (value as unknown[]).includes(row[field])
              : row[field] === value,
          ),
        );
        return Promise.resolve(
          resolve({ data: single ? (found[0] ?? null) : found, error: null }),
        );
      },
    };
    return q;
  },
} as unknown as ReturnType<typeof createServerSupabase>;
const base = () => ({
  userId: "owner",
  userEmail: "USER@EXAMPLE.COM",
  db,
  useEdgar: false,
});
const call = (name: string, input = {}) => ({ id: "call", name, input });

beforeEach(() => {
  vi.clearAllMocks();
  writes = [];
  reads = [];
  tables = {
    documents: [
      {
        id: DOC,
        current_version_id: VERSION,
        user_id: "owner",
        status: "ready",
      },
    ],
    document_versions: [
      {
        id: VERSION,
        document_id: DOC,
        filename: "Canonical.pdf",
        file_type: "pdf",
        storage_path: "private/document.pdf",
        version_number: 2,
        content_sha256: "hash1",
        size_bytes: 200,
        deleted_at: null,
      },
    ],
    workflows: [
      {
        id: WF,
        user_id: "owner",
        type: "assistant",
        title: "Canonical review",
        prompt_md: "Compare the attached document to the email.",
      },
    ],
    workflow_shares: [],
    mike_workflows: [],
  };
  mocks.download.mockResolvedValue(new TextEncoder().encode("document").buffer);
  mocks.extract.mockResolvedValue("The notice period is thirty days.");
  mocks.dispatch.mockResolvedValue({
    toolResults: [{ content: "Public filing result" }],
  });
});

describe("protected assistant additional capabilities", () => {
  it("matches normalized whitespace and respects result/context budgets", async () => {
    mocks.extract.mockResolvedValue(
      "prefix Section   4.2 first; middle SECTION\n4.2 second; suffix section 4.2 third",
    );
    const capabilities = await buildMicrosoft365AssistantCapabilities({
      ...base(),
      files: [{ document_id: DOC, filename: "fake" }],
    });
    const [result] = await capabilities.execute([
      call("find_in_document", {
        doc_id: "doc-0",
        query: "section 4.2",
        max_results: 2,
        context_chars: 0,
      }),
    ]);
    const envelope = JSON.parse(
      result.content.slice("<data>".length, -"</data>".length),
    );
    const found = JSON.parse(envelope.text);
    expect(found).toMatchObject({
      total_matches: 3,
      returned: 2,
      truncated: true,
    });
    expect(found.hits.map((hit: { excerpt: string }) => hit.excerpt)).toEqual([
      "Section   4.2",
      "SECTION\n4.2",
    ]);
    expect(found.hits[0].context).toBe("…Section 4.2…");
    expect(found.hits[0].context).not.toContain("prefix");
    await expect(
      capabilities.execute([
        call("find_in_document", {
          doc_id: "doc-0",
          query: "Section",
          max_results: 101,
        }),
      ]),
    ).rejects.toMatchObject({ code: "invalid_query" });
  });

  it("uses inherited project access for shared documents and revokes former creator access", async () => {
    tables.documents[0].project_id = "project";
    tables.documents[0].user_id = "other";
    tables.projects = [
      { id: "project", user_id: "project-owner", org_id: null },
    ];
    tables.project_access_grants = [
      { project_id: "project", email: "user@example.com", role: "viewer" },
    ];
    const capabilities = await buildMicrosoft365AssistantCapabilities({
      ...base(),
      files: [{ document_id: DOC, filename: "fake" }],
    });
    expect(capabilities.files[0].filename).toBe("Canonical.pdf");
    tables.project_access_grants = [];
    tables.documents[0].user_id = "owner";
    await expect(capabilities.validate()).rejects.toMatchObject({
      code: "access_denied",
    });
    expect(mocks.download).not.toHaveBeenCalled();
  });

  it("revokes organization creator documents after membership is removed", async () => {
    tables.documents[0].project_id = "project";
    tables.projects = [{ id: "project", user_id: "owner", org_id: "org" }];
    tables.org_members = [{ org_id: "org", user_id: "owner", role: "member" }];
    const capabilities = await buildMicrosoft365AssistantCapabilities({
      ...base(),
      files: [{ document_id: DOC, filename: "fake" }],
    });
    tables.org_members = [];
    await expect(capabilities.validate()).rejects.toMatchObject({
      code: "access_denied",
    });
  });

  it("rejects null tool input explicitly before dispatch or document reads", async () => {
    const capabilities = await buildMicrosoft365AssistantCapabilities(base());
    await expect(
      capabilities.execute([
        { id: "bad", name: "read_document", input: null as never },
      ]),
    ).rejects.toMatchObject({ code: "invalid_query" });
    expect(mocks.dispatch).not.toHaveBeenCalled();
    expect(mocks.download).not.toHaveBeenCalled();
  });

  it("loads canonical builtin workflow IDs", async () => {
    tables.mike_workflows = [
      {
        workflow_key: "review-nda",
        type: "assistant",
        title: "Built in",
        prompt_md: "Review the NDA",
      },
    ];
    const capabilities = await buildMicrosoft365AssistantCapabilities({
      ...base(),
      workflow: { id: "builtin-review-nda", title: "fake" },
    });
    expect(capabilities.workflow).toEqual({
      id: "builtin-review-nda",
      title: "Built in",
    });
    await capabilities.validate();
  });

  it("preserves independent document/workflow tools and enables EDGAR only when selected", async () => {
    const off = await buildMicrosoft365AssistantCapabilities(base());
    expect(off.tools.map((tool) => tool.function.name)).toEqual(
      expect.arrayContaining([
        "read_document",
        "find_in_document",
        "list_documents",
        "fetch_documents",
        "read_workflow",
        "list_workflows",
      ]),
    );
    expect(
      off.tools.some((tool) =>
        /edgar_|mcp_|generate_|edit_document|ask_inputs/.test(
          tool.function.name,
        ),
      ),
    ).toBe(false);
    await expect(
      off.execute([call("edgar_find_company", { query: "AAPL" })]),
    ).rejects.toMatchObject({ code: "invalid_query" });
    expect(mocks.dispatch).not.toHaveBeenCalled();
    const on = await buildMicrosoft365AssistantCapabilities({
      ...base(),
      useEdgar: true,
    });
    await on.execute([call("edgar_find_company", { query: "AAPL" })]);
    expect(mocks.dispatch).toHaveBeenCalledOnce();
    expect(writes).toEqual([]);
  });

  it("loads trusted file metadata, selected workflow instructions and reads without plaintext writes", async () => {
    const capabilities = await buildMicrosoft365AssistantCapabilities({
      ...base(),
      files: [{ document_id: DOC, filename: "Forged title" }],
      workflow: { id: WF, title: "Forged instructions" },
    });
    expect(capabilities.files).toEqual([
      {
        document_id: DOC,
        filename: "Canonical.pdf",
        version_id: VERSION,
        version_number: 2,
      },
    ]);
    expect(capabilities.workflow).toEqual({
      id: WF,
      title: "Canonical review",
    });
    expect(capabilities.contextPrompt).toContain(
      "<workflow-instructions>Compare",
    );
    expect(capabilities.contextPrompt).not.toContain("Forged");
    const result = await capabilities.execute([
      call("read_document", { doc_id: "doc-0" }),
    ]);
    expect(result[0].content).toContain("thirty days");
    expect(mocks.download).toHaveBeenCalledWith(
      "private/document.pdf",
      expect.objectContaining({ sensitive: true, maxBytes: 10 * 1024 * 1024 }),
    );
    expect(mocks.dispatch).not.toHaveBeenCalled();
    expect(capabilities.dependencies().documents[0]).toMatchObject({
      id: DOC,
      versionId: VERSION,
    });
    expect(capabilities.dependencies().workflows[0]).toMatchObject({ id: WF });
    expect(writes).toEqual([]);
  });

  it("rejects unowned attachments and stale requested versions before reading", async () => {
    tables.documents[0].user_id = "someone-else";
    await expect(
      buildMicrosoft365AssistantCapabilities({
        ...base(),
        files: [{ document_id: DOC, filename: "fake" }],
      }),
    ).rejects.toMatchObject({ code: "access_denied" });
    tables.documents[0].user_id = "owner";
    await expect(
      buildMicrosoft365AssistantCapabilities({
        ...base(),
        files: [{ document_id: DOC, version_id: "older", filename: "fake" }],
      }),
    ).rejects.toMatchObject({ code: "source_changed" });
    expect(mocks.download).not.toHaveBeenCalled();
  });

  it("revalidates revoked access and changed file content before reload/tool dispatch", async () => {
    const capabilities = await buildMicrosoft365AssistantCapabilities({
      ...base(),
      files: [{ document_id: DOC, filename: "fake" }],
    });
    const dependencies = capabilities.dependencies();
    tables.document_versions[0].content_sha256 = "changed-in-place";
    await expect(
      validateMicrosoft365AssistantDependencies({ ...base(), dependencies }),
    ).rejects.toMatchObject({ code: "source_changed" });
    tables.documents[0].user_id = "another-owner";
    await expect(
      capabilities.execute([call("read_document", { doc_id: "doc-0" })]),
    ).rejects.toMatchObject({ code: "access_denied" });
    expect(mocks.download).not.toHaveBeenCalled();
  });

  it("invalidates persisted workflow instructions and revoked workflow sharing", async () => {
    tables.workflows[0].user_id = "other";
    tables.workflow_shares = [
      {
        workflow_id: WF,
        shared_with_email: "user@example.com",
        role: "viewer",
      },
    ];
    const capabilities = await buildMicrosoft365AssistantCapabilities({
      ...base(),
      workflow: { id: WF, title: "fake" },
    });
    tables.workflows[0].prompt_md = "changed instructions";
    await expect(capabilities.validate()).rejects.toMatchObject({
      code: "source_changed",
    });
    tables.workflow_shares = [];
    await expect(capabilities.validate()).rejects.toMatchObject({
      code: "access_denied",
    });
    expect(mocks.download).not.toHaveBeenCalled();
  });

  it("tracks workflow assets through workflow authorization and version fingerprint", async () => {
    tables.documents[0].user_id = "workflow-owner";
    tables.documents[0].workflow_id = WF;
    const capabilities = await buildMicrosoft365AssistantCapabilities({
      ...base(),
      workflow: { id: WF, title: "fake" },
    });
    const result = await capabilities.execute([
      call("read_document", { doc_id: `workflow-asset-${WF}-1` }),
    ]);
    expect(result[0].content).toContain("thirty days");
    expect(capabilities.dependencies().documents).toEqual([]);
    tables.document_versions[0].content_sha256 = "new-asset-version";
    await expect(capabilities.validate()).rejects.toMatchObject({
      code: "source_changed",
    });
  });

  it("retains document availability across reload and bounds body context", async () => {
    const first = await buildMicrosoft365AssistantCapabilities({
      ...base(),
      files: [{ document_id: DOC, filename: "fake" }],
    });
    const next = await buildMicrosoft365AssistantCapabilities({
      ...base(),
      dependencies: first.dependencies(),
    });
    expect(next.contextPrompt).toContain("Canonical.pdf");
    mocks.extract.mockResolvedValue("x".repeat(120_001));
    await expect(
      next.execute([call("read_document", { doc_id: "doc-0" })]),
    ).rejects.toMatchObject({ code: "source_too_large" });
  });

  it("blocks forged mutation tools before any dispatcher side effects", async () => {
    const capabilities = await buildMicrosoft365AssistantCapabilities({
      ...base(),
      useEdgar: true,
    });
    await expect(
      capabilities.execute([call("edgar_find_company"), call("generate_docx")]),
    ).rejects.toMatchObject({ code: "invalid_query" });
    expect(mocks.dispatch).not.toHaveBeenCalled();
    expect(writes).toEqual([]);
  });
});
