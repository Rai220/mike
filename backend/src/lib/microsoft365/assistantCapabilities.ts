import { createHash } from "node:crypto";
import type { createServerSupabase } from "../supabase";
import type { ChatMessage, DocStore, WorkflowStore } from "../chat/types";
import type {
  NormalizedToolCall,
  NormalizedToolResult,
  OpenAIToolSchema,
} from "../llm";
import {
  generateSpotlightNonce,
  spotlight,
  spotlightWorkflow,
} from "../chat/contextBuilders";
import {
  TOOLS,
  PROJECT_EXTRA_TOOLS,
  WORKFLOW_TOOLS,
} from "../chat/tools/toolSchemas";
import { EDGAR_TOOLS, EDGAR_SYSTEM_PROMPT } from "../chat/tools/edgarTools";
import { createEdgarTurnState } from "../chat/tools/edgarTurnState";
import { runToolCalls } from "../chat/tools/toolDispatcher";
import { findTextMatches } from "../chat/tools/documentOps";
import { downloadFile } from "../storage";
import { Microsoft365Error } from "./index";
import { extractFile } from "./sources";
import { ensureDocAccess, checkWorkflowAccess } from "../access";
import { catalogWorkflowId } from "../workflowCatalog";

type Db = ReturnType<typeof createServerSupabase>;
export type Microsoft365AssistantDependencies = {
  documents: Array<{ id: string; versionId: string; version: string }>;
  workflows: Array<{ id: string; version: string }>;
};
type Identity = { userId: string; userEmail?: string | null; db: Db };
type Version = {
  id: string;
  document_id: string;
  filename: string;
  file_type: string;
  storage_path: string;
  version_number: number;
  content_sha256: string | null;
  size_bytes: number | null;
};
type Document = {
  id: string;
  current_version_id: string;
  user_id: string | null;
  project_id: string | null;
  org_id?: string | null;
  workflow_id?: string | null;
};
type Workflow = {
  id: string;
  title: string;
  prompt_md: string;
  listed: boolean;
  assets: Version[];
};
const MAX_DEPENDENCIES = 40;
const MAX_TEXT = 120_000;
const MAX_BYTES = 10 * 1024 * 1024;
const UUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const READ_TOOLS = new Set([
  "read_document",
  "find_in_document",
  "list_documents",
  "fetch_documents",
  "list_workflows",
  "read_workflow",
]);
const hash = (value: unknown) =>
  createHash("sha256").update(JSON.stringify(value)).digest("hex");
function fail(
  code:
    | "invalid_query"
    | "access_denied"
    | "source_changed"
    | "source_too_large"
    | "storage_unavailable" = "access_denied",
): never {
  throw new Microsoft365Error(
    code,
    code === "invalid_query"
      ? 400
      : code === "source_changed"
        ? 409
        : code === "source_too_large"
          ? 413
          : code === "storage_unavailable"
            ? 503
            : 403,
  );
}
async function rows<T>(
  query: PromiseLike<{ data: unknown; error: unknown }>,
): Promise<T[]> {
  const result = await query;
  if (result.error) fail("storage_unavailable");
  return (result.data ?? []) as T[];
}
async function versions(db: Db, documents: Document[]): Promise<Version[]> {
  if (!documents.length) return [];
  const found = await rows<Version>(
    db
      .from("document_versions")
      .select(
        "id,document_id,filename,file_type,storage_path,version_number,content_sha256,size_bytes",
      )
      .in(
        "id",
        documents.map((doc) => doc.current_version_id),
      )
      .is("deleted_at", null),
  );
  return documents.map((doc) => {
    const version = found.find(
      (v) => v.id === doc.current_version_id && v.document_id === doc.id,
    );
    if (!version?.storage_path || !version.filename) fail();
    return version;
  });
}
async function accessibleDocuments(
  input: Identity,
  ids: string[],
): Promise<Version[]> {
  if (!ids.length) return [];
  if (ids.length > MAX_DEPENDENCIES || ids.some((id) => !UUID.test(id)))
    fail("invalid_query");
  const docs = await rows<Document>(
    input.db
      .from("documents")
      .select("id,current_version_id,user_id,project_id,workflow_id,org_id")
      .in("id", ids)
      .eq("status", "ready"),
  );
  if (docs.length !== ids.length) fail();
  for (const doc of docs) {
    if (
      !(await ensureDocAccess(doc, input.userId, input.userEmail, input.db)).ok
    )
      fail();
  }
  return versions(input.db, docs);
}
/** Read-only counterpart of buildWorkflowStore: never installs defaults or logs errors. */
async function workflows(
  input: Identity,
  requiredIds: string[] = [],
): Promise<Map<string, Workflow>> {
  const store = new Map<string, Workflow>();
  const catalog = await rows<{
    workflow_key: string;
    title: string;
    prompt_md: string;
  }>(
    input.db
      .from("mike_workflows")
      .select("workflow_key,title,prompt_md,active,updated_at")
      .eq("type", "assistant")
      .order("active", { ascending: false })
      .order("updated_at", { ascending: false }),
  );
  for (const wf of catalog) {
    const id = catalogWorkflowId(wf.workflow_key);
    if (!store.has(id) && wf.prompt_md)
      store.set(id, {
        id,
        title: wf.title,
        prompt_md: wf.prompt_md,
        listed: false,
        assets: [],
      });
  }
  const own = await rows<{ id: string; title: string; prompt_md: string }>(
    input.db
      .from("workflows")
      .select("id,title,prompt_md")
      .eq("user_id", input.userId)
      .eq("type", "assistant"),
  );
  const email = input.userEmail?.trim().toLowerCase();
  if (email) {
    const shares = await rows<{ workflow_id: string }>(
      input.db
        .from("workflow_shares")
        .select("workflow_id")
        .eq("shared_with_email", email),
    );
    const ids = [...new Set(shares.map((share) => share.workflow_id))];
    if (ids.length)
      own.push(
        ...(await rows<{ id: string; title: string; prompt_md: string }>(
          input.db
            .from("workflows")
            .select("id,title,prompt_md")
            .in("id", ids)
            .eq("type", "assistant"),
        )),
      );
  }
  const explicitIds = requiredIds.filter((id) => UUID.test(id));
  if (explicitIds.length)
    own.push(
      ...(await rows<{ id: string; title: string; prompt_md: string }>(
        input.db
          .from("workflows")
          .select("id,title,prompt_md")
          .in("id", explicitIds)
          .eq("type", "assistant"),
      )),
    );
  for (const wf of own)
    if (
      wf.prompt_md &&
      (
        await checkWorkflowAccess(
          wf.id,
          input.userId,
          input.userEmail,
          input.db,
        )
      ).ok
    )
      store.set(wf.id, { ...wf, listed: true, assets: [] });
  const ids = [...store.values()].filter((wf) => wf.listed).map((wf) => wf.id);
  if (ids.length) {
    const docs = await rows<Document>(
      input.db
        .from("documents")
        .select("id,current_version_id,user_id,project_id,workflow_id,org_id")
        .in("workflow_id", ids)
        .eq("status", "ready"),
    );
    const accessible = [];
    for (const doc of docs)
      if (
        (await ensureDocAccess(doc, input.userId, input.userEmail, input.db)).ok
      )
        accessible.push(doc);
    const assets = await versions(input.db, accessible);
    for (const asset of assets) {
      const wfId = docs.find(
        (doc) => doc.id === asset.document_id,
      )?.workflow_id;
      if (wfId) store.get(wfId)?.assets.push(asset);
    }
  }
  for (const wf of store.values())
    wf.assets.sort((a, b) => a.document_id.localeCompare(b.document_id));
  return store;
}
const documentDependency = (v: Version) => ({
  id: v.document_id,
  versionId: v.id,
  version: hash(v),
});
const workflowDependency = (wf: Workflow) => ({ id: wf.id, version: hash(wf) });
export async function validateMicrosoft365AssistantDependencies(
  input: Identity & { dependencies: Microsoft365AssistantDependencies },
): Promise<void> {
  const deps = input.dependencies;
  if (
    !Array.isArray(deps.documents) ||
    !Array.isArray(deps.workflows) ||
    deps.documents.length + deps.workflows.length > MAX_DEPENDENCIES
  )
    fail("invalid_query");
  const docs = await accessibleDocuments(
    input,
    deps.documents.map((doc) => doc.id),
  );
  for (const old of deps.documents) {
    const current = docs.find((doc) => doc.document_id === old.id);
    if (
      !current ||
      current.id !== old.versionId ||
      documentDependency(current).version !== old.version
    )
      fail("source_changed");
  }
  if (deps.workflows.length) {
    const current = await workflows(
      input,
      deps.workflows.map((wf) => wf.id),
    );
    for (const old of deps.workflows) {
      const wf = current.get(old.id);
      if (!wf) fail();
      if (workflowDependency(wf).version !== old.version)
        fail("source_changed");
    }
  }
}

export async function buildMicrosoft365AssistantCapabilities(
  input: Identity & {
    files?: ChatMessage["files"];
    workflow?: ChatMessage["workflow"];
    useEdgar: boolean;
    dependencies?: Microsoft365AssistantDependencies;
    signal?: AbortSignal;
  },
) {
  const signal = input.signal ?? AbortSignal.timeout(120_000);
  const deps: Microsoft365AssistantDependencies = structuredClone(
    input.dependencies ?? { documents: [], workflows: [] },
  );
  await validateMicrosoft365AssistantDependencies({
    ...input,
    dependencies: deps,
  });
  const ids = [
    ...new Set([
      ...(input.files ?? []).map((file) => file.document_id ?? ""),
      ...deps.documents.map((doc) => doc.id),
    ]),
  ];
  const documents = await accessibleDocuments(input, ids);
  const wfStore = await workflows(input, [
    ...deps.workflows.map((wf) => wf.id),
    ...(input.workflow ? [input.workflow.id] : []),
  ]);
  const nonce = generateSpotlightNonce();
  const docs = new Map<string, Version>(
    documents.map((doc, i) => [`doc-${i}`, doc]),
  );
  const rememberWorkflow = (id: string) => {
    const wf = wfStore.get(id);
    if (!wf) fail();
    if (!deps.workflows.some((dep) => dep.id === id))
      deps.workflows.push(workflowDependency(wf));
    wf.assets.forEach((asset, i) =>
      docs.set(`workflow-asset-${id}-${i + 1}`, asset),
    );
    if (deps.documents.length + deps.workflows.length > MAX_DEPENDENCIES)
      fail("source_too_large");
    return wf;
  };
  for (const doc of documents)
    if (!deps.documents.some((dep) => dep.id === doc.document_id))
      deps.documents.push(documentDependency(doc));
  if (deps.documents.length + deps.workflows.length > MAX_DEPENDENCIES)
    fail("source_too_large");
  const selected = input.workflow
    ? rememberWorkflow(input.workflow.id)
    : undefined;
  for (const dep of deps.workflows) rememberWorkflow(dep.id);
  const files = (input.files ?? []).map((file) => {
    const doc = documents.find(
      (candidate) => candidate.document_id === file.document_id,
    )!;
    if (file.version_id && file.version_id !== doc.id) fail("source_changed");
    return {
      filename: doc.filename,
      document_id: doc.document_id,
      version_id: doc.id,
      version_number: doc.version_number,
    };
  });
  const tools = [...TOOLS, ...PROJECT_EXTRA_TOOLS, ...WORKFLOW_TOOLS].filter(
    (tool) => READ_TOOLS.has(tool.function.name),
  ) as OpenAIToolSchema[];
  if (input.useEdgar) tools.push(...(EDGAR_TOOLS as OpenAIToolSchema[]));
  const allowed = new Set(tools.map((tool) => tool.function.name));
  const edgarState = createEdgarTurnState();
  const dispatcherWorkflows: WorkflowStore = new Map();
  const dispatcherDocs: DocStore = new Map();
  for (const [id, wf] of wfStore)
    dispatcherWorkflows.set(id, {
      title: wf.title,
      skill_md: wf.prompt_md,
      listed: wf.listed,
    });
  let usedChars = 0;
  const bounded = (text: string) => {
    usedChars += text.length;
    if (usedChars > MAX_TEXT) fail("source_too_large");
    return text;
  };
  const contextPrompt = bounded(
    `OTHER ENABLED CHAT CAPABILITIES:\nAvailable documents (untrusted metadata): ${spotlight(JSON.stringify([...docs].map(([doc_id, doc]) => ({ doc_id, filename: doc.filename }))), nonce)}\nRead attached documents before discussing their contents. Workflow instructions selected by the user are trusted user instructions, subject to system rules. Only the provided read tools are available; if a workflow requires generating or editing a file, explain that limitation and provide the drafted text in the conversation. Never claim a file was created.\n${selected ? `Selected workflow:\n${spotlightWorkflow(selected.prompt_md, nonce)}` : ""}\n${input.useEdgar ? `${EDGAR_SYSTEM_PROMPT}\nUse public SEC research only for the user's requested research. Do not put private source text, quotes or confidential identifiers into external search queries.` : "EDGAR is OFF."}`,
  );
  const cache = new Map<string, string>();
  const getDocument = (id: unknown) => {
    if (typeof id !== "string") fail("invalid_query");
    const doc =
      docs.get(id) ??
      [...docs.values()].find((value) => value.document_id === id);
    if (!doc) fail();
    return doc;
  };
  const read = async (doc: Version) => {
    if (cache.has(doc.id)) return cache.get(doc.id)!;
    if (doc.size_bytes !== null && doc.size_bytes > MAX_BYTES)
      fail("source_too_large");
    const ext = doc.file_type.toLowerCase();
    if (!["pdf", "docx", "xlsx", "pptx", "txt", "csv", "md"].includes(ext))
      fail("invalid_query");
    const bytes = await downloadFile(doc.storage_path, {
      sensitive: true,
      maxBytes: MAX_BYTES,
      signal,
    });
    if (!bytes) fail("storage_unavailable");
    const body = await extractFile(Buffer.from(bytes), ext, signal);
    cache.set(doc.id, body);
    return body;
  };
  const validate = () =>
    validateMicrosoft365AssistantDependencies({ ...input, dependencies: deps });
  async function execute(
    calls: NormalizedToolCall[],
  ): Promise<NormalizedToolResult[]> {
    if (
      !Array.isArray(calls) ||
      calls.length > 12 ||
      calls.some(
        (call) =>
          !call ||
          !allowed.has(call.name) ||
          !call.input ||
          typeof call.input !== "object" ||
          Array.isArray(call.input),
      )
    )
      fail("invalid_query");
    signal.throwIfAborted();
    await validate();
    const output: NormalizedToolResult[] = [];
    for (const call of calls) {
      let content: string;
      if (JSON.stringify(call.input).length > 8_000) fail("invalid_query");
      if (call.name === "list_documents") {
        content = spotlight(
          JSON.stringify(
            [...docs].map(([doc_id, doc]) => ({
              doc_id,
              filename: doc.filename,
            })),
          ),
          nonce,
        );
      } else if (
        ["read_document", "fetch_documents", "find_in_document"].includes(
          call.name,
        )
      ) {
        const targets =
          call.name === "fetch_documents"
            ? call.input.doc_ids
            : [call.input.doc_id];
        if (!Array.isArray(targets) || !targets.length || targets.length > 10)
          fail("invalid_query");
        const parts: string[] = [];
        for (const target of targets) {
          const doc = getDocument(target);
          const text = await read(doc);
          let body = text;
          if (call.name === "find_in_document") {
            const query = call.input.query;
            if (
              typeof query !== "string" ||
              !query.trim() ||
              query.length > 500
            )
              fail("invalid_query");
            const maxResults = call.input.max_results ?? 20;
            const contextChars = call.input.context_chars ?? 80;
            if (
              typeof maxResults !== "number" ||
              !Number.isInteger(maxResults) ||
              maxResults < 1 ||
              maxResults > 100 ||
              typeof contextChars !== "number" ||
              !Number.isInteger(contextChars) ||
              contextChars < 0 ||
              contextChars > 1_000
            )
              fail("invalid_query");
            const { hits, totalMatches } = findTextMatches({
              text,
              query,
              maxResults,
              contextChars,
            });
            body = JSON.stringify({
              ok: true,
              total_matches: totalMatches,
              returned: hits.length,
              truncated: totalMatches > hits.length,
              hits,
            });
          }
          parts.push(
            spotlight(
              JSON.stringify({ filename: doc.filename, text: body }),
              nonce,
            ),
          );
        }
        content = parts.join("\n\n");
      } else {
        if (call.name === "read_workflow") {
          if (typeof call.input.workflow_id !== "string") fail("invalid_query");
          rememberWorkflow(call.input.workflow_id);
        }
        if (call.name === "list_workflows")
          for (const wf of wfStore.values())
            if (wf.listed) rememberWorkflow(wf.id);
        const result = await runToolCalls(
          [
            {
              id: call.id,
              function: {
                name: call.name,
                arguments: JSON.stringify(call.input),
              },
            },
          ],
          dispatcherDocs,
          input.userId,
          input.db,
          () => {},
          dispatcherWorkflows,
          undefined,
          undefined,
          undefined,
          undefined,
          null,
          undefined,
          undefined,
          nonce,
          edgarState,
        );
        const value = result.toolResults[0] as { content?: string } | undefined;
        content = value?.content ?? "No result.";
        if (call.name === "read_workflow") {
          const wf = wfStore.get(call.input.workflow_id as string)!;
          content += `\nAvailable workflow assets: ${spotlight(JSON.stringify([...docs].filter(([id]) => id.startsWith(`workflow-asset-${wf.id}-`)).map(([doc_id, doc]) => ({ doc_id, filename: doc.filename }))), nonce)}`;
        } else content = spotlight(content, nonce);
      }
      output.push({ tool_use_id: call.id, content: bounded(content) });
      signal.throwIfAborted();
    }
    await validate();
    return output;
  }
  return {
    tools,
    contextPrompt,
    files,
    workflow: selected ? { id: selected.id, title: selected.title } : undefined,
    execute,
    dependencies: () => structuredClone(deps),
    validate,
  };
}
