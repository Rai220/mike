import { createHash } from "node:crypto";
import { Worker } from "node:worker_threads";
import { pathToFileURL } from "node:url";
import { convert } from "html-to-text";
import type { createServerSupabase } from "../supabase";
import { Microsoft365Error, microsoft365GraphRequest } from "./index";
import {
  MICROSOFT365_MAIL_SELECT,
  normalizeMicrosoft365MailMetadata,
  type Microsoft365MailMetadata,
} from "./mailMetadata";
import {
  listMicrosoft365MailAttachments,
  readMicrosoft365MailAttachment,
  readMicrosoft365MailAttachmentMetadata,
  type AttachmentSourceMetadata,
  type Microsoft365AttachmentMetadata,
} from "./mailAttachments";

type Db = ReturnType<typeof createServerSupabase>;
export type SourceLocator =
  { kind: "mail"; id: string; attachmentId?: string } | { kind: "file"; driveId: string; id: string };
export interface Microsoft365Source {
  locator: SourceLocator;
  title: string;
  text: string;
  version: string;
  /** Graph changeKey/eTag for metadata-only ACL/version checks. */
  graphVersion?: string;
  webUrl?: string;
  fetchedAt: string;
  mail?: Microsoft365MailMetadata;
  uniqueBody?: string;
  attachment?: Microsoft365AttachmentMetadata;
  attachments?: { items: AttachmentSourceMetadata[]; more: boolean };
}
const MAX_BYTES = 10 * 1024 * 1024;
const MAX_JSON_BYTES = 2 * 1024 * 1024;
const MAX_TEXT = 100_000;
const PAGE_SIZE = 20;
const ID = /^[A-Za-z0-9_!.,=+-]{1,1024}$/;
function fail(
  code: ConstructorParameters<typeof Microsoft365Error>[0],
  status: number,
): never {
  throw new Microsoft365Error(code, status);
}
export function validateMicrosoft365SourceLocator(
  value: unknown,
): SourceLocator {
  if (!value || typeof value !== "object") fail("invalid_source", 400);
  const v = value as Record<string, unknown>;
  if (
    typeof v.id !== "string" ||
    !ID.test(v.id) ||
    v.id === "." ||
    v.id === ".."
  )
    fail("invalid_source", 400);
  if (v.kind === "mail") {
    if (v.attachmentId !== undefined &&
        (typeof v.attachmentId !== "string" || !ID.test(v.attachmentId) || v.attachmentId === "." || v.attachmentId === ".."))
      fail("invalid_source", 400);
    return { kind: "mail", id: v.id, ...(typeof v.attachmentId === "string" ? { attachmentId: v.attachmentId } : {}) };
  }
  if (
    v.kind === "file" &&
    typeof v.driveId === "string" &&
    ID.test(v.driveId) &&
    v.driveId !== "." &&
    v.driveId !== ".."
  ) {
    return { kind: "file", id: v.id, driveId: v.driveId };
  }
  return fail("invalid_source", 400);
}
function signalWithTimeout(signal?: AbortSignal) {
  const timeout = AbortSignal.timeout(30_000);
  return signal ? AbortSignal.any([signal, timeout]) : timeout;
}
function safeWebUrl(value: unknown): string | undefined {
  if (typeof value !== "string" || value.length > 4096) return undefined;
  try {
    const url = new URL(value);
    if (
      url.protocol !== "https:" ||
      url.username ||
      url.password ||
      (url.port && url.port !== "443")
    )
      return undefined;
    const host = url.hostname.toLowerCase();
    return host === "outlook.office.com" ||
      host === "outlook.office365.com" ||
      host.endsWith(".sharepoint.com") ||
      host === "onedrive.live.com"
      ? url.href
      : undefined;
  } catch {
    return undefined;
  }
}
async function checkResponse(response: Response): Promise<void> {
  if (response.ok) return;
  if (response.status === 401) fail("reconnect_required", 401);
  if (response.status === 403) fail("access_denied", 403);
  if (response.status === 404 || response.status === 410)
    fail("not_found", 404);
  if (response.status === 429 || response.status === 503) {
    const value = response.headers.get("retry-after");
    const seconds =
      value && /^\d+$/.test(value)
        ? Number(value)
        : value
          ? Math.ceil((Date.parse(value) - Date.now()) / 1000)
          : 30;
    throw new Microsoft365Error(
      "rate_limited",
      429,
      Math.min(3600, Math.max(1, Number.isFinite(seconds) ? seconds : 30)),
    );
  }
  fail("provider_unavailable", 502);
}
async function boundedBytes(
  response: Response,
  max: number,
  signal: AbortSignal,
): Promise<Buffer> {
  if (Number(response.headers.get("content-length")) > max)
    fail("source_too_large", 413);
  if (!response.body) return Buffer.alloc(0);
  const reader = response.body.getReader();
  const chunks: Uint8Array[] = [];
  let size = 0;
  const cancel = () => {
    void reader.cancel().catch(() => undefined);
  };
  signal.addEventListener("abort", cancel, { once: true });
  try {
    while (true) {
      signal.throwIfAborted();
      const { done, value } = await reader.read();
      signal.throwIfAborted();
      if (done) break;
      size += value.byteLength;
      if (size > max) fail("source_too_large", 413);
      chunks.push(value);
    }
    return Buffer.concat(chunks, size);
  } finally {
    signal.removeEventListener("abort", cancel);
    await reader.cancel().catch(() => undefined);
    reader.releaseLock();
  }
}
async function json(response: Response, signal: AbortSignal): Promise<any> {
  await checkResponse(response);
  try {
    return JSON.parse(
      (await boundedBytes(response, MAX_JSON_BYTES, signal)).toString("utf8"),
    );
  } catch (error) {
    if (error instanceof Microsoft365Error) throw error;
    return fail("provider_unavailable", 502);
  }
}
async function guarded<T>(fn: () => Promise<T>): Promise<T> {
  try {
    return await fn();
  } catch (error) {
    if (error instanceof Microsoft365Error) throw error;
    return fail("provider_unavailable", 502);
  }
}

/** Search metadata is disclosed only within the protected Microsoft 365 execution boundary. */
export async function searchMicrosoft365Sources(
  userId: string,
  connectionId: string,
  db: Db,
  input: {
    kind: "mail" | "file";
    query: string;
    offset?: number;
    mode?: "search" | "recent";
    folder?: "all" | "inbox" | "sentitems";
  },
  signal?: AbortSignal,
): Promise<{
  items: Array<{ locator: SourceLocator; title: string; webUrl?: string }>;
  more: boolean;
  nextOffset: number;
}> {
  if (
    !input ||
    !["mail", "file"].includes(input.kind) ||
    typeof input.query !== "string" ||
    (!input.query.trim() && input.mode !== "recent") ||
    (input.mode !== undefined && !["search", "recent"].includes(input.mode)) ||
    (input.mode === "recent" &&
      (input.kind !== "mail" || input.query.trim() !== "")) ||
    (input.folder !== undefined &&
      (input.kind !== "mail" || input.mode !== "recent" ||
        !["all", "inbox", "sentitems"].includes(input.folder))) ||
    input.query.length > 500 ||
    /[\u0000-\u001f]/.test(input.query)
  )
    fail("invalid_query", 400);
  const offset = input.offset ?? 0;
  if (!Number.isInteger(offset) || offset < 0 || offset > 980)
    fail("invalid_query", 400);
  const combined = signalWithTimeout(signal);
  return guarded(async () => {
    if (input.mode === "recent") {
      const folder = input.folder ?? "inbox";
      const path = folder === "all" ? "/me/messages" : `/me/mailFolders/${folder}/messages`;
      const query = new URLSearchParams({
        $top: String(PAGE_SIZE + 1),
        $skip: String(offset),
        $orderby: folder === "sentitems" ? "sentDateTime desc" : "receivedDateTime desc",
        $select: "id,subject",
      });
      const data = await json(
        await microsoft365GraphRequest(
          userId,
          connectionId,
          db,
          `${path}?${query}`,
          { headers: { Prefer: 'IdType="ImmutableId"' }, signal: combined },
        ),
        combined,
      );
      if (
        !Array.isArray(data?.value) ||
        data.value.some(
          (mail: any) => typeof mail?.id !== "string" || !ID.test(mail.id),
        )
      )
        fail("provider_unavailable", 502);
      return {
        items: data.value
          .slice(0, PAGE_SIZE)
          .map((mail: any) => ({
            locator: { kind: "mail" as const, id: mail.id },
            title: String(mail.subject ?? "(No subject)").slice(0, 500),
          })),
        more: offset < 980 && data.value.length > PAGE_SIZE,
        nextOffset: offset + PAGE_SIZE,
      };
    }
    const data = await json(
      await microsoft365GraphRequest(
        userId,
        connectionId,
        db,
        "/search/query",
        {
          method: "POST",
          signal: combined,
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            requests: [
              {
                entityTypes: [input.kind === "mail" ? "message" : "driveItem"],
                query: { queryString: input.query.trim() },
                from: offset,
                size: PAGE_SIZE,
              },
            ],
          }),
        },
      ),
      combined,
    );
    const container = data?.value?.[0]?.hitsContainers?.[0];
    // Graph omits hits entirely for a confirmed empty result (observed live).
    // Accept only that exact shape, not malformed or incomplete nonempty pages.
    const providerHits = container?.hits === undefined && container?.total === 0 && container?.moreResultsAvailable === false
      ? [] : container?.hits;
    if (!container || !Array.isArray(providerHits))
      fail("provider_unavailable", 502);
    const hits = providerHits.slice(0, PAGE_SIZE);
    let immutable = new Map<string, string>();
    if (input.kind === "mail" && hits.length) {
      // searchHit documents message hitId as RestId; resource.id may be absent.
      const ids = hits.map((hit: any) => hit?.hitId);
      if (ids.some((id: unknown) => typeof id !== "string" || !ID.test(id)))
        fail("provider_unavailable", 502);
      const translated = await json(
        await microsoft365GraphRequest(
          userId,
          connectionId,
          db,
          "/me/translateExchangeIds",
          {
            method: "POST",
            signal: combined,
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
              inputIds: ids,
              sourceIdType: "restId",
              targetIdType: "restImmutableEntryId",
            }),
          },
        ),
        combined,
      );
      if (!Array.isArray(translated?.value)) fail("provider_unavailable", 502);
      immutable = new Map(
        translated.value
          .filter(
            (row: any) =>
              typeof row.sourceId === "string" &&
              typeof row.targetId === "string" &&
              ID.test(row.targetId),
          )
          .map((row: any) => [row.sourceId, row.targetId]),
      );
    }
    const items = hits.flatMap((hit: any) => {
      const resource = hit?.resource;
      if (!resource) return [];
      if (input.kind === "mail") {
        const id = immutable.get(hit.hitId);
        if (!id) return [];
        return [
          {
            locator: { kind: "mail", id } as SourceLocator,
            title: String(resource.subject ?? "(No subject)").slice(0, 500),
            webUrl: safeWebUrl(resource.webLink),
          },
        ];
      }
      if (
        !resource.file ||
        !ID.test(resource.id ?? "") ||
        !ID.test(resource.parentReference?.driveId ?? "")
      )
        return [];
      return [
        {
          locator: {
            kind: "file",
            id: resource.id,
            driveId: resource.parentReference.driveId,
          } as SourceLocator,
          title: String(resource.name ?? "File").slice(0, 500),
          webUrl: safeWebUrl(resource.webUrl),
        },
      ];
    });
    return {
      items,
      more: offset < 980 && container.moreResultsAvailable === true,
      nextOffset: offset + PAGE_SIZE,
    };
  });
}

/** Downloads never forward a Graph bearer token to the signed content URL. */
async function download(
  response: Response,
  webUrl: string | undefined,
  signal: AbortSignal,
): Promise<Buffer> {
  let current = response;
  const sourceHost = webUrl ? new URL(webUrl).hostname : undefined;
  for (
    let redirects = 0;
    [301, 302, 303, 307, 308].includes(current.status);
    redirects++
  ) {
    if (redirects >= 3) fail("provider_unavailable", 502);
    const location = current.headers.get("location");
    if (!location) fail("provider_unavailable", 502);
    const target = new URL(location);
    const host = target.hostname.toLowerCase();
    if (
      target.protocol !== "https:" ||
      target.username ||
      target.password ||
      (target.port && target.port !== "443") ||
      !(
        (sourceHost &&
          host === sourceHost &&
          host.endsWith(".sharepoint.com")) ||
        host.endsWith(".files.1drv.com") ||
        host.endsWith(".storage.live.com")
      )
    )
      fail("access_denied", 403);
    await current.body?.cancel();
    current = await fetch(target, {
      method: "GET",
      redirect: "manual",
      signal,
      credentials: "omit",
    });
  }
  await checkResponse(current);
  return boundedBytes(current, MAX_BYTES, signal);
}

// A worker imposes a hard deadline on synchronous parsers and keeps malformed Office
// archives / huge PDF content streams away from the API event loop. No external IO.
const PARSER_WORKER = String.raw`
const { parentPort, workerData } = require('node:worker_threads');
(async () => {
 const buffer = Buffer.from(workerData.bytes);
 const ext = workerData.ext;
 const max = workerData.maxText;
 let text = '';
 if (['docx','xlsx','pptx'].includes(ext)) {
   const JSZip = require(workerData.modules.zip);
   const zip = await JSZip.loadAsync(buffer);
   const entries = Object.values(zip.files).filter(f => !f.dir);
   if (entries.length > 2000) throw new Error('limit');
   const declared = entries.reduce((sum, entry) => sum + (entry._data?.uncompressedSize || 0), 0);
   if (declared > 40 * 1024 * 1024) throw new Error('limit');
   let expanded = 0;
   for (const entry of entries) {
     await new Promise((resolve, reject) => {
       const stream = entry.nodeStream();
       stream.on('data', chunk => { expanded += chunk.length; if (expanded > 40 * 1024 * 1024) stream.destroy(new Error('limit')); });
       stream.on('error', reject); stream.on('end', resolve);
     });
   }
   if (ext === 'docx') {
     text = (await require(workerData.modules.mammoth).extractRawText({buffer})).value;
   } else if (ext === 'xlsx') {
     const XLSX = require(workerData.modules.xlsx);
     const workbook = XLSX.read(buffer, {type:'buffer', sheetRows:2001, dense:true, cellHTML:false, cellStyles:false, bookVBA:false});
     if (workbook.SheetNames.length > 20) throw new Error('limit');
     for (const name of workbook.SheetNames) {
       const sheet = workbook.Sheets[name];
       const range = XLSX.utils.decode_range(sheet['!fullref'] || sheet['!ref'] || 'A1');
       if (range.e.r > 1999 || range.e.c > 99) throw new Error('limit');
       text += '\n## ' + name + '\n' + XLSX.utils.sheet_to_csv(sheet);
       if (text.length > max) throw new Error('limit');
     }
   } else {
     const paths = Object.keys(zip.files).filter(p => /^ppt\/slides\/slide\d+\.xml$/i.test(p)).sort((a,b) => a.localeCompare(b, undefined, {numeric:true}));
     const { XMLParser } = require(workerData.modules.xml);
     const parser = new XMLParser({ignoreAttributes:true, processEntities:true, htmlEntities:false, trimValues:false});
     const collect = obj => {
       if (!obj || typeof obj !== 'object') return [];
       return Object.entries(obj).flatMap(([key,value]) => key === 'a:t' ? (Array.isArray(value) ? value : [value]).map(String) : Array.isArray(value) ? value.flatMap(collect) : collect(value));
     };
     for (let i = 0; i < paths.length; i++) {
       const xml = await zip.file(paths[i]).async('text');
       if (/<!DOCTYPE|<!ENTITY/i.test(xml)) throw new Error('format');
       const slideText = collect(parser.parse(xml)).join('\n').trim();
       if (slideText) text += '\n## Slide ' + (i+1) + '\n' + slideText;
       if (text.length > max) throw new Error('limit');
     }
   }
 } else if (ext === 'pdf') {
   const pdfjs = await import(workerData.modules.pdf);
   const task = pdfjs.getDocument({data:new Uint8Array(buffer), isEvalSupported:false, useSystemFonts:false, disableFontFace:true, verbosity:0, maxImageSize:1});
   const pdf = await task.promise;
   try {
     if (pdf.numPages > 200) throw new Error('limit');
     for (let i = 1; i <= pdf.numPages; i++) {
       const page = await pdf.getPage(i);
       const content = await page.getTextContent();
       const pageText = content.items.map(item => item.str || '').join(' ').trim();
       if (pageText) text += '\n## Page ' + i + '\n' + pageText;
       if (text.length > max) throw new Error('limit');
       page.cleanup();
     }
   } finally { await task.destroy(); }
 } else { text = buffer.toString('utf8'); }
 if (text.length > max) throw new Error('limit');
 if (!text.trim()) throw new Error('empty');
 parentPort.postMessage({text:text.trim()});
})().catch(error => parentPort.postMessage({error:error.message === 'limit' ? 'limit' : 'format'}));
`;
const MAX_PARSER_WORKERS = 2;
let activeParserWorkers = 0;

export async function extractFile(
  bytes: Buffer,
  ext: string,
  signal: AbortSignal,
): Promise<string> {
  if (activeParserWorkers >= MAX_PARSER_WORKERS)
    throw new Microsoft365Error("busy", 409, 1);
  signal.throwIfAborted();
  activeParserWorkers++;
  let released = false;
  const release = () => {
    if (!released) {
      released = true;
      activeParserWorkers--;
    }
  };
  return new Promise((resolve, reject) => {
    let worker: Worker;
    try {
      worker = new Worker(PARSER_WORKER, {
        eval: true,
        stdout: true,
        stderr: true,
        resourceLimits: {
          maxOldGenerationSizeMb: 128,
          maxYoungGenerationSizeMb: 16,
          stackSizeMb: 4,
        },
        workerData: {
          bytes,
          ext,
          maxText: MAX_TEXT,
          modules: {
            zip: require.resolve("jszip"),
            mammoth: require.resolve("mammoth"),
            xlsx: require.resolve("xlsx"),
            xml: require.resolve("fast-xml-parser"),
            pdf: pathToFileURL(
              require.resolve("pdfjs-dist/legacy/build/pdf.mjs"),
            ).href,
          },
        },
      });
    } catch {
      release();
      reject(new Microsoft365Error("provider_unavailable", 502));
      return;
    }
    // Consume parser diagnostics privately; source text must never reach application logs.
    worker.stdout?.resume();
    worker.stderr?.resume();
    let settled = false;
    const finish = (error?: Microsoft365Error, text?: string) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      signal.removeEventListener("abort", abort);
      // A completed parse is not necessarily an exited worker. Keep the slot
      // reserved until termination resolves or the actual exit event fires.
      void worker.terminate().then(
        () => {
          release();
          if (error) reject(error);
          else resolve(text!);
        },
        () => {
          // Do not release on failed termination: the worker could still be alive.
          reject(new Microsoft365Error("provider_unavailable", 502));
        },
      );
    };
    const abort = () =>
      finish(new Microsoft365Error("provider_unavailable", 502));
    const timer = setTimeout(
      () => finish(new Microsoft365Error("source_too_large", 413)),
      20_000,
    );
    signal.addEventListener("abort", abort, { once: true });
    if (signal.aborted) abort();
    worker.on("message", (value) => {
      if (typeof value?.text === "string" && value.text.length <= MAX_TEXT)
        finish(undefined, value.text);
      else
        finish(
          new Microsoft365Error(
            value?.error === "limit"
              ? "source_too_large"
              : "unsupported_source",
            value?.error === "limit" ? 413 : 415,
          ),
        );
    });
    worker.on("error", () =>
      finish(new Microsoft365Error("unsupported_source", 415)),
    );
    worker.on("exit", () => {
      release();
      if (!settled) finish(new Microsoft365Error("unsupported_source", 415));
    });
  });
}

/** Revalidate listed metadata without downloading mail bodies or file contents. */
export async function readMicrosoft365SourceMetadata(
  userId: string,
  connectionId: string,
  db: Db,
  sourceLocator: SourceLocator,
  signal?: AbortSignal,
): Promise<Omit<Microsoft365Source, "text">> {
  const locator = validateMicrosoft365SourceLocator(sourceLocator);
  const combined = signalWithTimeout(signal);
  return guarded(async () => {
    if (locator.kind === "mail" && locator.attachmentId)
      return readMicrosoft365MailAttachmentMetadata(userId, connectionId, db, locator, combined);
    const path =
      locator.kind === "mail"
        ? `/me/messages/${encodeURIComponent(locator.id)}?$select=id,subject,webLink,changeKey`
        : `/drives/${encodeURIComponent(locator.driveId)}/items/${encodeURIComponent(locator.id)}?$select=id,name,file,folder,webUrl,eTag,parentReference`;
    const data = await json(
      await microsoft365GraphRequest(userId, connectionId, db, path, {
        signal: combined,
        ...(locator.kind === "mail"
          ? { headers: { Prefer: 'IdType="ImmutableId"' } }
          : {}),
      }),
      combined,
    );
    if (
      data.id !== locator.id ||
      (locator.kind === "file" &&
        data.parentReference?.driveId !== locator.driveId)
    )
      fail("provider_unavailable", 502);
    if (
      locator.kind === "file" &&
      (!data.file || data.folder || typeof data.name !== "string")
    )
      fail("unsupported_source", 415);
    const version = locator.kind === "mail" ? data.changeKey : data.eTag;
    if (typeof version !== "string" || !version)
      fail("provider_unavailable", 502);
    return {
      locator,
      title: String(
        locator.kind === "mail" ? (data.subject ?? "(No subject)") : data.name,
      ).slice(0, 500),
      version,
      graphVersion: version,
      fetchedAt: new Date().toISOString(),
      webUrl: safeWebUrl(locator.kind === "mail" ? data.webLink : data.webUrl),
    };
  });
}

export async function readMicrosoft365Source(
  userId: string,
  connectionId: string,
  db: Db,
  locator: SourceLocator,
  signal?: AbortSignal,
): Promise<Microsoft365Source> {
  locator = validateMicrosoft365SourceLocator(locator);
  const combined = signalWithTimeout(signal);
  return guarded(async () => {
    if (locator.kind === "mail" && locator.attachmentId)
      return readMicrosoft365MailAttachment(userId, connectionId, db, locator, combined);
    if (locator.kind === "mail") {
      const path = `/me/messages/${encodeURIComponent(locator.id)}?$select=${MICROSOFT365_MAIL_SELECT}`;
      const headers = {
        Prefer: 'IdType="ImmutableId", outlook.body-content-type="text"',
      };
      const mail = await json(
        await microsoft365GraphRequest(userId, connectionId, db, path, {
          headers,
          signal: combined,
        }),
        combined,
      );
      if (
        mail.id !== locator.id ||
        typeof mail.body?.content !== "string" ||
        typeof mail.changeKey !== "string"
      )
        fail("provider_unavailable", 502);
      if (mail.body.content.length > MAX_TEXT) fail("source_too_large", 413);
      const text =
        mail.body.contentType?.toLowerCase() === "html"
          ? convert(mail.body.content, {
              wordwrap: false,
              selectors: [
                { selector: "img", format: "skip" },
                { selector: "a", options: { ignoreHref: true } },
              ],
            })
          : mail.body.content;
      if (text.length > MAX_TEXT) fail("source_too_large", 413);
      const metadata = normalizeMicrosoft365MailMetadata(mail);
      const attachments = await listMicrosoft365MailAttachments(userId, connectionId, db, locator.id, mail.changeKey, combined);
      let uniqueBody: string | undefined;
      if (mail.uniqueBody != null) {
        if (typeof mail.uniqueBody?.content !== "string") fail("provider_unavailable", 502);
        if (mail.uniqueBody.content.length > MAX_TEXT) fail("source_too_large", 413);
        uniqueBody = mail.uniqueBody.contentType?.toLowerCase() === "html"
          ? convert(mail.uniqueBody.content, { wordwrap: false, selectors: [{ selector: "img", format: "skip" }, { selector: "a", options: { ignoreHref: true } }] })
          : mail.uniqueBody.content;
        if (uniqueBody!.length > MAX_TEXT) fail("source_too_large", 413);
      }
      return {
        locator,
        title: String(mail.subject ?? "(No subject)").slice(0, 500),
        text,
        version: mail.changeKey,
        graphVersion: mail.changeKey,
        webUrl: safeWebUrl(mail.webLink),
        fetchedAt: new Date().toISOString(),
        mail: metadata,
        ...(uniqueBody !== undefined ? { uniqueBody } : {}),
        attachments,
      };
    }
    const path = `/drives/${encodeURIComponent(locator.driveId)}/items/${encodeURIComponent(locator.id)}`;
    const select =
      "?$select=id,name,size,file,folder,webUrl,eTag,cTag,parentReference";
    const metadata = await json(
      await microsoft365GraphRequest(userId, connectionId, db, path + select, {
        signal: combined,
      }),
      combined,
    );
    if (
      metadata.id !== locator.id ||
      metadata.parentReference?.driveId !== locator.driveId
    )
      fail("provider_unavailable", 502);
    if (!metadata.file || metadata.folder || typeof metadata.name !== "string")
      fail("unsupported_source", 415);
    if (
      typeof metadata.size !== "number" ||
      metadata.size < 0 ||
      metadata.size > MAX_BYTES
    )
      fail("source_too_large", 413);
    const ext = metadata.name.split(".").pop()?.toLowerCase() ?? "";
    if (!["txt", "md", "csv", "pdf", "docx", "xlsx", "pptx"].includes(ext))
      fail("unsupported_source", 415);
    const webUrl = safeWebUrl(metadata.webUrl);
    const version = metadata.eTag;
    if (typeof version !== "string" || !version)
      fail("provider_unavailable", 502);
    const content = await microsoft365GraphRequest(
      userId,
      connectionId,
      db,
      path + "/content",
      { signal: combined, redirect: "manual" },
    );
    const bytes = await download(content, webUrl, combined);
    const text = await extractFile(bytes, ext, combined);
    // Recheck ACL and version after the signed download and parsing, before persistence.
    const after = await json(
      await microsoft365GraphRequest(userId, connectionId, db, path + select, {
        signal: combined,
      }),
      combined,
    );
    if (
      after.id !== metadata.id ||
      after.eTag !== version ||
      after.parentReference?.driveId !== locator.driveId
    )
      fail("source_changed", 409);
    return {
      locator,
      title: metadata.name.slice(0, 500),
      text,
      graphVersion: version,
      version: createHash("sha256")
        .update(version)
        .update("\0")
        .update(bytes)
        .digest("hex"),
      webUrl,
      fetchedAt: new Date().toISOString(),
    };
  });
}

// Shared bounded Graph transport for mail-attachment reads. All callers keep
// account authorization inside microsoft365GraphRequest and use fixed paths.
export {
  json as readMicrosoft365GraphJson,
  boundedBytes as readMicrosoft365GraphBytes,
  checkResponse as checkMicrosoft365GraphResponse,
};
