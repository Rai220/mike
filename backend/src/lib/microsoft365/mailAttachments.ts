import { createHash } from "node:crypto";
import { convert } from "html-to-text";
import type { createServerSupabase } from "../supabase";
import { Microsoft365Error, microsoft365GraphRequest } from "./index";
import {
  formatMicrosoft365MailMetadata,
  normalizeMicrosoft365MailMetadata,
  type Microsoft365MailMetadata,
} from "./mailMetadata";
import {
  checkMicrosoft365GraphResponse,
  extractFile,
  readMicrosoft365GraphBytes,
  readMicrosoft365GraphJson,
  validateMicrosoft365SourceLocator,
  type Microsoft365Source,
  type SourceLocator,
} from "./sources";

type Db = ReturnType<typeof createServerSupabase>;
type MailLocator = Extract<SourceLocator, { kind: "mail" }>;
export interface Microsoft365AttachmentMetadata {
  name: string;
  contentType: string | null;
  size: number;
  isInline: boolean;
  type: "file" | "item" | "reference";
  lastModifiedDateTime?: string;
  readable: boolean;
  limitation?: string;
}
export interface AttachmentSourceMetadata extends Omit<
  Microsoft365Source,
  "text"
> {
  attachment: Microsoft365AttachmentMetadata;
}
const SELECT = "id,name,contentType,size,isInline,lastModifiedDateTime";
const MAX_BYTES = 10 * 1024 * 1024;
const MAX_ATTACHMENTS = 100;
const MAX_TEXT = 100_000;
const EXTENSIONS = new Set(["pdf", "docx", "xlsx", "pptx", "txt", "csv", "md"]);
const HEADERS = { Prefer: 'IdType="ImmutableId"' };
function fail(
  code: ConstructorParameters<typeof Microsoft365Error>[0],
  status: number,
): never {
  throw new Microsoft365Error(code, status);
}
function combinedSignal(signal?: AbortSignal): AbortSignal {
  const timeout = AbortSignal.timeout(60_000);
  return signal ? AbortSignal.any([signal, timeout]) : timeout;
}
async function guarded<T>(operation: () => Promise<T>): Promise<T> {
  try {
    return await operation();
  } catch (error) {
    if (error instanceof Microsoft365Error) throw error;
    return fail("provider_unavailable", 502);
  }
}
function parentLocator(parentId: string): MailLocator {
  return validateMicrosoft365SourceLocator({
    kind: "mail",
    id: parentId,
  }) as MailLocator;
}
function attachmentLocator(
  locator: SourceLocator,
): MailLocator & { attachmentId: string } {
  const checked = validateMicrosoft365SourceLocator(locator);
  if (checked.kind !== "mail" || !checked.attachmentId)
    fail("invalid_source", 400);
  return checked as MailLocator & { attachmentId: string };
}
function collectionPath(parentId: string) {
  return `/me/messages/${encodeURIComponent(parentId)}/attachments`;
}
function extension(name: string) {
  return name.split(".").pop()?.toLowerCase() ?? "";
}
function normalize(
  data: any,
  parentId: string,
  parentChangeKey: string,
): AttachmentSourceMetadata {
  if (!data || typeof data !== "object" || typeof data.id !== "string")
    fail("provider_unavailable", 502);
  let locator: MailLocator & { attachmentId: string };
  try {
    locator = attachmentLocator({
      kind: "mail",
      id: parentId,
      attachmentId: data.id,
    });
  } catch {
    return fail("provider_unavailable", 502);
  }
  const attachmentTypes: Record<
    string,
    Microsoft365AttachmentMetadata["type"]
  > = {
    "#microsoft.graph.fileAttachment": "file",
    "#microsoft.graph.itemAttachment": "item",
    "#microsoft.graph.referenceAttachment": "reference",
  };
  const graphType = data["@odata.type"];
  const type =
    typeof graphType === "string" && Object.hasOwn(attachmentTypes, graphType)
      ? attachmentTypes[graphType]
      : undefined;
  if (
    !type ||
    typeof data.name !== "string" ||
    !data.name ||
    data.name.length > 4096 ||
    (data.contentType !== null &&
      (typeof data.contentType !== "string" ||
        data.contentType.length > 512)) ||
    !Number.isSafeInteger(data.size) ||
    data.size < 0 ||
    typeof data.isInline !== "boolean" ||
    (data.lastModifiedDateTime !== undefined &&
      (typeof data.lastModifiedDateTime !== "string" ||
        !Number.isFinite(Date.parse(data.lastModifiedDateTime)))) ||
    typeof parentChangeKey !== "string" ||
    !parentChangeKey ||
    parentChangeKey.length > 4096
  )
    fail("provider_unavailable", 502);
  const limitation =
    type === "reference"
      ? "Reference attachment: external content is not downloaded."
      : data.size > MAX_BYTES
        ? "Attachment exceeds the 10 MiB reading limit."
        : type === "file" && !EXTENSIONS.has(extension(data.name))
          ? "Content extraction supports PDF, DOCX, XLSX, PPTX, TXT, CSV and MD only; this attachment is listed by metadata."
          : undefined;
  const attachment: AttachmentSourceMetadata["attachment"] = {
    name: data.name,
    contentType: data.contentType,
    size: data.size,
    isInline: data.isInline,
    type,
    ...(data.lastModifiedDateTime
      ? { lastModifiedDateTime: data.lastModifiedDateTime }
      : {}),
    readable: !limitation,
    ...(limitation ? { limitation } : {}),
  };
  const {
    readable: _readable,
    limitation: _limitation,
    ...graphMetadata
  } = attachment;
  const version = createHash("sha256")
    .update(
      JSON.stringify([parentChangeKey, locator.attachmentId, graphMetadata]),
    )
    .digest("hex");
  return {
    locator,
    title: data.name.slice(0, 500),
    attachment,
    version,
    graphVersion: version,
    fetchedAt: new Date().toISOString(),
  };
}
async function requestJson(
  userId: string,
  connectionId: string,
  db: Db,
  path: string,
  signal: AbortSignal,
) {
  return readMicrosoft365GraphJson(
    await microsoft365GraphRequest(userId, connectionId, db, path, {
      signal,
      headers: HEADERS,
      redirect: "manual",
    }),
    signal,
  );
}
async function parentVersion(
  userId: string,
  connectionId: string,
  db: Db,
  parentId: string,
  signal: AbortSignal,
): Promise<string> {
  const data = await requestJson(
    userId,
    connectionId,
    db,
    `/me/messages/${encodeURIComponent(parentId)}?$select=id,changeKey`,
    signal,
  );
  if (
    data?.id !== parentId ||
    typeof data.changeKey !== "string" ||
    !data.changeKey ||
    data.changeKey.length > 4096
  )
    fail("provider_unavailable", 502);
  return data.changeKey;
}
function continuation(value: unknown, parentId: string): string {
  if (typeof value !== "string" || value.length > 8192)
    fail("provider_unavailable", 502);
  let url: URL;
  try {
    url = new URL(value);
  } catch {
    return fail("provider_unavailable", 502);
  }
  const expected = `/v1.0${collectionPath(parentId)}`;
  if (
    url.origin !== "https://graph.microsoft.com" ||
    url.username ||
    url.password ||
    url.hash ||
    url.pathname !== expected
  )
    fail("provider_unavailable", 502);
  // Continuation may only paginate the same metadata projection, never expand bodies.
  for (const key of url.searchParams.keys())
    if (!["$skip", "$skiptoken", "$top", "$select"].includes(key))
      fail("provider_unavailable", 502);
  if (
    url.searchParams.has("$select") &&
    url.searchParams.get("$select") !== SELECT
  )
    fail("provider_unavailable", 502);
  url.searchParams.set("$select", SELECT);
  return url.pathname.slice("/v1.0".length) + url.search;
}

/** Includes inline attachments even when the parent message hasAttachments is false. */
export async function listMicrosoft365MailAttachments(
  userId: string,
  connectionId: string,
  db: Db,
  parentId: string,
  parentChangeKey: string,
  signal?: AbortSignal,
): Promise<{ items: AttachmentSourceMetadata[]; more: boolean }> {
  parentLocator(parentId);
  const combined = combinedSignal(signal);
  return guarded(async () => {
    if (
      typeof parentChangeKey !== "string" ||
      !parentChangeKey ||
      parentChangeKey.length > 4096
    )
      fail("invalid_source", 400);
    let path = `${collectionPath(parentId)}?$select=${SELECT}&$top=100`;
    const seen = new Set<string>();
    const ids = new Set<string>();
    const items: AttachmentSourceMetadata[] = [];
    let more = false;
    for (let page = 0; ; page++) {
      if (seen.has(path) || page >= MAX_ATTACHMENTS)
        fail("provider_unavailable", 502);
      seen.add(path);
      const data = await requestJson(userId, connectionId, db, path, combined);
      if (!data || !Array.isArray(data.value))
        fail("provider_unavailable", 502);
      for (const value of data.value) {
        const item = normalize(value, parentId, parentChangeKey);
        const id = (item.locator as MailLocator).attachmentId!;
        if (ids.has(id)) fail("provider_unavailable", 502);
        ids.add(id);
        if (items.length < MAX_ATTACHMENTS) items.push(item);
        else more = true;
      }
      const next = data["@odata.nextLink"];
      if (next !== undefined && next !== null) {
        const nextPath = continuation(next, parentId);
        if (items.length >= MAX_ATTACHMENTS) {
          more = true;
          break;
        }
        path = nextPath;
      } else break;
    }
    if (
      (await parentVersion(userId, connectionId, db, parentId, combined)) !==
      parentChangeKey
    )
      fail("source_changed", 409);
    return { items, more };
  });
}
export async function readMicrosoft365MailAttachmentMetadata(
  userId: string,
  connectionId: string,
  db: Db,
  sourceLocator: SourceLocator,
  signal?: AbortSignal,
): Promise<AttachmentSourceMetadata> {
  const locator = attachmentLocator(sourceLocator);
  const combined = combinedSignal(signal);
  return guarded(async () => {
    const before = await parentVersion(
      userId,
      connectionId,
      db,
      locator.id,
      combined,
    );
    const data = await requestJson(
      userId,
      connectionId,
      db,
      `${collectionPath(locator.id)}/${encodeURIComponent(locator.attachmentId)}?$select=${SELECT}`,
      combined,
    );
    if (data?.id !== locator.attachmentId) fail("provider_unavailable", 502);
    const metadata = normalize(data, locator.id, before);
    if (
      (await parentVersion(userId, connectionId, db, locator.id, combined)) !==
      before
    )
      fail("source_changed", 409);
    return metadata;
  });
}

function itemText(item: any): {
  text: string;
  mail?: Microsoft365MailMetadata;
} {
  if (!item || typeof item !== "object") fail("unsupported_source", 415);
  const type = item["@odata.type"];
  const fields =
    type === "#microsoft.graph.message"
      ? [
          "subject",
          "from",
          "sender",
          "toRecipients",
          "ccRecipients",
          "bccRecipients",
          "replyTo",
          "sentDateTime",
          "receivedDateTime",
          "internetMessageId",
          "importance",
          "isRead",
          "isDraft",
          "hasAttachments",
          "categories",
        ]
      : type === "#microsoft.graph.event"
        ? [
            "subject",
            "organizer",
            "attendees",
            "start",
            "end",
            "location",
            "isAllDay",
            "isCancelled",
            "importance",
            "sensitivity",
          ]
        : type === "#microsoft.graph.contact"
          ? [
              "displayName",
              "givenName",
              "surname",
              "companyName",
              "jobTitle",
              "emailAddresses",
              "businessPhones",
              "homePhones",
              "mobilePhone",
              "businessAddress",
              "homeAddress",
            ]
          : fail("unsupported_source", 415);
  const envelope = Object.fromEntries(
    fields
      .filter((field) => item[field] !== undefined)
      .map((field) => [field, item[field]]),
  );
  const mail =
    type === "#microsoft.graph.message"
      ? normalizeMicrosoft365MailMetadata(item)
      : undefined;
  let text = mail
    ? formatMicrosoft365MailMetadata(mail)
    : `Attached ${type.slice("#microsoft.graph.".length)} metadata:\n${JSON.stringify(envelope, null, 2)}`;
  if (item.body !== undefined) {
    if (
      !item.body ||
      typeof item.body.content !== "string" ||
      !["text", "html"].includes(String(item.body.contentType).toLowerCase())
    )
      fail("provider_unavailable", 502);
    const body =
      String(item.body.contentType).toLowerCase() === "html"
        ? convert(item.body.content, {
            wordwrap: false,
            selectors: [
              { selector: "a", options: { ignoreHref: true } },
              { selector: "img", format: "skip" },
            ],
          })
        : item.body.content;
    text += `\n\nBody:\n${body}`;
  }
  if (item.attachments !== undefined) {
    if (!Array.isArray(item.attachments)) fail("provider_unavailable", 502);
    const nested = item.attachments
      .slice(0, MAX_ATTACHMENTS)
      .map((attachment: any) => ({
        name:
          typeof attachment?.name === "string"
            ? attachment.name.slice(0, 4096)
            : null,
        contentType:
          typeof attachment?.contentType === "string"
            ? attachment.contentType.slice(0, 512)
            : null,
        size: Number.isSafeInteger(attachment?.size) ? attachment.size : null,
        isInline:
          typeof attachment?.isInline === "boolean"
            ? attachment.isInline
            : null,
        type:
          typeof attachment?.["@odata.type"] === "string"
            ? attachment["@odata.type"].slice(0, 128)
            : null,
      }));
    text += `\n\nNested attachments metadata (contents not read):\n${JSON.stringify(nested, null, 2)}`;
    if (item.attachments.length > MAX_ATTACHMENTS)
      text +=
        "\nAdditional nested attachments omitted (100 metadata entries limit).";
  } else if (
    type === "#microsoft.graph.message" ||
    type === "#microsoft.graph.event"
  ) {
    text +=
      "\n\nNested attachments were not retrieved; their absence has not been established.";
  }
  if (text.length > MAX_TEXT) fail("source_too_large", 413);
  return { text, ...(mail ? { mail } : {}) };
}
export async function readMicrosoft365MailAttachment(
  userId: string,
  connectionId: string,
  db: Db,
  sourceLocator: SourceLocator,
  signal?: AbortSignal,
): Promise<Microsoft365Source> {
  const locator = attachmentLocator(sourceLocator);
  const combined = combinedSignal(signal);
  return guarded(async () => {
    const before = await readMicrosoft365MailAttachmentMetadata(
      userId,
      connectionId,
      db,
      locator,
      combined,
    );
    if (before.attachment.size > MAX_BYTES) fail("source_too_large", 413);
    if (!before.attachment.readable) fail("unsupported_source", 415);
    const path = `${collectionPath(locator.id)}/${encodeURIComponent(locator.attachmentId)}`;
    let text: string;
    let mail: Microsoft365MailMetadata | undefined;
    if (before.attachment.type === "file") {
      const response = await microsoft365GraphRequest(
        userId,
        connectionId,
        db,
        `${path}/$value`,
        { signal: combined, headers: HEADERS, redirect: "manual" },
      );
      await checkMicrosoft365GraphResponse(response);
      text = await extractFile(
        await readMicrosoft365GraphBytes(response, MAX_BYTES, combined),
        extension(before.attachment.name),
        combined,
      );
    } else {
      const data = await requestJson(
        userId,
        connectionId,
        db,
        `${path}?$expand=microsoft.graph.itemattachment/item`,
        combined,
      );
      if (
        data?.id !== locator.attachmentId ||
        data["@odata.type"] !== "#microsoft.graph.itemAttachment"
      )
        fail("provider_unavailable", 502);
      ({ text, mail } = itemText(data.item));
    }
    const after = await readMicrosoft365MailAttachmentMetadata(
      userId,
      connectionId,
      db,
      locator,
      combined,
    );
    if (before.version !== after.version) fail("source_changed", 409);
    return { ...after, text, ...(mail ? { mail } : {}) };
  });
}
