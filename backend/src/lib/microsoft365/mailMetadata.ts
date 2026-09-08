import { Microsoft365Error } from "./index";

export interface Microsoft365MailAddress {
  name: string | null;
  address: string | null;
}

interface MailDateTime {
  dateTime: string | null;
  timeZone: string | null;
}

export interface Microsoft365MailMetadata {
  id: string | null;
  changeKey: string | null;
  subject: string | null;
  bodyPreview: string | null;
  sender: Microsoft365MailAddress | null;
  from: Microsoft365MailAddress | null;
  toRecipients: Microsoft365MailAddress[] | null;
  ccRecipients: Microsoft365MailAddress[] | null;
  bccRecipients: Microsoft365MailAddress[] | null;
  replyTo: Microsoft365MailAddress[] | null;
  createdDateTime: string | null;
  lastModifiedDateTime: string | null;
  receivedDateTime: string | null;
  sentDateTime: string | null;
  hasAttachments: boolean | null;
  importance: string | null;
  isRead: boolean | null;
  isDraft: boolean | null;
  isDeliveryReceiptRequested: boolean | null;
  isReadReceiptRequested: boolean | null;
  inferenceClassification: string | null;
  categories: string[] | null;
  flag: {
    flagStatus: string | null;
    completedDateTime: MailDateTime | null;
    dueDateTime: MailDateTime | null;
    startDateTime: MailDateTime | null;
  } | null;
  internetMessageId: string | null;
  internetMessageHeaders: Array<{
    name: string | null;
    value: string | null;
  }> | null;
  conversationId: string | null;
  conversationIndex: string | null;
  parentFolderId: string | null;
  webLink: string | null;
}

/** Navigation properties (including attachments) are retrieved separately. */
export const MICROSOFT365_MAIL_SELECT = [
  "id",
  "changeKey",
  "subject",
  "body",
  "uniqueBody",
  "bodyPreview",
  "sender",
  "from",
  "toRecipients",
  "ccRecipients",
  "bccRecipients",
  "replyTo",
  "createdDateTime",
  "lastModifiedDateTime",
  "receivedDateTime",
  "sentDateTime",
  "hasAttachments",
  "importance",
  "isRead",
  "isDraft",
  "isDeliveryReceiptRequested",
  "isReadReceiptRequested",
  "inferenceClassification",
  "categories",
  "flag",
  "internetMessageId",
  "internetMessageHeaders",
  "conversationId",
  "conversationIndex",
  "parentFolderId",
  "webLink",
].join(",");

const MAX_METADATA_BYTES = 500_000;

function malformed(): never {
  throw new Microsoft365Error("provider_unavailable", 502);
}

function oversized(): never {
  throw new Microsoft365Error("source_too_large", 413);
}

function record(value: unknown): Record<string, unknown> {
  if (!value || typeof value !== "object" || Array.isArray(value)) malformed();
  return value as Record<string, unknown>;
}

function string(value: unknown, limit = 4096): string | null {
  if (value == null) return null;
  if (typeof value !== "string") malformed();
  if (value.length > limit) oversized();
  return value;
}

function boolean(value: unknown): boolean | null {
  if (value == null) return null;
  if (typeof value !== "boolean") malformed();
  return value;
}

function array<T>(value: unknown, normalize: (item: unknown) => T): T[] | null {
  if (value == null) return null;
  if (!Array.isArray(value)) malformed();
  if (value.length > 1000) oversized();
  return value.map(normalize);
}

function address(value: unknown): Microsoft365MailAddress | null {
  if (value == null) return null;
  const recipient = record(value);
  // A recipient with unavailable address data is not evidence of an empty address.
  if (recipient.emailAddress == null) return null;
  const email = record(recipient.emailAddress);
  return { name: string(email.name), address: string(email.address) };
}

function recipients(value: unknown): Microsoft365MailAddress[] | null {
  return array(value, (item) => {
    const result = address(item);
    if (result === null) malformed();
    return result;
  });
}

function dateTime(value: unknown): MailDateTime | null {
  if (value == null) return null;
  const date = record(value);
  return {
    dateTime: string(date.dateTime, 128),
    timeZone: string(date.timeZone, 256),
  };
}

function flag(value: unknown): Microsoft365MailMetadata["flag"] {
  if (value == null) return null;
  const followup = record(value);
  return {
    flagStatus: string(followup.flagStatus, 128),
    completedDateTime: dateTime(followup.completedDateTime),
    dueDateTime: dateTime(followup.dueDateTime),
    startDateTime: dateTime(followup.startDateTime),
  };
}

/** Only documented metadata is retained. Missing fields remain explicitly unknown. */
export function normalizeMicrosoft365MailMetadata(
  mail: unknown,
): Microsoft365MailMetadata {
  const data = record(mail);
  const metadata: Microsoft365MailMetadata = {
    id: string(data.id),
    changeKey: string(data.changeKey),
    subject: string(data.subject, 8192),
    bodyPreview: string(data.bodyPreview, 100_000),
    sender: address(data.sender),
    from: address(data.from),
    toRecipients: recipients(data.toRecipients),
    ccRecipients: recipients(data.ccRecipients),
    bccRecipients: recipients(data.bccRecipients),
    replyTo: recipients(data.replyTo),
    createdDateTime: string(data.createdDateTime, 128),
    lastModifiedDateTime: string(data.lastModifiedDateTime, 128),
    receivedDateTime: string(data.receivedDateTime, 128),
    sentDateTime: string(data.sentDateTime, 128),
    hasAttachments: boolean(data.hasAttachments),
    importance: string(data.importance, 128),
    isRead: boolean(data.isRead),
    isDraft: boolean(data.isDraft),
    isDeliveryReceiptRequested: boolean(data.isDeliveryReceiptRequested),
    isReadReceiptRequested: boolean(data.isReadReceiptRequested),
    inferenceClassification: string(data.inferenceClassification, 128),
    categories: array(data.categories, (item) => {
      const result = string(item);
      if (result === null) malformed();
      return result;
    }),
    flag: flag(data.flag),
    internetMessageId: string(data.internetMessageId),
    internetMessageHeaders: array(data.internetMessageHeaders, (item) => {
      const header = record(item);
      return {
        name: string(header.name, 1024),
        value: string(header.value, 65_536),
      };
    }),
    conversationId: string(data.conversationId),
    conversationIndex: string(data.conversationIndex, 16_384),
    parentFolderId: string(data.parentFolderId),
    webLink: string(data.webLink),
  };
  formatMicrosoft365MailMetadata(metadata);
  return metadata;
}

/** This same representation is supplied to the model and used as citation evidence. */
export function formatMicrosoft365MailMetadata(
  metadata: Microsoft365MailMetadata,
): string {
  const text = JSON.stringify(metadata);
  if (Buffer.byteLength(text, "utf8") > MAX_METADATA_BYTES) oversized();
  return text;
}
