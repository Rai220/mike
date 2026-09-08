import { describe, expect, it } from "vitest";
import { Microsoft365Error } from "./index";
import {
  MICROSOFT365_MAIL_SELECT,
  formatMicrosoft365MailMetadata,
  normalizeMicrosoft365MailMetadata,
} from "./mailMetadata";

const recipient = (address: string, name = "Example") => ({
  emailAddress: { name, address },
});

describe("Microsoft 365 mail metadata", () => {
  it("selects all normalized message properties and both separately processed bodies", () => {
    const fields = MICROSOFT365_MAIL_SELECT.split(",");
    expect(fields).toEqual(
      expect.arrayContaining(
        Object.keys(normalizeMicrosoft365MailMetadata({})),
      ),
    );
    expect(fields).toContain("body");
    expect(fields).toContain("uniqueBody");
    expect(fields).not.toContain("attachments");
    expect(new Set(fields).size).toBe(fields.length);
  });

  it("retains sender, recipients, message status, threading, headers and follow-up dates", () => {
    const value = normalizeMicrosoft365MailMetadata({
      id: "immutable-id",
      changeKey: "version",
      subject: "Contract",
      bodyPreview: "Preview",
      sender: recipient("delegate@example.test"),
      from: recipient("author@example.test"),
      toRecipients: [recipient("to@example.test")],
      ccRecipients: [recipient("cc@example.test")],
      bccRecipients: [recipient("bcc@example.test")],
      replyTo: [recipient("reply@example.test")],
      createdDateTime: "2026-09-07T01:00:00Z",
      lastModifiedDateTime: "2026-09-07T02:00:00Z",
      receivedDateTime: "2026-09-07T03:00:00Z",
      sentDateTime: "2026-09-07T02:59:00Z",
      hasAttachments: true,
      importance: "high",
      isRead: false,
      isDraft: false,
      isDeliveryReceiptRequested: true,
      isReadReceiptRequested: false,
      inferenceClassification: "focused",
      categories: ["Legal"],
      flag: {
        flagStatus: "flagged",
        dueDateTime: { dateTime: "2026-09-08T12:00:00", timeZone: "UTC" },
      },
      internetMessageId: "<id@example.test>",
      internetMessageHeaders: [{ name: "X-Example", value: "value" }],
      conversationId: "conversation",
      conversationIndex: "base64",
      parentFolderId: "sent",
      webLink: "https://outlook.office.com/mail/id/example",
      body: { content: "Handled separately" },
      uniqueBody: { content: "Handled separately too" },
      unexpected: "must not propagate",
    });
    expect(value.from).toEqual({
      name: "Example",
      address: "author@example.test",
    });
    expect(value.sender?.address).toBe("delegate@example.test");
    expect(value.toRecipients?.[0].address).toBe("to@example.test");
    expect(value.ccRecipients?.[0].address).toBe("cc@example.test");
    expect(value.bccRecipients?.[0].address).toBe("bcc@example.test");
    expect(value.replyTo?.[0].address).toBe("reply@example.test");
    expect(value).toMatchObject({
      subject: "Contract",
      bodyPreview: "Preview",
      hasAttachments: true,
      isRead: false,
      isDraft: false,
      importance: "high",
      isDeliveryReceiptRequested: true,
      isReadReceiptRequested: false,
      categories: ["Legal"],
      flag: {
        flagStatus: "flagged",
        completedDateTime: null,
        startDateTime: null,
        dueDateTime: { dateTime: "2026-09-08T12:00:00", timeZone: "UTC" },
      },
      internetMessageHeaders: [{ name: "X-Example", value: "value" }],
      parentFolderId: "sent",
      conversationId: "conversation",
      conversationIndex: "base64",
    });
    const evidence = formatMicrosoft365MailMetadata(value);
    expect(evidence).toContain("to@example.test");
    expect(JSON.parse(evidence)).toEqual(value);
    expect(evidence).not.toContain("Handled separately");
    expect(evidence).not.toContain("must not propagate");
  });

  it("distinguishes unavailable recipients and booleans from provided empty or false values", () => {
    const value = normalizeMicrosoft365MailMetadata({
      toRecipients: [],
      ccRecipients: null,
      isRead: false,
    });
    expect(value.toRecipients).toEqual([]);
    expect(value.ccRecipients).toBeNull();
    expect(value.bccRecipients).toBeNull();
    expect(value.isRead).toBe(false);
    expect(value.isDraft).toBeNull();
    expect(value.internetMessageHeaders).toBeNull();
    expect(value.categories).toBeNull();
    expect(value.from).toBeNull();
  });

  it.each([
    null,
    [],
    "message",
    { subject: 1 },
    { isRead: "false" },
    { toRecipients: "recipient@example.test" },
    { toRecipients: [null] },
    { sender: { emailAddress: "sender@example.test" } },
    {
      bccRecipients: [
        recipient("example@example.test", 1 as unknown as string),
      ],
    },
    { internetMessageHeaders: ["bad"] },
    { categories: [null] },
    { flag: { dueDateTime: false } },
  ])("rejects malformed provider metadata without exposing it: %j", (input) => {
    expect(() => normalizeMicrosoft365MailMetadata(input)).toThrowError(
      Microsoft365Error,
    );
    expect(() => normalizeMicrosoft365MailMetadata(input)).toThrow(
      "provider_unavailable",
    );
  });

  it.each([
    { subject: "x".repeat(8193) },
    {
      toRecipients: Array.from({ length: 1001 }, () =>
        recipient("a@example.test"),
      ),
    },
    { internetMessageHeaders: [{ name: "X", value: "x".repeat(65_537) }] },
    { categories: Array.from({ length: 1001 }, () => "category") },
    { bodyPreview: "x".repeat(100_001) },
  ])("rejects oversized individual fields and collections", (input) => {
    expect(() => normalizeMicrosoft365MailMetadata(input)).toThrow(
      "source_too_large",
    );
  });

  it("bounds aggregate UTF-8 metadata bytes even when individual fields are within limits", () => {
    const input = {
      internetMessageHeaders: Array.from({ length: 8 }, () => ({
        name: "X",
        value: "я".repeat(40_000),
      })),
    };
    expect(() => normalizeMicrosoft365MailMetadata(input)).toThrow(
      "source_too_large",
    );
  });
});
