import { beforeEach, describe, expect, it, vi } from "vitest";
import { Document, Packer, Paragraph } from "docx";
vi.mock("./index", async (importOriginal) => ({
  ...(await importOriginal<typeof import("./index")>()),
  microsoft365GraphRequest: vi.fn(),
}));
import { microsoft365GraphRequest } from "./index";
import {
  listMicrosoft365MailAttachments,
  readMicrosoft365MailAttachment,
  readMicrosoft365MailAttachmentMetadata,
} from "./mailAttachments";
const graph = vi.mocked(microsoft365GraphRequest);
const db = {} as Parameters<typeof readMicrosoft365MailAttachment>[2];
const locator = {
  kind: "mail" as const,
  id: "mail-id",
  attachmentId: "attachment-id",
};
const response = (value: unknown, status = 200, headers?: HeadersInit) =>
  new Response(JSON.stringify(value), { status, headers });
const parent = (changeKey = "parent-v1") => ({ id: "mail-id", changeKey });
const metadata = (extra: Record<string, unknown> = {}) => ({
  id: "attachment-id",
  "@odata.type": "#microsoft.graph.fileAttachment",
  name: "note.txt",
  contentType: "text/plain",
  size: 12,
  isInline: false,
  lastModifiedDateTime: "2026-09-07T13:00:00Z",
  ...extra,
});
function metadataRead(
  extra: Record<string, unknown> = {},
  version = "parent-v1",
) {
  graph
    .mockResolvedValueOnce(response(parent(version)))
    .mockResolvedValueOnce(response(metadata(extra)))
    .mockResolvedValueOnce(response(parent(version)));
}
beforeEach(() => {
  graph.mockReset();
});
describe("mail attachment metadata", () => {
  it("lists inline metadata without requesting bytes and verifies the parent version", async () => {
    graph
      .mockResolvedValueOnce(
        response({ value: [metadata({ isInline: true })] }),
      )
      .mockResolvedValueOnce(response(parent()));
    const result = await listMicrosoft365MailAttachments(
      "u",
      "c",
      db,
      "mail-id",
      "parent-v1",
    );
    expect(result).toMatchObject({
      more: false,
      items: [
        {
          locator,
          attachment: { isInline: true, readable: true, type: "file" },
        },
      ],
    });
    expect(result.items[0]).not.toHaveProperty("text");
    const path = graph.mock.calls[0][3];
    expect(path).toContain("/me/messages/mail-id/attachments?");
    expect(path).not.toContain("contentBytes");
    expect(
      graph.mock.calls.every((call) => call[4]?.redirect === "manual"),
    ).toBe(true);
  });
  it("lists unsupported, reference and oversized attachments with honest limitations", async () => {
    graph
      .mockResolvedValueOnce(
        response({
          value: [
            metadata({ id: "image", name: "image.png", isInline: true }),
            metadata({
              id: "reference",
              "@odata.type": "#microsoft.graph.referenceAttachment",
              contentType: null,
            }),
            metadata({ id: "huge", size: 11 * 1024 * 1024 }),
          ],
        }),
      )
      .mockResolvedValueOnce(response(parent()));
    const result = await listMicrosoft365MailAttachments(
      "u",
      "c",
      db,
      "mail-id",
      "parent-v1",
    );
    expect(result.items).toHaveLength(3);
    expect(
      result.items.every(
        (item) => !item.attachment.readable && !!item.attachment.limitation,
      ),
    ).toBe(true);
  });
  it("follows bounded metadata pages for the exact same parent", async () => {
    graph
      .mockResolvedValueOnce(
        response({
          value: [metadata()],
          "@odata.nextLink":
            "https://graph.microsoft.com/v1.0/me/messages/mail-id/attachments?$skip=1",
        }),
      )
      .mockResolvedValueOnce(response({ value: [metadata({ id: "second" })] }))
      .mockResolvedValueOnce(response(parent()));
    const result = await listMicrosoft365MailAttachments(
      "u",
      "c",
      db,
      "mail-id",
      "parent-v1",
    );
    expect(result.items).toHaveLength(2);
    expect(result.more).toBe(false);
    expect(
      new URL(
        graph.mock.calls[1][3],
        "https://graph.microsoft.com",
      ).searchParams.get("$skip"),
    ).toBe("1");
  });
  it("returns explicit more after 100 entries", async () => {
    graph
      .mockResolvedValueOnce(
        response({
          value: Array.from({ length: 100 }, (_, n) =>
            metadata({ id: `a${n}` }),
          ),
          "@odata.nextLink":
            "https://graph.microsoft.com/v1.0/me/messages/mail-id/attachments?$skip=100",
        }),
      )
      .mockResolvedValueOnce(response(parent()));
    const result = await listMicrosoft365MailAttachments(
      "u",
      "c",
      db,
      "mail-id",
      "parent-v1",
    );
    expect(result.items).toHaveLength(100);
    expect(result.more).toBe(true);
    expect(graph).toHaveBeenCalledTimes(2);
  });
  it.each([
    "https://evil.test/v1.0/me/messages/mail-id/attachments?$skip=1",
    "https://graph.microsoft.com/v1.0/me/messages/other/attachments?$skip=1",
    "https://graph.microsoft.com/v1.0/me/messages/mail-id/attachments?$expand=item",
    "https://graph.microsoft.com/v1.0/me/messages/mail-id/attachments?$select=contentBytes",
  ])("rejects unsafe pagination %s before fetching it", async (next) => {
    graph.mockResolvedValueOnce(
      response({ value: [], "@odata.nextLink": next }),
    );
    await expect(
      listMicrosoft365MailAttachments("u", "c", db, "mail-id", "parent-v1"),
    ).rejects.toThrow("provider_unavailable");
    expect(graph).toHaveBeenCalledTimes(1);
  });
  it.each([
    { id: "../escape" },
    { "@odata.type": "unknown" },
    { "@odata.type": "__proto__" },
    { size: -1 },
    { size: "12" },
    { isInline: "yes" },
    { name: null },
    { lastModifiedDateTime: "bad" },
  ])("rejects malformed attachment metadata %j", async (extra) => {
    graph.mockResolvedValueOnce(response({ value: [metadata(extra)] }));
    await expect(
      listMicrosoft365MailAttachments("u", "c", db, "mail-id", "parent-v1"),
    ).rejects.toThrow("provider_unavailable");
  });
  it("rejects duplicate metadata IDs", async () => {
    graph.mockResolvedValueOnce(response({ value: [metadata(), metadata()] }));
    await expect(
      listMicrosoft365MailAttachments("u", "c", db, "mail-id", "parent-v1"),
    ).rejects.toThrow("provider_unavailable");
  });
  it("does not release listing after parent changes", async () => {
    graph
      .mockResolvedValueOnce(response({ value: [metadata()] }))
      .mockResolvedValueOnce(response(parent("changed")));
    await expect(
      listMicrosoft365MailAttachments("u", "c", db, "mail-id", "parent-v1"),
    ).rejects.toThrow("source_changed");
  });
  it("revalidates metadata with no content fetch", async () => {
    metadataRead();
    const result = await readMicrosoft365MailAttachmentMetadata(
      "u",
      "c",
      db,
      locator,
    );
    expect(result.locator).toEqual(locator);
    expect(result.version).toHaveLength(64);
    expect(result.graphVersion).toBe(result.version);
    expect(
      graph.mock.calls.every(
        (call) => !call[3].includes("$value") && !call[3].includes("$expand"),
      ),
    ).toBe(true);
  });
  it.each([
    { ...locator, attachmentId: ".." },
    { ...locator, attachmentId: "x/y" },
    { kind: "mail", id: "mail-id" },
    { ...locator, id: "../other" },
  ])("rejects unsafe attachment locator %j without Graph", async (unsafe) => {
    await expect(
      readMicrosoft365MailAttachmentMetadata(
        "u",
        "c",
        db,
        unsafe as typeof locator,
      ),
    ).rejects.toThrow("invalid_source");
    expect(graph).not.toHaveBeenCalled();
  });
  it("rejects a mismatched parent identity", async () => {
    graph.mockResolvedValueOnce(response({ id: "other", changeKey: "v" }));
    await expect(
      readMicrosoft365MailAttachmentMetadata("u", "c", db, locator),
    ).rejects.toThrow("provider_unavailable");
  });
});
describe("mail attachment evidence", () => {
  it("reads text through the isolated parser and verifies metadata again", async () => {
    metadataRead();
    graph.mockResolvedValueOnce(new Response("Contract evidence"));
    metadataRead();
    const result = await readMicrosoft365MailAttachment("u", "c", db, locator);
    expect(result).toMatchObject({ locator, text: "Contract evidence" });
    expect(graph.mock.calls[3][3]).toBe(
      "/me/messages/mail-id/attachments/attachment-id/$value",
    );
    expect(graph).toHaveBeenCalledTimes(7);
  });
  it("reads DOCX through the existing isolated Office parser", async () => {
    const bytes = await Packer.toBuffer(
      new Document({
        sections: [{ children: [new Paragraph("Attached Office evidence")] }],
      }),
    );
    const extra = {
      name: "letter.docx",
      contentType:
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      size: bytes.length,
    };
    metadataRead(extra);
    graph.mockResolvedValueOnce(new Response(bytes));
    metadataRead(extra);
    const result = await readMicrosoft365MailAttachment("u", "c", db, locator);
    expect(result.text).toContain("Attached Office evidence");
  });
  it("withholds parsed evidence when parent changes", async () => {
    metadataRead();
    graph.mockResolvedValueOnce(new Response("Contract evidence"));
    metadataRead({}, "changed");
    await expect(
      readMicrosoft365MailAttachment("u", "c", db, locator),
    ).rejects.toThrow("source_changed");
  });
  it("withholds parsed evidence when access is revoked", async () => {
    metadataRead();
    graph
      .mockResolvedValueOnce(new Response("Contract evidence"))
      .mockResolvedValueOnce(response({}, 403));
    await expect(
      readMicrosoft365MailAttachment("u", "c", db, locator),
    ).rejects.toThrow("access_denied");
  });
  it.each([302, 401, 403, 404, 429, 500])(
    "maps HTTP %s and never follows content redirects",
    async (status) => {
      metadataRead();
      graph.mockResolvedValueOnce(
        response({}, status, { location: "https://evil.test/secret" }),
      );
      await expect(
        readMicrosoft365MailAttachment("u", "c", db, locator),
      ).rejects.toBeDefined();
      expect(graph).toHaveBeenCalledTimes(4);
    },
  );
  it("bounds actual bytes despite a small metadata size", async () => {
    metadataRead();
    graph.mockResolvedValueOnce(
      new Response("x", {
        headers: { "content-length": String(11 * 1024 * 1024) },
      }),
    );
    await expect(
      readMicrosoft365MailAttachment("u", "c", db, locator),
    ).rejects.toThrow("source_too_large");
  });
  it("bounds a chunked response without trusting content-length", async () => {
    metadataRead();
    graph.mockResolvedValueOnce(
      new Response(
        new ReadableStream({
          start(controller) {
            controller.enqueue(new Uint8Array(6 * 1024 * 1024));
            controller.enqueue(new Uint8Array(6 * 1024 * 1024));
            controller.close();
          },
        }),
      ),
    );
    await expect(
      readMicrosoft365MailAttachment("u", "c", db, locator),
    ).rejects.toThrow("source_too_large");
  });
  it("does not return attachments after cancellation", async () => {
    const controller = new AbortController();
    controller.abort();
    graph.mockResolvedValueOnce(response(parent()));
    await expect(
      readMicrosoft365MailAttachment("u", "c", db, locator, controller.signal),
    ).rejects.toThrow("provider_unavailable");
  });
  it("rejects unsupported/reference formats before content fetch", async () => {
    metadataRead({ "@odata.type": "#microsoft.graph.referenceAttachment" });
    await expect(
      readMicrosoft365MailAttachment("u", "c", db, locator),
    ).rejects.toThrow("unsupported_source");
    expect(graph).toHaveBeenCalledTimes(3);
  });
  it("reads an attached email envelope and text without silently omitting nested attachment limitations", async () => {
    const extra = {
      "@odata.type": "#microsoft.graph.itemAttachment",
      name: "Forwarded message",
      contentType: null,
    };
    metadataRead(extra);
    graph.mockResolvedValueOnce(
      response({
        ...metadata(extra),
        item: {
          "@odata.type": "#microsoft.graph.message",
          subject: "Notice",
          toRecipients: [
            {
              emailAddress: {
                name: "Receiver",
                address: "receiver@example.test",
              },
            },
          ],
          body: { contentType: "html", content: "<p>Attached evidence</p>" },
          hasAttachments: true,
        },
      }),
    );
    metadataRead(extra);
    const result = await readMicrosoft365MailAttachment("u", "c", db, locator);
    expect(result.text).toContain("receiver@example.test");
    expect(result.text).toContain("Attached evidence");
    expect(result.text).toContain("Nested attachments were not retrieved");
    expect(result.text).not.toContain("<p>");
  });
  it("allows event/contact fields and does not copy arbitrary item properties", async () => {
    const extra = {
      "@odata.type": "#microsoft.graph.itemAttachment",
      name: "Contact",
    };
    metadataRead(extra);
    graph.mockResolvedValueOnce(
      response({
        ...metadata(extra),
        item: {
          "@odata.type": "#microsoft.graph.contact",
          displayName: "Name",
          emailAddresses: [{ address: "person@example.test" }],
          externalSecret: "DO NOT COPY",
        },
      }),
    );
    metadataRead(extra);
    const result = await readMicrosoft365MailAttachment("u", "c", db, locator);
    expect(result.text).toContain("person@example.test");
    expect(result.text).not.toContain("DO NOT COPY");
  });
});
