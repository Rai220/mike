import { beforeEach, describe, expect, it, vi } from "vitest";
import JSZip from "jszip";
import { EventEmitter } from "node:events";
const workerMock = vi.hoisted(() => ({
  create: null as null | (() => unknown),
}));
vi.mock("node:worker_threads", async (importOriginal) => {
  const actual = await importOriginal<typeof import("node:worker_threads")>();
  return {
    ...actual,
    Worker: function (...args: ConstructorParameters<typeof actual.Worker>) {
      return workerMock.create
        ? workerMock.create()
        : new actual.Worker(...args);
    },
  };
});
import * as XLSX from "xlsx";
import { Document, Packer, Paragraph } from "docx";
vi.mock("./index", async (importOriginal) => ({
  ...(await importOriginal<typeof import("./index")>()),
  microsoft365GraphRequest: vi.fn(),
}));
import { microsoft365GraphRequest } from "./index";
import {
  readMicrosoft365Source,
  readMicrosoft365SourceMetadata,
  searchMicrosoft365Sources,
  validateMicrosoft365SourceLocator,
} from "./sources";
const graph = vi.mocked(microsoft365GraphRequest);
const db = {} as Parameters<typeof readMicrosoft365Source>[2];
const file = { kind: "file" as const, driveId: "drive-id", id: "file-id" };
const metadata = (name = "test.txt", eTag = '"v1"') => ({
  id: file.id,
  parentReference: { driveId: file.driveId },
  file: { mimeType: "text/plain" },
  name,
  size: 100,
  eTag,
  webUrl: "https://company.sharepoint.com/sites/team/test.txt",
});
const response = (value: unknown, status = 200, headers?: HeadersInit) =>
  new Response(JSON.stringify(value), { status, headers });
function prepareFile(
  bytes: Uint8Array | string,
  name = "test.txt",
  after = metadata(name),
) {
  graph
    .mockResolvedValueOnce(response(metadata(name)))
    .mockResolvedValueOnce(new Response(bytes as BodyInit))
    .mockResolvedValueOnce(response(after));
}
beforeEach(() => {
  workerMock.create = null;
  vi.restoreAllMocks();
  graph.mockReset();
  vi.unstubAllGlobals();
});

describe("source identifiers", () => {
  it.each([
    null,
    {},
    { kind: "mail", id: "../users/other" },
    { kind: "mail", id: "a?x=1" },
    { kind: "mail", id: "parent", attachmentId: "../other" },
    { kind: "mail", id: "parent", attachmentId: ".." },
    { kind: "mail", id: "parent", attachmentId: 123 },
    { kind: "file", id: "x", driveId: ".." },
    { kind: "file", id: "x", driveId: "https://evil" },
  ])("rejects unsafe locator %j", (value) => {
    expect(() => validateMicrosoft365SourceLocator(value)).toThrow(
      "invalid_source",
    );
  });
  it("accepts Graph immutable IDs and discards unrelated fields", () => {
    expect(
      validateMicrosoft365SourceLocator({
        kind: "mail",
        id: "AAMk+-_=",
        url: "https://evil",
      }),
    ).toEqual({ kind: "mail", id: "AAMk+-_=" });
  });
});
describe("metadata-only source revalidation", () => {
  it("checks immutable mail identity and changeKey without requesting the body", async () => {
    graph.mockResolvedValueOnce(
      response({
        id: "mail-id",
        subject: "Notice",
        changeKey: "v1",
        body: { content: "Do not return" },
      }),
    );
    const result = await readMicrosoft365SourceMetadata("u", "c", db, {
      kind: "mail",
      id: "mail-id",
    });
    expect(result).toMatchObject({
      title: "Notice",
      version: "v1",
      graphVersion: "v1",
    });
    expect(result).not.toHaveProperty("text");
    expect(graph).toHaveBeenCalledTimes(1);
    expect(graph.mock.calls[0]![3]).not.toContain("body");
    expect(graph.mock.calls[0]![4]?.headers).toEqual({
      Prefer: 'IdType="ImmutableId"',
    });
  });
  it("checks file eTag and drive identity without downloading unsupported or large contents", async () => {
    graph.mockResolvedValueOnce(
      response({ ...metadata("archive.zip"), size: 100_000_000 }),
    );
    const result = await readMicrosoft365SourceMetadata("u", "c", db, file);
    expect(result).toMatchObject({ title: "archive.zip", version: '"v1"' });
    expect(result).not.toHaveProperty("text");
    expect(graph).toHaveBeenCalledTimes(1);
    expect(graph.mock.calls[0]![3]).not.toContain("/content");
  });
  it.each([
    { id: "wrong", subject: "Notice", changeKey: "v1" },
    { id: "mail-id", subject: "Notice", changeKey: "" },
    { id: "mail-id", subject: "Notice" },
  ])("rejects mismatched or unversioned mail metadata", async (value) => {
    graph.mockResolvedValueOnce(response(value));
    await expect(
      readMicrosoft365SourceMetadata("u", "c", db, {
        kind: "mail",
        id: "mail-id",
      }),
    ).rejects.toMatchObject({ code: "provider_unavailable" });
  });
  it("rejects file metadata belonging to another drive", async () => {
    graph.mockResolvedValueOnce(
      response({ ...metadata(), parentReference: { driveId: "other" } }),
    );
    await expect(
      readMicrosoft365SourceMetadata("u", "c", db, file),
    ).rejects.toMatchObject({ code: "provider_unavailable" });
  });
  it("fails closed when the source ACL has been revoked", async () => {
    graph.mockResolvedValueOnce(response({}, 403));
    await expect(
      readMicrosoft365SourceMetadata("u", "c", db, file),
    ).rejects.toMatchObject({ code: "access_denied" });
  });
});
describe("metadata search", () => {
  it("translates documented message hitId and does not return snippets", async () => {
    graph.mockResolvedValueOnce(
      response({
        value: [
          {
            hitsContainers: [
              {
                moreResultsAvailable: true,
                hits: [
                  {
                    hitId: "rest-id",
                    summary: "SECRET SNIPPET",
                    resource: {
                      subject: "Subject",
                      webLink: "https://outlook.office.com/mail/item",
                    },
                  },
                ],
              },
            ],
          },
        ],
      }),
    );
    graph.mockResolvedValueOnce(
      response({ value: [{ sourceId: "rest-id", targetId: "immutable-id" }] }),
    );
    const result = await searchMicrosoft365Sources("u", "c", db, {
      kind: "mail",
      query: "contracts",
    });
    expect(result).toEqual({
      items: [
        {
          locator: { kind: "mail", id: "immutable-id" },
          title: "Subject",
          webUrl: "https://outlook.office.com/mail/item",
        },
      ],
      more: true,
      nextOffset: 20,
    });
    expect(graph.mock.calls[1][3]).toBe("/me/translateExchangeIds");
    expect(JSON.parse(String(graph.mock.calls[1][4]?.body))).toEqual({
      inputIds: ["rest-id"],
      sourceIdType: "restId",
      targetIdType: "restImmutableEntryId",
    });
    expect(JSON.stringify(result)).not.toContain("SECRET");
  });
  it("skips folders, requires drive ownership locator and strips unsafe web links", async () => {
    graph.mockResolvedValueOnce(
      response({
        value: [
          {
            hitsContainers: [
              {
                moreResultsAvailable: true,
                hits: [
                  {
                    resource: {
                      id: "folder",
                      name: "Folder",
                      folder: {},
                      parentReference: { driveId: "d" },
                    },
                  },
                  {
                    resource: {
                      id: "file",
                      name: "Document",
                      file: {},
                      parentReference: { driveId: "d" },
                      webUrl: "https://evil.test/file",
                    },
                  },
                ],
              },
            ],
          },
        ],
      }),
    );
    expect(
      await searchMicrosoft365Sources("u", "c", db, {
        kind: "file",
        query: "test",
        offset: 20,
      }),
    ).toEqual({
      items: [
        {
          locator: { kind: "file", id: "file", driveId: "d" },
          title: "Document",
          webUrl: undefined,
        },
      ],
      more: true,
      nextOffset: 40,
    });
  });
  it.each([
    { kind: "file", query: "" },
    { kind: "file", query: "x", offset: -1 },
    { kind: "file", query: "x", offset: 1001 },
    { kind: "mail", query: "x\n" },
  ])("rejects bad search %j before IO", async (input) => {
    await expect(
      searchMicrosoft365Sources("u", "c", db, input as any),
    ).rejects.toMatchObject({ code: "invalid_query", status: 400 });
    expect(graph).not.toHaveBeenCalled();
  });
  it("maps throttling without retaining the upstream body", async () => {
    graph.mockResolvedValueOnce(
      response({ error: "SECRET" }, 429, { "Retry-After": "120" }),
    );
    await expect(
      searchMicrosoft365Sources("u", "c", db, { kind: "file", query: "x" }),
    ).rejects.toMatchObject({
      code: "rate_limited",
      status: 429,
      retryAfterSeconds: 120,
    });
  });
});
describe("mail reads", () => {
  it("uses immutable IDs and text preference; sanitizes HTML fallback", async () => {
    graph.mockResolvedValueOnce(
      response({
        id: "immutable",
        subject: "Hello",
        changeKey: "version",
        toRecipients: [{ emailAddress: { name: "Recipient", address: "recipient@example.test" } }],
        ccRecipients: [],
        from: { emailAddress: { name: "Sender", address: "sender@example.test" } },
        internetMessageHeaders: [{ name: "X-Test", value: "synthetic" }],
        hasAttachments: false,
        body: {
          contentType: "html",
          content:
            '<p>Allowed text</p><script>bad()</script><img src="https://tracker" />',
        },
      }),
    );
    graph.mockResolvedValueOnce(response({ value: [] }));
    graph.mockResolvedValueOnce(response({ id: "immutable", changeKey: "version" }));
    const result = await readMicrosoft365Source("u", "c", db, {
      kind: "mail",
      id: "immutable",
    });
    expect(result.text).toBe("Allowed text");
    expect(result.version).toBe("version");
    expect(result.mail).toMatchObject({ toRecipients: [{ name: "Recipient", address: "recipient@example.test" }], ccRecipients: [], bccRecipients: null, internetMessageHeaders: [{ name: "X-Test", value: "synthetic" }] });
    expect(result.attachments).toEqual({ items: [], more: false });
    expect(graph.mock.calls[0][3]).toContain("toRecipients");
    expect(graph.mock.calls[0][3]).toContain("internetMessageHeaders");
    expect(graph.mock.calls[1][3]).toContain("/attachments?");
    expect(graph.mock.calls[0][4]?.headers).toEqual({
      Prefer: 'IdType="ImmutableId", outlook.body-content-type="text"',
    });
  });
  it.each([
    [401, "reconnect_required"],
    [403, "access_denied"],
    [404, "not_found"],
    [410, "not_found"],
  ])("maps HTTP %s without exposing server errors", async (status, code) => {
    graph.mockResolvedValueOnce(
      response({ message: "SECRET" }, status as number),
    );
    await expect(
      readMicrosoft365Source("u", "c", db, { kind: "mail", id: "id" }),
    ).rejects.toMatchObject({ code });
  });
  it("rejects oversized response without reading content", async () => {
    graph.mockResolvedValueOnce(
      new Response("{}", { headers: { "content-length": "999999999" } }),
    );
    await expect(
      readMicrosoft365Source("u", "c", db, { kind: "mail", id: "id" }),
    ).rejects.toMatchObject({ code: "source_too_large" });
  });
});
describe("file reads and parser isolation", () => {
  it("reads text and rechecks ACL/version after download", async () => {
    prepareFile("Corporate text");
    const result = await readMicrosoft365Source("u", "c", db, file);
    expect(result.text).toBe("Corporate text");
    expect(result.version).toMatch(/^[a-f0-9]{64}$/);
    expect(graph).toHaveBeenCalledTimes(3);
  });
  it("discards content if source changed during the download", async () => {
    prepareFile("Corporate text", "test.txt", metadata("test.txt", '"v2"'));
    await expect(
      readMicrosoft365Source("u", "c", db, file),
    ).rejects.toMatchObject({ code: "source_changed", status: 409 });
  });
  it("discards content if access is revoked after download", async () => {
    graph
      .mockResolvedValueOnce(response(metadata()))
      .mockResolvedValueOnce(new Response("text"))
      .mockResolvedValueOnce(response({}, 403));
    await expect(
      readMicrosoft365Source("u", "c", db, file),
    ).rejects.toMatchObject({ code: "access_denied" });
  });
  it("does not forward credentials to a tenant signed download", async () => {
    graph
      .mockResolvedValueOnce(response(metadata()))
      .mockResolvedValueOnce(
        new Response(null, {
          status: 302,
          headers: {
            Location: "https://company.sharepoint.com/download?sig=private",
          },
        }),
      )
      .mockResolvedValueOnce(response(metadata()));
    const fetcher = vi.fn().mockResolvedValue(new Response("Downloaded text"));
    vi.stubGlobal("fetch", fetcher);
    expect((await readMicrosoft365Source("u", "c", db, file)).text).toBe(
      "Downloaded text",
    );
    expect(fetcher.mock.calls[0][1]).toMatchObject({
      redirect: "manual",
      credentials: "omit",
    });
    expect(fetcher.mock.calls[0][1].headers).toBeUndefined();
  });
  it.each([
    "http://company.sharepoint.com/a",
    "https://evil.test/a",
    "https://another.sharepoint.com/a",
    "https://company.sharepoint.com.evil.test/a",
    "https://127.0.0.1/a",
    "https://user:pass@company.sharepoint.com/a",
  ])("rejects download redirect %s", async (location) => {
    graph
      .mockResolvedValueOnce(response(metadata()))
      .mockResolvedValueOnce(
        new Response(null, { status: 302, headers: { Location: location } }),
      );
    const fetcher = vi.fn();
    vi.stubGlobal("fetch", fetcher);
    await expect(
      readMicrosoft365Source("u", "c", db, file),
    ).rejects.toMatchObject({ code: "access_denied" });
    expect(fetcher).not.toHaveBeenCalled();
  });
  it("parses PPTX text in slide order with XML decoding", async () => {
    const zip = new JSZip();
    zip.file(
      "ppt/slides/slide2.xml",
      "<p:sld><a:t>Second &amp; final</a:t></p:sld>",
    );
    zip.file("ppt/slides/slide1.xml", "<p:sld><a:t>First</a:t></p:sld>");
    prepareFile(await zip.generateAsync({ type: "nodebuffer" }), "slides.pptx");
    expect((await readMicrosoft365Source("u", "c", db, file)).text).toBe(
      "## Slide 1\nFirst\n## Slide 2\nSecond & final",
    );
  });
  it("parses XLSX and preserves sheet/cell positions as CSV", async () => {
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(
      workbook,
      XLSX.utils.aoa_to_sheet([
        ["Name", "Value"],
        ["Budget", 123],
      ]),
      "Finance",
    );
    prepareFile(
      XLSX.write(workbook, { type: "buffer", bookType: "xlsx" }),
      "budget.xlsx",
    );
    expect((await readMicrosoft365Source("u", "c", db, file)).text).toContain(
      "Name,Value\nBudget,123",
    );
  });
  it("parses DOCX without invoking the document logging pipeline", async () => {
    const document = new Document({
      sections: [{ children: [new Paragraph("Confidential agreement")] }],
    });
    prepareFile(await Packer.toBuffer(document), "agreement.docx");
    expect((await readMicrosoft365Source("u", "c", db, file)).text).toBe(
      "Confidential agreement",
    );
  });
  it("parses PDF text with page provenance", async () => {
    const stream = "BT /F1 12 Tf 10 50 Td (Agreement text) Tj ET";
    const objects = [
      "<< /Type /Catalog /Pages 2 0 R >>",
      "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
      "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 100] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>",
      "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
      `<< /Length ${stream.length} >>\nstream\n${stream}\nendstream`,
    ];
    let pdf = "%PDF-1.4\n";
    const offsets = [0];
    objects.forEach((object, index) => {
      offsets.push(Buffer.byteLength(pdf));
      pdf += `${index + 1} 0 obj\n${object}\nendobj\n`;
    });
    const xref = Buffer.byteLength(pdf);
    pdf += `xref\n0 6\n0000000000 65535 f \n${offsets
      .slice(1)
      .map((offset) => String(offset).padStart(10, "0") + " 00000 n ")
      .join(
        "\n",
      )}\ntrailer\n<< /Size 6 /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF`;
    prepareFile(Buffer.from(pdf), "agreement.pdf");
    expect((await readMicrosoft365Source("u", "c", db, file)).text).toContain(
      "## Page 1\nAgreement text",
    );
  });
  it("rejects a compressed archive with excessive expanded bytes", async () => {
    const zip = new JSZip();
    zip.file("ppt/slides/slide1.xml", "a".repeat(41 * 1024 * 1024));
    prepareFile(
      await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }),
      "bomb.pptx",
    );
    await expect(
      readMicrosoft365Source("u", "c", db, file),
    ).rejects.toMatchObject({ code: "source_too_large" });
  });
  it("rejects unknown extensions before downloading", async () => {
    graph.mockResolvedValueOnce(response(metadata("test.exe")));
    await expect(
      readMicrosoft365Source("u", "c", db, file),
    ).rejects.toMatchObject({ code: "unsupported_source" });
    expect(graph).toHaveBeenCalledTimes(1);
  });
  it("does not silently truncate long text", async () => {
    prepareFile("a".repeat(100001));
    await expect(
      readMicrosoft365Source("u", "c", db, file),
    ).rejects.toMatchObject({ code: "source_too_large" });
  });
});

describe("parser worker admission", () => {
  function mockFileRequests() {
    graph.mockImplementation(async (_user, _connection, _db, path) =>
      path.endsWith("/content")
        ? new Response("content")
        : response(metadata()),
    );
  }
  function controlledWorker() {
    const worker = new EventEmitter();
    let stop!: (code: number) => void;
    const termination = new Promise<number>((resolve) => {
      stop = resolve;
    });
    return Object.assign(worker, { terminate: vi.fn(() => termination), stop });
  }
  it("caps active workers at two, keeps terminating/cancelled slots, and releases on exit", async () => {
    mockFileRequests();
    const workers: ReturnType<typeof controlledWorker>[] = [];
    workerMock.create = () => {
      const worker = controlledWorker();
      workers.push(worker);
      return worker;
    };
    const controller = new AbortController();
    const first = readMicrosoft365Source("u", "c", db, file);
    const second = readMicrosoft365Source(
      "u",
      "c",
      db,
      file,
      controller.signal,
    );
    await vi.waitFor(() => expect(workers).toHaveLength(2));
    workers[0].emit("message", { text: "first" });
    await expect(
      readMicrosoft365Source("u", "c", db, file),
    ).rejects.toMatchObject({
      code: "busy",
      status: 409,
      retryAfterSeconds: 1,
    });
    expect(workers).toHaveLength(2);
    workers[0].stop(0);
    expect((await first).text).toBe("first");
    const third = readMicrosoft365Source("u", "c", db, file);
    await vi.waitFor(() => expect(workers).toHaveLength(3));
    controller.abort();
    expect(workers[1].terminate).toHaveBeenCalledTimes(1);
    await expect(
      readMicrosoft365Source("u", "c", db, file),
    ).rejects.toMatchObject({ code: "busy" });
    const secondFailure = expect(second).rejects.toMatchObject({
      code: "provider_unavailable",
    });
    workers[1].stop(0);
    await secondFailure;
    // An exit itself releases capacity even if terminate's Promise is still pending.
    workers[2].emit("exit", 1);
    const fourth = readMicrosoft365Source("u", "c", db, file);
    await vi.waitFor(() => expect(workers).toHaveLength(4));
    const thirdFailure = expect(third).rejects.toMatchObject({
      code: "unsupported_source",
    });
    workers[2].stop(1);
    await thirdFailure;
    workers[3].emit("message", { text: "fourth" });
    workers[3].stop(0);
    expect((await fourth).text).toBe("fourth");
  });
  it("releases reservations after constructor failures", async () => {
    mockFileRequests();
    workerMock.create = () => {
      throw new Error("Cannot spawn");
    };
    for (let i = 0; i < 3; i++) {
      await expect(
        readMicrosoft365Source("u", "c", db, file),
      ).rejects.toMatchObject({ code: "provider_unavailable", status: 502 });
    }
    workerMock.create = null;
    expect((await readMicrosoft365Source("u", "c", db, file)).text).toBe(
      "content",
    );
  });
});

describe("recent mail folder metadata", () => {
  it.each([
    ["sentitems", "/me/mailFolders/sentitems/messages", "sentDateTime desc"],
    ["inbox", "/me/mailFolders/inbox/messages", "receivedDateTime desc"],
    ["all", "/me/messages", "receivedDateTime desc"],
  ] as const)("lists %s with the requested scope and ordering", async (folder, path, order) => {
    graph.mockResolvedValueOnce(response({ value: [{ id: "immutable-sent", subject: "Subject", bodyPreview: "private body" }] }));
    const result = await searchMicrosoft365Sources("u", "c", db, {
      kind: "mail", query: "", mode: "recent", folder, offset: 20,
    });
    const url = new URL(`https://graph.microsoft.com${graph.mock.calls[0][3]}`);
    expect(url.pathname).toBe(path);
    expect(url.searchParams.get("$orderby")).toBe(order);
    expect(url.searchParams.get("$select")).toBe("id,subject");
    expect(url.searchParams.get("$skip")).toBe("20");
    expect(graph.mock.calls[0][4]?.headers).toEqual({ Prefer: 'IdType="ImmutableId"' });
    expect(result).toEqual({ items: [{ locator: { kind: "mail", id: "immutable-sent" }, title: "Subject" }], more: false, nextOffset: 40 });
    expect(JSON.stringify(result)).not.toContain("private body");
  });
  it("lists newest immutable message IDs without reading message bodies or calling Search", async () => {
    graph.mockResolvedValueOnce(
      response({
        value: Array.from({ length: 21 }, (_, i) => ({
          id: `immutable-${i}`,
          subject: `Message ${i}`,
          bodyPreview: "must not escape",
        })),
      }),
    );
    const result = await searchMicrosoft365Sources("u", "c", db, {
      kind: "mail",
      query: "",
      mode: "recent",
      offset: 20,
    });
    const path = graph.mock.calls[0][3];
    expect(path).toContain("/me/mailFolders/inbox/messages?");
    const params = new URL(`https://graph.microsoft.com${path}`).searchParams;
    expect(params.get("$select")).toBe("id,subject");
    expect(params.get("$orderby")).toBe("receivedDateTime desc");
    expect(params.get("$skip")).toBe("20");
    expect(graph.mock.calls[0][4]?.headers).toEqual({
      Prefer: 'IdType="ImmutableId"',
    });
    expect(result.items).toHaveLength(20);
    expect(result.more).toBe(true);
    expect(result.nextOffset).toBe(40);
    expect(JSON.stringify(result)).not.toContain("must not escape");
  });
  it("reports empty inbox distinctly and never returns malformed provider IDs", async () => {
    graph.mockResolvedValueOnce(response({ value: [] }));
    expect(
      await searchMicrosoft365Sources("u", "c", db, {
        kind: "mail",
        query: "",
        mode: "recent",
      }),
    ).toEqual({ items: [], more: false, nextOffset: 20 });
    graph.mockResolvedValueOnce(
      response({ value: [{ id: "../users/other" }] }),
    );
    await expect(
      searchMicrosoft365Sources("u", "c", db, {
        kind: "mail",
        query: "",
        mode: "recent",
      }),
    ).rejects.toMatchObject({ code: "provider_unavailable" });
  });
  it.each([
    { kind: "file", query: "", mode: "recent" },
    { kind: "mail", query: "ignored query", mode: "recent" },
    { kind: "mail", query: "", mode: "search" },
    { kind: "mail", query: "", mode: "recent", folder: "../other" },
    { kind: "mail", query: "contract", mode: "search", folder: "sentitems" },
    { kind: "file", query: "contract", folder: "all" },
  ])("rejects ambiguous recent mode %j", async (input) => {
    await expect(
      searchMicrosoft365Sources("u", "c", db, input as any),
    ).rejects.toMatchObject({ code: "invalid_query" });
    expect(graph).not.toHaveBeenCalled();
  });
});


describe("Graph Search empty results observed on the live tenant", () => {
  it.each(["mail", "file"] as const)("accepts omitted hits only for confirmed empty %s results", async (kind) => {
    graph.mockResolvedValueOnce(response({ value: [{ hitsContainers: [{ total: 0, moreResultsAvailable: false }] }] }));
    await expect(searchMicrosoft365Sources("u", "c", db, { kind, query: "test" })).resolves.toEqual({ items: [], more: false, nextOffset: 20 });
    expect(graph).toHaveBeenCalledTimes(1);
  });
  it.each([
    { total: 1, moreResultsAvailable: false },
    { total: 0, moreResultsAvailable: true },
    { total: 0 },
    { total: 0, moreResultsAvailable: false, hits: null },
  ])("rejects malformed or incomplete hit containers %j", async (container) => {
    graph.mockResolvedValueOnce(response({ value: [{ hitsContainers: [container] }] }));
    await expect(searchMicrosoft365Sources("u", "c", db, { kind: "file", query: "test" })).rejects.toMatchObject({ code: "provider_unavailable" });
  });
});
