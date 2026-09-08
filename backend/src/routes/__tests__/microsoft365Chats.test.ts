import express from "express";
import request from "supertest";
import { randomUUID } from "node:crypto";
import { once } from "node:events";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

const calls = vi.hoisted(() => ({
  status: vi.fn(),
  start: vi.fn(),
  complete: vi.fn(),
  check: vi.fn(),
  disconnect: vi.fn(),
  policy: vi.fn(),
  list: vi.fn(),
  create: vi.fn(),
  delete: vi.fn(),
  get: vi.fn(),
  search: vi.fn(),
  send: vi.fn(),
  view: vi.fn(),
}));
vi.mock("../../lib/supabase", () => ({
  createServerSupabase: () => ({ fixture: true }),
}));
vi.mock("../../middleware/auth", () => ({
  requireAuth: (
    req: express.Request,
    res: express.Response,
    next: express.NextFunction,
  ) => {
    if (!req.headers["x-test-user"])
      return res.status(401).json({ code: "unauthenticated" });
    res.locals.userId = req.headers["x-test-user"];
    res.locals.token = "verified-session";
    next();
  },
  requireMfaIfEnrolled: (
    req: express.Request,
    res: express.Response,
    next: express.NextFunction,
  ) => {
    if (req.headers["x-test-mfa"] === "required")
      return res.status(403).json({ code: "mfa_verification_required" });
    next();
  },
}));
vi.mock("../../lib/microsoft365", async (importOriginal) => ({
  ...(await importOriginal<typeof import("../../lib/microsoft365")>()),
  getMicrosoft365Status: calls.status,
  startMicrosoft365Connection: calls.start,
  completeMicrosoft365Connection: calls.complete,
  checkMicrosoft365Connection: calls.check,
  disconnectMicrosoft365Connection: calls.disconnect,
}));
vi.mock("../../lib/microsoft365/protected", () => ({
  getMicrosoft365ChatPolicy: calls.policy,
  listMicrosoft365Chats: calls.list,
  createMicrosoft365Chat: calls.create,
  deleteMicrosoft365Chat: calls.delete,
}));
vi.mock("../../lib/microsoft365/chat", () => ({
  getProtectedMicrosoft365Chat: calls.get,
  searchProtectedMicrosoft365Chat: calls.search,
  sendProtectedMicrosoft365Message: calls.send,
  viewProtectedMicrosoft365Source: calls.view,
}));
import { microsoft365Router } from "../microsoft365";
import { Microsoft365Error } from "../../lib/microsoft365";

const app = express();
app.use(express.json());
app.use("/integrations/microsoft365", microsoft365Router);
const base = "/integrations/microsoft365/chats";
const chatId = "33333333-3333-4333-8333-333333333333";
const connectionId = "44444444-4444-4444-8444-444444444444";
const sourceRef = "55555555-5555-4555-8555-555555555555";
let owner: string;
const policy = { model: "approved-model", retentionDays: 7 };
const endpoints = [
  ["get", ""],
  ["post", ""],
  ["get", `/${chatId}`],
  ["post", `/${chatId}/search`],
  ["post", `/${chatId}/messages`],
  ["get", `/${chatId}/sources/${sourceRef}`],
  ["delete", `/${chatId}`],
] as const;
function headers(response: request.Response) {
  expect(response.headers["cache-control"]).toBe("no-store");
  expect(response.headers["referrer-policy"]).toBe("no-referrer");
}

describe("protected Microsoft 365 chat routes", () => {
  beforeEach(() => {
    vi.resetAllMocks();
    owner = randomUUID(); // Keep real per-user rate limiter isolated across cases.
    vi.spyOn(console, "error").mockImplementation(() => {});
    calls.status.mockResolvedValue({
      available: true,
      connection: { id: connectionId, status: "connected" },
    });
    calls.policy.mockReturnValue(policy);
    calls.list.mockResolvedValue([{ id: chatId }]);
    calls.create.mockResolvedValue({ id: chatId });
    calls.delete.mockResolvedValue(undefined);
    calls.get.mockResolvedValue({ id: chatId, messages: [] });
    calls.search.mockResolvedValue({ sources: [], offset: 0 });
    calls.send.mockResolvedValue({ answer: "Protected answer", citations: [] });
    calls.view.mockResolvedValue({ title: "Source", text: "Source excerpt" });
  });
  afterEach(() => {
    vi.restoreAllMocks();
  });

  it.each(endpoints)(
    "requires authentication for %s %s",
    async (method, path) => {
      const response = await request(app)[method](`${base}${path}`);
      expect(response.status).toBe(401);
      headers(response);
      for (const fn of Object.values(calls)) expect(fn).not.toHaveBeenCalled();
    },
  );

  it.each(endpoints)("requires MFA for %s %s", async (method, path) => {
    const response = await request(app)
      [method](`${base}${path}`)
      .set("x-test-user", owner)
      .set("x-test-mfa", "required");
    expect(response.status).toBe(403);
    headers(response);
    for (const fn of Object.values(calls)) expect(fn).not.toHaveBeenCalled();
  });

  it("lists only the session user's connected account", async () => {
    const response = await request(app)
      .get(`${base}?userId=forged`)
      .set("x-test-user", owner);
    expect(response.status).toBe(200);
    headers(response);
    expect(response.body).toEqual({
      available: true,
      chats: [{ id: chatId }],
      policy,
    });
    expect(calls.status).toHaveBeenCalledWith(owner, { fixture: true });
    expect(calls.list).toHaveBeenCalledWith(owner, connectionId, {
      fixture: true,
    });
  });

  it.each(["disabled", "disconnected", "no-policy"])(
    "does not retrieve histories when %s",
    async (condition) => {
      if (condition === "disabled")
        calls.status.mockResolvedValue({ available: false, connection: null });
      if (condition === "disconnected")
        calls.status.mockResolvedValue({
          available: true,
          connection: { id: connectionId, status: "reconnect_required" },
        });
      if (condition === "no-policy") calls.policy.mockReturnValue(null);
      const response = await request(app).get(base).set("x-test-user", owner);
      expect(response.status).toBe(200);
      expect(response.body.available).toBe(false);
      expect(response.body.chats).toEqual([]);
      expect(calls.list).not.toHaveBeenCalled();
    },
  );

  it.each([
    "userId",
    "model",
    "url",
    "locator",
    "tools",
    "runTools",
    "apiKeys",
    "systemPrompt",
  ])("rejects forged %s in all mutation payloads", async (field) => {
    for (const [path, body] of [
      ["", { connectionId }],
      [`/${chatId}/search`, { kind: "mail", query: "Quarterly" }],
      [
        `/${chatId}/messages`,
        { message: "Summarize", sourceRefs: [sourceRef] },
      ],
    ] as const) {
      const response = await request(app)
        .post(`${base}${path}`)
        .set("x-test-user", owner)
        .send({ ...body, [field]: "forged-value" });
      expect(response.status).toBe(400);
      expect(response.body.code).toBe("invalid_query");
      headers(response);
    }
    expect(calls.create).not.toHaveBeenCalled();
    expect(calls.search).not.toHaveBeenCalled();
    expect(calls.send).not.toHaveBeenCalled();
  });

  it("passes validated inputs with session ownership and a cancellation signal", async () => {
    const created = await request(app)
      .post(base)
      .set("x-test-user", owner)
      .send({ connectionId });
    expect(created.status).toBe(201);
    expect(created.body).toEqual({ chatId });
    expect(calls.create).toHaveBeenCalledWith(owner, connectionId, {
      fixture: true,
    });
    const loaded = await request(app)
      .get(`${base}/${chatId}?userId=forged`)
      .set("x-test-user", owner);
    expect(loaded.status).toBe(200);
    expect(calls.get).toHaveBeenCalledWith(
      owner,
      chatId,
      { fixture: true },
      expect.any(AbortSignal),
    );
    const searched = await request(app)
      .post(`${base}/${chatId}/search`)
      .set("x-test-user", owner)
      .send({ kind: "file", query: " Quarterly ", offset: 25 });
    expect(searched.status).toBe(200);
    expect(calls.search).toHaveBeenCalledWith(
      owner,
      chatId,
      { fixture: true },
      { kind: "file", query: "Quarterly", offset: 25 },
      expect.any(AbortSignal),
    );
    const sent = await request(app)
      .post(`${base}/${chatId}/messages`)
      .set("x-test-user", owner)
      .send({ message: " Summarize ", sourceRefs: [sourceRef] });
    expect(sent.status).toBe(200);
    expect(calls.send).toHaveBeenCalledWith(
      owner,
      chatId,
      { fixture: true },
      { message: "Summarize", sourceRefs: [sourceRef] },
      expect.any(AbortSignal),
    );
    const viewed = await request(app)
      .get(`${base}/${chatId}/sources/${sourceRef}`)
      .set("x-test-user", owner);
    expect(viewed.status).toBe(200);
    expect(calls.view).toHaveBeenCalledWith(
      owner,
      chatId,
      sourceRef,
      { fixture: true },
      expect.any(AbortSignal),
    );
    const deleted = await request(app)
      .delete(`${base}/${chatId}`)
      .set("x-test-user", owner);
    expect(deleted.status).toBe(204);
    expect(calls.delete).toHaveBeenCalledWith(owner, chatId, { fixture: true });
    for (const response of [created, loaded, searched, sent, viewed, deleted])
      headers(response);
    for (const fn of [calls.start, calls.check, calls.disconnect])
      expect(fn).not.toHaveBeenCalled();
  });

  it("validates source handles, input limits and pagination before service IO", async () => {
    for (const body of [
      { message: "", sourceRefs: [] },
      { message: "x".repeat(8001), sourceRefs: [] },
      { message: "Summarize", sourceRefs: ["https://attacker.test/source"] },
      { message: "Summarize", sourceRefs: Array(6).fill(sourceRef) },
    ])
      expect(
        (
          await request(app)
            .post(`${base}/${chatId}/messages`)
            .set("x-test-user", owner)
            .send(body)
        ).status,
      ).toBe(400);
    for (const body of [
      { kind: "shared-mailbox", query: "x" },
      { kind: "mail", query: " " },
      { kind: "file", query: "x".repeat(501) },
      { kind: "file", query: "x", offset: -1 },
      { kind: "file", query: "x", offset: 1001 },
      { kind: "file", query: "x", offset: 1.5 },
    ])
      expect(
        (
          await request(app)
            .post(`${base}/${chatId}/search`)
            .set("x-test-user", owner)
            .send(body)
        ).status,
      ).toBe(400);
    expect(
      (await request(app).get(`${base}/not-uuid`).set("x-test-user", owner))
        .status,
    ).toBe(400);
    expect(
      (
        await request(app)
          .get(`${base}/${chatId}/sources/not-uuid`)
          .set("x-test-user", owner)
      ).status,
    ).toBe(400);
    for (const fn of [calls.send, calls.search, calls.get, calls.view])
      expect(fn).not.toHaveBeenCalled();
  });

  it("never returns or logs raw provider failures", async () => {
    calls.send.mockRejectedValue(new Error("SECRET_TOKEN PRIVATE_EMAIL BODY"));
    const response = await request(app)
      .post(`${base}/${chatId}/messages`)
      .set("x-test-user", owner)
      .send({ message: "Summarize", sourceRefs: [sourceRef] });
    expect(response.status).toBe(502);
    expect(response.body.code).toBe("provider_unavailable");
    headers(response);
    expect(
      JSON.stringify({
        response: response.body,
        logs: vi.mocked(console.error).mock.calls,
      }),
    ).not.toMatch(/SECRET_TOKEN|PRIVATE_EMAIL/);
  });

  it("preserves safe permission failures and Microsoft Retry-After", async () => {
    calls.get.mockRejectedValueOnce(
      new Microsoft365Error("rate_limited", 429, 45),
    );
    const limited = await request(app)
      .get(`${base}/${chatId}`)
      .set("x-test-user", owner);
    expect(limited.status).toBe(429);
    expect(limited.headers["retry-after"]).toBe("45");
    expect(limited.body.code).toBe("rate_limited");
    calls.get.mockRejectedValueOnce(
      new Microsoft365Error("access_denied", 403),
    );
    const denied = await request(app)
      .get(`${base}/${chatId}`)
      .set("x-test-user", owner);
    expect(denied.status).toBe(403);
    expect(denied.body.code).toBe("access_denied");
    expect(console.error).not.toHaveBeenCalled();
  });

  it("limits repeated requests per authenticated user", async () => {
    for (let index = 0; index < 30; index++)
      expect(
        (await request(app).get(base).set("x-test-user", owner)).status,
      ).toBe(200);
    const limited = await request(app).get(base).set("x-test-user", owner);
    expect(limited.status).toBe(429);
    expect(limited.body.code).toBe("rate_limited");
    expect(Number(limited.headers["retry-after"])).toBeGreaterThan(0);
    expect(calls.list).toHaveBeenCalledTimes(30);
    expect(
      (await request(app).get(base).set("x-test-user", randomUUID())).status,
    ).toBe(200);
  });

  it("does not interpret protected chat paths as connection operations", async () => {
    expect(
      (
        await request(app)
          .post(`${base}/${chatId}/check`)
          .set("x-test-user", owner)
      ).status,
    ).toBe(404);
    expect(
      (await request(app).delete(`${base}/${chatId}`).set("x-test-user", owner))
        .status,
    ).toBe(204);
    expect(calls.check).not.toHaveBeenCalled();
    expect(calls.disconnect).not.toHaveBeenCalled();
  });

  it("aborts service IO when the HTTP client disconnects", async () => {
    let reached!: (signal: AbortSignal) => void;
    const started = new Promise<AbortSignal>((resolve) => {
      reached = resolve;
    });
    calls.send.mockImplementation(
      (_user, _chat, _db, _input, signal: AbortSignal) =>
        new Promise((_resolve, reject) => {
          signal.addEventListener(
            "abort",
            () => reject(new Error("aborted private request")),
            { once: true },
          );
          reached(signal);
        }),
    );
    const server = app.listen(0, "127.0.0.1");
    await once(server, "listening");
    const address = server.address();
    if (!address || typeof address === "string")
      throw new Error("Missing test listener");
    const controller = new AbortController();
    try {
      const pending = fetch(
        `http://127.0.0.1:${address.port}${base}/${chatId}/messages`,
        {
          method: "POST",
          headers: { "content-type": "application/json", "x-test-user": owner },
          body: JSON.stringify({ message: "Summarize", sourceRefs: [] }),
          signal: controller.signal,
        },
      ).catch((error: unknown) => error);
      const signal = await started;
      expect(signal.aborted).toBe(false);
      const aborted = once(signal, "abort");
      controller.abort();
      await aborted;
      expect(signal.aborted).toBe(true);
      await pending;
      expect(console.error).not.toHaveBeenCalled();
    } finally {
      controller.abort();
      server.closeAllConnections();
      await new Promise<void>((resolve) => server.close(() => resolve()));
    }
  });
  it("keeps protected messages out of the ordinary account chat export", async () => {
    const { buildUserChatsExport } = await import("../../lib/userDataExport");
    const from = vi.fn((table: string) => {
      const rows = table.startsWith("microsoft365_")
        ? [{ id: chatId, body: "PRIVATE_CORPORATE_BODY" }]
        : table === "chats"
          ? [{ id: "ordinary-chat", user_id: owner }]
          : [];
      const query = {
        select: vi.fn(() => query),
        range: vi.fn(() => query),
        eq: vi.fn(() => query),
        order: vi.fn(() => query),
        in: vi.fn(() => query),
        then: (resolve: (result: unknown) => unknown) =>
          Promise.resolve({ data: rows, error: null }).then(resolve),
      };
      return query;
    });
    const db = { from } as unknown as Parameters<
      typeof buildUserChatsExport
    >[0];
    const result = await buildUserChatsExport(db, owner);
    expect(result.assistant_chats.chats).toEqual([
      { id: "ordinary-chat", user_id: owner },
    ]);
    expect(JSON.stringify(result)).not.toContain("PRIVATE_CORPORATE_BODY");
    expect(from.mock.calls.map(([table]) => table)).not.toEqual(
      expect.arrayContaining([expect.stringMatching(/^microsoft365_/)]),
    );
  });
});
