import express from "express";
import request from "supertest";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

const calls = vi.hoisted(() => ({
  status: vi.fn(),
  start: vi.fn(),
  complete: vi.fn(),
  check: vi.fn(),
  disconnect: vi.fn(),
  originalUrls: [] as string[],
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
    calls.originalUrls.push(req.originalUrl);
    if (!req.headers["x-test-user"])
      return res.status(401).json({ detail: "Sign in required" });
    res.locals.userId = req.headers["x-test-user"];
    res.locals.token = "verified-mike-session";
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
vi.mock("../../lib/microsoft365", () => ({
  Microsoft365Error: class extends Error {},
  getMicrosoft365Status: calls.status,
  startMicrosoft365Connection: calls.start,
  completeMicrosoft365Connection: calls.complete,
  checkMicrosoft365Connection: calls.check,
  disconnectMicrosoft365Connection: calls.disconnect,
}));
import { microsoft365Router } from "../microsoft365";
import { Microsoft365Error } from "../../lib/microsoft365";

const app = express();
app.use(express.json());
app.use("/integrations/microsoft365", microsoft365Router);
const base = "/integrations/microsoft365";
const id = "33333333-3333-4333-8333-333333333333";
const state = "s".repeat(43);

describe("Microsoft 365 connection routes", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    calls.originalUrls.length = 0;
    vi.stubEnv("FRONTEND_URL", "https://mike.example.test");
    vi.stubEnv("NODE_ENV", "production");
    vi.spyOn(console, "error").mockImplementation(() => {});
  });
  afterEach(() => {
    vi.unstubAllEnvs();
    vi.restoreAllMocks();
  });

  it("requires an authenticated Mike session on every route including callback", async () => {
    for (const result of await Promise.all([
      request(app).get(base),
      request(app).post(`${base}/connect`),
      request(app).get(`${base}/callback?state=${state}&code=secret-code`),
      request(app).post(`${base}/${id}/check`),
      request(app).delete(`${base}/${id}`),
    ]))
      expect(result.status).toBe(401);
    for (const fn of [
      calls.status,
      calls.start,
      calls.complete,
      calls.check,
      calls.disconnect,
    ])
      expect(fn).not.toHaveBeenCalled();
    expect(calls.originalUrls.join(" ")).not.toContain("secret-code");
  });

  it("resolves status for the session user and never accepts a caller-selected owner", async () => {
    calls.status.mockResolvedValue({ available: false, connection: null });
    const result = await request(app)
      .get(`${base}?userId=other-user`)
      .set("x-test-user", "user-1");
    expect(result.status).toBe(200);
    expect(result.body).toEqual({ available: false, connection: null });
    expect(calls.status).toHaveBeenCalledWith("user-1", { fixture: true });
    expect(result.headers["cache-control"]).toBe("no-store");
  });

  it("enforces MFA before connect, check, and disconnect", async () => {
    for (const result of await Promise.all([
      request(app)
        .post(`${base}/connect`)
        .set("x-test-user", "u")
        .set("x-test-mfa", "required"),
      request(app)
        .post(`${base}/${id}/check`)
        .set("x-test-user", "u")
        .set("x-test-mfa", "required"),
      request(app)
        .delete(`${base}/${id}`)
        .set("x-test-user", "u")
        .set("x-test-mfa", "required"),
    ]))
      expect(result.status).toBe(403);
    expect(calls.start).not.toHaveBeenCalled();
    expect(calls.check).not.toHaveBeenCalled();
    expect(calls.disconnect).not.toHaveBeenCalled();
  });

  it("binds connect to the authenticated session", async () => {
    calls.start.mockResolvedValue({
      authorizationUrl: "https://login.microsoftonline.com/authorize",
      callbackOrigin: "https://mike.example.test",
    });
    const result = await request(app)
      .post(`${base}/connect`)
      .set("x-test-user", "u")
      .send({ userId: "other", token: "forged" });
    expect(result.status).toBe(200);
    expect(calls.start).toHaveBeenCalledWith("u", "verified-mike-session", {
      fixture: true,
    });
  });

  it("consumes callback with the session owner and redirects without exposing OAuth values", async () => {
    calls.complete.mockResolvedValue(undefined);
    const result = await request(app)
      .get(`${base}/callback?state=${state}&code=private-code`)
      .set("x-test-user", "u");
    expect(calls.complete).toHaveBeenCalledWith(
      "u",
      "verified-mike-session",
      { state, code: "private-code" },
      { fixture: true },
    );
    expect(result.status).toBe(303);
    expect(result.headers.location).toBe(
      "https://mike.example.test/settings/connectors?microsoft365=connected",
    );
    expect(result.headers["referrer-policy"]).toBe("no-referrer");
    expect(calls.originalUrls).toEqual([`${base}/callback`]);
    expect(result.text).not.toContain("private-code");
  });

  it("rejects ambiguous callbacks before exchanging a code", async () => {
    for (const query of [
      `state=${state}&state=other&code=c`,
      `state=${state}&code=c&code=d`,
      `state=${state}&code=c&error=denied`,
      "code=c",
      `state=${state}&error=a&error=b`,
    ]) {
      const result = await request(app)
        .get(`${base}/callback?${query}`)
        .set("x-test-user", "u");
      expect(result.status).toBe(303);
      expect(result.headers.location).toBe(
        "https://mike.example.test/settings/connectors?microsoft365=error",
      );
    }
    expect(calls.complete).not.toHaveBeenCalled();
  });

  it("passes a cancellation through state consumption but never reflects provider messages", async () => {
    calls.complete.mockRejectedValue(new Error("PRIVATE_SECRET"));
    const result = await request(app)
      .get(
        `${base}/callback?state=${state}&error=access_denied&error_description=PRIVATE_SECRET`,
      )
      .set("x-test-user", "u");
    expect(calls.complete).toHaveBeenCalledWith(
      "u",
      "verified-mike-session",
      { state, code: "", error: "access_denied" },
      { fixture: true },
    );
    expect(result.headers.location).toMatch(/microsoft365=error$/);
    expect(
      JSON.stringify({
        body: result.body,
        headers: result.headers,
        logs: vi.mocked(console.error).mock.calls,
      }),
    ).not.toContain("PRIVATE_SECRET");
  });

  it("checks or deletes only the requested connection for the current user", async () => {
    calls.check.mockResolvedValue({
      ok: true,
      mailAccess: true,
      checkedAt: "2026-09-07T00:00:00.000Z",
    });
    calls.disconnect.mockResolvedValue(undefined);
    const checked = await request(app)
      .post(`${base}/${id}/check`)
      .set("x-test-user", "u");
    const deleted = await request(app)
      .delete(`${base}/${id}`)
      .set("x-test-user", "u");
    expect(checked.status).toBe(200);
    expect(calls.check).toHaveBeenCalledWith("u", id, { fixture: true });
    expect(deleted.status).toBe(204);
    expect(calls.disconnect).toHaveBeenCalledWith("u", id, { fixture: true });
    expect(
      (
        await request(app)
          .post(`${base}/not-an-id/check`)
          .set("x-test-user", "u")
      ).status,
    ).toBe(400);
  });

  it("returns intentional errors with Retry-After and sanitizes unexpected failures", async () => {
    const error = Object.assign(Object.create(Microsoft365Error.prototype), {
      status: 429,
      code: "microsoft365_busy",
      detail: "Please try again shortly.",
      retryAfterSeconds: 30,
    });
    calls.check
      .mockRejectedValueOnce(error)
      .mockRejectedValueOnce(new Error("TOKEN_AND_MAIL_BODY"));
    const limited = await request(app)
      .post(`${base}/${id}/check`)
      .set("x-test-user", "u");
    expect(limited.status).toBe(429);
    expect(limited.headers["retry-after"]).toBe("30");
    const failed = await request(app)
      .post(`${base}/${id}/check`)
      .set("x-test-user", "u");
    expect(failed.status).toBe(500);
    expect(failed.body.code).toBe("internal_error");
    expect(
      JSON.stringify({
        body: failed.body,
        logs: vi.mocked(console.error).mock.calls,
      }),
    ).not.toContain("TOKEN_AND_MAIL_BODY");
  });

  it("does not redirect to an insecure or missing production frontend", async () => {
    calls.complete.mockResolvedValue(undefined);
    for (const origin of ["http://insecure.example.test", ""]) {
      vi.stubEnv("FRONTEND_URL", origin);
      const result = await request(app)
        .get(`${base}/callback?state=${state}&code=c`)
        .set("x-test-user", "u");
      expect(result.status).toBe(500);
      expect(result.headers.location).toBeUndefined();
    }
  });
});
