import express from "express";
import request from "supertest";
import { describe, it, expect, vi } from "vitest";
import { microsoft365JsonBody } from "./microsoft365Body";
describe("protected body parser", () => {
  it("never forwards malformed corporate body or parser exceptions to generic logging", async () => {
    const app = express();
    const errorLog = vi.fn();
    app.use(microsoft365JsonBody);
    app.post("/", (_req, res) => res.json({ ok: true }));
    app.use(((
      error: unknown,
      _req: express.Request,
      res: express.Response,
      _next: express.NextFunction,
    ) => {
      errorLog(error);
      res.status(500).end();
    }) as express.ErrorRequestHandler);
    const response = await request(app)
      .post("/")
      .type("json")
      .send('{"secret-corporate-text":');
    expect(response.status).toBe(400);
    expect(response.headers["cache-control"]).toBe("no-store");
    expect(response.text).not.toContain("secret-corporate-text");
    expect(errorLog).not.toHaveBeenCalled();
  });
  it("caps input before the global large body parser but accepts legitimate UTF8 questions", async () => {
    const app = express();
    app.use(microsoft365JsonBody);
    app.use(express.json({ limit: "50mb" }));
    app.post("/", (_req, res) => res.json({ ok: true }));
    expect(
      (
        await request(app)
          .post("/")
          .send({ message: "x".repeat(66000) })
      ).status,
    ).toBe(413);
    expect(
      (
        await request(app)
          .post("/")
          .send({ message: "я".repeat(8000) })
      ).status,
    ).toBe(200);
  });
});
