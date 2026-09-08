import { Router, type Response } from "express";
import { requireAuth, requireMfaIfEnrolled } from "../middleware/auth";
import { createServerSupabase } from "../lib/supabase";
import { INTERNAL_ERROR_CODE, INTERNAL_ERROR_MESSAGE } from "../lib/httpError";
import {
  Microsoft365Error,
  getMicrosoft365Status,
  startMicrosoft365Connection,
  completeMicrosoft365Connection,
  checkMicrosoft365Connection,
  disconnectMicrosoft365Connection,
} from "../lib/microsoft365";

import { microsoft365ChatsRouter } from "./microsoft365Chats";

export const microsoft365Router = Router();

// OAuth query parameters must not reach generic middleware error logs. They
// remain available in req.query for the one-time callback exchange.
microsoft365Router.use((req, res, next) => {
  req.originalUrl = `${req.baseUrl}${req.path}`;
  res.set("Cache-Control", "no-store");
  res.set("Referrer-Policy", "no-referrer");
  next();
});
microsoft365Router.use(requireAuth);
microsoft365Router.use("/chats", microsoft365ChatsRouter);

function fail(res: Response, error: unknown, operation: string) {
  if (error instanceof Microsoft365Error) {
    if (error.retryAfterSeconds !== undefined) {
      res.set("Retry-After", String(error.retryAfterSeconds));
    }
    return res
      .status(error.status)
      .json({ code: error.code, detail: error.detail });
  }
  // The existing sendInternalError helper logs the supplied error. Never pass
  // provider exceptions or database rows containing credentials to that logger.
  console.error("[microsoft365] operation failed", { operation });
  return res
    .status(500)
    .json({ code: INTERNAL_ERROR_CODE, detail: INTERNAL_ERROR_MESSAGE });
}

function settingsReturnUrl(success: boolean): string {
  const configured = process.env.FRONTEND_URL?.trim();
  if (!configured && process.env.NODE_ENV === "production")
    throw new Error("Invalid frontend configuration");
  const url = new URL(configured || "http://localhost:3000");
  if (
    url.username ||
    url.password ||
    (process.env.NODE_ENV === "production"
      ? url.protocol !== "https:"
      : !["http:", "https:"].includes(url.protocol))
  ) {
    throw new Error("Invalid frontend configuration");
  }
  const target = new URL("/settings/connectors", url);
  target.searchParams.set("microsoft365", success ? "connected" : "error");
  return target.href;
}

const uuid = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

microsoft365Router.get("/", async (_req, res) => {
  try {
    res.json(
      await getMicrosoft365Status(res.locals.userId, createServerSupabase()),
    );
  } catch (error) {
    fail(res, error, "status");
  }
});

microsoft365Router.post("/connect", requireMfaIfEnrolled, async (_req, res) => {
  try {
    res.json(
      await startMicrosoft365Connection(
        res.locals.userId,
        res.locals.token,
        createServerSupabase(),
      ),
    );
  } catch (error) {
    fail(res, error, "connect");
  }
});

microsoft365Router.get("/callback", async (req, res) => {
  let success = false;
  try {
    const { state, code, error } = req.query;
    if (
      typeof state !== "string" ||
      state.length < 32 ||
      state.length > 256 ||
      (error === undefined
        ? typeof code !== "string" || !code || code.length > 8192
        : typeof error !== "string" || !error || code !== undefined)
    ) {
      throw new Error("Invalid callback");
    }
    await completeMicrosoft365Connection(
      res.locals.userId,
      res.locals.token,
      {
        state,
        code: typeof code === "string" ? code : "",
        ...(typeof error === "string" ? { error } : {}),
      },
      createServerSupabase(),
    );
    success = true;
  } catch {
    // Do not log OAuth codes/state, provider error_description, or raw errors.
    console.error("[microsoft365] callback could not be completed");
  }
  try {
    res.redirect(303, settingsReturnUrl(success));
  } catch (error) {
    fail(res, error, "callback_configuration");
  }
});

microsoft365Router.post(
  "/:connectionId/check",
  requireMfaIfEnrolled,
  async (req, res) => {
    if (!uuid.test(req.params.connectionId)) {
      return res
        .status(400)
        .json({ code: "invalid_connection_id", detail: "Invalid connection." });
    }
    try {
      res.json(
        await checkMicrosoft365Connection(
          res.locals.userId,
          req.params.connectionId,
          createServerSupabase(),
        ),
      );
    } catch (error) {
      fail(res, error, "check");
    }
  },
);

microsoft365Router.delete(
  "/:connectionId",
  requireMfaIfEnrolled,
  async (req, res) => {
    if (!uuid.test(req.params.connectionId)) {
      return res
        .status(400)
        .json({ code: "invalid_connection_id", detail: "Invalid connection." });
    }
    try {
      await disconnectMicrosoft365Connection(
        res.locals.userId,
        req.params.connectionId,
        createServerSupabase(),
      );
      res.status(204).send();
    } catch (error) {
      fail(res, error, "disconnect");
    }
  },
);
