import {
  Router,
  type Request,
  type Response,
  type NextFunction,
} from "express";
import rateLimit from "express-rate-limit";
import { z } from "zod";
import { requireMfaIfEnrolled } from "../middleware/auth";
import { createServerSupabase } from "../lib/supabase";
import { getMicrosoft365Status, Microsoft365Error } from "../lib/microsoft365";
import {
  createMicrosoft365Chat,
  deleteMicrosoft365Chat,
  getMicrosoft365ChatPolicy,
  listMicrosoft365Chats,
} from "../lib/microsoft365/protected";
import {
  getProtectedMicrosoft365Chat,
  searchProtectedMicrosoft365Chat,
  sendProtectedMicrosoft365Message,
  viewProtectedMicrosoft365Source,
} from "../lib/microsoft365/chat";

/** Mounted after requireAuth on microsoft365Router; no route aliases in ordinary chats/exports. */
export const microsoft365ChatsRouter = Router();
microsoft365ChatsRouter.use(requireMfaIfEnrolled);
microsoft365ChatsRouter.use(
  rateLimit({
    windowMs: 60_000,
    limit: 30,
    standardHeaders: "draft-7",
    legacyHeaders: false,
    keyGenerator: (_req, res) => String(res.locals.userId),
    message: {
      code: "rate_limited",
      detail: "Please wait before trying again.",
    },
  }),
);
const uuid = z.string().uuid();
function route(
  handler: (req: Request, res: Response, signal: AbortSignal) => Promise<void>,
) {
  return async (req: Request, res: Response, _next: NextFunction) => {
    const controller = new AbortController();
    const cancel = () => {
      if (!res.writableEnded) controller.abort();
    };
    res.on("close", cancel);
    const signal = AbortSignal.any([
      controller.signal,
      AbortSignal.timeout(120_000),
    ]);
    try {
      await handler(req, res, signal);
    } catch (error) {
      if (res.destroyed || res.writableEnded) return;
      if (error instanceof Microsoft365Error) {
        if (error.retryAfterSeconds)
          res.set("Retry-After", String(error.retryAfterSeconds));
        res
          .status(error.status)
          .json({ code: error.code, detail: error.detail });
      } else if (error instanceof z.ZodError) {
        res
          .status(400)
          .json({
            code: "invalid_query",
            detail: "Check the request and source selection.",
          });
      } else {
        // Do not log model/provider errors: these may embed request bodies.
        console.error("[microsoft365] protected chat operation failed");
        res
          .status(502)
          .json({
            code: "provider_unavailable",
            detail:
              "The protected chat could not complete this request. Please try again.",
          });
      }
    } finally {
      res.off("close", cancel);
    }
  };
}
microsoft365ChatsRouter.get(
  "/",
  route(async (_req, res) => {
    const db = createServerSupabase();
    const status = await getMicrosoft365Status(res.locals.userId, db);
    const policy = getMicrosoft365ChatPolicy();
    const available =
      status.available && status.connection?.status === "connected" && !!policy;
    const chats = available
      ? await listMicrosoft365Chats(
          res.locals.userId,
          status.connection!.id,
          db,
        )
      : [];
    res.json({ available, chats, policy });
  }),
);
microsoft365ChatsRouter.post(
  "/",
  route(async (req, res) => {
    const input = z.object({ connectionId: uuid }).strict().parse(req.body);
    const chat = await createMicrosoft365Chat(
      res.locals.userId,
      input.connectionId,
      createServerSupabase(),
    );
    res.status(201).json({ chatId: chat.id });
  }),
);
microsoft365ChatsRouter.get(
  "/:chatId",
  route(async (req, res, signal) => {
    res.json(
      await getProtectedMicrosoft365Chat(
        res.locals.userId,
        uuid.parse(req.params.chatId),
        createServerSupabase(),
        signal,
      ),
    );
  }),
);
microsoft365ChatsRouter.post(
  "/:chatId/search",
  route(async (req, res, signal) => {
    const input = z
      .object({
        kind: z.enum(["mail", "file"]),
        query: z.string().trim().max(500),
        mode: z.enum(["search", "recent"]).optional(),
        offset: z.number().int().min(0).max(980).optional(),
        selectedSourceRefs: z.array(uuid).max(5).optional(),
      })
      .strict()
      .refine(input => input.mode === "recent" ? input.kind === "mail" && input.query === "" : input.query.length > 0)
      .parse(req.body);
    res.json(
      await searchProtectedMicrosoft365Chat(
        res.locals.userId,
        uuid.parse(req.params.chatId),
        createServerSupabase(),
        input,
        signal,
      ),
    );
  }),
);
microsoft365ChatsRouter.post(
  "/:chatId/messages",
  route(async (req, res, signal) => {
    const input = z
      .object({
        message: z.string().trim().min(1).max(8000),
        sourceRefs: z.array(uuid).max(5),
      })
      .strict()
      .parse(req.body);
    res.json(
      await sendProtectedMicrosoft365Message(
        res.locals.userId,
        uuid.parse(req.params.chatId),
        createServerSupabase(),
        input,
        signal,
      ),
    );
  }),
);
microsoft365ChatsRouter.get(
  "/:chatId/sources/:sourceRef",
  route(async (req, res, signal) => {
    res.json(
      await viewProtectedMicrosoft365Source(
        res.locals.userId,
        uuid.parse(req.params.chatId),
        uuid.parse(req.params.sourceRef),
        createServerSupabase(),
        signal,
      ),
    );
  }),
);
microsoft365ChatsRouter.delete(
  "/:chatId",
  route(async (req, res) => {
    await deleteMicrosoft365Chat(
      res.locals.userId,
      uuid.parse(req.params.chatId),
      createServerSupabase(),
    );
    res.status(204).send();
  }),
);
