import type { createServerSupabase } from "../supabase";
import {
  getMicrosoft365Config,
  rotateMicrosoft365ConnectionEncryption,
} from "./index";
import {
  cleanupMicrosoft365Chats,
  rotateMicrosoft365ChatEncryption,
} from "./protected";
import { enqueueDbJob } from "../dbq/enqueue";
import type { DbJob } from "../dbq/types";
/** No corporate content or credentials enter queue payloads or audit results. */
export async function maintainMicrosoft365(
  db: ReturnType<typeof createServerSupabase>,
  job: DbJob,
) {
  const removed = await cleanupMicrosoft365Chats(db);
  if (!getMicrosoft365Config()) return { removed };
  const connectionCursor =
    typeof job.payload.connectionCursor === "string"
      ? job.payload.connectionCursor
      : undefined;
  const chatCursor =
    typeof job.payload.chatCursor === "string"
      ? job.payload.chatCursor
      : undefined;
  const connections =
    job.payload.connectionsDone === true
      ? null
      : await rotateMicrosoft365ConnectionEncryption(db, connectionCursor);
  const chats =
    job.payload.chatsDone === true
      ? null
      : await rotateMicrosoft365ChatEncryption(db, chatCursor);
  if (connections?.nextCursor || chats?.nextCursor) {
    await enqueueDbJob(db, {
      kind: "microsoft365.maintenance",
      payload: {
        connectionCursor: connections?.nextCursor,
        chatCursor: chats?.nextCursor,
        connectionsDone: !connections?.nextCursor,
        chatsDone: !chats?.nextCursor,
      },
      dedupeKey: `microsoft365.maintenance:${connections?.nextCursor ?? "done"}:${chats?.nextCursor ?? "done"}`,
      maxAttempts: 3,
    });
  }
  return {
    removed,
    rotated: (connections?.rotated ?? 0) + (chats?.rotated ?? 0),
    skipped: (connections?.skipped ?? 0) + (chats?.skipped ?? 0),
  };
}
