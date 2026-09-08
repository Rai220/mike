import type { createServerSupabase } from "../supabase";
import { getMicrosoft365Status } from "./index";

export function microsoft365ChatRoutingInstruction(): string {
  return `
Microsoft 365 mail and corporate files are available to Mike through the Microsoft 365 toggle beside the message composer, in this same personal conversation.
This turn has Microsoft 365 access OFF. If the user asks to read, list, search, or discuss their mailbox or OneDrive/SharePoint files, explain that they can enable Microsoft 365 beside the composer and resend their question. Never redirect to a separate chat.
Do not claim Mike has no email capability or refuse merely because mail is outside legal work. Never invent mailbox contents or imply you read them with access disabled. Do not claim an account is connected without evidence.
Continue ordinary legal drafting, email templates, and work on documents already attached to this chat normally.
`;
}

/** Deliberately narrow: uncertainty falls through to the normal conversation. */
export function isMicrosoft365ReadRequest(text: string): boolean {
  if (text.length > 2_000) return false;
  const normalized = text.toLowerCase().replace(/ё/g, "е");
  // Drafting and explaining email content are ordinary legal-chat tasks.
  if (
    /(?:напиши|написать|составь|составить|подготовь|подготовить|шаблон|проект письма|переведи|перевести|отредактируй|объясни|формат письма|\b(?:draft|write|compose|template|translate|rewrite|format|explain)\b)/u.test(
      normalized,
    )
  )
    return false;
  const reads =
    /(?:какие|что нового|покажи|покажите|прочитай|прочитайте|найди|найдите|поищи|проверь|список|последние|непрочитанные|\b(?:what|show|list|read|find|search|check|latest|recent|unread|summari[sz]e)\b)/u.test(
      normalized,
    );
  if (!reads) return false;
  const mailbox =
    /(?:почт|письм|\b(?:e-?mails?|mailbox|inbox|outlook)\b)/u.test(normalized);
  const ownMailbox =
    /(?:^|[^\p{L}])(?:у меня|моя|моей|мои|моих|моим|моими|моего|моему|моем|мое|мою|my|our|inbox|outlook)(?=$|[^\p{L}])/u.test(
      normalized,
    );
  const cloudFiles =
    /(?:\b(?:onedrive|sharepoint)\b|корпоративн[а-я]*\s+(?:диск|файл)|(?:моем|моего|моему)\s+диск)/u.test(
      normalized,
    );
  return (mailbox && ownMailbox) || cloudFiles;
}

/** Only connection status is inspected. No Graph requests, source metadata or content. */
export async function getMicrosoft365ChatHandoff(
  messages: unknown[],
  userId: string,
  db: ReturnType<typeof createServerSupabase>,
): Promise<string | null> {
  const lastUser = [...messages]
    .reverse()
    .find(
      (message) =>
        message &&
        typeof message === "object" &&
        "role" in message &&
        message.role === "user",
    ) as { content?: unknown } | undefined;
  if (
    typeof lastUser?.content !== "string" ||
    !isMicrosoft365ReadRequest(lastUser.content)
  )
    return null;
  try {
    const status = await getMicrosoft365Status(userId, db);
    if (!status.available || status.connection?.status !== "connected")
      return null;
  } catch {
    // A disabled/unavailable integration must not break ordinary chat or expose DB errors.
    return null;
  }
  return /[а-я]/iu.test(lastUser.content)
    ? "Включите Microsoft 365 рядом с полем ввода и повторите вопрос — я смогу искать и читать вашу почту и корпоративные файлы прямо в этом чате. Сейчас доступ выключен."
    : "Enable Microsoft 365 beside the message composer and send your question again. I can then search and read your mailbox and corporate files in this conversation. Access is currently off.";
}
