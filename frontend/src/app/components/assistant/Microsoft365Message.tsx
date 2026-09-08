import type { Message } from "../shared/types";

/** Corporate text stays inert: sources and model output cannot create requests or links. */
export function Microsoft365Message({ message, isLoading }: { message: Message; isLoading: boolean }) {
    const text = message.events?.filter((event) => event.type === "content").map((event) => event.text).join("") || message.content;
    return <div className="whitespace-pre-wrap break-words py-3 text-sm" aria-live="polite">
        {message.error || text || (isLoading ? "Lispenard is working…" : "")}
    </div>;
}
