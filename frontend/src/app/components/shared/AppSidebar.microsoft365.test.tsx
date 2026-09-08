import { cleanup, render, screen } from "@testing-library/react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

const state = vi.hoisted(() => ({ pathname: "/assistant/microsoft365", push: vi.fn(), setCurrentChatId: vi.fn() }));
vi.mock("next/navigation", () => ({ useRouter: () => ({ push: state.push }), usePathname: () => state.pathname }));
vi.mock("@/app/contexts/AuthContext", () => ({ useAuth: () => ({ user: { id: "sidebar-user", email: "test@example.com" }, signOut: vi.fn() }) }));
vi.mock("@/app/contexts/UserProfileContext", () => ({ useUserProfile: () => ({ profile: { displayName: "Test" } }) }));
vi.mock("@/app/contexts/ChatHistoryContext", () => ({ useChatHistoryContext: () => ({ chats: [], loadMoreChats: vi.fn(), setCurrentChatId: state.setCurrentChatId }) }));
vi.mock("@/app/lib/mikeApi", () => ({ listProjectSummaries: vi.fn().mockResolvedValue([]) }));
vi.mock("@/app/components/shared/SidebarChatItem", () => ({ SidebarChatItem: () => null }));
import { AppSidebar } from "./AppSidebar";

describe("Microsoft 365 navigation", () => {
    beforeEach(() => { vi.clearAllMocks(); state.pathname = "/assistant/microsoft365"; });
    afterEach(cleanup);

    it("keeps Microsoft 365 in the main assistant instead of a separate navigation item", () => {
        state.pathname = "/assistant";
        render(<AppSidebar isOpen onToggle={vi.fn()} />);
        expect(screen.queryByRole("button", { name: "Microsoft 365" })).not.toBeInTheDocument();
        expect(screen.getByRole("button", { name: "Assistant" })).toHaveAttribute("aria-current", "page");
    });
});
