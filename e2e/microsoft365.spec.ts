/** Synthetic main-agent integration smoke. All APIs are mocked and external
 * requests blocked; no Microsoft account, database, or model keys are used. */
import { test, expect } from "@playwright/test";

test.use({ storageState: { cookies: [], origins: [] } });

test("Microsoft 365 access switches within the same main-agent conversation", async ({ page }, testInfo) => {
    const chatId = "synthetic-microsoft365";
    const chatPath = `/assistant/chat/${chatId}`;
    const expiresAt = "2099-09-14T12:00:00Z";
    const answer = "Your latest message is Synthetic project update.\nApproved budget: 120 units.\n[Untrusted link](https://source.example.invalid/document)\n![Untrusted image](https://source.example.invalid/tracker.png)";
    const followup = "The approved budget remains 120 units. No new Microsoft 365 sources were read.";
    let protectedChat = false;
    const requests: { use_microsoft365: boolean; chat_id: string; use_edgar: boolean }[] = [];
    const messages: { id: string; role: string; content: unknown; useMicrosoft365?: boolean; useEdgar?: boolean; model?: string; reasoning?: string }[] = [
        { id: "initial-user", role: "user", content: "Help me with this project." },
        { id: "initial-assistant", role: "assistant", content: [{ type: "content", text: "What would you like to review?" }] },
    ];
    const browserErrors: string[] = [];
    const externalRequests: string[] = [];
    const legacyRequests: string[] = [];
    const chat = () => ({ id: chatId, project_id: null, user_id: "synthetic-user", title: "Synthetic project", model: "claude-sonnet-4-6", reasoning_level: null, created_at: "2026-09-07T07:00:00Z", is_owner: true, access_role: "owner", microsoft365_protected: protectedChat, microsoft365_expires_at: protectedChat ? expiresAt : null });
    page.on("pageerror", error => browserErrors.push(error.message));
    await page.route("**/*", async route => {
        const request = route.request();
        const url = new URL(request.url());
        if (url.origin !== new URL(testInfo.project.use.baseURL as string).origin) {
            externalRequests.push(url.href);
            return route.abort();
        }
        if (!url.pathname.startsWith("/api/")) return route.continue();
        const path = url.pathname;
        let body: unknown = [];
        if (path.startsWith("/api/integrations/microsoft365/chats")) legacyRequests.push(path);
        if (path === "/api/auth/session") body = { user: { id: "synthetic-user", email: "test@example.invalid", pendingEmail: null, createdWithGoogle: false } };
        else if (path === "/api/user/profile") body = { displayName: "Synthetic tester", apiKeyStatus: { claude: true }, onboardingComplete: true, onboardingVersion: 1, lastSelectedChatModel: "claude-sonnet-4-6", darkMode: false, practiceAreas: [], creditsRemaining: 100, quickActionsVisible: false };
        else if (path === "/api/user/api-keys") body = { claude: true };
        else if (path === "/api/integrations/microsoft365") body = { available: true, connection: { id: "synthetic-connection", status: "connected" } };
        else if (path === `/api/chat/${chatId}`) body = { chat: chat(), messages, is_owner: true, access_role: "owner" };
        else if (path === "/api/chat" && request.method() === "GET") body = [chat()];
        else if (path === "/api/chat" && request.method() === "POST") {
            const input = request.postDataJSON();
            requests.push(input);
            protectedChat = true;
            const text = input.use_microsoft365 ? answer : followup;
            messages.push({ id: `user-${requests.length}`, role: "user", content: input.messages.at(-1).content, useMicrosoft365: input.use_microsoft365, useEdgar: input.use_edgar, model: input.model, reasoning: input.reasoning }, { id: `assistant-${requests.length}`, role: "assistant", content: [{ type: "content", text }] });
            const events = [
                { type: "microsoft365", protected: true, enabled: input.use_microsoft365, expiresAt, model: "claude-sonnet-4-6" },
                { type: "chat_id", chatId },
                { type: "content_delta", text },
                { type: "citations", citations: [] },
            ];
            return route.fulfill({ contentType: "text/event-stream", body: events.map(event => `data: ${JSON.stringify(event)}\n\n`).join("") + "data: [DONE]\n\n" });
        }
        await route.fulfill({ contentType: "application/json", body: JSON.stringify(body) });
    });

    await page.goto(chatPath);
    await expect(page.getByText("What would you like to review?", { exact: true })).toBeVisible();
    const toggle = page.getByRole("button", { name: "Use Microsoft 365 mail and files" });
    await expect(toggle).toHaveAttribute("aria-pressed", "false");
    await expect(page.getByRole("button", { name: "Use SEC EDGAR research" })).toBeVisible();
    await page.getByRole("button", { name: "Use SEC EDGAR research" }).click();
    const expectIndependentControls = async () => {
        for (const name of ["Add documents", "Open workflows", "Use SEC EDGAR research", "Choose model"]) {
            await expect(page.getByRole("button", { name, exact: true })).toBeVisible();
            await expect(page.getByRole("button", { name, exact: true })).toBeEnabled();
        }
        await expect(page.getByRole("button", { name: "Use SEC EDGAR research" })).toHaveAttribute("aria-pressed", "true");
    };
    await toggle.click();
    await expect(toggle).toHaveAttribute("aria-pressed", "true");
    await expectIndependentControls();
    await page.getByRole("combobox").fill("Какие письма у меня на почте?");
    await page.getByRole("button", { name: "Send message" }).click();
    await expect(page.getByText(answer, { exact: true })).toBeVisible();
    expect(requests[0]).toMatchObject({ chat_id: chatId, use_microsoft365: true, use_edgar: true });
    await expect(page).toHaveURL(new RegExp(`${chatPath}$`));
    await expect(page.getByRole("status").filter({ hasText: "Private Microsoft 365 conversation" })).toBeVisible();
    await page.getByRole("button", { name: "Chat actions", exact: true }).click();
    await expect(page.getByRole("menuitem", { name: "Share", exact: true })).toBeDisabled();
    await expect(page.getByRole("menuitem", { name: "Rename", exact: true })).toBeDisabled();
    await page.keyboard.press("Escape");
    await expect(page.getByRole("menuitem", { name: "Share", exact: true })).toHaveCount(0);
    await expect(page.locator('a[href*="source.example.invalid"], img[src*="source.example.invalid"]')).toHaveCount(0);
    await page.screenshot({ path: testInfo.outputPath("unified-microsoft365-desktop.png"), fullPage: true, animations: "disabled", style: "nextjs-portal { display: none !important; }" });

    await page.reload();
    await expect(page.getByText(answer, { exact: true })).toBeVisible();
    await expect(toggle).toHaveAttribute("aria-pressed", "true");
    await expectIndependentControls();
    await toggle.click();
    await expect(toggle).toHaveAttribute("aria-pressed", "false");
    await expectIndependentControls();
    await page.getByRole("combobox").fill("Explain that budget without reading new sources.");
    await page.getByRole("button", { name: "Send message" }).click();
    await expect(page.getByText(followup, { exact: true })).toBeVisible();
    expect(requests[1]).toMatchObject({ chat_id: chatId, use_microsoft365: false, use_edgar: true });
    await expect(page.getByText(answer, { exact: true })).toBeVisible();

    await page.reload();
    await expect(page).toHaveURL(new RegExp(`${chatPath}$`));
    await expect(page.getByText(answer, { exact: true })).toBeVisible();
    await expect(page.getByText(followup, { exact: true })).toBeVisible();
    await expect(toggle).toHaveAttribute("aria-pressed", "false");
    await expectIndependentControls();
    await page.setViewportSize({ width: 390, height: 844 });
    await expect(toggle).toBeVisible();
    await expect(page.getByText(followup, { exact: true })).toBeVisible();
    await expect(page.locator(".transition-opacity").filter({ hasText: followup })).toHaveCSS("opacity", "1");
    expect(await page.evaluate(() => document.documentElement.scrollWidth <= window.innerWidth)).toBe(true);
    await page.screenshot({ path: testInfo.outputPath("unified-microsoft365-mobile.png"), fullPage: true, animations: "disabled", style: "nextjs-portal { display: none !important; }" });
    expect(await page.evaluate(() => localStorage.getItem("microsoft365Enabled"))).toBeNull();
    expect(legacyRequests).toEqual([]);
    expect(externalRequests.filter(url => url.includes("source.example.invalid"))).toEqual([]);
    expect(browserErrors).toEqual([]);
});
