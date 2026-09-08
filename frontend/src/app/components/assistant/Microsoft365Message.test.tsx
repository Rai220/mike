import { render, screen } from "@testing-library/react";
import { describe, expect, it } from "vitest";
import { Microsoft365Message } from "./Microsoft365Message";

describe("Microsoft365Message", () => {
    it("keeps corporate markdown inert without remote requests or generated links", () => {
        const text = "![secret](https://outside.example/image)\n[link](https://outside.example)";
        const { container } = render(<Microsoft365Message message={{ role: "assistant", content: text }} isLoading={false} />);
        expect(container).toHaveTextContent(text.replace("\n", " "));
        expect(container.querySelector("a,img,iframe")).toBeNull();
    });
    it("shows a waiting state inside the ordinary transcript", () => {
        render(<Microsoft365Message message={{ role: "assistant", content: "" }} isLoading />);
        expect(screen.getByText("Lispenard is working…")).toBeInTheDocument();
    });
});
