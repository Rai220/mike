import express, { type RequestHandler } from "express";
const parser = express.json({ limit: "64kb" });
/** Body-parser errors may carry the submitted corporate text; never pass them to generic logging. */
export const microsoft365JsonBody: RequestHandler = (req, res, next) => {
  res.set("Cache-Control", "no-store");
  res.set("Referrer-Policy", "no-referrer");
  parser(req, res, (error?: { type?: string }) => {
    if (!error) return next();
    const large = error.type === "entity.too.large";
    res
      .status(large ? 413 : 400)
      .json({
        code: "invalid_query",
        detail: large
          ? "The request is too large."
          : "The request is not valid JSON.",
      });
  });
};
