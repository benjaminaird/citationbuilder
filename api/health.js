import { onRequestGet } from "../functions/api/health.js";
import { sendFetchResponse, toFetchRequest } from "./_bridge.js";

export default async function handler(req, res) {
  if (req.method !== "GET") {
    res.status(405).json({ error: "Method not allowed." });
    return;
  }

  const response = await onRequestGet({
    request: toFetchRequest(req),
    env: process.env,
    ctx: {},
  });
  await sendFetchResponse(res, response);
}

