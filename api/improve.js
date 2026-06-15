import { onRequest } from "../functions/api/improve.js";
import { sendFetchResponse, toFetchRequest } from "./_bridge.js";

export const config = {
  maxDuration: 60,
};

export default async function handler(req, res) {
  const response = await onRequest({
    request: toFetchRequest(req),
    env: process.env,
    ctx: {},
  });
  await sendFetchResponse(res, response);
}

