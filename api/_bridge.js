function headerObject(headers) {
  const out = {};
  headers.forEach((value, key) => {
    out[key] = value;
  });
  return out;
}

function requestUrl(req) {
  const proto = req.headers["x-forwarded-proto"] || "https";
  const host = req.headers.host || "citationbuilder.vercel.app";
  return `${proto}://${host}${req.url || "/"}`;
}

function requestBody(req) {
  if (req.body === undefined || req.body === null) return undefined;
  return typeof req.body === "string" ? req.body : JSON.stringify(req.body);
}

export function toFetchRequest(req) {
  return new Request(requestUrl(req), {
    method: req.method,
    headers: req.headers,
    body: ["GET", "HEAD"].includes(req.method) ? undefined : requestBody(req),
  });
}

export async function sendFetchResponse(res, response) {
  res.status(response.status);
  const headers = headerObject(response.headers);
  for (const [key, value] of Object.entries(headers)) {
    res.setHeader(key, value);
  }
  res.send(await response.text());
}

