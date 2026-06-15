export async function onRequestGet({ env }) {
  return Response.json({
    ok: true,
    aiAvailable: Boolean(env.ANTHROPIC_API_KEY),
    runtime: "serverless",
  });
}
