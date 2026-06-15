import { onRequestGet as health } from "../functions/api/health.js";
import { onRequest as improve } from "../functions/api/improve.js";

export default {
  async fetch(request, env, ctx) {
    const url = new URL(request.url);

    if (url.pathname === "/api/health") {
      return health({ request, env, ctx });
    }

    if (url.pathname === "/api/improve") {
      return improve({ request, env, ctx });
    }

    return env.ASSETS.fetch(request);
  },
};

