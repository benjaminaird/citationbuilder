const DEFAULT_MODEL = "claude-sonnet-4-6";
const DEFAULT_FALLBACK_MODEL = "claude-haiku-4-5-20251001";

function json(body, status = 200) {
  return Response.json(body, { status });
}

function normalizeSpaces(text) {
  if (typeof text !== "string") return "";
  return text
    .replace(/[ \t]+/g, " ")
    .replace(/ *\n */g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function buildPrompt(body) {
  const {
    mode = "all",
    award = "",
    soa = "",
    citation = "",
    opening = "",
    closing = "",
    charLimit = 0,
    targetLow = 0,
    primaryBillet = "",
    additionalBillets = "",
    achievements = "",
    metricsToPreserve = [],
    validationFindings = [],
    awardJustificationFindings = [],
    realityFindings = [],
  } = body || {};

  const reviewContext = [
    Array.isArray(validationFindings) && validationFindings.length
      ? ["VALIDATION FINDINGS TO FIX:", ...validationFindings.map((x) => `- ${String(x)}`)].join("\n")
      : "",
    Array.isArray(awardJustificationFindings) && awardJustificationFindings.length
      ? ["AWARD JUSTIFICATION FINDINGS:", ...awardJustificationFindings.map((x) => `- ${String(x)}`)].join("\n")
      : "",
    Array.isArray(realityFindings) && realityFindings.length
      ? ["REALITY CHECK FINDINGS (do not invent facts; preserve questionable claims but make wording professional):", ...realityFindings.map((x) => `- ${String(x)}`)].join("\n")
      : "",
    primaryBillet || additionalBillets
      ? [
          "DUTY CONTEXT:",
          primaryBillet ? `- Primary billet: ${String(primaryBillet)}` : "",
          additionalBillets ? `- Additional billets/collateral duties: ${String(additionalBillets)}` : "",
          "- Do not force additional billets into the citation opening sentence. Use them in the body only when relevant to the accomplishments.",
        ].filter(Boolean).join("\n")
      : "",
    Array.isArray(metricsToPreserve) && metricsToPreserve.length
      ? ["QUANTITATIVE IMPACT TO PRESERVE WHENEVER POSSIBLE:", ...metricsToPreserve.map((x) => `- ${String(x)}`)].join("\n")
      : "",
    achievements ? ["ORIGINAL ACCOMPLISHMENTS FOR FACT CHECKING:", String(achievements)].join("\n") : "",
  ].filter(Boolean).join("\n\n");

  const metricRule = [
    "- Preserve quantitative impact whenever possible, including personnel counts, dollars, percentages, events, ceremonies, duration, volunteer hours, beneficiaries, Defense Travel System authorizations and vouchers, travel claims, inspections, and training events.",
    "- If the citation must be shortened, remove generic adjectives, filler phrases, and the weakest accomplishments before removing numbers or measurable impact.",
    "- Group related accomplishments by duty or topic so the citation and Summary of Action do not bounce between unrelated subjects.",
    "- Summary of Action text must remain paragraph-format normal prose, not ALL CAPS. Citation casing rules apply only to the citation.",
    "- Prioritize major leadership, quantified impact, duration, personnel, money or resources, command advisory, readiness, ceremonial or operational or community impact, and place routine professional development last.",
    "- When facts support it, retain public visibility and scope indicators such as national, international, diplomatic, presidential, state ceremony, senior military or civilian leader, battalion-wide, command-wide, and installation-wide impact.",
    "- Keep the award-specific citation frame: opening sentence first, accomplishment body second, standard credit/traditions closing last.",
    "- For Meritorious Mast and Certificate of Commendation, treat the output as a citation/certificate only; do not create or require a Summary of Action.",
    "- Meritorious Mast must remain ALL CAPS, use the mandatory 'DURING THE PERIOD OF ... THROUGH ...' opening, use an initiative/perseverance/total dedication closing, and fit the portrait certificate style of about 14 lines.",
    "- Certificate of Commendation must remain ALL CAPS, use the mandatory 'EXCEPTIONAL PERFORMANCE OF ... DUTIES' opening, use a 'reflected great credit' closing, and fit the landscape certificate style of about 9 lines.",
  ].join("\n");

  const isUpper = ["MMAST", "NAM", "NMC", "CERTCOM"].includes(award);
  const caseRule = isUpper
    ? "The citation must remain in ALL CAPITAL LETTERS."
    : "The citation must remain in sentence case (mixed case, like normal prose).";

  if (mode === "soa") {
    if (!soa) return { error: "Provide a Summary of Action to improve.", status: 400 };
    return {
      mode,
      isUpper,
      system: [
        "You are an expert United States Marine Corps awards writer and editor.",
        "You refine the WORDING of a Summary of Action to be crisp, active-voice, narrative, and competition-ready.",
        "",
        "HARD CONSTRAINTS - you must obey every one:",
        "- You may ONLY improve the wording, transitions, impact statements, and military writing style of the SOA body.",
        "- Keep the Summary of Action in paragraph format with normal sentence case. Do not add headings or ALL CAPS sections.",
        "- Do NOT invent or add facts, numbers, awards, events, outcomes, units, or personnel names.",
        metricRule,
        "- Do NOT alter dates, rank, EDIPI, names, billet, or unit names.",
        "- Do NOT introduce ANY abbreviations. Spell everything out.",
        '- The only permitted abbreviation is "Washington, D.C." Always write the unit exactly as "Marine Barracks, Washington, D.C.,".',
        "- Uphold professional Marine Corps prose - strong action verbs, specific impacts, results-focused.",
        "",
        "Respond with ONLY a JSON object, no markdown, of the shape:",
        '{ "soa": string, "notes": string[] }',
        'Put short reviewer notes (what you improved / cautions) in "notes".',
      ].join("\n"),
      user: [
        `Award type: ${award}`,
        reviewContext ? `\n${reviewContext}` : "",
        "",
        "CURRENT SUMMARY OF ACTION (refine wording, transitions, and impact - preserve the same information):",
        soa || "(none)",
      ].join("\n"),
    };
  }

  if (mode === "expand") {
    if (!citation) return { error: "Provide a citation to expand.", status: 400 };
    const targetMsg = targetLow
      ? `- The expanded citation, including spaces, should approach ${targetLow} characters but MUST NOT exceed ${charLimit || targetLow + 200} characters.`
      : charLimit
        ? `- The expanded citation, including spaces, MUST NOT exceed ${charLimit} characters.`
        : "";

    return {
      mode,
      isUpper,
      system: [
        "You are an expert United States Marine Corps awards writer and editor.",
        "You EXPAND the body of a citation to increase its length and impact WITHOUT inventing new facts.",
        "",
        "HARD CONSTRAINTS - you must obey every one:",
        "- You may ONLY expand the body wording of the citation.",
        "- You MUST NOT change the opening sentence. It is provided and is fixed.",
        "- You MUST NOT change the closing sentence. It is provided and is fixed.",
        "- " + caseRule,
        "- The Summary of Action must remain paragraph-format normal prose, not ALL CAPS. Citation casing rules apply only to the citation.",
        "- Do NOT invent or add facts, numbers, awards, events, outcomes, units, or personnel names.",
        metricRule,
        "- Do NOT alter dates, rank, EDIPI, names, billet, or unit names.",
        "- Do NOT introduce ANY abbreviations. Spell everything out.",
        '- The only permitted abbreviation is "Washington, D.C." Always write the unit exactly as "Marine Barracks, Washington, D.C.,".',
        targetMsg,
        "",
        "EXPANSION TECHNIQUES (you may use any combination):",
        "- Elaborate on the impact and significance of each accomplishment.",
        "- Improve transitions between accomplishments for better flow.",
        "- Strengthen military writing style with more vivid active-voice verbs.",
        "- Expand on how the Marine's actions benefited the unit, mission, or Marine Corps.",
        "- Add appropriate connective language between related achievements.",
        "",
        "Respond with ONLY a JSON object, no markdown, of the shape:",
        '{ "citation": string, "notes": string[] }',
        'Put short reviewer notes (what you expanded / cautions) in "notes".',
      ].filter(Boolean).join("\n"),
      user: [
        `Award type: ${award}`,
        reviewContext ? `\n${reviewContext}` : "",
        "",
        "FIXED OPENING (do not change, reproduce exactly at the start of the citation):",
        opening || "(none provided)",
        "",
        "FIXED CLOSING (do not change, reproduce exactly at the end of the citation):",
        closing || "(none provided)",
        "",
        "CURRENT CITATION (expand the body between opening and closing only):",
        citation || "(none)",
      ].join("\n"),
    };
  }

  if (mode === "loa") {
    if (!soa) return { error: "Provide a Letter of Authorization to improve.", status: 400 };
    return {
      mode,
      isUpper,
      system: [
        "You are an expert United States Marine Corps awards writer and editor.",
        "You refine the WORDING of a Letter of Authorization for the Outstanding Volunteer Service Medal to be crisp, active-voice, and competition-ready.",
        "",
        "HARD CONSTRAINTS - you must obey every one:",
        "- You may ONLY improve the wording, transitions, impact statements, and military writing style of the LOA body.",
        "- You MUST NOT change the structure: keep the header, identification, volunteer narrative, impact, and recommendation sections.",
        "- Do NOT invent or add facts, numbers, awards, events, outcomes, units, hours volunteered, or organization names.",
        metricRule,
        "- Do NOT alter dates, rank, EDIPI, names, billet, or unit names.",
        "- Do NOT introduce ANY abbreviations. Spell everything out.",
        '- The only permitted abbreviation is "Washington, D.C." Always write the unit exactly as "Marine Barracks, Washington, D.C.,".',
        "- Uphold professional Marine Corps prose - strong action verbs, specific impacts, results-focused.",
        "",
        "Respond with ONLY a JSON object, no markdown, of the shape:",
        '{ "loa": string, "notes": string[] }',
        'Put short reviewer notes (what you improved / cautions) in "notes".',
      ].join("\n"),
      user: [
        `Award type: ${award}`,
        reviewContext ? `\n${reviewContext}` : "",
        "",
        "CURRENT LETTER OF AUTHORIZATION (refine wording, transitions, and impact - preserve the same information):",
        soa || "(none)",
      ].join("\n"),
    };
  }

  if (!soa && !citation) {
    return { error: "Provide at least a Summary of Action or a citation to refine.", status: 400 };
  }

  return {
    mode,
    isUpper,
    system: [
      "You are an expert United States Marine Corps awards writer and editor.",
      "You refine the WORDING of award documents to be crisp, active-voice, and competition-ready.",
      "",
      "HARD CONSTRAINTS - you must obey every one:",
      "- You may ONLY improve the body wording of the Summary of Action and the citation.",
      "- You MUST NOT change the opening sentence. It is provided and is fixed.",
      "- You MUST NOT change the closing sentence. It is provided and is fixed.",
      "- " + caseRule,
      "- The Summary of Action must remain paragraph-format normal prose, not ALL CAPS. Citation casing rules apply only to the citation.",
      "- Do NOT invent or add facts, numbers, awards, events, or outcomes.",
      metricRule,
      "- Do NOT alter dates, rank, EDIPI, names, billet, or unit names.",
      "- Do NOT introduce ANY abbreviations. Spell everything out.",
      '- The only permitted abbreviation is "Washington, D.C." Always write the unit exactly as "Marine Barracks, Washington, D.C.,".',
      charLimit ? `- The citation, including spaces, MUST NOT exceed ${charLimit} characters.` : "",
      "",
      "Respond with ONLY a JSON object, no markdown, of the shape:",
      '{ "soa": string, "citation": string, "notes": string[] }',
      'Put short reviewer notes (what you improved / cautions) in "notes".',
    ].filter(Boolean).join("\n"),
    user: [
      `Award type: ${award}`,
      reviewContext ? `\n${reviewContext}` : "",
      "",
      "FIXED OPENING (do not change, reproduce exactly at the start of the citation):",
      opening || "(none provided)",
      "",
      "FIXED CLOSING (do not change, reproduce exactly at the end of the citation):",
      closing || "(none provided)",
      "",
      "CURRENT SUMMARY OF ACTION (refine wording only):",
      soa || "(none)",
      "",
      "CURRENT CITATION (refine the body between opening and closing only):",
      citation || "(none)",
    ].join("\n"),
  };
}

async function requestAnthropic({ apiKey, model, fallbackModel, system, user }) {
  const makeRequest = (chosenModel) => fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: {
      "content-type": "application/json",
      "x-api-key": apiKey,
      "anthropic-version": "2023-06-01",
    },
    body: JSON.stringify({
      model: chosenModel,
      max_tokens: 2000,
      system,
      messages: [{ role: "user", content: user }],
    }),
  });

  let usedModel = model;
  let response = await makeRequest(usedModel);
  let detail = "";

  if (!response.ok) {
    detail = await response.text().catch(() => "");
    const modelUnavailable = response.status === 404 && /model/i.test(detail) && fallbackModel !== usedModel;
    if (modelUnavailable) {
      usedModel = fallbackModel;
      response = await makeRequest(usedModel);
      detail = "";
    }
  }

  return { response, detail, usedModel };
}

function providerError(status, detail, usedModel) {
  let providerMessage = "";
  try {
    const parsedDetail = JSON.parse(detail);
    providerMessage = parsedDetail?.error?.message || parsedDetail?.message || "";
  } catch {
    providerMessage = detail;
  }

  const cleanProviderMessage = String(providerMessage).replace(/\s+/g, " ").trim();
  const modelHint = status === 404 || /model/i.test(cleanProviderMessage)
    ? ` Check ANTHROPIC_MODEL; this deploy tried ${usedModel}.`
    : "";
  const authHint = status === 401
    ? " Check that ANTHROPIC_API_KEY is set correctly in Cloudflare Pages."
    : "";
  const quotaHint = status === 402 || status === 429
    ? " Check Anthropic billing, credits, or rate limits."
    : "";

  return `Anthropic returned ${status}.${authHint}${quotaHint}${modelHint}${
    cleanProviderMessage ? ` ${cleanProviderMessage}` : " Please try again."
  }`;
}

function parseAIJson(raw) {
  const match = raw.match(/\{[\s\S]*\}/);
  return JSON.parse(match ? match[0] : raw);
}

function repinCitation({ citation, opening, closing }) {
  let outCitation = normalizeSpaces(citation);

  if (opening && !outCitation.toLowerCase().startsWith(opening.slice(0, 24).toLowerCase())) {
    const cleaned = outCitation.replace(/^.*?\.\s*/, "");
    outCitation = `${opening} ${cleaned}`.trim();
  }

  if (closing) {
    const closeKey = closing.slice(-24).toLowerCase();
    if (!outCitation.toLowerCase().endsWith(closeKey)) {
      const cleaned = outCitation.replace(/\s*[^.]*\.\s*$/, "");
      outCitation = `${cleaned} ${closing}`.trim();
    }
  }

  return outCitation;
}

export async function onRequest({ request, env }) {
  if (request.method === "OPTIONS") return new Response(null, { status: 204 });
  if (request.method !== "POST") return json({ error: "Method not allowed." }, 405);

  const apiKey = env.ANTHROPIC_API_KEY || "";
  if (!apiKey) {
    return json({
      error: "AI refinement is unavailable: ANTHROPIC_API_KEY is not configured for this deployment. All formatting, validation, and drafting features still work without it.",
    }, 503);
  }

  let body;
  try {
    body = await request.json();
  } catch {
    return json({ error: "Invalid JSON request body." }, 400);
  }

  const prompt = buildPrompt(body);
  if (prompt.error) return json({ error: prompt.error }, prompt.status || 400);

  try {
    const { response, detail, usedModel } = await requestAnthropic({
      apiKey,
      model: env.ANTHROPIC_MODEL || DEFAULT_MODEL,
      fallbackModel: env.ANTHROPIC_FALLBACK_MODEL || DEFAULT_FALLBACK_MODEL,
      system: prompt.system,
      user: prompt.user,
    });

    if (!response.ok) {
      const fullDetail = detail || await response.text().catch(() => "");
      return json({ error: providerError(response.status, fullDetail, usedModel) }, 502);
    }

    const data = await response.json();
    const raw = Array.isArray(data.content)
      ? data.content.map((block) => block.text || "").join("")
      : "";
    let parsed;
    try {
      parsed = parseAIJson(raw);
    } catch {
      return json({ error: "The AI response could not be parsed. Please try again." }, 502);
    }

    const notes = Array.isArray(parsed.notes)
      ? parsed.notes.map((note) => String(note)).slice(0, 8)
      : [];

    if (prompt.mode === "soa") return json({ soa: normalizeSpaces(parsed.soa || body.soa), notes });
    if (prompt.mode === "loa") return json({ loa: normalizeSpaces(parsed.loa || body.soa), notes });

    const outSoa = prompt.mode === "expand" ? body.soa : normalizeSpaces(parsed.soa || body.soa);
    let outCitation = repinCitation({
      citation: parsed.citation || body.citation,
      opening: body.opening || "",
      closing: body.closing || "",
    });
    if (prompt.isUpper) outCitation = outCitation.toUpperCase();

    return json({ soa: outSoa, citation: outCitation, notes });
  } catch (err) {
    return json({
      error: err?.name === "AbortError"
        ? "The AI request timed out. Please try again."
        : "An unexpected error occurred. Please try again.",
    }, err?.name === "AbortError" ? 504 : 500);
  }
}
