/**
 * CitationBuilder v2 — Marine Corps award writing aid.
 *
 * Deployment model (preserved from v1):
 *   - Node/Express server
 *   - Static frontend served from dist/
 *   - Render-compatible (npm start -> node server.js)
 *   - Anthropic API key stays server-side as ANTHROPIC_API_KEY
 *
 * Philosophy:
 *   This is a formatting, validation, and drafting engine first. AI is optional.
 *   Every core feature works with the AI offline; only AI refinement endpoints
 *   need the key. If the key is missing endpoints return clean 503s and the
 *   frontend keeps working.
 */

import path from "path";
import { fileURLToPath } from "url";
import express from "express";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const PORT = process.env.PORT || 3000;
const ANTHROPIC_API_KEY = process.env.ANTHROPIC_API_KEY || "";
const ANTHROPIC_MODEL = process.env.ANTHROPIC_MODEL || "claude-sonnet-4-6";
const ANTHROPIC_FALLBACK_MODEL =
  process.env.ANTHROPIC_FALLBACK_MODEL || "claude-haiku-4-5-20251001";

app.use(express.json({ limit: "256kb" }));
app.use(express.static(path.join(__dirname, "dist")));

/** Collapse runaway whitespace without touching newline structure. */
function normalizeSpaces(text) {
  if (typeof text !== "string") return "";
  return text
    .replace(/[ \t]+/g, " ")
    .replace(/ *\n */g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

/** Health check (Render & uptime monitors). */
app.get("/api/health", (req, res) => {
  res.json({ ok: true, aiAvailable: Boolean(ANTHROPIC_API_KEY) });
});

/**
 * POST /api/improve
 * Body: { mode, award, soa, citation, opening, closing, charLimit, targetLow }
 *
 * Modes:
 *   "all"       — refine both SOA and citation body (default, legacy)
 *   "soa"       — refine SOA only; citation is ignored
 *   "expand"    — expand citation body toward target length, no new facts
 *   "loa"       — refine Letter of Authorization (OVSM), no citation changes
 *
 * AI may ONLY refine/expand the body wording. Opening and closing sentences
 * are pinned: the model is instructed not to touch them, and the server
 * re-pins them after the response regardless.
 */
app.post("/api/improve", async (req, res) => {
  if (!ANTHROPIC_API_KEY) {
    return res.status(503).json({
      error:
        "AI refinement is unavailable: ANTHROPIC_API_KEY is not configured on the server. All formatting, validation, and drafting features still work without it.",
    });
  }

  const {
    mode = "all",
    award = "",
    soa = "",
    citation = "",
    opening = "",
    closing = "",
    charLimit = 0,
    targetLow = 0,
    validationFindings = [],
    awardJustificationFindings = [],
    realityFindings = [],
  } = req.body || {};

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
  ].filter(Boolean).join("\n\n");

  const isUpper = ["NAM", "NMC", "CERTCOM"].includes(award);
  const caseRule = isUpper
    ? "The citation must remain in ALL CAPITAL LETTERS."
    : "The citation must remain in sentence case (mixed case, like normal prose).";

  let system = "";
  let user = "";

  if (mode === "soa") {
    // ---- SOA-only improvement ----
    if (!soa) {
      return res.status(400).json({ error: "Provide a Summary of Action to improve." });
    }

    system = [
      "You are an expert United States Marine Corps awards writer and editor.",
      "You refine the WORDING of a Summary of Action to be crisp, active-voice, narrative, and competition-ready.",
      "",
      "HARD CONSTRAINTS — you must obey every one:",
      "- You may ONLY improve the wording, transitions, impact statements, and military writing style of the SOA body.",
      "- You MUST NOT change the structure: keep Background, Accomplishments, and Recommendation sections.",
      "- Do NOT invent or add facts, numbers, awards, events, outcomes, units, or personnel names.",
      "- Do NOT alter dates, rank, EDIPI, names, billet, or unit names.",
      "- Do NOT introduce ANY abbreviations. Spell everything out.",
      '- The only permitted abbreviation is "Washington, D.C." Always write the unit exactly as "Marine Barracks, Washington, D.C.,".',
      "- Uphold professional Marine Corps prose — strong action verbs, specific impacts, results-focused.",
      "",
      "Respond with ONLY a JSON object, no markdown, of the shape:",
      '{ "soa": string, "notes": string[] }',
      'Put short reviewer notes (what you improved / cautions) in "notes".',
    ].join("\n");

    user = [
      `Award type: ${award}`,
      "",
      "CURRENT SUMMARY OF ACTION (refine wording, transitions, and impact — preserve the same information):",
      soa || "(none)",
    ].join("\n");

  } else if (mode === "expand") {
    // ---- Citation expansion ----
    if (!citation) {
      return res.status(400).json({ error: "Provide a citation to expand." });
    }

    const targetMsg = targetLow
      ? `- The expanded citation, including spaces, should approach ${targetLow} characters but MUST NOT exceed ${charLimit || targetLow + 200} characters.`
      : charLimit
        ? `- The expanded citation, including spaces, MUST NOT exceed ${charLimit} characters.`
        : "";

    system = [
      "You are an expert United States Marine Corps awards writer and editor.",
      "You EXPAND the body of a citation to increase its length and impact WITHOUT inventing new facts.",
      "",
      "HARD CONSTRAINTS — you must obey every one:",
      "- You may ONLY expand the body wording of the citation.",
      "- You MUST NOT change the opening sentence. It is provided and is fixed.",
      "- You MUST NOT change the closing sentence. It is provided and is fixed.",
      "- " + caseRule,
      "- Do NOT invent or add facts, numbers, awards, events, outcomes, units, or personnel names.",
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
    ]
      .filter(Boolean)
      .join("\n");

    user = [
      `Award type: ${award}`,
      "",
      "FIXED OPENING (do not change, reproduce exactly at the start of the citation):",
      opening || "(none provided)",
      "",
      "FIXED CLOSING (do not change, reproduce exactly at the end of the citation):",
      closing || "(none provided)",
      "",
      "CURRENT CITATION (expand the body between opening and closing only):",
      citation || "(none)",
    ].join("\n");

  } else if (mode === "loa") {
    // ---- LOA improvement (OVSM) ----
    if (!soa) {
      return res.status(400).json({ error: "Provide a Letter of Authorization to improve." });
    }

    system = [
      "You are an expert United States Marine Corps awards writer and editor.",
      "You refine the WORDING of a Letter of Authorization for the Outstanding Volunteer Service Medal to be crisp, active-voice, and competition-ready.",
      "",
      "HARD CONSTRAINTS — you must obey every one:",
      "- You may ONLY improve the wording, transitions, impact statements, and military writing style of the LOA body.",
      "- You MUST NOT change the structure: keep the header, identification, volunteer narrative, impact, and recommendation sections.",
      "- Do NOT invent or add facts, numbers, awards, events, outcomes, units, hours volunteered, or organization names.",
      "- Do NOT alter dates, rank, EDIPI, names, billet, or unit names.",
      "- Do NOT introduce ANY abbreviations. Spell everything out.",
      '- The only permitted abbreviation is "Washington, D.C." Always write the unit exactly as "Marine Barracks, Washington, D.C.,".',
      "- Uphold professional Marine Corps prose — strong action verbs, specific impacts, results-focused.",
      "",
      "Respond with ONLY a JSON object, no markdown, of the shape:",
      '{ "loa": string, "notes": string[] }',
      'Put short reviewer notes (what you improved / cautions) in "notes".',
    ].join("\n");

    user = [
      `Award type: ${award}`,
      reviewContext ? `\n${reviewContext}` : "",
      "",
      "CURRENT LETTER OF AUTHORIZATION (refine wording, transitions, and impact — preserve the same information):",
      soa || "(none)",
    ].join("\n");

  } else {
    // ---- Default: refine both ----
    if (!soa && !citation) {
      return res
        .status(400)
        .json({ error: "Provide at least a Summary of Action or a citation to refine." });
    }

    system = [
      "You are an expert United States Marine Corps awards writer and editor.",
      "You refine the WORDING of award documents to be crisp, active-voice, and competition-ready.",
      "",
      "HARD CONSTRAINTS — you must obey every one:",
      "- You may ONLY improve the body wording of the Summary of Action and the citation.",
      "- You MUST NOT change the opening sentence. It is provided and is fixed.",
      "- You MUST NOT change the closing sentence. It is provided and is fixed.",
      "- " + caseRule,
      "- Do NOT invent or add facts, numbers, awards, events, or outcomes.",
      "- Do NOT alter dates, rank, EDIPI, names, billet, or unit names.",
      "- Do NOT introduce ANY abbreviations. Spell everything out.",
      '- The only permitted abbreviation is "Washington, D.C." Always write the unit exactly as "Marine Barracks, Washington, D.C.,".',
      charLimit
        ? `- The citation, including spaces, MUST NOT exceed ${charLimit} characters.`
        : "",
      "",
      "Respond with ONLY a JSON object, no markdown, of the shape:",
      '{ "soa": string, "citation": string, "notes": string[] }',
      'Put short reviewer notes (what you improved / cautions) in "notes".',
    ]
      .filter(Boolean)
      .join("\n");

    user = [
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
    ].join("\n");
  }

  try {
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 45000);

    const makeAnthropicRequest = (model) => fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "content-type": "application/json",
        "x-api-key": ANTHROPIC_API_KEY,
        "anthropic-version": "2023-06-01",
      },
      body: JSON.stringify({
        model,
        max_tokens: 2000,
        system,
        messages: [{ role: "user", content: user }],
      }),
      signal: controller.signal,
    });

    let usedModel = ANTHROPIC_MODEL;
    let response = await makeAnthropicRequest(usedModel);
    let detail = "";

    if (!response.ok) {
      detail = await response.text().catch(() => "");
      const modelUnavailable =
        response.status === 404 && /model/i.test(detail) && ANTHROPIC_FALLBACK_MODEL !== usedModel;

      if (modelUnavailable) {
        console.warn(
          `Anthropic model ${usedModel} unavailable; retrying with ${ANTHROPIC_FALLBACK_MODEL}`,
        );
        usedModel = ANTHROPIC_FALLBACK_MODEL;
        response = await makeAnthropicRequest(usedModel);
        detail = "";
      }
    }

    clearTimeout(timeout);

    if (!response.ok) {
      if (!detail) detail = await response.text().catch(() => "");
      console.error("Anthropic API error:", response.status, detail.slice(0, 500));

      let providerMessage = "";
      try {
        const parsedDetail = JSON.parse(detail);
        providerMessage = parsedDetail?.error?.message || parsedDetail?.message || "";
      } catch {
        providerMessage = detail;
      }

      const cleanProviderMessage = String(providerMessage).replace(/\s+/g, " ").trim();
      const modelHint = response.status === 404 || /model/i.test(cleanProviderMessage)
        ? ` Check ANTHROPIC_MODEL; this deploy tried ${usedModel}.`
        : "";
      const authHint = response.status === 401
        ? " Check that ANTHROPIC_API_KEY is set correctly in Render."
        : "";
      const quotaHint = response.status === 402 || response.status === 429
        ? " Check Anthropic billing, credits, or rate limits."
        : "";

      return res.status(502).json({
        error: `Anthropic returned ${response.status}.${authHint}${quotaHint}${modelHint}${
          cleanProviderMessage ? ` ${cleanProviderMessage}` : " Please try again."
        }`,
      });
    }

    const data = await response.json();
    const raw = Array.isArray(data.content)
      ? data.content.map((block) => block.text || "").join("")
      : "";

    let parsed = null;
    try {
      const match = raw.match(/\{[\s\S]*\}/);
      parsed = JSON.parse(match ? match[0] : raw);
    } catch (err) {
      console.error("Failed to parse AI JSON:", err.message);
      return res
        .status(502)
        .json({ error: "The AI response could not be parsed. Please try again." });
    }

    if (mode === "soa") {
      // SOA-only: return improved SOA
      const outSoa = normalizeSpaces(parsed.soa || soa);
      const notes = Array.isArray(parsed.notes)
        ? parsed.notes.map((n) => String(n)).slice(0, 8)
        : [];
      return res.json({ soa: outSoa, notes });
    }

    if (mode === "loa") {
      // LOA-only: return improved LOA
      const outLOA = normalizeSpaces(parsed.loa || soa);
      const notes = Array.isArray(parsed.notes)
        ? parsed.notes.map((n) => String(n)).slice(0, 8)
        : [];
      return res.json({ loa: outLOA, notes });
    }

    // For "all" and "expand" modes, handle citation + opening/closing pinning
    let outSoa = mode === "expand" ? soa : normalizeSpaces(parsed.soa || soa);
    let outCitation = normalizeSpaces(parsed.citation || citation);

    // Re-pin the fixed opening/closing — never trust the model to leave them alone.
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

    if (isUpper) outCitation = outCitation.toUpperCase();

    const notes = Array.isArray(parsed.notes)
      ? parsed.notes.map((n) => String(n)).slice(0, 8)
      : [];

    res.json({ soa: outSoa, citation: outCitation, notes });
  } catch (err) {
    if (err.name === "AbortError") {
      return res.status(504).json({ error: "The AI request timed out. Please try again." });
    }
    console.error("Unexpected /api/improve error:", err.message);
    res.status(500).json({ error: "An unexpected error occurred. Please try again." });
  }
});

// Fallback to the single page for any unmatched GET route.
app.use((req, res) => {
  res.sendFile(path.join(__dirname, "dist", "index.html"));
});

app.listen(PORT, () => {
  console.log(`CitationBuilder v2 running on port ${PORT}`);
  console.log(`AI refinement: ${ANTHROPIC_API_KEY ? "enabled" : "disabled (no ANTHROPIC_API_KEY)"}`);
});
