import { useState, useEffect, useCallback, useRef } from "react";
import { toast } from "sonner";
import { Document, Packer, Paragraph, TextRun, HeadingLevel } from "docx";
import { saveAs } from "file-saver";

/* ================================================================
   CitationBuilder v1 — Marine Corps Award Drafting Engine
   Formatting, validation & drafting first. AI is optional.
   ================================================================ */

// ---- Types ----
type AwardKey = "MMAST" | "CERTCOM" | "OVSM" | "NAM" | "NMC" | "MSM" | "LOM";
type PronounKey = "m" | "f";

interface AwardConfig {
  label: string;
  casing: "upper" | "sentence";
  maxChars: number;
  target: [number, number] | null;
  closing: "lesser" | "great" | "loa";
  greatCredit: boolean;
  isLOA: boolean;
  citationOnly?: boolean;
}

interface FormState {
  award: AwardKey;
  rank: string;
  lastName: string;
  firstName: string;
  edipi: string;
  pronoun: PronounKey;
  billet: string;
  additionalBillets: string;
  unit: string;
  dateFrom: string;
  dateTo: string;
  attr1: string;
  attr2: string;
  adj: string;
  achievements: string;
}

interface CheckItem {
  status: "ok" | "warn" | "err";
  title: string;
  detail: string;
  fixId?: string;
}

type AchievementCategory = "Leadership" | "Operations" | "Training" | "Administration" | "Innovation" | "Community Relations" | "Uncategorized";

interface ClassifiedAchievement {
  text: string;
  category: AchievementCategory;
}

interface QualityScores {
  quantifiableImpact: number;
  strongVerbs: number;
  leadershipLanguage: number;
  resultOriented: number;
  awardLevelMatch: number;
  overall: number;
}

interface AwardMatchResult {
  score: number;
  severity: "none" | "possible" | "severe";
  title: string;
  detail: string;
  recommendedAward: AwardKey;
  recommendations: string[];
}

interface WeakInputIssue {
  title: string;
  detail: string;
}

interface SavedDraft {
  id: string;
  name: string;
  form: FormState;
  soa: string;
  citation: string;
  aiNotes: string[];
  updatedAt: string;
  createdAt: string;
}

// ---- Constants ----
const RANKS = [
  "Private", "Private First Class", "Lance Corporal", "Corporal", "Sergeant",
  "Staff Sergeant", "Gunnery Sergeant", "Master Sergeant", "First Sergeant",
  "Master Gunnery Sergeant", "Sergeant Major",
  "Warrant Officer", "Chief Warrant Officer 2", "Chief Warrant Officer 3", "Chief Warrant Officer 4", "Chief Warrant Officer 5",
  "Second Lieutenant", "First Lieutenant", "Captain", "Major", "Lieutenant Colonel", "Colonel",
] as const;

const AWARDS: Record<AwardKey, AwardConfig> = {
  MMAST:   { label: "Meritorious Mast",                       casing: "upper",    maxChars: 0,    target: null,          closing: "lesser", greatCredit: false, isLOA: false, citationOnly: true },
  CERTCOM: { label: "Certificate of Commendation",            casing: "upper",    maxChars: 1250, target: [1200, 1245], closing: "lesser", greatCredit: true,  isLOA: false, citationOnly: true },
  OVSM:    { label: "Outstanding Volunteer Service Medal",     casing: "sentence", maxChars: 0,    target: null,          closing: "loa",    greatCredit: true,  isLOA: true },
  NAM:     { label: "Navy & Marine Corps Achievement Medal",  casing: "upper",    maxChars: 1250, target: [1200, 1245], closing: "lesser", greatCredit: false, isLOA: false },
  NMC:     { label: "Navy & Marine Corps Commendation Medal", casing: "upper",    maxChars: 1250, target: [1200, 1245], closing: "lesser", greatCredit: false, isLOA: false },
  MSM:     { label: "Meritorious Service Medal",              casing: "sentence", maxChars: 1800, target: [1700, 1790], closing: "great",  greatCredit: true,  isLOA: false },
  LOM:     { label: "Legion of Merit",                        casing: "sentence", maxChars: 1650, target: [1550, 1640], closing: "great",  greatCredit: true,  isLOA: false },
};

const PRONOUNS: Record<PronounKey, { subj: string; obj: string; poss: string; refl: string }> = {
  m: { subj: "he", obj: "him", poss: "his", refl: "himself" },
  f: { subj: "she", obj: "her", poss: "her", refl: "herself" },
};

const EXPANSIONS: Record<string, string> = {
  "SSgt": "Staff Sergeant", "GySgt": "Gunnery Sergeant", "MSgt": "Master Sergeant",
  "1stLt": "First Lieutenant", "Capt": "Captain", "Maj": "Major", "LtCol": "Lieutenant Colonel",
  "NCOIC": "Noncommissioned Officer in Charge", "SNCOIC": "Staff Noncommissioned Officer in Charge",
  "CO": "Commanding Officer", "XO": "Executive Officer",
};

// ---- Achievement Classification Keywords ----
const CATEGORY_KEYWORDS: Record<Exclude<AchievementCategory, "Uncategorized">, RegExp[]> = {
  Leadership: [
    /\b(led|lead|directed|supervised|managed|mentored|guided|oversaw|commanded|headed|chaired|spearheaded|galvanized|inspired|motivated)\b/i,
    /\b(leadership|supervision|direction|guidance|mentorship|stewardship)\b/i,
    /\b(NCO|NCOIC|SNCOIC|officer in charge|section head|team lead)\b/i,
  ],
  Operations: [
    /\b(operations?|mission|deployment|exercise|drill|readiness|contingency|mobilization)\b/i,
    /\b(executed|coordinated|orchestrated|conducted|spearheaded)\s+(operations?|missions?|exercises?|events?)/i,
    /\b(operational|tactical|logistics|supply chain|maintenance)\b/i,
  ],
  Training: [
    /\b(trained|instructed|taught|certified|qualified|coached|school|course|curriculum)\b/i,
    /\b(training|instruction|education|MOS|PME|professional military)\b/i,
    /\b(range|qualification|marksmanship|PFT|CFT|MCMAP)\b/i,
  ],
  Administration: [
    /\b(processed|administered|tracked|documented|records?|reports?|rosters?|database)\b/i,
    /\b(organized|streamlined|automated|overhauled|standardized)\s+(process|system|workflow|procedure)\b/i,
    /\b(administrative|clerical|personnel|manpower|fiscal|budget|pay|audit)\b/i,
  ],
  Innovation: [
    /\b(created|designed|developed|innovated|modernized|improved|redesigned|built|launched|established|pioneered)\b/i,
    /\b(innovation|initiative|program|system|solution|tool|application|platform)\b/i,
    /\b(from scratch|ground up|new (program|system|process|initiative))\b/i,
  ],
  "Community Relations": [
    /\b(community|volunteer|outreach|ceremony|ceremonial|event|public|media|relations?|partnership|engagement)\b/i,
    /\b(color guard|funeral detail|parade|memorial|toys for tots|blood drive)\b/i,
  ],
};

const CATEGORY_COLORS: Record<AchievementCategory, string> = {
  Leadership: "#1a4d8f",
  Operations: "#7c3a12",
  Training: "#2f7d44",
  Administration: "#6b4e8a",
  Innovation: "#b5751a",
  "Community Relations": "#0d6b6b",
  Uncategorized: "#6b6f76",
};

const UNIT_CANON = "Marine Barracks, Washington, D.C.,";
const UNIT_PRESETS = [
  "Marine Barracks, Washington, D.C.",
  "A Company",
  "B Company",
  "Headquarters and Service Company",
  "Guard Company",
  '"The Commandant\'s Own", U.S. Marine Drum & Bugle Corps',
  '"The President\'s Own", U.S. Marine Band',
] as const;
const STORAGE_KEY = "citationbuilder.form";
const SAVED_DRAFTS_KEY = "citationbuilder.savedDrafts";
const RELEASE_NOTICE_KEY = "citationbuilder.v1ReleaseNoticeAccepted";
const APP_VERSION = "v1.0";
const SUPPORT_EMAIL = "mailto:benjaminaird@yahoo.com?subject=CitationBuilder%20V1%20Issue";

// ---- Helpers ----
function expandAbbr(text: string): string {
  let out = text;
  for (const [abbr, full] of Object.entries(EXPANSIONS)) {
    const re = new RegExp("\\b" + abbr.replace(/[.*+?^${}()|[\]\\]/g, "\\$&") + "\\b", "g");
    out = out.replace(re, full);
  }
  return out;
}

// ---- Classification ----
function classifyAchievement(line: string): AchievementCategory {
  const clean = line.replace(/^[-•*\d.)\s]+/, "").trim();
  if (!clean) return "Uncategorized";
  let best: AchievementCategory = "Uncategorized";
  let bestScore = 0;
  for (const [cat, patterns] of Object.entries(CATEGORY_KEYWORDS)) {
    let score = 0;
    for (const re of patterns) {
      if (re.test(clean)) score++;
    }
    if (score > bestScore) { bestScore = score; best = cat as AchievementCategory; }
  }
  return best;
}

function classifyAll(lines: string[]): ClassifiedAchievement[] {
  return lines.map((l) => ({ text: l.replace(/^[-•*\d.)\s]+/, "").trim(), category: classifyAchievement(l) }));
}

// ---- Quality & Award Fit Scoring ----
const STRONG_VERBS = /\b(led|spearheaded|orchestrated|directed|championed|galvanized|transformed|revitalized|modernized|overhauled|pioneered|executed|coordinated|mentored|cultivated|engineered|fortified|streamlined|optimized|accelerated|catapulted)\b/i;
const LEADERSHIP_TERMS = /\b(leadership|initiative|judgment|responsibility|accountability|stewardship|guidance|mentorship|example|standard)\b/i;
const RESULT_TERMS = /\b(resulted in|yielded|achieved|attained|produced|generated|delivered|improved|increased|reduced|eliminated|exceeded|surpassed|enhanced)\b/i;
const QUANTIFIABLE = /\b(\d+\s*(percent|%|Marines?|Sailors?|personnel|service members?|members|events?|ceremonies?|inspections?|trainings?|hours?|days?|weeks?|months?|years?|beneficiaries|profiles?|authorizations?|vouchers?|claims?|dollars?|\$))\b|\$\s?\d+|\b(zero|no)\s+(discrepanc|error|failure|incident|loss)/i;
const VISIBILITY_TERMS = /\b(national|international|diplomatic|state ceremony|presidential|senior military|senior civilian|distinguished visitors?|public visibility|global prestige|strategic|battalion-wide|command-wide|installation-wide)\b/i;

const SCOPE_TERMS = {
  individual: /\b(individual|personally|single|shop|desk|task)\b/i,
  section: /\b(section|team|squad|platoon|detail|watch|shift|cell)\b/i,
  unit: /\b(unit|company|battery|detachment|battalion|squadron|command)\b/i,
  command: /\b(command-wide|battalion-wide|installation-wide|regiment|group|wing|base|installation|headquarters|enterprise|national|international|diplomatic|presidential)\b/i,
  service: /\b(service-wide|marine corps|navy|department|institutional|enterprise-wide|force-wide)\b/i,
};

function achievementLines(form: FormState): string[] {
  return form.achievements.split("\n").map((l) => l.trim()).filter(Boolean);
}

function metricsToPreserve(form: FormState): string[] {
  const metricContext = /\b\d+\s*(?:percent|%|marines?|sailors?|personnel|service members?|members|events?|ceremonies?|inspections?|trainings?|hours?|days?|weeks?|months?|years?|beneficiaries|profiles?|authorizations?|vouchers?|claims?)\b|\$\s?\d[\d,]*(?:\.\d+)?|\b\d[\d,]*\s*(?:dollars?|volunteer hours?)\b|\b(?:defense travel system|dts|travel claims?|authorizations?|vouchers?|profiles?|commander advisory|advised the commander)\b/i;
  return achievementLines(form)
    .filter((line) => metricContext.test(line))
    .slice(0, 12);
}

function countMatches(text: string, re: RegExp): number {
  const flags = re.flags.includes("g") ? re.flags : re.flags + "g";
  return (text.match(new RegExp(re.source, flags)) || []).length;
}

function rankSeniority(rank: string): number {
  const r = rank.toLowerCase();
  if (/private|lance corporal/.test(r)) return 1;
  if (/corporal|sergeant$/.test(r)) return 2;
  if (/staff sergeant|gunnery sergeant/.test(r)) return 3;
  if (/master sergeant|first sergeant|master gunnery sergeant|sergeant major/.test(r)) return 4;
  if (/second lieutenant|first lieutenant|captain|warrant officer|chief warrant officer 2/.test(r)) return 3;
  if (/major|lieutenant colonel|colonel|chief warrant officer 3|chief warrant officer 4|chief warrant officer 5/.test(r)) return 4;
  return 2;
}

function awardLevel(award: AwardKey): number {
  if (award === "MMAST" || award === "CERTCOM") return 0;
  if (award === "NAM" || award === "OVSM") return 1;
  if (award === "NMC") return 2;
  if (award === "MSM") return 3;
  if (award === "LOM") return 4;
  return 1;
}

function billetScope(billet: string): number {
  const b = billet.toLowerCase();
  if (/commander|commanding officer|executive officer|sergeant major|director|chief of staff|department head/.test(b)) return 4;
  if (/officer in charge|oic|sncoic|ncoic|first sergeant|operations chief|company|battalion|section head/.test(b)) return 3;
  if (/platoon|section|team|squad|supervisor|leader|chief/.test(b)) return 2;
  if (b.trim()) return 1;
  return 0;
}

function scopeLevel(text: string): number {
  if (SCOPE_TERMS.service.test(text)) return 4;
  if (SCOPE_TERMS.command.test(text)) return 4;
  if (SCOPE_TERMS.unit.test(text)) return 3;
  if (SCOPE_TERMS.section.test(text)) return 2;
  if (SCOPE_TERMS.individual.test(text)) return 1;
  return 1;
}

function recommendedAwardFromSupport(support: number): AwardKey {
  if (support <= 0) return "MMAST";
  if (support <= 1) return "NAM";
  if (support === 2) return "NMC";
  if (support === 3) return "MSM";
  return "LOM";
}

function awardShortLabel(award: AwardKey): string {
  if (award === "MMAST") return "Meritorious Mast";
  if (award === "NMC") return "Navy and Marine Corps Commendation Medal";
  return AWARDS[award].label;
}

function analyzeWeakInput(form: FormState): WeakInputIssue[] {
  const text = form.achievements;
  if (!text.trim()) {
    return [{
      title: "Accomplishments missing",
      detail: "Add actions, measurable impact, scope of responsibility, results, and operational significance before submission.",
    }];
  }

  const issues: WeakInputIssue[] = [];
  if (!QUANTIFIABLE.test(text)) {
    issues.push({
      title: "Missing measurable impact",
      detail: "Add numbers, personnel affected, events supported, readiness impact, time saved, money saved, or command-level effect.",
    });
  }
  if (!/(responsib|supervis|managed|led|oversaw|section|platoon|team|unit|command|personnel|marines|sailors)/i.test(text + " " + form.billet)) {
    issues.push({
      title: "Missing scope of responsibility",
      detail: "State who or what the Marine was responsible for: personnel, programs, events, equipment, mission areas, or command functions.",
    });
  }
  if (!RESULT_TERMS.test(text)) {
    issues.push({
      title: "Missing result",
      detail: "Show what changed because of the action: readiness improved, errors reduced, events completed, timelines met, or mission risk lowered.",
    });
  }
  if (!LEADERSHIP_TERMS.test(text) && !/led|supervised|mentored|trained|guided|directed/i.test(text)) {
    issues.push({
      title: "Missing leadership effect",
      detail: "Describe how the Marine influenced others, raised standards, trained personnel, improved performance, or enabled the chain of command.",
    });
  }
  if (!/(mission|readiness|operational|command|inspection|ceremonial|deployment|exercise|training|support)/i.test(text)) {
    issues.push({
      title: "Missing operational significance",
      detail: "Tie the accomplishment to mission execution, command priorities, readiness, ceremonial support, inspections, or unit effectiveness.",
    });
  }
  return issues;
}

function analyzeRealityIssues(form: FormState): CheckItem[] {
  const text = form.achievements;
  if (!text.trim()) return [];
  const checks: CheckItem[] = [];
  const personalRoutine = /\b(fed (my )?cat|cleaned (my )?litter box|walked (my )?dog|washed (my )?car|made (my )?bed|did (my )?laundry|read one book|read a book)\b/i;
  const trivial = /\b(showed up|was on time|did my job|completed daily tasks|answered emails|attended formation)\b/i;
  const hugeClaim = /\b(saved|led|trained|managed|impacted|supported)\s+(?:over\s+|more than\s+)?(?:10,000|10000|[5-9]\d{3,})\s+(marines|sailors|personnel|people|families)\b/i;
  const absurd = /\b(defeated sharks?|became king of france|ended all wars|personally ended all wars|single[- ]handedly saved the world|invented freedom|meme|skibidi|yeet|sigma|rizz)\b/i;
  const noContext = hugeClaim.test(text) && !/(command-wide|service-wide|installation|multi-year|enterprise|across|throughout|program|initiative)/i.test(text);

  if (personalRoutine.test(text) || trivial.test(text) || absurd.test(text)) {
    checks.push({
      status: "warn",
      title: "Reality Check Triggered",
      detail: "One or more accomplishments may not be realistic and should be verified.",
    });
  }
  if (noContext) {
    checks.push({
      status: "warn",
      title: "Reality Check Triggered",
      detail: "Impact claims appear unusually large and may require supporting context.",
    });
  }
  return checks;
}

function analyzeOVSMIssues(form: FormState): CheckItem[] {
  if (form.award !== "OVSM") return [];
  const text = `${form.achievements} ${form.dateFrom} ${form.dateTo}`;
  const checks: CheckItem[] = [];
  const hasHours = /\b\d+\s*(hours?|hrs?)\b/i.test(text);
  const hasDuration = /\b(months?|years?|weekly|monthly|sustained|from\b.+\bto\b)\b/i.test(text) || Boolean(form.dateFrom && form.dateTo);
  const hasCommunity = /\b(community|families|youth|veterans|students|residents|beneficiaries|civilians|charity|nonprofit|organization)\b/i.test(text);
  const hasBeneficiaries = /\b\d+\s*(families|youth|students|veterans|residents|people|beneficiaries|children|members)\b/i.test(text);
  const hasVolunteerLeadership = /\b(led|organized|coordinated|supervised|managed|trained|mentored|chaired)\b/i.test(text);
  const hasSustainability = /\b(sustained|recurring|weekly|monthly|program|established|continued|ongoing|long-term)\b/i.test(text);

  checks.push(hasHours
    ? { status: "ok", title: "OVSM volunteer hours", detail: "Volunteer hours are documented." }
    : { status: "warn", title: "OVSM volunteer hours", detail: "Add total volunteer hours or a clear estimate." });
  checks.push(hasDuration
    ? { status: "ok", title: "OVSM duration", detail: "Duration of service is documented." }
    : { status: "warn", title: "OVSM duration", detail: "Describe sustained service over time, not a one-time event." });
  checks.push(hasCommunity
    ? { status: "ok", title: "OVSM community impact", detail: "Community impact is described." }
    : { status: "warn", title: "OVSM community impact", detail: "Volunteer effort is documented but community impact is not clearly described." });
  checks.push(hasBeneficiaries
    ? { status: "ok", title: "OVSM beneficiaries", detail: "Beneficiaries served are quantified." }
    : { status: "warn", title: "OVSM beneficiaries", detail: "Add who benefited and how many people or organizations were served." });
  checks.push(hasVolunteerLeadership
    ? { status: "ok", title: "OVSM leadership", detail: "Volunteer leadership is shown." }
    : { status: "warn", title: "OVSM leadership", detail: "If applicable, describe leadership within the volunteer organization or event." });
  checks.push(hasSustainability
    ? { status: "ok", title: "OVSM sustainability", detail: "Sustained or recurring effort is shown." }
    : { status: "warn", title: "OVSM sustainability", detail: "Explain whether the effort was recurring, sustained, or created lasting benefit." });
  return checks;
}

type AchievementTopic =
  | "leadership"
  | "police"
  | "administration"
  | "operations"
  | "training"
  | "ceremonial"
  | "community"
  | "professionalDevelopment"
  | "other";

function additionalBilletList(form: FormState): string[] {
  return form.additionalBillets
    .split(/[;,/]+|\band\b/i)
    .map((item) => item.trim())
    .filter(Boolean);
}

function hasDutyWord(line: string, duty: string): boolean {
  const words = duty.toLowerCase().match(/[a-z0-9]+/g) || [];
  if (!words.length) return false;
  const text = line.toLowerCase();
  return words.some((word) => word.length > 3 && new RegExp(`\\b${word.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}\\b`, "i").test(text));
}

function achievementTopic(line: string, form: FormState): AchievementTopic {
  const text = line.toLowerCase();
  if (/\b(?:defense travel system|dts|travel|voucher|authorization|profile|claim|administrative|admin|orders|roster|report|records?|pay|personnel action)\b/i.test(text)) {
    return "administration";
  }
  if (/\b(?:police|sergeant|patrol|security|law enforcement|call|response|watch|post|area supervisor)\b/i.test(text)) {
    return "police";
  }
  if (/\b(?:funeral|ceremon|bugler|color guard|parade|honors|memorial)\b/i.test(text)) {
    return "ceremonial";
  }
  if (/\b(?:train|training|instruct|course|class|qualification|professional development)\b/i.test(text)) {
    return /professional development|read one book|completed course/i.test(text) ? "professionalDevelopment" : "training";
  }
  if (/\b(?:volunteer|community|families|youth|veterans|charity|nonprofit|beneficiaries)\b/i.test(text)) {
    return "community";
  }
  if (/\b(?:led|supervised|managed|mentored|advised|briefed|commander|commanding officer|marines?|sailors?|personnel)\b/i.test(text)) {
    return "leadership";
  }
  if (/\b(?:mission|readiness|operation|inspection|deployment|maintenance|logistics|resources?)\b/i.test(text)) {
    return "operations";
  }
  const dutyHints = [form.billet, ...additionalBilletList(form)].filter(Boolean);
  if (dutyHints.some((duty) => hasDutyWord(line, duty))) return "operations";
  return "other";
}

function groupedAchievementLines(form: FormState, lines: string[], forCitation: boolean): string[] {
  const groups = new Map<AchievementTopic, { line: string; index: number; score: number }[]>();
  lines.forEach((line, index) => {
    const topic = achievementTopic(line, form);
    if (!groups.has(topic)) groups.set(topic, []);
    groups.get(topic)!.push({ line, index, score: accomplishmentPriorityScore(line, form) });
  });

  const topicBase: Record<AchievementTopic, number> = {
    leadership: 95,
    police: 88,
    administration: 84,
    operations: 80,
    training: 64,
    ceremonial: 60,
    community: 56,
    professionalDevelopment: 25,
    other: 35,
  };

  const orderedTopics = Array.from(groups.keys()).sort((a, b) => {
    if (forCitation) {
      const aScore = Math.max(...groups.get(a)!.map((item) => item.score)) + topicBase[a];
      const bScore = Math.max(...groups.get(b)!.map((item) => item.score)) + topicBase[b];
      return bScore - aScore;
    }
    const aFirst = Math.min(...groups.get(a)!.map((item) => item.index));
    const bFirst = Math.min(...groups.get(b)!.map((item) => item.index));
    return aFirst - bFirst;
  });

  return orderedTopics.flatMap((topic) => groups.get(topic)!.sort((a, b) => {
    if (forCitation) return b.score - a.score || a.index - b.index;
    return a.index - b.index;
  }).map((item) => item.line));
}

function accomplishmentPriorityScore(line: string, form: FormState): number {
  const text = `${line} ${form.billet} ${form.additionalBillets}`;
  let score = 0;
  score += countMatches(text, QUANTIFIABLE) * 18;
  score += countMatches(text, STRONG_VERBS) * 12;
  score += countMatches(text, LEADERSHIP_TERMS) * 12;
  score += countMatches(text, RESULT_TERMS) * 14;
  if (/\b\d+\s*(marines?|sailors?|personnel|service members?|members|civilians|beneficiaries|profiles?|authorizations?|vouchers?|claims?|events?|ceremonies?|inspections?|trainings?|hours?|months?|years?)\b/i.test(text)) score += 22;
  if (/\$\s?\d|(?:dollars?|funds?|budget|resources?|claims?)/i.test(text)) score += 18;
  if (/\b(?:defense travel system|dts|travel|voucher|authorization|profile|claim)\b/i.test(text)) score += 18;
  if (/\b(?:advised|briefed|counseled|recommended|informed)\b.*\b(?:commander|commanding officer|senior enlisted|leadership)\b/i.test(text)) score += 18;
  if (VISIBILITY_TERMS.test(text)) score += 16;
  if (/\b(?:funerals?|parades?|ceremonies|ceremonial|diplomatic honors?|state ceremonies|presidential|evening parade|sunset parade|battle color)\b/i.test(text)) score += 12;
  score += scopeLevel(text) * 10;
  if (/(mission|readiness|operational|ceremonial|inspection|deployment|training|command)/i.test(text)) score += 12;
  if (/(read one book|professional development|completed course|fed my cat|litter box|showed up|did my job)/i.test(text)) score -= 28;
  if (form.award === "LOM" || form.award === "MSM") {
    if (/(command-wide|institutional|sustained|program|enterprise|organization|led|supervised|managed)/i.test(text)) score += 15;
  }
  if (form.award === "NAM" || form.award === "NMC") {
    if (/(section|team|detail|event|readiness|zero discrepancies|personnel|marines)/i.test(text)) score += 10;
  }
  return score;
}

function prioritizedAchievementLines(form: FormState, forCitation: boolean): string[] {
  const lines = achievementLines(form);
  if (!forCitation) return groupedAchievementLines(form, lines, false);
  const ranked = lines
    .map((line, index) => ({ line, index, score: accomplishmentPriorityScore(line, form) }))
    .sort((a, b) => b.score - a.score || a.index - b.index);
  const cfg = AWARDS[form.award];
  const limit = cfg.maxChars ? (cfg.maxChars <= 1250 ? 6 : 8) : 8;
  return groupedAchievementLines(
    form,
    ranked.slice(0, Math.max(1, Math.min(limit, ranked.length))).map((item) => item.line),
    true,
  );
}

function awardMatchScore(form: FormState, citation = ""): AwardMatchResult {
  const text = `${form.achievements} ${citation}`.trim();
  const lines = achievementLines(form);
  if (!form.achievements.trim() && !citation.trim()) {
    return {
      score: 70,
      severity: "none",
      title: "Award level not yet assessed",
      detail: "Add accomplishments to assess award level.",
      recommendedAward: form.award,
      recommendations: [],
    };
  }
  const rank = rankSeniority(form.rank);
  const dutyText = `${form.billet} ${form.additionalBillets}`;
  const billet = billetScope(dutyText);
  const scope = scopeLevel(`${dutyText} ${text}`);
  const quant = Math.min(4, countMatches(text, QUANTIFIABLE));
  const leadership = Math.min(4, countMatches(text, STRONG_VERBS) + countMatches(text, LEADERSHIP_TERMS));
  const results = Math.min(4, countMatches(text, RESULT_TERMS));
  const months = serviceMonths(form);
  let language = /(transformational|institutional|service-wide|command-wide|enduring|exceptional|monumental|strategic|enterprise)/i.test(text) ? 2 : 0;
  if (VISIBILITY_TERMS.test(text)) language += 1;
  if (months !== null && months >= 24) language += 1;
  const support = Math.min(4, Math.round((rank + billet + scope + Math.min(4, lines.length) + quant + leadership + results + language) / 7));
  const selected = awardLevel(form.award);
  const diff = selected - support;
  const recommendedAward = recommendedAwardFromSupport(support);

  let severity: AwardMatchResult["severity"] = "none";
  let title = "Award level appears supportable";
  let detail = "Rank, billet, scope, and accomplishment detail generally match the selected award level.";
  const recommendations: string[] = [];

  if (form.award === "LOM" && rank <= 2 && support <= 2) {
    severity = "severe";
    title = "Award Mismatch Warning";
    detail = "The selected award may not be appropriate for the rank, billet, and scope entered. Review award level before submission.";
    recommendations.push("This does not currently support an LOM.");
    recommendations.push(`Recommended award level: ${awardShortLabel(recommendedAward)}.`);
  } else if (diff >= 2) {
    severity = "severe";
    title = "Award Mismatch Warning";
    detail = "The selected award may not be appropriate for the rank, billet, and scope entered. Review award level before submission.";
    recommendations.push(`This package does not currently support ${AWARDS[form.award].label}.`);
    recommendations.push(`Recommended award level: ${awardShortLabel(recommendedAward)}.`);
    recommendations.push("Add sustained leadership, measurable command-level impact, and broader organizational results.");
  } else if (diff === 1) {
    severity = "possible";
    title = "Possible award mismatch";
    detail = "The selected award may be high for the current rank, billet, and accomplishment scope.";
    recommendations.push(form.award === "NMC"
      ? "This may support a Navy Comm if stronger command-level impact is added."
      : `This package currently supports ${awardShortLabel(recommendedAward)}.`);
  }

  if (selected <= 1 && rank >= 4 && (scope >= 3 || leadership >= 2 || quant >= 2)) {
    severity = severity === "severe" ? "severe" : "possible";
    title = "Possible award mismatch";
    detail = "The selected award may be low for the rank, billet, and apparent scope entered.";
    recommendations.push("This may be under-awarded; consider Navy Comm/MSM.");
  }

  if (!recommendations.length) {
    if (selected === 1) recommendations.push("This package currently supports a NAM when impact remains local and individual.");
    if (selected === 2) recommendations.push("This package may support a Navy and Marine Corps Commendation Medal when command-level impact is clear.");
    if (selected >= 3) recommendations.push("Ensure the write-up shows sustained leadership, organizational impact, and measurable results.");
  }

  const score = Math.max(15, Math.min(100, 100 - Math.abs(diff) * 25 - (severity === "severe" ? 20 : severity === "possible" ? 8 : 0)));
  return { score, severity, title, detail, recommendedAward, recommendations: Array.from(new Set(recommendations)).slice(0, 3) };
}

function scoreQuality(citation: string, form: FormState): QualityScores {
  const text = citation || "";
  const qi = (text.match(QUANTIFIABLE) || []).length;
  const sv = (text.match(STRONG_VERBS) || []).length;
  const ll = (text.match(LEADERSHIP_TERMS) || []).length;
  const ro = (text.match(RESULT_TERMS) || []).length;

  // Normalize to 0-100 based on expected counts for a strong award
  const qiScore = Math.min(100, Math.round((qi / 4) * 100));
  const svScore = Math.min(100, Math.round((sv / 6) * 100));
  const llScore = Math.min(100, Math.round((ll / 4) * 100));
  const roScore = Math.min(100, Math.round((ro / 4) * 100));

  const awardMatch = awardMatchScore(form, citation).score;

  const overall = Math.round((qiScore * 0.25) + (svScore * 0.25) + (llScore * 0.20) + (roScore * 0.20) + (awardMatch * 0.10));

  return { quantifiableImpact: qiScore, strongVerbs: svScore, leadershipLanguage: llScore, resultOriented: roScore, awardLevelMatch: awardMatch, overall };
}

// ---- Citation expansion (client-side) ----
function expandCitationLocally(citation: string, targetLow: number, maxChars: number, form: FormState): string {
  if (!citation) return citation;
  if (citation.length >= targetLow) return citation;

  const p = PRONOUNS[form.pronoun];
  const rl = rankLast(form);
  const closingText = buildClosing(form);
  const heShe = p.subj.charAt(0).toUpperCase() + p.subj.slice(1);

  // Strip closing temporarily so we can insert text before it
  let body = citation;
  const closingIdx = body.lastIndexOf(closingText);
  if (closingIdx > -1) {
    body = body.slice(0, closingIdx).trimEnd();
  } else {
    // Try to find closing by common patterns
    const closeMatch = body.match(/\s+(By .+?traditions of the Marine Corps and the United States Naval Service\.)\s*$/i);
    if (closeMatch) {
      body = body.slice(0, body.lastIndexOf(closeMatch[1])).trimEnd();
    }
  }

  // Build expansion phrases from available facts (no invented facts)
  const elaborations: string[] = [];

  // Impact elaboration
  elaborations.push(
    `${rl}'s exceptional leadership and technical expertise directly contributed to the unit's mission success and operational readiness.`,
    `${heShe} consistently set the standard for excellence, inspiring those around ${p.obj} and elevating the performance of the entire section.`,
    `The lasting impact of ${p.poss} efforts was felt across the command, setting a benchmark for professionalism and dedication.`,
    `${heShe} demonstrated sound judgment, initiative, and an unwavering commitment to the highest standards of the Marine Corps.`,
  );

  let expanded = body;
  const remaining = maxChars - expanded.length - closingText.length - 2; // -2 for space + safety

  // Add elaborations one at a time until we approach the target
  for (const phrase of elaborations) {
    if (expanded.length + phrase.length + 1 <= maxChars - closingText.length - 1) {
      expanded = expanded.trimEnd() + " " + phrase;
      if (expanded.length + closingText.length + 1 >= targetLow) break;
    }
  }

  // If still too short, add transitional connectors within the body
  if (expanded.length + closingText.length + 1 < targetLow) {
    expanded = expanded
      .replace(/\. (?=[A-Z])/g, ". Moreover, ")
      .replace(/\. Moreover, \. Moreover, /g, ". "); // prevent double connectors
  }

  // Reattach closing
  expanded = expanded.trimEnd() + " " + closingText;

  // Final trim to max
  if (expanded.length > maxChars) {
    // Try to cut at the last sentence boundary before maxChars
    const cutPoint = expanded.lastIndexOf(". ", maxChars - closingText.length - 1);
    if (cutPoint > 0) {
      expanded = expanded.slice(0, cutPoint + 1) + " " + closingText;
    }
  }

  return expanded;
}

function normalizeWashingtonDC(text: string): string {
  let out = text;
  out = out.replace(/Washington\s*,?\s*D\s*\.?\s*C\.?/gi, "Washington, D.C.");
  out = out.replace(/Washington,D\.C\./gi, "Washington, D.C.");
  out = out.replace(/Marine Barracks,?\s+Washington,\s*D\.C\./gi, UNIT_CANON);
  out = out.replace(/D\.C\.,\s*,+/gi, "D.C.,");
  out = out.replace(/D\.C\.,\s*\./gi, "D.C.");
  out = out.replace(/D\.C\.\s*,\s*(?=(from|while|for|during|and|in|with)\b)/gi, "D.C., ");
  out = out.replace(/D\.C\.,\s*$/gi, "D.C.");
  out = out.replace(/,\s*,/g, ",");
  out = out.replace(/\s+([,.])/g, "$1");
  return out;
}

function enforceWashington(text: string): string {
  return normalizeWashingtonDC(text);
}

function cleanup(text: string): string {
  return normalizeWashingtonDC(text)
    .replace(/[ \t]+/g, " ")
    .replace(/ *\n */g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function redactSensitiveForAI(value: unknown): unknown {
  if (typeof value === "string") {
    return value
      .replace(/\b\d{10}\b/g, "[REDACTED EDIPI]")
      .replace(/\b\d{3}-\d{2}-\d{4}\b/g, "[REDACTED SSN]")
      .replace(/\bEDIPI\s*[:#]?\s*\d+\b/gi, "EDIPI: [REDACTED]");
  }
  if (Array.isArray(value)) return value.map((item) => redactSensitiveForAI(item));
  if (value && typeof value === "object") {
    return Object.fromEntries(
      Object.entries(value as Record<string, unknown>)
        .filter(([key]) => !/edipi|ssn|social/i.test(key))
        .map(([key, item]) => [key, redactSensitiveForAI(item)]),
    );
  }
  return value;
}

function citationSentences(text: string): string[] {
  return cleanup(text)
    .split(/(?<=[.!?])\s+/)
    .map((s) => s.trim())
    .filter(Boolean);
}

function estimatedCitationLines(text: string, charsPerLine: number): number {
  const cleaned = cleanup(text);
  if (!cleaned) return 0;
  return cleaned
    .split(/\n+/)
    .map((line) => Math.max(1, Math.ceil(line.trim().length / charsPerLine)))
    .reduce((sum, lines) => sum + lines, 0);
}

function enforceCitationLimit(text: string, form: FormState): string {
  const cfg = AWARDS[form.award];
  if (!cfg.maxChars) return applyCase(cleanup(text), cfg.casing);

  const limit = cfg.maxChars;
  let out = applyCase(cleanup(text), cfg.casing);
  if (out.length <= limit) return out;

  const opening = applyCase(cleanup(buildOpening(form)), cfg.casing);
  const closing = applyCase(cleanup(buildClosing(form)), cfg.casing);
  const lower = out.toLowerCase();
  const openingLower = opening.toLowerCase();
  const closingLower = closing.toLowerCase();
  let body = out;

  if (lower.startsWith(openingLower)) body = body.slice(opening.length).trim();
  if (body.toLowerCase().endsWith(closingLower)) body = body.slice(0, -closing.length).trim();

  const ranked = citationSentences(body)
    .filter((s) => s && s.toLowerCase() !== openingLower && s.toLowerCase() !== closingLower)
    .map((sentence, index) => ({ sentence, index, score: accomplishmentPriorityScore(sentence, form) }))
    .sort((a, b) => b.score - a.score || a.index - b.index);

  const kept: string[] = [];
  for (const item of ranked) {
    const candidate = cleanup([opening, ...kept, item.sentence, closing].join(" "));
    if (candidate.length <= limit) kept.push(item.sentence);
  }

  out = cleanup([opening, ...kept, closing].join(" "));
  if (out.length <= limit) return applyCase(out, cfg.casing);

  const available = Math.max(0, limit - opening.length - closing.length - 2);
  let compressedBody = kept.join(" ");
  if (compressedBody.length > available) {
    compressedBody = compressedBody.slice(0, Math.max(0, available)).replace(/\s+\S*$/, "").trim();
    if (compressedBody && !/[.!?]$/.test(compressedBody)) compressedBody += ".";
  }
  out = cleanup([opening, compressedBody, closing].filter(Boolean).join(" "));

  if (out.length > limit) {
    const reserved = ` ${closing}`;
    const openingAndBodyLimit = Math.max(0, limit - reserved.length);
    out = cleanup(out.slice(0, openingAndBodyLimit).replace(/\s+\S*$/, "") + reserved);
  }

  return applyCase(out.slice(0, limit), cfg.casing);
}

function parseServiceDate(value: string): Date | null {
  const raw = value.trim();
  if (!raw) return null;
  const monthMap: Record<string, number> = {
    jan: 0, january: 0, feb: 1, february: 1, mar: 2, march: 2, apr: 3, april: 3,
    may: 4, jun: 5, june: 5, jul: 6, july: 6, aug: 7, august: 7, sep: 8,
    sept: 8, september: 8, oct: 9, october: 9, nov: 10, november: 10, dec: 11, december: 11,
  };
  const monthYear = raw.match(/\b([A-Za-z]+)\s+(\d{4})\b/);
  if (monthYear) {
    const month = monthMap[monthYear[1].toLowerCase()];
    if (month !== undefined) return new Date(Number(monthYear[2]), month, 1);
  }
  const iso = raw.match(/\b(\d{4})[-/](\d{1,2})(?:[-/]\d{1,2})?\b/);
  if (iso) return new Date(Number(iso[1]), Number(iso[2]) - 1, 1);
  const short = raw.match(/\b(\d{1,2})[-/](\d{4})\b/);
  if (short) return new Date(Number(short[2]), Number(short[1]) - 1, 1);
  const parsed = new Date(raw);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

function dateOrderError(form: FormState): string | null {
  if (!form.dateFrom || !form.dateTo) return null;
  const from = parseServiceDate(form.dateFrom);
  const to = parseServiceDate(form.dateTo);
  if (!from || !to) return null;
  return to.getTime() < from.getTime() ? "End date must be after start date." : null;
}

function serviceMonths(form: FormState): number | null {
  const from = parseServiceDate(form.dateFrom);
  const to = parseServiceDate(form.dateTo);
  if (!from || !to || to.getTime() < from.getTime()) return null;
  return ((to.getFullYear() - from.getFullYear()) * 12) + (to.getMonth() - from.getMonth()) + 1;
}

function applyCase(text: string, mode: "upper" | "sentence"): string {
  return mode === "upper" ? text.toUpperCase() : text;
}

function restoreSOATerms(text: string, form: FormState): string {
  const canonicalTerms = [
    "Marine Corps",
    "Marine",
    "Marines",
    "United States Marine Corps",
    "United States Naval Service",
    "National Capital Region",
    "Summary of Action",
    "Washington, D.C.",
    "S-1",
    "EDIPI",
    "OMPF",
    "APS",
    AWARDS[form.award].label,
    form.rank,
    form.firstName,
    form.lastName,
    rankLast(form),
    displayUnit(form.unit),
  ].filter(Boolean);

  let out = text;
  for (const term of canonicalTerms) {
    const escaped = term.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    out = out.replace(new RegExp(`\\b${escaped}\\b`, "gi"), term);
  }
  return out;
}

function paragraphCase(text: string, form: FormState): string {
  const letters = text.replace(/[^A-Za-z]/g, "");
  if (!letters) return text;
  const upperRatio = letters.replace(/[^A-Z]/g, "").length / letters.length;
  if (upperRatio < 0.72) return restoreSOATerms(text, form);

  let out = text.toLowerCase();
  out = out.replace(/(^|[.!?]\s+)([a-z])/g, (_, prefix: string, letter: string) => prefix + letter.toUpperCase());
  out = out.replace(/\bi\b/g, "I");
  return restoreSOATerms(out, form);
}

function normalizeSOA(text: string, form: FormState): string {
  const headingOnly = /^(background|accomplishments?|recommendation|summary of action|soa)$/i;
  const paragraphs = enforceWashington(expandAbbr(text))
    .split(/\n{2,}/)
    .map((paragraph) => paragraph.replace(/\n+/g, " ").trim())
    .filter((paragraph) => paragraph && !headingOnly.test(paragraph.replace(/[:.]/g, "").trim()))
    .map((paragraph) => paragraphCase(cleanup(paragraph), form));

  return cleanup(paragraphs.join("\n\n"));
}

function rankLast(form: FormState): string {
  return [form.rank, form.lastName].filter(Boolean).join(" ").trim() || "[Rank Lastname]";
}

function displayUnit(unit: string): string {
  return normalizeWashingtonDC(unit || UNIT_PRESETS[0]).replace(/,\s*$/g, "");
}

function unitInSentence(unit: string): string {
  const cleanUnit = displayUnit(unit);
  return cleanUnit === UNIT_PRESETS[0] ? UNIT_CANON : `${cleanUnit},`;
}

function buildOpening(form: FormState): string {
  const p = PRONOUNS[form.pronoun];
  const billet = form.billet || "[Billet]";
  const from = form.dateFrom || "[Month Year]";
  const to = form.dateTo || "[Month Year]";
  const rl = rankLast(form);
  const unit = unitInSentence(form.unit);

  switch (form.award) {
    case "MMAST":
      return `During the period of ${from} through ${to}, ${rl} performed ${p.poss} demanding duties in an outstanding manner while serving as ${billet}, ${unit}.`;
    case "NAM":
      return `Professional achievement in the superior performance of ${p.poss} duties while serving as ${billet}, ${unit} from ${from} to ${to}.`;
    case "NMC":
      return `Meritorious service while serving as ${billet}, ${unit} from ${from} to ${to}.`;
    case "CERTCOM":
      return `Exceptional performance of ${p.poss} duties while serving as ${billet}, ${unit} from ${from} to ${to}. ${rl} performed ${p.poss} demanding duties in an exemplary and highly professional manner.`;
    case "MSM":
      return `For outstanding meritorious service while serving as ${billet}, ${unit} from ${from} to ${to}.`;
    case "LOM":
      return `For exceptionally meritorious conduct in the performance of outstanding services while serving as ${billet}, ${unit} from ${from} to ${to}.`;
    default:
      return "";
  }
}

function buildClosing(form: FormState): string {
  const p = PRONOUNS[form.pronoun];
  const rl = rankLast(form);
  const cfg = AWARDS[form.award];

  if (form.award === "MMAST") {
    return `${rl}'s initiative, perseverance, and total dedication to duty reflected credit upon ${p.obj} and were in keeping with the highest traditions of the Marine Corps and the United States Naval Service.`;
  }

  if (cfg.greatCredit || cfg.closing === "great" || cfg.closing === "loa") {
    return `${rl}'s professionalism, perseverance, and loyal dedication to duty reflected great credit on ${p.obj} and were in keeping with the highest traditions of the Marine Corps and the United States Naval Service.`;
  }
  return `By ${p.poss} ${form.attr1}, ${form.attr2}, and ${form.adj} dedication to duty, ${rl} reflected credit upon ${p.refl} and upheld the highest traditions of the Marine Corps and the United States Naval Service.`;
}

// ---- Build citation body: synthesize accomplishments into impact-focused narrative ----
function buildBody(form: FormState, forCitation = false): string {
  const lines = prioritizedAchievementLines(form, forCitation);
  if (!lines.length) {
    const p = PRONOUNS[form.pronoun];
    return `${rankLast(form)} consistently performed ${p.poss} demanding duties with exceptional skill, sound judgment, and unwavering commitment.`;
  }

  // Clean and classify each line
  const classified = classifyAll(lines);
  if (classified.length === 0) {
    const p = PRONOUNS[form.pronoun];
    return `${rankLast(form)} consistently performed ${p.poss} demanding duties with exceptional skill, sound judgment, and unwavering commitment.`;
  }

  const cleaned = classified.map((c) => {
    let s = c.text;
    s = s.charAt(0).toUpperCase() + s.slice(1);
    if (!/[.!?]$/.test(s)) s += ".";
    return s;
  });

  // If only 1-2 achievements, synthesize into a single compound sentence with connective tissue
  if (cleaned.length <= 2) {
    return cleaned.join(" Furthermore, ");
  }

  // Group achievements by category for synthesis
  const groups = new Map<AchievementCategory, string[]>();
  for (const c of classified) {
    const cat = c.category;
    if (!groups.has(cat)) groups.set(cat, []);
    let s = c.text;
    s = s.charAt(0).toUpperCase() + s.slice(1);
    if (!/[.!?]$/.test(s)) s += ".";
    groups.get(cat)!.push(s);
  }

  // Build paragraphs by category with transition phrases
  const paragraphs: string[] = [];
  const orderedCats: AchievementCategory[] = ["Leadership", "Operations", "Training", "Administration", "Innovation", "Community Relations", "Uncategorized"];
  const transitions = [
    "",
    "Furthermore,",
    "In addition,",
    "Moreover,",
  ];

  let globalIdx = 0;
  for (const cat of orderedCats) {
    const items = groups.get(cat);
    if (!items || items.length === 0) continue;

    if (items.length === 1) {
      const prefix = globalIdx === 0 ? "" : transitions[Math.min(globalIdx, transitions.length - 1)] + " ";
      paragraphs.push(prefix + items[0]);
      globalIdx++;
    } else {
      // Synthesize multiple items in same category
      const joined = items.join(" ");
      const prefix = globalIdx === 0 ? "" : transitions[Math.min(globalIdx, transitions.length - 1)] + " ";
      paragraphs.push(prefix + joined);
      globalIdx++;
    }
  }

  return paragraphs.join(" ");
}

// ---- Build SOA: narrative format with Background, Accomplishments by theme, Recommendation ----
function buildSOA(form: FormState): string {
  const name = [form.rank, form.firstName, form.lastName].filter(Boolean).join(" ").trim() || "[Rank First Last]";
  const p = PRONOUNS[form.pronoun];
  const rl = rankLast(form);
  const billet = form.billet || "[Billet]";
  const additionalBillets = additionalBilletList(form).join("; ");
  const from = form.dateFrom || "[Month Year]";
  const to = form.dateTo || "[Month Year]";
  const awardLabel = AWARDS[form.award].label;
  const unit = unitInSentence(form.unit);

  const lines = prioritizedAchievementLines(form, false);
  const classified = classifyAll(lines);

  // ---- 1. BACKGROUND PARAGRAPH ----
  const heShe = p.subj.charAt(0).toUpperCase() + p.subj.slice(1);
  const background = [
    `${name} is enthusiastically recommended for award of the ${awardLabel} ` +
      `for ${p.poss} outstanding performance of duty while serving as ${billet}, ${unit} ` +
      `from ${from} to ${to}.`,
    `Throughout ${p.poss} tour, ${rl} demonstrated exceptional leadership, technical expertise, ` +
      `and steadfast dedication, consistently surpassing the demanding standards expected of ` +
      `Marines assigned to the National Capital Region. ${heShe} performed ${p.poss} duties with ` +
      `initiative and sound judgment, earning the respect and confidence of seniors, peers, ` +
      `and subordinates alike.`,
    additionalBillets
      ? `${heShe} also assumed additional duties as ${additionalBillets}, extending ${p.poss} influence across multiple mission areas while maintaining excellence in ${p.poss} primary billet.`
      : "",
  ].join(" ");

  // ---- 2. ACCOMPLISHMENTS SECTION (grouped by theme) ----
  const groups = new Map<AchievementCategory, ClassifiedAchievement[]>();
  for (const c of classified) {
    if (!c.text) continue;
    if (!groups.has(c.category)) groups.set(c.category, []);
    groups.get(c.category)!.push(c);
  }

  const orderedCats: AchievementCategory[] = [
    "Leadership", "Operations", "Training", "Administration", "Innovation", "Community Relations", "Uncategorized",
  ];

  const accomplishmentParagraphs: string[] = [];
  for (const cat of orderedCats) {
    const items = groups.get(cat);
    if (!items || items.length === 0) continue;

    const expanded = items.map((c) => expandAbbr(c.text));

    // Write a narrative paragraph for this category
    const categoryIntros: Record<AchievementCategory, string> = {
      Leadership: "As a leader,",
      Operations: "In the realm of operations,",
      Training: "In the area of training and professional development,",
      Administration: "Demonstrating superior administrative skill,",
      Innovation: "Through keen initiative and resourcefulness,",
      "Community Relations": "In the arena of community and ceremonial engagement,",
      Uncategorized: "Furthermore,",
    };

    let paragraph = categoryIntros[cat];
    if (expanded.length === 1) {
      paragraph += " " + expanded[0].charAt(0).toLowerCase() + expanded[0].slice(1);
      if (!/[.!?]$/.test(paragraph)) paragraph += ".";
    } else {
      const sentences = expanded.map((e) => {
        let s = e.charAt(0).toLowerCase() + e.slice(1);
        if (!/[.!?]$/.test(s)) s += ".";
        return s;
      });
      paragraph += " " + sentences.join(" ");
    }
    accomplishmentParagraphs.push(paragraph);
  }

  if (accomplishmentParagraphs.length === 0) {
    accomplishmentParagraphs.push(
      `Throughout ${p.poss} tour, ${rl} performed all assigned duties with the utmost professionalism and attention to detail. ` +
        `${p.subj.charAt(0).toUpperCase() + p.subj.slice(1)} consistently delivered superior results across every task and responsibility entrusted to ${p.obj}.`
    );
  }

  // ---- 3. RECOMMENDATION PARAGRAPH ----
  const closingAttrs = form.attr1 && form.attr2 && form.adj
    ? `${p.poss} ${form.attr1}, ${form.attr2}, and ${form.adj} dedication to duty`
    : `${p.poss} exceptional and sustained performance`;

  const recommendation = [
    `In view of the foregoing, ${rl}'s exemplary record of achievement, ` +
    `${closingAttrs}, and ` +
    `significant contributions to mission accomplishment clearly merit recognition through award of the ${awardLabel}.`,
    `${heShe} is most deserving of this honor and is ` +
    `enthusiastically recommended for approval.`,
  ].join(" ");

  return [
    background,
    "",
    ...accomplishmentParagraphs,
    "",
    recommendation,
  ].join("\n\n");
}

// ---- Build Letter of Authorization for OVSM ----
function buildLOA(form: FormState): string {
  const name = [form.rank, form.firstName, form.lastName].filter(Boolean).join(" ").trim() || "[Rank First Last]";
  const rl = rankLast(form);
  const p = PRONOUNS[form.pronoun];
  const billet = form.billet || "[Billet]";
  const additionalBillets = additionalBilletList(form).join("; ");
  const from = form.dateFrom || "[Date]";
  const to = form.dateTo || "[Date]";
  const unit = displayUnit(form.unit);
  const heShe = p.subj.charAt(0).toUpperCase() + p.subj.slice(1);

  const lines = prioritizedAchievementLines(form, false);
  const classified = classifyAll(lines);

  // ---- 1. HEADER - Marine identification ----
  const header = [
    `LETTER OF AUTHORIZATION`,
    `FOR AWARD OF THE`,
    `OUTSTANDING VOLUNTEER SERVICE MEDAL`,
  ].join("\n");

  // ---- 2. IDENTIFICATION BLOCK ----
  const identification = [
    `${name}, United States Marine Corps, ${unit}, is`,
    `recommended for award of the Outstanding Volunteer Service Medal for sustained`,
    `and exemplary volunteer service to the community during the period ${from} to ${to}.`,
  ].join(" ");

  // ---- 3. VOLUNTEER SERVICE NARRATIVE ----
  const serviceParagraphs: string[] = [];

  // Description of volunteer activities from achievements
  if (classified.length > 0) {
    serviceParagraphs.push(
      `${rl} dedicated ${p.poss} personal time while serving as ${billet} at ${unit}` +
      `${additionalBillets ? ` and while carrying additional duties as ${additionalBillets}` : ""}` +
      ` to the following volunteer activities:`
    );

    const cleaned = classified.map((c) => {
      let s = expandAbbr(c.text);
      s = s.charAt(0).toUpperCase() + s.slice(1);
      if (!/[.!?]$/.test(s)) s += ".";
      return s;
    });

    // Group into paragraphs
    if (cleaned.length <= 2) {
      serviceParagraphs.push(cleaned.join(" Additionally, "));
    } else {
      // Group by category for organized narrative
      const groups = new Map<AchievementCategory, string[]>();
      for (const c of classified) {
        const cat = c.category;
        if (!groups.has(cat)) groups.set(cat, []);
        let s = expandAbbr(c.text);
        s = s.charAt(0).toUpperCase() + s.slice(1);
        if (!/[.!?]$/.test(s)) s += ".";
        groups.get(cat)!.push(s);
      }

      const orderedCats: AchievementCategory[] = [
        "Leadership", "Operations", "Community Relations", "Training", "Administration", "Innovation", "Uncategorized",
      ];

      const categoryIntros: Record<AchievementCategory, string> = {
        Leadership: "In a leadership capacity,",
        Operations: "Through direct service,",
        Training: "In mentoring and training,",
        Administration: "Providing organizational support,",
        Innovation: "Through initiative and creativity,",
        "Community Relations": "In direct community engagement,",
        Uncategorized: "Additionally,",
      };

      for (const cat of orderedCats) {
        const items = groups.get(cat);
        if (!items || items.length === 0) continue;

        let paragraph = categoryIntros[cat];
        if (items.length === 1) {
          paragraph += " " + items[0].charAt(0).toLowerCase() + items[0].slice(1);
        } else {
          const sentences = items.map((e) => {
            let s = e.charAt(0).toLowerCase() + e.slice(1);
            if (!/[.!?]$/.test(s)) s += ".";
            return s;
          });
          paragraph += " " + sentences.join(" ");
        }
        serviceParagraphs.push(paragraph);
      }
    }
  } else {
    serviceParagraphs.push(
      `${rl} performed extensive volunteer service in the local community,` +
      ` dedicating substantial personal time to programs and organizations that` +
      ` directly benefited both civilian and military communities. ${heShe}` +
      ` consistently demonstrated the Marine Corps' commitment to community` +
      ` engagement and selfless service.`
    );
  }

  // ---- 4. IMPACT AND RECOGNITION ----
  const impact = [
    `${rl}'s sustained volunteer efforts have directly and meaningfully`,
    `benefited the community, embodying the Marine Corps ethos of service beyond self.`,
    `${heShe} has consistently upheld the highest standards of citizenship and`,
    `community involvement, serving as an outstanding representative of the`,
    `United States Marine Corps.`,
  ].join(" ");

  // ---- 5. RECOMMENDATION ----
  const recommendation = [
    `${rl} is most deserving of this recognition. ${p.poss} selfless dedication`,
    `to volunteer service reflects great credit upon ${p.refl} and is in`,
    `keeping with the highest traditions of the Marine Corps and the`,
    `United States Naval Service. Approval is enthusiastically recommended.`,
  ].join(" ");

  // Assemble: Header, Identification, Service narrative, Impact, Recommendation
  const sections = [
    header,
    "",
    identification,
    "",
    ...serviceParagraphs,
    "",
    impact,
    "",
    recommendation,
  ].filter(Boolean);

  return cleanup(sections.join("\n\n"));
}

function assembleCitation(form: FormState): string {
  const cfg = AWARDS[form.award];
  let body = buildBody(form, true);
  body = expandAbbr(body);
  let text = `${buildOpening(form)} ${body} ${buildClosing(form)}`;
  text = expandAbbr(text);
  text = enforceWashington(text);
  text = cleanup(text);
  text = applyCase(text, cfg.casing);
  return enforceCitationLimit(text, form);
}

function runChecks(citation: string, soa: string, form: FormState): CheckItem[] {
  const cfg = AWARDS[form.award];
  const checks: CheckItem[] = [];
  const dateErr = dateOrderError(form);
  if (dateErr) {
    checks.push({ status: "err", title: "Date order", detail: dateErr });
  }

  const match = awardMatchScore(form, citation || soa);
  if (match.severity !== "none") {
    checks.push({
      status: match.severity === "severe" ? "err" : "warn",
      title: match.title,
      detail: `${match.detail} ${match.recommendations.join(" ")}`,
    });
  }

  for (const issue of analyzeWeakInput(form)) {
    checks.push({ status: "warn", title: issue.title, detail: issue.detail });
  }
  const months = serviceMonths(form);
  if ((form.award === "MSM" || form.award === "LOM") && months !== null && months < 18) {
    checks.push({
      status: "warn",
      title: "Sustained meritorious period",
      detail: "Example MSM/LOM packages usually show sustained, broad impact over a longer meritorious period. Add stronger scope and duration context if this is an impact award.",
    });
  }
  if ((form.award === "MSM" || form.award === "LOM") && !VISIBILITY_TERMS.test(`${form.achievements} ${citation} ${soa}`)) {
    checks.push({
      status: "warn",
      title: "Senior-scope impact",
      detail: "Higher awards are stronger when they show command-wide, installation-wide, national, international, senior-leader, or strategic visibility.",
    });
  }
  checks.push(...analyzeRealityIssues(form));
  checks.push(...analyzeOVSMIssues(form));

  // OVSM uses LOA format — different validation
  if (cfg.isLOA) {
    if (!soa) return checks;

    // LOA content check
    if (/LETTER OF AUTHORIZATION/i.test(soa)) {
      checks.push({ status: "ok", title: "LOA header", detail: 'Contains "LETTER OF AUTHORIZATION" heading.' });
    } else {
      checks.push({ status: "err", title: "LOA header", detail: "Letter of Authorization heading missing." });
    }

    // OVSM mention
    if (/Outstanding Volunteer Service Medal/i.test(soa)) {
      checks.push({ status: "ok", title: "Award named", detail: "Outstanding Volunteer Service Medal is referenced." });
    } else {
      checks.push({ status: "warn", title: "Award named", detail: "The OVSM should be mentioned explicitly." });
    }

    // Unit formatting
    if (/Marine Barracks/i.test(soa)) {
      const badUnit = /Marine Barracks\s+Washington\s+DC\b/i.test(soa)
        || /Marine Barracks,?\s+Washington,?\s+DC\b/i.test(soa)
        || !/Marine Barracks,\s*Washington,\s*D\.C\./i.test(soa);
      checks.push(badUnit
        ? { status: "err", title: "Unit formatting", detail: 'Must read "Marine Barracks, Washington, D.C.,".', fixId: "fixUnit" }
        : { status: "ok", title: "Unit formatting", detail: 'Correct "Marine Barracks, Washington, D.C.,".' }
      );
    }

    // Abbreviations
    const found = Object.keys(EXPANSIONS).filter((abbr) =>
      new RegExp("\\b" + abbr.replace(/[.*+?^${}()|[\]\\]/g, "\\$&") + "\\b").test(soa)
    );
    checks.push(found.length
      ? { status: "err", title: "Abbreviations present", detail: "Expand: " + found.join(", ") + ".", fixId: "fixAbbr" }
      : { status: "ok", title: "No abbreviations", detail: "Only \"Washington, D.C.\" abbreviated." }
    );

    // EDIPI
    if (form.edipi) {
      checks.push(
        /^\d{10}$/.test(form.edipi)
          ? { status: "ok", title: "EDIPI", detail: "Valid 10-digit EDIPI." }
          : { status: "warn", title: "EDIPI", detail: "EDIPI should be exactly 10 digits." }
      );
    }

    // Closing sentence for OVSM
    checks.push(
      /traditions of the Marine Corps and the United States Naval Service/i.test(soa)
        ? { status: "ok", title: "Closing sentence", detail: "Required closing present." }
        : { status: "warn", title: "Closing sentence", detail: "Standard closing statement is recommended." }
    );

    return checks;
  }

  if (!citation) return checks;

  const expectedOpening = applyCase(cleanup(buildOpening(form)), cfg.casing);
  const normalizedCitation = cleanup(citation);
  if (expectedOpening && !normalizedCitation.toLowerCase().startsWith(expectedOpening.toLowerCase())) {
    checks.push({
      status: "warn",
      title: "Opening sentence",
      detail: "Citation opening no longer matches the selected award, billet, unit, and dates. Regenerate or use Fix With AI before submission.",
    });
  } else if (expectedOpening) {
    checks.push({ status: "ok", title: "Opening sentence", detail: "Opening matches selected award and service period." });
  }

  // Unit formatting
  const badUnit = /Marine Barracks\s+Washington\s+DC\b/i.test(citation)
    || /Marine Barracks,?\s+Washington,?\s+DC\b/i.test(citation)
    || (/Marine Barracks/i.test(citation) && !/Marine Barracks,\s*Washington,\s*D\.C\.,/i.test(citation));
  if (/Marine Barracks/i.test(citation)) {
    checks.push(badUnit
      ? { status: "err", title: "Unit formatting", detail: "Must read \"Marine Barracks, Washington, D.C.,\".", fixId: "fixUnit" }
      : { status: "ok", title: "Unit formatting", detail: "Correct \"Marine Barracks, Washington, D.C.,\"." }
    );
  }

  // Washington, D.C.
  if (/Washington/i.test(citation)) {
    checks.push(
      /Washington,\s*D\.C\./i.test(citation) && !/D\.C\.,[,.]/i.test(citation)
        ? { status: "ok", title: "Washington, D.C.", detail: "Properly formatted." }
        : { status: "err", title: "Washington, D.C.", detail: "Use \"Washington, D.C.\" and avoid \"D.C.,.\" or \"D.C.,,\".", fixId: "fixUnit" }
    );
  }

  // Abbreviations
  const found = Object.keys(EXPANSIONS).filter((abbr) =>
    new RegExp("\\b" + abbr.replace(/[.*+?^${}()|[\]\\]/g, "\\$&") + "\\b").test(citation)
  );
  checks.push(found.length
    ? { status: "err", title: "Abbreviations present", detail: "Expand: " + found.join(", ") + ".", fixId: "fixAbbr" }
    : { status: "ok", title: "No abbreviations", detail: "Only \"Washington, D.C.\" abbreviated." }
  );

  // Capitalization
  if (cfg.casing === "upper") {
    const alpha = citation.replace(/[^A-Za-z]/g, "");
    checks.push(alpha === alpha.toUpperCase()
      ? { status: "ok", title: "Capitalization", detail: "Correctly in ALL CAPS." }
      : { status: "err", title: "Capitalization", detail: "This award must be ALL CAPS.", fixId: "fixCase" }
    );
  } else {
    const upperRun = /[A-Z]{6,}/.test(citation.replace(/D\.C\./g, ""));
    checks.push(upperRun
      ? { status: "warn", title: "Capitalization", detail: "Sentence case expected — long all-caps runs found.", fixId: "fixCase" }
      : { status: "ok", title: "Capitalization", detail: "Sentence case as required." }
    );
  }

  // Great credit rule
  const hasGreat = /great credit/i.test(citation);
  if (cfg.greatCredit) {
    checks.push(hasGreat
      ? { status: "ok", title: "\"Great credit\"", detail: "Correct for this award level." }
      : { status: "warn", title: "\"Great credit\"", detail: `${AWARDS[form.award].label} closing should read "reflected great credit."` }
    );
  } else {
    checks.push(hasGreat
      ? { status: "err", title: "\"Great credit\" misuse", detail: "This award closing should read \"reflected credit\" without \"great.\"", fixId: "fixGreat" }
      : { status: "ok", title: "Credit phrasing", detail: "\"Reflected credit\" — correct for this award." }
    );
  }

  // Character limit
  if (cfg.maxChars) {
    if (citation.length > cfg.maxChars) {
      checks.push({ status: "err", title: "Character limit", detail: `${citation.length}/${cfg.maxChars} — over by ${citation.length - cfg.maxChars}.` });
    } else if (cfg.target && citation.length < cfg.target[0]) {
      checks.push({ status: "warn", title: "Length", detail: `${citation.length} chars — below target ${cfg.target[0]}–${cfg.target[1]}.` });
    } else {
      checks.push({ status: "ok", title: "Character limit", detail: `${citation.length}/${cfg.maxChars} — within target.` });
    }
  }

  if (form.award === "MMAST") {
    const lines = estimatedCitationLines(citation, 78);
    checks.push(lines <= 14
      ? { status: "ok", title: "Meritorious Mast format", detail: "Citation-only recognition; not processed through APS and should be forwarded for OMPF entry." }
      : { status: "warn", title: "Meritorious Mast format", detail: `Estimated ${lines}/14 lines. Shorten to fit the portrait Meritorious Mast certificate format.` }
    );
  }

  if (form.award === "CERTCOM") {
    const lines = estimatedCitationLines(citation, 128);
    checks.push(lines <= 9
      ? { status: "ok", title: "CertCom format", detail: "Citation-only recognition; use all caps and forward for OMPF entry." }
      : { status: "warn", title: "CertCom format", detail: `Estimated ${lines}/9 lines. Shorten to fit the landscape Certificate of Commendation format.` }
    );
  }

  // Closing sentence
  checks.push(
    /traditions of the Marine Corps and the United States Naval Service/i.test(citation)
      ? { status: "ok", title: "Closing sentence", detail: "Required closing present." }
      : { status: "err", title: "Closing sentence", detail: "Required closing statement missing." }
  );

  // EDIPI
  if (form.edipi) {
    checks.push(
      /^\d{10}$/.test(form.edipi)
        ? { status: "ok", title: "EDIPI", detail: "Valid 10-digit EDIPI." }
        : { status: "warn", title: "EDIPI", detail: "EDIPI should be exactly 10 digits." }
    );
  }

  return checks;
}

function autoFixCitation(text: string, fixId: string, casing: "upper" | "sentence"): string {
  if (fixId === "fixUnit") text = enforceWashington(text);
  if (fixId === "fixAbbr") text = expandAbbr(text);
  if (fixId === "fixGreat") text = text.replace(/reflected great credit/gi, "reflected credit");
  if (fixId === "fixCase") text = applyCase(text, casing);
  return applyCase(cleanup(enforceWashington(expandAbbr(text))), casing);
}

// ---- Default form ----
const DEFAULT_FORM: FormState = {
  award: "NAM",
  rank: "Sergeant",
  lastName: "",
  firstName: "",
  edipi: "",
  pronoun: "m",
  billet: "",
  additionalBillets: "",
  unit: UNIT_CANON.replace(/,$/, ""),
  dateFrom: "",
  dateTo: "",
  attr1: "initiative",
  attr2: "professionalism",
  adj: "unwavering",
  achievements: "",
};

// ---- Storage ----
function loadForm(): FormState {
  return { ...DEFAULT_FORM };
}

function saveForm(form: FormState, soa: string, citation: string) {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify({ ...form, _soa: soa, _citation: citation }));
  } catch { /* ignore */ }
}

function loadSavedDrafts(): SavedDraft[] {
  try {
    const raw = localStorage.getItem(SAVED_DRAFTS_KEY);
    if (!raw) return [];
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function persistSavedDrafts(drafts: SavedDraft[]) {
  try {
    localStorage.setItem(SAVED_DRAFTS_KEY, JSON.stringify(drafts));
  } catch { /* ignore */ }
}

function draftLabel(form: FormState): string {
  const who = [form.rank, form.lastName].filter(Boolean).join(" ").trim() || "Untitled award";
  return `${who} ${form.award}`.trim();
}

function makeDraft(form: FormState, soa: string, citation: string, aiNotes: string[], existing?: SavedDraft): SavedDraft {
  const now = new Date().toISOString();
  return {
    id: existing?.id || `draft-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
    name: existing?.name || draftLabel(form),
    form: { ...form },
    soa,
    citation,
    aiNotes: [...aiNotes],
    createdAt: existing?.createdAt || now,
    updatedAt: now,
  };
}

function formatDraftDate(iso: string): string {
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return "";
  return d.toLocaleString(undefined, { month: "short", day: "numeric", year: "numeric", hour: "numeric", minute: "2-digit" });
}

// ---- Sub-components ----
function ScoreRing({ pct, errors }: { pct: number; errors: number }) {
  const r = 22;
  const circ = 2 * Math.PI * r;
  const stroke = errors ? "#b3261e" : pct < 100 ? "#b5751a" : "#2f7d44";
  return (
    <div className="relative w-[52px] h-[52px] shrink-0">
      <svg width="52" height="52" className="-rotate-90">
        <circle cx="26" cy="26" r={r} stroke="#e9e3d6" strokeWidth="6" fill="none" />
        <circle
          cx="26" cy="26" r={r}
          stroke={stroke}
          strokeWidth="6"
          fill="none"
          strokeLinecap="round"
          strokeDasharray={circ.toFixed(1)}
          strokeDashoffset={(circ * (1 - pct / 100)).toFixed(1)}
        />
      </svg>
      <div className="absolute inset-0 grid place-items-center font-extrabold text-[15px] text-[#11161d]">
        {pct}%
      </div>
    </div>
  );
}

function CheckRow({ check, onFix }: { check: CheckItem; onFix: (fixId: string) => void }) {
  const icon = { ok: "\u2713", warn: "!", err: "\u2715" };
  const bg = { ok: "bg-[#2f7d44]", warn: "bg-[#b5751a]", err: "bg-[#b3261e]" };
  const border = { ok: "border-[#e0efe4]", warn: "border-[#f5e8d3]", err: "border-[#fce4e2]" };
  const bg2 = { ok: "bg-[#f6fbf8]", warn: "bg-[#fdf8f0]", err: "bg-[#fef7f6]" };
  return (
    <li className={`flex gap-[9px] items-start text-[12.5px] leading-[1.4] p-[9px_11px] rounded-[9px] border ${border[check.status]} ${bg2[check.status]}`}>
      <div className={`shrink-0 w-4 h-4 rounded-full mt-px grid place-items-center text-white text-[10px] font-extrabold ${bg[check.status]}`}>
        {icon[check.status]}
      </div>
      <div className="flex-1">
        <b className="block text-[#11161d] mb-px text-[12.5px]">{check.title}</b>
        <span className="text-[#6b6f76]">{check.detail}</span>
      </div>
      {check.fixId && (
        <button
          onClick={() => onFix(check.fixId!)}
          className="shrink-0 text-[12px] font-semibold px-[11px] py-[6px] rounded-[9px] bg-[#f0ece2] text-[#3a414b] hover:bg-[#e7e1d4] transition-colors"
        >
          Fix
        </button>
      )}
    </li>
  );
}

// ---- Main Component ----
export default function Index() {
  const [form, setForm] = useState<FormState>(loadForm);
  const [soa, setSoa] = useState<string>("");
  const [citation, setCitation] = useState<string>("");
  const [checks, setChecks] = useState<CheckItem[]>([]);
  const [aiAvailable, setAiAvailable] = useState<boolean>(false);
  const [aiLoading, setAiLoading] = useState<boolean>(false);
  const [aiEnhancement, setAiEnhancement] = useState<boolean>(true);
  const [aiNotes, setAiNotes] = useState<string[]>([]);
  const [aiFixSummary, setAiFixSummary] = useState<{ count: number; details: string[]; open: boolean } | null>(null);
  const [spellMode, setSpellMode] = useState<boolean>(false);
  const [aiBanner, setAiBanner] = useState<{ show: boolean; over: boolean; message: string }>({ show: false, over: false, message: "" });
  const [classifiedAchievements, setClassifiedAchievements] = useState<ClassifiedAchievement[]>([]);
  const [qualityScores, setQualityScores] = useState<QualityScores | null>(null);
  const [exportOpen, setExportOpen] = useState<boolean>(false);
  const [showSplash, setShowSplash] = useState<boolean>(true);
  const [showStartupDialog, setShowStartupDialog] = useState<boolean>(false);
  const [showReleaseNotice, setShowReleaseNotice] = useState<boolean>(false);
  const [savedDrafts, setSavedDrafts] = useState<SavedDraft[]>(loadSavedDrafts);
  const [activeDraftId, setActiveDraftId] = useState<string | null>(null);
  const [showCatPopup, setShowCatPopup] = useState<boolean>(false);
  const [catEasterEggDismissed, setCatEasterEggDismissed] = useState<boolean>(false);
  const hasRestored = useRef(false);
  const exportRef = useRef<HTMLDivElement>(null);

  const cfg = AWARDS[form.award];
  const p = PRONOUNS[form.pronoun];
  const dateErr = dateOrderError(form);
  const currentAwardMatch = awardMatchScore(form, citation || soa);
  const currentWeakInput = analyzeWeakInput(form);
  const hasHardValidationError = Boolean(dateErr);
  const unitIsPreset = (UNIT_PRESETS as readonly string[]).includes(displayUnit(form.unit));
  const awardConcernKind = currentAwardMatch.severity === "none"
    ? "none"
    : awardLevel(currentAwardMatch.recommendedAward) > awardLevel(form.award)
      ? "upgrade"
      : "downgrade";

  // Show splash for 3.5s, then check for saved draft
  useEffect(() => {
    const timer = setTimeout(() => {
      setShowSplash(false);
      try {
        if (localStorage.getItem(RELEASE_NOTICE_KEY) !== "1") {
          setShowReleaseNotice(true);
        }
      } catch {
        setShowReleaseNotice(true);
      }
      if (hasRestored.current) return;
      hasRestored.current = true;
      try {
        const raw = localStorage.getItem(STORAGE_KEY);
        if (raw) {
          const saved = JSON.parse(raw);
          const hasFormData = saved.lastName || saved.edipi || saved.achievements || saved._soa || saved._citation;
          if (hasFormData) {
            setShowStartupDialog(true);
            return;
          }
        }
        if (loadSavedDrafts().length) setShowStartupDialog(true);
      } catch { /* ignore */ }
    }, 3500);
    return () => clearTimeout(timer);
  }, []);

  function handleAcceptReleaseNotice() {
    try { localStorage.setItem(RELEASE_NOTICE_KEY, "1"); } catch { /* ignore */ }
    setShowReleaseNotice(false);
  }

  // Close export dropdown on outside click
  useEffect(() => {
    if (!exportOpen) return;
    const handler = (e: MouseEvent) => {
      if (exportRef.current && !exportRef.current.contains(e.target as Node)) {
        setExportOpen(false);
      }
    };
    document.addEventListener("mousedown", handler);
    return () => document.removeEventListener("mousedown", handler);
  }, [exportOpen]);

  function handleRestoreDraft() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (raw) {
        const saved = JSON.parse(raw);
        const restored = { ...DEFAULT_FORM, ...saved };
        setForm(restored);
        if (saved._soa) setSoa(saved._soa);
        if (saved._citation) {
          setCitation(saved._citation);
          if (saved._soa || saved._citation) {
            setChecks(runChecks(saved._citation || "", saved._soa || "", restored));
          }
        }
        setActiveDraftId(null);
      }
    } catch { /* ignore */ }
    setShowStartupDialog(false);
    toast.success("Previous draft restored");
  }

  function handleStartNew() {
    try { localStorage.removeItem(STORAGE_KEY); } catch { /* ignore */ }
    setForm({ ...DEFAULT_FORM });
    setSoa("");
    setCitation("");
    setChecks([]);
    setAiNotes([]);
    setAiFixSummary(null);
    setClassifiedAchievements([]);
    setQualityScores(null);
    setActiveDraftId(null);
    setShowCatPopup(false);
    setCatEasterEggDismissed(false);
    setShowStartupDialog(false);
  }

  function handleNewAward() {
    if (form.lastName || form.achievements) {
      if (!window.confirm("Start a completely new award? Current unsaved fields will be cleared. Saved drafts will remain.")) return;
    }
    try { localStorage.removeItem(STORAGE_KEY); } catch { /* ignore */ }
    setForm({ ...DEFAULT_FORM });
    setSoa("");
    setCitation("");
    setChecks([]);
    setAiNotes([]);
    setAiFixSummary(null);
    setClassifiedAchievements([]);
    setQualityScores(null);
    setAiBanner({ show: false, over: false, message: "" });
    setActiveDraftId(null);
    setShowCatPopup(false);
    setCatEasterEggDismissed(false);
    toast.success("New award started");
  }

  function handleOpenDraftsFromStartup() {
    setShowStartupDialog(false);
    window.setTimeout(() => document.getElementById("saved-drafts")?.scrollIntoView({ behavior: "smooth", block: "start" }), 50);
  }

  function handleSaveDraft() {
    const existing = activeDraftId ? savedDrafts.find((d) => d.id === activeDraftId) : undefined;
    const draft = makeDraft(form, soa, citation, aiNotes, existing);
    const next = existing
      ? savedDrafts.map((d) => d.id === draft.id ? draft : d)
      : [draft, ...savedDrafts];
    setSavedDrafts(next);
    persistSavedDrafts(next);
    setActiveDraftId(draft.id);
    toast.success(existing ? "Draft updated" : "Draft saved");
  }

  function handleRenameDraft(id: string) {
    const draft = savedDrafts.find((d) => d.id === id);
    if (!draft) return;
    const name = window.prompt("Rename draft", draft.name);
    if (!name || !name.trim()) return;
    const next = savedDrafts.map((d) => d.id === id ? { ...d, name: name.trim(), updatedAt: new Date().toISOString() } : d);
    setSavedDrafts(next);
    persistSavedDrafts(next);
    toast.success("Draft renamed");
  }

  function handleOpenSavedDraft(id: string) {
    const draft = savedDrafts.find((d) => d.id === id);
    if (!draft) return;
    setForm({ ...DEFAULT_FORM, ...draft.form });
    setSoa(draft.soa || "");
    setCitation(draft.citation || "");
    setAiNotes(draft.aiNotes || []);
    setAiFixSummary(null);
    setChecks(runChecks(draft.citation || "", draft.soa || "", { ...DEFAULT_FORM, ...draft.form }));
    setQualityScores(draft.citation ? scoreQuality(draft.citation, { ...DEFAULT_FORM, ...draft.form }) : null);
    setClassifiedAchievements(classifyAll(achievementLines(draft.form)));
    setAiBanner({ show: false, over: false, message: "" });
    setActiveDraftId(id);
    setShowCatPopup(false);
    setCatEasterEggDismissed(false);
    toast.success("Draft opened");
  }

  function handleDuplicateDraft(id: string) {
    const draft = savedDrafts.find((d) => d.id === id);
    if (!draft) return;
    const copy = makeDraft(draft.form, draft.soa, draft.citation, draft.aiNotes);
    copy.name = `${draft.name} copy`;
    const next = [copy, ...savedDrafts];
    setSavedDrafts(next);
    persistSavedDrafts(next);
    toast.success("Draft duplicated");
  }

  function handleDeleteDraft(id: string) {
    if (!window.confirm("Delete this saved draft?")) return;
    const next = savedDrafts.filter((d) => d.id !== id);
    setSavedDrafts(next);
    persistSavedDrafts(next);
    if (activeDraftId === id) setActiveDraftId(null);
    toast.success("Draft deleted");
  }

  function handleClearLocalDrafts() {
    if (!window.confirm("Clear autosave and all saved drafts from this browser? This cannot be undone.")) return;
    try {
      localStorage.removeItem(STORAGE_KEY);
      localStorage.removeItem(SAVED_DRAFTS_KEY);
    } catch { /* ignore */ }
    setSavedDrafts([]);
    setActiveDraftId(null);
    setSoa("");
    setCitation("");
    setChecks([]);
    setAiNotes([]);
    setAiFixSummary(null);
    toast.success("Local drafts cleared");
  }

  // Check AI status
  useEffect(() => {
    fetch("/api/health")
      .then((r) => r.json())
      .then((d) => setAiAvailable(Boolean(d.aiAvailable)))
      .catch(() => setAiAvailable(false));
  }, []);

  // Persist on form change
  const updateForm = useCallback((patch: Partial<FormState>) => {
    setForm((prev) => ({ ...prev, ...patch }));
  }, []);

  useEffect(() => {
    const timer = setTimeout(() => saveForm(form, soa, citation), 400);
    return () => clearTimeout(timer);
  }, [form, soa, citation]);

  useEffect(() => {
    if (catEasterEggDismissed) return;
    const timer = setTimeout(() => {
      const trigger = /(?:^|[^A-Za-z])(?:fostered\s+cats|fostered\s+kittens|litter\s+box|litterbox|cats|cat|kittens|kitten)(?=$|[^A-Za-z])/i;
      if (trigger.test(form.achievements)) {
        setShowCatPopup(true);
      }
    }, 900);
    return () => clearTimeout(timer);
  }, [form.achievements, catEasterEggDismissed]);

  // Keyboard shortcut
  useEffect(() => {
    const handler = (e: KeyboardEvent) => {
      if ((e.metaKey || e.ctrlKey) && e.key === "Enter") {
        e.preventDefault();
        handleGenerate();
      }
    };
    document.addEventListener("keydown", handler);
    return () => document.removeEventListener("keydown", handler);
    // Keep the shortcut bound to the current form state without rebinding on unrelated UI state.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [form]);

  async function requestAIImprove(payload: Record<string, unknown>) {
    const safePayload = redactSensitiveForAI(payload) as Record<string, unknown>;
    const r = await fetch("/api/improve", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify(safePayload),
    });
    const d = await r.json();
    if (!r.ok) throw new Error(d.error || "AI request failed");
    return d;
  }

  function aiContextPayload() {
    return {
      primaryBillet: form.billet,
      additionalBillets: form.additionalBillets,
      achievements: form.achievements,
      metricsToPreserve: metricsToPreserve(form),
    };
  }

  function currentFindingSummary() {
    const validationFindings = checks
      .filter((check) => check.status !== "ok")
      .map((check) => `${check.title}: ${check.detail}`);
    const awardJustificationFindings = [
      `Award Justification: ${currentAwardMatch.score}%`,
      currentAwardMatch.detail,
      ...currentAwardMatch.recommendations,
    ].filter(Boolean);
    const realityFindings = checks
      .filter((check) => /Reality Check/i.test(check.title))
      .map((check) => check.detail);
    return { validationFindings, awardJustificationFindings, realityFindings };
  }

  async function handleFixWithAI() {
    if (!soa && !citation) {
      toast.error("Generate an award package first");
      return;
    }
    if (!aiAvailable) {
      toast.error("AI unavailable. Generate Award Package can still use standard drafting mode.");
      return;
    }

    setAiLoading(true);
    const beforeIssues = checks.filter((check) => check.status !== "ok").length;
    const findings = currentFindingSummary();

    try {
      const d = await requestAIImprove({
        mode: cfg.isLOA ? "loa" : "all",
        award: form.award,
        soa: cfg.citationOnly ? "" : soa,
        citation: cfg.isLOA ? "" : citation,
        opening: cfg.isLOA ? "" : applyCase(buildOpening(form), cfg.casing),
        closing: cfg.isLOA ? "" : applyCase(buildClosing(form), cfg.casing),
        charLimit: cfg.maxChars || 0,
        targetLow: cfg.target?.[0] || 0,
        ...aiContextPayload(),
        validationFindings: findings.validationFindings,
        awardJustificationFindings: findings.awardJustificationFindings,
        realityFindings: findings.realityFindings,
      });

      const nextSoa = cfg.citationOnly ? "" : cfg.isLOA
        ? cleanup(expandAbbr(d.loa || soa))
        : normalizeSOA(d.soa || soa, form);
      const nextCitation = cfg.isLOA ? "" : enforceCitationLimit(expandAbbr(d.citation || citation), form);
      const nextChecks = runChecks(nextCitation, nextSoa, form);
      const afterIssues = nextChecks.filter((check) => check.status !== "ok").length;
      const details = Array.isArray(d.notes)
        ? d.notes.map((note: unknown) => String(note)).slice(0, 8)
        : [];
      const improvementCount = Math.max(1, Math.min(9, details.length || beforeIssues - afterIssues || 1));

      setSoa(nextSoa);
      setCitation(nextCitation);
      setChecks(nextChecks);
      setQualityScores(nextCitation ? scoreQuality(nextCitation, form) : null);
      setAiNotes(details);
      setAiFixSummary({
        count: improvementCount,
        details: details.length ? details : ["Updated wording, formatting, and validation alignment."],
        open: false,
      });
      toast.success(`AI applied ${improvementCount} improvement${improvementCount === 1 ? "" : "s"}`);
    } catch (err) {
      toast.error(err instanceof Error ? err.message : "AI request failed");
    } finally {
      setAiLoading(false);
    }
  }

  async function handleGenerate() {
    if (dateErr) {
      setChecks(runChecks(citation, soa, form));
      toast.error(dateErr);
      return;
    }
    const missing: string[] = [];
    if (!form.lastName) missing.push("last name");
    if (!form.billet) missing.push("billet");
    if (!form.dateFrom) missing.push("start date");
    if (!form.dateTo) missing.push("end date");
    if (missing.length) {
      toast.error("Add: " + missing.join(", "));
    }

    setAiLoading(Boolean(aiEnhancement && aiAvailable));
    const collectedNotes: string[] = [];

    // OVSM uses Letter of Authorization format, not SOA + Citation
    if (cfg.isLOA) {
      let loa = buildLOA(form);

      if (aiEnhancement && aiAvailable) {
        try {
          const improved = await requestAIImprove({ mode: "loa", award: form.award, soa: loa, ...aiContextPayload() });
          loa = cleanup(expandAbbr(improved.loa || loa));
          if (Array.isArray(improved.notes)) collectedNotes.push(...improved.notes.map((n: unknown) => String(n)));
        } catch (err) {
          toast.error(err instanceof Error ? err.message : "AI request failed");
        }
      } else if (aiEnhancement && !aiAvailable) {
        setAiBanner({ show: true, over: false, message: "AI unavailable. Using standard drafting mode." });
        toast.info("AI unavailable. Using standard drafting mode.");
      }

      // Classify achievements
      const lines = form.achievements.split("\n").map((l) => l.trim()).filter(Boolean);
      setClassifiedAchievements(classifyAll(lines));

      setSoa(loa); // LOA stored in SOA field for consistency
      setCitation("");
      setChecks(runChecks("", loa, form));
      setQualityScores(null); // No citation quality score for LOA
      setAiNotes(collectedNotes.slice(0, 8));
      setAiFixSummary(null);
      if (!aiEnhancement || aiAvailable) setAiBanner({ show: false, over: false, message: "" });
      setAiLoading(false);
      toast.success(aiEnhancement && aiAvailable ? "Award package generated with AI" : "Award package generated");
      return;
    }

    let newSoa = cfg.citationOnly ? "" : normalizeSOA(buildSOA(form), form);
    if (!cfg.citationOnly && aiEnhancement && aiAvailable) {
      try {
        const improved = await requestAIImprove({ mode: "soa", award: form.award, soa: newSoa, ...aiContextPayload() });
        newSoa = normalizeSOA(improved.soa || newSoa, form);
        if (Array.isArray(improved.notes)) collectedNotes.push(...improved.notes.map((n: unknown) => String(n)));
      } catch (err) {
        toast.error(err instanceof Error ? err.message : "AI request failed");
      }
    } else if (!cfg.citationOnly && aiEnhancement && !aiAvailable) {
      setAiBanner({ show: true, over: false, message: "AI unavailable. Using standard drafting mode." });
      toast.info("AI unavailable. Using standard drafting mode.");
    }

    let newCitation = assembleCitation(form);

    // Auto-expand citation if below target range (no AI needed)
    if (cfg.target && newCitation.length < cfg.target[0] && form.achievements.trim()) {
      const expanded = expandCitationLocally(newCitation, cfg.target[0], cfg.maxChars, form);
      if (expanded.length > newCitation.length) {
        newCitation = enforceCitationLimit(expanded, form);
      }
    }

    if (aiEnhancement && aiAvailable && citation !== newCitation) {
      try {
        const expanded = await requestAIImprove({
          mode: "expand",
          award: form.award,
          citation: newCitation,
          opening: applyCase(buildOpening(form), cfg.casing),
          closing: applyCase(buildClosing(form), cfg.casing),
          charLimit: cfg.maxChars || 0,
          targetLow: cfg.target?.[0] || 0,
          ...aiContextPayload(),
        });
        newCitation = enforceCitationLimit(expandAbbr(expanded.citation || newCitation), form);
        if (Array.isArray(expanded.notes)) collectedNotes.push(...expanded.notes.map((n: unknown) => String(n)));
      } catch (err) {
        toast.error(err instanceof Error ? err.message : "AI request failed");
      }
    }

    // Classify achievements
    const lines = form.achievements.split("\n").map((l) => l.trim()).filter(Boolean);
    setClassifiedAchievements(classifyAll(lines));

    // Compute quality score
    newCitation = enforceCitationLimit(newCitation, form);
    setQualityScores(scoreQuality(newCitation, form));

    setSoa(newSoa);
    setCitation(newCitation);
    setChecks(runChecks(newCitation, newSoa, form));
    setAiNotes(collectedNotes.slice(0, 8));
    setAiFixSummary(null);
    if (!aiEnhancement || aiAvailable) setAiBanner({ show: false, over: false, message: "" });
    setAiLoading(false);
    toast.success(aiEnhancement && aiAvailable ? "Award package generated with AI" : "Award package generated");
  }

  function handleClear() {
    if (!window.confirm("Clear the form? Saved drafts will remain.")) return;
    setForm({ ...DEFAULT_FORM });
    setSoa("");
    setCitation("");
    setChecks([]);
    setAiNotes([]);
    setAiFixSummary(null);
    setClassifiedAchievements([]);
    setQualityScores(null);
    setAiBanner({ show: false, over: false, message: "" });
    setActiveDraftId(null);
    try { localStorage.removeItem(STORAGE_KEY); } catch { /* ignore */ }
    toast.success("Form cleared");
  }

  function handleAutoFix(fixId: string) {
    if (!citation) return;
    const fixed = enforceCitationLimit(autoFixCitation(citation, fixId, cfg.casing), form);
    setCitation(fixed);
    setChecks(runChecks(fixed, soa, form));
    toast.success("Auto-fix applied");
  }

  function handleRevalidate() {
    const limited = enforceCitationLimit(citation, form);
    if (limited !== citation) setCitation(limited);
    setChecks(runChecks(limited, soa, form));
    setQualityScores(scoreQuality(limited, form));
  }

  function handleCopy(kind: "soa" | "citation") {
    const text = kind === "soa" ? soa : citation;
    if (!text) { toast.error("Nothing to copy yet"); return; }
    navigator.clipboard.writeText(text).then(
      () => toast.success((kind === "soa" ? "SOA" : "Citation") + " copied"),
      () => {
        const ta = document.createElement("textarea");
        ta.value = text;
        document.body.appendChild(ta);
        ta.select();
        document.execCommand("copy");
        document.body.removeChild(ta);
        toast.success("Copied");
      }
    );
  }

  function handleExportWord() {
    if (!soa && !citation) { toast.error("Generate a draft first"); return; }
    setExportOpen(false);

    const name = [form.firstName, form.lastName].filter(Boolean).join(" ") || "[Name]";
    const awardLabel = AWARDS[form.award].label;
    const awardKey = form.award;
    const fileName = `${form.lastName || "Draft"}_${awardKey}`;
    const isLOA = cfg.isLOA;
    const isCitationOnly = Boolean(cfg.citationOnly);

    const children: Paragraph[] = [
      new Paragraph({
        children: [
          new TextRun({ text: `EDIPI: ${form.edipi || "[EDIPI]"}`, bold: true, font: "Times New Roman", size: 22 }),
        ],
        spacing: { after: 60 },
      }),
      new Paragraph({
        children: [new TextRun({ text: `Rank: ${form.rank || "[Rank]"}`, font: "Times New Roman", size: 22 })],
        spacing: { after: 60 },
      }),
      new Paragraph({
        children: [new TextRun({ text: `Name: ${name}`, font: "Times New Roman", size: 22 })],
        spacing: { after: 60 },
      }),
      new Paragraph({
        children: [new TextRun({ text: `Award: ${awardLabel}`, font: "Times New Roman", size: 22 })],
        spacing: { after: 200 },
      }),
      new Paragraph({
        children: [new TextRun({ text: "─".repeat(50), font: "Times New Roman", size: 22 })],
        spacing: { after: 200 },
      }),
    ];

    if (isLOA) {
      // OVSM: only the Letter of Authorization
      children.push(
        new Paragraph({
          children: [new TextRun({ text: "LETTER OF AUTHORIZATION", bold: true, font: "Times New Roman", size: 24 })],
          heading: HeadingLevel.HEADING_2,
          spacing: { after: 120 },
        }),
        ...(soa ? soa.split("\n\n").filter(Boolean).map((para: string) =>
          new Paragraph({
            children: [new TextRun({ text: para.trim(), font: "Times New Roman", size: 22 })],
            spacing: { after: 120 },
          })
        ) : [new Paragraph({ children: [new TextRun({ text: "(No LOA generated)", font: "Times New Roman", size: 22, italics: true })] })])
      );
    } else if (isCitationOnly) {
      children.push(
        new Paragraph({
          children: [new TextRun({ text: "CITATION", bold: true, font: "Times New Roman", size: 24 })],
          heading: HeadingLevel.HEADING_2,
          spacing: { after: 120 },
        }),
        ...(citation ? citation.match(/.{1,500}(?:\. |$)/g)?.map((chunk: string) =>
          new Paragraph({
            children: [new TextRun({ text: chunk.trim(), font: "Times New Roman", size: 22 })],
            spacing: { after: 120 },
          })
        ) || [] : [new Paragraph({ children: [new TextRun({ text: "(No citation generated)", font: "Times New Roman", size: 22, italics: true })] })])
      );
    } else {
      // Standard: SOA + Citation
      children.push(
        new Paragraph({
          children: [new TextRun({ text: "SUMMARY OF ACTION", bold: true, font: "Times New Roman", size: 24 })],
          heading: HeadingLevel.HEADING_2,
          spacing: { after: 120 },
        }),
        ...(soa ? soa.split("\n\n").filter(Boolean).map((para: string) =>
          new Paragraph({
            children: [new TextRun({ text: para.trim(), font: "Times New Roman", size: 22 })],
            spacing: { after: 120 },
          })
        ) : [new Paragraph({ children: [new TextRun({ text: "(No SOA generated)", font: "Times New Roman", size: 22, italics: true })] })]),
        new Paragraph({
          children: [new TextRun({ text: "─".repeat(50), font: "Times New Roman", size: 22 })],
          spacing: { before: 200, after: 200 },
        }),
        new Paragraph({
          children: [new TextRun({ text: "CITATION", bold: true, font: "Times New Roman", size: 24 })],
          heading: HeadingLevel.HEADING_2,
          spacing: { after: 120 },
        }),
        ...(citation ? citation.match(/.{1,500}(?:\. |$)/g)?.map((chunk: string) =>
          new Paragraph({
            children: [new TextRun({ text: chunk.trim(), font: "Times New Roman", size: 22 })],
            spacing: { after: 60 },
          })
        ) ?? [new Paragraph({ children: [new TextRun({ text: citation, font: "Times New Roman", size: 22 })] })]
        : [new Paragraph({ children: [new TextRun({ text: "(No citation generated)", font: "Times New Roman", size: 22, italics: true })] })])
      );
    }

    const doc = new Document({
      styles: {
        default: {
          document: {
            run: { font: "Times New Roman", size: 24 },
          },
        },
      },
      sections: [{
        properties: {},
        children,
      }],
    });

    Packer.toBlob(doc).then((blob) => {
      saveAs(blob, `${fileName}.docx`);
      toast.success(`Exported ${fileName}.docx`);
    }).catch(() => {
      toast.error("Word export failed. Try PDF instead.");
    });
  }

  function handleExportPDF() {
    if (!soa && !citation) { toast.error("Generate a draft first"); return; }
    setExportOpen(false);

    const name = [form.firstName, form.lastName].filter(Boolean).join(" ") || "[Name]";
    const awardLabel = AWARDS[form.award].label;
    const awardKey = form.award;
    const fileName = `${form.lastName || "Draft"}_${awardKey}`;
    const isLOA = cfg.isLOA;
    const isCitationOnly = Boolean(cfg.citationOnly);

    const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

    let bodyContent = "";
    if (isLOA) {
      const loaParas = soa ? soa.split("\n\n").filter(Boolean).map((p) => `<p>${esc(p.trim())}</p>`).join("\n") : "<p><em>(No LOA generated)</em></p>";
      bodyContent = `<h2>LETTER OF AUTHORIZATION</h2>\n${loaParas}`;
    } else if (isCitationOnly) {
      const citeText = citation ? `<p>${esc(citation)}</p>` : "<p><em>(No citation generated)</em></p>";
      bodyContent = `<h2>CITATION</h2>\n${citeText}`;
    } else {
      const soaParas = soa ? soa.split("\n\n").filter(Boolean).map((p) => `<p>${esc(p.trim())}</p>`).join("\n") : "<p><em>(No SOA generated)</em></p>";
      const citeText = citation ? `<p>${esc(citation)}</p>` : "<p><em>(No citation generated)</em></p>";
      bodyContent = `<h2>SUMMARY OF ACTION</h2>\n${soaParas}\n<hr>\n<h2>CITATION</h2>\n${citeText}`;
    }

    const html = `<!DOCTYPE html>
<html><head><meta charset="utf-8"><title>${fileName}</title>
<style>
  body { font-family: "Times New Roman", Georgia, serif; font-size: 12pt; line-height: 1.6; color: #000; max-width: 6.5in; margin: 0.75in auto; padding: 0; }
  .header { margin-bottom: 18pt; }
  .header p { margin: 2pt 0; }
  hr { border: none; border-top: 1px solid #333; margin: 18pt 0; }
  h2 { font-size: 14pt; margin: 14pt 0 8pt 0; }
  p { margin: 0 0 8pt 0; text-align: justify; }
  @media print { body { margin: 0.75in; } }
</style></head><body>
<div class="header">
<p><strong>EDIPI:</strong> ${form.edipi || "[EDIPI]"}</p>
<p><strong>Rank:</strong> ${form.rank || "[Rank]"}</p>
<p><strong>Name:</strong> ${name}</p>
<p><strong>Award:</strong> ${awardLabel}</p>
</div>
<hr>
${bodyContent}
</body></html>`;

    const printWindow = window.open("", "_blank");
    if (printWindow) {
      printWindow.document.write(html);
      printWindow.document.close();
      printWindow.onload = () => printWindow.print();
      setTimeout(() => printWindow.print(), 500);
    } else {
      const blob = new Blob([html], { type: "text/html" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `${fileName}.html`;
      a.click();
      URL.revokeObjectURL(url);
      toast.success("Opened for print — use Save as PDF in your print dialog");
    }
  }

  function handleExportBoth() {
    handleExportWord();
    setTimeout(() => handleExportPDF(), 300);
  }

  function handleToggleSpell() {
    if (!citation) { toast.error("Generate a citation first"); return; }
    if (spellMode) {
      setSpellMode(false);
      const fixed = enforceCitationLimit(expandAbbr(citation), form);
      setCitation(fixed);
      setChecks(runChecks(fixed, soa, form));
      toast.success("Spell-check off — re-validated");
    } else {
      setSpellMode(true);
      toast.success("Spell-check on — edit inline, your browser underlines errors");
    }
  }

  async function handleImproveSOA() {
    if (!soa) { toast.error("Generate a SOA first"); return; }
    setAiLoading(true);
    try {
      const r = await fetch("/api/improve", {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: JSON.stringify({
          mode: "soa",
          award: form.award,
          soa,
          ...aiContextPayload(),
        }),
      });
      const d = await r.json();
      if (!r.ok) { toast.error(d.error || "AI request failed"); return; }

      const outSoa = normalizeSOA(d.soa || soa, form);
      setSoa(outSoa);

      const notes = Array.isArray(d.notes) ? d.notes.map((n: unknown) => String(n)).slice(0, 8) : [];
      setAiNotes(notes);
      setAiBanner({ show: true, over: false, message: "SOA wording improved. Regenerate citation to reflect changes." });
      toast.success("SOA refined");
    } catch {
      toast.error("Network error reaching AI service");
    } finally {
      setAiLoading(false);
    }
  }

  async function handleImproveLOA() {
    if (!soa) { toast.error("Generate a Letter of Authorization first"); return; }
    setAiLoading(true);
    try {
      const r = await fetch("/api/improve", {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: JSON.stringify({
          mode: "loa",
          award: form.award,
          soa,
          ...aiContextPayload(),
        }),
      });
      const d = await r.json();
      if (!r.ok) { toast.error(d.error || "AI request failed"); return; }

      const outLOA = cleanup(expandAbbr(d.loa || soa));
      setSoa(outLOA);

      const notes = Array.isArray(d.notes) ? d.notes.map((n: unknown) => String(n)).slice(0, 8) : [];
      setAiNotes(notes);
      setAiBanner({ show: true, over: false, message: "Letter of Authorization wording improved." });
      toast.success("LOA refined");
    } catch {
      toast.error("Network error reaching AI service");
    } finally {
      setAiLoading(false);
    }
  }

  async function handleExpandCitation() {
    if (!citation) { toast.error("Generate a citation first"); return; }
    if (!aiAvailable) { toast.error("AI expansion requires the AI service"); return; }
    setAiLoading(true);
    try {
      const r = await fetch("/api/improve", {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: JSON.stringify({
          mode: "expand",
          award: form.award,
          citation,
          opening: applyCase(buildOpening(form), cfg.casing),
          closing: applyCase(buildClosing(form), cfg.casing),
          charLimit: cfg.maxChars || 1500,
          targetLow: cfg.target?.[0] || 1100,
          ...aiContextPayload(),
        }),
      });
      const d = await r.json();
      if (!r.ok) { toast.error(d.error || "AI request failed"); return; }

      const outCite = enforceCitationLimit(expandAbbr(d.citation || citation), form);
      setCitation(outCite);
      setChecks(runChecks(outCite, soa, form));
      setQualityScores(scoreQuality(outCite, form));

      const notes = Array.isArray(d.notes) ? d.notes.map((n: unknown) => String(n)).slice(0, 8) : [];
      setAiNotes(notes);

      const over = cfg.maxChars > 0 && outCite.length > cfg.maxChars;
      setAiBanner({
        show: true,
        over,
        message: over
          ? `Citation expanded to ${outCite.length}/${cfg.maxChars} — over limit, trim before submitting.`
          : `Citation expanded to ${outCite.length} chars.`,
      });
      toast.success("Citation expanded");
    } catch {
      toast.error("Network error reaching AI service");
    } finally {
      setAiLoading(false);
    }
  }

  async function handleRefineAI() {
    if (!citation && !soa) { toast.error("Generate a draft first"); return; }
    setAiLoading(true);
    try {
      const mode = cfg.isLOA ? "loa" : "all";
      const r = await fetch("/api/improve", {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: JSON.stringify({
          mode,
          award: form.award,
          soa,
          citation: cfg.isLOA ? "" : citation,
          opening: cfg.isLOA ? "" : applyCase(buildOpening(form), cfg.casing),
          closing: cfg.isLOA ? "" : applyCase(buildClosing(form), cfg.casing),
          charLimit: cfg.maxChars || 0,
          targetLow: cfg.target?.[0] || 0,
          ...aiContextPayload(),
        }),
      });
      const d = await r.json();
      if (!r.ok) { toast.error(d.error || "AI request failed"); return; }

      if (cfg.isLOA) {
        const outLOA = cleanup(expandAbbr(d.loa || soa));
        setSoa(outLOA);
        setChecks(runChecks("", outLOA, form));
        setQualityScores(null);

        const notes = Array.isArray(d.notes) ? d.notes.map((n: unknown) => String(n)).slice(0, 8) : [];
        setAiNotes(notes);
        setAiBanner({
          show: true,
          over: false,
          message: "AI wording applied and re-validated. Dates, rank, and formatting were re-pinned automatically.",
        });
        toast.success("AI refinement applied");
      } else {
        const outCite = enforceCitationLimit(expandAbbr(d.citation || citation), form);
        const outSoa = normalizeSOA(d.soa || soa, form);

        setSoa(outSoa);
        setCitation(outCite);
        setChecks(runChecks(outCite, outSoa, form));
        setQualityScores(scoreQuality(outCite, form));

        const notes = Array.isArray(d.notes) ? d.notes.map((n: unknown) => String(n)).slice(0, 8) : [];
        setAiNotes(notes);

        const over = cfg.maxChars > 0 && outCite.length > cfg.maxChars;
        setAiBanner({
          show: true,
          over,
          message: over
            ? `AI draft applied, then re-validated. Citation is ${outCite.length}/${cfg.maxChars} — over limit, trim before submitting.`
            : "AI wording applied and re-validated. Opening, closing, dates, and formatting were re-pinned automatically.",
        });
        toast.success("AI refinement applied");
      }
    } catch {
      toast.error("Network error reaching AI service");
    } finally {
      setAiLoading(false);
    }
  }

  // Character counter
  const charCount = cfg.isLOA ? soa.length : citation.length;
  const maxChars = cfg.maxChars;
  const charPct = maxChars ? Math.min(100, (charCount / maxChars) * 100) : Math.min(100, (charCount / 1500) * 100);
  const charBarColor = maxChars
    ? charCount > maxChars ? "#b3261e" : cfg.target && charCount >= cfg.target[0] ? "#2f7d44" : "#b5751a"
    : "#6b6f76";

  // Validation score
  const totalChecks = checks.length;
  const passedChecks = checks.filter((c) => c.status === "ok").length;
  const errorChecks = checks.filter((c) => c.status === "err").length;
  const scorePct = totalChecks ? Math.round((passedChecks / totalChecks) * 100) : 0;
  const reviewerChecklist = [
    { label: "Draft generated", ok: Boolean(cfg.isLOA ? soa : citation), detail: cfg.isLOA ? "LOA present" : cfg.citationOnly ? "Citation/certificate present" : "SOA and citation present" },
    { label: "No blocking validation errors", ok: checks.length > 0 && errorChecks === 0, detail: errorChecks ? `${errorChecks} error${errorChecks === 1 ? "" : "s"} remaining` : "No errors found" },
    { label: "Opening and closing checked", ok: checks.some((check) => check.title === "Opening sentence" && check.status === "ok") || cfg.isLOA, detail: cfg.isLOA ? "LOA format checked" : "Opening frame validated" },
    { label: "Character or format limit checked", ok: checks.some((check) => /Character limit|format/i.test(check.title)), detail: cfg.maxChars ? `${charCount}/${cfg.maxChars}` : "Certificate/LOA format reviewed" },
    { label: "Privacy posture", ok: true, detail: "EDIPI excluded/redacted from AI requests" },
    { label: "Final human review", ok: false, detail: "S-1/adjutant and command review still required" },
  ];

  return (
    <>
      {/* Splash Screen */}
      {showSplash && (
        <div
          className="fixed inset-0 z-[100] flex flex-col items-center justify-center gap-6"
          style={{
            background: "linear-gradient(180deg, #11161d, #1b232e)",
            animation: "splashEnter .15s ease",
          }}
        >
          <div
            className="w-[100px] h-[100px] rounded-[24px] overflow-hidden"
            style={{
              background: "linear-gradient(135deg, #1b232e, #c5a44e)",
              boxShadow: "0 0 0 4px rgba(197,164,78,.25), 0 16px 48px rgba(160,23,34,.4)",
              animation: "splashPulse 3.5s ease infinite",
            }}
          >
            <img
              src="/icon.png"
              alt="CitationBuilder"
              className="w-full h-full object-cover"
            />
          </div>
          <h1
            className="text-[28px] font-bold text-white tracking-[.2px] text-center px-4"
            style={{ fontFamily: '"Inter", -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif' }}
          >
            CitationBuilder
          </h1>
          <p className="text-[14px] text-[#aeb6c2] text-center px-4 opacity-0" style={{ fontFamily: 'inherit', animation: 'fadeIn .6s ease .6s forwards' }}>
            Marine Corps award formatting, validation & drafting engine
          </p>
        </div>
      )}

      {/* Main App */}
      <div
        className="min-h-screen"
        style={{
          background: "radial-gradient(1200px 480px at 80% -10%, rgba(160,23,34,.06), transparent 60%), radial-gradient(900px 420px at 0% 0%, rgba(197,164,78,.10), transparent 55%), #f4f1ea",
          fontFamily: '"Inter", -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif',
          visibility: showSplash ? "hidden" : "visible",
          position: showSplash ? "absolute" : "relative",
        }}
      >
      {/* Header */}
      <header className="sticky top-0 z-40 border-b-[3px] border-[#c5a44e]" style={{
        background: "linear-gradient(180deg, #1b232e, #11161d)",
      }}>
        <div className="max-w-[1480px] mx-auto px-[14px] sm:px-[18px] lg:px-[22px] py-[10px] sm:py-[12px] lg:py-[14px] flex items-center gap-3 sm:gap-4">
          <img
            src="/icon.png"
            alt="CitationBuilder"
            className="w-[34px] h-[34px] sm:w-[38px] sm:h-[38px] lg:w-[42px] lg:h-[42px] rounded-[8px] sm:rounded-[10px] shrink-0 object-cover"
            style={{
              boxShadow: "inset 0 0 0 1px rgba(255,255,255,.12), 0 6px 18px rgba(160,23,34,.4)",
            }}
          />
          <div className="min-w-0">
            <h1 className="text-[15px] sm:text-[16px] lg:text-[17px] font-bold tracking-[.3px] text-white m-0 leading-tight">
              CitationBuilder <span className="text-[#e6d29a] font-semibold">{APP_VERSION}</span>
            </h1>
            <p className="mt-[1px] sm:mt-[2px] text-[10.5px] sm:text-[11.5px] lg:text-[12px] text-[#aeb6c2] m-0 hidden sm:block">
              Battalion-ready Marine Corps award formatting, validation & drafting engine
            </p>
          </div>
          <div className="flex-1" />
          <a
            href={SUPPORT_EMAIL}
            className="hidden sm:inline-flex text-[12px] font-semibold px-[11px] py-[7px] rounded-[8px] border border-[#4d5664] text-[#d8dee8] hover:text-white hover:border-[#c5a44e] transition-colors"
          >
            Report Issue
          </a>
        </div>
      </header>

      {/* V1 Privacy and Use Notice */}
      {showReleaseNotice && !showStartupDialog && (
        <div className="fixed inset-0 z-50 flex items-center justify-center" style={{ background: "rgba(17,22,29,.58)", backdropFilter: "blur(6px)" }}>
          <div className="bg-white rounded-[14px] shadow-[0_16px_48px_rgba(17,22,29,.28)] max-w-[620px] w-[92%] p-[24px]" style={{ animation: "fadeIn .2s ease" }}>
            <div className="flex items-start gap-[13px]">
              <div className="w-[42px] h-[42px] rounded-[10px] grid place-items-center shrink-0" style={{ background: "linear-gradient(160deg, #a01722, #7c0f19)" }}>
                <span className="text-[#e6d29a] font-extrabold text-[18px]">V1</span>
              </div>
              <div>
                <h2 className="m-0 text-[18px] font-bold text-[#11161d]">CitationBuilder {APP_VERSION} Release Notice</h2>
                <p className="m-0 mt-[7px] text-[13.5px] text-[#3a414b] leading-[1.5]">
                  Use this as a drafting and validation aid. Final awards still require chain-of-command, S-1/adjutant, and current SECNAV/unit SOP review.
                </p>
              </div>
            </div>
            <div className="mt-[16px] grid gap-[9px] text-[12.5px] leading-[1.45] text-[#3a414b]">
              <div className="rounded-[9px] border border-[#f0b8b3] bg-[#fef7f6] p-[10px_12px]">
                Do not enter classified information, CUI, medical/legal/disciplinary details, or sensitive operational details.
              </div>
              <div className="rounded-[9px] border border-[#dcd6c8] bg-[#faf8f3] p-[10px_12px]">
                AI refinement sends draft wording and accomplishments to the configured AI provider. EDIPI/SSN-like values are redacted before AI requests, and EDIPI is not included in AI payload fields.
              </div>
              <div className="rounded-[9px] border border-[#dcd6c8] bg-[#faf8f3] p-[10px_12px]">
                Drafts are saved locally in this browser. Clear local drafts before using a shared computer or handing the workstation to another user.
              </div>
            </div>
            <div className="mt-[16px] flex flex-wrap gap-[9px] justify-end">
              <a href={SUPPORT_EMAIL} className="text-[13px] font-semibold px-[13px] py-[9px] rounded-[9px] bg-white border border-[#dcd6c8] text-[#3a414b] hover:border-[#a01722] hover:text-[#a01722] transition-colors">
                Report an Issue
              </a>
              <button
                onClick={handleAcceptReleaseNotice}
                className="cursor-pointer text-[13px] font-semibold px-[15px] py-[9px] rounded-[9px] text-white border-none"
                style={{ background: "linear-gradient(160deg, #a01722, #7c0f19)", boxShadow: "0 6px 16px rgba(160,23,34,.22)" }}
              >
                Acknowledge and Continue
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Startup Privacy Dialog */}
      {showStartupDialog && (
        <div className="fixed inset-0 z-50 flex items-center justify-center" style={{ background: "rgba(17,22,29,.55)", backdropFilter: "blur(6px)" }}>
          <div className="bg-white rounded-[16px] shadow-[0_16px_48px_rgba(17,22,29,.28)] max-w-[460px] w-[92%] p-[28px]" style={{ animation: "fadeIn .2s ease" }}>
            <div className="text-center mb-[20px]">
              <div className="w-[56px] h-[56px] mx-auto mb-[14px] rounded-[14px] grid place-items-center" style={{ background: "linear-gradient(160deg, #a01722, #7c0f19)", boxShadow: "0 8px 20px rgba(160,23,34,.25)" }}>
                <svg width="26" height="26" viewBox="0 0 24 24" fill="none" stroke="#e6d29a" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
                  <rect x="3" y="11" width="18" height="11" rx="2" ry="2" />
                  <path d="M7 11V7a5 5 0 0 1 10 0v4" />
                </svg>
              </div>
              <h2 className="m-0 text-[18px] font-bold text-[#11161d]">Previous Draft Found</h2>
              <p className="m-0 mt-[8px] text-[13.5px] text-[#6b6f76] leading-[1.5]">
                Saved award work exists in this browser. Start blank, restore the previous autosave, or open Saved Drafts.
              </p>
            </div>
            <div className="grid gap-[10px]">
              <button
                onClick={handleRestoreDraft}
                className="cursor-pointer text-[14px] font-semibold px-[20px] py-[13px] rounded-[10px] text-white transition-transform duration-[.08s] active:translate-y-px"
                style={{ background: "linear-gradient(160deg, #a01722, #7c0f19)", boxShadow: "0 6px 16px rgba(160,23,34,.22)", border: "none" }}
              >
                Restore Previous Draft
              </button>
              <button
                onClick={handleStartNew}
                className="cursor-pointer text-[14px] font-semibold px-[20px] py-[13px] rounded-[10px] bg-white border border-[#dcd6c8] text-[#3a414b] hover:border-[#a01722] hover:text-[#a01722] transition-[border-color,color] duration-150 active:translate-y-px"
              >
                Start New Award
              </button>
              <button
                onClick={handleOpenDraftsFromStartup}
                className="cursor-pointer text-[14px] font-semibold px-[20px] py-[13px] rounded-[10px] bg-[#f6f3ea] border border-[#dcd6c8] text-[#3a414b] hover:border-[#a01722] hover:text-[#a01722] transition-[border-color,color] duration-150 active:translate-y-px"
              >
                Open Saved Drafts
              </button>
            </div>
            <p className="m-0 mt-[14px] text-[11px] text-[#aeb6c2] text-center">
              Drafts are saved only in this browser. No data is shared across users or devices.
            </p>
          </div>
        </div>
      )}

      <style>{`@keyframes fadeIn{from{opacity:0;transform:translateY(6px)}to{opacity:1;transform:translateY(0)}}@keyframes splashEnter{from{opacity:0}to{opacity:1}}@keyframes splashPulse{0%,100%{transform:scale(1)}50%{transform:scale(1.03)}}`}</style>

      {/* Main layout — responsive: stacked on mobile, 3-col on desktop */}
      <nav className="lg:hidden sticky top-[58px] z-30 max-w-[1480px] mx-auto px-[14px] pt-[10px]">
        <div className="grid grid-cols-5 gap-[5px] rounded-[10px] border border-[#dcd6c8] bg-white/95 p-[6px] shadow-[0_6px_18px_rgba(17,22,29,.08)]">
          {[
            ["details", "Details"],
            ["accomplishments", "Accomplishments"],
            ["outputs", cfg.isLOA ? "LOA" : cfg.citationOnly ? "Citation" : "SOA"],
            ["citation-output", cfg.isLOA ? "Validation" : "Citation"],
            ["validation", "Validation"],
          ].map(([id, label]) => (
            <button
              key={id}
              onClick={() => document.getElementById(id)?.scrollIntoView({ behavior: "smooth", block: "start" })}
              className="min-w-0 text-[11px] font-semibold px-[5px] py-[7px] rounded-[8px] bg-[#f6f3ea] text-[#3a414b] border border-transparent truncate"
            >
              {label}
            </button>
          ))}
        </div>
      </nav>
      <main className="cb-grid max-w-[1480px] mx-auto p-[14px] sm:p-[18px] lg:p-[22px] items-start">
        {/* ---- SECTION 1: Inputs ---- */}
        <section className="flex flex-col gap-[18px]">
          {/* Award & Service Details */}
          <div id="details" className="scroll-mt-[96px] bg-white border border-[#dcd6c8] rounded-[12px] shadow-[0_1px_2px_rgba(17,22,29,.05),0_8px_24px_rgba(17,22,29,.07)]">
            <div className="flex items-center gap-[10px] p-[14px_16px] border-b border-[#dcd6c8]">
              <span className="w-[22px] h-[22px] rounded-[6px] shrink-0 grid place-items-center text-[12px] font-extrabold bg-[#a01722] text-[#e6d29a]">1</span>
              <h2 className="m-0 text-[13px] uppercase tracking-[.12em] text-[#11161d]">Award & Service Details</h2>
            </div>
            <div className="p-4 space-y-[13px]">
              {/* Award */}
              <div>
                <label className="block text-[12px] font-semibold text-[#3a414b] mb-[5px]">
                  Award <span className="text-[#a01722]">*</span>
                </label>
                <select
                  value={form.award}
                  onChange={(e) => updateForm({ award: e.target.value as AwardKey })}
                  className="w-full text-[14px] text-[#1c222b] bg-[#fcfbf8] border border-[#dcd6c8] rounded-[9px] p-[9px_11px] focus:outline-none focus:border-[#a01722] focus:shadow-[0_0_0_3px_rgba(160,23,34,.12)] focus:bg-white transition-[border-color,box-shadow] duration-150"
                  style={{ fontFamily: "inherit" }}
                >
                  <option value="MMAST">Meritorious Mast</option>
                  <option value="CERTCOM">Certificate of Commendation (CertCom)</option>
                  <option value="OVSM">Outstanding Volunteer Service Medal (OVSM)</option>
                  <option value="NAM">Navy & Marine Corps Achievement Medal (NAM)</option>
                  <option value="NMC">Navy & Marine Corps Commendation Medal (NAVCOM)</option>
                  <option value="MSM">Meritorious Service Medal (MSM)</option>
                  <option value="LOM">Legion of Merit (LOM)</option>
                </select>
                <div className="text-[11px] text-[#6b6f76] mt-[4px]">
                  {cfg.isLOA
                    ? "Generates a Letter of Authorization instead of a citation."
                    : cfg.citationOnly
                      ? "Generates a citation/certificate only."
                    : cfg.casing === "upper"
                      ? "Citation renders in ALL CAPS."
                      : "Citation renders in sentence case."}
                </div>
              </div>

              {/* Rank + Last Name */}
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "10px" }}>
                <div>
                  <label className="block text-[12px] font-semibold text-[#3a414b] mb-[5px]">
                    Rank <span className="text-[#a01722]">*</span>
                  </label>
                  <select
                    value={form.rank}
                    onChange={(e) => updateForm({ rank: e.target.value })}
                    className="w-full text-[14px] text-[#1c222b] bg-[#fcfbf8] border border-[#dcd6c8] rounded-[9px] p-[9px_11px] focus:outline-none focus:border-[#a01722] focus:shadow-[0_0_0_3px_rgba(160,23,34,.12)] focus:bg-white transition-[border-color,box-shadow] duration-150"
                    style={{ fontFamily: "inherit" }}
                  >
                    {RANKS.map((r) => <option key={r} value={r}>{r}</option>)}
                  </select>
                </div>
                <div>
                  <label className="block text-[12px] font-semibold text-[#3a414b] mb-[5px]">
                    Last name <span className="text-[#a01722]">*</span>
                  </label>
                  <input
                    type="text"
                    value={form.lastName}
                    onChange={(e) => updateForm({ lastName: e.target.value })}
                    placeholder="Doe"
                    autoComplete="off"
                    className="w-full text-[14px] text-[#1c222b] bg-[#fcfbf8] border border-[#dcd6c8] rounded-[9px] p-[9px_11px] focus:outline-none focus:border-[#a01722] focus:shadow-[0_0_0_3px_rgba(160,23,34,.12)] focus:bg-white transition-[border-color,box-shadow] duration-150"
                    style={{ fontFamily: "inherit" }}
                  />
                </div>
              </div>
              {dateErr && (
                <div className="text-[12px] font-semibold text-[#7c1d13] bg-[#fef7f6] border border-[#f0b8b3] rounded-[9px] p-[9px_11px]">
                  {dateErr}
                </div>
              )}

              {/* First Name + EDIPI */}
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "10px" }}>
                <div>
                  <label className="block text-[12px] font-semibold text-[#3a414b] mb-[5px]">First name</label>
                  <input
                    type="text"
                    value={form.firstName}
                    onChange={(e) => updateForm({ firstName: e.target.value })}
                    placeholder="John"
                    autoComplete="off"
                    className="w-full text-[14px] text-[#1c222b] bg-[#fcfbf8] border border-[#dcd6c8] rounded-[9px] p-[9px_11px] focus:outline-none focus:border-[#a01722] focus:shadow-[0_0_0_3px_rgba(160,23,34,.12)] focus:bg-white transition-[border-color,box-shadow] duration-150"
                    style={{ fontFamily: "inherit" }}
                  />
                </div>
                <div>
                  <label className="block text-[12px] font-semibold text-[#3a414b] mb-[5px]">
                    EDIPI <span className="font-normal text-[#6b6f76]">(10 digits)</span>
                  </label>
                  <input
                    type="text"
                    value={form.edipi}
                    onChange={(e) => updateForm({ edipi: e.target.value.replace(/\D/g, "").slice(0, 10) })}
                    placeholder="1234567890"
                    inputMode="numeric"
                    maxLength={10}
                    autoComplete="off"
                    className="w-full text-[14px] text-[#1c222b] bg-[#fcfbf8] border border-[#dcd6c8] rounded-[9px] p-[9px_11px] focus:outline-none focus:border-[#a01722] focus:shadow-[0_0_0_3px_rgba(160,23,34,.12)] focus:bg-white transition-[border-color,box-shadow] duration-150"
                    style={{ fontFamily: "inherit" }}
                  />
                </div>
              </div>

              {/* Pronouns */}
              <div>
                <label className="block text-[12px] font-semibold text-[#3a414b] mb-[5px]">Pronouns</label>
                <div className="flex gap-[6px]">
                  {(["m", "f"] as PronounKey[]).map((k) => (
                    <button
                      key={k}
                      type="button"
                      onClick={() => updateForm({ pronoun: k })}
                      className="flex-1 min-w-[64px] cursor-pointer text-[13px] font-semibold py-[8px] px-[6px] rounded-[8px] border transition-all duration-[.12s]"
                      style={{
                        background: form.pronoun === k ? "#a01722" : "#fcfbf8",
                        borderColor: form.pronoun === k ? "#a01722" : "#dcd6c8",
                        color: form.pronoun === k ? "#fff" : "#3a414b",
                        boxShadow: form.pronoun === k ? "0 4px 12px rgba(160,23,34,.25)" : "none",
                      }}
                    >
                      {k === "m" ? "He / His" : "She / Her"}
                    </button>
                  ))}
                </div>
              </div>

              {/* Billet */}
              <div>
                <label className="block text-[12px] font-semibold text-[#3a414b] mb-[5px]">
                  Billet <span className="text-[#a01722]">*</span>
                </label>
                <input
                  type="text"
                  value={form.billet}
                  onChange={(e) => updateForm({ billet: e.target.value })}
                  placeholder="Platoon Sergeant"
                  autoComplete="off"
                  className="w-full text-[14px] text-[#1c222b] bg-[#fcfbf8] border border-[#dcd6c8] rounded-[9px] p-[9px_11px] focus:outline-none focus:border-[#a01722] focus:shadow-[0_0_0_3px_rgba(160,23,34,.12)] focus:bg-white transition-[border-color,box-shadow] duration-150"
                  style={{ fontFamily: "inherit" }}
                />
                <div className="text-[11px] text-[#6b6f76] mt-[4px]">Spell out fully — no abbreviations.</div>
              </div>

              <div>
                <label className="block text-[12px] font-semibold text-[#3a414b] mb-[5px]">
                  Additional Billets / Collateral Duties
                </label>
                <input
                  type="text"
                  value={form.additionalBillets}
                  onChange={(e) => updateForm({ additionalBillets: e.target.value })}
                  placeholder="Acting Company Administrative Chief; Funeral Bugler"
                  autoComplete="off"
                  className="w-full text-[14px] text-[#1c222b] bg-[#fcfbf8] border border-[#dcd6c8] rounded-[9px] p-[9px_11px] focus:outline-none focus:border-[#a01722] focus:shadow-[0_0_0_3px_rgba(160,23,34,.12)] focus:bg-white transition-[border-color,box-shadow] duration-150"
                  style={{ fontFamily: "inherit" }}
                />
                <div className="text-[11px] text-[#6b6f76] mt-[4px]">Optional. Used in the body when relevant.</div>
              </div>

              {/* Unit */}
              <div>
                <label className="block text-[12px] font-semibold text-[#3a414b] mb-[5px]">Unit</label>
                <select
                  value={unitIsPreset ? displayUnit(form.unit) : "Custom Unit"}
                  onChange={(e) => updateForm({ unit: e.target.value === "Custom Unit" ? "" : e.target.value })}
                  className="w-full text-[14px] text-[#1c222b] bg-[#fcfbf8] border border-[#dcd6c8] rounded-[9px] p-[9px_11px] focus:outline-none focus:border-[#a01722] focus:shadow-[0_0_0_3px_rgba(160,23,34,.12)] focus:bg-white transition-[border-color,box-shadow] duration-150"
                  style={{ fontFamily: "inherit" }}
                >
                  {UNIT_PRESETS.map((unit) => <option key={unit} value={unit}>{unit}</option>)}
                  <option value="Custom Unit">Custom Unit</option>
                </select>
                {!unitIsPreset && (
                  <input
                    type="text"
                    value={form.unit}
                    onChange={(e) => updateForm({ unit: e.target.value })}
                    placeholder="Enter custom unit"
                    className="mt-[8px] w-full text-[14px] text-[#1c222b] bg-[#fcfbf8] border border-[#dcd6c8] rounded-[9px] p-[9px_11px] focus:outline-none focus:border-[#a01722] focus:shadow-[0_0_0_3px_rgba(160,23,34,.12)] focus:bg-white transition-[border-color,box-shadow] duration-150"
                    style={{ fontFamily: "inherit" }}
                  />
                )}
                <div className="text-[11px] text-[#6b6f76] mt-[4px]">Use an MBW preset or choose Custom Unit.</div>
              </div>

              {/* Dates */}
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "10px" }}>
                <div>
                  <label className="block text-[12px] font-semibold text-[#3a414b] mb-[5px]">
                    From <span className="text-[#a01722]">*</span>
                  </label>
                  <input
                    type="text"
                    value={form.dateFrom}
                    onChange={(e) => updateForm({ dateFrom: e.target.value })}
                    placeholder="June 2023"
                    autoComplete="off"
                    className="w-full text-[14px] text-[#1c222b] bg-[#fcfbf8] border border-[#dcd6c8] rounded-[9px] p-[9px_11px] focus:outline-none focus:border-[#a01722] focus:shadow-[0_0_0_3px_rgba(160,23,34,.12)] focus:bg-white transition-[border-color,box-shadow] duration-150"
                    style={{ fontFamily: "inherit" }}
                  />
                </div>
                <div>
                  <label className="block text-[12px] font-semibold text-[#3a414b] mb-[5px]">
                    To <span className="text-[#a01722]">*</span>
                  </label>
                  <input
                    type="text"
                    value={form.dateTo}
                    onChange={(e) => updateForm({ dateTo: e.target.value })}
                    placeholder="May 2024"
                    autoComplete="off"
                    className="w-full text-[14px] text-[#1c222b] bg-[#fcfbf8] border border-[#dcd6c8] rounded-[9px] p-[9px_11px] focus:outline-none focus:border-[#a01722] focus:shadow-[0_0_0_3px_rgba(160,23,34,.12)] focus:bg-white transition-[border-color,box-shadow] duration-150"
                    style={{ fontFamily: "inherit" }}
                  />
                </div>
              </div>

              {/* Closing attributes (only for lesser awards, not OVSM) */}
              {cfg.closing === "lesser" && !cfg.isLOA && (
                <div>
                  <label className="block text-[12px] font-semibold text-[#3a414b] mb-[5px]">
                    Closing attributes <span className="font-normal text-[#6b6f76]">(used in the closing sentence)</span>
                  </label>
                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: "10px" }}>
                    <input
                      type="text"
                      value={form.attr1}
                      onChange={(e) => updateForm({ attr1: e.target.value })}
                      placeholder="initiative"
                      autoComplete="off"
                      className="w-full text-[14px] text-[#1c222b] bg-[#fcfbf8] border border-[#dcd6c8] rounded-[9px] p-[9px_11px] focus:outline-none focus:border-[#a01722] focus:shadow-[0_0_0_3px_rgba(160,23,34,.12)] focus:bg-white transition-[border-color,box-shadow] duration-150"
                      style={{ fontFamily: "inherit" }}
                    />
                    <input
                      type="text"
                      value={form.attr2}
                      onChange={(e) => updateForm({ attr2: e.target.value })}
                      placeholder="professionalism"
                      autoComplete="off"
                      className="w-full text-[14px] text-[#1c222b] bg-[#fcfbf8] border border-[#dcd6c8] rounded-[9px] p-[9px_11px] focus:outline-none focus:border-[#a01722] focus:shadow-[0_0_0_3px_rgba(160,23,34,.12)] focus:bg-white transition-[border-color,box-shadow] duration-150"
                      style={{ fontFamily: "inherit" }}
                    />
                    <input
                      type="text"
                      value={form.adj}
                      onChange={(e) => updateForm({ adj: e.target.value })}
                      placeholder="unwavering"
                      autoComplete="off"
                      className="w-full text-[14px] text-[#1c222b] bg-[#fcfbf8] border border-[#dcd6c8] rounded-[9px] p-[9px_11px] focus:outline-none focus:border-[#a01722] focus:shadow-[0_0_0_3px_rgba(160,23,34,.12)] focus:bg-white transition-[border-color,box-shadow] duration-150"
                      style={{ fontFamily: "inherit" }}
                    />
                  </div>
                  <div className="text-[11px] text-[#6b6f76] mt-[4px]">
                    e.g. "By {p.poss} <b>initiative</b>, <b>professionalism</b>, and <b>unwavering</b> dedication to duty…"
                  </div>
                </div>
              )}
            </div>
          </div>

          {/* Accomplishments */}
          <div id="accomplishments" className="scroll-mt-[96px] bg-white border border-[#dcd6c8] rounded-[12px] shadow-[0_1px_2px_rgba(17,22,29,.05),0_8px_24px_rgba(17,22,29,.07)]">
            <div className="flex items-center gap-[10px] p-[14px_16px] border-b border-[#dcd6c8]">
              <span className="w-[22px] h-[22px] rounded-[6px] shrink-0 grid place-items-center text-[12px] font-extrabold bg-[#a01722] text-[#e6d29a]">2</span>
              <h2 className="m-0 text-[13px] uppercase tracking-[.12em] text-[#11161d]">
                {cfg.isLOA ? "Volunteer Service" : "Accomplishments"}
              </h2>
            </div>
            <div className="p-4 space-y-[13px]">
              <div>
                <label className="block text-[12px] font-semibold text-[#3a414b] mb-[5px]">
                  {cfg.isLOA ? "Volunteer activities & community impact" : "Key actions & impact"}
                </label>
                <div className="relative">
                <textarea
                  rows={7}
                  value={form.achievements}
                  onChange={(e) => updateForm({ achievements: e.target.value })}
                  placeholder={cfg.isLOA
                    ? "One activity per line. Describe the organization, your role, hours contributed, and community impact, e.g.\n" +
                      "Volunteered 150+ hours as a youth mentor with the local Boys & Girls Club, guiding at-risk teens.\n" +
                      "Led 20 volunteers in organizing three community food drives serving over 300 families in need."
                    : "One accomplishment per line. Lead with the action and its measurable impact, e.g.\n" +
                      "Led a 12-Marine detail supporting 40+ ceremonial events with zero discrepancies.\n" +
                      "Overhauled the duty roster, reducing scheduling conflicts by 30 percent."}
                  className="w-full text-[14px] text-[#1c222b] bg-[#fcfbf8] border border-[#dcd6c8] rounded-[9px] p-[9px_11px] focus:outline-none focus:border-[#a01722] focus:shadow-[0_0_0_3px_rgba(160,23,34,.12)] focus:bg-white transition-[border-color,box-shadow] duration-150 resize-y leading-[1.5]"
                  style={{ fontFamily: "inherit" }}
                />
                {showCatPopup && (
                  <div className="absolute right-[10px] top-[10px] z-10 rounded-[10px] border border-[#dcd6c8] bg-white px-[10px] py-[7px] text-[12px] font-semibold text-[#3a414b] shadow-[0_8px_22px_rgba(17,22,29,.14)]">
                    <button
                      onClick={() => {
                        setShowCatPopup(false);
                        setCatEasterEggDismissed(true);
                      }}
                      className="ml-[8px] float-right text-[#6b6f76]"
                      aria-label="Dismiss"
                    >
                      x
                    </button>
                    🐱 Meow.
                    <span className="block font-normal text-[#6b6f76]">This may require additional justification.</span>
                  </div>
                )}
                </div>
                <div className="text-[11px] text-[#6b6f76] mt-[4px]">
                  {cfg.isLOA
                    ? "These feed the volunteer service narrative in the Letter of Authorization."
                    : "These feed both the Summary of Action and the citation body. The opening & closing are generated automatically."}
                </div>
              </div>

              {/* Category tags */}
              {classifiedAchievements.length > 0 && (
                <div className="flex flex-wrap gap-[6px]">
                  {(() => {
                    const catCounts = new Map<AchievementCategory, number>();
                    for (const ca of classifiedAchievements) {
                      catCounts.set(ca.category, (catCounts.get(ca.category) || 0) + 1);
                    }
                    const ordered: AchievementCategory[] = [
                      "Leadership", "Operations", "Training", "Administration", "Innovation", "Community Relations", "Uncategorized",
                    ];
                    return ordered.map((cat) => {
                      const count = catCounts.get(cat);
                      if (!count) return null;
                      return (
                        <span
                          key={cat}
                          className="inline-flex items-center gap-[5px] text-[11px] font-semibold px-[8px] py-[4px] rounded-[6px]"
                          style={{
                            background: CATEGORY_COLORS[cat] + "14",
                            color: CATEGORY_COLORS[cat],
                            border: `1px solid ${CATEGORY_COLORS[cat]}33`,
                          }}
                        >
                          {cat}
                          <span
                            className="inline-flex items-center justify-center w-[16px] h-[16px] rounded-full text-[10px]"
                            style={{ background: CATEGORY_COLORS[cat] + "22", color: CATEGORY_COLORS[cat] }}
                          >
                            {count}
                          </span>
                        </span>
                      );
                    });
                  })()}
                </div>
              )}

              {form.achievements.trim() && currentWeakInput.length > 0 && (
                <div className="rounded-[10px] border border-[#e6c98b] bg-[#fdf8f0] p-[10px_12px] text-[12.5px] leading-[1.45] text-[#5b4a13]">
                  <b className="block uppercase tracking-[.08em] text-[11px] mb-[3px]">Weak input guidance</b>
                  <ul className="m-0 pl-[18px]">
                    {currentWeakInput.slice(0, 3).map((issue) => (
                      <li key={issue.title} className="mb-[3px]">
                        <b>{issue.title}:</b> {issue.detail}
                      </li>
                    ))}
                  </ul>
                </div>
              )}

              {currentAwardMatch.severity !== "none" && (
                <div
                  className="text-[12.5px] leading-[1.45] rounded-[10px] p-[10px_12px] border"
                  style={{
                    background: currentAwardMatch.severity === "severe" ? "#fef7f6" : "#fdf8f0",
                    borderColor: currentAwardMatch.severity === "severe" ? "#f0b8b3" : "#e6c98b",
                    color: currentAwardMatch.severity === "severe" ? "#7c1d13" : "#5b4a13",
                  }}
                >
                  <b className="block uppercase tracking-[.08em] text-[11px] mb-[3px]">{currentAwardMatch.title}</b>
                  {currentAwardMatch.detail}
                  <span className="block mt-[4px]">{currentAwardMatch.recommendations[0]}</span>
                </div>
              )}

              <div className="flex flex-wrap items-center gap-[9px]">
                <button
                  type="button"
                  onClick={() => setAiEnhancement((v) => !v)}
                  className="cursor-pointer text-[13.5px] font-semibold px-[13px] py-[10px] rounded-[9px] inline-flex items-center gap-[8px] bg-white border border-[#dcd6c8] text-[#3a414b] hover:border-[#a01722] transition-colors active:translate-y-px"
                >
                  <span
                    className="w-[34px] h-[18px] rounded-full p-[2px] inline-flex"
                    style={{ background: aiEnhancement ? "#2f7d44" : "#aeb6c2", justifyContent: aiEnhancement ? "flex-end" : "flex-start" }}
                  >
                    <span className="w-[14px] h-[14px] rounded-full bg-white" />
                  </span>
                  AI Enhancement {aiEnhancement ? "ON" : "OFF"}
                </button>
                <button
                  onClick={handleGenerate}
                  disabled={hasHardValidationError || aiLoading}
                  className="cursor-pointer text-[13.5px] font-semibold px-[15px] py-[10px] rounded-[9px] inline-flex items-center gap-2 text-white transition-transform duration-[.08s] active:translate-y-px"
                  style={{
                    background: hasHardValidationError || aiLoading ? "#aeb6c2" : "linear-gradient(160deg, #a01722, #7c0f19)",
                    boxShadow: "0 6px 16px rgba(160,23,34,.28)",
                    cursor: hasHardValidationError || aiLoading ? "not-allowed" : "pointer",
                  }}
                >
                  {aiLoading ? "Generating..." : "Generate Award Package"}
                </button>
                <button
                  onClick={handleClear}
                  className="cursor-pointer text-[13.5px] font-semibold px-[15px] py-[10px] rounded-[9px] inline-flex items-center gap-2 bg-[#f0ece2] text-[#3a414b] border-none transition-colors hover:bg-[#e7e1d4] active:translate-y-px"
                >
                  Clear form
                </button>
                <button
                  onClick={handleNewAward}
                  className="cursor-pointer text-[13.5px] font-semibold px-[15px] py-[10px] rounded-[9px] inline-flex items-center gap-2 bg-white border border-[#b3261e]/30 text-[#b3261e] hover:bg-[#fef7f6] hover:border-[#b3261e] transition-colors active:translate-y-px"
                >
                  New Award
                </button>
              </div>
            </div>
          </div>

          {/* Saved Drafts */}
          <div id="saved-drafts" className="scroll-mt-[96px] bg-white border border-[#dcd6c8] rounded-[12px] shadow-[0_1px_2px_rgba(17,22,29,.05),0_8px_24px_rgba(17,22,29,.07)]">
            <div className="flex items-center gap-[10px] p-[14px_16px] border-b border-[#dcd6c8]">
              <span className="w-[22px] h-[22px] rounded-[6px] shrink-0 grid place-items-center text-[12px] font-extrabold bg-[#a01722] text-[#e6d29a]">S</span>
              <h2 className="m-0 text-[13px] uppercase tracking-[.12em] text-[#11161d]">Saved Drafts</h2>
              <button
                onClick={handleSaveDraft}
                className="ml-auto cursor-pointer text-[12px] font-semibold px-[11px] py-[7px] rounded-[8px] bg-[#f0ece2] text-[#3a414b] hover:bg-[#e7e1d4] transition-colors"
              >
                {activeDraftId ? "Update Draft" : "Save Current"}
              </button>
              <button
                onClick={handleClearLocalDrafts}
                className="cursor-pointer text-[12px] font-semibold px-[11px] py-[7px] rounded-[8px] bg-white border border-[#f0b8b3] text-[#b3261e] hover:bg-[#fef7f6] transition-colors"
              >
                Clear Local Drafts
              </button>
            </div>
            <div className="p-3 grid gap-[8px]">
              {savedDrafts.length ? savedDrafts.map((draft) => (
                <div key={draft.id} className="rounded-[10px] border border-[#efe9dc] bg-[#faf8f3] p-[10px]">
                  <div className="flex items-start gap-[8px]">
                    <div className="min-w-0 flex-1">
                      <b className="block text-[13px] text-[#11161d] truncate">{draft.name}</b>
                      <div className="text-[11.5px] text-[#6b6f76] leading-[1.45]">
                        {[draft.form.lastName || "No last name", draft.form.rank, draft.form.award, draft.form.billet || "No billet"].join(" • ")}
                        <br />
                        {[draft.form.dateFrom || "No start", draft.form.dateTo || "No end"].join(" to ")} • Modified {formatDraftDate(draft.updatedAt)}
                      </div>
                    </div>
                    {activeDraftId === draft.id && <span className="text-[10.5px] font-bold text-[#2f7d44] bg-[#e8f4eb] px-[7px] py-[3px] rounded-full">Open</span>}
                  </div>
                  <div className="flex flex-wrap gap-[6px] mt-[9px]">
                    <button onClick={() => handleOpenSavedDraft(draft.id)} className="text-[11.5px] font-semibold px-[9px] py-[6px] rounded-[7px] bg-white border border-[#dcd6c8] text-[#3a414b]">Open</button>
                    <button onClick={() => handleRenameDraft(draft.id)} className="text-[11.5px] font-semibold px-[9px] py-[6px] rounded-[7px] bg-white border border-[#dcd6c8] text-[#3a414b]">Rename</button>
                    <button onClick={() => handleDuplicateDraft(draft.id)} className="text-[11.5px] font-semibold px-[9px] py-[6px] rounded-[7px] bg-white border border-[#dcd6c8] text-[#3a414b]">Duplicate</button>
                    <button onClick={() => handleDeleteDraft(draft.id)} className="text-[11.5px] font-semibold px-[9px] py-[6px] rounded-[7px] bg-white border border-[#f0b8b3] text-[#b3261e]">Delete</button>
                  </div>
                </div>
              )) : (
                <div className="text-[12.5px] text-[#6b6f76] bg-[#faf8f3] border border-[#efe9dc] rounded-[9px] p-[10px]">
                  No saved drafts yet. Save current work to keep multiple awards available in this browser.
                </div>
              )}
            </div>
          </div>
        </section>

        {/* ---- SECTION 2: Output ---- */}
        <section id="outputs" className="scroll-mt-[96px]">
          <div className="bg-white border border-[#dcd6c8] rounded-[12px] shadow-[0_1px_2px_rgba(17,22,29,.05),0_8px_24px_rgba(17,22,29,.07)]">
            <div className="flex items-center gap-[10px] p-[14px_16px] border-b border-[#dcd6c8]">
              <span className="w-[22px] h-[22px] rounded-[6px] shrink-0 grid place-items-center text-[12px] font-extrabold bg-[#a01722] text-[#e6d29a]">3</span>
              <h2 className="m-0 text-[13px] uppercase tracking-[.12em] text-[#11161d]">
                {cfg.isLOA ? "Letter of Authorization" : "Drafts"}
              </h2>
              {!cfg.isLOA && (
                <span className="ml-auto text-[12px] text-[#6b6f76]">
                  {cfg.casing === "upper" ? "ALL CAPS" : "Sentence case"}
                </span>
              )}
            </div>
            <div className="p-4">
              {currentAwardMatch.severity !== "none" && (
                <div
                  className="rounded-[12px] border p-[13px_15px] mb-4 text-[13px] leading-[1.45]"
                  style={{
                    background: awardConcernKind === "upgrade" ? "#eef6ff" : "#fff1f0",
                    borderColor: awardConcernKind === "upgrade" ? "#6aa9e8" : "#b3261e",
                    color: awardConcernKind === "upgrade" ? "#173a5e" : "#7c1d13",
                    boxShadow: awardConcernKind === "upgrade" ? "0 8px 24px rgba(26,77,143,.12)" : "0 8px 24px rgba(179,38,30,.12)",
                  }}
                >
                  <b className="block text-[13px] uppercase tracking-[.1em] mb-[5px]">
                    {awardConcernKind === "upgrade" ? "⬆ POTENTIAL UPGRADE" : "⚠ AWARD LEVEL CONCERN"}
                  </b>
                  <span>
                    {awardConcernKind === "upgrade"
                      ? "Accomplishments may justify consideration for a higher award level."
                      : `Accomplishments currently appear below the typical threshold for ${AWARDS[form.award].label}.`}
                  </span>
                  <span className="block mt-[5px] font-semibold">
                    {awardConcernKind === "upgrade"
                      ? `Recommended review: ${awardShortLabel(currentAwardMatch.recommendedAward)}.`
                      : `Recommended award level: ${awardShortLabel(currentAwardMatch.recommendedAward)}.`}
                  </span>
                </div>
              )}

              {/* AI banner */}
              {aiBanner.show && (
                <div
                  className="flex items-center gap-2 text-[12.5px] p-[9px_12px] rounded-[9px] mb-3"
                  style={{
                    background: aiBanner.over ? "#fde6e3" : "#fbf2d9",
                    borderColor: aiBanner.over ? "#b3261e" : "#c5a44e",
                    borderWidth: "1px",
                    borderStyle: "solid",
                    color: aiBanner.over ? "#7c1d13" : "#5b4a13",
                  }}
                >
                  {aiBanner.message}
                </div>
              )}

              {/* ---- OVSM: Letter of Authorization output ---- */}
              {cfg.isLOA ? (
                <>
                  <div className="flex items-center gap-[10px] mb-[10px]">
                    <h3 className="m-0 text-[14px] text-[#11161d]">Letter of Authorization</h3>
                    <div className="ml-auto flex gap-[7px]">
                      <button
                        onClick={() => handleCopy("soa")}
                        className="cursor-pointer text-[12px] font-semibold px-[11px] py-[6px] rounded-[9px] bg-white border border-[#dcd6c8] text-[#3a414b] hover:border-[#a01722] hover:text-[#a01722] transition-[border-color,color] duration-150"
                      >
                        Copy LOA
                      </button>
                    </div>
                  </div>
                  <div
                    className="w-full min-h-[300px] text-[15px] leading-[1.62] text-[#1a1f27] bg-[#fffdf8] border border-[#dcd6c8] rounded-[10px] p-[16px_18px] whitespace-pre-wrap break-words"
                    style={{ fontFamily: '"Iowan Old Style", "Palatino Linotype", Georgia, serif' }}
                  >
                    {soa || <span className="text-[#6b6f76] italic" style={{ fontFamily: "inherit" }}>Your Letter of Authorization will appear here.</span>}
                  </div>
                </>
              ) : cfg.citationOnly ? (
                <>
                  <div id="citation-output" className="scroll-mt-[96px] flex items-center gap-[10px] mb-[10px]">
                    <h3 className="m-0 text-[14px] text-[#11161d]">Proposed Citation</h3>
                    <div className="ml-auto flex gap-[7px]">
                      <button
                        onClick={handleToggleSpell}
                        className={`cursor-pointer text-[12px] font-semibold px-[11px] py-[6px] rounded-[9px] transition-colors ${
                          spellMode
                            ? "text-[#3a2e08] border-none"
                            : "bg-white border border-[#dcd6c8] text-[#3a414b] hover:border-[#a01722] hover:text-[#a01722]"
                        }`}
                        style={spellMode ? { background: "linear-gradient(160deg, #d8bb63, #c5a44e)" } : undefined}
                      >
                        {spellMode ? "Done editing" : "Spell-check mode"}
                      </button>
                      <button
                        onClick={() => handleCopy("citation")}
                        className="cursor-pointer text-[12px] font-semibold px-[11px] py-[6px] rounded-[9px] bg-white border border-[#dcd6c8] text-[#3a414b] hover:border-[#a01722] hover:text-[#a01722] transition-[border-color,color] duration-150"
                      >
                        Copy Citation
                      </button>
                    </div>
                  </div>
                  <div
                    contentEditable={spellMode}
                    suppressContentEditableWarning
                    spellCheck={spellMode}
                    onInput={(e) => {
                      if (spellMode) {
                        setCitation(e.currentTarget.textContent || "");
                      }
                    }}
                    onBlur={() => {
                      if (spellMode) {
                        setChecks(runChecks(citation, soa, form));
                      }
                    }}
                    className={`w-full min-h-[300px] text-[15px] leading-[1.62] text-[#1a1f27] border rounded-[10px] p-[16px_18px] whitespace-pre-wrap break-words outline-none ${
                      spellMode ? "bg-[#fffef9] shadow-[inset_0_0_0_2px_rgba(197,164,78,.4)]" : "bg-[#fffdf8] border-[#dcd6c8]"
                    }`}
                    style={{ fontFamily: '"Iowan Old Style", "Palatino Linotype", Georgia, serif' }}
                  >
                    {citation || (!spellMode && <span className="text-[#6b6f76] italic" style={{ fontFamily: "inherit" }}>Your proposed citation will appear here.</span>)}
                  </div>
                </>
              ) : (
                <>
              {/* SOA Output */}
              <div className="flex items-center gap-[10px] mb-[10px]">
                <h3 className="m-0 text-[14px] text-[#11161d]">Summary of Action</h3>
                <div className="ml-auto flex gap-[7px]">
                  <button
                    onClick={() => handleCopy("soa")}
                    className="cursor-pointer text-[12px] font-semibold px-[11px] py-[6px] rounded-[9px] bg-white border border-[#dcd6c8] text-[#3a414b] hover:border-[#a01722] hover:text-[#a01722] transition-[border-color,color] duration-150"
                  >
                    Copy SOA
                  </button>
                </div>
              </div>
              <div
                className="w-full min-h-[150px] text-[15px] leading-[1.62] text-[#1a1f27] bg-[#fffdf8] border border-[#dcd6c8] rounded-[10px] p-[16px_18px] whitespace-pre-wrap break-words"
                style={{ fontFamily: '"Iowan Old Style", "Palatino Linotype", Georgia, serif' }}
              >
                {soa || <span className="text-[#6b6f76] italic" style={{ fontFamily: "inherit" }}>Your Summary of Action will appear here.</span>}
              </div>

              {/* Citation Output */}
              <div id="citation-output" className="scroll-mt-[96px] flex items-center gap-[10px] mb-[10px] mt-5">
                <h3 className="m-0 text-[14px] text-[#11161d]">Proposed Citation</h3>
                <div className="ml-auto flex gap-[7px]">
                  <button
                    onClick={handleToggleSpell}
                    className={`cursor-pointer text-[12px] font-semibold px-[11px] py-[6px] rounded-[9px] transition-colors ${
                      spellMode
                        ? "text-[#3a2e08] border-none"
                        : "bg-white border border-[#dcd6c8] text-[#3a414b] hover:border-[#a01722] hover:text-[#a01722]"
                    }`}
                    style={spellMode ? { background: "linear-gradient(160deg, #d8bb63, #c5a44e)" } : undefined}
                  >
                    {spellMode ? "Done editing" : "Spell-check mode"}
                  </button>
                  <button
                    onClick={() => handleCopy("citation")}
                    className="cursor-pointer text-[12px] font-semibold px-[11px] py-[6px] rounded-[9px] bg-white border border-[#dcd6c8] text-[#3a414b] hover:border-[#a01722] hover:text-[#a01722] transition-[border-color,color] duration-150"
                  >
                    Copy Citation
                  </button>
                </div>
              </div>
              <div
                contentEditable={spellMode}
                suppressContentEditableWarning
                spellCheck={spellMode}
                onInput={(e) => {
                  if (spellMode) {
                    setCitation(e.currentTarget.textContent || "");
                  }
                }}
                onBlur={() => {
                  if (spellMode) {
                    setChecks(runChecks(citation, soa, form));
                  }
                }}
                className={`w-full min-h-[150px] text-[15px] leading-[1.62] text-[#1a1f27] border rounded-[10px] p-[16px_18px] whitespace-pre-wrap break-words outline-none ${
                  spellMode ? "bg-[#fffef9] shadow-[inset_0_0_0_2px_rgba(197,164,78,.4)]" : "bg-[#fffdf8] border-[#dcd6c8]"
                }`}
                style={{ fontFamily: '"Iowan Old Style", "Palatino Linotype", Georgia, serif' }}
              >
                {citation || (!spellMode && <span className="text-[#6b6f76] italic" style={{ fontFamily: "inherit" }}>Your proposed citation will appear here.</span>)}
              </div>

              {/* Character counter */}
              <div className="flex items-center gap-[10px] mt-[9px] text-[12px]">
                <span
                  className="font-mono text-[#3a414b]"
                  style={{ fontVariantNumeric: "tabular-nums" }}
                >
                  <b className={charCount > maxChars && maxChars > 0 ? "text-[#b3261e]" : (cfg.target && charCount >= cfg.target[0] ? "text-[#2f7d44]" : "text-[#11161d]")}>
                    {charCount}
                  </b>
                  {maxChars ? ` / ${maxChars}` : " chars"}
                </span>
                <div className="flex-1 h-[7px] rounded-full bg-[#e9e3d6] overflow-hidden">
                  <div
                    className="h-full rounded-full transition-[width,background-color] duration-[.35s]"
                    style={{ width: `${charPct}%`, background: charBarColor }}
                  />
                </div>
                <span className="text-[#6b6f76]">
                  {!maxChars
                    ? "No hard limit"
                    : charCount > maxChars
                      ? `Over limit — trim ${charCount - maxChars} chars`
                      : cfg.target && charCount >= cfg.target[0]
                        ? "In target zone"
                        : `Target ${cfg.target?.[0]}–${cfg.target?.[1]}`}
                </span>
              </div>
              </>
              )}

              {/* ---- Action buttons (shown for both OVSM and standard awards) ---- */}
              <div className="flex flex-wrap gap-[9px] mt-4">
                <div className="relative" ref={exportRef}>
                  <button
                    onClick={() => setExportOpen(!exportOpen)}
                    className="cursor-pointer text-[13.5px] font-semibold px-[15px] py-[10px] rounded-[9px] inline-flex items-center gap-2 bg-white border border-[#dcd6c8] text-[#3a414b] hover:border-[#a01722] hover:text-[#a01722] transition-[border-color,color] duration-150 active:translate-y-px"
                  >
                    Export Award Package
                    <svg width="10" height="6" viewBox="0 0 10 6" fill="none" style={{ transform: exportOpen ? "rotate(180deg)" : "none", transition: "transform .15s" }}>
                      <path d="M1 1l4 4 4-4" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" />
                    </svg>
                  </button>
                  {exportOpen && (
                    <div className="absolute bottom-full mb-[6px] left-0 bg-white border border-[#dcd6c8] rounded-[10px] shadow-[0_4px_24px_rgba(17,22,29,.12)] overflow-hidden z-50 min-w-[180px]" style={{ animation: "fadeIn .12s ease" }}>
                      <button
                        onClick={handleExportWord}
                        className="w-full text-left text-[13px] font-medium px-[14px] py-[10px] text-[#3a414b] hover:bg-[#f6f3ea] hover:text-[#a01722] transition-colors border-b border-[#efe9dc]"
                      >
                        Export Word (.docx)
                      </button>
                      <button
                        onClick={handleExportPDF}
                        className="w-full text-left text-[13px] font-medium px-[14px] py-[10px] text-[#3a414b] hover:bg-[#f6f3ea] hover:text-[#a01722] transition-colors border-b border-[#efe9dc]"
                      >
                        Export PDF
                      </button>
                      <button
                        onClick={handleExportBoth}
                        className="w-full text-left text-[13px] font-medium px-[14px] py-[10px] text-[#3a414b] hover:bg-[#f6f3ea] hover:text-[#a01722] transition-colors"
                      >
                        Export Award Package
                      </button>
                    </div>
                  )}
                </div>
              </div>
            </div>
          </div>
        </section>

        {/* ---- SECTION 3: Validation ---- */}
        <aside id="validation" className="scroll-mt-[96px]">
          <div className="bg-white border border-[#dcd6c8] rounded-[12px] shadow-[0_1px_2px_rgba(17,22,29,.05),0_8px_24px_rgba(17,22,29,.07)]">
            <div className="flex items-center gap-[10px] p-[14px_16px] border-b border-[#dcd6c8]">
              <span className="w-[22px] h-[22px] rounded-[6px] shrink-0 grid place-items-center text-[12px] font-extrabold bg-[#a01722] text-[#e6d29a]">✓</span>
              <h2 className="m-0 text-[13px] uppercase tracking-[.12em] text-[#11161d]">Validation</h2>
            </div>

            {currentAwardMatch.severity !== "none" && (
              <div
                className="m-[12px_12px_0] rounded-[10px] border p-[11px_12px] text-[12.5px] leading-[1.45]"
                style={{
                  background: currentAwardMatch.severity === "severe" ? "#fef7f6" : "#fdf8f0",
                  borderColor: currentAwardMatch.severity === "severe" ? "#b3261e" : "#b5751a",
                  color: currentAwardMatch.severity === "severe" ? "#7c1d13" : "#5b4a13",
                }}
              >
                <b className="block text-[11px] uppercase tracking-[.08em] mb-[4px]">
                  {currentAwardMatch.title}
                </b>
                <span>{currentAwardMatch.detail}</span>
                <ul className="m-[7px_0_0] pl-[18px]">
                  {currentAwardMatch.recommendations.map((rec) => <li key={rec}>{rec}</li>)}
                </ul>
              </div>
            )}

            {dateErr && (
              <div className="m-[12px_12px_0] rounded-[10px] border border-[#b3261e] bg-[#fef7f6] p-[11px_12px] text-[12.5px] font-semibold text-[#7c1d13]">
                {dateErr}
              </div>
            )}

            <div className="m-[12px_12px_0] rounded-[10px] border border-[#dcd6c8] bg-[#faf8f3] p-[11px_12px]">
              <button
                onClick={handleFixWithAI}
                disabled={!aiAvailable || aiLoading || (!soa && !citation)}
                className="w-full cursor-pointer text-[13.5px] font-semibold px-[14px] py-[10px] rounded-[9px] inline-flex items-center justify-center gap-2 text-[#3a2e08] transition-transform duration-[.08s] active:translate-y-px disabled:opacity-50 disabled:cursor-not-allowed"
                style={{
                  background: "linear-gradient(160deg, #d8bb63, #c5a44e)",
                  boxShadow: "0 6px 16px rgba(197,164,78,.24)",
                }}
              >
                {aiLoading ? "Fixing..." : "Fix With AI"}
              </button>
              {aiFixSummary && (
                <div className="mt-[9px] text-[12.5px] text-[#3a414b]">
                  <button
                    onClick={() => setAiFixSummary((prev) => prev ? { ...prev, open: !prev.open } : prev)}
                    className="font-semibold text-[#11161d] underline decoration-[#c5a44e] underline-offset-2"
                  >
                    AI applied {aiFixSummary.count} improvement{aiFixSummary.count === 1 ? "" : "s"}
                  </button>
                  {aiFixSummary.open && (
                    <ul className="m-[7px_0_0] pl-[18px] text-[#6b6f76]">
                      {aiFixSummary.details.map((detail, i) => <li key={i} className="mb-[4px]">{detail}</li>)}
                    </ul>
                  )}
                </div>
              )}
            </div>

            {/* Scores — Compliance + Quality */}
            <div className="p-[14px_16px] border-b border-[#dcd6c8] grid gap-[14px]">
              {/* Compliance Score */}
              <div className="flex items-center gap-3">
                <ScoreRing pct={checks.length ? scorePct : 0} errors={errorChecks} />
                <div className="text-[12.5px] text-[#6b6f76]">
                  <b className="block text-[#11161d] text-[14px]">
                    {!checks.length
                      ? "Not yet generated"
                      : errorChecks
                        ? `${errorChecks} issue${errorChecks > 1 ? "s" : ""} to fix`
                        : scorePct === 100
                          ? "Competition-ready"
                          : "Review warnings"}
                  </b>
                  Compliance Score
                </div>
              </div>
              {form.achievements.trim() && (
                <div className="flex items-center gap-3">
                  <ScoreRing
                    pct={currentAwardMatch.score}
                    errors={currentAwardMatch.severity === "severe" ? 1 : 0}
                  />
                  <div className="text-[12.5px] text-[#6b6f76]">
                    <b className="block text-[#11161d] text-[14px]">
                      Award Justification: {currentAwardMatch.score}%
                    </b>
                    <span>{currentAwardMatch.recommendations[0] || `Recommended award level: ${awardShortLabel(currentAwardMatch.recommendedAward)}.`}</span>
                  </div>
                </div>
              )}
              {/* Quality Score */}
              {qualityScores && (
                <>
                  <div className="flex items-center gap-3">
                    <ScoreRing
                      pct={qualityScores.overall}
                      errors={qualityScores.overall < 50 ? 1 : 0}
                    />
                    <div className="text-[12.5px] text-[#6b6f76]">
                      <b className="block text-[#11161d] text-[14px]">
                        {qualityScores.overall >= 80
                          ? "Excellent"
                          : qualityScores.overall >= 60
                            ? "Strong"
                            : qualityScores.overall >= 40
                              ? "Adequate"
                              : "Needs work"}
                      </b>
                      Writing Quality Score
                    </div>
                  </div>
                  <div className="grid gap-[6px]">
                    {[
                      { label: "Quantifiable impact", score: qualityScores.quantifiableImpact },
                      { label: "Strong action verbs", score: qualityScores.strongVerbs },
                      { label: "Leadership language", score: qualityScores.leadershipLanguage },
                      { label: "Result-oriented", score: qualityScores.resultOriented },
                    ].map((item) => (
                      <div key={item.label} className="flex items-center gap-[8px]">
                        <span className="text-[11px] text-[#6b6f76] w-[120px] shrink-0 truncate">{item.label}</span>
                        <div className="flex-1 h-[5px] rounded-full bg-[#e9e3d6] overflow-hidden">
                          <div
                            className="h-full rounded-full transition-[width] duration-[.35s]"
                            style={{
                              width: `${item.score}%`,
                              background: item.score >= 70 ? "#2f7d44" : item.score >= 40 ? "#b5751a" : "#b3261e",
                            }}
                          />
                        </div>
                        <span className="text-[11px] font-mono text-[#3a414b] w-[28px] text-right" style={{ fontVariantNumeric: "tabular-nums" }}>
                          {item.score}
                        </span>
                      </div>
                    ))}
                  </div>
                </>
              )}
            </div>

            {/* Reviewer checklist */}
            <div className="p-[12px_16px] border-b border-[#dcd6c8] bg-[#fbfaf6]">
              <h4 className="m-0 mb-[8px] text-[11px] uppercase tracking-[.1em] text-[#6b6f76]">V1 Reviewer Checklist</h4>
              <div className="grid gap-[6px]">
                {reviewerChecklist.map((item) => (
                  <div key={item.label} className="flex items-start gap-[8px] text-[12.3px] leading-[1.35]">
                    <span
                      className="mt-[1px] w-[16px] h-[16px] rounded-full grid place-items-center text-[10px] font-extrabold text-white shrink-0"
                      style={{ background: item.ok ? "#2f7d44" : "#b5751a" }}
                    >
                      {item.ok ? "\u2713" : "!"}
                    </span>
                    <span className="min-w-0">
                      <b className="text-[#11161d]">{item.label}</b>
                      <span className="block text-[#6b6f76]">{item.detail}</span>
                    </span>
                  </div>
                ))}
              </div>
            </div>

            {/* Checks list */}
            {checks.length > 0 ? (
              <ul className="list-none m-0 p-2 grid gap-[6px]">
                {checks.map((check, i) => (
                  <CheckRow key={i} check={check} onFix={handleAutoFix} />
                ))}
              </ul>
            ) : (
              <ul className="list-none m-0 p-2">
                <li className="flex gap-[9px] items-start text-[12.5px] leading-[1.4] p-[9px_11px] rounded-[9px] bg-[#faf8f3] border border-[#efe9dc]">
                  <div className="shrink-0 w-4 h-4 rounded-full mt-px grid place-items-center text-white text-[10px] font-extrabold bg-[#6b6f76]">i</div>
                  <div>
                    <b className="block text-[#11161d] mb-px text-[12.5px]">Generate a draft</b>
                    <span className="text-[#6b6f76]">Run the engine to validate formatting and compliance.</span>
                  </div>
                </li>
              </ul>
            )}

            {/* AI notes */}
            {aiNotes.length > 0 && (
              <div className="p-[12px_16px] border-t border-[#dcd6c8]">
                <h4 className="m-0 mb-2 text-[11px] uppercase tracking-[.1em] text-[#6b6f76]">AI reviewer notes</h4>
                <ul className="m-0 pl-[18px] text-[12.5px] text-[#3a414b]">
                  {aiNotes.map((note, i) => (
                    <li key={i} className="mb-[5px]">{note}</li>
                  ))}
                </ul>
              </div>
            )}
          </div>
        </aside>
      </main>

      {/* Footer */}
      <footer className="max-w-[1480px] mx-auto px-[16px] lg:px-[22px] pt-2 pb-[30px]">
        <div className="rounded-[10px] border border-[#dcd6c8] bg-white/70 p-[11px_13px] text-[11.5px] text-[#6b6f76] leading-[1.45]">
          <b className="text-[#11161d]">CitationBuilder {APP_VERSION}</b> is a drafting aid only. Verify every package against current SECNAV M-1650.1, command SOP, and S-1/adjutant guidance before submission. AI may produce incorrect wording; do not enter classified, CUI, medical, legal, disciplinary, or sensitive operational details. Drafts autosave locally in this browser.{" "}
          <kbd className="font-mono bg-[#ece6d8] px-[5px] py-[1px] rounded text-[11px]">Ctrl/⌘ + Enter</kbd> generates.{" "}
          <a href={SUPPORT_EMAIL} className="font-semibold text-[#a01722] underline decoration-[#d8bb63] underline-offset-2">Report an issue</a>.
        </div>
      </footer>
    </div>
    </>
  );
}

/* eslint-disable react-refresh/only-export-components -- V1 smoke tests exercise the citation engine in this single-file app. */
export {
  APP_VERSION,
  AWARDS,
  DEFAULT_FORM,
  assembleCitation,
  buildClosing,
  buildLOA,
  buildOpening,
  buildSOA,
  enforceCitationLimit,
  normalizeSOA,
  redactSensitiveForAI,
  runChecks,
};
export type { AwardKey, FormState };
