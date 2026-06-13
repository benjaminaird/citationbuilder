import { describe, expect, it } from "vitest";

import {
  APP_VERSION,
  AWARDS,
  DEFAULT_FORM,
  assembleCitation,
  buildLOA,
  buildSOA,
  normalizeSOA,
  redactSensitiveForAI,
  runChecks,
  type AwardKey,
  type FormState,
} from "@/pages/Index";

const baseAchievements = [
  "Led 24 Marines through 18 ceremonies supporting national and diplomatic events with zero discrepancies.",
  "Managed 75 travel profiles and 500 authorizations and vouchers, resolving $100,000 in travel claims.",
  "Advised the commander during inspections, improving readiness from 65 percent to 85 percent over six months.",
].join("\n");

function makeForm(award: AwardKey, patch: Partial<FormState> = {}): FormState {
  return {
    ...DEFAULT_FORM,
    award,
    rank: "Sergeant",
    firstName: "Jane",
    lastName: "Marine",
    edipi: "1234567890",
    billet: "Platoon Sergeant",
    additionalBillets: "Acting Company Administrative Chief",
    unit: "Marine Barracks, Washington, D.C.",
    dateFrom: "January 2025",
    dateTo: "December 2025",
    achievements: baseAchievements,
    ...patch,
  };
}

describe("CitationBuilder V1 award engine", () => {
  it("labels the release as V1", () => {
    expect(APP_VERSION).toBe("v1.0");
  });

  it.each([
    ["MMAST", "DURING THE PERIOD OF", "REFLECTED CREDIT", true],
    ["CERTCOM", "EXCEPTIONAL PERFORMANCE", "REFLECTED GREAT CREDIT", true],
    ["NAM", "PROFESSIONAL ACHIEVEMENT", "REFLECTED CREDIT", false],
    ["NMC", "MERITORIOUS SERVICE", "REFLECTED CREDIT", false],
  ] satisfies Array<[AwardKey, string, string, boolean]>)("%s follows uppercase citation rules", (award, opening, closing, citationOnly) => {
    const form = makeForm(award);
    const citation = assembleCitation(form);
    expect(citation).toContain(opening);
    expect(citation).toContain(closing);
    expect(citation.replace(/[^A-Za-z]/g, "")).toBe(citation.replace(/[^A-Za-z]/g, "").toUpperCase());
    expect(AWARDS[award].citationOnly === true).toBe(citationOnly);
    if (AWARDS[award].maxChars) expect(citation.length).toBeLessThanOrEqual(AWARDS[award].maxChars);
    expect(runChecks(citation, "", form).filter((check) => check.status === "err")).toHaveLength(0);
  });

  it.each([
    ["MSM", "For outstanding meritorious service", "reflected great credit"],
    ["LOM", "For exceptionally meritorious conduct", "reflected great credit"],
  ] satisfies Array<[AwardKey, string, string]>)("%s follows sentence-case senior award rules", (award, opening, closing) => {
    const form = makeForm(award, { dateFrom: "January 2023", dateTo: "December 2025" });
    const citation = assembleCitation(form);
    expect(citation).toContain(opening);
    expect(citation).toContain(closing);
    expect(citation.length).toBeLessThanOrEqual(AWARDS[award].maxChars);
    expect(runChecks(citation, buildSOA(form), form).filter((check) => check.status === "err")).toHaveLength(0);
  });

  it("builds OVSM as a Letter of Authorization instead of a citation", () => {
    const form = makeForm("OVSM", {
      achievements: "Volunteered 160 hours with a local nonprofit, coordinated monthly events, and served 450 community beneficiaries.",
    });
    const loa = buildLOA(form);
    expect(loa).toContain("LETTER OF AUTHORIZATION");
    expect(loa).toContain("Outstanding Volunteer Service Medal");
    expect(AWARDS.OVSM.isLOA).toBe(true);
    expect(runChecks("", loa, form).filter((check) => check.status === "err")).toHaveLength(0);
  });

  it("keeps SOA paragraph prose in normal case while citations follow award casing", () => {
    const form = makeForm("NAM");
    const soa = normalizeSOA(
      "SERGEANT MARINE LED 24 MARINES THROUGH 18 CEREMONIES WITH ZERO DISCREPANCIES.\n\nBACKGROUND\n\nSHE IMPROVED READINESS FROM 65 PERCENT TO 85 PERCENT.",
      form,
    );
    const citation = assembleCitation(form);

    expect(soa).toContain("Sergeant Marine led 24 Marines");
    expect(soa).toContain("\n\nShe improved readiness");
    expect(soa).not.toContain("BACKGROUND");
    expect(soa.replace(/[^A-Za-z]/g, "")).not.toBe(soa.replace(/[^A-Za-z]/g, "").toUpperCase());
    expect(citation.replace(/[^A-Za-z]/g, "")).toBe(citation.replace(/[^A-Za-z]/g, "").toUpperCase());
  });

  it("redacts EDIPI and SSN-looking values from AI payloads", () => {
    const redacted = redactSensitiveForAI({
      edipi: "1234567890",
      achievements: "EDIPI 1234567890 and SSN 123-45-6789 should not go to AI.",
      nested: ["Marine 0987654321"],
    });
    expect(JSON.stringify(redacted)).not.toContain("1234567890");
    expect(JSON.stringify(redacted)).not.toContain("123-45-6789");
    expect(JSON.stringify(redacted)).not.toContain("0987654321");
    expect(JSON.stringify(redacted)).toContain("[REDACTED");
  });
});
