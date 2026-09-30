/**
 * Format rules for the scholarly identifiers SPEC §4.2 names — DOI, PMID, PMCID, arXiv,
 * ClinicalTrials.gov and ORCID (task E02-02.1).
 *
 * These are **format** checks only. §14's "Automatic DOI validation" and "PMID and PMCID
 * linking" are network lookups against the registrar and belong to Phase 7-adjacent work; a
 * well-formed DOI is not a resolvable one, and this module never claims otherwise.
 *
 * Every pattern is exported as a plain source string rather than a `RegExp`, because field
 * definitions are data (ADR-0005) and must stay JSON-serialisable for the day the registry is
 * seeded from a file. `FieldRegistry` anchors and compiles them.
 */
/**
 * DOI: a registrant prefix `10.` plus 4–9 digits, then any non-blank suffix. Suffixes are
 * case-insensitive and may contain almost anything, so the rule stays deliberately loose — a
 * stricter one rejects real DOIs, and rejecting a real identifier is the worse failure.
 */
export const DOI_PATTERN = "10\\.\\d{4,9}(?:\\.\\d+)*/\\S+";
/** PMID: a positive integer with no leading zero, no prefix and no version suffix. */
export const PMID_PATTERN = "[1-9]\\d{0,8}";
/** PMCID: PubMed Central keeps the `PMC` prefix as part of the identifier. */
export const PMCID_PATTERN = "PMC[1-9]\\d*";
/**
 * arXiv: the post-2007 `YYMM.NNNNN` form and the pre-2007 `archive[.subject]/YYMMNNN` form,
 * either optionally versioned. Bare `arXiv:` prefixes are not part of the identifier.
 */
export const ARXIV_PATTERN = "\\d{4}\\.\\d{4,5}(?:v\\d+)?|[a-z][a-z-]*(?:\\.[A-Z]{2})?/\\d{7}(?:v\\d+)?";
/** ClinicalTrials.gov NCT number: `NCT` plus exactly eight digits. */
export const CLINICAL_TRIAL_PATTERN = "NCT\\d{8}";
/**
 * ORCID: four groups of four characters where only the final checksum character may be `X`.
 * The mod-11-2 checksum itself is not verified here — a checksum failure is a typo worth
 * surfacing in the editor, but this layer only decides whether a value is structurally an ORCID.
 */
export const ORCID_PATTERN = "\\d{4}-\\d{4}-\\d{4}-\\d{3}[\\dX]";
function anchored(source) {
    return new RegExp(`^(?:${source})$`, "u");
}
const DOI = anchored(DOI_PATTERN);
const PMID = anchored(PMID_PATTERN);
const ORCID = anchored(ORCID_PATTERN);
/** True when `value` is a well-formed DOI. Does not check that it resolves. */
export function isWellFormedDoi(value) {
    return DOI.test(value);
}
/** True when `value` is a well-formed PMID. Does not check that the record exists. */
export function isWellFormedPmid(value) {
    return PMID.test(value);
}
/** True when `value` is a structurally valid ORCID iD. The checksum is not verified. */
export function isWellFormedOrcid(value) {
    return ORCID.test(value);
}
/**
 * The stored form of an ORCID, or `undefined` when the string is not one.
 *
 * Crossref citeproc-JSON (and many other CSL-JSON producers) put the resolver URL on the name
 * object. The canonical creator column is the bare `0000-0000-0000-0000` id; this unwraps the
 * URL so import can fill that column instead of parking a well-formed identifier in extensions.
 */
const ORCID_RESOLVER_PREFIX = /^(?:https?:\/\/)?(?:www\.)?orcid\.org\/+/iu;
export function canonicalOrcid(value) {
    const trimmed = value.trim();
    const bare = trimmed.replace(ORCID_RESOLVER_PREFIX, "");
    return isWellFormedOrcid(bare) ? bare : undefined;
}
/**
 * The §14 scholarly-status vocabulary for `retractionStatus`.
 *
 * It is a closed list because this product sets the value: §14 asks for retraction, correction
 * and expression-of-concern warnings, and a warning can only be rendered from a term the UI
 * knows. `none` is meaningful and distinct from an absent field — it records that the status was
 * checked and the work is clean, where absent records that nobody has looked.
 */
export const RETRACTION_STATUSES = [
    "none",
    "corrected",
    "expression-of-concern",
    "retracted",
    "withdrawn",
];
