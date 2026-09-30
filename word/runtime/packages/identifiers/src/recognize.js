/**
 * Identifier recognition for §5.1 "add records using DOI, PMID, PMCID, ISBN, arXiv ID, URL,
 * clinical trial identifier, multiple identifiers pasted at once" (task E04-08.1).
 *
 * This module answers one question — *what did the user just paste?* — and answers it offline.
 * Retrieval of metadata for a recognised identifier is E04-08.2; comparing and merging what the
 * sources return is E04-09. Keeping recognition separate is what makes it exhaustively testable:
 * there is no network, no clock and no ordering, so every rule below is pinned by a fixture.
 *
 * Three rules shape the design:
 *
 * 1. **Nothing is dropped silently.** Every non-blank input line either contributes an
 *    identifier, is reported as a duplicate of one, or appears in `problems` with a reason. A
 *    test asserts that accounting holds, because "we imported 7 of your 9 identifiers" is only
 *    an acceptable outcome if the product can say which two and why.
 * 2. **A malformed identifier is refused, not repaired.** A failed ISBN check digit names a
 *    different real book, so it is reported rather than accepted or quietly discarded.
 * 3. **Ambiguity is decided by shape, never by guessing.** Bare numbers are read as PMIDs or
 *    ISBNs only on a line that contains nothing but numbers; inside prose an identifier must be
 *    labelled or carry its own prefix. The alternative — treating a page number or a year in a
 *    pasted reference as a PMID — produces confident nonsense.
 *
 * Values returned here echo user input (`matchedText`, `problems[].text`). They are for display
 * and must never reach a log (invariant 6, SPEC §23).
 */
import { ARXIV_PATTERN, CLINICAL_TRIAL_PATTERN, DOI_PATTERN, PMCID_PATTERN, } from "@refmgr/core";
import { isValidIsbn, isbnComparisonKey, normalizeIsbn } from "@refmgr/core";
/** The identifier classes §5.1 requires. `url` is the deliberate fallback for a web page. */
export const IDENTIFIER_KINDS = [
    "doi",
    "pmid",
    "pmcid",
    "isbn",
    "arxiv",
    "nct",
    "url",
];
/* ------------------------------------------------------------------ text helpers */
const CLOSERS = { ")": "(", "]": "[", "}": "{", ">": "<" };
const TRAILING_PUNCTUATION = ".,;:!?'\"";
function occurrences(text, character) {
    let count = 0;
    for (const candidate of text)
        if (candidate === character)
            count += 1;
    return count;
}
/**
 * Drop punctuation a sentence left behind. `\S+` in the DOI grammar is greedy by design — a DOI
 * suffix may contain almost anything — so a DOI quoted mid-sentence arrives with the full stop
 * attached. A closing bracket is only dropped when nothing opened it, because `10.1000/(x)` is
 * a legal DOI.
 */
export function trimTrailingPunctuation(text) {
    let end = text.length;
    while (end > 0) {
        const character = text[end - 1];
        if (TRAILING_PUNCTUATION.includes(character)) {
            end -= 1;
            continue;
        }
        const opener = CLOSERS[character];
        if (opener !== undefined) {
            const slice = text.slice(0, end);
            if (occurrences(slice, character) > occurrences(slice, opener)) {
                end -= 1;
                continue;
            }
        }
        break;
    }
    return text.slice(0, end);
}
function stripLeadingZeros(digits) {
    const trimmed = digits.replace(/^0+/u, "");
    return trimmed === "" ? undefined : trimmed;
}
/* ------------------------------------------------------------------ per-kind values */
function doiCandidate(start, end, doi, form, sourceUrl) {
    const value = trimTrailingPunctuation(doi);
    if (!new RegExp(`^(?:${DOI_PATTERN})$`, "u").test(value))
        return undefined;
    return {
        start,
        end,
        kind: "doi",
        value,
        // DOI syntax is case-insensitive, so comparison folds case; the registrant's own casing is
        // kept in `value` because that is how publishers print it.
        key: value.toLowerCase(),
        form,
        ...(sourceUrl === undefined ? {} : { sourceUrl }),
    };
}
function pmidCandidate(start, end, digits, form, sourceUrl) {
    const value = stripLeadingZeros(digits);
    if (value === undefined || value.length > 9)
        return undefined;
    return {
        start,
        end,
        kind: "pmid",
        value,
        key: value,
        form,
        ...(sourceUrl === undefined ? {} : { sourceUrl }),
    };
}
function pmcidCandidate(start, end, digits, form, sourceUrl) {
    const body = stripLeadingZeros(digits);
    if (body === undefined)
        return undefined;
    const value = `PMC${body}`;
    return {
        start,
        end,
        kind: "pmcid",
        value,
        key: value,
        form,
        ...(sourceUrl === undefined ? {} : { sourceUrl }),
    };
}
const ARXIV_ANCHORED = new RegExp(`^(?:${ARXIV_PATTERN})$`, "u");
/**
 * The post-2007 arXiv identifier is `YYMM.NNNNN`. A bare four-digit group is only accepted as
 * one when its second pair is a real month: without that check every `2020.12345` in a pasted
 * table of numbers becomes an arXiv id.
 */
function plausibleArxivYearMonth(value) {
    const modern = /^(\d{2})(\d{2})\./u.exec(value);
    if (modern === null)
        return true;
    const month = Number(modern[2]);
    return month >= 1 && month <= 12;
}
function arxivCandidate(start, end, id, form, sourceUrl) {
    if (!ARXIV_ANCHORED.test(id) || !plausibleArxivYearMonth(id))
        return undefined;
    return {
        start,
        end,
        kind: "arxiv",
        value: id,
        // The version suffix stays in the key: `2401.01234v1` and `v2` are different documents, and
        // collapsing them would hand the user a record they did not ask for.
        key: id.toLowerCase(),
        form,
        ...(sourceUrl === undefined ? {} : { sourceUrl }),
    };
}
function nctCandidate(start, end, id, form, sourceUrl) {
    const value = id.toUpperCase();
    if (!new RegExp(`^(?:${CLINICAL_TRIAL_PATTERN})$`, "u").test(value))
        return undefined;
    return {
        start,
        end,
        kind: "nct",
        value,
        key: value,
        form,
        ...(sourceUrl === undefined ? {} : { sourceUrl }),
    };
}
function isbnCandidate(start, end, raw, form) {
    const value = normalizeIsbn(raw);
    const key = isbnComparisonKey(value);
    if (!isValidIsbn(value) || key === undefined)
        return undefined;
    return { start, end, kind: "isbn", value, key, form };
}
/* ------------------------------------------------------------------ URL resolvers */
const DOI_HOSTS = new Set(["doi.org", "dx.doi.org", "www.doi.org"]);
const ARXIV_HOSTS = new Set(["arxiv.org", "www.arxiv.org", "export.arxiv.org"]);
const TRIAL_HOSTS = new Set(["clinicaltrials.gov", "www.clinicaltrials.gov"]);
function segments(pathname) {
    return pathname.split("/").filter((segment) => segment !== "");
}
/**
 * Read a resolver URL as the identifier it addresses, or fall back to the URL itself.
 *
 * The URL is kept on the result as `sourceUrl` rather than discarded: it is where the user's
 * paste came from, and §5.1.7's open-access fetch may want the landing page even when the DOI
 * resolves elsewhere.
 */
function fromUrl(start, end, raw) {
    let url;
    try {
        url = new URL(raw);
    }
    catch {
        return undefined;
    }
    const host = url.hostname.toLowerCase();
    const parts = segments(url.pathname);
    const normalized = `${url.protocol}//${url.host}${url.pathname}${url.search}${url.hash}`;
    if (DOI_HOSTS.has(host)) {
        const suffix = decodeURIComponent(url.pathname.slice(1)) + url.search + url.hash;
        const doi = doiCandidate(start, end, suffix, "url", normalized);
        if (doi !== undefined)
            return doi;
    }
    if (host === "pubmed.ncbi.nlm.nih.gov" && parts.length >= 1 && /^\d+$/u.test(parts[0])) {
        const pmid = pmidCandidate(start, end, parts[0], "url", normalized);
        if (pmid !== undefined)
            return pmid;
    }
    if (host === "www.ncbi.nlm.nih.gov" && parts[0] === "pubmed" && /^\d+$/u.test(parts[1] ?? "")) {
        const pmid = pmidCandidate(start, end, parts[1], "url", normalized);
        if (pmid !== undefined)
            return pmid;
    }
    const pmcSegment = host === "pmc.ncbi.nlm.nih.gov" && parts[0] === "articles"
        ? parts[1]
        : host === "www.ncbi.nlm.nih.gov" && parts[0] === "pmc" && parts[1] === "articles"
            ? parts[2]
            : undefined;
    if (pmcSegment !== undefined && new RegExp(`^(?:${PMCID_PATTERN})$`, "iu").test(pmcSegment)) {
        const pmcid = pmcidCandidate(start, end, pmcSegment.slice(3), "url", normalized);
        if (pmcid !== undefined)
            return pmcid;
    }
    if (ARXIV_HOSTS.has(host) && (parts[0] === "abs" || parts[0] === "pdf") && parts.length >= 2) {
        const id = parts.slice(1).join("/").replace(/\.pdf$/iu, "");
        const arxiv = arxivCandidate(start, end, id, "url", normalized);
        if (arxiv !== undefined)
            return arxiv;
    }
    const trialSegment = TRIAL_HOSTS.has(host) && (parts[0] === "study" || (parts[0] === "ct2" && parts[1] === "show"))
        ? parts[parts[0] === "study" ? 1 : 2]
        : undefined;
    if (trialSegment !== undefined) {
        const nct = nctCandidate(start, end, trialSegment, "url", normalized);
        if (nct !== undefined)
            return nct;
    }
    return { start, end, kind: "url", value: normalized, key: normalized, form: "url" };
}
/* ------------------------------------------------------------------ line scanning */
const URL_PATTERN = /\bhttps?:\/\/[^\s<>"]+/giu;
const LABELLED_DOI = /\bdoi\b[\s:]*(10\.\d{4,9}(?:\.\d+)*\/\S+)/giu;
const BARE_DOI = new RegExp(DOI_PATTERN, "giu");
const LABELLED_PMID = /\bpmid\b[\s:]*(\d{1,9})\b/giu;
const LABELLED_PMCID = /\bpmcid\b[\s:]*(?:PMC)?(\d+)\b/giu;
const BARE_PMCID = /\bPMC(\d+)\b/giu;
const BARE_NCT = /\bNCT\d{8}\b/giu;
const LABELLED_ARXIV = /\barxiv\b[\s:]*([^\s,;]+)/giu;
const BARE_ARXIV_MODERN = /(?<![\d.])\d{4}\.\d{4,5}(?:v\d+)?(?![\d.])/gu;
const BARE_ARXIV_LEGACY = /\b[a-z][a-z-]*(?:\.[A-Z]{2})?\/\d{7}(?:v\d+)?\b/gu;
const LABELLED_ISBN = /\bisbn(?:-1[03])?\b[\s:]*([0-9][0-9Xx\s-]{8,20}[0-9Xx])/giu;
/** A line made only of numbers and their separators, where a bare number can only be an id. */
const NUMERIC_LINE = /^[\s0-9Xx,;|·‐‑‒–—-]*$/u;
const NUMERIC_TOKEN = /[0-9][0-9Xx‐‑‒–—-]*/gu;
function scanTrimmedMatch(text, match, group) {
    const raw = match[group];
    const offset = match.index + match[0].indexOf(raw);
    const trimmed = trimTrailingPunctuation(raw);
    return { start: offset, end: offset + trimmed.length, text: trimmed };
}
function collect(line, pattern, group, build, into) {
    pattern.lastIndex = 0;
    let match;
    while ((match = pattern.exec(line)) !== null) {
        const span = scanTrimmedMatch(line, match, group);
        if (span.text === "")
            continue;
        // A labelled match must span its label so it outranks the bare form underneath it.
        const start = group === 0 ? span.start : match.index;
        const candidate = build(start, span.end, span.text);
        if (candidate !== undefined)
            into.push(candidate);
    }
}
/**
 * Bare numbers, read only on a line that holds nothing else. Ten- and thirteen-character values
 * are ISBNs — no PMID is that long — so a check-digit failure there is reported rather than
 * reinterpreted as something shorter.
 */
function collectNumericTokens(line, lineNumber, scan) {
    NUMERIC_TOKEN.lastIndex = 0;
    let match;
    while ((match = NUMERIC_TOKEN.exec(line)) !== null) {
        const raw = match[0];
        const start = match.index;
        const end = start + raw.length;
        const compact = normalizeIsbn(raw);
        if (compact.length === 10 || compact.length === 13) {
            const isbn = isbnCandidate(start, end, compact, "bare");
            if (isbn !== undefined) {
                scan.candidates.push(isbn);
            }
            else {
                scan.problems.push({ line: lineNumber, text: raw, reason: "isbn-check-digit" });
            }
            continue;
        }
        if (/^\d{1,9}$/u.test(compact)) {
            const pmid = pmidCandidate(start, end, compact, "bare");
            if (pmid !== undefined)
                scan.candidates.push(pmid);
        }
    }
}
function scanLine(line, lineNumber) {
    const scan = { candidates: [], problems: [] };
    collect(line, URL_PATTERN, 0, (start, end, text) => fromUrl(start, end, text), scan.candidates);
    collect(line, LABELLED_DOI, 1, (start, end, text) => doiCandidate(start, end, text, "labelled"), scan.candidates);
    collect(line, BARE_DOI, 0, (start, end, text) => doiCandidate(start, end, text, "bare"), scan.candidates);
    collect(line, LABELLED_PMID, 1, (start, end, text) => pmidCandidate(start, end, text, "labelled"), scan.candidates);
    collect(line, LABELLED_PMCID, 1, (start, end, text) => pmcidCandidate(start, end, text, "labelled"), scan.candidates);
    collect(line, BARE_PMCID, 1, (start, end, text) => pmcidCandidate(start, end, text, "bare"), scan.candidates);
    collect(line, BARE_NCT, 0, (start, end, text) => nctCandidate(start, end, text, "bare"), scan.candidates);
    collect(line, LABELLED_ARXIV, 1, (start, end, text) => arxivCandidate(start, end, text, "labelled"), scan.candidates);
    collect(line, BARE_ARXIV_MODERN, 0, (start, end, text) => arxivCandidate(start, end, text, "bare"), scan.candidates);
    collect(line, BARE_ARXIV_LEGACY, 0, (start, end, text) => arxivCandidate(start, end, text, "bare"), scan.candidates);
    // A labelled ISBN is handled here rather than through `collect` so that a check-digit failure
    // is reported per occurrence: one bad ISBN on a line must not be hidden by a good one.
    LABELLED_ISBN.lastIndex = 0;
    let isbnMatch;
    while ((isbnMatch = LABELLED_ISBN.exec(line)) !== null) {
        const raw = isbnMatch[1].trim();
        const start = isbnMatch.index;
        const end = start + isbnMatch[0].length;
        const candidate = isbnCandidate(start, end, raw, "labelled");
        if (candidate === undefined) {
            scan.problems.push({ line: lineNumber, text: raw, reason: "isbn-check-digit" });
        }
        else {
            scan.candidates.push(candidate);
        }
    }
    if (NUMERIC_LINE.test(line) && /\d/u.test(line)) {
        collectNumericTokens(line, lineNumber, scan);
    }
    return scan;
}
/**
 * Keep the leftmost, then longest, non-overlapping candidates.
 *
 * This is what makes `https://doi.org/10.1000/xyz` one DOI rather than a URL *and* a DOI, and
 * what stops an `NCT` sequence inside a DOI suffix being reported as a separate trial.
 */
function resolveOverlaps(candidates) {
    const ordered = [...candidates].sort((a, b) => a.start !== b.start ? a.start - b.start : b.end - b.start - (a.end - a.start));
    const accepted = [];
    let boundary = -1;
    for (const candidate of ordered) {
        if (candidate.start < boundary)
            continue;
        accepted.push(candidate);
        boundary = candidate.end;
    }
    return accepted;
}
/* ------------------------------------------------------------------ public API */
/**
 * Recognise every identifier in a pasted block of text.
 *
 * Order of appearance is preserved. Repeats are folded into `duplicates` rather than dropped,
 * so the UI can say "you pasted this one twice" instead of quietly importing fewer records than
 * the user counted.
 */
export function recognizeIdentifiers(input) {
    const identifiers = [];
    const duplicates = [];
    const problems = [];
    const seen = new Map();
    const lines = input.split(/\r\n|\r|\n/u);
    for (const [index, line] of lines.entries()) {
        const lineNumber = index + 1;
        if (line.trim() === "")
            continue;
        const scan = scanLine(line, lineNumber);
        problems.push(...scan.problems);
        const accepted = resolveOverlaps(scan.candidates);
        if (accepted.length === 0) {
            if (!scan.problems.some((problem) => problem.line === lineNumber)) {
                problems.push({ line: lineNumber, text: line.trim(), reason: "no-identifier" });
            }
            continue;
        }
        for (const candidate of accepted) {
            const matchedText = line.slice(candidate.start, candidate.end);
            const identity = `${candidate.kind}:${candidate.key}`;
            const keptIndex = seen.get(identity);
            if (keptIndex !== undefined) {
                duplicates.push({
                    kind: candidate.kind,
                    key: candidate.key,
                    matchedText,
                    line: lineNumber,
                    keptIndex,
                });
                continue;
            }
            seen.set(identity, identifiers.length);
            identifiers.push({
                kind: candidate.kind,
                value: candidate.value,
                key: candidate.key,
                matchedText,
                form: candidate.form,
                line: lineNumber,
                ...(candidate.sourceUrl === undefined ? {} : { sourceUrl: candidate.sourceUrl }),
            });
        }
    }
    return { identifiers, duplicates, problems };
}
/**
 * Recognise a single pasted identifier, for the one-field case.
 *
 * Deliberately strict: anything that yields more than one identifier, or that left a problem
 * behind, returns `undefined` so the caller routes it to the batch UI rather than importing
 * whichever match happened to come first.
 */
export function recognizeIdentifier(input) {
    const batch = recognizeIdentifiers(input);
    if (batch.identifiers.length !== 1 || batch.problems.length > 0)
        return undefined;
    return batch.identifiers[0];
}
