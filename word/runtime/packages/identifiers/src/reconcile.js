/**
 * Multi-source reconciliation for retrieved metadata candidates (SPEC §5.1 steps 2–6, §14;
 * task E04-09).
 *
 * E04-08.2 stops at independent provider candidates. This layer compares them field by field,
 * keeps the provenance of every value, and — this is the point of the task — **refuses to pick a
 * winner when the sources genuinely disagree**. Invariant 3 says two-sided differences are
 * surfaced, never auto-resolved, and §14 says a stored value is never replaced without showing
 * the proposed change. So a conflicting field leaves this module with *no* chosen value, and
 * `resolveReconciliation` refuses to produce anything writable until every conflict carries an
 * explicit decision.
 *
 * What is *not* a conflict is equally important: two sources writing the same fact differently
 * (`456-78` vs `456-478`, `Smith, J` vs `Smith, John`, a JATS-marked-up abstract vs a plain one)
 * agree. Treating those as conflicts would bury the real disagreements in noise, which is the
 * same failure as auto-picking, just louder. Equivalence is decided by a per-path comparison key;
 * the *stored* value is always one of the values a source actually returned, never a synthesised
 * blend.
 *
 * Nothing here writes a `ReferenceItem`, formats a citation, or performs I/O.
 */
import { isbnComparisonKey, jsonEquals } from "@refmgr/core";
/** Marks a value that came from the library rather than a provider. */
export const EXISTING_SOURCE = "existing";
const DEFAULT_USER_SOURCES = ["user"];
function isJsonObject(value) {
    return typeof value === "object" && value !== null && !Array.isArray(value);
}
/**
 * Comparison text: what two values must share to count as the same fact. Diacritics, case,
 * dash flavour, JATS/HTML markup and terminal punctuation are presentation, not fact.
 */
function normalizeText(value) {
    return value
        .replace(/<[^>]+>/gu, " ")
        .replace(/&(?:amp|#38);/giu, "&")
        .replace(/&(?:nbsp|#160);/giu, " ")
        .replace(/&(?:lt|#60);/giu, "<")
        .replace(/&(?:gt|#62);/giu, ">")
        .normalize("NFKD")
        .replace(/\p{Diacritic}/gu, "")
        .replace(/[‐-―−]/gu, "-")
        .replace(/[‘’‛]/gu, "'")
        .replace(/[“”]/gu, '"')
        .toLowerCase()
        .replace(/[\s\p{P}]+/gu, " ")
        .trim();
}
/**
 * `456-78` and `456-478` are the same pages: providers abbreviate the end page against the start.
 * Expanding the short form is a normalisation, not a repair — the returned string is untouched.
 */
function normalizePageRange(value) {
    const match = /^\s*(\d+)\s*[-‐-―−]\s*(\d+)\s*$/u.exec(value);
    if (match === null)
        return normalizeText(value);
    const [, start = "", end = ""] = match;
    const expanded = end.length < start.length ? `${start.slice(0, start.length - end.length)}${end}` : end;
    return `${start}-${expanded}`;
}
/** A personal name matches on family plus first given initial; `Smith, J` is `Smith, John`. */
function creatorKey(creator) {
    const family = normalizeText(creator.family ?? "");
    // Do not discard non-Latin initials. `李, 伟` and `李, 娜` are different people even though
    // neither given name contains an ASCII letter; treating both as `李:` silently picks one.
    const initial = normalizeText(creator.given ?? "").replace(/\P{Letter}/gu, "").slice(0, 1);
    const literal = normalizeText(creator.literal ?? "");
    return `${creator.role}:${family || literal}:${initial}`;
}
function isCreatorArray(value) {
    return Array.isArray(value) && value.every((entry) => isJsonObject(entry) && "role" in entry);
}
function asCreators(value) {
    return value;
}
function comparisonKey(path, value) {
    if (value === null)
        return "null";
    if (typeof value === "boolean" || typeof value === "number")
        return String(value);
    if (typeof value === "string") {
        if (path === "doi")
            return value.trim().toLowerCase();
        if (path === "pmid")
            return value.trim();
        if (path === "fields.isbn")
            return isbnComparisonKey(value) ?? value.trim();
        return path === "fields.pages" ? normalizePageRange(value) : normalizeText(value);
    }
    if (isCreatorArray(value))
        return asCreators(value).map(creatorKey).join("|");
    if (Array.isArray(value)) {
        return [...value].map((entry) => comparisonKey(path, entry)).sort().join("|");
    }
    return Object.entries(value)
        .map(([key, entry]) => `${key}=${entry === undefined ? "" : comparisonKey(path, entry)}`)
        .sort()
        .join("&");
}
/**
 * Which of two equivalent forms to keep. More information wins: full given names and ORCIDs over
 * initials, mixed case over shouted all-caps, longer plain text over a truncated variant.
 */
function completeness(value) {
    if (typeof value === "string") {
        const text = value.replace(/<[^>]+>/gu, "").trim();
        const mixedCase = /\p{Ll}/u.test(text) && /\p{Lu}/u.test(text) ? 1 : 0;
        return text.length + mixedCase * 2;
    }
    if (isCreatorArray(value)) {
        return asCreators(value).reduce((total, creator) => {
            const given = creator.given ?? "";
            const initialOnly = /^\p{Lu}\.?$/u.test(given.trim());
            return total + (creator.family === undefined ? 0 : 2) + (initialOnly ? 1 : given.length) + (creator.orcid === undefined ? 0 : 5);
        }, 0);
    }
    if (Array.isArray(value))
        return value.length;
    return 0;
}
function rankOf(source, provenance, sourceRank) {
    if (sourceRank === undefined)
        return provenance.confidence;
    const index = sourceRank.indexOf(source);
    // Unranked sources sort below every ranked one but keep their confidence ordering.
    return index === -1 ? provenance.confidence - sourceRank.length - 1 : sourceRank.length - index;
}
/** Independent agreeing evidence compounds: noisy-or, capped so nothing ever claims certainty. */
function combinedConfidence(provenances) {
    const combined = 1 - provenances.reduce((product, entry) => product * (1 - clamp(entry.confidence)), 1);
    return Math.min(0.99, Number(combined.toFixed(4)));
}
function clamp(value) {
    return Number.isFinite(value) ? Math.min(1, Math.max(0, value)) : 0;
}
/** Keep every raw response when one source id contributes more than one candidate. */
function addExtension(extensions, namespace, value) {
    let key = namespace;
    let suffix = 1;
    while (Object.hasOwn(extensions, key)) {
        if (jsonEquals(extensions[key], value))
            return;
        suffix += 1;
        key = `${namespace}:${suffix}`;
    }
    Object.defineProperty(extensions, key, {
        configurable: true,
        enumerable: true,
        value,
        writable: true,
    });
}
function representative(group, sourceRank) {
    let bestIndex = 0;
    for (let index = 1; index < group.forms.length; index += 1) {
        const form = group.forms[index];
        const best = group.forms[bestIndex];
        if (form === undefined || best === undefined)
            continue;
        const formRank = rankOf(group.sources[index] ?? "", group.provenances[index] ?? { source: "", retrievedAt: "", confidence: 0 }, sourceRank);
        const bestRank = rankOf(group.sources[bestIndex] ?? "", group.provenances[bestIndex] ?? { source: "", retrievedAt: "", confidence: 0 }, sourceRank);
        if (completeness(form) > completeness(best) || (completeness(form) === completeness(best) && formRank > bestRank)) {
            bestIndex = index;
        }
    }
    const value = group.forms[bestIndex] ?? group.forms[0];
    const provenance = group.provenances[bestIndex] ?? group.provenances[0];
    if (value === undefined || provenance === undefined) {
        throw new Error("reconciliation value group was empty");
    }
    return { value, provenance };
}
function toFieldValueCandidate(group, sourceRank) {
    const { value, provenance } = representative(group, sourceRank);
    return { value, sources: [...group.sources], provenance, forms: [...group.forms] };
}
function existingProvenance(entry, userSources) {
    return entry.provenance ?? { source: userSources[0] ?? "user", retrievedAt: "", confidence: 1 };
}
function isUserOwned(entry, userSources) {
    return entry.provenance === undefined || userSources.includes(entry.provenance.source);
}
/**
 * Reconcile one identifier's candidates into a single field view.
 *
 * Every candidate must belong to the same identifier; `reconcileBatch` groups them for you.
 */
export function reconcileCandidates(candidates, options = {}) {
    const first = candidates[0];
    if (first === undefined)
        throw new Error("reconcileCandidates requires at least one candidate");
    const userSources = options.userSources ?? DEFAULT_USER_SOURCES;
    const existing = options.existing ?? {};
    const groupsByPath = new Map();
    const extensions = {};
    const sources = [];
    let order = 0;
    const addValue = (path, value, source, provenance) => {
        const key = comparisonKey(path, value);
        const groups = groupsByPath.get(path) ?? [];
        const match = groups.find((group) => group.key === key);
        const rank = rankOf(source, provenance, options.sourceRank);
        if (match === undefined) {
            groups.push({ key, forms: [value], sources: [source], provenances: [provenance], rank, order });
        }
        else {
            match.forms.push(value);
            match.sources.push(source);
            match.provenances.push(provenance);
            match.rank = Math.max(match.rank, rank);
        }
        order += 1;
        groupsByPath.set(path, groups);
    };
    // The stored value goes in first so it is the earliest-seen form and survives every tie-break.
    for (const [path, entry] of Object.entries(existing)) {
        addValue(path, entry.value, EXISTING_SOURCE, existingProvenance(entry, userSources));
    }
    for (const candidate of candidates) {
        if (!sources.includes(candidate.source))
            sources.push(candidate.source);
        for (const [namespace, raw] of Object.entries(candidate.extensions))
            addExtension(extensions, namespace, raw);
        for (const [path, value] of Object.entries(candidate.fields)) {
            const provenance = candidate.provenance[path] ?? {
                source: candidate.source,
                retrievedAt: candidate.retrievedAt,
                confidence: candidate.confidence,
            };
            addValue(path, value, candidate.source, provenance);
        }
    }
    const fields = [];
    const conflicts = [];
    for (const [path, groups] of [...groupsByPath.entries()].sort(([a], [b]) => a.localeCompare(b))) {
        const ordered = [...groups].sort((a, b) => (b.rank - a.rank) || (a.order - b.order));
        const valueCandidates = ordered.map((group) => toFieldValueCandidate(group, options.sourceRank));
        const best = ordered[0];
        const bestCandidate = valueCandidates[0];
        if (best === undefined || bestCandidate === undefined)
            continue;
        const existingEntry = existing[path];
        if (ordered.length === 1) {
            // One agreed fact. If the library already held it, this is a confirmation, not a change.
            const providerProvenances = best.provenances.filter((_, index) => best.sources[index] !== EXISTING_SOURCE);
            const supporting = providerProvenances.length === 0 ? best.provenances : providerProvenances;
            const agreement = best.sources.filter((source) => source !== EXISTING_SOURCE).length > 1
                ? "agreed"
                : "single-source";
            fields.push({
                path,
                agreement,
                value: bestCandidate.value,
                provenance: bestCandidate.provenance,
                confidence: combinedConfidence(supporting),
                candidates: valueCandidates,
            });
            continue;
        }
        // Genuine disagreement. No value is chosen here, by design (invariant 3, §14).
        const reason = existingEntry === undefined
            ? "source-disagreement"
            : isUserOwned(existingEntry, userSources)
                ? "user-edited"
                : "stored-value-differs";
        const suggested = ordered.find((group) => !group.sources.includes(EXISTING_SOURCE)) ?? best;
        const suggestedCandidate = valueCandidates[ordered.indexOf(suggested)] ?? bestCandidate;
        conflicts.push({
            path,
            reason,
            candidates: valueCandidates,
            suggested: suggestedCandidate.value,
            ...(existingEntry === undefined ? {} : { existingValue: existingEntry.value }),
        });
        fields.push({ path, agreement: "conflict", candidates: valueCandidates });
    }
    const warnings = detectWarnings(candidates);
    return {
        ...(first.identifier === undefined ? {} : { identifier: first.identifier }),
        sources,
        fields,
        conflicts,
        warnings,
        confidence: recordConfidence(fields),
        extensions,
    };
}
/**
 * §14's metadata confidence score: the mean field confidence, with every conflicting field
 * counting as zero. A record whose sources disagree is, correctly, a less trustworthy record.
 */
function recordConfidence(fields) {
    if (fields.length === 0)
        return 0;
    const total = fields.reduce((sum, field) => sum + (field.confidence ?? 0), 0);
    return Number((total / fields.length).toFixed(4));
}
/** Reconcile a whole retrieval result, one record per identifier, plus duplicate warnings. */
export function reconcileBatch(input, options = {}) {
    const candidates = Array.isArray(input)
        ? input
        : input.records;
    const failures = Array.isArray(input) ? [] : input.failures;
    const grouped = new Map();
    const identifierOrder = [];
    for (const [position, candidate] of candidates.entries()) {
        // A candidate with no identifier is its own group, and it is **never** grouped with another
        // one. Two identifier-less records agree about nothing that could establish they are the same
        // work — that is what having no identifier means — so merging them would be the reconciler
        // choosing, which is precisely what it exists not to do (invariant 3). The key is the
        // candidate's position, and the NUL prefix cannot collide with any `kind:key`: every kind
        // comes from IDENTIFIER_KINDS and none of them contains one.
        const key = candidate.identifier === undefined
            ? `\u0000unidentified:${position}`
            : `${candidate.identifier.kind}:${candidate.identifier.key}`;
        const bucket = grouped.get(key);
        if (bucket === undefined) {
            grouped.set(key, [candidate]);
            identifierOrder.push(key);
        }
        else {
            bucket.push(candidate);
        }
    }
    const records = identifierOrder.map((key) => reconcileCandidates(grouped.get(key) ?? [], options));
    const withDuplicates = addDuplicateWarnings(records);
    const unresolved = [];
    const seenUnresolved = new Set();
    for (const failure of failures) {
        const key = `${failure.identifier.kind}:${failure.identifier.key}`;
        if (grouped.has(key) || seenUnresolved.has(key))
            continue;
        seenUnresolved.add(key);
        unresolved.push(failure.identifier);
    }
    return { records: withDuplicates, unresolved };
}
/** Identity keys a record can be recognised by across two different pasted identifiers. */
function identityKeys(record) {
    const keys = [];
    for (const field of record.fields) {
        if (field.path !== "doi" && field.path !== "pmid" && field.path !== "fields.isbn")
            continue;
        for (const candidate of field.candidates) {
            if (typeof candidate.value === "string")
                keys.push(`${field.path}=${comparisonKey(field.path, candidate.value)}`);
        }
    }
    return keys;
}
/** Same work, different metadata identity: normalised title, year and first author family. */
function workKey(record) {
    const field = (path) => record.fields.find((entry) => entry.path === path)?.candidates[0]?.value;
    const title = field("title");
    if (typeof title !== "string" || normalizeText(title) === "")
        return undefined;
    const year = field("issuedYear");
    const creators = field("creators");
    const firstAuthor = creators !== undefined && isCreatorArray(creators)
        ? asCreators(creators).find((creator) => creator.role === "author")
        : undefined;
    const firstFamily = firstAuthor === undefined ? "" : creatorKey(firstAuthor);
    return `${normalizeText(title)}|${typeof year === "number" ? year : ""}|${firstFamily}`;
}
/**
 * How the *other* record in a duplicate pair is named in the warning's evidence.
 *
 * A record with no identifier has nothing to name it by, and `null` says that rather than an empty
 * string that reads like a value. The pair is still reported: two identifier-less records that
 * describe the same work is exactly the case a file import produces, and it is the user's to
 * settle (§14, invariant 3).
 */
function duplicateDetail(other) {
    if (other === undefined)
        return { otherIdentifierKind: null, otherIdentifierValue: null };
    return { otherIdentifierKind: other.kind, otherIdentifierValue: other.value };
}
function addDuplicateWarnings(records) {
    const extra = records.map(() => []);
    for (let a = 0; a < records.length; a += 1) {
        for (let b = a + 1; b < records.length; b += 1) {
            const left = records[a];
            const right = records[b];
            if (left === undefined || right === undefined)
                continue;
            const shared = identityKeys(left).find((key) => identityKeys(right).includes(key));
            const leftWork = workKey(left);
            const sameWork = leftWork !== undefined && leftWork === workKey(right);
            if (shared === undefined && !sameWork)
                continue;
            const code = shared === undefined ? "duplicate-work" : "duplicate-identifier";
            const message = shared === undefined
                ? "another identifier in this batch describes the same work"
                : `another identifier in this batch shares ${shared.split("=")[0] ?? "an identifier"}`;
            extra[a]?.push({ code, role: "subject", source: "reconciler", message, detail: duplicateDetail(right.identifier) });
            extra[b]?.push({ code, role: "subject", source: "reconciler", message, detail: duplicateDetail(left.identifier) });
        }
    }
    return records.map((record, index) => {
        const added = extra[index];
        return added === undefined || added.length === 0
            ? record
            : { ...record, warnings: [...record.warnings, ...added] };
    });
}
function stringsOf(value) {
    if (typeof value === "string")
        return [value];
    if (Array.isArray(value))
        return value.flatMap(stringsOf);
    return [];
}
function noticeCode(text) {
    const lowered = text.toLowerCase();
    if (lowered.includes("expression of concern") || lowered.includes("expression_of_concern"))
        return "expression-of-concern";
    if (lowered.includes("retract"))
        return "retraction";
    if (lowered.includes("withdraw"))
        return "withdrawn";
    if (lowered.includes("erratum") || lowered.includes("corrigend") || lowered.includes("correction"))
        return "correction";
    return undefined;
}
/**
 * §5.1.6 warnings, read from the provider payloads the retrieval layer preserved verbatim under
 * `extensions[source]`. Detection is deliberately conservative: a warning that fires on healthy
 * records trains the user to ignore it.
 */
function detectWarnings(candidates) {
    const warnings = [];
    const seen = new Set();
    const push = (warning) => {
        const key = `${warning.code}:${warning.role}:${warning.source}`;
        if (seen.has(key))
            return;
        seen.add(key);
        warnings.push(warning);
    };
    for (const candidate of candidates) {
        const raw = candidate.extensions[candidate.source];
        if (raw === undefined)
            continue;
        switch (candidate.source) {
            case "crossref":
                detectCrossref(raw, push);
                break;
            case "pubmed":
                // PubMed keys its `result` object by the id that was requested, so with no identifier
                // there is no row to look up and no notice to detect. Only a provider response reaches
                // this branch, and a provider is always asked by identifier.
                if (candidate.identifier !== undefined)
                    detectPubMed(raw, candidate.identifier.value, push);
                break;
            case "clinicaltrials":
                detectClinicalTrials(raw, push);
                break;
            case "arxiv":
                detectArxiv(raw, push);
                break;
            default:
                break;
        }
    }
    return warnings;
}
function objectAt(value, key) {
    if (!isJsonObject(value))
        return undefined;
    const nested = value[key];
    return isJsonObject(nested) ? nested : undefined;
}
function detectCrossref(raw, push) {
    const message = objectAt(raw, "message") ?? raw;
    // `updated-by`: something has retracted or corrected *this* record.
    // `update-to`:  *this* record is the notice acting on something else.
    for (const [key, role] of [["updated-by", "subject"], ["update-to", "notice"]]) {
        const entries = message[key];
        if (!Array.isArray(entries))
            continue;
        for (const entry of entries) {
            if (!isJsonObject(entry))
                continue;
            const type = typeof entry.type === "string" ? entry.type : "";
            const code = noticeCode(type);
            if (code === undefined)
                continue;
            const doi = typeof entry.DOI === "string" ? entry.DOI : undefined;
            push({
                code,
                role,
                source: "crossref",
                message: role === "subject"
                    ? `Crossref reports this record as the subject of a ${type.replace(/_/gu, " ")}`
                    : `Crossref reports this record as a ${type.replace(/_/gu, " ")} notice`,
                ...(doi === undefined ? {} : { detail: { relatedDoi: doi } }),
            });
        }
    }
}
function detectPubMed(raw, id, push) {
    const result = objectAt(raw, "result") ?? raw;
    const row = objectAt(result, id);
    if (row === undefined)
        return;
    for (const type of stringsOf(row.pubtype)) {
        const lowered = type.toLowerCase();
        const code = noticeCode(lowered);
        if (code === undefined)
            continue;
        // "Retracted Publication" is the retracted article; "Retraction of Publication" is the notice.
        const role = /\bof publication\b|erratum|expression of concern in/u.test(lowered)
            ? "notice"
            : "subject";
        push({ code, role, source: "pubmed", message: `PubMed publication type "${type}"` });
    }
}
function detectClinicalTrials(raw, push) {
    const status = objectAt(objectAt(raw, "protocolSection"), "statusModule")?.overallStatus;
    if (typeof status !== "string" || !/^withdrawn$/iu.test(status.trim()))
        return;
    push({
        code: "withdrawn",
        role: "subject",
        source: "clinicaltrials",
        message: "ClinicalTrials.gov reports the study as withdrawn",
        detail: { overallStatus: status },
    });
}
function detectArxiv(raw, push) {
    const xml = typeof raw.rawRecord === "string" ? raw.rawRecord : undefined;
    if (xml === undefined)
        return;
    const comment = /<arxiv:comment(?:\s[^>]*)?>([\s\S]*?)<\/arxiv:comment>/iu.exec(xml)?.[1];
    if (comment === undefined || !/has been withdrawn/iu.test(comment))
        return;
    push({ code: "withdrawn", role: "subject", source: "arxiv", message: "the arXiv submission comment states the paper was withdrawn" });
}
/**
 * Turn a reconciled record plus explicit decisions into the fields a host may write.
 *
 * **Refuses on any undecided conflict.** `fields` is empty while `status` is `unresolved`, so a
 * caller cannot half-apply an import and cannot reach a written value without having chosen it.
 */
export function resolveReconciliation(record, decisions = []) {
    const conflictsByPath = new Map(record.conflicts.map((conflict) => [conflict.path, conflict]));
    const decided = new Map();
    const rejected = [];
    for (const decision of decisions) {
        const conflict = conflictsByPath.get(decision.path);
        if (conflict === undefined) {
            rejected.push(decision.path);
            continue;
        }
        if (decision.choose === "existing") {
            decided.set(decision.path, "keep-existing");
            continue;
        }
        if (decision.choose === "value") {
            decided.set(decision.path, {
                value: decision.value,
                provenance: decision.provenance ?? { source: "user", retrievedAt: "", confidence: 1 },
            });
            continue;
        }
        const chosen = conflict.candidates.find((candidate) => candidate.sources.includes(decision.source));
        if (chosen === undefined) {
            rejected.push(decision.path);
            continue;
        }
        decided.set(decision.path, { value: chosen.value, provenance: chosen.provenance });
    }
    const pending = record.conflicts.map((conflict) => conflict.path).filter((path) => !decided.has(path));
    if (pending.length > 0)
        return { status: "unresolved", fields: {}, pending, rejected };
    const fields = {};
    for (const field of record.fields) {
        if (field.agreement === "conflict") {
            const choice = decided.get(field.path);
            if (choice !== undefined && choice !== "keep-existing")
                fields[field.path] = choice;
            continue;
        }
        if (field.value === undefined || field.provenance === undefined)
            continue;
        fields[field.path] = { value: field.value, provenance: field.provenance };
    }
    return { status: "resolved", fields, pending: [], rejected };
}
