/**
 * Canonical JSON serialization and fingerprints — ADR-0015 decision 5.
 *
 * The fingerprint embedded in a Word content-control tag (ADR-0007) must be reproducible by
 * any future implementation, so the serialization is fixed here: recursively sorted object
 * keys, no insignificant whitespace, UTF-8, SHA-256.
 */
import { sha256Hex } from "./sha256.js";
/** Hex characters of SHA-256 kept in the content-control tag fingerprint. */
export const SHORT_FINGERPRINT_LENGTH = 12;
export const CLUSTER_TAG_PREFIX = "refmgr-cite";
export const BIBLIOGRAPHY_TAG = "refmgr-bib";
function sortValue(value) {
    if (Array.isArray(value)) {
        return value.map(sortValue);
    }
    if (value !== null && typeof value === "object") {
        const entries = Object.entries(value)
            .filter(([, v]) => v !== undefined)
            .sort(([a], [b]) => (a < b ? -1 : a > b ? 1 : 0));
        return Object.fromEntries(entries.map(([k, v]) => [k, sortValue(v)]));
    }
    return value;
}
/**
 * Deterministic JSON: object keys sorted recursively (code-unit order), undefined-valued
 * properties dropped, arrays kept in order. Two structurally equal values always produce
 * byte-identical output.
 */
export function canonicalJson(value) {
    return JSON.stringify(sortValue(value));
}
/**
 * Full SHA-256 hex fingerprint of a value's canonical JSON.
 *
 * Computed by `sha256.ts` rather than `node:crypto` so the same planners run in the Word task pane
 * (ADR-0242). The digest is byte-for-byte the one `node:crypto` gives, and `sha256.test.ts` holds it
 * there: the fingerprints already written into users' documents must not change.
 */
export function fingerprint(value) {
    return sha256Hex(canonicalJson(value));
}
/** Truncated fingerprint used in content-control tags. */
export function shortFingerprint(value) {
    return fingerprint(value).slice(0, SHORT_FINGERPRINT_LENGTH);
}
/**
 * The ADR-0007 content-control tag for a citation cluster:
 * `refmgr-cite:<clusterId>:<shortFingerprint(payload)>`. If the custom XML part is lost but
 * the control survives, the tag alone identifies the cluster and detects payload mismatch.
 */
export function clusterTag(clusterId, clusterPayload) {
    return `${CLUSTER_TAG_PREFIX}:${clusterId}:${shortFingerprint(clusterPayload)}`;
}
/**
 * The reserved scope id of a whole-document bibliography — the version-1 shape, and what the
 * 1 → 2 migration gives an existing document. Its control keeps the bare `refmgr-bib` tag, so
 * a document written before scopes existed needs no rewrite of its manuscript to stay valid.
 */
export const DOCUMENT_SCOPE_ID = "document";
export function bibliographyScopeCoverage(scope) {
    return {
        scopeId: scope.scopeId,
        mode: scope.mode,
        ...(scope.itemIds === undefined ? {} : { itemIds: [...scope.itemIds] }),
    };
}
/**
 * The tag of a bibliography control: bare `refmgr-bib` for the whole-document scope, and
 * `refmgr-bib:<scopeId>:<shortFingerprint(coverage)>` for any other (ADR-0123, R-072).
 *
 * ADR-0007 deliberately left this tag unfingerprinted, on the reasoning that there is one
 * bibliography per document whose content derives from every cluster. R-072 is that reasoning
 * failing: a thesis has one per chapter, and with nothing in the tag no refresh could tell
 * which control it was standing in front of.
 */
export function bibliographyTag(scope) {
    if (scope.scopeId === DOCUMENT_SCOPE_ID) {
        return BIBLIOGRAPHY_TAG;
    }
    return `${BIBLIOGRAPHY_TAG}:${scope.scopeId}:${shortFingerprint(bibliographyScopeCoverage(scope))}`;
}
/**
 * Parse a bibliography control tag; null when it is not one of ours.
 *
 * The bare tag is accepted forever: it is in every document written before schema version 2,
 * and refusing to recognise it would turn an existing bibliography into an unknown control.
 */
export function parseBibliographyTag(tag) {
    if (tag === BIBLIOGRAPHY_TAG) {
        return { scopeId: DOCUMENT_SCOPE_ID, shortFingerprint: null };
    }
    const parts = tag.split(":");
    if (parts.length !== 3 || parts[0] !== BIBLIOGRAPHY_TAG) {
        return null;
    }
    const [, scopeId, short] = parts;
    if (!scopeId || !short || short.length !== SHORT_FINGERPRINT_LENGTH) {
        return null;
    }
    return { scopeId, shortFingerprint: short };
}
/** Parse a content-control tag; null when it is not one of ours or is malformed. */
export function parseClusterTag(tag) {
    const parts = tag.split(":");
    if (parts.length !== 3 || parts[0] !== CLUSTER_TAG_PREFIX) {
        return null;
    }
    const [, clusterId, short] = parts;
    if (!clusterId || !short || short.length !== SHORT_FINGERPRINT_LENGTH) {
        return null;
    }
    return { clusterId, shortFingerprint: short };
}
