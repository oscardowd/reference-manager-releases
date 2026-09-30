/**
 * The read/write seam for a portable document part — SPEC §9.3, ADR-0015, ADR-0098.
 *
 * Every integration (Word add-in, bridge service, CLI, Google Docs) goes through these two
 * functions rather than calling `JSON.parse` and casting. What the seam is actually for:
 *
 * - **A payload is untrusted input.** It came out of a file a collaborator emailed. It gets
 *   parsed, migrated and validated before anybody sees a typed object.
 * - **A newer payload is read-only, never rewritten** (ADR-0015 decision 6). The reader hands
 *   back no typed part in that case, so there is nothing convenient to pass to the writer.
 * - **What we do not understand survives.** The reader returns the payload itself, not a
 *   reconstruction of the fields we know about, and the writer serialises whatever it is given.
 *   A reconstruction would drop a future version's fields on every save — the §5.3 no-silent-
 *   loss rule, applied to documents.
 */
import { canonicalJson } from "./canonical.js";
import { DOC_SCHEMA_VERSION } from "./types.js";
import { validateDocumentPart } from "./validate.js";
import { migratePayload, } from "./version.js";
/**
 * Parse (when given a string), migrate and validate a document payload.
 *
 * Returns a result rather than throwing: a damaged document is something to report to the user
 * through §9.5 diagnostics, not an exception thrown out of an add-in task pane.
 */
export function readDocumentPart(input, options = {}) {
    let value = input;
    if (typeof input === "string") {
        try {
            value = JSON.parse(input);
        }
        catch (cause) {
            return {
                status: "unreadable",
                writable: false,
                reason: "not-json",
                message: `document part is not valid JSON: ${cause instanceof Error ? cause.message : String(cause)}`,
                issues: [],
                raw: input,
            };
        }
    }
    const migrated = migratePayload(value, options);
    if (migrated.status === "future") {
        return {
            status: "read-only",
            writable: false,
            reason: "future-schema-version",
            declaredVersion: migrated.declaredVersion,
            supportedVersion: migrated.targetVersion,
            message: `document was written with document-schema version ${migrated.declaredVersion}; this build supports ${migrated.targetVersion}. It can be read but must not be written back.`,
            raw: value,
        };
    }
    if (migrated.status === "refused") {
        return {
            status: "unreadable",
            writable: false,
            reason: migrated.reason,
            message: migrated.message,
            issues: [],
            raw: value,
        };
    }
    const validation = validateDocumentPart(migrated.part, {
        supportedVersion: options.targetVersion ?? DOC_SCHEMA_VERSION,
    });
    if (!validation.ok) {
        return {
            status: "unreadable",
            writable: false,
            reason: "invalid-payload",
            message: "document part failed schema validation",
            issues: validation.issues,
            raw: value,
        };
    }
    return {
        status: "ok",
        writable: true,
        part: migrated.part,
        fromVersion: migrated.fromVersion,
        applied: migrated.applied,
        validation,
        raw: value,
    };
}
/**
 * Serialise a part for storage in a document, or refuse.
 *
 * Two refusals, and both exist because the alternative damages a file the user cannot easily
 * get back. An **invalid** payload must never reach a document — a half-written part is worse
 * than no part, because the controls stay in the text and point at nothing. And a payload whose
 * `schemaVersion` is not the current one must never be written: a newer one would lose the
 * newer build's fields, and an older one would silently downgrade a document that a colleague's
 * newer build will read next.
 *
 * Output is canonical JSON (ADR-0015 decision 5), so the bytes are reproducible and the
 * cluster fingerprints in the content-control tags mean the same thing to every implementation.
 */
export function writeDocumentPart(part, targetVersion = DOC_SCHEMA_VERSION) {
    const declared = typeof part === "object" && part !== null && !Array.isArray(part)
        ? part["schemaVersion"]
        : undefined;
    if (declared !== targetVersion) {
        return {
            status: "refused",
            reason: "not-current-version",
            message: `refusing to write a document part at schemaVersion ${String(declared)}; this build writes version ${targetVersion} only`,
            issues: [],
        };
    }
    const validation = validateDocumentPart(part, { supportedVersion: targetVersion });
    if (!validation.ok) {
        return {
            status: "refused",
            reason: "invalid-payload",
            message: "refusing to write a document part that fails schema validation",
            issues: validation.issues,
        };
    }
    return { status: "ok", json: canonicalJson(part), validation };
}
