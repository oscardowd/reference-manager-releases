/**
 * Document-payload versioning — SPEC §9.3 ("Schema version"), ADR-0015 decision 6, ADR-0098.
 *
 * ADR-0010 says every schema change ships with a migration and a migration test. That rule was
 * written for the database, but a `PortableDocumentPart` lives in somebody else's `.docx`, on
 * a machine we do not control, possibly written by a build we have never seen. It needs the
 * same discipline and one extra rule the database does not need: a payload from the *future*
 * is read-only, because the only alternative is to rewrite a document while silently dropping
 * whatever the newer version put in it.
 *
 * There is exactly one direction. Migrations run forward, 1 → 2 → 3, never back.
 */
import { DOCUMENT_SCOPE_ID } from "./canonical.js";
import { DOC_SCHEMA_VERSION } from "./types.js";
/**
 * Version 1 → 2: give an existing bibliography an explicit whole-document scope (ADR-0123).
 *
 * What it must **not** do is invent scopes for a document that already holds two bibliography
 * controls. That document is R-072's victim, and the migration cannot know which records
 * Chapter 1 listed — the version-1 payload never recorded it, and the visible text in the
 * control is the only remaining evidence, which lives in the manuscript rather than here.
 * Inventing a plausible split would be the same silent flattening it is meant to end, so a
 * migrated document declares one document-wide scope and the refresh path refuses when it
 * finds more controls than scopes.
 *
 * `renderedHash` is copied, not moved: version 1's document-level hash keeps its meaning for
 * any reader that has not learned about scopes, and the per-control hash starts life equal to
 * it because before scopes existed the one control *was* the whole document.
 */
export const MIGRATION_1_TO_2 = {
    from: 1,
    to: 2,
    description: "bibliography controls carry an explicit scope (R-072)",
    migrate: (part) => {
        const bibliography = part["bibliography"];
        if (!isJsonObject(bibliography) || bibliography["scopes"] !== undefined) {
            return { ...part, schemaVersion: 2 };
        }
        const renderedHash = bibliography["renderedHash"];
        const scope = {
            scopeId: DOCUMENT_SCOPE_ID,
            mode: "document",
            ...(typeof renderedHash === "string" ? { renderedHash } : {}),
        };
        return { ...part, schemaVersion: 2, bibliography: { ...bibliography, scopes: [scope] } };
    },
};
/**
 * The real chain. When `DOC_SCHEMA_VERSION` is bumped, the step that lands the new version
 * goes here, and `checkMigrationChain` fails the suite until it does.
 */
export const DOC_SCHEMA_MIGRATIONS = [MIGRATION_1_TO_2];
/**
 * Is a chain a usable ladder from version 1 to `targetVersion`? Called by the reader before it
 * migrates anything (a broken ladder must refuse, not half-climb) and by the test that stops
 * `DOC_SCHEMA_VERSION` being bumped without a migration.
 */
export function checkMigrationChain(migrations, targetVersion = DOC_SCHEMA_VERSION) {
    const problems = [];
    let expected = 1;
    for (const step of migrations) {
        if (step.to <= step.from) {
            problems.push({
                code: "wrong-direction",
                message: `migration ${step.from} → ${step.to} does not move forward; payload migrations are forward-only`,
            });
            return { ok: false, problems };
        }
        if (step.from !== expected) {
            problems.push({
                code: "not-contiguous",
                message: `migration chain jumps from version ${expected} to ${step.from}; every version needs a step`,
            });
            return { ok: false, problems };
        }
        expected = step.to;
    }
    if (expected < targetVersion) {
        problems.push({
            code: "does-not-reach-target",
            message: `migration chain reaches version ${expected} but DOC_SCHEMA_VERSION is ${targetVersion}; the version bump needs a migration`,
        });
    }
    else if (expected > targetVersion) {
        problems.push({
            code: "overshoots-target",
            message: `migration chain reaches version ${expected}, past DOC_SCHEMA_VERSION ${targetVersion}`,
        });
    }
    return { ok: problems.length === 0, problems };
}
function isJsonObject(value) {
    return typeof value === "object" && value !== null && !Array.isArray(value);
}
/**
 * Deep copy that keeps a payload a payload: objects are rebuilt with `Object.fromEntries`,
 * which defines own data properties, so a `"__proto__"` key that arrived over the wire stays
 * ordinary data instead of reaching `Object.prototype`.
 *
 * `structuredClone` would also be safe, but it throws on values JSON can hold no opinion
 * about; here the input has always come through `JSON.parse` or a typed literal.
 */
export function clonePayload(value) {
    if (Array.isArray(value)) {
        return value.map((entry) => clonePayload(entry));
    }
    if (isJsonObject(value)) {
        return Object.fromEntries(Object.entries(value).map(([key, entry]) => [key, clonePayload(entry)]));
    }
    return value;
}
/**
 * Bring a parsed payload up to `targetVersion`, or say precisely why not.
 *
 * A payload declaring a version *newer* than we support is not an error and not garbage: it is
 * a document written by a newer build, and the answer is "read it, never write it"
 * (ADR-0015 decision 6). The caller gets `status: "future"` and no part, because handing back a
 * part invites somebody to write it.
 */
export function migratePayload(value, options = {}) {
    const migrations = options.migrations ?? DOC_SCHEMA_MIGRATIONS;
    const targetVersion = options.targetVersion ?? DOC_SCHEMA_VERSION;
    if (!isJsonObject(value)) {
        return { status: "refused", reason: "not-an-object", message: "document part must be a JSON object" };
    }
    const declared = value["schemaVersion"];
    if (typeof declared !== "number" || !Number.isInteger(declared) || declared < 1) {
        return {
            status: "refused",
            reason: "no-schema-version",
            message: "schemaVersion must be a positive integer; a payload without one cannot be placed on the ladder",
        };
    }
    if (declared > targetVersion) {
        return { status: "future", declaredVersion: declared, targetVersion };
    }
    const chain = checkMigrationChain(migrations, targetVersion);
    if (!chain.ok) {
        return {
            status: "refused",
            reason: "broken-chain",
            message: chain.problems.map((p) => p.message).join("; "),
        };
    }
    let part = clonePayload(value);
    const applied = [];
    for (const step of migrations) {
        if (step.from < declared) {
            continue;
        }
        let next;
        try {
            next = step.migrate(part);
        }
        catch (cause) {
            return {
                status: "refused",
                reason: "migration-threw",
                message: `migration ${step.from} → ${step.to} (${step.description}) failed: ${cause instanceof Error ? cause.message : String(cause)}`,
            };
        }
        if (!isJsonObject(next)) {
            return {
                status: "refused",
                reason: "migration-returned-non-object",
                message: `migration ${step.from} → ${step.to} (${step.description}) did not return an object`,
            };
        }
        if (next["schemaVersion"] !== step.to) {
            return {
                status: "refused",
                reason: "migration-did-not-set-version",
                message: `migration ${step.from} → ${step.to} (${step.description}) left schemaVersion at ${String(next["schemaVersion"])}`,
            };
        }
        part = next;
        applied.push({ from: step.from, to: step.to, description: step.description });
    }
    // A contiguous chain can still skip a version: a single 1 → 3 step is contiguous, and a
    // payload declaring 2 matches none of its steps. Landing short is a silent downgrade, so it
    // is a refusal rather than an "ok" at the wrong version.
    if (part["schemaVersion"] !== targetVersion) {
        return {
            status: "refused",
            reason: "no-path-from-version",
            message: `no migration path from schemaVersion ${declared} to ${targetVersion}; the chain skips version ${declared}`,
        };
    }
    return {
        status: applied.length === 0 ? "current" : "migrated",
        part,
        fromVersion: declared,
        applied,
    };
}
