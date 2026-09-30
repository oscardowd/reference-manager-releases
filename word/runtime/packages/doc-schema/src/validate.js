/**
 * Structural validation for the portable document citation schema — ADR-0015.
 *
 * Hand-rolled on purpose: this package must stay dependency-free so every integration
 * (add-in bundle, CLI, desktop) can carry it without licence or size cost.
 *
 * Severity contract:
 * - "error": the payload violates the schema and must not be written back.
 * - "warning": the payload is usable but carries something newer or unknown (e.g. an
 *   unrecognised locator label). Data is preserved either way — never dropped.
 */
import { CSL_LOCATOR_LABELS, DOC_SCHEMA_VERSION } from "./types.js";
const LOCATOR_LABELS = new Set(CSL_LOCATOR_LABELS);
function isRecord(value) {
    return typeof value === "object" && value !== null && !Array.isArray(value);
}
function isNonEmptyString(value) {
    return typeof value === "string" && value.length > 0;
}
/**
 * Validate an untrusted value as a PortableDocumentPart. Returns issues rather than
 * throwing: documents from collaborators or older versions are data to be surfaced, not
 * exceptions (§9.5 diagnostics build on this).
 */
export function validateDocumentPart(value, options = {}) {
    const supportedVersion = options.supportedVersion ?? DOC_SCHEMA_VERSION;
    const issues = [];
    const error = (path, code, message) => issues.push({ severity: "error", path, code, message });
    const warning = (path, code, message) => issues.push({ severity: "warning", path, code, message });
    if (!isRecord(value)) {
        error("", "invalid-type", "document part must be a JSON object");
        return { ok: false, issues };
    }
    const part = value;
    if (typeof part.schemaVersion !== "number" || !Number.isInteger(part.schemaVersion) || part.schemaVersion < 1) {
        error("/schemaVersion", "invalid-type", "schemaVersion must be a positive integer");
    }
    else if (part.schemaVersion > supportedVersion) {
        error("/schemaVersion", "unsupported-version", `schemaVersion ${part.schemaVersion} is newer than supported version ${supportedVersion}; treat the document as read-only`);
    }
    for (const field of ["docId", "generator", "locale"]) {
        if (!isNonEmptyString(part[field])) {
            error(`/${field}`, "missing-field", `${field} must be a non-empty string`);
        }
    }
    if (!isRecord(part.style)) {
        error("/style", "missing-field", "style must be an object");
    }
    else if (!isNonEmptyString(part.style.id)) {
        error("/style/id", "missing-field", "style.id must be a non-empty string");
    }
    if (!isRecord(part.preferences)) {
        error("/preferences", "missing-field", "preferences must be an object");
    }
    const snapshotIds = new Set();
    if (!isRecord(part.snapshots)) {
        error("/snapshots", "missing-field", "snapshots must be an object keyed by itemId");
    }
    else {
        for (const [key, snap] of Object.entries(part.snapshots)) {
            const base = `/snapshots/${key}`;
            if (!isRecord(snap)) {
                error(base, "invalid-type", "snapshot must be an object");
                continue;
            }
            snapshotIds.add(key);
            if (snap["itemId"] !== key) {
                error(`${base}/itemId`, "snapshot-key-mismatch", `snapshot itemId must equal its key "${key}"`);
            }
            if (!isNonEmptyString(snap["snapshotAt"])) {
                error(`${base}/snapshotAt`, "missing-field", "snapshotAt must be a non-empty ISO 8601 string");
            }
            if (!isNonEmptyString(snap["contentHash"])) {
                error(`${base}/contentHash`, "missing-field", "contentHash must be a non-empty string");
            }
            const csl = snap["csl"];
            if (!isRecord(csl)) {
                error(`${base}/csl`, "missing-field", "csl must be a CSL-JSON item object");
            }
            else {
                if (!isNonEmptyString(csl["id"])) {
                    error(`${base}/csl/id`, "missing-field", "csl.id must be a non-empty string");
                }
                else if (csl["id"] !== snap["itemId"] || csl["id"] !== key) {
                    error(`${base}/csl/id`, "snapshot-csl-id-mismatch", `csl.id must equal the snapshot itemId and key "${key}"`);
                }
                if (!isNonEmptyString(csl["type"])) {
                    error(`${base}/csl/type`, "missing-field", "csl.type must be a non-empty string");
                }
            }
        }
    }
    // §9.4 bibliography scopes (schema version 2, R-072). A scope names which records one
    // control lists, so a scope that names nothing, or two controls claiming one id, is the
    // ambiguity the scope model exists to remove — an error, not a warning.
    if (isRecord(part.bibliography)) {
        const scopes = part.bibliography["scopes"];
        if (scopes !== undefined) {
            if (!Array.isArray(scopes)) {
                error("/bibliography/scopes", "invalid-type", "bibliography.scopes must be an array");
            }
            else {
                const seenScopeIds = new Set();
                scopes.forEach((scope, si) => {
                    const base = `/bibliography/scopes/${si}`;
                    if (!isRecord(scope)) {
                        error(base, "invalid-type", "bibliography scope must be an object");
                        return;
                    }
                    const scopeId = scope["scopeId"];
                    if (!isNonEmptyString(scopeId)) {
                        error(`${base}/scopeId`, "missing-field", "scopeId must be a non-empty string");
                    }
                    else if (seenScopeIds.has(scopeId)) {
                        error(`${base}/scopeId`, "duplicate-scope-id", `scopeId "${scopeId}" appears more than once`);
                    }
                    else {
                        seenScopeIds.add(scopeId);
                    }
                    const mode = scope["mode"];
                    if (mode !== "document" && mode !== "items") {
                        error(`${base}/mode`, "invalid-type", 'bibliography scope mode must be "document" or "items"');
                    }
                    const itemIds = scope["itemIds"];
                    if (mode === "items") {
                        if (!Array.isArray(itemIds) || itemIds.length === 0) {
                            error(`${base}/itemIds`, "empty", 'a scope with mode "items" must list at least one itemId');
                        }
                        else {
                            itemIds.forEach((itemId, ii) => {
                                if (!isNonEmptyString(itemId)) {
                                    error(`${base}/itemIds/${ii}`, "invalid-type", "itemId must be a non-empty string");
                                }
                                else if (isRecord(part.snapshots) && !snapshotIds.has(itemId)) {
                                    error(`${base}/itemIds/${ii}`, "missing-snapshot", `scoped item "${itemId}" has no snapshot; this bibliography could not be rendered offline (§9.3)`);
                                }
                            });
                        }
                    }
                    else if (itemIds !== undefined) {
                        error(`${base}/itemIds`, "conflicting-options", 'itemIds belongs to a scope with mode "items"; a document-wide scope lists every required record');
                    }
                });
            }
        }
    }
    if (!Array.isArray(part.clusters)) {
        error("/clusters", "missing-field", "clusters must be an array");
        return { ok: issues.every((i) => i.severity !== "error"), issues };
    }
    const seenClusterIds = new Set();
    part.clusters.forEach((cluster, ci) => {
        const base = `/clusters/${ci}`;
        if (!isRecord(cluster)) {
            error(base, "invalid-type", "cluster must be an object");
            return;
        }
        const clusterId = cluster["clusterId"];
        if (!isNonEmptyString(clusterId)) {
            error(`${base}/clusterId`, "missing-field", "clusterId must be a non-empty string");
        }
        else if (seenClusterIds.has(clusterId)) {
            error(`${base}/clusterId`, "duplicate-cluster-id", `clusterId "${clusterId}" appears more than once`);
        }
        else {
            seenClusterIds.add(clusterId);
        }
        const manualEdit = cluster["manualEdit"];
        if (!isRecord(manualEdit) || !["none", "kept", "pending"].includes(manualEdit["state"])) {
            error(`${base}/manualEdit`, "missing-field", 'manualEdit.state must be "none", "kept" or "pending"');
        }
        const items = cluster["items"];
        if (!Array.isArray(items) || items.length === 0) {
            error(`${base}/items`, "empty", "a cluster must cite at least one item");
            return;
        }
        items.forEach((item, ii) => {
            const ibase = `${base}/items/${ii}`;
            if (!isRecord(item)) {
                error(ibase, "invalid-type", "cluster item must be an object");
                return;
            }
            const itemId = item["itemId"];
            if (!isNonEmptyString(itemId)) {
                error(`${ibase}/itemId`, "missing-field", "itemId must be a non-empty string");
            }
            else if (snapshotIds.size > 0 || isRecord(part.snapshots)) {
                if (!snapshotIds.has(itemId)) {
                    error(`${ibase}/itemId`, "missing-snapshot", `cited item "${itemId}" has no snapshot; the document would not be editable offline (§9.3)`);
                }
            }
            if (item["suppressAuthor"] === true && item["authorOnly"] === true) {
                error(`${ibase}`, "conflicting-options", "suppressAuthor and authorOnly are mutually exclusive");
            }
            const locator = item["locator"];
            if (locator !== undefined) {
                if (!isRecord(locator) || !isNonEmptyString(locator["label"]) || typeof locator["value"] !== "string") {
                    error(`${ibase}/locator`, "invalid-type", "locator must be {label, value} with string fields");
                }
                else if (!LOCATOR_LABELS.has(locator["label"])) {
                    warning(`${ibase}/locator/label`, "unknown-locator-label", `locator label "${locator["label"]}" is not a known CSL 1.0.2 label; preserved as-is`);
                }
            }
        });
    });
    return { ok: issues.every((i) => i.severity !== "error"), issues };
}
