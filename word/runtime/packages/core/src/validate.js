import { DEFAULT_FIELD_REGISTRY, fieldConstraintViolation, } from "./field-registry.js";
import { DEFAULT_ITEM_TYPE_REGISTRY } from "./item-types.js";
import { isJsonObject, isJsonValue } from "./json.js";
import { isWellFormedDoi, isWellFormedOrcid, isWellFormedPmid } from "./scholarly-identifiers.js";
/**
 * Codes that describe the quality of the user's metadata rather than the integrity of the
 * record. They are reported and never block a save (SPEC §4.4, §14).
 */
const ADVISORY_CODES = new Set([
    "invalid-format",
    "missing-required-field",
]);
const ISO_INSTANT = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d{1,9})?(?:Z|[+-]\d{2}:\d{2})$/;
function isNonEmptyString(value) {
    return typeof value === "string" && value.trim().length > 0;
}
function validateFieldProvenance(value, path, issues) {
    if (!isJsonObject(value)) {
        issues.push({ path, code: "invalid-provenance", message: "field provenance must be an object" });
        return false;
    }
    const provenance = value;
    if (!isNonEmptyString(provenance.source)) {
        issues.push({
            path: `${path}/source`,
            code: "invalid-provenance",
            message: "field provenance source must be a non-empty string",
        });
    }
    if (!isNonEmptyString(provenance.retrievedAt) ||
        !ISO_INSTANT.test(provenance.retrievedAt) ||
        Number.isNaN(Date.parse(provenance.retrievedAt))) {
        issues.push({
            path: `${path}/retrievedAt`,
            code: "invalid-provenance",
            message: "field provenance retrievedAt must be an ISO 8601 instant with a timezone",
        });
    }
    if (typeof provenance.confidence !== "number" ||
        !Number.isFinite(provenance.confidence) ||
        provenance.confidence < 0 ||
        provenance.confidence > 1) {
        issues.push({
            path: `${path}/confidence`,
            code: "invalid-provenance",
            message: "field provenance confidence must be a number from 0 to 1",
        });
    }
    return true;
}
function validateExtensions(extensions, path, issues) {
    if (!isJsonObject(extensions)) {
        issues.push({
            path,
            code: "invalid-extension",
            message: "extensions must be an object of source namespaces",
        });
        return false;
    }
    let valid = true;
    for (const [namespace, value] of Object.entries(extensions)) {
        if (!isNonEmptyString(namespace) || !isJsonObject(value) || !isJsonValue(value)) {
            valid = false;
            issues.push({
                path: `${path}/${namespace}`,
                code: "invalid-extension",
                message: "each extension namespace must be a non-empty key containing a JSON object",
            });
        }
    }
    return valid;
}
function validateCreator(value, index, issues) {
    const path = `/creators/${index}`;
    if (!isJsonObject(value)) {
        issues.push({ path, code: "invalid-creator", message: "creator must be an object" });
        return;
    }
    const creator = value;
    if (!Number.isInteger(creator.ordinal) || creator.ordinal < 0) {
        issues.push({ path: `${path}/ordinal`, code: "invalid-type", message: "ordinal must be a non-negative integer" });
    }
    if (!isNonEmptyString(creator.role)) {
        issues.push({ path: `${path}/role`, code: "empty", message: "creator role must be non-empty" });
    }
    const personalParts = [
        creator.family,
        creator.given,
        creator.suffix,
        creator.droppingParticle,
        creator.nonDroppingParticle,
    ];
    const hasPersonalName = personalParts.some((part) => part !== undefined);
    const hasLiteral = creator.literal !== undefined;
    if (hasLiteral === hasPersonalName) {
        issues.push({
            path,
            code: "invalid-creator",
            message: "creator must have either literal or personal-name components, never both",
        });
    }
    for (const [key, value] of Object.entries(creator)) {
        if ([
            "role",
            "family",
            "given",
            "suffix",
            "droppingParticle",
            "nonDroppingParticle",
            "literal",
            "orcid",
            "affiliation",
        ].includes(key) &&
            value !== undefined &&
            typeof value === "string" &&
            value.trim().length === 0) {
            issues.push({ path: `${path}/${key}`, code: "empty", message: `${key} must not be blank` });
        }
    }
    if (creator.orcid !== undefined && !isWellFormedOrcid(creator.orcid)) {
        issues.push({
            path: `${path}/orcid`,
            code: "invalid-format",
            message: "ORCID must use the 0000-0000-0000-0000 form, where only the last digit may be X",
        });
    }
    if (creator.extensions !== undefined) {
        validateExtensions(creator.extensions, `${path}/extensions`, issues);
    }
}
function hasRequiredField(item, field) {
    switch (field) {
        case "title":
            return isNonEmptyString(item.title);
        case "containerTitle":
            return isNonEmptyString(item.containerTitle);
        case "issued":
            return item.issuedYear !== undefined || isNonEmptyString(item.issuedRaw);
        default: {
            if (!field.startsWith("field:"))
                return false;
            const key = field.slice("field:".length);
            return item.fields.some((entry) => entry.key === key && isNonEmptyString(entry.value));
        }
    }
}
export function validateReferenceItem(value, registry = DEFAULT_FIELD_REGISTRY, itemTypeRegistry = DEFAULT_ITEM_TYPE_REGISTRY) {
    const issues = [];
    if (!isJsonObject(value)) {
        return {
            ok: false,
            issues: [{ path: "", code: "invalid-type", message: "reference item must be an object" }],
        };
    }
    const item = value;
    for (const field of ["id", "libraryId", "type"]) {
        if (!isNonEmptyString(item[field])) {
            issues.push({ path: `/${field}`, code: "empty", message: `${field} must be a non-empty string` });
        }
    }
    const itemType = itemTypeRegistry.get(item.type);
    if (itemType === undefined && isNonEmptyString(item.type)) {
        issues.push({
            path: "/type",
            code: "unknown-item-type",
            message: `item type "${item.type}" is not registered; add it to an ItemTypeRegistry`,
        });
    }
    else if (itemType !== undefined) {
        for (const requiredField of itemType.requiredFields) {
            if (!hasRequiredField(item, requiredField)) {
                issues.push({
                    path: `/${requiredField.startsWith("field:") ? `fields/${requiredField.slice(6)}` : requiredField}`,
                    code: "missing-required-field",
                    message: `${itemType.label} requires ${requiredField}`,
                });
            }
        }
    }
    for (const field of [
        "citationKey",
        "title",
        "containerTitle",
        "issuedRaw",
        "doi",
        "pmid",
    ]) {
        const fieldValue = item[field];
        if (fieldValue !== undefined && !isNonEmptyString(fieldValue)) {
            issues.push({ path: `/${field}`, code: "empty", message: `${field} must not be blank` });
        }
    }
    if (item.doi !== undefined && isNonEmptyString(item.doi) && !isWellFormedDoi(item.doi)) {
        issues.push({
            path: "/doi",
            code: "invalid-format",
            message: "DOI must be a registrant prefix and suffix, such as 10.1000/example",
        });
    }
    if (item.pmid !== undefined && isNonEmptyString(item.pmid) && !isWellFormedPmid(item.pmid)) {
        issues.push({
            path: "/pmid",
            code: "invalid-format",
            message: "PMID must be digits only, without a PMID: prefix",
        });
    }
    if (item.issuedYear !== undefined && (!Number.isInteger(item.issuedYear) || item.issuedYear < 0)) {
        issues.push({
            path: "/issuedYear",
            code: "invalid-type",
            message: "issuedYear must be a non-negative integer",
        });
    }
    if (!Number.isInteger(item.version) || item.version < 1) {
        issues.push({ path: "/version", code: "invalid-type", message: "version must be a positive integer" });
    }
    for (const field of ["createdAt", "updatedAt"]) {
        const fieldValue = item[field];
        if (typeof fieldValue !== "string" ||
            !ISO_INSTANT.test(fieldValue) ||
            Number.isNaN(Date.parse(fieldValue))) {
            issues.push({
                path: `/${field}`,
                code: "invalid-timestamp",
                message: `${field} must be an ISO 8601 instant with a timezone`,
            });
        }
    }
    if (item.deletedAt !== undefined &&
        (!ISO_INSTANT.test(item.deletedAt) || Number.isNaN(Date.parse(item.deletedAt)))) {
        issues.push({
            path: "/deletedAt",
            code: "invalid-timestamp",
            message: "deletedAt must be an ISO 8601 instant with a timezone",
        });
    }
    if (!Array.isArray(item.creators)) {
        issues.push({ path: "/creators", code: "invalid-type", message: "creators must be an array" });
    }
    else {
        item.creators.forEach((creator, index) => validateCreator(creator, index, issues));
        const ordinalsByRole = new Map();
        for (const creator of item.creators.filter((entry) => isJsonObject(entry))) {
            const ordinals = ordinalsByRole.get(creator.role) ?? [];
            ordinals.push(creator.ordinal);
            ordinalsByRole.set(creator.role, ordinals);
        }
        for (const [role, ordinals] of ordinalsByRole) {
            const ordered = [...ordinals].sort((left, right) => left - right);
            const valid = ordered.every((ordinal, index) => ordinal === index);
            if (!valid) {
                issues.push({
                    path: "/creators",
                    code: "duplicate-ordinal",
                    message: `creator ordinals for role "${role}" must be unique and contiguous from zero`,
                });
            }
        }
    }
    if (!Array.isArray(item.fields)) {
        issues.push({ path: "/fields", code: "invalid-type", message: "fields must be an array" });
    }
    else {
        const rowsByKey = new Map();
        item.fields.forEach((field, index) => {
            const path = `/fields/${index}`;
            if (!isJsonObject(field)) {
                issues.push({ path, code: "invalid-type", message: "field must be an object" });
                return;
            }
            const definition = registry.get(field.key);
            if (definition === undefined) {
                issues.push({
                    path: `${path}/key`,
                    code: "unknown-field",
                    message: `field "${field.key}" is not registered; imported values belong in extensions`,
                });
            }
            else if (!isJsonValue(field.value)) {
                issues.push({
                    path: `${path}/value`,
                    code: "invalid-type",
                    message: `field "${field.key}" must contain ${definition.valueType}`,
                });
            }
            else {
                const violation = fieldConstraintViolation(field.value, definition);
                if (violation === "type") {
                    issues.push({
                        path: `${path}/value`,
                        code: "invalid-type",
                        message: `field "${field.key}" must contain ${definition.valueType}`,
                    });
                }
                else if (violation === "value") {
                    issues.push({
                        path: `${path}/value`,
                        code: "invalid-value",
                        message: definition.allowedValues === undefined
                            ? `field "${field.key}" must be a ${definition.shape ?? definition.valueType} value`
                            : `field "${field.key}" must be one of: ${definition.allowedValues.join(", ")}`,
                    });
                }
                else if (violation === "format") {
                    issues.push({
                        path: `${path}/value`,
                        code: "invalid-format",
                        message: `field "${field.key}" is not in the expected format`,
                    });
                }
            }
            if (!Number.isInteger(field.ordinal) || field.ordinal < 0) {
                issues.push({
                    path: `${path}/ordinal`,
                    code: "invalid-type",
                    message: "field ordinal must be a non-negative integer",
                });
            }
            if (field.provenance !== undefined) {
                validateFieldProvenance(field.provenance, `${path}/provenance`, issues);
            }
            const rows = rowsByKey.get(field.key) ?? [];
            rows.push(field);
            rowsByKey.set(field.key, rows);
        });
        for (const [key, rows] of rowsByKey) {
            const definition = registry.get(key);
            const ordered = rows.map((row) => row.ordinal).sort((left, right) => left - right);
            if (definition?.cardinality === "single" && (rows.length !== 1 || ordered[0] !== 0)) {
                issues.push({
                    path: "/fields",
                    code: "cardinality",
                    message: `single-valued field "${key}" must have exactly one row at ordinal zero`,
                });
            }
            else if (!ordered.every((ordinal, index) => ordinal === index)) {
                issues.push({
                    path: "/fields",
                    code: "duplicate-ordinal",
                    message: `field ordinals for "${key}" must be unique and contiguous from zero`,
                });
            }
        }
    }
    validateExtensions(item.extensions, "/extensions", issues);
    // Required metadata and identifier formats are advisory at this layer: callers can save an
    // incomplete or mis-typed capture and surface the warning in the editor, which is what SPEC
    // §14's metadata-quality checks are for. Structural, type and closed-vocabulary errors stay
    // blocking, because storing one of those makes the field meaningless to everything that reads
    // it afterwards (SPEC §4.4).
    return { ok: !issues.some((issue) => !ADVISORY_CODES.has(issue.code)), issues };
}
export class InvalidReferenceItemError extends Error {
    issues;
    constructor(issues) {
        super(`reference item failed validation with ${issues.length} issue(s)`);
        this.name = "InvalidReferenceItemError";
        this.issues = issues;
    }
}
export function assertValidReferenceItem(value, registry = DEFAULT_FIELD_REGISTRY, itemTypeRegistry = DEFAULT_ITEM_TYPE_REGISTRY) {
    const result = validateReferenceItem(value, registry, itemTypeRegistry);
    if (!result.ok)
        throw new InvalidReferenceItemError(result.issues);
}
