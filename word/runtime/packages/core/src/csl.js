import { cslDatePartValue, DEFAULT_FIELD_REGISTRY, fieldValueSatisfies, isCslDate, } from "./field-registry.js";
import { cloneJson, isJsonObject, isJsonValue, jsonEquals } from "./json.js";
import { canonicalItemType } from "./item-types.js";
import { canonicalOrcid, isWellFormedDoi, isWellFormedPmid } from "./scholarly-identifiers.js";
import { assertValidReferenceItem } from "./validate.js";
const CSL_NAME_VARIABLES = new Set([
    "author",
    "chair",
    "collection-editor",
    "compiler",
    "composer",
    "container-author",
    "contributor",
    "curator",
    "director",
    "editor",
    "editorial-director",
    "executive-producer",
    "guest",
    "host",
    "illustrator",
    "interviewer",
    "narrator",
    "organizer",
    "original-author",
    "performer",
    "producer",
    "recipient",
    "reviewed-author",
    "script-writer",
    "series-creator",
    "translator",
]);
const CSL_NAME_KEYS = new Set([
    "family",
    "given",
    "suffix",
    "dropping-particle",
    "non-dropping-particle",
    "literal",
]);
export class CslProjectionConflictError extends Error {
    variable;
    constructor(variable) {
        super(`CSL projection has conflicting values for variable "${variable}"`);
        this.name = "CslProjectionConflictError";
        this.variable = variable;
    }
}
function extensionNamespace(extensions, namespace) {
    const value = extensions?.[namespace];
    return value === undefined ? {} : cloneJson(value);
}
function setProjected(target, variable, value) {
    const existing = target[variable];
    if (existing !== undefined && !jsonEquals(existing, value)) {
        throw new CslProjectionConflictError(variable);
    }
    target[variable] = cloneJson(value);
}
function creatorToCsl(creator) {
    const projected = extensionNamespace(creator.extensions, "csl");
    const values = [
        ["family", creator.family],
        ["given", creator.given],
        ["suffix", creator.suffix],
        ["dropping-particle", creator.droppingParticle],
        ["non-dropping-particle", creator.nonDroppingParticle],
        ["literal", creator.literal],
        ["ORCID", creator.orcid],
    ];
    for (const [key, value] of values) {
        if (value !== undefined)
            setProjected(projected, key, value);
    }
    return projected;
}
function issuedYearFromCslDate(value) {
    const dateParts = value["date-parts"];
    if (!Array.isArray(dateParts) || !Array.isArray(dateParts[0]) || dateParts[0].length === 0) {
        return undefined;
    }
    // One reader for both spellings the schema allows, so a `"2020"` year fills the column exactly
    // as `2020` does (R-091). Negative years stay unread here, as they always were: the column is a
    // non-negative year and a BCE date belongs in `extensions.csl` with its era intact.
    const year = cslDatePartValue(dateParts[0][0]);
    if (year !== undefined && year >= 0)
        return year;
    return undefined;
}
function coreIssued(item) {
    const issued = {};
    if (item.issuedRaw !== undefined)
        issued["raw"] = item.issuedRaw;
    if (item.issuedYear !== undefined)
        issued["date-parts"] = [[item.issuedYear]];
    return issued;
}
/**
 * True when `issued` is exactly what the core columns can store: an optional `raw` string and
 * an optional year-only `date-parts`. Anything richer (month, day, circa, season, literal, a
 * range) stays on `extensions.csl.issued` so the processor still sees it.
 */
function isCoreRepresentableIssued(value) {
    const keys = Object.keys(value);
    if (keys.length === 0)
        return false;
    const allowed = new Set(["date-parts", "raw"]);
    if (!keys.every((key) => allowed.has(key)))
        return false;
    const raw = value["raw"];
    const dateParts = value["date-parts"];
    if (raw === undefined && dateParts === undefined)
        return false;
    if (raw !== undefined && !isNonEmptyString(raw))
        return false;
    if (dateParts === undefined)
        return true;
    // A numeric safe-integer year can be held by the column and given back unchanged. A string-form
    // year cannot, because the columns would re-emit it as `[[2020]]` and quietly rewrite what the
    // file said. A string form is therefore never taken out of `extensions.csl` — only its year is
    // read into the column, and the object itself survives byte-for-byte (invariant 2).
    if (!Array.isArray(dateParts) ||
        dateParts.length !== 1 ||
        !Array.isArray(dateParts[0]) ||
        dateParts[0].length !== 1) {
        return false;
    }
    const year = dateParts[0][0];
    return typeof year === "number" && cslDatePartValue(year) !== undefined && year >= 0;
}
/**
 * Overlay core `issuedYear` / `issuedRaw` onto a copied `extensions.csl.issued` without
 * flattening month/day. A mismatched year is a real two-sided edit and is refused (invariant 3).
 */
function issuedValue(item) {
    const fromExtensions = item.extensions["csl"]?.["issued"];
    const hasCore = item.issuedRaw !== undefined || item.issuedYear !== undefined;
    if (isJsonObject(fromExtensions)) {
        if (!hasCore)
            return undefined;
        const merged = cloneJson(fromExtensions);
        if (item.issuedRaw !== undefined)
            setProjected(merged, "raw", item.issuedRaw);
        if (item.issuedYear !== undefined) {
            const existingYear = issuedYearFromCslDate(merged);
            if (existingYear === undefined) {
                if (merged["date-parts"] !== undefined)
                    throw new CslProjectionConflictError("issued");
                merged["date-parts"] = [[item.issuedYear]];
            }
            else if (existingYear !== item.issuedYear) {
                throw new CslProjectionConflictError("issued");
            }
        }
        return merged;
    }
    if (!hasCore)
        return undefined;
    return coreIssued(item);
}
function projectedFieldValue(rows, definition) {
    const ordered = [...rows].sort((left, right) => left.ordinal - right.ordinal);
    if (definition.cardinality === "single")
        return cloneJson(ordered[0].value);
    return ordered.map((row) => cloneJson(row.value));
}
/**
 * Pure canonical-item → CSL-JSON projection (ADR-0005). It produces processor input only;
 * citation strings remain exclusively the responsibility of `packages/citation`.
 */
export function toCslItem(item, registry = DEFAULT_FIELD_REGISTRY) {
    assertValidReferenceItem(item, registry);
    const projected = extensionNamespace(item.extensions, "csl");
    setProjected(projected, "id", item.id);
    setProjected(projected, "type", item.type);
    if (item.citationKey !== undefined)
        setProjected(projected, "citation-key", item.citationKey);
    if (item.title !== undefined)
        setProjected(projected, "title", item.title);
    if (item.containerTitle !== undefined) {
        setProjected(projected, "container-title", item.containerTitle);
    }
    const issued = issuedValue(item);
    if (issued !== undefined)
        setProjected(projected, "issued", issued);
    if (item.doi !== undefined)
        setProjected(projected, "DOI", item.doi);
    if (item.pmid !== undefined)
        setProjected(projected, "PMID", item.pmid);
    const creatorsByRole = new Map();
    for (const creator of item.creators) {
        const creators = creatorsByRole.get(creator.role) ?? [];
        creators.push(creator);
        creatorsByRole.set(creator.role, creators);
    }
    for (const [role, creators] of creatorsByRole) {
        setProjected(projected, role, creators
            .sort((left, right) => left.ordinal - right.ordinal)
            .map((creator) => creatorToCsl(creator)));
    }
    const fieldsByKey = new Map();
    for (const field of item.fields) {
        const rows = fieldsByKey.get(field.key) ?? [];
        rows.push(field);
        fieldsByKey.set(field.key, rows);
    }
    for (const [key, rows] of fieldsByKey) {
        const definition = registry.get(key);
        if (definition.cslVariable !== undefined) {
            setProjected(projected, definition.cslVariable, projectedFieldValue(rows, definition));
        }
    }
    return projected;
}
function isNonEmptyString(value) {
    return typeof value === "string" && value.trim().length > 0;
}
function takeString(source, key) {
    const value = source[key];
    if (!isNonEmptyString(value))
        return undefined;
    delete source[key];
    return value;
}
/**
 * Take a canonical-column string only when it is well formed. A mis-typed identifier is left in
 * `source`, so it survives intact in `extensions.csl` instead of being written into a canonical
 * column or dropped (invariant 2). Import is deliberately not where a format rule is enforced by
 * throwing: the user still gets their reference, and SPEC §14's quality checks report the
 * problem afterwards.
 */
function takeWellFormedString(source, key, isWellFormed) {
    const value = source[key];
    if (!isNonEmptyString(value) || !isWellFormed(value))
        return undefined;
    delete source[key];
    return value;
}
function creatorFromCsl(role, ordinal, value) {
    const source = cloneJson(value);
    for (const key of CSL_NAME_KEYS) {
        const part = source[key];
        if (part !== undefined && !isNonEmptyString(part))
            return undefined;
    }
    const family = takeString(source, "family");
    const given = takeString(source, "given");
    const suffix = takeString(source, "suffix");
    const droppingParticle = takeString(source, "dropping-particle");
    const nonDroppingParticle = takeString(source, "non-dropping-particle");
    const literal = takeString(source, "literal");
    for (const key of CSL_NAME_KEYS)
        delete source[key];
    const personal = [family, given, suffix, droppingParticle, nonDroppingParticle].some((part) => part !== undefined);
    if ((literal === undefined) === !personal)
        return undefined;
    const creator = { role, ordinal };
    if (family !== undefined)
        creator.family = family;
    if (given !== undefined)
        creator.given = given;
    if (suffix !== undefined)
        creator.suffix = suffix;
    if (droppingParticle !== undefined)
        creator.droppingParticle = droppingParticle;
    if (nonDroppingParticle !== undefined)
        creator.nonDroppingParticle = nonDroppingParticle;
    if (literal !== undefined)
        creator.literal = literal;
    const orcid = takeCreatorOrcid(source);
    if (orcid !== undefined)
        creator.orcid = orcid;
    if (Object.keys(source).length > 0)
        creator.extensions = { csl: source };
    return creator;
}
function takeCreatorOrcid(source) {
    for (const key of ["ORCID", "orcid"]) {
        const value = source[key];
        if (typeof value !== "string")
            continue;
        const orcid = canonicalOrcid(value);
        if (orcid === undefined)
            continue;
        delete source[key];
        return orcid;
    }
    return undefined;
}
function takeCreators(source) {
    const creators = [];
    for (const role of CSL_NAME_VARIABLES) {
        const value = source[role];
        if (!Array.isArray(value) || value.length === 0 || !value.every(isJsonObject))
            continue;
        const parsed = value.map((entry, ordinal) => creatorFromCsl(role, ordinal, entry));
        if (parsed.some((entry) => entry === undefined))
            continue;
        delete source[role];
        creators.push(...parsed);
    }
    return creators;
}
function takeIssued(source) {
    const value = source["issued"];
    if (!isJsonObject(value))
        return {};
    if (isCoreRepresentableIssued(value)) {
        delete source["issued"];
        const result = {};
        if (typeof value["raw"] === "string")
            result.issuedRaw = value["raw"];
        const year = issuedYearFromCslDate(value);
        if (year !== undefined)
            result.issuedYear = year;
        return result;
    }
    if (!isCslDate(value))
        return {};
    const result = {};
    const year = issuedYearFromCslDate(value);
    if (year !== undefined)
        result.issuedYear = year;
    if (isNonEmptyString(value["raw"]))
        result.issuedRaw = value["raw"];
    // Leave the full date on the source: month, day, circa, season and literal are not core
    // columns, and deleting them here would be silent loss. Overlaying a matching issuedYear
    // onto the preserved object must not conflict on the way back out.
    return result;
}
function takeRegisteredFields(source, registry) {
    const fields = [];
    for (const definition of registry.definitions()) {
        if (definition.cslVariable === undefined)
            continue;
        const value = source[definition.cslVariable];
        if (value === undefined)
            continue;
        if (definition.cardinality === "single") {
            if (fieldValueSatisfies(value, definition)) {
                fields.push({ key: definition.key, ordinal: 0, value: cloneJson(value) });
                delete source[definition.cslVariable];
            }
            continue;
        }
        if (Array.isArray(value) &&
            value.length > 0 &&
            value.every((entry) => fieldValueSatisfies(entry, definition))) {
            value.forEach((entry, ordinal) => fields.push({ key: definition.key, ordinal, value: cloneJson(entry) }));
            delete source[definition.cslVariable];
        }
    }
    return fields;
}
/**
 * Pure CSL-JSON → canonical item projection. Unsupported or malformed CSL properties survive
 * byte-for-byte in `extensions.csl`; they are never silently discarded (X-01).
 */
export function fromCslItem(value, context, registry = DEFAULT_FIELD_REGISTRY) {
    if (!isJsonObject(value) || !isJsonValue(value)) {
        throw new TypeError("CSL item must be a JSON object");
    }
    const source = cloneJson(value);
    const id = takeString(source, "id");
    const type = takeString(source, "type");
    if (id === undefined || type === undefined) {
        throw new TypeError("CSL item id and type must be non-empty strings");
    }
    const title = takeString(source, "title");
    const citationKey = takeString(source, "citation-key");
    const containerTitle = takeString(source, "container-title");
    const doi = takeWellFormedString(source, "DOI", isWellFormedDoi);
    const pmid = takeWellFormedString(source, "PMID", isWellFormedPmid);
    const issued = takeIssued(source);
    const creators = takeCreators(source);
    const fields = takeRegisteredFields(source, registry);
    const extensions = {};
    if (Object.keys(source).length > 0)
        extensions["csl"] = source;
    const item = {
        id,
        libraryId: context.libraryId,
        type: canonicalItemType(type),
        createdAt: context.createdAt,
        updatedAt: context.updatedAt ?? context.createdAt,
        version: context.version ?? 1,
        fields,
        creators,
        extensions,
        ...issued,
    };
    if (title !== undefined)
        item.title = title;
    if (citationKey !== undefined)
        item.citationKey = citationKey;
    if (containerTitle !== undefined)
        item.containerTitle = containerTitle;
    if (doi !== undefined)
        item.doi = doi;
    if (pmid !== undefined)
        item.pmid = pmid;
    assertValidReferenceItem(item, registry);
    return item;
}
