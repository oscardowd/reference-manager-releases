import { isJsonObject } from "./json.js";
import { ARXIV_PATTERN, CLINICAL_TRIAL_PATTERN, PMCID_PATTERN, RETRACTION_STATUSES, } from "./scholarly-identifiers.js";
function requireKey(value, label) {
    if (!/^[A-Za-z][A-Za-z0-9._-]*$/.test(value)) {
        throw new Error(`${label} must be a non-empty portable key`);
    }
}
const STRING_VALUE_TYPES = new Set(["string", "string-list"]);
/**
 * A constraint that could never match is a definition bug, not a validation failure at write
 * time, so it is rejected when the registry is built rather than silently rejecting every value
 * a caller later tries to store.
 */
function requireUsableConstraints(definition) {
    const { key, valueType, pattern, allowedValues, shape } = definition;
    if (pattern !== undefined) {
        if (!STRING_VALUE_TYPES.has(valueType)) {
            throw new Error(`field "${key}" cannot constrain a ${valueType} value with a pattern`);
        }
        compilePattern(key, pattern);
    }
    if (allowedValues !== undefined) {
        if (!STRING_VALUE_TYPES.has(valueType)) {
            throw new Error(`field "${key}" cannot constrain a ${valueType} value with allowedValues`);
        }
        if (allowedValues.length === 0) {
            throw new Error(`field "${key}" has an empty allowedValues list, so no value could ever be stored`);
        }
    }
    if (shape !== undefined && valueType !== "json") {
        throw new Error(`field "${key}" shape "${shape}" requires a json value type`);
    }
}
const patternCache = new Map();
function compilePattern(key, source) {
    const cached = patternCache.get(source);
    if (cached !== undefined)
        return cached;
    let compiled;
    try {
        compiled = new RegExp(`^(?:${source})$`, "u");
    }
    catch {
        throw new Error(`field "${key}" has an invalid pattern`);
    }
    patternCache.set(source, compiled);
    return compiled;
}
/**
 * CSL date variables (`accessed`, `issued`, …) are objects. This accepts the properties CSL
 * defines and requires at least one of them, so `{}` is not silently treated as a date.
 */
const CSL_DATE_KEYS = new Set(["circa", "date-parts", "literal", "raw", "season"]);
/**
 * The numeric value of one `date-parts` entry, or `undefined` when the entry is not one.
 *
 * The official CSL-JSON schema types a date part as `string | number`, and a great deal of real
 * data — citeproc's own corpus, several exporters — writes the year as `"2020"`. Reading only the
 * number left `issuedYear` empty for those records, which is the column duplicate detection,
 * sorting and the §14 checks all read (R-091).
 *
 * Only a plain decimal integer is accepted, not every string the schema permits. A date part
 * reading `"spring"` is schema-valid and is still not a number, and writing it into a canonical
 * date column would assert a precision the source does not have; left unrecognised it survives
 * verbatim in `extensions.csl` instead, which is the same trade `takeWellFormedString` makes for
 * a mistyped identifier (invariant 2). The safe-integer bound is what keeps the parse reversible.
 *
 * This is the single definition of "a date part" for the whole product. It used to be written out
 * twice — here and in `issuedYearFromCslDate` — and the two disagreeing is what let an item hold a
 * year column and a string-form `issued` that `toCslItem` then refused to project at all.
 */
export function cslDatePartValue(entry) {
    if (typeof entry === "number")
        return Number.isSafeInteger(entry) ? entry : undefined;
    if (typeof entry !== "string" || !/^-?\d+$/u.test(entry))
        return undefined;
    const parsed = Number(entry);
    return Number.isSafeInteger(parsed) ? parsed : undefined;
}
export function isCslDate(value) {
    if (!isJsonObject(value))
        return false;
    const keys = Object.keys(value);
    if (keys.length === 0 || !keys.every((key) => CSL_DATE_KEYS.has(key)))
        return false;
    const dateParts = value["date-parts"];
    if (dateParts !== undefined) {
        const valid = Array.isArray(dateParts) &&
            dateParts.length > 0 &&
            dateParts.length <= 2 &&
            dateParts.every((part) => Array.isArray(part) &&
                part.length > 0 &&
                part.length <= 3 &&
                part.every((entry) => cslDatePartValue(entry) !== undefined));
        if (!valid)
            return false;
    }
    for (const key of ["literal", "raw", "season"]) {
        const entry = value[key];
        if (entry !== undefined && typeof entry !== "string")
            return false;
    }
    if (value["circa"] !== undefined && typeof value["circa"] !== "boolean")
        return false;
    return true;
}
export class FieldRegistry {
    #byKey;
    #byCslVariable;
    constructor(definitions) {
        const byKey = new Map();
        const byCslVariable = new Map();
        for (const raw of definitions) {
            requireKey(raw.key, "field key");
            if (byKey.has(raw.key))
                throw new Error(`duplicate field key "${raw.key}"`);
            if (raw.cslVariable !== undefined) {
                requireKey(raw.cslVariable, "CSL variable");
                if (byCslVariable.has(raw.cslVariable)) {
                    throw new Error(`duplicate CSL variable "${raw.cslVariable}"`);
                }
            }
            requireUsableConstraints(raw);
            const definition = Object.freeze({
                ...raw,
                ...(raw.allowedValues === undefined
                    ? {}
                    : { allowedValues: Object.freeze([...raw.allowedValues]) }),
            });
            byKey.set(definition.key, definition);
            if (definition.cslVariable !== undefined) {
                byCslVariable.set(definition.cslVariable, definition);
            }
        }
        this.#byKey = byKey;
        this.#byCslVariable = byCslVariable;
    }
    get(key) {
        return this.#byKey.get(key);
    }
    getByCslVariable(variable) {
        return this.#byCslVariable.get(variable);
    }
    definitions() {
        return [...this.#byKey.values()];
    }
    with(definition) {
        return new FieldRegistry([...this.#byKey.values(), definition]);
    }
}
export function fieldValueMatches(value, type) {
    switch (type) {
        case "boolean":
            return typeof value === "boolean";
        case "integer":
            return typeof value === "number" && Number.isInteger(value);
        case "json":
            return true;
        case "number":
            return typeof value === "number";
        case "string":
            return typeof value === "string";
        case "string-list":
            return Array.isArray(value) && value.every((entry) => typeof entry === "string");
    }
}
/**
 * Grade a value against a field definition. Returns `undefined` when the value is fully legal.
 *
 * `fieldValueMatches` answers the storage question — can the column hold it. This answers the
 * canonical question — is it a legal value for *this* field. Importers use the difference: a
 * value that fails here is not coerced and not dropped, it stays in `extensions[<namespace>]`
 * (invariant 2, X-01).
 */
export function fieldConstraintViolation(value, definition) {
    if (!fieldValueMatches(value, definition.valueType))
        return "type";
    if (definition.shape === "csl-date" && !isCslDate(value))
        return "value";
    const strings = typeof value === "string"
        ? [value]
        : definition.valueType === "string-list" && Array.isArray(value)
            ? value.filter((entry) => typeof entry === "string")
            : [];
    if (definition.allowedValues !== undefined) {
        const allowed = new Set(definition.allowedValues);
        if (!strings.every((entry) => allowed.has(entry)))
            return "value";
    }
    if (definition.pattern !== undefined) {
        const compiled = compilePattern(definition.key, definition.pattern);
        if (!strings.every((entry) => compiled.test(entry)))
            return "format";
    }
    return undefined;
}
/** True when the value is legal under every constraint the definition declares. */
export function fieldValueSatisfies(value, definition) {
    return fieldConstraintViolation(value, definition) === undefined;
}
/**
 * Every SPEC §4.2 metadata field that is not already a core column on `ReferenceItem` or a
 * property of a structural `Creator` (task E02-02.1).
 *
 * Three deliberate choices, all of them about not lying to the citation processor:
 *
 * 1. **Not every field has a `cslVariable`.** `subtitle`, `arxivId`, `clinicalTrialId`,
 *    `keywords`, `retractionStatus`, `peerReviewed`, `addedBy` and `modifiedBy` have no standard
 *    CSL equivalent, so they are stored canonically and simply not projected. Inventing a
 *    variable would hand the processor data no style can read correctly (invariant 4).
 * 2. **Only fields we own carry a closed `allowedValues` vocabulary.** `retractionStatus` is set
 *    by our own §14 monitoring, so it is closed. `publicationStatus` maps to CSL `status`, which
 *    is free text in the wild ("in press", "forthcoming"), so constraining it would push valid
 *    imported metadata into `extensions` for no benefit.
 * 3. **`accessed` is a CSL date object, not a formatted string.** Flattening a date to a display
 *    string is the classic way to lose it; the `csl-date` shape keeps it structured.
 *
 * Where the rest of §4.2 lives: Title, Publication title, Publication date, DOI and PMID are core
 * columns; Authors, Editors, Translators, Other contributors, Institutional authors, ORCID and
 * Author affiliations are structural `Creator`s; Date added and Date modified are `createdAt`
 * and `updatedAt`.
 */
export const DEFAULT_FIELD_DEFINITIONS = [
    { key: "shortTitle", valueType: "string", cardinality: "single", cslVariable: "title-short" },
    { key: "subtitle", valueType: "string", cardinality: "single" },
    {
        key: "journalAbbreviation",
        valueType: "string",
        cardinality: "single",
        cslVariable: "container-title-short",
    },
    { key: "volume", valueType: "string", cardinality: "single", cslVariable: "volume" },
    { key: "issue", valueType: "string", cardinality: "single", cslVariable: "issue" },
    { key: "pages", valueType: "string", cardinality: "single", cslVariable: "page" },
    { key: "articleNumber", valueType: "string", cardinality: "single", cslVariable: "number" },
    { key: "edition", valueType: "string", cardinality: "single", cslVariable: "edition" },
    { key: "publisher", valueType: "string", cardinality: "single", cslVariable: "publisher" },
    {
        key: "publisherPlace",
        valueType: "string",
        cardinality: "single",
        cslVariable: "publisher-place",
    },
    {
        key: "accessed",
        valueType: "json",
        cardinality: "single",
        cslVariable: "accessed",
        shape: "csl-date",
    },
    { key: "url", valueType: "string", cardinality: "single", cslVariable: "URL" },
    {
        key: "pmcid",
        valueType: "string",
        cardinality: "single",
        cslVariable: "PMCID",
        pattern: PMCID_PATTERN,
    },
    { key: "arxivId", valueType: "string", cardinality: "single", pattern: ARXIV_PATTERN },
    {
        key: "clinicalTrialId",
        valueType: "string",
        cardinality: "single",
        pattern: CLINICAL_TRIAL_PATTERN,
    },
    { key: "isbn", valueType: "string", cardinality: "single", cslVariable: "ISBN" },
    { key: "issn", valueType: "string", cardinality: "single", cslVariable: "ISSN" },
    { key: "language", valueType: "string", cardinality: "single", cslVariable: "language" },
    { key: "abstract", valueType: "string", cardinality: "single", cslVariable: "abstract" },
    { key: "keywords", valueType: "string", cardinality: "multiple" },
    { key: "rights", valueType: "string", cardinality: "single", cslVariable: "rights" },
    { key: "archive", valueType: "string", cardinality: "single", cslVariable: "archive" },
    { key: "callNumber", valueType: "string", cardinality: "single", cslVariable: "call-number" },
    { key: "extra", valueType: "string", cardinality: "single", cslVariable: "note" },
    {
        key: "retractionStatus",
        valueType: "string",
        cardinality: "single",
        allowedValues: RETRACTION_STATUSES,
    },
    { key: "publicationStatus", valueType: "string", cardinality: "single", cslVariable: "status" },
    { key: "peerReviewed", valueType: "boolean", cardinality: "single" },
    { key: "addedBy", valueType: "string", cardinality: "single" },
    { key: "modifiedBy", valueType: "string", cardinality: "single" },
];
export const DEFAULT_FIELD_REGISTRY = new FieldRegistry(DEFAULT_FIELD_DEFINITIONS);
