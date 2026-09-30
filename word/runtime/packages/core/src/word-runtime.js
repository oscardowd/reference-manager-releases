/**
 * The `@refmgr/core` surface the Word task pane's runtime reaches (ADR-0242 decision 3).
 *
 * The pane resolves `@refmgr/core` to this file, not to the package index: the index re-exports the
 * whole canonical model — notes, annotations, sync conflicts, duplicate detection — none of which a
 * citation in Word needs, and every module served to the pane is attack surface and load time.
 * `apps/word-taskpane/tsconfig.runtime.json` maps the specifier here, so the compiler refuses a
 * closure module that imports anything this file does not export.
 */
export { fromCslItem, toCslItem } from "./csl.js";
export { cloneJson, jsonEquals } from "./json.js";
export { DEFAULT_ITEM_TYPE_REGISTRY, processorCslType } from "./item-types.js";
export { DEFAULT_FIELD_REGISTRY } from "./field-registry.js";
export { validateReferenceItem } from "./validate.js";
export { generateUuidV7 } from "./identifiers.js";
// Identifier recognition in document mode (E10-20.4).
export { ARXIV_PATTERN, CLINICAL_TRIAL_PATTERN, DOI_PATTERN, PMCID_PATTERN } from "./scholarly-identifiers.js";
export { isValidIsbn, isbnComparisonKey, normalizeIsbn } from "./isbn.js";
