/**
 * The `@refmgr/identifiers` surface the Word task pane's runtime reaches (ADR-0242 decision 5).
 *
 * Recognition, retrieval over an injected transport, reconciliation and the provider disclosures —
 * not the package index, which also exports the open-access full-text fetcher and its `node:crypto`
 * hashing, neither of which a task pane uses.
 */
export { recognizeIdentifiers } from "./recognize.js";
export { createMetadataSources, retrieveMetadata, } from "./retrieval.js";
export { reconcileBatch, resolveReconciliation, } from "./reconcile.js";
export { RETRIEVAL_DISCLOSURES } from "./retrieval-policy.js";
