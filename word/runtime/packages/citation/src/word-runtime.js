/**
 * The `@refmgr/citation` surface the Word task pane's runtime reaches (ADR-0242 decision 3).
 *
 * Narrower than `desktop-runtime.ts`: the pane renders and plans citations in Word and nothing else,
 * so note rendering and the CSL test-suite runner are not served to it.
 */
export { CslStyleRegistry, bibliographyLayout, citationSortsItems } from "./styles.js";
export { CiteprocJsAdapter } from "./adapters/citeproc-js/index.js";
export { CslBibliographyRenderer } from "./bibliography.js";
export { CitationClusterRenderer, toProcessorItem } from "./clusters.js";
export { CitationDocumentEditor } from "./document-edits.js";
