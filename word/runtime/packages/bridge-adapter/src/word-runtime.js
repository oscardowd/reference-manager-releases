/**
 * The Word task pane's citation runtime (SPEC §9.1–§9.4; ADR-0242 decision 3).
 *
 * The one entry `apps/word-taskpane/tsconfig.runtime.json` compiles for the browser. Everything the
 * pane can run of the domain is reachable from here and nothing else is served: the Word session
 * handler, the ports that back it with a document's own snapshots, and the style registry those
 * ports render with. The desktop composes the same handler over the library
 * (`word-session-library.ts`), so the two modes cannot drift apart in what they write.
 */
export { createWordSession } from "./word-session.js";
export { createDocumentBridgePorts, documentReferences, documentRequirements, mergeDocumentReferences, } from "./document-ports.js";
export { DOCUMENT_LIBRARY_ID, createDocumentImporter, documentModeDisclosures, manualReference, } from "./document-import.js";
export { CslStyleRegistry } from "@refmgr/citation";
