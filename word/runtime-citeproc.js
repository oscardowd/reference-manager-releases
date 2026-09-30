/**
 * citeproc-js as an ES module, for the task pane's citation runtime (ADR-0242 decision 3).
 *
 * citeproc-js publishes only CommonJS and is never modified (ADR-0013), so a browser cannot import
 * it directly. `document-runtime.js` loads its file byte-for-byte as a classic script with a
 * `module` object in place and keeps what it exported; the pane's import map sends the specifier
 * `citeproc` here, and this module hands that export to the runtime. It must be loaded first — the
 * runtime is imported only after — so an absent export is a loading-order defect, and it says so.
 */
const CSL = globalThis.refmgrCiteproc;
if (CSL === undefined || typeof CSL.Engine !== "function") {
  throw new Error("citeproc-js was not loaded before the citation runtime");
}
export default CSL;
