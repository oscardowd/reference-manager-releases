/**
 * Load the task pane's own citation runtime (ADR-0242 decisions 2 and 3).
 *
 * The runtime is the domain's Word session handler compiled to browser modules. It is loaded only
 * when the pane needs it — when the desktop cannot be reached — so the desktop path pays nothing
 * for it, and in particular not for citeproc-js, which is most of its weight.
 *
 * citeproc-js is CommonJS and is never modified (ADR-0013). It is loaded here as a classic script,
 * from the verbatim copy the build serves, with a `module` object in place for its last line
 * (`module.exports = CSL`) to assign to. The object is removed straight after, so no other script
 * mistakes the page for a CommonJS environment. `runtime-citeproc.js` then hands the export to the
 * runtime under the specifier `citeproc`.
 */

const CITEPROC_URL = "./vendor/citeproc/citeproc_commonjs.js";
const RUNTIME_URL = "./runtime/packages/bridge-adapter/src/word-runtime.js";

function appendScript(url) {
  return new Promise((resolve, reject) => {
    const script = globalThis.document.createElement("script");
    script.src = url;
    script.async = false;
    script.onload = () => resolve();
    script.onerror = () => reject(new Error("The citation engine could not be loaded."));
    globalThis.document.head.append(script);
  });
}

let loading = null;

/** Resolve to the runtime module; loads once per page, and a failure may be retried. */
export function loadDocumentRuntime({ loadScript = appendScript, importModule = (url) => import(url) } = {}) {
  if (loading === null) {
    loading = (async () => {
      if (globalThis.refmgrCiteproc === undefined) {
        const holder = { exports: {} };
        globalThis.module = holder;
        try {
          await loadScript(CITEPROC_URL);
        } finally {
          delete globalThis.module;
        }
        if (typeof holder.exports?.Engine !== "function") {
          throw new Error("The citation engine did not load correctly.");
        }
        globalThis.refmgrCiteproc = holder.exports;
      }
      return importModule(new URL(RUNTIME_URL, import.meta.url).href);
    })().catch((error) => {
      loading = null;
      throw error;
    });
  }
  return loading;
}
