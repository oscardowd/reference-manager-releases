/**
 * Document mode: the pane's own Word session, for when the desktop cannot be reached (ADR-0242
 * decision 2).
 *
 * Word on iPad and on the web never has the desktop, and on a Mac or PC it may simply be closed.
 * §9.3 already puts a snapshot of every cited reference in the document so that its citations stay
 * editable then. This module composes the domain's own Word session handler — the same one the
 * desktop runs — over those snapshots and the pane's reference shelf, with styles from the pinned
 * catalog, and hands back a client with exactly the desktop client's shape. Every surface in the
 * pane keeps calling the same ports; only the transport differs.
 */

import { boundedFetch, createStyleSource } from "./document-styles.js";
import { createReferenceShelf } from "./reference-shelf.js";
import { createDocumentSessionClient } from "./word-session.js";

/** Session paths that render or plan with a style, and so need it installed first. */
const STYLED_PATHS = new Set([
  "/session",
  "/citation/preview",
  "/insertion/prepare",
  "/refresh/prepare",
  "/bibliography/prepare",
  "/edit/open",
  "/edit/prepare",
]);

/**
 * The retrieval transport a lookup runs over: the page's `fetch`, with the pane's deadline, and only
 * the headers a browser sends without a CORS preflight. The domain asks for a `User-Agent`, which a
 * browser either refuses to set or turns into a preflight some providers do not answer; the
 * browser's own is sent instead. No cookie and no referrer go with the request.
 */
export function lookupTransport(fetchImpl) {
  return {
    async request(request) {
      const accept = request.headers && typeof request.headers.Accept === "string" ? request.headers.Accept : "application/json";
      const response = await boundedFetch(fetchImpl, request.url, {
        headers: { Accept: accept },
        credentials: "omit",
        referrerPolicy: "no-referrer",
      });
      return { status: response.status, json: () => response.json(), text: () => response.text() };
    },
  };
}

async function fetchCatalog(fetchImpl) {
  const response = await boundedFetch(fetchImpl, new URL("./csl-catalog.json", import.meta.url).href, { cache: "no-cache" });
  if (!response.ok) throw new Error("The list of citation styles could not be loaded.");
  return response.json();
}

/**
 * Start document mode. `loadRuntime` resolves to the emitted runtime module (see
 * `document-runtime.js`); it is injected so tests can hand in the package itself.
 */
export async function startDocumentMode({
  loadRuntime,
  storage,
  fetchImpl,
  loadCatalog = () => fetchCatalog(fetchImpl),
  /** Replaces the network lookup; tests supply one. */
  importIdentifiers,
} = {}) {
  const [runtime, catalog] = await Promise.all([loadRuntime(), loadCatalog()]);
  const registry = new runtime.CslStyleRegistry();
  const styles = createStyleSource({ catalog, storage, fetchImpl });
  const shelf = await createReferenceShelf(storage);
  // The last document evidence the pane read. Every request that carries the document replaces it,
  // so a search after an edit lists what the document cites now.
  let evidence = { customXmlParts: [], controls: [] };

  const references = () => runtime.mergeDocumentReferences(runtime.documentReferences(evidence), shelf.list());
  // §9.2's "import a missing reference": the domain's own importer, over the document and the shelf.
  const lookup = importIdentifiers ?? runtime.createDocumentImporter({
    references,
    add: (reference) => { shelf.add(reference).catch(() => { /* storage fell back to memory */ }); },
    transport: lookupTransport(fetchImpl),
  });

  const session = runtime.createWordSession({
    ports: runtime.createDocumentBridgePorts({ references, styles: registry, importIdentifiers: lookup }),
    styles: registry,
  });

  async function send(payload) {
    const document = payload && payload.body && typeof payload.body === "object" ? payload.body.document : undefined;
    if (document && typeof document === "object") evidence = document;
    if (STYLED_PATHS.has(payload.path)) {
      try {
        await styles.prepare(registry, payload, runtime.documentRequirements(evidence));
      } catch (error) {
        return { status: 400, body: { error: error.code ?? "style-unavailable", message: error.message } };
      }
    }
    return session(payload);
  }

  return Object.freeze({
    client: createDocumentSessionClient(send),
    shelf,
    /** Look identifiers up and keep what is found on the shelf; answers the bridge's import result. */
    lookup: (identifiers) => lookup({ identifiers }),
    /** Keep a reference typed by hand, once the domain has validated it. */
    async addManual(input) {
      const outcome = runtime.manualReference(input);
      if (outcome.ok) await shelf.add(outcome.reference);
      return outcome;
    },
    /** Who a lookup contacts, for the sentence shown before it runs. */
    disclosures: () => runtime.documentModeDisclosures(),
    /** Every style this mode can offer, whether or not it is downloaded yet. */
    catalogStyles: () => styles.catalogStyles(),
    /** Record the document the pane just read, before any request carries it. */
    setDocument(next) {
      if (next && typeof next === "object") evidence = next;
    },
  });
}
