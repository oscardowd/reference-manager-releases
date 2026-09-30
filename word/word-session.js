/**
 * The pane's client for the Word session (ADR-0241, ADR-0242).
 *
 * Two transports, one client. With the desktop reachable, requests go to the HTTPS host's
 * same-origin `/word-api`, which holds the private desktop socket. Without it, they go to the pane's
 * own session (`document-mode.js`). Either way a request is `{ path, body?, query? }` and an answer is
 * `{ status, body }` (the HTTP transport adds `ok`), and the surfaces above cannot tell which carried
 * it — except through `mode`, which the pane shows.
 */
import { librarySearchPresentation, libraryCollectionsPresentation, citationPreviewPresentation } from "./library-bridge.js";

function clientOver(send, mode) {
  async function request(path, body, query) {
    const payload = { path, ...(body === undefined ? {} : { body }), ...(query === undefined ? {} : { query }) };
    const answer = await send(payload);
    const result = answer.body ?? {};
    const ok = typeof answer.ok === "boolean" ? answer.ok : answer.status >= 200 && answer.status < 300;
    if (!ok) {
      // The runtime's code travels with the sentence, so a caller can tell a question it must
      // put to the writer (409 `confirm-…`) from a refusal it can only report.
      const error = new Error(result.message ?? "Open Reference Manager, then choose Reconnect.");
      error.code = typeof result.error === "string" ? result.error : "unavailable";
      error.details = result;
      throw error;
    }
    return result;
  }
  return {
    mode,
    request,
    recent: async (query) => librarySearchPresentation(await request("/recent", undefined, { limit: String(query?.limit ?? 25) })),
    search: async (query) => librarySearchPresentation(await request("/search", undefined, {
      q: String(query.q), limit: String(query.limit ?? 25),
      ...(query.collection ? { collection: query.collection } : {}),
      ...(query.prefer?.length ? { prefer: query.prefer.join(",") } : {}),
    })),
    collections: async () => libraryCollectionsPresentation(await request("/collections")),
    previewCitation: async (body) => citationPreviewPresentation(await request("/citation/preview", body)),
  };
}

/**
 * The desktop, through the HTTPS host's `/word-api`. Same-origin when the host served this pane; when
 * the pane is served from a public origin (a store listing, ADR-0242 §7), `desktopOrigin` names the
 * host and the request is a CORS request that the host admits only for the one origin it was
 * configured with. Credentials are never sent either way.
 */
export function createWordSessionClient(fetchImpl = globalThis.fetch.bind(globalThis), { desktopOrigin } = {}) {
  const crossOrigin = typeof desktopOrigin === "string" && desktopOrigin !== globalThis.location?.origin;
  const endpoint = crossOrigin ? `${desktopOrigin}/word-api` : "/word-api";
  return clientOver(async (payload) => {
    const controller = new AbortController();
    const timeout = globalThis.setTimeout(() => controller.abort(), 15_000);
    try {
      const response = await fetchImpl(endpoint, {
        method: "POST", mode: crossOrigin ? "cors" : "same-origin", credentials: "omit", cache: "no-store",
        headers: { "Content-Type": "application/json", "X-RefMgr-Word": "1" },
        body: JSON.stringify(payload),
        signal: controller.signal,
      });
      return { ok: response.ok, status: response.status, body: await response.json() };
    } finally { globalThis.clearTimeout(timeout); }
  }, "library");
}

/**
 * The pane's own session, over the document's references (document mode). `send` is the session
 * handler, or a wrapper that prepares styles first; it answers `{ status, body }` and never throws
 * for a refusal. Anything it does throw is a defect, reported as one sentence rather than a stack.
 */
export function createDocumentSessionClient(send) {
  return clientOver(async (payload) => {
    try {
      return await send(payload);
    } catch {
      return { status: 500, body: { error: "document-mode-failed", message: "The pane could not complete that. Reopen it and try again." } };
    }
  }, "document");
}

/**
 * Evidence is re-read immediately before the first write; concurrent changes abort the plan.
 *
 * A numeric style, or an author-date one that now needs "2020a", cannot know a citation's final
 * text until it is in the document. The runtime says so (`refreshRequired`), the citation is
 * written with the text the writer previewed — payload and visible text agreeing — and the whole
 * document is then re-rendered in Word's real order. If that second step cannot run, the insertion
 * still stands and the writer is told how to finish it.
 */
export async function insertFromPicker({ client, readDocument, writeInsertion, writeRefresh, draft, request, previewText }) {
  const document = await readDocument();
  const prepared = await client.request("/insertion/prepare", { draft, styleId: request.styleId, locale: request.locale, previewText, document });
  const current = await readDocument();
  if (JSON.stringify(current) !== JSON.stringify(document)) throw new Error("The document changed during preparation. Review the preview and try again.");
  const result = await writeInsertion(prepared.plan, prepared.stalePartIds);
  if (result.status === "failed") throw new Error(result.cleanupRequiredPartIds.length ? "Word could not finish the insertion. Check the document before trying again." : "Word could not insert this citation. The draft has been kept.");
  if (result.status === "written-with-cleanup-required") {
    return { message: "Citation inserted. Extra citation data needs cleanup before another insertion." };
  }
  if (prepared.refreshRequired === true && typeof writeRefresh === "function") {
    try {
      await refreshDocumentFromPane({ client, readDocument, writeRefresh });
      return { message: "Citation inserted and the document's citations renumbered.", refreshed: true };
    } catch {
      return { message: "Citation inserted. Choose Update document on the Document tab to finish numbering." };
    }
  }
  return { message: "Citation inserted." };
}

/** Pull the confirmation answers a Document-tab action may carry, and nothing else. */
function confirmations(options) {
  const answers = {};
  if (Array.isArray(options?.acceptRemovedClusterIds)) answers.acceptRemovedClusterIds = options.acceptRemovedClusterIds;
  if (Array.isArray(options?.acceptManualReplacementIds)) answers.acceptManualReplacementIds = options.acceptManualReplacementIds;
  if (options?.acceptBibliographyReplacement === true) answers.acceptBibliographyReplacement = true;
  return answers;
}

/**
 * §9.1 Refresh document and Change citation style. One read, one plan, one re-read, one write:
 * the Office batch then re-checks the whole citation walk itself before it touches a control.
 */
export async function refreshDocumentFromPane({ client, readDocument, writeRefresh, styleId, ...options }) {
  const document = await readDocument();
  const prepared = await client.request("/refresh/prepare", {
    document,
    ...(typeof styleId === "string" && styleId.length > 0 ? { styleId } : {}),
    ...confirmations(options),
  });
  const current = await readDocument();
  if (JSON.stringify(current) !== JSON.stringify(document)) throw new Error("The document changed while it was being prepared. Try again.");
  const result = await writeRefresh(prepared.plan, prepared.stalePartIds);
  if (result.status === "failed") {
    throw new Error(result.cleanupRequiredPartIds?.length
      ? "Word could not finish the update. Check the document before trying again."
      : "Word could not update the document. Nothing was changed.");
  }
  return { summary: prepared.summary, cleanupRequired: result.status === "written-with-cleanup-required" };
}

/** §9.1 Insert bibliography: written at the cursor, or over the bibliography already there. */
export async function insertBibliographyFromPane({ client, readDocument, writeBibliography, ...options }) {
  const document = await readDocument();
  const prepared = await client.request("/bibliography/prepare", { document, ...confirmations(options) });
  const current = await readDocument();
  if (JSON.stringify(current) !== JSON.stringify(document)) throw new Error("The document changed while it was being prepared. Try again.");
  const result = await writeBibliography(prepared.plan, prepared.stalePartIds);
  if (result.status === "failed") {
    throw new Error(result.cleanupRequiredPartIds?.length
      ? "Word could not finish the bibliography. Check the document before trying again."
      : "Word could not write the bibliography. Nothing was changed.");
  }
  const entries = prepared.summary?.entries ?? 0;
  const noun = entries === 1 ? "1 entry" : `${String(entries)} entries`;
  return { message: prepared.summary?.replacing ? `Bibliography updated (${noun}).` : `Bibliography inserted (${noun}).` };
}

/**
 * §9.2's Edit citation, opened on the citation the user has selected in Word (E10-05.1).
 *
 * The pane sends evidence and gets back a projection; it never parses a payload or verifies a
 * fingerprint itself (ADR-0240, ADR-0241). `readSelected` is the one Office read that returns both
 * halves at once — the selected control as Word shows it, and every custom XML part — so the two
 * cannot be read at different instants and disagree.
 */
export async function openCitationEdit({ client, readSelected }) {
  const evidence = await readSelected();
  const session = await client.request("/edit/open", {
    document: { customXmlParts: evidence.customXmlParts },
    control: evidence.control,
  });
  return { ...session, control: evidence.control };
}

/**
 * Save an edited citation.
 *
 * The document is re-read **immediately before the write**, exactly as `insertFromPicker` does and
 * for the same reason: preparation is a round trip, and a citation that changed during it is a
 * citation the plan no longer describes. A mismatch aborts before anything is written rather than
 * overwriting whatever arrived in the meantime.
 */
export async function saveCitationEdit({
  client,
  readSelected,
  writeEdit,
  readDocument,
  writeRefresh,
  draft,
  previewText,
  acceptManualEditReplacement,
}) {
  const before = await readSelected();
  const prepared = await client.request("/edit/prepare", {
    document: { customXmlParts: before.customXmlParts },
    control: before.control,
    draft,
    previewText,
    ...(acceptManualEditReplacement === true ? { acceptManualEditReplacement: true } : {}),
  });
  const current = await readSelected();
  if (JSON.stringify(current) !== JSON.stringify(before)) {
    throw new Error("The citation changed while it was being prepared. Review the preview and save again.");
  }
  const result = await writeEdit(prepared.expectedTag, prepared);
  if (result.status === "failed") {
    throw new Error(
      result.cleanupRequiredPartIds.length
        ? "Word could not finish the change. Check the citation in the document before trying again."
        : "Word could not change this citation. Your edits have been kept.",
    );
  }
  if (result.status === "written-with-cleanup-required") {
    return { message: "Citation updated. Extra citation data needs cleanup before another change." };
  }
  // A numeric style: the edit was written with its single-cite text; the document's order decides
  // the number, so the whole document is re-rendered straight after — as after an insertion.
  if (prepared.refreshRequired === true && typeof readDocument === "function" && typeof writeRefresh === "function") {
    try {
      await refreshDocumentFromPane({ client, readDocument, writeRefresh });
      return { message: "Citation updated and the document's citations renumbered.", refreshed: true };
    } catch {
      return { message: "Citation updated. Choose Update document on the Document tab to finish numbering." };
    }
  }
  return { message: "Citation updated." };
}
