/**
 * Authenticated loopback client for the citation picker's four reads (SPEC §9.2; ADR-0008,
 * ADR-0240; E10-04.5).
 *
 * `/recent` is the picker's opening state, `/search` is what a typed term asks for, `/collections`
 * fills the scope selector and `POST /citation/preview` is the formatted result the user confirms
 * before anything is written. Transport only: selection policy is `@refmgr/word-addin`, and the
 * preview text is the CSL processor's answer, carried unaltered (invariant 4).
 *
 * **Every one of the four is a read.** None of them changes the library, so an abort here is
 * always a no-op on the user's data and is reported simply as nobody having answered — unlike the
 * collaborator client, which has a write to be honest about.
 *
 * The search term never enters a logger and never becomes a DOM attribute (§23, invariant 6).
 */

function connection(options) {
  const baseUrl = typeof options.baseUrl === "string" ? options.baseUrl : "";
  const token = typeof options.token === "string" ? options.token : "";
  let parsed;
  try {
    parsed = new URL(baseUrl);
  } catch {
    throw new Error("InvalidLibraryBridgeOrigin");
  }
  if (
    parsed.protocol !== "http:" ||
    parsed.hostname !== "127.0.0.1" ||
    parsed.pathname !== "/" ||
    parsed.username !== "" ||
    parsed.password !== "" ||
    parsed.search !== "" ||
    parsed.hash !== "" ||
    token.length === 0
  ) {
    throw new Error("InvalidLibraryBridgeOrigin");
  }
  return { baseUrl: parsed.origin, token };
}

/**
 * How long one library read may take, end to end, before the client stops waiting for it.
 *
 * The same reasoning and the same number as the annotation search (R-098): §21 budgets an ordinary
 * search under 200 ms, so a request still running at ten seconds has missed that budget fiftyfold
 * and cannot be told apart from a port that accepted the connection and will never answer. Larger
 * than the pane's 2 s `/health` probe, because these routes read a library rather than a status
 * line, and a first query over a cold index is allowed to be slow without being allowed to be
 * endless. The deadline covers the whole exchange, headers and body.
 */
export const LIBRARY_BRIDGE_TIMEOUT_MS = 10_000;

function optionalText(value) {
  if (value === undefined || value === null) return undefined;
  if (typeof value !== "string") throw new Error("InvalidLibraryBridgeHit");
  return value;
}

/**
 * One search hit, validated before it reaches the DOM.
 *
 * An unknown shape refuses rather than rendering a half-row: the next thing the user does with a
 * row is put the reference in a manuscript, and a row whose id did not survive the wire is a row
 * that inserts nothing or inserts the wrong thing.
 */
function presentHit(entry) {
  if (!entry || typeof entry !== "object") throw new Error("InvalidLibraryBridgeHit");
  const itemId = typeof entry.itemId === "string" ? entry.itemId : "";
  if (itemId.trim().length === 0) throw new Error("InvalidLibraryBridgeHit");
  return Object.freeze({
    itemId,
    itemType: typeof entry.itemType === "string" ? entry.itemType : "",
    title: typeof entry.title === "string" ? entry.title : "",
    creatorSummary: typeof entry.creatorSummary === "string" ? entry.creatorSummary : "",
    ...(optionalText(entry.year) === undefined ? {} : { year: entry.year }),
    ...(optionalText(entry.containerTitle) === undefined ? {} : { containerTitle: entry.containerTitle }),
    ...(optionalText(entry.doi) === undefined ? {} : { doi: entry.doi }),
  });
}

export function librarySearchPresentation(body) {
  if (!body || !Array.isArray(body.results)) throw new Error("InvalidLibraryBridgeResponse");
  return Object.freeze({ results: Object.freeze(body.results.map(presentHit)) });
}

export function libraryCollectionsPresentation(body) {
  if (!body || !Array.isArray(body.collections)) throw new Error("InvalidLibraryBridgeResponse");
  return Object.freeze({
    collections: Object.freeze(
      body.collections.map((entry) => {
        if (!entry || typeof entry !== "object") throw new Error("InvalidLibraryBridgeScope");
        const collectionId = typeof entry.collectionId === "string" ? entry.collectionId : "";
        if (collectionId.trim().length === 0) throw new Error("InvalidLibraryBridgeScope");
        return Object.freeze({
          collectionId,
          name: typeof entry.name === "string" ? entry.name : "",
          ...(optionalText(entry.parentId) === undefined ? {} : { parentId: entry.parentId }),
        });
      }),
    ),
  });
}

/**
 * The preview, as the CSL processor rendered it.
 *
 * `text` is taken verbatim — not trimmed, not wrapped, not punctuated. Anything this module did to
 * it would be the add-in formatting a citation, which invariant 4 forbids outright; the warnings
 * travel beside it because a preview that silently dropped "this style has no such locator" is a
 * preview the user cannot rely on.
 */
export function citationPreviewPresentation(body) {
  if (!body || typeof body.text !== "string") throw new Error("InvalidLibraryBridgeResponse");
  const warnings = Array.isArray(body.warnings) ? body.warnings : [];
  return Object.freeze({
    text: body.text,
    warnings: Object.freeze(
      warnings.map((warning) =>
        Object.freeze({
          code: typeof warning?.code === "string" ? warning.code : "unknown",
          message: typeof warning?.message === "string" ? warning.message : "",
        }),
      ),
    ),
    ...(body.context === undefined ? {} : { context: body.context }),
  });
}

/**
 * Build the client.
 *
 * `timeoutMs` exists so a test can state a deadline instead of waiting one out; the shipped task
 * pane passes nothing and gets {@link LIBRARY_BRIDGE_TIMEOUT_MS}. A present-but-unusable value is
 * refused rather than silently replaced by the default — a caller that wrote a deadline down meant
 * it, and quietly substituting another is the silent adjustment this pane refuses everywhere else.
 */
export function createLibraryBridgeClient(options) {
  const configured = connection(options || {});
  const fetchImpl = typeof options.fetch === "function" ? options.fetch : globalThis.fetch;
  if (typeof fetchImpl !== "function") throw new Error("LibraryBridgeFetchUnavailable");
  const timeoutMs = options.timeoutMs === undefined ? LIBRARY_BRIDGE_TIMEOUT_MS : options.timeoutMs;
  if (!Number.isSafeInteger(timeoutMs) || timeoutMs <= 0) {
    throw new Error("InvalidLibraryBridgeTimeout");
  }

  async function request(path, init, present) {
    const controller = new AbortController();
    // Through `globalThis` for the same reason `fetch` is: this module is loaded by a task pane
    // where the timers hang off `window`, and by tests where there is no `window` at all.
    const timer = globalThis.setTimeout(() => controller.abort(), timeoutMs);
    try {
      let response;
      try {
        response = await fetchImpl(`${configured.baseUrl}${path}`, {
          mode: "cors",
          credentials: "omit",
          cache: "no-store",
          referrerPolicy: "no-referrer",
          ...init,
          headers: {
            Accept: "application/json",
            Authorization: `Bearer ${configured.token}`,
            ...(init.headers ?? {}),
          },
          signal: controller.signal,
        });
      } catch {
        // Our own deadline, a refused connection and a browser that declined the request are
        // indistinguishable here and mean the same thing: nobody answered. **A deadline must land
        // here and nowhere else** — reporting it as `LibraryBridgeRefused` would tell the reader
        // the desktop application decided something it never got the chance to decide, and "we
        // could not ask" and "we asked and were refused" are different facts.
        throw new Error("LibraryBridgeUnreachable");
      }
      // A status the peer chose is the peer's answer, whatever it says — never a timeout.
      if (!response || response.ok !== true) throw new Error("LibraryBridgeRefused");
      let body;
      try {
        body = await response.json();
      } catch {
        // The body is where the two causes separate: an abort is still nobody answering, while
        // anything else is a peer that answered with something this client cannot read.
        throw new Error(
          controller.signal.aborted ? "LibraryBridgeUnreachable" : "InvalidLibraryBridgeResponse",
        );
      }
      return present(body);
    } finally {
      globalThis.clearTimeout(timer);
    }
  }

  return Object.freeze({
    /** §9.2 "Show recent references" — what the picker shows before a term exists. */
    async recent(query) {
      const limit = query && Number.isSafeInteger(query.limit) ? `?limit=${query.limit}` : "";
      return request(`/recent${limit}`, { method: "GET" }, librarySearchPresentation);
    },

    /**
     * §9.2's search, scoped and prioritised.
     *
     * `collection` is omitted rather than sent empty when the scope is the whole library: the
     * bridge refuses `collection=` written empty, precisely so that a scoped search cannot be
     * widened by an accident of string building.
     */
    async search(query) {
      const parts = [`q=${encodeURIComponent(String(query?.q ?? ""))}`];
      if (query && Number.isSafeInteger(query.limit)) parts.push(`limit=${query.limit}`);
      const collection = typeof query?.collection === "string" ? query.collection : "";
      if (collection.length > 0) parts.push(`collection=${encodeURIComponent(collection)}`);
      const prefer = Array.isArray(query?.prefer) ? query.prefer.filter((id) => typeof id === "string" && id.length > 0) : [];
      if (prefer.length > 0) parts.push(`prefer=${prefer.map((id) => encodeURIComponent(id)).join(",")}`);
      return request(`/search?${parts.join("&")}`, { method: "GET" }, librarySearchPresentation);
    },

    /** §9.2 "search … collection" — the scopes the selector may offer. */
    async collections() {
      return request("/collections", { method: "GET" }, libraryCollectionsPresentation);
    },

    /** §9.2 "Preview the formatted result before insertion". */
    async previewCitation(citationRequest) {
      return request(
        "/citation/preview",
        {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(citationRequest),
        },
        citationPreviewPresentation,
      );
    },
  });
}
