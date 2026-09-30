/**
 * Authenticated loopback client for GET /annotations/search (SPEC §9.1; ADR-0173/0176; E14-04.4).
 *
 * Transport only. Selection policy and citation formatting stay in `@refmgr/word-addin`. Query,
 * quote and comment never enter a logger (§23).
 *
 * The request is bounded (R-098). A refused connection fails in microseconds, so the state this
 * guards against is not an absent desktop application but a loopback port **answered and then
 * held** — another program listening on it, a wedged desktop process, a proxy that accepts and
 * stalls. An unbounded `fetch` waits on that for as long as the browser will, and the picker
 * waits with it.
 */

function connection(options) {
  const baseUrl = typeof options.baseUrl === "string" ? options.baseUrl : "";
  const token = typeof options.token === "string" ? options.token : "";
  let parsed;
  try {
    parsed = new URL(baseUrl);
  } catch {
    throw new Error("InvalidAnnotationBridgeOrigin");
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
    throw new Error("InvalidAnnotationBridgeOrigin");
  }
  return { baseUrl: parsed.origin, token };
}

function optionalText(value) {
  if (value === undefined || value === null) return undefined;
  if (typeof value !== "string") throw new Error("InvalidAnnotationBridgeHit");
  return value;
}

function presentHit(entry) {
  if (!entry || typeof entry !== "object") throw new Error("InvalidAnnotationBridgeHit");
  const annotationId = typeof entry.annotationId === "string" ? entry.annotationId : "";
  const attachmentId = typeof entry.attachmentId === "string" ? entry.attachmentId : "";
  const type = typeof entry.type === "string" ? entry.type : "";
  if (annotationId.length === 0 || attachmentId.length === 0 || type.length === 0) {
    throw new Error("InvalidAnnotationBridgeHit");
  }
  if (!Number.isSafeInteger(entry.pageIndex) || entry.pageIndex < 0) {
    throw new Error("InvalidAnnotationBridgeHit");
  }
  let reference = null;
  if (entry.reference !== null && entry.reference !== undefined) {
    if (typeof entry.reference !== "object" || typeof entry.reference.itemId !== "string") {
      throw new Error("InvalidAnnotationBridgeHit");
    }
    const itemId = entry.reference.itemId;
    if (itemId.trim().length === 0) throw new Error("InvalidAnnotationBridgeHit");
    reference = Object.freeze({
      itemId,
      title: typeof entry.reference.title === "string" ? entry.reference.title : "",
      creatorSummary:
        typeof entry.reference.creatorSummary === "string" ? entry.reference.creatorSummary : "",
      ...(typeof entry.reference.year === "string" ? { year: entry.reference.year } : {}),
    });
  }
  return Object.freeze({
    annotationId,
    attachmentId,
    type,
    pageIndex: entry.pageIndex,
    ...(optionalText(entry.pageLabel) === undefined ? {} : { pageLabel: entry.pageLabel }),
    ...(optionalText(entry.quoteText) === undefined ? {} : { quoteText: entry.quoteText }),
    ...(optionalText(entry.commentText) === undefined ? {} : { commentText: entry.commentText }),
    ...(optionalText(entry.color) === undefined ? {} : { color: entry.color }),
    reference,
  });
}

/**
 * Validate the bridge page before it reaches the DOM. Unknown shapes refuse rather than render
 * a half-hit the user could select by mistake.
 */
export function annotationSearchPresentation(page) {
  if (!page || !Array.isArray(page.results) || typeof page.total !== "number") {
    throw new Error("InvalidAnnotationBridgeResponse");
  }
  return Object.freeze({
    results: Object.freeze(page.results.map(presentHit)),
    total: page.total,
    totalExact: page.totalExact === true,
    scanned: typeof page.scanned === "number" ? page.scanned : 0,
  });
}

/**
 * How long one search may take, end to end, before the client stops waiting for it.
 *
 * §21 budgets an ordinary search at under 200 ms, so a request still running at ten seconds has
 * already missed that budget fiftyfold; the number is not a performance assertion but the point
 * past which continuing to wait cannot be told apart from a port that will never answer. It is
 * larger than the task pane's 2 s `/health` probe because this route reads a library rather than
 * a status line — a first query over a cold index on a large library is allowed to be slow
 * without being allowed to be endless.
 *
 * The deadline covers the **whole exchange**, headers and body: a peer that sends a status line
 * and then holds the body open would otherwise hang exactly where an unbounded request did.
 */
export const ANNOTATION_BRIDGE_TIMEOUT_MS = 10_000;

/**
 * Build the client.
 *
 * `timeoutMs` exists so a test can state a deadline instead of waiting one out; the shipped task
 * pane passes nothing and gets {@link ANNOTATION_BRIDGE_TIMEOUT_MS}. A present-but-unusable value
 * is refused rather than silently replaced by the default — a caller that wrote a deadline down
 * meant it, and quietly substituting another is the silent-adjustment failure this module refuses
 * everywhere else.
 */
export function createAnnotationBridgeClient(options) {
  const configured = connection(options || {});
  const fetchImpl = typeof options.fetch === "function" ? options.fetch : globalThis.fetch;
  if (typeof fetchImpl !== "function") throw new Error("AnnotationBridgeFetchUnavailable");
  const timeoutMs = options.timeoutMs === undefined ? ANNOTATION_BRIDGE_TIMEOUT_MS : options.timeoutMs;
  if (!Number.isSafeInteger(timeoutMs) || timeoutMs <= 0) {
    throw new Error("InvalidAnnotationBridgeTimeout");
  }

  return Object.freeze({
    async search(query) {
      const parts = [];
      const text = query && typeof query.q === "string" ? query.q : undefined;
      if (text !== undefined) parts.push(`q=${encodeURIComponent(text)}`);
      if (query && Number.isSafeInteger(query.limit)) parts.push(`limit=${query.limit}`);
      if (query && Number.isSafeInteger(query.offset)) parts.push(`offset=${query.offset}`);
      const suffix = parts.length === 0 ? "" : `?${parts.join("&")}`;
      const controller = new AbortController();
      // Through `globalThis` for the same reason `fetch` is: this module is loaded by a task pane
      // where the timers hang off `window`, and by tests where there is no `window` at all.
      const timer = globalThis.setTimeout(() => controller.abort(), timeoutMs);
      try {
        let response;
        try {
          response = await fetchImpl(`${configured.baseUrl}/annotations/search${suffix}`, {
            method: "GET",
            mode: "cors",
            credentials: "omit",
            cache: "no-store",
            referrerPolicy: "no-referrer",
            headers: {
              Accept: "application/json",
              Authorization: `Bearer ${configured.token}`,
            },
            signal: controller.signal,
          });
        } catch {
          // Our own deadline, a refused connection and a browser that declined the request are
          // indistinguishable here and mean the same thing: nobody answered. **A deadline must
          // land here and nowhere else** — reporting it as `AnnotationBridgeRefused` would tell
          // the reader the desktop application decided something it never got the chance to
          // decide, and "we could not ask" and "we asked and were refused" are different facts.
          throw new Error("AnnotationBridgeUnreachable");
        }
        // A status the peer chose is the peer's answer, whatever it says — never a timeout.
        if (!response || response.ok !== true) throw new Error("AnnotationBridgeRefused");
        let body;
        try {
          body = await response.json();
        } catch {
          // The body is where the two causes separate: an abort is still nobody answering, while
          // anything else is a peer that answered with something this client cannot read.
          throw new Error(
            controller.signal.aborted ? "AnnotationBridgeUnreachable" : "InvalidAnnotationBridgeResponse",
          );
        }
        return annotationSearchPresentation(body);
      } finally {
        globalThis.clearTimeout(timer);
      }
    },
  });
}
