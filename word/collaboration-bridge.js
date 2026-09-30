/**
 * Authenticated loopback client for E10-13.2's collaborator-review route (SPEC §9.6; ADR-0120).
 *
 * The request is bounded (R-098). A refused connection fails in microseconds, so the state this
 * guards against is not an absent desktop application but a loopback port **answered and then
 * held** — another program listening on it, a wedged desktop process, a proxy that accepts and
 * stalls. An unbounded `fetch` waits on that for as long as the browser will, and the review
 * surface waits with it.
 *
 * **This route is a write, and that changes what an abort means.** `/collaboration/references`
 * inserts a reference for every id in `importItemIds`, so a deadline ends *our wait*, never the
 * server's transaction: past the deadline the import may have happened, may be part-done, or may
 * never have started, and this client does not know which. §21 makes an interrupted import the
 * server's problem to resume or roll back safely; it makes it this client's problem never to
 * report the interruption as though nothing happened. So an abort on a request that asked for
 * imports is reported with {@link CollaboratorBridgeUnreachable.outcomeUnknown} set, and the
 * surface above must offer a re-read rather than telling the user their references were not
 * imported. A review that asked for no imports writes nothing, and there the abort really is a
 * no-op — the two are different facts and are kept different.
 */

function connection(options) {
  const baseUrl = typeof options.baseUrl === "string" ? options.baseUrl : "";
  const token = typeof options.token === "string" ? options.token : "";
  let parsed;
  try { parsed = new URL(baseUrl); } catch { throw new Error("InvalidCollaboratorBridgeOrigin"); }
  if (parsed.protocol !== "http:" || parsed.hostname !== "127.0.0.1" || parsed.pathname !== "/" ||
      parsed.username !== "" || parsed.password !== "" || parsed.search !== "" || parsed.hash !== "" ||
      token.length === 0) {
    throw new Error("InvalidCollaboratorBridgeOrigin");
  }
  return { baseUrl: parsed.origin, token };
}

/**
 * How long one collaborator review may take, end to end, before the client stops waiting for it.
 *
 * Sized from what the request does rather than from the number next door. The annotation search
 * is one indexed query on one term and is bounded at 10 s; this posts **every** reference in the
 * document and the server projects each one, compares it pairwise against the library, diffs the
 * matches and inserts the chosen ones. §21's shape for that is a 1,000-citation document against
 * a 100,000-reference library, over SQLite, on a cold page cache — work that is allowed to be
 * slow without being allowed to be endless. Thirty seconds is the point past which continuing to
 * wait cannot be told apart from a port that will never answer.
 *
 * The deadline covers the **whole exchange**, headers and body: a peer that sends a status line
 * and then holds the body open would otherwise hang exactly where an unbounded request did.
 */
export const COLLABORATOR_BRIDGE_TIMEOUT_MS = 30_000;

/**
 * The error a deadline, a refused connection or a declined request produces.
 *
 * `outcomeUnknown` is the write half of the classification: true when the request that did not
 * come back had asked for imports, because the server may have applied some, all or none of them.
 * A caller that treats this as "nothing was imported" would tell the user their library is
 * unchanged when it may not be.
 */
export class CollaboratorBridgeUnreachable extends Error {
  constructor(outcomeUnknown) {
    super("CollaboratorBridgeUnreachable");
    this.name = "CollaboratorBridgeUnreachable";
    this.outcomeUnknown = outcomeUnknown === true;
  }
}

function asksForImports(request) {
  return Array.isArray(request?.importItemIds) && request.importItemIds.length > 0;
}

/**
 * Build the client.
 *
 * `timeoutMs` exists so a test can state a deadline instead of waiting one out; the shipped task
 * pane passes nothing and gets {@link COLLABORATOR_BRIDGE_TIMEOUT_MS}. A present-but-unusable
 * value is refused rather than silently replaced by the default — a caller that wrote a deadline
 * down meant it.
 */
export function createCollaboratorBridgeClient(options) {
  const configured = connection(options || {});
  const fetchImpl = typeof options.fetch === "function" ? options.fetch : globalThis.fetch;
  if (typeof fetchImpl !== "function") throw new Error("CollaboratorBridgeFetchUnavailable");
  const timeoutMs =
    options.timeoutMs === undefined ? COLLABORATOR_BRIDGE_TIMEOUT_MS : options.timeoutMs;
  if (!Number.isSafeInteger(timeoutMs) || timeoutMs <= 0) {
    throw new Error("InvalidCollaboratorBridgeTimeout");
  }

  return Object.freeze({
    async review(request) {
      const writing = asksForImports(request);
      const controller = new AbortController();
      // Through `globalThis` for the same reason `fetch` is: this module is loaded by a task pane
      // where the timers hang off `window`, and by tests where there is no `window` at all.
      const timer = globalThis.setTimeout(() => controller.abort(), timeoutMs);
      try {
        let response;
        try {
          response = await fetchImpl(`${configured.baseUrl}/collaboration/references`, {
            method: "POST",
            headers: {
              Authorization: `Bearer ${configured.token}`,
              "Content-Type": "application/json",
            },
            body: JSON.stringify(request),
            signal: controller.signal,
          });
        } catch {
          // Our own deadline, a refused connection and a browser that declined the request are
          // indistinguishable here and mean the same thing: nobody answered. **A deadline must
          // land here and nowhere else** — reporting it as `CollaboratorBridgeRefused` would tell
          // the reader the desktop application decided something it never got the chance to
          // decide, and "we could not ask" and "we asked and were refused" are different facts.
          throw new CollaboratorBridgeUnreachable(writing);
        }
        // A status the peer chose is the peer's answer, whatever it says — never a timeout.
        if (!response || response.ok !== true) throw new Error("CollaboratorBridgeRefused");
        let result;
        try {
          result = await response.json();
        } catch {
          // The body is where the two causes separate: an abort is still nobody answering, while
          // anything else is a peer that answered with something this client cannot read. A body
          // that never arrived leaves a write undecided; a body we could not parse does not —
          // the server finished, and its answer is simply unreadable here.
          if (controller.signal.aborted) throw new CollaboratorBridgeUnreachable(writing);
          throw new Error("InvalidCollaboratorBridgeResponse");
        }
        if (!result || !result.library || !Array.isArray(result.comparisons)) {
          throw new Error("InvalidCollaboratorBridgeResponse");
        }
        return result;
      } finally {
        globalThis.clearTimeout(timer);
      }
    },
  });
}
