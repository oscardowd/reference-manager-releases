/**
 * The bridge's route table (ADR-0008), on its own so it has no runtime dependency.
 *
 * `policy.ts` decides whether a request may reach a route and needs `node:crypto` to do it; the Word
 * session only needs to *name* a route, and it runs in the task pane as well as in the desktop
 * (ADR-0242). Keeping the table here lets the pane import it without the gate.
 */
/** The routes ADR-0008 permits. The listener refuses anything not on this list. */
export const BRIDGE_ROUTES = [
    { method: "GET", path: "/health", requiresToken: false },
    { method: "GET", path: "/search", requiresToken: true },
    { method: "GET", path: "/annotations/search", requiresToken: true },
    { method: "GET", path: "/recent", requiresToken: true },
    // §9.2 "Search … collection" needs the picker to be able to *name* the collections before it
    // can scope a search to one. `/search?collection=` has always accepted an id; nothing told the
    // add-in what ids exist, so the scope selector had nothing to offer (E10-04.5).
    { method: "GET", path: "/collections", requiresToken: true },
    { method: "GET", path: "/items/:id", requiresToken: true },
    { method: "POST", path: "/citation/preview", requiresToken: true },
    { method: "POST", path: "/citation/format", requiresToken: true },
    { method: "POST", path: "/bibliography", requiresToken: true },
    { method: "GET", path: "/styles", requiresToken: true },
    { method: "POST", path: "/import", requiresToken: true },
    { method: "POST", path: "/collaboration/references", requiresToken: true },
    // §5.5's browser capture (E11-02.2, ADR-0128). Token-bearing like every other route: the
    // extension is a caller, not a peer, and an unauthenticated capture route would let any process
    // that can open a loopback socket push a reference into the preview a user is about to accept.
    { method: "POST", path: "/capture", requiresToken: true },
    { method: "GET", path: "/capture/options", requiresToken: true },
    { method: "POST", path: "/capture/duplicates", requiresToken: true },
    { method: "POST", path: "/capture/save", requiresToken: true },
    // §9.5's pre-repair copy of the user's document (E10-11.1, ADR-0116). Four paths rather than
    // one, because a 64 MiB document arrives in slices and the size must be declared before any of
    // them is sent. `read` is a separate path from `slice` rather than the same path with a
    // different method: `resolveRoute` matches a path and *then* checks the method, so one path
    // cannot carry two verbs.
    { method: "POST", path: "/document-backup/begin", requiresToken: true },
    { method: "POST", path: "/document-backup/slice", requiresToken: true },
    { method: "POST", path: "/document-backup/commit", requiresToken: true },
    { method: "GET", path: "/document-backup/read", requiresToken: true },
];
