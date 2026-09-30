/**
 * The bridge's read surface — SPEC §9.2, ADR-0008, ADR-0100.
 *
 * Domain routes over the ports in `ports.ts`. The listener (E10-02.1) has already decided the
 * caller may be here; this file decides what they may *ask for*, and the two jobs are separate
 * because they fail differently — the gate's failure is an intruder, and this file's failure is
 * a legitimate caller being handed something it should not see, or a malformed request reaching
 * a repository.
 *
 * Three rules run through all of it:
 *
 * - **A request is validated before a port sees it, and validation is strict.** An unknown field
 *   is a refusal, not something to ignore. This is an API between two halves of one product that
 *   ship together; a silently-ignored misspelling is an integration bug that surfaces as "the
 *   locator did nothing" months later, in a user's manuscript.
 * - **A bound is refused, never clamped.** `limit=100000` gets a 400 saying the maximum, not 25
 *   results and the impression that the library holds 25 papers. Same reasoning as the batch-edit
 *   surface (ADR-0096): a silently-adjusted request is a request the caller believes was honoured.
 * - **Nothing that leaves here is composed by this package.** Citation text comes from the
 *   processor through the port (invariant 4), and no response carries a filesystem path (ADR-0008)
 *   because the ports have nowhere to put one.
 */
import { BridgeRequestError, badRequest } from "./errors.js";
/** §9.2's picker shows a page of results, not a library. */
export const MAX_SEARCH_LIMIT = 100;
export const DEFAULT_SEARCH_LIMIT = 25;
/** One paste of identifiers. Past this the user wants the desktop importer, not a task pane. */
export const MAX_IMPORT_IDENTIFIERS = 200;
/** A bounded manuscript-sized review request; larger documents are handled in pages. */
export const MAX_COLLABORATOR_SNAPSHOTS = 500;
/** A grouped citation. Beyond this something has gone wrong upstream. */
export const MAX_CITATION_ITEMS = 50;
/** A document's bibliography, bounded so one request cannot render a whole library. */
export const MAX_BIBLIOGRAPHY_ITEMS = 5000;
export const MAX_CAPTURE_COLLECTIONS = 32;
export const MAX_CAPTURE_TAGS = 64;
export const MAX_CAPTURE_NOTE_LENGTH = 4_000;
/**
 * What a route can change on the user's machine.
 *
 * Declared, not inferred from the verb or the method — the rule E06-01.1 arrived at for IPC
 * channels. A test asserts the `write` set is exactly `/import`, so a second route that touches
 * the library cannot be added without the addition being visible.
 *
 * `store` is the third value, added by E10-11.1: those routes never touch the library, but they do
 * write a copy of the user's **document** to disk and delete older copies under a retention policy
 * (ADR-0116). Folding them into `write` would have made the library-write assertion vaguer, and
 * calling them `read` would have been false.
 */
export const BRIDGE_ROUTE_EFFECTS = {
    "/health": "read",
    "/search": "read",
    "/annotations/search": "read",
    "/recent": "read",
    "/collections": "read",
    "/items/:id": "read",
    "/citation/preview": "read",
    "/citation/format": "read",
    "/bibliography": "read",
    "/styles": "read",
    "/import": "write",
    "/collaboration/references": "write",
    // `read` because it changes nothing: it neither touches the library nor writes a file. It also
    // does not *read* the library, which no band expresses — the three bands are about what a route
    // can change, and this one changes nothing. E11-03's save route will be `write`, added beside
    // this one rather than by widening it, so the assertion below keeps naming every writer.
    "/capture": "read",
    "/capture/options": "read",
    "/capture/duplicates": "read",
    "/capture/save": "write",
    "/document-backup/begin": "store",
    "/document-backup/slice": "store",
    "/document-backup/commit": "store",
    "/document-backup/read": "store",
};
function isRecord(value) {
    return typeof value === "object" && value !== null && !Array.isArray(value);
}
/**
 * Parse a JSON body and refuse any field the route does not know.
 *
 * The allowed-key list is passed in rather than derived, so adding a field to a request means
 * editing this call — the same "make it a visible act" reasoning as the IPC allow-list.
 */
function readJsonObject(body, allowed) {
    if (body.trim().length === 0) {
        badRequest("body-required", "this route requires a JSON object body");
    }
    let parsed;
    try {
        parsed = JSON.parse(body);
    }
    catch {
        // The parser's message quotes the body, and the body may hold a search term or a title.
        badRequest("body-not-json", "request body is not valid JSON");
    }
    if (!isRecord(parsed)) {
        badRequest("body-not-an-object", "request body must be a JSON object");
    }
    for (const key of Object.keys(parsed)) {
        if (!allowed.includes(key)) {
            // The key is named because a caller's own field name is not user content, and a refusal
            // that will not say which field is a refusal nobody can act on.
            badRequest("unknown-field", `unknown request field "${key}"`);
        }
    }
    return parsed;
}
function requireString(source, key) {
    const value = source[key];
    if (typeof value !== "string" || value.length === 0) {
        badRequest("invalid-field", `"${key}" must be a non-empty string`);
    }
    return value;
}
function optionalString(source, key) {
    const value = source[key];
    if (value === undefined) {
        return undefined;
    }
    if (typeof value !== "string") {
        badRequest("invalid-field", `"${key}" must be a string when present`);
    }
    return value;
}
function optionalBoolean(source, key) {
    const value = source[key];
    if (value === undefined) {
        return undefined;
    }
    if (typeof value !== "boolean") {
        badRequest("invalid-field", `"${key}" must be a boolean when present`);
    }
    return value;
}
function requireStringArray(source, key, max) {
    const value = source[key];
    if (!Array.isArray(value)) {
        badRequest("invalid-field", `"${key}" must be an array of strings`);
    }
    if (value.length === 0) {
        badRequest("invalid-field", `"${key}" must not be empty`);
    }
    if (value.length > max) {
        badRequest("too-many", `"${key}" accepts at most ${max} entries`);
    }
    for (const entry of value) {
        if (typeof entry !== "string" || entry.length === 0) {
            badRequest("invalid-field", `every entry in "${key}" must be a non-empty string`);
        }
    }
    return value;
}
function parseLimit(query) {
    const raw = query["limit"];
    if (raw === undefined) {
        return DEFAULT_SEARCH_LIMIT;
    }
    if (!/^\d+$/.test(raw)) {
        badRequest("invalid-query", '"limit" must be a whole number');
    }
    const limit = Number(raw);
    if (limit < 1 || limit > MAX_SEARCH_LIMIT) {
        // Refused, not clamped. See the module note.
        badRequest("invalid-query", `"limit" must be between 1 and ${MAX_SEARCH_LIMIT}`);
    }
    return limit;
}
function parseOffset(query) {
    const raw = query["offset"];
    if (raw === undefined)
        return 0;
    if (!/^\d+$/.test(raw) || !Number.isSafeInteger(Number(raw))) {
        badRequest("invalid-query", '"offset" must be a non-negative safe integer');
    }
    return Number(raw);
}
function parseCitationItems(source) {
    const raw = source["items"];
    if (!Array.isArray(raw) || raw.length === 0) {
        badRequest("invalid-field", '"items" must be a non-empty array');
    }
    if (raw.length > MAX_CITATION_ITEMS) {
        badRequest("too-many", `"items" accepts at most ${MAX_CITATION_ITEMS} entries`);
    }
    return raw.map((entry) => {
        if (!isRecord(entry)) {
            badRequest("invalid-field", 'every entry in "items" must be an object');
        }
        for (const key of Object.keys(entry)) {
            if (!["itemId", "locator", "prefix", "suffix", "suppressAuthor", "authorOnly"].includes(key)) {
                badRequest("unknown-field", `unknown field "${key}" in a citation item`);
            }
        }
        const item = {
            itemId: requireString(entry, "itemId"),
            locator: parseLocator(entry),
            prefix: optionalString(entry, "prefix"),
            suffix: optionalString(entry, "suffix"),
            suppressAuthor: optionalBoolean(entry, "suppressAuthor"),
            authorOnly: optionalBoolean(entry, "authorOnly"),
        };
        // The same refusal `packages/citation` makes (ADR-0059): citeproc renders the pair as
        // `[NO_PRINTED_FORM]`, so it must not reach the processor.
        if (item.suppressAuthor === true && item.authorOnly === true) {
            badRequest("conflicting-options", '"suppressAuthor" and "authorOnly" are mutually exclusive');
        }
        return item;
    });
}
function parseLocator(entry) {
    const locator = entry["locator"];
    if (locator === undefined) {
        return undefined;
    }
    if (!isRecord(locator)) {
        badRequest("invalid-field", '"locator" must be an object with label and value');
    }
    for (const key of Object.keys(locator)) {
        if (!["label", "value"].includes(key)) {
            badRequest("unknown-field", `unknown field "${key}" in a locator`);
        }
    }
    // The label is not checked against CSL's vocabulary here: `packages/citation` warns rather than
    // refuses on an unknown label (ADR-0059), and a second, stricter opinion at the seam would make
    // the bridge reject documents the desktop app accepts.
    return { label: requireString(locator, "label"), value: requireString(locator, "value") };
}
function parseCitationRequest(body) {
    const source = readJsonObject(body, ["styleId", "locale", "items"]);
    return {
        styleId: requireString(source, "styleId"),
        locale: requireString(source, "locale"),
        items: parseCitationItems(source),
    };
}
function parseBibliographyRequest(body) {
    const source = readJsonObject(body, ["styleId", "locale", "itemIds"]);
    return {
        styleId: requireString(source, "styleId"),
        locale: requireString(source, "locale"),
        itemIds: requireStringArray(source, "itemIds", MAX_BIBLIOGRAPHY_ITEMS),
    };
}
function parseImportRequest(body) {
    const source = readJsonObject(body, ["identifiers"]);
    return { identifiers: requireStringArray(source, "identifiers", MAX_IMPORT_IDENTIFIERS) };
}
/**
 * The capture envelope, and only the envelope (ADR-0128).
 *
 * One field, which must be a JSON object. What is *inside* it is ADR-0127's snapshot contract, and
 * that contract lives in `connector`, which this package may not import — so validating it here
 * would mean maintaining a second copy of it, in the layer least able to keep it in step. The
 * port bounds and refuses the snapshot itself, and the route turns its refusal into a 400 with
 * the port's own code. The strictness that belongs here — an unknown field is a refusal — is
 * applied exactly as it is for every other route.
 */
function parseCaptureSnapshot(body) {
    const source = readJsonObject(body, ["snapshot"]);
    const snapshot = source["snapshot"];
    if (typeof snapshot !== "object" || snapshot === null || Array.isArray(snapshot)) {
        badRequest("invalid-field", '"snapshot" must be a JSON object');
    }
    return snapshot;
}
function parseCaptureRecord(source) {
    const record = source["record"];
    if (!isRecord(record))
        badRequest("invalid-field", '"record" must be a JSON object');
    return record;
}
function parseCaptureSaveRequest(body) {
    const source = readJsonObject(body, ["record", "collectionIds", "tagNames", "note", "review"]);
    const collectionIds = source["collectionIds"];
    if (!Array.isArray(collectionIds) || collectionIds.length > MAX_CAPTURE_COLLECTIONS) {
        badRequest("invalid-field", '"collectionIds" must be a bounded array of strings');
    }
    if (collectionIds.some((entry) => typeof entry !== "string" || entry.length === 0)) {
        badRequest("invalid-field", '"collectionIds" must contain non-empty strings');
    }
    const tagNames = source["tagNames"];
    if (!Array.isArray(tagNames) || tagNames.length > MAX_CAPTURE_TAGS) {
        badRequest("invalid-field", '"tagNames" must be a bounded array of strings');
    }
    if (tagNames.some((entry) => typeof entry !== "string" || entry.trim() === "" || entry.length > 128)) {
        badRequest("invalid-field", '"tagNames" must contain bounded non-empty strings');
    }
    const note = source["note"];
    if (typeof note !== "string" || note.length > MAX_CAPTURE_NOTE_LENGTH) {
        badRequest("invalid-field", '"note" must be a bounded string');
    }
    const review = source["review"];
    if (!isRecord(review))
        badRequest("invalid-field", '"review" must be an object');
    for (const key of Object.keys(review)) {
        if (!["fingerprint", "itemIds", "action", "targetItemId"].includes(key)) {
            badRequest("unknown-field", `unknown review field "${key}"`);
        }
    }
    const fingerprint = review["fingerprint"];
    const itemIds = review["itemIds"];
    const action = review["action"];
    if (typeof fingerprint !== "string" || fingerprint.length === 0 || fingerprint.length > 128) {
        badRequest("invalid-field", '"review.fingerprint" must be a bounded string');
    }
    if (!Array.isArray(itemIds) || itemIds.length > MAX_CAPTURE_TAGS || itemIds.some((entry) => typeof entry !== "string" || entry.length === 0)) {
        badRequest("invalid-field", '"review.itemIds" must be a bounded array of strings');
    }
    if (action !== "save-new" && action !== "update-existing") {
        badRequest("invalid-field", '"review.action" is not supported');
    }
    const targetItemId = review["targetItemId"];
    if (targetItemId !== undefined && (typeof targetItemId !== "string" || targetItemId.length === 0)) {
        badRequest("invalid-field", '"review.targetItemId" must be a non-empty string when present');
    }
    return {
        record: parseCaptureRecord(source),
        collectionIds: collectionIds,
        tagNames: tagNames,
        note,
        review: {
            fingerprint,
            itemIds: itemIds,
            action,
            ...(targetItemId === undefined ? {} : { targetItemId }),
        },
    };
}
function captureWorkflowFailure(outcome) {
    const status = outcome.code === "stale-review" || outcome.code === "target-not-found" ? 409 : 422;
    return { status, body: { error: outcome.code, message: outcome.message } };
}
function parseCollaboratorReviewRequest(body) {
    const source = readJsonObject(body, ["snapshots", "importItemIds"]);
    const rawSnapshots = source["snapshots"];
    if (!Array.isArray(rawSnapshots) || rawSnapshots.length === 0) {
        badRequest("invalid-field", '"snapshots" must be a non-empty array');
    }
    if (rawSnapshots.length > MAX_COLLABORATOR_SNAPSHOTS) {
        badRequest("too-many", `"snapshots" accepts at most ${MAX_COLLABORATOR_SNAPSHOTS} entries`);
    }
    const seen = new Set();
    const snapshots = rawSnapshots.map((entry) => {
        if (!isRecord(entry))
            badRequest("invalid-field", 'every entry in "snapshots" must be an object');
        for (const key of Object.keys(entry)) {
            if (!["itemId", "contentHash", "snapshotAt", "csl"].includes(key)) {
                badRequest("unknown-field", `unknown field "${key}" in a snapshot`);
            }
        }
        const itemId = requireString(entry, "itemId");
        if (seen.has(itemId))
            badRequest("duplicate-item", '"snapshots" must not repeat an item id');
        seen.add(itemId);
        const contentHash = requireString(entry, "contentHash");
        if (!/^[0-9a-f]{64}$/u.test(contentHash)) {
            badRequest("invalid-field", '"contentHash" must be a lowercase SHA-256 fingerprint');
        }
        const snapshotAt = requireString(entry, "snapshotAt");
        const csl = entry["csl"];
        if (!isRecord(csl))
            badRequest("invalid-field", '"csl" must be a JSON object');
        return { itemId, contentHash, snapshotAt, csl };
    });
    const rawImports = source["importItemIds"] ?? [];
    if (!Array.isArray(rawImports) || rawImports.length > snapshots.length) {
        badRequest("invalid-field", '"importItemIds" must be an array no longer than "snapshots"');
    }
    const importSeen = new Set();
    const importItemIds = rawImports.map((entry) => {
        if (typeof entry !== "string" || entry.length === 0 || !seen.has(entry)) {
            badRequest("invalid-field", 'every import id must name a supplied snapshot');
        }
        if (importSeen.has(entry))
            badRequest("duplicate-item", '"importItemIds" must not repeat an item id');
        importSeen.add(entry);
        return entry;
    });
    return { snapshots, importItemIds };
}
export { BridgeRequestError } from "./errors.js";
/**
 * Every route answered from the ports alone — search, items, citations, bibliography, styles,
 * import, collaborator review and capture — and not the document-backup upload, which holds
 * session state and decodes bytes (`router.ts` adds it for the listener).
 *
 * The Word session runs on this (ADR-0242): it asks only for library reads and citation renders,
 * and it runs in the task pane too, where the upload's byte handling has no place. Refusals are the
 * same `{ error, message }` bodies `createBridgeRouter` returns; anything else is the port's
 * failure and is rethrown for the caller's 500 path, which says nothing about it (invariant 6).
 */
export function createPortRouter(ports) {
    return async (context) => {
        try {
            return await route(ports, context);
        }
        catch (error) {
            if (error instanceof BridgeRequestError) {
                return { status: error.status, body: { error: error.code, message: error.message } };
            }
            throw error;
        }
    };
}
/**
 * A minimal composition root for hosts that only provide browser capture.
 *
 * The desktop shell can start the authenticated bridge before the full Word/citation adapter is
 * staged. Keeping this handler in the bridge package preserves the same route validation while
 * letting the capture runtime inject only its two reviewed ports.
 */
export function createCaptureRouter(ports) {
    return async (context) => captureRoute(ports, context);
}
async function route(ports, context) {
    switch (context.route.path) {
        case "/search": {
            const text = context.query["q"];
            if (text === undefined || text.trim().length === 0) {
                badRequest("invalid-query", '"q" is required');
            }
            for (const key of Object.keys(context.query)) {
                if (!["q", "limit", "prefer", "collection"].includes(key)) {
                    badRequest("unknown-field", `unknown query parameter "${key}"`);
                }
            }
            const prefer = context.query["prefer"];
            // §9.2 lists collection among the fields the picker searches. An empty value is refused
            // rather than read as "the whole library": the caller wrote `collection=` for a reason, and
            // quietly widening a scoped search is the silent-adjustment failure this module refuses.
            const collection = context.query["collection"];
            if (collection !== undefined && collection.length === 0) {
                badRequest("invalid-query", '"collection" must not be empty when present');
            }
            const hits = await ports.library.search({
                text,
                limit: parseLimit(context.query),
                // §9.2: the add-in passes the items the document already cites, comma-separated.
                preferItemIds: prefer === undefined || prefer.length === 0 ? [] : prefer.split(","),
                ...(collection === undefined ? {} : { collectionId: collection }),
            });
            return { status: 200, body: { results: hits } };
        }
        case "/annotations/search": {
            if (ports.annotations === undefined)
                return null;
            for (const key of Object.keys(context.query)) {
                if (!["q", "limit", "offset"].includes(key)) {
                    badRequest("unknown-field", `unknown query parameter "${key}"`);
                }
            }
            const text = context.query["q"];
            if (text !== undefined) {
                const hasControl = [...text].some((character) => {
                    const code = character.codePointAt(0) ?? 0;
                    return code <= 0x1f || code === 0x7f;
                });
                if (text.trim().length === 0 || text.length > 1_000 || hasControl) {
                    badRequest("invalid-query", '"q" must be non-empty, at most 1000 characters, and contain no controls');
                }
            }
            return {
                status: 200,
                body: await ports.annotations.search({
                    ...(text === undefined ? {} : { text }),
                    limit: parseLimit(context.query),
                    offset: parseOffset(context.query),
                }),
            };
        }
        case "/recent": {
            // §9.2 "Show recent references": the opening state of the picker, which `/search` cannot
            // express because `q` is required there and an empty `q` must never mean "everything".
            for (const key of Object.keys(context.query)) {
                if (key !== "limit") {
                    badRequest("unknown-field", `unknown query parameter "${key}"`);
                }
            }
            return { status: 200, body: { results: await ports.library.listRecent({ limit: parseLimit(context.query) }) } };
        }
        case "/collections": {
            // §9.2's picker scope selector. No parameters at all: a library's collection tree is small
            // and the whole of it is what a selector offers, so there is no page to ask for and no
            // filter to get wrong. An unknown parameter is refused rather than ignored, for the reason
            // every other route refuses one — a caller that wrote `?limit=10` meant it, and answering
            // with the whole list while saying nothing is the silent adjustment this module refuses.
            for (const key of Object.keys(context.query)) {
                badRequest("unknown-field", `unknown query parameter "${key}"`);
            }
            return { status: 200, body: { collections: await ports.library.listCollections() } };
        }
        case "/items/:id": {
            const item = await ports.library.getItem(context.route.params["id"] ?? "");
            if (item === null) {
                // Not an error: the add-in asks about items a document cites, and a cited item may have
                // been deleted from the library. §9.3's snapshot is what keeps the document working.
                return { status: 404, body: { error: "not-found", message: "no such item in this library" } };
            }
            return { status: 200, body: { item } };
        }
        case "/citation/preview":
            return { status: 200, body: await ports.citation.previewCitation(parseCitationRequest(context.body)) };
        case "/citation/format":
            return { status: 200, body: await ports.citation.formatCitation(parseCitationRequest(context.body)) };
        case "/bibliography":
            return { status: 200, body: await ports.citation.formatBibliography(parseBibliographyRequest(context.body)) };
        case "/styles":
            return { status: 200, body: { styles: await ports.library.listStyles() } };
        case "/import":
            return { status: 200, body: await ports.library.importIdentifiers(parseImportRequest(context.body)) };
        case "/collaboration/references":
            return {
                status: 200,
                body: await ports.library.reviewCollaboratorReferences(parseCollaboratorReviewRequest(context.body)),
            };
        case "/capture":
        case "/capture/options":
        case "/capture/duplicates":
        case "/capture/save":
            return captureRoute(ports, context);
        case "/document-backup/begin":
        case "/document-backup/slice":
        case "/document-backup/commit":
        case "/document-backup/read":
            // Served by `handleDocumentBackup` before this switch is reached. Listed rather than left to
            // a default, so adding a route to ADR-0008's list without serving it fails to compile.
            return null;
        case "/health":
            // Answered by the listener itself; a router that also answered it would be a second
            // opinion about whether the bridge is alive.
            return null;
    }
}
async function captureRoute(ports, context) {
    switch (context.route.path) {
        case "/capture": {
            if (ports.capture === undefined)
                return null;
            const outcome = await ports.capture.translate(parseCaptureSnapshot(context.body));
            if (!outcome.ok)
                return { status: 400, body: { error: outcome.code, message: outcome.message } };
            return { status: 200, body: { candidates: outcome.candidates, findings: outcome.findings } };
        }
        case "/capture/options": {
            if (ports.captureWorkflow === undefined)
                return null;
            return { status: 200, body: await ports.captureWorkflow.options() };
        }
        case "/capture/duplicates": {
            if (ports.captureWorkflow === undefined)
                return null;
            const outcome = await ports.captureWorkflow.findDuplicates(parseCaptureRecord(readJsonObject(context.body, ["record"])));
            if (!outcome.ok)
                return captureWorkflowFailure(outcome);
            return { status: 200, body: { fingerprint: outcome.fingerprint, duplicates: outcome.duplicates } };
        }
        case "/capture/save": {
            if (ports.captureWorkflow === undefined)
                return null;
            const outcome = await ports.captureWorkflow.save(parseCaptureSaveRequest(context.body));
            if (!outcome.ok)
                return captureWorkflowFailure(outcome);
            return { status: 200, body: outcome.data };
        }
        default:
            return null;
    }
}
