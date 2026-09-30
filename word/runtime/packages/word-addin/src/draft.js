/**
 * The insertion draft — SPEC §9.2, task E10-04.3; ADR-0106.
 *
 * A draft is **the citation the user is assembling, before it exists anywhere**. §9.2 asks the
 * insertion experience to "allow one or multiple items", to "support locators, prefix, suffix and
 * author suppression", to "allow reordering within a grouped citation" and to "preview the
 * formatted result before insertion". Everything in that sentence except the preview itself is
 * this module, and the preview is the request this module produces.
 *
 * Three properties are the point of keeping it here rather than in the task pane's DOM code:
 *
 * - **It is pure.** No socket, no Office.js, no clock, no randomness. Every function takes a draft
 *   and returns a new one; a draft is a value the pane can hold, store and re-render. The whole
 *   failure mode this module exists to prevent — a locator or an affix quietly dropped between the
 *   picker and the document — is invisible to a test that has to drive a UI to see it.
 * - **It refuses rather than repairs.** Every operation either produces a new draft or a refusal
 *   naming a code; nothing is silently trimmed, clamped, substituted or deduplicated. A picker
 *   that quietly drops an empty locator teaches the user their input was accepted.
 * - **It formats nothing.** No text this module returns is ever shown as a citation (invariant 4).
 *   It builds the *request* whose answer is the citation, and the answer comes from the CSL
 *   processor on the other side of the bridge.
 *
 * ## What it deliberately does not do
 *
 * **It does not validate against the library.** Whether an item id exists is the bridge's answer,
 * not a guess made from the pane's cached search results.
 *
 * **It does not add spacing to affixes, and must not.** This was measured rather than assumed: with
 * the repository's own citeproc-js adapter, a prefix of `"see"` and a prefix of `"see "` both
 * render `(see Smith, 2020)`, and a suffix of `"emphasis added"` and `" emphasis added"` both
 * render `(Smith, 2020 emphasis added)`, while `", emphasis added"` renders
 * `(Smith, 2020, emphasis added)`. The processor already knows where a space belongs and where
 * punctuation means it does not. Adding one here would double it — and would be this package
 * formatting a citation, which invariant 4 forbids outright. The measurement is pinned by
 * `tests/insertion-draft-contract.test.ts` so a processor change cannot quietly invalidate this
 * paragraph. `packages/citation` states the same rule from the other end: affixes cross unchanged.
 *
 * **It does not promise the rendered order is the draft's order.** Reordering is what §9.2 asks
 * for and what `moveEntry` does, but a style may carry `<sort>` in its citation layout — a numeric
 * style sorted by citation number will print a cluster in its own order however the user arranged
 * it. The draft's order is what citeproc is *given*; the preview is the only honest account of
 * what comes out. Sorting the draft to match, or hiding the reorder control, would both be this
 * package having a second opinion about a style rule.
 */
/**
 * The ceiling on one cluster, restated from the bridge's `MAX_CITATION_ITEMS` because this package
 * imports no other (ARCHITECTURE §2) and asserted equal to it by
 * `tests/insertion-draft-contract.test.ts`.
 *
 * It is enforced here as well as there so the refusal arrives on the click that would exceed it,
 * naming the limit, rather than as a 400 at preview time after the user has built the cluster.
 */
export const DRAFT_MAX_ITEMS = 50;
/**
 * CSL 1.0.2's locator labels, restated for the same reason and asserted identical to
 * `CSL_LOCATOR_LABELS` in `@refmgr/doc-schema` by the same test.
 *
 * The pane needs this list to offer labels, and offering a label CSL does not define is not a
 * cosmetic error: citeproc does not reject an unknown label, it prints the locator value with **no
 * label at all**, so "sprocket 3" silently becomes "3".
 */
export const DRAFT_LOCATOR_LABELS = Object.freeze([
    "act",
    "appendix",
    "article-locator",
    "book",
    "canon",
    "chapter",
    "column",
    "elocation",
    "equation",
    "figure",
    "folio",
    "issue",
    "line",
    "note",
    "opus",
    "page",
    "paragraph",
    "part",
    "rule",
    "scene",
    "section",
    "sub-verbo",
    "supplement",
    "table",
    "timestamp",
    "title-locator",
    "verse",
    "version",
    "volume",
]);
const LOCATOR_LABELS = new Set(DRAFT_LOCATOR_LABELS);
function refuse(code, message, entryId) {
    return { ok: false, refusal: entryId === undefined ? { code, message } : { code, message, entryId } };
}
function isBlank(value) {
    return value.trim().length === 0;
}
/** A draft citing nothing. The state the picker opens in. */
export function emptyDraft() {
    return { entries: [], nextOrdinal: 1 };
}
/** How many references the draft cites. Repeats count once each: two entries, two cites. */
export function draftSize(draft) {
    return draft.entries.length;
}
export function findEntry(draft, entryId) {
    return draft.entries.find((entry) => entry.entryId === entryId);
}
/**
 * §9.2 "Allow one or multiple items" — append a reference to the draft.
 *
 * Appends rather than inserts: the picker's own order is the order the user built, and an item
 * that arrived somewhere other than the end of the list is an item the user has to hunt for.
 * `moveEntry` is how it gets somewhere else.
 */
export function addEntry(draft, itemId) {
    if (isBlank(itemId)) {
        return refuse("invalid-item-id", "an item id is required");
    }
    if (draft.entries.length >= DRAFT_MAX_ITEMS) {
        return refuse("too-many-items", `a citation may cite at most ${DRAFT_MAX_ITEMS} references; this one already cites ${draft.entries.length}`);
    }
    const entry = {
        entryId: `entry-${draft.nextOrdinal}`,
        itemId,
        authorMode: "normal",
    };
    const notices = [];
    if (draft.entries.some((existing) => existing.itemId === itemId)) {
        notices.push({
            code: "repeated-item",
            entryId: entry.entryId,
            message: `this citation already cites item ${itemId}; it will be cited twice`,
        });
    }
    return {
        ok: true,
        draft: { entries: [...draft.entries, entry], nextOrdinal: draft.nextOrdinal + 1 },
        notices,
    };
}
/** Remove one entry. Removing the last one leaves an empty draft, which is a legitimate state. */
export function removeEntry(draft, entryId) {
    if (findEntry(draft, entryId) === undefined) {
        return refuse("unknown-entry", `no entry ${entryId} in this draft`, entryId);
    }
    return {
        ok: true,
        // `nextOrdinal` does not go back. See the note on `InsertionDraft`.
        draft: { entries: draft.entries.filter((entry) => entry.entryId !== entryId), nextOrdinal: draft.nextOrdinal },
        notices: [],
    };
}
/**
 * §9.2 "Allow reordering within a grouped citation" — move one entry to an absolute position.
 *
 * Absolute rather than "up"/"down" because the pane may offer a drag as well as a button, and a
 * position is the only form that expresses both without the caller composing several moves and
 * having to reason about what the intermediate states looked like.
 *
 * The position is where the entry ends up **after** it is lifted out, which is what a drop target
 * means. Moving entry 0 of three to position 2 therefore puts it last, rather than second — the
 * alternative reading makes the last position unreachable for the first entry.
 */
export function moveEntry(draft, entryId, toPosition) {
    const from = draft.entries.findIndex((entry) => entry.entryId === entryId);
    if (from === -1) {
        return refuse("unknown-entry", `no entry ${entryId} in this draft`, entryId);
    }
    if (!Number.isInteger(toPosition) || toPosition < 0 || toPosition > draft.entries.length - 1) {
        // Refused, not clamped, for the reason the bridge refuses a limit out of range (ADR-0100): a
        // clamp answers a question nobody asked and the caller never learns it was wrong.
        return refuse("position-out-of-range", `position must be a whole number between 0 and ${draft.entries.length - 1}`, entryId);
    }
    const entries = [...draft.entries];
    const [moved] = entries.splice(from, 1);
    entries.splice(toPosition, 0, moved);
    return { ok: true, draft: { ...draft, entries }, notices: [] };
}
function replaceEntry(draft, entryId, next) {
    return { ...draft, entries: draft.entries.map((entry) => (entry.entryId === entryId ? next : entry)) };
}
/**
 * §9.2 "Support locators" — set or clear one entry's locator.
 *
 * `null` clears it. A label with a blank value is refused rather than treated as a clear: the two
 * are different intentions, and citeproc's own behaviour on the second — drop the locator, say
 * nothing — is exactly what makes a page number vanish between the pane and the manuscript.
 */
export function setLocator(draft, entryId, locator) {
    const entry = findEntry(draft, entryId);
    if (entry === undefined) {
        return refuse("unknown-entry", `no entry ${entryId} in this draft`, entryId);
    }
    if (locator === null) {
        const { locator: _removed, ...rest } = entry;
        return { ok: true, draft: replaceEntry(draft, entryId, rest), notices: [] };
    }
    if (isBlank(locator.label)) {
        return refuse("invalid-locator-label", "a locator label is required", entryId);
    }
    if (isBlank(locator.value)) {
        return refuse("empty-locator-value", `the "${locator.label}" locator has no value; clear the locator instead`, entryId);
    }
    const notices = [];
    if (!LOCATOR_LABELS.has(locator.label)) {
        notices.push({
            code: "unknown-locator-label",
            entryId,
            message: `"${locator.label}" is not a CSL 1.0.2 locator label; the locator will print with no label`,
        });
    }
    // Verbatim, including whatever spacing the user typed inside the value.
    return { ok: true, draft: replaceEntry(draft, entryId, { ...entry, locator }), notices };
}
/**
 * §9.2 "Support … prefix, suffix" — set or clear the text before or after one cite.
 *
 * `null` clears. The text is stored **exactly as typed**: no trimming, no added spacing, no added
 * punctuation. See the module note — the processor supplies the space, and a prefix is the user's
 * own writing, which this package is in no position to edit.
 *
 * An affix that is only whitespace is a clear, and the one normalisation here: it is indivisible
 * from what the user sees in an empty box, and sending `" "` would put a stray space in the
 * manuscript that nothing downstream could attribute to anything.
 */
export function setAffix(draft, entryId, which, text) {
    const entry = findEntry(draft, entryId);
    if (entry === undefined) {
        return refuse("unknown-entry", `no entry ${entryId} in this draft`, entryId);
    }
    if (text === null || isBlank(text)) {
        return { ok: true, draft: replaceEntry(draft, entryId, withoutAffix(entry, which)), notices: [] };
    }
    return { ok: true, draft: replaceEntry(draft, entryId, { ...entry, [which]: text }), notices: [] };
}
/**
 * Drops one affix key entirely rather than setting it to `undefined`, because
 * `exactOptionalPropertyTypes` and the wire agree that a field which is absent and a field which is
 * present-and-undefined are different things.
 */
function withoutAffix(entry, which) {
    if (which === "prefix") {
        const { prefix: _dropped, ...rest } = entry;
        return rest;
    }
    const { suffix: _dropped, ...rest } = entry;
    return rest;
}
/** §9.2 "Support … author suppression". See `DraftAuthorMode` for why this is one setting. */
export function setAuthorMode(draft, entryId, mode) {
    const entry = findEntry(draft, entryId);
    if (entry === undefined) {
        return refuse("unknown-entry", `no entry ${entryId} in this draft`, entryId);
    }
    return { ok: true, draft: replaceEntry(draft, entryId, { ...entry, authorMode: mode }), notices: [] };
}
/**
 * §9.2 "Preview the formatted result before insertion" — the request whose answer is the preview.
 *
 * The same request serves `/citation/preview` and `/citation/format`: they differ in what the
 * caller does with the answer, not in what is asked. Which is why this returns a request rather
 * than sending one — the draft has no socket, and the choice between previewing and inserting is
 * not the draft's to make.
 *
 * Notices from `addEntry` and `setLocator` are **re-derived here rather than accumulated**, so that
 * a preview always reports the draft as it now stands: an unknown locator label the user has since
 * corrected must not still be reported, and a repeat that arrived by two separate edits must be.
 */
export function citationRequest(draft, options) {
    if (draft.entries.length === 0) {
        return { ok: false, refusal: { code: "empty-draft", message: "this citation cites nothing yet" } };
    }
    if (isBlank(options.styleId)) {
        return { ok: false, refusal: { code: "invalid-style-id", message: "a style id is required" } };
    }
    if (isBlank(options.locale)) {
        return { ok: false, refusal: { code: "invalid-locale", message: "a locale is required" } };
    }
    return { ok: true, request: { styleId: options.styleId, locale: options.locale, items: draft.entries.map(toItemRequest) }, notices: draftNotices(draft) };
}
/**
 * Everything currently worth saying about a draft, in entry order.
 *
 * Exported because the pane shows these beside the entries themselves, not only at preview time.
 */
export function draftNotices(draft) {
    const notices = [];
    const seen = new Set();
    for (const entry of draft.entries) {
        if (seen.has(entry.itemId)) {
            notices.push({
                code: "repeated-item",
                entryId: entry.entryId,
                message: `this citation already cites item ${entry.itemId}; it will be cited twice`,
            });
        }
        seen.add(entry.itemId);
        if (entry.locator !== undefined && !LOCATOR_LABELS.has(entry.locator.label)) {
            notices.push({
                code: "unknown-locator-label",
                entryId: entry.entryId,
                message: `"${entry.locator.label}" is not a CSL 1.0.2 locator label; the locator will print with no label`,
            });
        }
    }
    return notices;
}
/**
 * One entry, as the bridge expects it.
 *
 * Absent fields are **omitted rather than sent as `undefined`**: the bridge refuses a `prefix` that
 * is present and not a string, and `JSON.stringify` drops an explicit `undefined` anyway, so
 * building the object without the key is the shape that survives serialisation intact. The author
 * booleans are written only when true, which is the only way `DraftAuthorMode`'s guarantee — that
 * the forbidden pair cannot be expressed — reaches the wire.
 */
function toItemRequest(entry) {
    const item = { itemId: entry.itemId };
    if (entry.locator !== undefined)
        item.locator = { label: entry.locator.label, value: entry.locator.value };
    if (entry.prefix !== undefined)
        item.prefix = entry.prefix;
    if (entry.suffix !== undefined)
        item.suffix = entry.suffix;
    if (entry.authorMode === "suppress-author")
        item.suppressAuthor = true;
    if (entry.authorMode === "author-only")
        item.authorOnly = true;
    return item;
}
