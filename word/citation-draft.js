/**
 * The insertion draft, as the task pane itself can run it (SPEC §9.2; ADR-0086, ADR-0106,
 * ADR-0240; E10-04.5).
 *
 * ## Why this file exists at all
 *
 * `packages/word-addin/src/draft.ts` is the authority on what a draft is and what it refuses. The
 * task pane cannot import it: **no bundler enters this tree** (ADR-0086), so a bare specifier does
 * not resolve in the pane's webview, and compiling one in is the transform step ADR-0083/ADR-0086
 * exclude. Until E10-04.5 that did not matter, because the pane had no picker and never held a
 * draft. It holds one now, so the choice was between a pane that cannot assemble a citation and a
 * second implementation of the rules.
 *
 * **The duplication is deliberate and it is guarded.** `apps/word-taskpane/src/citation-draft.test.ts`
 * runs this module and `@refmgr/word-addin` over the same case table and requires identical
 * outcomes, operation by operation and refusal code by refusal code. A divergence is a failing
 * test, not a behaviour the user discovers in Word. Change one side and the test tells you about
 * the other; it is not a comment asking you to remember.
 *
 * Everything the TypeScript module promises holds here for the same reasons, and the reasoning is
 * not repeated — read `draft.ts` for it. In particular: it is pure, it refuses rather than
 * repairs, it adds no spacing to affixes, and **it formats nothing** (invariant 4). No text this
 * module returns is ever shown as a citation; it builds the *request* whose answer is one.
 *
 * No refusal or notice message names a title, a creator, a locator value or an affix — a draft
 * holds the user's own writing and a prefix may quote the source (§23, invariant 6).
 */

/** The ceiling on one cluster. Equal to `DRAFT_MAX_ITEMS`, and asserted equal by the drift test. */
export const DRAFT_MAX_ITEMS = 50;

/**
 * CSL 1.0.2's locator labels. citeproc does not reject an unknown label — it prints the locator
 * value with **no label at all** — so offering one CSL does not define silently turns
 * "sprocket 3" into "3".
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

/** Shown when a reference has no title. Never a guess assembled from other fields. */
export const UNTITLED_REFERENCE_LABEL = "(untitled reference)";

/** The scope that means "the whole library" — the absence of `?collection=` on the wire. */
export const LIBRARY_SCOPE_ID = "";

function isBlank(value) {
  return typeof value !== "string" || value.trim().length === 0;
}

function refuse(code, message, entryId) {
  return Object.freeze({
    ok: false,
    refusal: Object.freeze(entryId === undefined ? { code, message } : { code, message, entryId }),
  });
}

function accept(draft, notices) {
  return Object.freeze({ ok: true, draft, notices: Object.freeze(notices) });
}

/** A draft citing nothing. The state the picker opens in. */
export function emptyDraft() {
  return Object.freeze({ entries: Object.freeze([]), nextOrdinal: 1 });
}

export function draftSize(draft) {
  return draft.entries.length;
}

export function findEntry(draft, entryId) {
  return draft.entries.find((entry) => entry.entryId === entryId);
}

/** §9.2 "Allow one or multiple items" — append a reference. Appending is the order the user built. */
export function addEntry(draft, itemId) {
  if (isBlank(itemId)) return refuse("invalid-item-id", "an item id is required");
  if (draft.entries.length >= DRAFT_MAX_ITEMS) {
    return refuse(
      "too-many-items",
      `a citation may cite at most ${DRAFT_MAX_ITEMS} references; this one already cites ${draft.entries.length}`,
    );
  }
  const entry = Object.freeze({ entryId: `entry-${draft.nextOrdinal}`, itemId, authorMode: "normal" });
  const notices = [];
  if (draft.entries.some((existing) => existing.itemId === itemId)) {
    notices.push(
      Object.freeze({
        code: "repeated-item",
        entryId: entry.entryId,
        message: `this citation already cites item ${itemId}; it will be cited twice`,
      }),
    );
  }
  return accept(
    Object.freeze({ entries: Object.freeze([...draft.entries, entry]), nextOrdinal: draft.nextOrdinal + 1 }),
    notices,
  );
}

/** Remove one entry. `nextOrdinal` does not go back: an entry id is never reused. */
export function removeEntry(draft, entryId) {
  if (findEntry(draft, entryId) === undefined) {
    return refuse("unknown-entry", `no entry ${entryId} in this draft`, entryId);
  }
  return accept(
    Object.freeze({
      entries: Object.freeze(draft.entries.filter((entry) => entry.entryId !== entryId)),
      nextOrdinal: draft.nextOrdinal,
    }),
    [],
  );
}

/**
 * §9.2 "Allow reordering within a grouped citation" — move one entry to an absolute position,
 * where the position is where it lands **after** it is lifted out. Refused, never clamped.
 */
export function moveEntry(draft, entryId, toPosition) {
  const from = draft.entries.findIndex((entry) => entry.entryId === entryId);
  if (from === -1) return refuse("unknown-entry", `no entry ${entryId} in this draft`, entryId);
  if (!Number.isInteger(toPosition) || toPosition < 0 || toPosition > draft.entries.length - 1) {
    return refuse(
      "position-out-of-range",
      `position must be a whole number between 0 and ${draft.entries.length - 1}`,
      entryId,
    );
  }
  const entries = [...draft.entries];
  const [moved] = entries.splice(from, 1);
  entries.splice(toPosition, 0, moved);
  return accept(Object.freeze({ ...draft, entries: Object.freeze(entries) }), []);
}

function replaceEntry(draft, entryId, next) {
  return Object.freeze({
    ...draft,
    entries: Object.freeze(draft.entries.map((entry) => (entry.entryId === entryId ? next : entry))),
  });
}

function without(entry, key) {
  const copy = { ...entry };
  delete copy[key];
  return Object.freeze(copy);
}

/**
 * §9.2 "Support locators" — set or clear one entry's locator.
 *
 * `null` clears. A label with a blank value is refused rather than treated as a clear: citeproc's
 * own behaviour on the second — drop the locator, say nothing — is exactly what makes a page
 * number vanish between the pane and the manuscript.
 */
export function setLocator(draft, entryId, locator) {
  const entry = findEntry(draft, entryId);
  if (entry === undefined) return refuse("unknown-entry", `no entry ${entryId} in this draft`, entryId);
  if (locator === null) return accept(replaceEntry(draft, entryId, without(entry, "locator")), []);
  if (isBlank(locator.label)) return refuse("invalid-locator-label", "a locator label is required", entryId);
  if (isBlank(locator.value)) {
    return refuse(
      "empty-locator-value",
      `the "${locator.label}" locator has no value; clear the locator instead`,
      entryId,
    );
  }
  const notices = [];
  if (!LOCATOR_LABELS.has(locator.label)) {
    notices.push(
      Object.freeze({
        code: "unknown-locator-label",
        entryId,
        message: `"${locator.label}" is not a CSL 1.0.2 locator label; the locator will print with no label`,
      }),
    );
  }
  // Verbatim, including whatever spacing the user typed inside the value.
  return accept(
    replaceEntry(draft, entryId, Object.freeze({ ...entry, locator: Object.freeze({ label: locator.label, value: locator.value }) })),
    notices,
  );
}

/**
 * §9.2 "Support … prefix, suffix" — stored exactly as typed. No trimming, no added spacing, no
 * added punctuation: the processor supplies the space, and a prefix is the user's own writing.
 * An affix that is only whitespace is a clear, and is the one normalisation here.
 */
export function setAffix(draft, entryId, which, text) {
  const entry = findEntry(draft, entryId);
  if (entry === undefined) return refuse("unknown-entry", `no entry ${entryId} in this draft`, entryId);
  if (text === null || isBlank(text)) {
    return accept(replaceEntry(draft, entryId, without(entry, which)), []);
  }
  return accept(replaceEntry(draft, entryId, Object.freeze({ ...entry, [which]: text })), []);
}

/** §9.2 "Support … author suppression". One setting with three values; the pair is not expressible. */
export function setAuthorMode(draft, entryId, mode) {
  const entry = findEntry(draft, entryId);
  if (entry === undefined) return refuse("unknown-entry", `no entry ${entryId} in this draft`, entryId);
  return accept(replaceEntry(draft, entryId, Object.freeze({ ...entry, authorMode: mode })), []);
}

/** Everything currently worth saying about a draft, in entry order. Re-derived, never accumulated. */
export function draftNotices(draft) {
  const notices = [];
  const seen = new Set();
  for (const entry of draft.entries) {
    if (seen.has(entry.itemId)) {
      notices.push(
        Object.freeze({
          code: "repeated-item",
          entryId: entry.entryId,
          message: `this citation already cites item ${entry.itemId}; it will be cited twice`,
        }),
      );
    }
    seen.add(entry.itemId);
    if (entry.locator !== undefined && !LOCATOR_LABELS.has(entry.locator.label)) {
      notices.push(
        Object.freeze({
          code: "unknown-locator-label",
          entryId: entry.entryId,
          message: `"${entry.locator.label}" is not a CSL 1.0.2 locator label; the locator will print with no label`,
        }),
      );
    }
  }
  return Object.freeze(notices);
}

function toItemRequest(entry) {
  const item = { itemId: entry.itemId };
  if (entry.locator !== undefined) item.locator = { label: entry.locator.label, value: entry.locator.value };
  if (entry.prefix !== undefined) item.prefix = entry.prefix;
  if (entry.suffix !== undefined) item.suffix = entry.suffix;
  if (entry.authorMode === "suppress-author") item.suppressAuthor = true;
  if (entry.authorMode === "author-only") item.authorOnly = true;
  return item;
}

/**
 * §9.2 "Preview the formatted result before insertion" — the request whose answer is the preview.
 *
 * The same request serves `/citation/preview` and `/citation/format`; they differ in what the
 * caller does with the answer, not in what is asked.
 */
export function citationRequest(draft, options) {
  if (draft.entries.length === 0) {
    return Object.freeze({ ok: false, refusal: Object.freeze({ code: "empty-draft", message: "this citation cites nothing yet" }) });
  }
  if (isBlank(options?.styleId)) {
    return Object.freeze({ ok: false, refusal: Object.freeze({ code: "invalid-style-id", message: "a style id is required" }) });
  }
  if (isBlank(options?.locale)) {
    return Object.freeze({ ok: false, refusal: Object.freeze({ code: "invalid-locale", message: "a locale is required" }) });
  }
  return Object.freeze({
    ok: true,
    request: Object.freeze({
      styleId: options.styleId,
      locale: options.locale,
      items: Object.freeze(draft.entries.map(toItemRequest)),
    }),
    notices: draftNotices(draft),
  });
}

// --- the picker's own view of a draft -------------------------------------------------------------

/** How many entries cite each item: the checkbox state, and the repeat count. */
export function draftSelectionCounts(draft) {
  const counts = new Map();
  for (const entry of draft.entries) counts.set(entry.itemId, (counts.get(entry.itemId) ?? 0) + 1);
  return counts;
}

/** The distinct references the draft cites, in the order they were first added. */
export function selectedItemIds(draft) {
  return Object.freeze([...new Set(draft.entries.map((entry) => entry.itemId))]);
}

/**
 * §9.2 "Allow one or multiple items" — tick or untick one reference.
 *
 * Both directions refuse when the box and the draft disagree rather than making the draft match
 * the box: a tick on a reference already cited would add a cite the user did not ask for, and an
 * untick on a reference cited twice cannot say which cite to drop.
 */
export function toggleCitationSelection(draft, itemId, selected) {
  const id = typeof itemId === "string" ? itemId : "";
  if (id.trim().length === 0) return refuse("invalid-item-id", "an item id is required");
  const count = draftSelectionCounts(draft).get(id) ?? 0;

  if (selected) {
    if (count > 0) {
      return Object.freeze({
        ok: false,
        refusal: Object.freeze({ code: "already-selected", message: `item ${id} is already in this citation`, itemId: id }),
      });
    }
    const outcome = addEntry(draft, id);
    if (!outcome.ok) {
      return Object.freeze({
        ok: false,
        refusal: Object.freeze({
          code: outcome.refusal.code === "too-many-items" ? "too-many-items" : "invalid-item-id",
          message: outcome.refusal.message,
          itemId: id,
        }),
      });
    }
    return accept(outcome.draft, outcome.notices);
  }

  if (count === 0) {
    return Object.freeze({ ok: false, refusal: Object.freeze({ code: "not-selected", message: `item ${id} is not in this citation`, itemId: id }) });
  }
  if (count > 1) {
    return Object.freeze({
      ok: false,
      refusal: Object.freeze({
        code: "ambiguous-repeat",
        message: `this citation cites item ${id} ${count} times; remove the entries individually`,
        itemId: id,
      }),
    });
  }
  const entry = draft.entries.find((candidate) => candidate.itemId === id);
  const outcome = removeEntry(draft, entry?.entryId ?? "");
  if (!outcome.ok) {
    return Object.freeze({ ok: false, refusal: Object.freeze({ code: "not-selected", message: outcome.refusal.message, itemId: id }) });
  }
  return accept(outcome.draft, draftNotices(outcome.draft));
}

function text(value) {
  return typeof value === "string" ? value.trim() : "";
}

function optional(value) {
  const trimmed = text(value);
  return trimmed.length === 0 ? null : trimmed;
}

/**
 * One library row, ready to render.
 *
 * `itemType` is the library's own key, shown unmapped: the display-label registry is in
 * `@refmgr/core`, and a second copy of it here could disagree with that one silently.
 */
export function citationPickerRow(hit, context) {
  const itemId = text(hit?.itemId);
  if (itemId.length === 0) throw new Error("InvalidCitationPickerHit");
  const counts = draftSelectionCounts(context.draft);
  const selectionCount = counts.get(itemId) ?? 0;
  const year = optional(hit.year);
  const containerTitle = optional(hit.containerTitle);
  const itemType = text(hit.itemType);
  return Object.freeze({
    itemId,
    title: optional(hit.title) ?? UNTITLED_REFERENCE_LABEL,
    creatorSummary: text(hit.creatorSummary),
    publication: [containerTitle, year, itemType].filter((part) => part !== null && part.length > 0).join(" · "),
    year,
    containerTitle,
    doi: optional(hit.doi),
    itemType,
    selected: selectionCount > 0,
    selectionCount,
    citedInDocument: (context.citedItemIds ?? []).includes(itemId),
  });
}

export function citationPickerRows(hits, context) {
  return Object.freeze(hits.map((hit) => citationPickerRow(hit, context)));
}

/**
 * Order the collection tree for a selector: a pre-order walk, parents before their children.
 *
 * An orphan or a cycle becomes a flat, reachable row rather than an empty dropdown — a scope the
 * library holds and the selector omits is a part of the user's library they cannot reach from Word.
 */
export function citationScopeOptions(scopes) {
  const byId = new Map();
  for (const scope of scopes) {
    const id = text(scope?.collectionId);
    if (id.length === 0 || byId.has(id)) throw new Error("InvalidCitationPickerScopes");
    byId.set(id, scope);
  }
  const parentOf = (scope) => {
    const parent = text(scope.parentId);
    return parent.length > 0 && byId.has(parent) ? parent : null;
  };
  const inCycle = (scope) => {
    const seen = new Set([text(scope.collectionId)]);
    let parent = parentOf(scope);
    while (parent !== null) {
      if (seen.has(parent)) return true;
      seen.add(parent);
      parent = parentOf(byId.get(parent));
    }
    return false;
  };

  const children = new Map();
  const roots = [];
  const detached = new Set();
  for (const scope of scopes) {
    const id = text(scope.collectionId);
    if (inCycle(scope)) {
      detached.add(id);
      roots.push(scope);
      continue;
    }
    const parent = parentOf(scope);
    if (parent === null) {
      if (text(scope.parentId).length > 0) detached.add(id);
      roots.push(scope);
      continue;
    }
    const siblings = children.get(parent) ?? [];
    siblings.push(scope);
    children.set(parent, siblings);
  }

  const options = [];
  const walk = (scope, depth) => {
    const id = text(scope.collectionId);
    options.push(Object.freeze({ collectionId: id, name: text(scope.name), depth, detached: detached.has(id) }));
    for (const child of children.get(id) ?? []) walk(child, depth + 1);
  };
  for (const root of roots) walk(root, 0);
  return Object.freeze(options);
}
