/**
 * Task-pane presentation for Insert citation (SPEC §9.2; ADR-0240; E10-04.5).
 *
 * The surface §9.2 always described and the pane never had: the library listed, scoped to a
 * collection, searched, ticked one reference at a time, ordered, given locators and affixes, and
 * previewed as the CSL processor renders it — before anything is written.
 *
 * Policy is `citation-draft.js`, which the pane may run and which
 * `apps/word-taskpane/src/citation-draft.test.ts` holds equal to `@refmgr/word-addin`. This module
 * owns the DOM and nothing else: it decides no citation, formats no citation, and re-renders from
 * the draft it was handed back rather than from what it believes it just did.
 *
 * Three rules it keeps deliberately:
 *
 * - **No library string ever enters an attribute or a class name.** Titles, creators and container
 *   titles are written with `textContent`; only ids travel in `dataset` (ADR-0084's rule, and the
 *   same one the annotation picker keeps). The search term is never logged (§23, invariant 6).
 * - **The preview is the processor's answer, shown verbatim.** This module never assembles,
 *   trims or punctuates citation text (invariant 4). While a preview is in flight the old one is
 *   marked stale rather than left looking current.
 * - **A refusal is shown, never swallowed.** Every refusal the draft can produce has a sentence
 *   here, and an unrecognised code still shows the draft's own message rather than nothing.
 */

import {
  LIBRARY_SCOPE_ID,
  citationPickerRows,
  citationRequest,
  citationScopeOptions,
  draftNotices,
  emptyDraft,
  moveEntry,
  removeEntry,
  selectedItemIds,
  setAffix,
  setAuthorMode,
  setLocator,
  toggleCitationSelection,
  DRAFT_LOCATOR_LABELS,
} from "./citation-draft.js";

/** Fixed pane copy for every refusal the picker can produce. Names no reference contents (§23). */
const REFUSAL_COPY = Object.freeze({
  "invalid-item-id": "That reference could not be used.",
  "already-selected": "This citation already includes that reference.",
  "not-selected": "That reference is not in this citation.",
  "ambiguous-repeat": "This citation includes that reference more than once. Remove the entries individually below.",
  "too-many-items": "This citation already includes as many references as one citation may hold.",
  "unknown-entry": "That entry is no longer in this citation.",
  "position-out-of-range": "That entry cannot move there.",
  "invalid-locator-label": "Choose a locator label.",
  "empty-locator-value": "Enter a locator value, or clear the locator.",
  "empty-draft": "Select a reference first.",
  "invalid-style-id": "Choose a citation style first.",
  "invalid-locale": "Choose a language first.",
});

/** The label shown for the whole-library scope. */
export const ALL_REFERENCES_LABEL = "All references";

/** How many references one page of the list holds. The bridge's own ceiling is 100. */
export const PICKER_PAGE_SIZE = 25;

/** How long typing must pause before the list is searched. Enter searches at once. */
export const SEARCH_DEBOUNCE_MS = 250;

const AUTHOR_MODE_LABELS = Object.freeze([
  ["normal", "Normal"],
  ["suppress-author", "Suppress author"],
  ["author-only", "Author only"],
]);

/**
 * The processor's warnings, in a writer's words. Its own messages name style URIs and CSL
 * vocabulary ("declares citation-format numeric"); a code this does not know still shows the
 * processor's sentence rather than nothing. `quiet` marks a note that is information, not a problem.
 */
const PREVIEW_WARNING_COPY = Object.freeze({
  "isolated-render": { text: "Numbered by position: the final number is set when the citation is inserted.", quiet: true },
  "style-sorted-order": { text: "This style sorts the references inside a citation, so the order above may change.", quiet: true },
  "locale-region-substituted": { text: "Your language variant is not installed, so the closest one is used.", quiet: false },
});

export function citationPreviewWarningCopy(warning) {
  const known = PREVIEW_WARNING_COPY[warning && warning.code];
  if (known) return known;
  const message = warning && typeof warning.message === "string" ? warning.message : "";
  return { text: message.length > 0 ? message : "The preview came with a warning.", quiet: false };
}

/** CSL's type codes, as a writer names them. Unknown codes are spaced and capitalised. */
const ITEM_TYPE_LABELS = Object.freeze({
  "article-journal": "Journal article",
  "article-magazine": "Magazine article",
  "article-newspaper": "Newspaper article",
  article: "Article",
  book: "Book",
  chapter: "Book chapter",
  "paper-conference": "Conference paper",
  thesis: "Thesis",
  report: "Report",
  webpage: "Web page",
  dataset: "Dataset",
  patent: "Patent",
  manuscript: "Manuscript",
});

export function itemTypeLabel(code) {
  if (typeof code !== "string" || code.length === 0) return "";
  const known = ITEM_TYPE_LABELS[code];
  if (known) return known;
  const spaced = code.replace(/-/gu, " ");
  return spaced.charAt(0).toUpperCase() + spaced.slice(1);
}

export function citationPickerRefusalCopy(code, fallback) {
  const message = REFUSAL_COPY[code];
  if (typeof message === "string") return message;
  // A code this module does not know is still a refusal the user has to see. Showing the draft's
  // own message is worse copy than a written sentence and far better than silence.
  return typeof fallback === "string" && fallback.length > 0 ? fallback : "That change was not made.";
}

/**
 * The sentence under the list: how many references the list holds, and how many are selected.
 *
 * Separate from the DOM so the wording is testable without a document, and so "1 reference" never
 * reads "1 references" — the kind of thing nobody notices until it is in front of a user.
 */
export function citationPickerSummary(shown, selected) {
  const references = `${shown} ${shown === 1 ? "reference" : "references"}`;
  if (selected === 0) return `${references}. Tick the ones to cite.`;
  return `${references}. ${selected} selected for this citation.`;
}

function element(document, name, className, textContent) {
  const node = document.createElement(name);
  if (className !== undefined) node.className = className;
  if (textContent !== undefined) node.textContent = textContent;
  return node;
}

function dispatchDetail(document, target, name, detail) {
  const event = document.createEvent("CustomEvent");
  event.initCustomEvent(name, true, false, detail);
  target.dispatchEvent(event);
}

/**
 * Mount the picker.
 *
 * `ports` are the bridge reads and the Word write. Every one is optional and an absent port is
 * **stated on the surface rather than discovered by clicking**: a pane whose desktop app is not
 * connected can still be read, and a button that cannot do its job says why it is disabled.
 */
export function mountCitationPickerSurface(document, ports = {}) {
  const form = document.getElementById("citation-search-form");
  const scope = document.getElementById("citation-scope");
  const query = document.getElementById("citation-search-query");
  const results = document.getElementById("citation-results");
  const status = document.getElementById("citation-picker-status");
  const count = document.getElementById("citation-draft-count");
  const entries = document.getElementById("citation-draft-entries");
  const notices = document.getElementById("citation-draft-notices");
  const preview = document.getElementById("citation-preview");
  const previewWarnings = document.getElementById("citation-preview-warnings");
  const clear = document.getElementById("citation-clear");
  const insert = document.getElementById("citation-insert");
  if (!form || !scope || !query || !results || !status || !count || !entries || !notices || !preview || !previewWarnings || !clear || !insert) {
    throw new Error("CitationPickerSurfaceMissing");
  }

  /** Full hits, in memory. Never written into attributes — see the module note. */
  const hitsById = new Map();
  let shown = [];
  let draft = emptyDraft();
  let citedItemIds = [];
  let style = { styleId: "", locale: "" };
  /** Rising, so a slow preview for an older draft can never overwrite a newer one. */
  let previewSequence = 0;
  let freshPreview = null;
  let inserting = false;
  const expandedEntries = new Set();

  function setStatus(message) {
    status.textContent = message;
  }

  function referenceLabel(itemId) {
    const hit = hitsById.get(itemId);
    if (hit === undefined) return itemId;
    const title = typeof hit.title === "string" && hit.title.trim().length > 0 ? hit.title.trim() : itemId;
    const year = typeof hit.year === "string" && hit.year.length > 0 ? ` (${hit.year})` : "";
    const creators = typeof hit.creatorSummary === "string" && hit.creatorSummary.length > 0 ? `${hit.creatorSummary} — ` : "";
    return `${creators}${title}${year}`;
  }

  function renderResults() {
    results.replaceChildren();
    const rows = citationPickerRows(shown, { draft, citedItemIds });
    for (const row of rows) {
      const item = element(document, "li", row.selected ? "reference-row reference-row-selected" : "reference-row");
      item.dataset.itemId = row.itemId;

      const check = document.createElement("input");
      check.type = "checkbox";
      check.checked = row.selected;
      check.dataset.selectItemId = row.itemId;
      const checkLabel = element(document, "label", "reference-check");
      checkLabel.append(element(document, "span", "sr-only", `Select ${row.title}`));
      checkLabel.append(check);
      item.append(checkLabel);

      const body = element(document, "div", "reference-body");
      body.append(element(document, "p", "reference-title", row.title));
      if (row.creatorSummary.length > 0) {
        body.append(element(document, "p", "reference-creators", row.creatorSummary));
      }
      // Where it appeared and when — the two things a writer scans a list by. The CSL type code
      // the row also carries is not a word anybody uses; it is in the details, named properly.
      const publication = [row.containerTitle, row.year].filter((part) => typeof part === "string" && part.length > 0).join(" · ");
      if (publication.length > 0) {
        body.append(element(document, "p", "reference-publication", publication));
      }
      if (row.citedInDocument) {
        // §9.2 prioritises what the document already cites; this says why a row is near the top.
        body.append(element(document, "p", "reference-badge", "Already cited in this document"));
      }
      if (row.selectionCount > 1) {
        body.append(
          element(document, "p", "reference-badge", `In this citation ${row.selectionCount} times`),
        );
      }

      const details = element(document, "dl", "reference-details");
      details.hidden = true;
      for (const [term, value] of [
        ["Title", row.title],
        ["Authors", row.creatorSummary],
        ["Type", itemTypeLabel(row.itemType)],
        ["Journal or publication", row.containerTitle],
        ["Year", row.year],
        ["DOI", row.doi],
      ]) {
        if (value === null || value === "") continue;
        details.append(element(document, "dt", undefined, term), element(document, "dd", undefined, value));
      }

      const disclose = element(document, "button", "reference-disclose", "i");
      disclose.type = "button";
      disclose.setAttribute("aria-label", "Reference details");
      disclose.dataset.discloseItemId = row.itemId;
      disclose.setAttribute("aria-expanded", "false");
      disclose.addEventListener("click", () => {
        const open = details.hidden;
        details.hidden = !open;
        disclose.setAttribute("aria-expanded", open ? "true" : "false");
      });
      body.append(disclose, details);

      item.append(body);
      results.append(item);
    }
    setStatus(citationPickerSummary(rows.length, selectedItemIds(draft).length));
  }

  function renderNotices() {
    notices.replaceChildren();
    for (const notice of draftNotices(draft)) {
      notices.append(element(document, "li", "notice", notice.message));
    }
  }

  function renderDraft() {
    const selectedCount = selectedItemIds(draft).length;
    const currentEntryIds = new Set(draft.entries.map((entry) => entry.entryId));
    for (const entryId of expandedEntries) {
      if (!currentEntryIds.has(entryId)) expandedEntries.delete(entryId);
    }
    count.textContent = `${draft.entries.length} in this citation`;
    entries.replaceChildren();

    for (const [index, entry] of draft.entries.entries()) {
      const item = element(document, "li", "draft-entry");
      item.dataset.entryId = entry.entryId;
      const options = element(document, "details", "draft-options");
      options.open = expandedEntries.has(entry.entryId);
      // The label is its own element so it can be clamped to two lines; a <summary> is not a box
      // a line clamp applies to cleanly.
      const summary = element(document, "summary", "draft-entry-title");
      summary.append(element(document, "span", "draft-entry-label", referenceLabel(entry.itemId)));
      options.append(summary);
      options.addEventListener("toggle", () => {
        if (options.isConnected === false) return;
        if (options.open) expandedEntries.add(entry.entryId);
        else expandedEntries.delete(entry.entryId);
      });
      item.append(options);

      const controls = element(document, "div", "draft-entry-controls");

      const authorLabel = element(document, "label", "draft-field");
      authorLabel.append(element(document, "span", undefined, "Author"));
      const author = document.createElement("select");
      author.dataset.authorModeEntryId = entry.entryId;
      for (const [value, label] of AUTHOR_MODE_LABELS) {
        const option = document.createElement("option");
        option.value = value;
        option.textContent = label;
        option.selected = entry.authorMode === value;
        author.append(option);
      }
      author.addEventListener("change", () => apply(setAuthorMode(draft, entry.entryId, author.value)));
      authorLabel.append(author);

      const locatorLabel = element(document, "label", "draft-field");
      locatorLabel.append(element(document, "span", undefined, "Locator"));
      const label = document.createElement("select");
      label.dataset.locatorLabelEntryId = entry.entryId;
      for (const candidate of DRAFT_LOCATOR_LABELS) {
        const option = document.createElement("option");
        option.value = candidate;
        option.textContent = candidate;
        option.selected = (entry.locator?.label ?? "page") === candidate;
        label.append(option);
      }
      const value = document.createElement("input");
      value.type = "text";
      value.dataset.locatorValueEntryId = entry.entryId;
      value.value = entry.locator?.value ?? "";
      const applyLocator = () => {
        // A blank box clears the locator; a value with no label cannot happen, because the label
        // select always holds one. The draft refuses the combinations this cannot produce.
        apply(
          value.value.trim().length === 0
            ? setLocator(draft, entry.entryId, null)
            : setLocator(draft, entry.entryId, { label: label.value, value: value.value }),
        );
      };
      value.addEventListener("change", applyLocator);
      label.addEventListener("change", applyLocator);
      locatorLabel.append(label, value);

      const affixes = [];
      for (const which of ["prefix", "suffix"]) {
        const affixLabel = element(document, "label", "draft-field");
        affixLabel.append(element(document, "span", undefined, which === "prefix" ? "Prefix" : "Suffix"));
        const input = document.createElement("input");
        input.type = "text";
        input.dataset.affixEntryId = entry.entryId;
        input.dataset.affix = which;
        input.value = entry[which] ?? "";
        // Exactly as typed: the processor supplies the space between an affix and a cite.
        input.addEventListener("change", () => apply(setAffix(draft, entry.entryId, which, input.value)));
        affixLabel.append(input);
        affixes.push(affixLabel);
      }

      controls.append(authorLabel, locatorLabel, ...affixes);
      options.append(controls);

      const actions = element(document, "div", "draft-entry-actions");
      const up = element(document, "button", "button button-small", "Move up");
      up.type = "button";
      up.dataset.moveUpEntryId = entry.entryId;
      up.disabled = index === 0;
      up.addEventListener("click", () => apply(moveEntry(draft, entry.entryId, index - 1)));
      const down = element(document, "button", "button button-small", "Move down");
      down.type = "button";
      down.dataset.moveDownEntryId = entry.entryId;
      down.disabled = index === draft.entries.length - 1;
      down.addEventListener("click", () => apply(moveEntry(draft, entry.entryId, index + 1)));
      const remove = element(document, "button", "button button-small", "Remove");
      remove.type = "button";
      remove.dataset.removeEntryId = entry.entryId;
      remove.addEventListener("click", () => apply(removeEntry(draft, entry.entryId)));
      actions.append(up, down, remove);
      options.append(actions);
      entries.append(item);
    }

    renderNotices();
    clear.disabled = draft.entries.length === 0;
    insert.disabled = inserting || freshPreview === null || draft.entries.length === 0 || typeof ports.insertCitation !== "function";
    if (draft.entries.length > 0 && typeof ports.insertCitation !== "function") {
      // Stated, not discovered by clicking. The picker assembles and previews without a write
      // path; the write is a separate composition and its absence is not hidden here.
      insert.title = "Inserting into the document is not available in this build.";
    }
    return selectedCount;
  }

  /**
   * Apply a draft outcome: adopt the new draft and re-render, or show the refusal and change
   * nothing. The DOM is rebuilt from the draft that came back, so a refused change cannot leave a
   * control showing a value the draft does not hold.
   */
  function apply(outcome) {
    if (inserting) return false;
    if (!outcome.ok) {
      // Re-render **first**, then say why nothing changed. `renderResults` ends by writing the
      // list summary into the same status line, so setting the refusal before it would wipe the
      // sentence in the same tick and leave the user with a control that silently did nothing.
      renderResults();
      renderDraft();
      setStatus(citationPickerRefusalCopy(outcome.refusal.code, outcome.refusal.message));
      return false;
    }
    draft = outcome.draft;
    renderResults();
    renderDraft();
    requestPreview();
    return true;
  }

  function markPreviewStale() {
    if (draft.entries.length === 0) {
      preview.textContent = "Select a reference to preview the citation.";
      preview.className = "preview-text";
      previewWarnings.replaceChildren();
      return;
    }
    preview.className = "preview-text preview-text-stale";
  }

  /**
   * Ask the bridge for §9.2's "preview the formatted result before insertion".
   *
   * The answer is the CSL processor's, shown verbatim. A response for a draft the user has since
   * changed is dropped rather than rendered: a preview that lags the selection is a preview that
   * says a different citation is about to be inserted than the one that is.
   */
  function requestPreview() {
    previewSequence += 1;
    freshPreview = null;
    insert.disabled = true;
    markPreviewStale();
    if (draft.entries.length === 0) return;
    const built = citationRequest(draft, style);
    if (!built.ok) {
      preview.className = "preview-text";
      preview.textContent = citationPickerRefusalCopy(built.refusal.code, built.refusal.message);
      return;
    }
    if (typeof ports.previewCitation !== "function") {
      preview.className = "preview-text";
      preview.textContent = "Connect the desktop app to preview this citation.";
      return;
    }
    previewSequence += 1;
    const sequence = previewSequence;
    Promise.resolve(ports.previewCitation(built.request))
      .then((rendered) => {
        if (sequence !== previewSequence) return;
        preview.className = "preview-text";
        preview.textContent = rendered.text;
        freshPreview = rendered.text;
        renderDraft();
        previewWarnings.replaceChildren();
        for (const warning of rendered.warnings ?? []) {
          const copy = citationPreviewWarningCopy(warning);
          previewWarnings.append(element(document, "li", copy.quiet ? "notice notice-quiet" : "notice", copy.text));
        }
      })
      .catch((error) => {
        if (sequence !== previewSequence) return;
        preview.className = "preview-text";
        preview.textContent = "The preview could not be rendered.";
        dispatchDetail(document, form, "refmgr:citation-preview-failed", {
          reason: error && error.message ? String(error.message) : "unknown",
        });
      });
  }

  function presentResults(page) {
    hitsById.clear();
    shown = [];
    for (const hit of page.results) {
      hitsById.set(hit.itemId, hit);
      shown.push(hit);
    }
    renderResults();
    // The draft may cite a reference that this page does not show; its label comes from whatever
    // page it was found on, so re-rendering keeps a selected entry readable.
    renderDraft();
  }

  function presentScopes(collections) {
    const options = citationScopeOptions(collections);
    const previous = scope.value;
    scope.replaceChildren();
    const all = document.createElement("option");
    all.value = LIBRARY_SCOPE_ID;
    all.textContent = ALL_REFERENCES_LABEL;
    scope.append(all);
    for (const option of options) {
      const node = document.createElement("option");
      node.value = option.collectionId;
      // Indentation is text, not a computed style or class: `style-src 'self'` would refuse the
      // first and ADR-0086 constrains the second.
      const indent = "  ".repeat(option.depth);
      node.textContent = option.detached ? `${indent}${option.name} (unfiled)` : `${indent}${option.name}`;
      scope.append(node);
    }
    // A scope that survived the reload stays chosen; one that did not falls back to the library.
    scope.value = options.some((option) => option.collectionId === previous) ? previous : LIBRARY_SCOPE_ID;
  }

  /**
   * Run the query the controls currently describe.
   *
   * A blank term is `/recent`, never `/search` with an empty `q`: the bridge refuses that, and it
   * refuses it precisely so an empty box can never mean "the whole library" by accident.
   */
  /** Rising with every query, so a slow answer to an older term can never replace a newer one. */
  let querySequence = 0;
  let debounce = null;

  function runQuery() {
    if (debounce !== null) {
      globalThis.clearTimeout(debounce);
      debounce = null;
    }
    querySequence += 1;
    const sequence = querySequence;
    const current = (page) => {
      if (sequence === querySequence) presentResults(page);
    };
    const failed = (message) => () => {
      if (sequence === querySequence) setStatus(message);
    };
    const term = query.value.trim();
    const collection = scope.value;
    dispatchDetail(document, form, "refmgr:citation-search", { scoped: collection.length > 0 });
    if (term.length === 0) {
      if (typeof ports.recent !== "function") {
        setStatus("Connect the desktop app to list your library.");
        return Promise.resolve();
      }
      setStatus("Loading recent references…");
      return Promise.resolve(ports.recent({ limit: PICKER_PAGE_SIZE }))
        .then(current)
        .catch(failed("Your library could not be reached. Check the desktop app connection."));
    }
    if (typeof ports.search !== "function") {
      setStatus("Connect the desktop app to search your library.");
      return Promise.resolve();
    }
    setStatus("Searching…");
    return Promise.resolve(
      ports.search({
        q: term,
        limit: PICKER_PAGE_SIZE,
        // §9.2 "Prioritise references already cited in the document."
        prefer: citedItemIds,
        ...(collection === LIBRARY_SCOPE_ID ? {} : { collection }),
      }),
    )
      .then(current)
      .catch(failed("That search could not be run. Check the desktop app connection."));
  }

  form.addEventListener("submit", (event) => {
    event.preventDefault();
    void runQuery();
  });
  scope.addEventListener("change", () => {
    void runQuery();
  });
  // Search as the writer types, once they pause. Enter (the form's submit) still searches at once.
  query.addEventListener("input", () => {
    if (debounce !== null) globalThis.clearTimeout(debounce);
    debounce = globalThis.setTimeout(() => {
      debounce = null;
      void runQuery();
    }, SEARCH_DEBOUNCE_MS);
  });

  results.addEventListener("change", (event) => {
    const target = event && event.target;
    const itemId = target && target.dataset ? target.dataset.selectItemId : undefined;
    if (typeof itemId !== "string" || itemId.length === 0) return;
    apply(toggleCitationSelection(draft, itemId, target.checked === true));
  });

  clear.addEventListener("click", () => {
    if (inserting) return;
    draft = emptyDraft();
    renderResults();
    renderDraft();
    requestPreview();
    setStatus("This citation was cleared. Nothing in your document changed.");
  });

  insert.addEventListener("click", () => {
    if (inserting || freshPreview === null) return;
    const built = citationRequest(draft, style);
    if (!built.ok) {
      setStatus(citationPickerRefusalCopy(built.refusal.code, built.refusal.message));
      return;
    }
    if (typeof ports.insertCitation !== "function") {
      setStatus("Inserting into the document is not available in this build.");
      return;
    }
    inserting = true;
    renderDraft();
    setStatus("Inserting…");
    Promise.resolve().then(() => ports.insertCitation({ draft, request: built.request, previewText: freshPreview }))
      .then((outcome) => {
        draft = emptyDraft();
        renderResults();
        renderDraft();
        requestPreview();
        setStatus(outcome?.message ?? "Citation inserted.");
      })
      .catch((error) => setStatus(error.message || "The citation could not be inserted. The draft has been kept."))
      .finally(() => { inserting = false; renderDraft(); });
  });

  renderResults();
  renderDraft();

  return Object.freeze({
    /** The opening state: the scope selector filled, then §9.2's recent references. */
    async open() {
      if (typeof ports.collections === "function") {
        try {
          presentScopes((await ports.collections()).collections);
        } catch {
          // A scope selector that could not be filled still offers the whole library, which is
          // what an unscoped search already does. Refusing to list anything would be worse.
          setStatus("Collections could not be listed. Searching the whole library.");
        }
      }
      await runQuery();
    },
    /** The item ids the open document already cites (§9.2's prioritisation). */
    setCitedItemIds(ids) {
      citedItemIds = [...new Set((Array.isArray(ids) ? ids : []).filter((id) => typeof id === "string" && id.length > 0))];
      renderResults();
    },
    /** The style and locale every preview and insertion is rendered in. */
    setStyle(next) {
      style = { styleId: String(next?.styleId ?? ""), locale: String(next?.locale ?? "") };
      requestPreview();
    },
    presentResults,
    presentScopes,
    setStatus,
    draft: () => draft,
    runQuery,
  });
}
