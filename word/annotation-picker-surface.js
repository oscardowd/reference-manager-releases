/**
 * Task-pane presentation for Insert PDF annotation (SPEC §9.1; ADR-0176; E14-04.4).
 *
 * Policy stays in `@refmgr/word-addin` (`resolveAnnotationPickerSelection`). This unbundled module
 * searches, displays, collects an explicit page label when needed, and requires every returned
 * omission notice to be acknowledged before the composition may commit. It never invents a page
 * label from `pageIndex`, never puts excerpt text in DOM attributes, and never logs query/quote/
 * comment (§23).
 */

const REFUSAL_COPY = Object.freeze({
  "reference-unavailable": "This annotation is not linked to a live reference, so it cannot be inserted.",
  "page-label-required": "Enter the printed page label. The physical page index is not used.",
  "empty-excerpt": "This annotation has neither a quote nor a comment.",
  "invalid-annotation-id": "The selected annotation could not be used.",
  "invalid-item-id": "The selected annotation's reference could not be used.",
  "draft-failed": "The citation draft could not be built from this annotation.",
});

const DISPLAY_EXCERPT_MAX = 120;

function text(value) {
  return typeof value === "string" ? value : "";
}

function truncateForDisplay(value) {
  const raw = text(value);
  if (raw.length <= DISPLAY_EXCERPT_MAX) return raw;
  return `${raw.slice(0, DISPLAY_EXCERPT_MAX - 1)}…`;
}

function hasStoredPageLabel(hit) {
  return typeof hit.pageLabel === "string" && hit.pageLabel.trim().length > 0;
}

/**
 * Pure view-model for one search hit. Selection availability is declared here so the DOM cannot
 * offer a button for an unlinked row; the domain still re-checks on commit.
 */
export function annotationPickerRowPresentation(hit) {
  if (!hit || typeof hit.annotationId !== "string") throw new Error("InvalidAnnotationPickerHit");
  const quote = truncateForDisplay(hit.quoteText);
  const comment = truncateForDisplay(hit.commentText);
  const excerpt =
    quote.length > 0 ? quote : comment.length > 0 ? comment : "(no excerpt text)";
  const referenceLabel =
    hit.reference === null
      ? "Not linked to a reference"
      : [text(hit.reference.creatorSummary), text(hit.reference.year), text(hit.reference.title)]
          .filter((part) => part.length > 0)
          .join(" · ") || text(hit.reference.itemId);
  const pageLabel = hasStoredPageLabel(hit) ? hit.pageLabel.trim() : null;
  return Object.freeze({
    annotationId: hit.annotationId,
    selectable: hit.reference !== null,
    needsPageLabel: hit.reference !== null && pageLabel === null,
    pageLabel,
    pageIndex: hit.pageIndex,
    excerpt,
    referenceLabel,
    // Full fields stay in memory for the composition event — not written into attributes.
    hit,
  });
}

export function annotationPickerRefusalCopy(code) {
  const message = REFUSAL_COPY[code];
  if (typeof message !== "string") throw new Error("InvalidAnnotationPickerRefusal");
  return message;
}

/**
 * Build the acknowledgement only when the DOM shows exactly the notices the selection returned
 * and every displayed checkbox is checked — same contract as the unlink surface (ADR-0119 shape).
 */
export function acknowledgeAnnotationOmissions(notices, displayedFacets, checkedFacets, acceptedAt) {
  const expected = (Array.isArray(notices) ? notices : []).map((entry) => text(entry && entry.facet));
  if (expected.some((facet) => facet.length === 0) || new Set(expected).size !== expected.length) {
    return Object.freeze({ ok: false, reason: "invalid-notices" });
  }
  const displayed = [...displayedFacets].map(text);
  const checked = new Set([...checkedFacets].map(text));
  const sameDisplay =
    displayed.length === expected.length && expected.every((facet, index) => displayed[index] === facet);
  const allChecked = checked.size === expected.length && expected.every((facet) => checked.has(facet));
  if (!sameDisplay || !allChecked) {
    return Object.freeze({ ok: false, reason: "omissions-not-acknowledged" });
  }
  const instant = text(acceptedAt);
  if (Number.isNaN(Date.parse(instant))) {
    return Object.freeze({ ok: false, reason: "invalid-accepted-at" });
  }
  return Object.freeze({
    ok: true,
    acknowledgement: Object.freeze({
      acceptedFacets: Object.freeze([...expected]),
      acceptedAt: instant,
    }),
  });
}

function appendText(document, parent, name, value, className) {
  const element = document.createElement(name);
  element.textContent = value;
  if (className !== undefined) element.className = className;
  parent.append(element);
  return element;
}

function dispatchDetail(document, target, name, detail) {
  const event = document.createEvent("CustomEvent");
  event.initCustomEvent(name, true, false, detail);
  target.dispatchEvent(event);
}

export function mountAnnotationPickerSurface(document) {
  const form = document.getElementById("annotation-search-form");
  const query = document.getElementById("annotation-search-query");
  const results = document.getElementById("annotation-search-results");
  const status = document.getElementById("annotation-picker-status");
  const pageLabelField = document.getElementById("annotation-page-label");
  const pageLabelRow = document.getElementById("annotation-page-label-row");
  const omissions = document.getElementById("annotation-omissions");
  const insert = document.getElementById("annotation-insert");
  if (!form || !query || !results || !status || !pageLabelField || !pageLabelRow || !omissions || !insert) {
    throw new Error("AnnotationPickerSurfaceMissing");
  }

  /** @type {Map<string, object>} */
  const hitsById = new Map();
  let selectedId = null;
  let selection = null;

  function setStatus(message) {
    status.textContent = message;
  }

  function clearSelectionUi() {
    selectedId = null;
    selection = null;
    omissions.replaceChildren();
    insert.disabled = true;
    pageLabelRow.hidden = true;
    pageLabelField.value = "";
  }

  function presentRefusal(refusal) {
    selection = null;
    omissions.replaceChildren();
    insert.disabled = true;
    setStatus(annotationPickerRefusalCopy(refusal.code));
  }

  function presentSelection(next) {
    selection = next;
    omissions.replaceChildren();
    for (const notice of next.omissionNotices) {
      const item = document.createElement("li");
      const label = document.createElement("label");
      const checkbox = document.createElement("input");
      checkbox.type = "checkbox";
      checkbox.dataset.omissionFacet = notice.facet;
      label.append(checkbox, document.createTextNode(` ${notice.label}`));
      item.append(label);
      omissions.append(item);
    }
    insert.disabled = next.omissionNotices.length > 0;
    setStatus(
      next.omissionNotices.length === 0
        ? "Ready to insert."
        : "Acknowledge every limitation before inserting.",
    );
  }

  function selectedHit() {
    return selectedId === null ? null : hitsById.get(selectedId) ?? null;
  }

  function requestSelectionResolve() {
    const hit = selectedHit();
    if (hit === null) return;
    const explicit = pageLabelField.value;
    dispatchDetail(document, form, "refmgr:annotation-select", {
      hit,
      explicitPageLabel: explicit.trim().length === 0 ? undefined : explicit,
    });
  }

  function renderResults(page) {
    hitsById.clear();
    clearSelectionUi();
    results.replaceChildren();
    for (const hit of page.results) {
      hitsById.set(hit.annotationId, hit);
      const row = annotationPickerRowPresentation(hit);
      const item = document.createElement("li");
      const button = document.createElement("button");
      button.type = "button";
      button.disabled = !row.selectable;
      button.dataset.annotationId = row.annotationId;
      // Display text only — full excerpt stays in hitsById, never in attributes.
      button.textContent = `${row.referenceLabel} · ${
        row.pageLabel === null ? "page label needed" : `p. ${row.pageLabel}`
      } — ${row.excerpt}`;
      button.addEventListener("click", () => {
        selectedId = row.annotationId;
        pageLabelRow.hidden = !row.needsPageLabel;
        if (!row.needsPageLabel) pageLabelField.value = "";
        requestSelectionResolve();
      });
      item.append(button);
      if (!row.selectable) {
        appendText(document, item, "p", annotationPickerRefusalCopy("reference-unavailable"), "note");
      }
      results.append(item);
    }
    const countLabel = page.totalExact ? String(page.total) : `${page.total}+`;
    setStatus(
      page.results.length === 0
        ? "No annotations matched."
        : `Showing ${page.results.length} of ${countLabel}.`,
    );
  }

  form.addEventListener("submit", (event) => {
    event.preventDefault();
    clearSelectionUi();
    dispatchDetail(document, form, "refmgr:annotation-search", {
      q: query.value.trim().length === 0 ? undefined : query.value,
    });
  });

  pageLabelField.addEventListener("input", () => {
    if (selectedId !== null) requestSelectionResolve();
  });

  omissions.addEventListener("change", () => {
    if (selection === null) return;
    const displayed = [...omissions.querySelectorAll("input[data-omission-facet]")].map(
      (input) => input.dataset.omissionFacet,
    );
    const checked = [...omissions.querySelectorAll("input[data-omission-facet]:checked")].map(
      (input) => input.dataset.omissionFacet,
    );
    const ack = acknowledgeAnnotationOmissions(
      selection.omissionNotices,
      displayed,
      checked,
      new Date().toISOString(),
    );
    insert.disabled = !ack.ok;
  });

  insert.addEventListener("click", () => {
    if (selection === null) return;
    const displayed = [...omissions.querySelectorAll("input[data-omission-facet]")].map(
      (input) => input.dataset.omissionFacet,
    );
    const checked = [...omissions.querySelectorAll("input[data-omission-facet]:checked")].map(
      (input) => input.dataset.omissionFacet,
    );
    const ack = acknowledgeAnnotationOmissions(
      selection.omissionNotices,
      displayed,
      checked,
      new Date().toISOString(),
    );
    if (!ack.ok) {
      setStatus("Acknowledge every limitation before inserting.");
      return;
    }
    dispatchDetail(document, form, "refmgr:annotation-insert", {
      plan: selection.plan,
      acknowledgement: ack.acknowledgement,
    });
  });

  return Object.freeze({
    presentSearchResults: renderResults,
    presentSelection,
    presentRefusal,
    setStatus,
    clearSelection: clearSelectionUi,
  });
}
