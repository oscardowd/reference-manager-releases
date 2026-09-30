/**
 * Task-pane presentation for Edit citation (SPEC §9.2, §9.3; ADR-0240, ADR-0241; E10-05.1).
 *
 * The other half of §9.2's sentence. E10-04.5 gave the pane a way to assemble a citation;
 * `openCitationForEdit`/`prepareCitationEdit` have been able to change one since E10-05 and had no
 * surface. This is it: select a citation in Word, see what it currently says, change how it is
 * displayed, its locator and its affixes, preview the result, and save it back in place.
 *
 * Policy is `citation-draft.js`, held equal to `@refmgr/word-addin`'s by
 * `apps/word-taskpane/src/citation-draft.test.ts`. Reading and writing the document belong to the
 * desktop runtime (ADR-0241) — this module parses no payload and verifies no fingerprint.
 *
 * Four rules it keeps, three shared with the picker and one its own:
 *
 * - **No library string ever enters an attribute or a class name**, and no citation text is
 *   logged (§23, ADR-0084's rule).
 * - **The preview is the processor's answer, shown verbatim** (invariant 4). It opens on what
 *   Word currently shows — not on a re-render, which would present the user with a citation that
 *   differs from their document and call it the current one.
 * - **A refusal is shown, never swallowed.**
 * - **A manual edit is surfaced, and saving over it requires an explicit acknowledgement** (§9.3).
 *   The user typed that text; replacing it silently is the one thing this panel must not do.
 */

import {
  DRAFT_LOCATOR_LABELS,
  citationRequest,
  draftNotices,
  setAffix,
  setAuthorMode,
  setLocator,
} from "./citation-draft.js";
import { citationPreviewWarningCopy } from "./citation-picker-surface.js";

/** "Display as", in the vocabulary a writer uses rather than CSL's. */
export const DISPLAY_MODES = Object.freeze([
  ["normal", "Author (Year)"],
  ["suppress-author", "(Year) only — author is already in the sentence"],
  ["author-only", "Author only"],
]);

/**
 * Copy for the notices the domain can return when a citation is opened.
 *
 * Each says what it means for the document rather than what it means for the payload: "custom XML
 * part 3 is redundant" is true and tells a writer nothing they can act on.
 */
const NOTICE_COPY = Object.freeze({
  "manual-edit-detected": "This citation has been edited by hand in the document. Saving will replace that wording.",
  "stale-custom-xml": "This document carries leftover citation data. Saving tidies it.",
  "unreadable-custom-xml": "Part of this document's citation data was written by a newer version and is being left alone.",
});

export function citationEditNoticeCopy(code, fallback) {
  const message = NOTICE_COPY[code];
  if (typeof message === "string") return message;
  return typeof fallback === "string" && fallback.length > 0 ? fallback : "";
}

/**
 * The line that says which citation is being edited.
 *
 * Built from the references the runtime resolved, and falling back to the document's own record
 * when the library no longer holds one — a citation whose reference was deleted still has to be
 * nameable, because it is still in the manuscript.
 */
export function editingLabel(references) {
  const parts = (Array.isArray(references) ? references : []).map((entry) => {
    const reference = entry && entry.reference;
    if (!reference) return "a reference no longer in your library";
    const creators = typeof reference.creatorSummary === "string" ? reference.creatorSummary.trim() : "";
    const year = typeof reference.year === "string" && reference.year.length > 0 ? ` ${reference.year}` : "";
    const title = typeof reference.title === "string" ? reference.title.trim() : "";
    if (creators.length > 0) return `${creators}${year}`;
    return title.length > 0 ? title : "(untitled reference)";
  });
  if (parts.length === 0) return "";
  if (parts.length === 1) return parts[0];
  if (parts.length === 2) return `${parts[0]} and ${parts[1]}`;
  return `${parts[0]} and ${parts.length - 1} others`;
}

function element(document, name, className, textContent) {
  const node = document.createElement(name);
  if (className !== undefined) node.className = className;
  if (textContent !== undefined) node.textContent = textContent;
  return node;
}

/**
 * Mount the panel.
 *
 * `ports.openEdit` reads the selected citation and `ports.saveEdit` writes it; both are absent
 * until the desktop is connected, and an absent one is **stated rather than discovered by
 * clicking** — the rule the picker keeps for the same reason.
 */
export function mountCitationEditSurface(document, ports = {}) {
  const panel = document.getElementById("citation-edit");
  const open = document.getElementById("citation-edit-open");
  const heading = document.getElementById("citation-edit-target");
  const status = document.getElementById("citation-edit-status");
  const preview = document.getElementById("citation-edit-preview");
  const notices = document.getElementById("citation-edit-notices");
  const display = document.getElementById("citation-edit-display");
  const locatorLabel = document.getElementById("citation-edit-locator-label");
  const locatorValue = document.getElementById("citation-edit-locator-value");
  const prefix = document.getElementById("citation-edit-prefix");
  const suffix = document.getElementById("citation-edit-suffix");
  const acceptRow = document.getElementById("citation-edit-accept-row");
  const accept = document.getElementById("citation-edit-accept");
  const save = document.getElementById("citation-edit-save");
  const cancel = document.getElementById("citation-edit-cancel");
  const form = document.getElementById("citation-edit-form");
  if (!panel || !open || !heading || !status || !preview || !notices || !display || !locatorLabel || !locatorValue || !prefix || !suffix || !acceptRow || !accept || !save || !cancel || !form) {
    throw new Error("CitationEditSurfaceMissing");
  }

  /** The session the runtime returned, or null when no citation is open. */
  let session = null;
  let draft = null;
  let style = { styleId: "", locale: "" };
  /** Rising, so a slow preview for an older edit cannot overwrite a newer one. */
  let previewSequence = 0;

  for (const [value, label] of DISPLAY_MODES) {
    const option = document.createElement("option");
    option.value = value;
    option.textContent = label;
    display.append(option);
  }
  for (const label of DRAFT_LOCATOR_LABELS) {
    const option = document.createElement("option");
    option.value = label;
    option.textContent = label;
    locatorLabel.append(option);
  }

  function setStatus(message) {
    status.textContent = message;
  }

  /**
   * The single entry this panel edits.
   *
   * A grouped citation has several, and **this panel deliberately does not try to edit them all**:
   * the controls below describe one cite, and silently applying them to three would be the pane
   * deciding something the user did not. A grouped citation opens read-only with its own sentence,
   * and the picker is where a cluster is assembled.
   */
  function soleEntry() {
    return draft !== null && draft.entries.length === 1 ? draft.entries[0] : null;
  }

  function setControlsEnabled(enabled) {
    for (const control of [display, locatorLabel, locatorValue, prefix, suffix]) control.disabled = !enabled;
    save.disabled = !enabled;
  }

  function closePanel() {
    session = null;
    draft = null;
    previewFresh = false;
    previewNotes = [];
    previewSequence += 1;
    panel.hidden = true;
    notices.replaceChildren();
    acceptRow.hidden = true;
    accept.checked = false;
    heading.textContent = "";
    preview.textContent = "";
    preview.className = "preview-text";
    setControlsEnabled(false);
  }

  function renderControls() {
    const entry = soleEntry();
    if (entry === null) return;
    display.value = entry.authorMode;
    locatorLabel.value = entry.locator === undefined ? "page" : entry.locator.label;
    locatorValue.value = entry.locator === undefined ? "" : entry.locator.value;
    prefix.value = entry.prefix === undefined ? "" : entry.prefix;
    suffix.value = entry.suffix === undefined ? "" : entry.suffix;
  }

  function renderNotices(list) {
    notices.replaceChildren();
    for (const notice of list) {
      const copy = citationEditNoticeCopy(notice.code, notice.message);
      if (copy.length === 0) continue;
      notices.append(element(document, "li", "notice", copy));
    }
    for (const note of previewNotes) notices.append(element(document, "li", "notice notice-quiet", note));
  }

  /**
   * The text in the preview box is the processor's answer for the draft as it now stands — not
   * the citation as the document showed it when the panel opened, and not an answer to an earlier
   * edit still in flight. Save sends that text, and the runtime refuses any other, so Save waits.
   */
  let previewFresh = false;
  /** Notes about the preview itself (e.g. "numbered by position"), shown after the session's. */
  let previewNotes = [];

  function saveEnabled() {
    if (draft === null || soleEntry() === null || !previewFresh) return false;
    // §9.3: a manual edit is the user's own writing. Nothing replaces it until they say so.
    if (session !== null && session.manualEditDetected && accept.checked !== true) return false;
    return true;
  }

  function refreshSaveState() {
    save.disabled = !saveEnabled();
  }

  function requestPreview() {
    if (draft === null) return;
    const built = citationRequest(draft, style);
    if (!built.ok) {
      preview.className = "preview-text";
      preview.textContent = built.refusal.message;
      return;
    }
    if (typeof ports.previewCitation !== "function") return;
    preview.className = "preview-text preview-text-stale";
    previewFresh = false;
    refreshSaveState();
    previewSequence += 1;
    const sequence = previewSequence;
    Promise.resolve(ports.previewCitation(built.request))
      .then((rendered) => {
        if (sequence !== previewSequence) return;
        preview.className = "preview-text";
        preview.textContent = rendered.text;
        previewNotes = (rendered.warnings ?? []).map((warning) => citationPreviewWarningCopy(warning).text);
        renderNotices([...(session?.notices ?? []), ...draftNotices(draft)]);
        previewFresh = true;
        refreshSaveState();
      })
      .catch(() => {
        if (sequence !== previewSequence) return;
        preview.className = "preview-text";
        preview.textContent = "The preview could not be rendered.";
      });
  }

  /** Adopt a draft outcome, or show why nothing changed and leave the controls as the draft has them. */
  function apply(outcome) {
    if (!outcome.ok) {
      renderControls();
      setStatus(outcome.refusal.message);
      return false;
    }
    draft = outcome.draft;
    renderControls();
    renderNotices([...(session?.notices ?? []), ...draftNotices(draft)]);
    refreshSaveState();
    requestPreview();
    setStatus("");
    return true;
  }

  function applyLocator() {
    const entry = soleEntry();
    if (entry === null) return;
    apply(
      locatorValue.value.trim().length === 0
        ? setLocator(draft, entry.entryId, null)
        : setLocator(draft, entry.entryId, { label: locatorLabel.value, value: locatorValue.value }),
    );
  }

  display.addEventListener("change", () => {
    const entry = soleEntry();
    if (entry !== null) apply(setAuthorMode(draft, entry.entryId, display.value));
  });
  locatorLabel.addEventListener("change", applyLocator);
  locatorValue.addEventListener("change", applyLocator);
  for (const [control, which] of [[prefix, "prefix"], [suffix, "suffix"]]) {
    control.addEventListener("change", () => {
      const entry = soleEntry();
      // Exactly as typed: the processor supplies the space between an affix and a cite.
      if (entry !== null) apply(setAffix(draft, entry.entryId, which, control.value));
    });
  }
  // Choosing to replace hand-typed wording is choosing the processor's text: show it before Save.
  accept.addEventListener("change", () => {
    if (accept.checked === true && !previewFresh) requestPreview();
    refreshSaveState();
  });
  form.addEventListener("submit", (event) => {
    event.preventDefault();
  });

  async function openSelected() {
    if (typeof ports.openEdit !== "function") {
      panel.hidden = false;
      setStatus("Connect the desktop app to edit a citation.");
      setControlsEnabled(false);
      return;
    }
    panel.hidden = false;
    setStatus("Reading the selected citation…");
    setControlsEnabled(false);
    try {
      const opened = await ports.openEdit();
      session = opened;
      draft = opened.draft;
      style = { styleId: String(opened.styleId ?? ""), locale: String(opened.locale ?? "") };
      heading.textContent = editingLabel(opened.references);
      // What Word shows right now, before anything is changed.
      preview.className = "preview-text";
      preview.textContent = String(opened.visibleText ?? "");
      renderNotices([...(opened.notices ?? []), ...draftNotices(draft)]);
      acceptRow.hidden = opened.manualEditDetected !== true;
      accept.checked = false;

      if (soleEntry() === null) {
        // A grouped citation. Readable, nameable, and not editable from these five controls —
        // see `soleEntry`.
        setControlsEnabled(false);
        setStatus(
          `This citation cites ${draft.entries.length} references. Editing a grouped citation one reference at a time is not available in this build.`,
        );
        return;
      }
      renderControls();
      setControlsEnabled(true);
      refreshSaveState();
      setStatus(
        opened.manualEditDetected === true
          ? "This citation was edited by hand. Tick the box below to let Reference Manager replace that wording."
          : "",
      );
    } catch (error) {
      closePanel();
      panel.hidden = false;
      setStatus(error && error.message ? String(error.message) : "That citation could not be opened.");
    }
  }

  open.addEventListener("click", () => {
    void openSelected();
  });
  cancel.addEventListener("click", () => {
    closePanel();
    setStatus("Nothing in your document changed.");
  });

  save.addEventListener("click", () => {
    if (draft === null || !saveEnabled()) return;
    if (typeof ports.saveEdit !== "function") {
      setStatus("Connect the desktop app to save this change.");
      return;
    }
    const built = citationRequest(draft, style);
    if (!built.ok) {
      setStatus(built.refusal.message);
      return;
    }
    setStatus("Saving…");
    save.disabled = true;
    Promise.resolve(
      ports.saveEdit({
        draft,
        // The text the user is looking at is the text that gets written, or nothing is written.
        previewText: preview.textContent,
        acceptManualEditReplacement: accept.checked === true,
      }),
    )
      .then((result) => {
        closePanel();
        setStatus(result && result.message ? result.message : "Citation updated.");
      })
      .catch((error) => {
        refreshSaveState();
        setStatus(error && error.message ? String(error.message) : "That change was not saved. Your document is unchanged.");
      });
  });

  closePanel();

  return Object.freeze({
    openSelected,
    close: closePanel,
    setStatus,
    draft: () => draft,
    session: () => session,
  });
}
