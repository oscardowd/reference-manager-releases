/**
 * The Document tab (SPEC §9.1 Insert bibliography, Refresh document, Change citation style; E10-19.2).
 *
 * Presentation only, like the picker: the ports prepare and write, and this module decides no
 * citation and formats none. It states the document's citations and bibliography, offers the three
 * document-wide actions, and turns the two questions the runtime can ask back into one explicit
 * confirmation each — never an automatic "yes":
 *
 * - `confirm-removed-citations`: citations the document data records but Word no longer shows.
 *   Usually deleted by the writer; possibly unreachable (a text box). Dropping them is theirs to say.
 * - `confirm-bibliography-replacement`: the bibliography was edited by hand (§9.4).
 *
 * A refresh that **kept** hand-edited citations is a success that still has a question in it, so it
 * is reported as done and then offered as a separate, optional replacement.
 *
 * Every sentence here is product copy built from counts. No reference title, citation text or style
 * name the library supplied is ever put in an attribute or a class (ADR-0084).
 */

function plural(count, one, many) {
  return `${String(count)} ${count === 1 ? one : many}`;
}

/** The summary line at the top of the tab. Exported so the wording is testable without a DOM. */
export function documentSummaryText(state) {
  if (!state || state.connected !== true) return "Connect the desktop app to work with this document.";
  if (state.citations === 0) return "No citations in this document yet. Insert one from the Cite tab.";
  const bibliography = state.bibliography ? "a bibliography" : "no bibliography yet";
  return `${plural(state.citations, "citation", "citations")} and ${bibliography}.`;
}

/** What a finished refresh says, from the runtime's summary counts. */
export function refreshOutcomeText(summary) {
  if (!summary) return "Document updated.";
  const parts = [];
  if (summary.styleChanged) parts.push("Style changed");
  if (summary.updatedCitations > 0) parts.push(`${plural(summary.updatedCitations, "citation", "citations")} updated`);
  else if (!summary.styleChanged) parts.push("Citations already up to date");
  if (summary.removedCitations > 0) parts.push(`${plural(summary.removedCitations, "deleted citation", "deleted citations")} removed`);
  if (summary.bibliographyUpdated) parts.push("bibliography updated");
  return `${parts.join(", ")}.`;
}

export function heldManualText(count) {
  return count === 1
    ? "One citation was edited by hand and kept as you wrote it. Replace it with formatted text?"
    : `${String(count)} citations were edited by hand and kept as you wrote them. Replace them with formatted text?`;
}

export function mountDocumentSurface(document, ports = {}) {
  const summary = document.getElementById("document-summary");
  const style = document.getElementById("citation-style");
  const styleNote = document.getElementById("style-card-note");
  const applyStyle = document.getElementById("document-style-apply");
  const bibliography = document.getElementById("document-bibliography");
  const bibliographyNote = document.getElementById("bibliography-card-note");
  const refresh = document.getElementById("document-refresh");
  const confirm = document.getElementById("document-confirm");
  const confirmMessage = document.getElementById("document-confirm-message");
  const confirmAccept = document.getElementById("document-confirm-accept");
  const confirmCancel = document.getElementById("document-confirm-cancel");
  const status = document.getElementById("document-status");
  if (!summary || !style || !styleNote || !applyStyle || !bibliography || !bibliographyNote || !refresh ||
      !confirm || !confirmMessage || !confirmAccept || !confirmCancel || !status) {
    throw new Error("DocumentSurfaceMissing");
  }

  let state = { connected: false, citations: 0, bibliography: false, bibliographyLocked: false, styleId: "", styles: [] };
  let busy = false;
  /** The action to run if the writer confirms; null while nothing is being asked. */
  let pending = null;

  function setStatus(message) {
    status.textContent = message;
  }

  function render() {
    summary.textContent = documentSummaryText(state);
    const idle = state.connected && !busy && pending === null;
    const cited = state.citations > 0;
    style.disabled = !state.connected || busy || pending !== null;
    applyStyle.hidden = !cited;
    applyStyle.disabled = !idle || style.value === state.styleId || style.value === "";
    styleNote.hidden = cited;
    styleNote.textContent = cited ? "" : "New citations will use this style.";
    bibliography.textContent = state.bibliography ? "Update bibliography" : "Insert bibliography";
    bibliography.disabled = !idle || !cited || state.bibliographyLocked;
    bibliographyNote.textContent = state.bibliographyLocked
      ? "This document's bibliography is locked."
      : state.bibliography
        ? "Re-formats the bibliography already in the document."
        : "Inserted at the cursor. Put the cursor where the list should go, usually the end of the document.";
    refresh.disabled = !idle || !cited;
    confirm.hidden = pending === null;
  }

  /** Replace the style list, keeping the document's own style selected. */
  function renderStyles() {
    style.replaceChildren();
    for (const entry of state.styles) {
      const option = document.createElement("option");
      option.value = entry.id;
      option.textContent = entry.title;
      style.append(option);
    }
    style.value = state.styleId;
  }

  function ask(message, run) {
    pending = run;
    confirmMessage.textContent = message;
    render();
    focus(confirmAccept);
  }

  function focus(element) {
    if (element && typeof element.focus === "function") {
      try { element.focus({ preventScroll: true }); } catch { element.focus(); }
    }
  }

  /**
   * Run one document action, turning the runtime's two questions into confirmations. `accepted`
   * accumulates the writer's answers, so confirming one question and then being asked the other
   * sends both answers on the final attempt.
   */
  async function run(action, options, accepted = {}) {
    if (busy) return;
    busy = true;
    pending = null;
    render();
    setStatus(action === "bibliography" ? "Writing the bibliography…" : options.styleId ? "Changing the style…" : "Updating the document…");
    const port = action === "bibliography" ? ports.insertBibliography : ports.refreshDocument;
    try {
      if (typeof port !== "function") throw new Error("Connect the desktop app first.");
      const outcome = await port({ ...options, ...accepted });
      busy = false;
      if (action === "bibliography") {
        setStatus(outcome?.message ?? "Bibliography written.");
      } else {
        const held = outcome?.summary?.heldManualClusterIds ?? [];
        setStatus(outcome?.message ?? refreshOutcomeText(outcome?.summary));
        if (held.length > 0) {
          ask(heldManualText(held.length), () => run("refresh", {}, { acceptManualReplacementIds: held }));
        }
      }
      if (typeof ports.onChanged === "function") await ports.onChanged();
    } catch (error) {
      busy = false;
      const code = error && error.code;
      const details = (error && error.details) || {};
      if (code === "confirm-removed-citations" && Array.isArray(details.removedClusterIds)) {
        setStatus("");
        ask(String(error.message), () => run(action, options, { ...accepted, acceptRemovedClusterIds: details.removedClusterIds }));
        return;
      }
      if (code === "confirm-bibliography-replacement") {
        setStatus("");
        ask(String(error.message), () => run(action, options, { ...accepted, acceptBibliographyReplacement: true }));
        return;
      }
      setStatus(error && error.message ? String(error.message) : "That could not be done. Your document is unchanged.");
    } finally {
      busy = false;
      render();
    }
  }

  style.addEventListener("change", () => {
    if (state.citations === 0 && style.value !== "") {
      // No citation to rewrite: the choice is simply the style the next citation is written in.
      state = { ...state, styleId: style.value };
      if (typeof ports.setNewDocumentStyle === "function") ports.setNewDocumentStyle(style.value);
      setStatus("New citations will use this style.");
    }
    render();
  });
  applyStyle.addEventListener("click", () => void run("refresh", { styleId: style.value }));
  bibliography.addEventListener("click", () => void run("bibliography", {}));
  refresh.addEventListener("click", () => void run("refresh", {}));
  confirmAccept.addEventListener("click", () => {
    const next = pending;
    pending = null;
    render();
    if (typeof next === "function") void next();
  });
  confirmCancel.addEventListener("click", () => {
    pending = null;
    setStatus("Nothing was changed.");
    render();
  });

  render();

  return Object.freeze({
    /** The document and library state the runtime reported for the open document. */
    setState(next) {
      state = {
        connected: next?.connected === true,
        citations: Number.isInteger(next?.citations) ? next.citations : 0,
        bibliography: next?.bibliography === true,
        bibliographyLocked: next?.bibliographyLocked === true,
        styleId: typeof next?.styleId === "string" ? next.styleId : "",
        styles: Array.isArray(next?.styles) ? next.styles : [],
      };
      renderStyles();
      render();
    },
    setStatus,
    /** Run the bibliography action as though its button had been pressed (ribbon command). */
    insertBibliography: () => run("bibliography", {}),
    /** Run the update action as though its button had been pressed (ribbon command). */
    refreshDocument: () => run("refresh", {}),
  });
}
