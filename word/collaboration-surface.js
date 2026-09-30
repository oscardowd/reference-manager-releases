/** Presentation-only collaborator review (SPEC §9.6; task E10-13.2). */

const VERDICTS = new Set(["duplicate", "probable", "possible", "version-variant"]);
const IMPORT_STATUSES = new Set(["imported", "already-present", "conflict"]);
const CHANGE_KINDS = new Set(["added", "changed", "removed"]);

function text(value) {
  return typeof value === "string" ? value : "";
}

function displayValue(value) {
  if (value === undefined) return "—";
  if (typeof value === "string") return value;
  const encoded = JSON.stringify(value);
  return encoded === undefined ? "—" : encoded;
}

/**
 * Validate and label the bridge response before it reaches the DOM. There is deliberately no
 * `selected`, `recommended` or winning-side field: a comparison is evidence for a decision, not
 * the decision itself (invariant 3).
 */
export function collaboratorReviewPresentation(result) {
  if (!result || !result.library || result.library.status !== "available") {
    throw new Error("InvalidCollaboratorReview");
  }
  const imports = Array.isArray(result.imports) ? result.imports.map((entry) => {
    const status = text(entry && entry.status);
    if (!IMPORT_STATUSES.has(status)) throw new Error("InvalidCollaboratorImportOutcome");
    return Object.freeze({
      documentItemId: text(entry.documentItemId),
      libraryItemId: text(entry.libraryItemId),
      status,
    });
  }) : [];
  const candidates = Array.isArray(result.candidates) ? result.candidates.map((entry) => {
    const verdict = text(entry && entry.verdict);
    if (!VERDICTS.has(verdict) || !Array.isArray(entry.signalCodes)) {
      throw new Error("InvalidCollaboratorCandidate");
    }
    return Object.freeze({
      documentItemId: text(entry.documentItemId),
      libraryItemId: text(entry.libraryItemId),
      verdict,
      signalCodes: Object.freeze(entry.signalCodes.map(text)),
    });
  }) : [];
  const comparisons = Array.isArray(result.comparisons) ? result.comparisons.map((entry) => {
    if (!Array.isArray(entry.differences)) throw new Error("InvalidCollaboratorComparison");
    const differences = entry.differences.map((difference) => {
      const kind = text(difference && difference.kind);
      if (!CHANGE_KINDS.has(kind) || text(difference.path).length === 0) {
        throw new Error("InvalidCollaboratorDifference");
      }
      return Object.freeze({
        field: text(difference.path),
        kind,
        document: displayValue(difference.document),
        library: displayValue(difference.library),
      });
    });
    return Object.freeze({
      documentItemId: text(entry.documentItemId),
      libraryItemId: text(entry.libraryItemId),
      relationship: text(entry.relationship),
      differences: Object.freeze(differences),
    });
  }) : [];
  return Object.freeze({ imports: Object.freeze(imports), candidates: Object.freeze(candidates), comparisons: Object.freeze(comparisons) });
}

function appendText(document, parent, name, value) {
  const element = document.createElement(name);
  element.textContent = value;
  parent.append(element);
  return element;
}

export function mountCollaboratorReviewSurface(document) {
  const status = document.getElementById("collaboration-status");
  const imports = document.getElementById("collaboration-imports");
  const candidates = document.getElementById("collaboration-candidates");
  const comparisons = document.getElementById("collaboration-comparisons");
  if (!status || !imports || !candidates || !comparisons) throw new Error("CollaboratorSurfaceMissing");

  function presentReview(result) {
    const view = collaboratorReviewPresentation(result);
    imports.replaceChildren();
    candidates.replaceChildren();
    comparisons.replaceChildren();
    for (const outcome of view.imports) {
      appendText(document, imports, "li", `${outcome.documentItemId}: ${outcome.status}`);
    }
    for (const candidate of view.candidates) {
      appendText(
        document,
        candidates,
        "li",
        `${candidate.documentItemId} → ${candidate.libraryItemId}: ${candidate.verdict} (${candidate.signalCodes.join(", ")})`,
      );
    }
    for (const comparison of view.comparisons) {
      const section = document.createElement("section");
      appendText(document, section, "h4", `${comparison.documentItemId} versus ${comparison.libraryItemId}`);
      const table = document.createElement("table");
      const head = document.createElement("thead");
      const headingRow = document.createElement("tr");
      for (const heading of ["Field", "Document version", "Library version"]) {
        const cell = appendText(document, headingRow, "th", heading);
        cell.scope = "col";
      }
      head.append(headingRow);
      const body = document.createElement("tbody");
      for (const difference of comparison.differences) {
        const row = document.createElement("tr");
        row.dataset.changeKind = difference.kind;
        appendText(document, row, "th", difference.field).scope = "row";
        appendText(document, row, "td", difference.document);
        appendText(document, row, "td", difference.library);
        body.append(row);
      }
      table.append(head, body);
      section.append(table);
      comparisons.append(section);
    }
    status.textContent = view.comparisons.length === 0
      ? "No metadata differences need review."
      : "Compare both versions. Reference Manager will not choose a side for you.";
  }

  return Object.freeze({ presentReview });
}
