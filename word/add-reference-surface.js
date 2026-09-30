/**
 * "Add a reference" in document mode (SPEC §9.2; ADR-0242 decision 5).
 *
 * Presentation only. It collects a DOI or PMID, or the details of a reference typed by hand, hands
 * them to its ports and says what happened. Recognition, lookup, reconciliation, validation and the
 * CSL projection are the domain's (`document-import.ts`); the reference lands on the pane's shelf,
 * and the picker lists it for the writer to tick. Nothing here writes to the document.
 *
 * **Before anything leaves the pane it says where it goes.** The disclosure is built from the
 * domain's own table of providers, not written out here, so it cannot name fewer hosts than the
 * lookup actually asks (ADR-0146).
 */

const KIND_LABELS = Object.freeze({
  doi: "a DOI",
  pmid: "a PMID",
  pmcid: "a PMCID",
  isbn: "an ISBN",
  nct: "a trial number",
  arxiv: "an arXiv id",
});

/** "Looking up sends only the identifier: a DOI to api.crossref.org and api.datacite.org; …" */
export function lookupDisclosureText(disclosures) {
  const hostsByKind = new Map();
  for (const entry of Array.isArray(disclosures) ? disclosures : []) {
    for (const kind of entry.kinds) {
      if (!hostsByKind.has(kind)) hostsByKind.set(kind, []);
      if (!hostsByKind.get(kind).includes(entry.host)) hostsByKind.get(kind).push(entry.host);
    }
  }
  const clauses = [...hostsByKind].map(([kind, hosts]) => `${KIND_LABELS[kind] ?? kind} to ${hosts.join(" and ")}`);
  if (clauses.length === 0) return "Looking up is not available in this pane.";
  return `Looking up sends only what you type: ${clauses.join("; ")}. Nothing from your document is sent.`;
}

function plural(count, one, many) {
  return `${String(count)} ${count === 1 ? one : many}`;
}

/** One sentence for a lookup's outcome, from counts and the domain's own reasons. */
export function lookupOutcomeText(result) {
  const imported = Array.isArray(result?.imported) ? result.imported : [];
  const unresolved = Array.isArray(result?.unresolved) ? result.unresolved : [];
  const added = imported.filter((entry) => entry.created).length;
  const already = imported.length - added;
  const parts = [];
  if (added > 0) {
    parts.push(added === 1 ? "Added 1 reference; tick it below to cite it." : `Added ${String(added)} references; tick them below to cite them.`);
  }
  if (already > 0) parts.push(`${plural(already, "reference was", "references were")} already here.`);
  if (unresolved.length > 0) {
    // The domain's reason names no host and repeats no metadata; the identifier is the writer's own.
    parts.push(unresolved.map((entry) => `${entry.identifier}: ${entry.reason}.`).join(" "));
  }
  return parts.length > 0 ? parts.join(" ") : "Nothing to look up. Paste a DOI or PMID.";
}

function value(document, id) {
  const element = document.getElementById(id);
  return element && typeof element.value === "string" ? element.value : "";
}

export function mountAddReferenceSurface(document, ports = {}) {
  const panel = document.getElementById("add-reference");
  const lookup = document.getElementById("add-reference-lookup");
  const identifier = document.getElementById("add-reference-identifier");
  const disclosure = document.getElementById("add-reference-disclosure");
  const manual = document.getElementById("add-reference-manual-form");
  const status = document.getElementById("add-reference-status");
  if (!panel || !lookup || !identifier || !disclosure || !manual || !status) throw new Error("AddReferenceSurfaceMissing");

  let busy = false;
  const setStatus = (message) => { status.textContent = message; };

  lookup.addEventListener("submit", (event) => {
    event.preventDefault();
    if (busy || typeof ports.lookup !== "function") return;
    const entries = identifier.value.split(/[\n,;]+/u).map((entry) => entry.trim()).filter((entry) => entry.length > 0);
    if (entries.length === 0) {
      setStatus("Paste a DOI or PMID first.");
      return;
    }
    busy = true;
    setStatus("Looking up…");
    Promise.resolve()
      .then(() => ports.lookup(entries))
      .then(async (result) => {
        setStatus(lookupOutcomeText(result));
        if ((result?.imported ?? []).length > 0) {
          identifier.value = "";
          if (typeof ports.onAdded === "function") await ports.onAdded();
        }
      })
      .catch((error) => setStatus(error && error.message ? String(error.message) : "The lookup did not complete. Try again."))
      .finally(() => { busy = false; });
  });

  manual.addEventListener("submit", (event) => {
    event.preventDefault();
    if (busy || typeof ports.addManual !== "function") return;
    const input = {
      type: value(document, "add-reference-type") || "article-journal",
      title: value(document, "add-reference-title"),
      authors: value(document, "add-reference-authors").split(/\n+/u),
      year: value(document, "add-reference-year"),
      containerTitle: value(document, "add-reference-container"),
      volume: value(document, "add-reference-volume"),
      pages: value(document, "add-reference-pages"),
      publisher: value(document, "add-reference-publisher"),
      doi: value(document, "add-reference-doi"),
    };
    busy = true;
    Promise.resolve()
      .then(() => ports.addManual(input))
      .then(async (outcome) => {
        if (!outcome || outcome.ok !== true) {
          setStatus(outcome?.message ?? "That reference could not be added.");
          return;
        }
        manual.reset();
        setStatus("Added 1 reference; tick it below to cite it.");
        if (typeof ports.onAdded === "function") await ports.onAdded();
      })
      .catch((error) => setStatus(error && error.message ? String(error.message) : "That reference could not be added."))
      .finally(() => { busy = false; });
  });

  return Object.freeze({
    /** Offer the surface (document mode) or withdraw it (library mode, where the desktop adds). */
    setAvailable(available, disclosures) {
      panel.hidden = available !== true;
      if (available === true) disclosure.textContent = lookupDisclosureText(disclosures);
      else setStatus("");
    },
    setStatus,
  });
}
