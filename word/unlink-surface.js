/**
 * Task-pane presentation for E10-12's unlink plan (SPEC §9.7; ADR-0118).
 *
 * Policy stays in `@refmgr/word-addin`. This unbundled browser module only reads the four choices,
 * presents the plan it is handed, proves the warning codes on screen were checked, and presents
 * the actual outcome. Changing an option invalidates the review: an acknowledgement must always
 * describe the exact plan that will be committed.
 */

const MODES = new Set(["unlink", "convert-to-plain-text"]);
const OUTCOME_COPY = Object.freeze({
  unlinked: "Citation links were removed after a verified linked copy was saved.",
  refused: "Nothing was changed. Resolve the refusal and review the document again.",
  partial: "The operation stopped part-way through. Use the checklist below to inspect what remains.",
});

/**
 * Refusals whose reason changes what the reader has to do, keyed by the domain's own reason code.
 *
 * A reason with no entry here keeps {@link OUTCOME_COPY}.refused: a refusal this surface cannot
 * name is still a refusal it must show, so an unrecognised reason must not throw and must not
 * become silence. Only reasons that send the user somewhere new earn a sentence of their own.
 */
const REFUSAL_COPY = Object.freeze({
  // Says the two things separately because they are separately true: the document was not touched
  // (the copy is taken strictly before the first write), and a copy may or may not have been
  // saved. The shared refusal sentence would claim the second.
  "persist-outcome-unknown":
    "Nothing in the document was changed. Reference Manager stopped waiting while the linked copy was being saved, "
    + "so a copy may or may not have been saved — check your document backups before reviewing the document again.",
});

function stringValue(value) {
  return typeof value === "string" ? value : "";
}

export function documentUnlinkOptionsFromControls(controls) {
  const mode = stringValue(controls.mode && controls.mode.value);
  if (!MODES.has(mode)) throw new Error("InvalidDocumentUnlinkMode");
  return Object.freeze({
    mode,
    citations: controls.citations && controls.citations.checked === true,
    bibliography: controls.bibliography && controls.bibliography.checked === true,
    retainRecoveryData: controls.retainRecoveryData && controls.retainRecoveryData.checked === true,
  });
}

function uniqueCodes(entries) {
  const codes = entries.map((entry) => stringValue(entry && entry.code));
  if (codes.some((code) => code.length === 0) || new Set(codes).size !== codes.length) {
    throw new Error("InvalidDocumentUnlinkWarnings");
  }
  return codes;
}

/**
 * Build the acknowledgement only when the DOM contains exactly the plan's warnings and all are
 * checked. The domain checks the codes again at commit; this check proves the surface did not omit
 * a warning before asking the domain to commit.
 */
export function acknowledgeDisplayedWarnings(plan, displayedCodes, checkedCodes, acceptedAt) {
  const expected = uniqueCodes(Array.isArray(plan && plan.warnings) ? plan.warnings : []);
  const displayed = [...displayedCodes].map(stringValue);
  const checked = new Set([...checkedCodes].map(stringValue));
  const sameDisplay = displayed.length === expected.length &&
    expected.every((code, index) => displayed[index] === code);
  const allChecked = checked.size === expected.length && expected.every((code) => checked.has(code));
  if (!sameDisplay || !allChecked) return Object.freeze({ ok: false, reason: "warnings-not-acknowledged" });
  const instant = stringValue(acceptedAt);
  if (Number.isNaN(Date.parse(instant))) return Object.freeze({ ok: false, reason: "invalid-accepted-at" });
  return Object.freeze({
    ok: true,
    acknowledgement: Object.freeze({
      acceptedWarningCodes: Object.freeze([...expected]),
      acceptedAt: instant,
    }),
  });
}

export function documentUnlinkOutcomePresentation(outcome) {
  const status = stringValue(outcome && outcome.status);
  if (!Object.prototype.hasOwnProperty.call(OUTCOME_COPY, status)) throw new Error("InvalidDocumentUnlinkOutcome");
  const modified = stringValue(outcome.documentModified);
  if (!new Set(["yes", "no", "unknown"]).has(modified)) throw new Error("InvalidDocumentModifiedOutcome");
  const modification = modified === "unknown"
    ? "Word stopped during a write. Whether the document changed is unknown; do not submit it until you inspect it."
    : modified === "yes"
      ? "The open document was changed."
      : "The open document was not changed.";
  const stage = status === "partial" && stringValue(outcome.stage).length > 0
    ? ` Failed stage: ${String(outcome.stage)}.`
    : "";
  // Read only off a refusal. A `partial` has a verified copy behind it by construction — the
  // write cannot begin without one — so a stale `reason` there would raise a false alarm about a
  // backup that is known to exist.
  const reason = status === "refused" ? stringValue(outcome.reason) : "";
  const named = Object.prototype.hasOwnProperty.call(REFUSAL_COPY, reason) ? REFUSAL_COPY[reason] : undefined;
  return Object.freeze({
    status,
    heading: status === "unlinked" ? "Unlink complete" : status === "partial" ? "Unlink incomplete" : "Unlink refused",
    message: `${named ?? OUTCOME_COPY[status]} ${modification}${stage}`,
    checklist: Object.freeze(Array.isArray(outcome.checklist) ? [...outcome.checklist] : []),
    documentModified: modified,
    // Exposed as its own field rather than left for a caller to sniff out of the sentence: the
    // undecided copy is the one refusal that leaves the user with something to go and look at.
    backupOutcomeUnknown: reason === "persist-outcome-unknown",
  });
}

function appendTextElement(document, parent, name, text, className) {
  const element = document.createElement(name);
  element.textContent = text;
  if (className !== undefined) element.className = className;
  parent.append(element);
  return element;
}

function dispatchDetail(document, target, name, detail) {
  const event = document.createEvent("CustomEvent");
  event.initCustomEvent(name, true, false, detail);
  target.dispatchEvent(event);
}

function renderChecklist(document, list, checklist) {
  list.replaceChildren();
  for (const item of checklist) {
    const row = document.createElement("li");
    row.className = `checklist-${stringValue(item.status)}`;
    appendTextElement(document, row, "strong", `${stringValue(item.status).toUpperCase()}: `);
    row.append(document.createTextNode(stringValue(item.statement)));
    list.append(row);
  }
}

export function mountDocumentUnlinkSurface(document, now = () => new Date()) {
  const form = document.getElementById("document-unlink-form");
  const controls = {
    mode: document.getElementById("unlink-mode"),
    citations: document.getElementById("unlink-citations"),
    bibliography: document.getElementById("unlink-bibliography"),
    retainRecoveryData: document.getElementById("unlink-retain-recovery"),
  };
  const warnings = document.getElementById("unlink-warnings");
  const refusals = document.getElementById("unlink-refusals");
  const checklist = document.getElementById("unlink-checklist");
  const status = document.getElementById("unlink-status");
  const commit = document.getElementById("unlink-commit");
  if (!form || Object.values(controls).some((control) => !control) || !warnings || !refusals ||
      !checklist || !status || !commit) throw new Error("DocumentUnlinkSurfaceMissing");

  let currentPlan = null;

  function invalidate() {
    currentPlan = null;
    warnings.replaceChildren();
    refusals.replaceChildren();
    checklist.replaceChildren();
    commit.disabled = true;
    status.textContent = "Options changed. Review the unlink plan again before continuing.";
  }

  function warningState() {
    const inputs = [...warnings.querySelectorAll("input[data-warning-code]")];
    return {
      displayed: inputs.map((input) => input.dataset.warningCode),
      checked: inputs.filter((input) => input.checked).map((input) => input.dataset.warningCode),
    };
  }

  function updateCommit() {
    if (currentPlan === null || currentPlan.refusals.length > 0) {
      commit.disabled = true;
      return;
    }
    const state = warningState();
    commit.disabled = !acknowledgeDisplayedWarnings(
      currentPlan,
      state.displayed,
      state.checked,
      now().toISOString(),
    ).ok;
  }

  function presentPlan(plan) {
    currentPlan = plan;
    warnings.replaceChildren();
    refusals.replaceChildren();
    for (const refusal of plan.refusals) {
      appendTextElement(document, refusals, "li", stringValue(refusal.remedy), "unlink-refusal");
    }
    for (const warning of plan.warnings) {
      const row = document.createElement("li");
      const label = document.createElement("label");
      const input = document.createElement("input");
      input.type = "checkbox";
      input.dataset.warningCode = stringValue(warning.code);
      input.addEventListener("change", updateCommit);
      label.append(input, document.createTextNode(` ${stringValue(warning.message)}`));
      row.append(label);
      warnings.append(row);
    }
    renderChecklist(document, checklist, plan.checklist);
    status.textContent = plan.refusals.length > 0
      ? "Nothing can be changed with these options."
      : "Read and acknowledge every warning before unlinking.";
    updateCommit();
  }

  function presentOutcome(outcome) {
    const presentation = documentUnlinkOutcomePresentation(outcome);
    status.textContent = `${presentation.heading}. ${presentation.message}`;
    status.className = `unlink-outcome-${presentation.status}`
      + (presentation.documentModified === "unknown" ? " unlink-modified-unknown" : "")
      + (presentation.backupOutcomeUnknown ? " unlink-backup-unknown" : "");
    renderChecklist(document, checklist, presentation.checklist);
    commit.disabled = true;
    currentPlan = null;
  }

  form.addEventListener("change", invalidate);
  form.addEventListener("submit", (event) => {
    event.preventDefault();
    dispatchDetail(document, form, "refmgr:unlink-review", documentUnlinkOptionsFromControls(controls));
  });
  commit.addEventListener("click", () => {
    if (currentPlan === null) return;
    const state = warningState();
    const accepted = acknowledgeDisplayedWarnings(
      currentPlan,
      state.displayed,
      state.checked,
      now().toISOString(),
    );
    if (!accepted.ok) {
      updateCommit();
      return;
    }
    commit.disabled = true;
    dispatchDetail(document, form, "refmgr:unlink-commit", {
      plan: currentPlan,
      acknowledgement: accepted.acknowledgement,
    });
  });

  return Object.freeze({
    options: () => documentUnlinkOptionsFromControls(controls),
    presentPlan,
    presentOutcome,
    setBusy(message) {
      commit.disabled = true;
      status.textContent = stringValue(message);
    },
  });
}
