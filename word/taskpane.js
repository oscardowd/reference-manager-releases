/**
 * The task pane (SPEC §9.1; ADR-0101).
 *
 * This file is the only code in the product that calls an Office.js API, and it is deliberately
 * dull: it asks Word what it supports, shows the answers, and formats them as the probe report
 * `@refmgr/word-addin` parses. Nothing here decides anything — interpretation is
 * `interpretProbeReport`, which runs where it can be tested (`npm run addin:report`).
 *
 * No bundler enters this tree (ADR-0086), so the requirement sets arrive through a generated data
 * module rather than an import of the package.
 */

import {
  createWordSessionClient,
  insertBibliographyFromPane,
  insertFromPicker,
  openCitationEdit,
  refreshDocumentFromPane,
  saveCitationEdit,
} from "./word-session.js";
import { mountCitationEditSurface } from "./citation-edit-surface.js";
import { mountDocumentSurface } from "./document-surface.js";
import { PROBED_REQUIREMENT_SETS, PROBE_REPORT_VERSION } from "./probed-sets.js";
import {
  createDocumentBackupBridgeClient,
  formatDocumentBackupRefusal,
} from "./document-backup-bridge.js";
import { mountDocumentUnlinkSurface } from "./unlink-surface.js";
import { mountCollaboratorReviewSurface } from "./collaboration-surface.js";
import { createCollaboratorBridgeClient } from "./collaboration-bridge.js";
import { createAnnotationBridgeClient } from "./annotation-bridge.js";
import { mountAnnotationPickerSurface } from "./annotation-picker-surface.js";
import { createLibraryBridgeClient } from "./library-bridge.js";
import { mountCitationPickerSurface } from "./citation-picker-surface.js";
import { mountAddReferenceSurface } from "./add-reference-surface.js";
import { loadDocumentRuntime } from "./document-runtime.js";
import { startDocumentMode } from "./document-mode.js";
import { createPaneStorage } from "./pane-storage.js";
import { DESKTOP_WORD_ORIGIN } from "./pane-config.js";
import {
  focusTaskPaneWorkflow,
  selectTaskPaneTab,
  setConnectionBanner,
  setupTaskPaneNavigation,
  setupTaskPaneTabs,
} from "./taskpane-ui.js";
import {
  officeDocumentUnlinkPort,
  officeDocumentRepairPort,
  readBibliographyFromOffice,
  readCitationOccurrencesFromOffice,
  readDocumentBytesFromOffice,
  readDocumentForRefreshFromOffice,
  readDocumentDiagnosticsFromOffice,
  readDocumentPreferencesFromOffice,
  readDocumentUnlinkEvidenceFromOffice,
  readSelectedCitationFromOffice,
  writeBibliographyToOffice,
  writeCitationEditToOffice,
  writeInsertionToOffice,
  writeAnnotationInsertionToOffice,
  writeDocumentPreferencesToOffice,
  writeDocumentRefreshToOffice,
  writeNoteRelocationToOffice,
} from "./word-host.js";


let report = null;

setupTaskPaneTabs(document);
setupTaskPaneNavigation(document);

Office.onReady((info) => {
  report = capture(info);
  render(report);
  // Keyboard shortcuts are a desktop Word feature; iPad has none, so the tip naming one is hidden.
  const shortcut = document.getElementById("citation-shortcut");
  if (shortcut) shortcut.hidden = !isSetSupported({ name: "KeyboardShortcuts", version: "1.1" });
  void connectDesktop();
});

/**
 * Show the pane, where the host can be asked to. Word on iPad has no shared runtime, so no
 * `Office.addin` (ADR-0242 §6); there the pane is already what the writer opened, and nothing
 * needs showing.
 */
async function showPane() {
  try {
    if (Office.addin && typeof Office.addin.showAsTaskpane === "function") await Office.addin.showAsTaskpane();
  } catch {
    // Already showing, or the host declined; the workflow below still runs in the pane.
  }
}

async function openInsertCitation() {
  await showPane();
  focusTaskPaneWorkflow(document, "insert-citation", "citation-search-query");
}

async function openInsertPdfAnnotation() {
  await showPane();
  focusTaskPaneWorkflow(document, "annotation-picker", "annotation-search-query");
}

/**
 * The ribbon's functions (E10-19.2). This page is the manifest's `FunctionFile` — the add-in runs
 * one shared runtime — so a ribbon button's `ExecuteFunction` is answered here or nowhere. Each
 * opens the pane on the workflow it names and then completes the event, whatever happened: an
 * uncompleted event leaves Word's button spinning. Anything that needs the writer's decision is
 * asked in the pane, never by the button.
 */
function ribbonCommand(run) {
  return async (event) => {
    try {
      await showPane();
      await run();
    } catch {
      // The pane states its own failures; a ribbon event has no surface to show one on.
    } finally {
      if (event && typeof event.completed === "function") event.completed();
    }
  };
}

// Word on iPad has neither add-in commands nor a shared runtime (ADR-0242 §6): there is no ribbon
// to answer, and `Office.actions` may be missing or refuse. Every workflow a button opens is reachable
// in the pane by touch, so a host without the ribbon loses a shortcut, never a capability.
try {
  if (Office.actions && typeof Office.actions.associate === "function") {
    Office.actions.associate("OpenInsertCitation", openInsertCitation);
    Office.actions.associate("OpenInsertPdfAnnotation", openInsertPdfAnnotation);
    Office.actions.associate("insertCitation", ribbonCommand(async () => {
      focusTaskPaneWorkflow(document, "insert-citation", "citation-search-query");
    }));
    Office.actions.associate("editCitation", ribbonCommand(async () => {
      selectTaskPaneTab(document, "tab-cite");
      await globalThis.refmgrCitationEditSurface?.openSelected();
    }));
    Office.actions.associate("insertBibliography", ribbonCommand(async () => {
      selectTaskPaneTab(document, "tab-document");
      await documentSurface.insertBibliography();
    }));
    Office.actions.associate("refreshDocument", ribbonCommand(async () => {
      selectTaskPaneTab(document, "tab-document");
      await documentSurface.refreshDocument();
    }));
    Office.actions.associate("openCitationStyle", ribbonCommand(async () => {
      focusTaskPaneWorkflow(document, "document-update", "citation-style");
    }));
  }
} catch {
  // No add-in commands on this host.
}
// The picker/bridge composition calls this only with a plan prepared by @refmgr/word-addin.
globalThis.refmgrWriteInsertion = writeInsertionToOffice;
// E14-04.4: host write for prose + live citation; picker surface + bridge client are below.
globalThis.refmgrWriteAnnotationInsertion = writeAnnotationInsertionToOffice;
globalThis.refmgrReadSelectedCitation = readSelectedCitationFromOffice;
globalThis.refmgrWriteCitationEdit = writeCitationEditToOffice;
globalThis.refmgrReadBibliography = readBibliographyFromOffice;
globalThis.refmgrWriteBibliography = writeBibliographyToOffice;
globalThis.refmgrReadDocumentForRefresh = readDocumentForRefreshFromOffice;
globalThis.refmgrWriteDocumentRefresh = writeDocumentRefreshToOffice;
// A style change is a refresh: it reads and writes through the refresh seam above. These two are
// the settings-only path, which never touches a content control.
globalThis.refmgrReadDocumentPreferences = readDocumentPreferencesFromOffice;
globalThis.refmgrWriteDocumentPreferences = writeDocumentPreferencesToOffice;
// E10-09.2. Both take the probe's feature verdicts: footnotes are above the manifest's declared
// floor, so whether this host has them is a measurement, never an assumption.
globalThis.refmgrReadCitationOccurrences = readCitationOccurrencesFromOffice;
globalThis.refmgrWriteNoteRelocation = writeNoteRelocationToOffice;
// E10-10 is deliberately read-only. The domain scanner returns nine structural sections and E10-11
// owns every repair, including its mandatory backup.
globalThis.refmgrReadDocumentDiagnostics = readDocumentDiagnosticsFromOffice;
// E10-11.2 joins Word's compressed-document reader to E10-11.1's authenticated bridge routes.
// The desktop composition supplies the ephemeral base URL and bearer token; the token is never
// accepted through the URL. Any refusal is rendered as fixed product copy, never a response body.
globalThis.refmgrReadDocumentBytes = readDocumentBytesFromOffice;
globalThis.refmgrReadDocumentForUnlink = readDocumentUnlinkEvidenceFromOffice;
globalThis.refmgrDocumentRepairPort = (bridge) => {
  const connection = bridge ?? {};
  return officeDocumentRepairPort(createDocumentBackupBridgeClient({
    ...connection,
    onFailure(error) {
      document.getElementById("document-repair-status").textContent = formatDocumentBackupRefusal(error);
      if (typeof connection.onFailure === "function") connection.onFailure(error);
    },
  }));
};

// E10-12.1 is presentation, not policy. The desktop/domain composition listens for
// `refmgr:unlink-review` and `refmgr:unlink-commit`, then returns the plan/outcome through this
// surface. The host port below implements E10-12's exact remove contract over Office.js.
// The `onFailure` is not decoration and is the same one the repair port above installs: the
// bridge client is the only layer that knows *which* request was abandoned, so an aborted commit
// — the one failure that leaves a copy undecided — has its sentence here or nowhere. The
// domain's refusal, presented after this, tells the fuller story; this covers the wait before it
// arrives and any failure the composition above does not present. A caller-supplied handler is
// chained rather than replaced, exactly as the repair composition does.
globalThis.refmgrDocumentUnlinkPort = (bridge) => {
  const connection = bridge ?? {};
  return officeDocumentUnlinkPort(createDocumentBackupBridgeClient({
    ...connection,
    onFailure(error) {
      document.getElementById("unlink-status").textContent = formatDocumentBackupRefusal(error);
      if (typeof connection.onFailure === "function") connection.onFailure(error);
    },
  }));
};
globalThis.refmgrDocumentUnlinkSurface = mountDocumentUnlinkSurface(document);
// E10-13.2 only presents the bridge/domain result. It never chooses a relink or a metadata side.
globalThis.refmgrCollaboratorReviewSurface = mountCollaboratorReviewSurface(document);
globalThis.refmgrCollaboratorBridgeClient = (bridge) => createCollaboratorBridgeClient(bridge ?? {});
// E14-04.4: search over the bridge; selection/omit acknowledgements on the surface; composition
// listens for `refmgr:annotation-select` / `refmgr:annotation-insert` and supplies domain outcomes.
globalThis.refmgrAnnotationBridgeClient = (bridge) => createAnnotationBridgeClient(bridge ?? {});
globalThis.refmgrAnnotationPickerSurface = mountAnnotationPickerSurface(document);

// One mount; canonical insertion preparation runs in the desktop domain runtime (ADR-0241).
const citationPickerPorts = {};
const picker = mountCitationPickerSurface(document, citationPickerPorts);
// One mount each, with the ports filled in when the desktop connects: a second mount over the
// same elements would install a second listener on every one of them.
const citationEditPorts = {};
globalThis.refmgrCitationEditSurface = mountCitationEditSurface(document, citationEditPorts);
// Editing replaces the search while it is open — one task on screen at a time. The edit surface
// owns its own `hidden`; this mirrors it onto the list and the tray, so what is not shown is also
// out of the tab order and the accessibility tree rather than merely covered.
{
  const sheet = document.getElementById("citation-edit");
  const behind = [document.getElementById("insert-citation"), document.getElementById("citation-tray")];
  const mirror = () => { for (const element of behind) if (element) element.hidden = !sheet.hidden; };
  new globalThis.MutationObserver(mirror).observe(sheet, { attributes: true, attributeFilter: ["hidden"] });
  mirror();
}
// The Document tab. Its ports are filled when the desktop connects, like the picker's.
const documentPorts = {};
const documentSurface = mountDocumentSurface(document, documentPorts);
globalThis.refmgrCitationPickerSurface = picker;
globalThis.refmgrConnectCitationPicker = (bridge) => {
  const client = createLibraryBridgeClient(bridge ?? {});
  Object.assign(citationPickerPorts, client);
  return picker.open();
};
/**
 * The session this pane is talking to: the desktop's (library mode) or its own over the document's
 * references (document mode, ADR-0242). Every port below reads it at call time, so switching mode
 * switches every surface at once.
 */
let client = desktopClient();
const storage = createPaneStorage();
/** Started once, the first time the desktop cannot be reached; reused by later reconnects. */
let documentMode = null;

/** The desktop's session: same-origin from its own host, cross-origin from a hosted pane. */
function desktopClient() {
  return createWordSessionClient(undefined, { desktopOrigin: DESKTOP_WORD_ORIGIN });
}

/** Word on iPad or on the web: hosts that never have the desktop app beside them. */
function hostWithoutDesktop() {
  const platform = String(report?.platform ?? "");
  return platform === "iOS" || platform === "OfficeOnline";
}

/**
 * The Document tab's style list. In document mode the session lists only the styles downloaded so
 * far; every style the pane could download is offered too, and the document's own style stays in
 * the list even when it is one the catalog does not have.
 */
function offeredStyles(installed) {
  if (client.mode !== "document" || documentMode === null) return installed;
  const offered = [...installed];
  for (const entry of documentMode.catalogStyles()) {
    if (!offered.some((style) => style.id === entry.id)) offered.push(entry);
  }
  return offered.sort((left, right) => left.title.localeCompare(right.title));
}

/**
 * Ask the runtime what the open document holds, and show it: the style every new citation is
 * written in, the references already cited, and the Document tab's summary. Run on connect and
 * after every write, so the pane never describes the document as it was before the last action.
 */
async function refreshSessionState() {
  const state = await client.request("/session", { document: await readDocumentForRefreshFromOffice() });
  picker.setStyle({ styleId: state.styleId, locale: state.locale });
  picker.setCitedItemIds(state.citedItemIds);
  documentSurface.setState({
    connected: true,
    citations: state.document?.citations ?? 0,
    bibliography: state.document?.bibliography === true,
    bibliographyLocked: state.document?.bibliographyLocked === true,
    styleId: state.styleId,
    styles: offeredStyles(state.styles),
  });
  if (state.document?.problem) documentSurface.setStatus(state.document.problem);
  else if (state.document?.needsUpdate) {
    documentSurface.setStatus("Some citations were deleted or edited by hand. Choose Update document to bring the document up to date.");
  }
  return state;
}

/** A write succeeded: re-read the document state, without letting a failed re-read hide the success. */
async function afterDocumentWrite() {
  try { await refreshSessionState(); } catch { /* The next action re-reads anyway. */ }
}

// The ports every surface calls. Each reads `client` when it is called, never when it is assigned.
Object.assign(citationPickerPorts, {
  recent: (query) => client.recent(query),
  search: (query) => client.search(query),
  collections: () => client.collections(),
  previewCitation: (body) => client.previewCitation(body),
  insertCitation: async (input) => {
    const outcome = await insertFromPicker({
      ...input,
      client,
      readDocument: readDocumentForRefreshFromOffice,
      writeInsertion: writeInsertionToOffice,
      writeRefresh: writeDocumentRefreshToOffice,
    });
    await afterDocumentWrite();
    return outcome;
  },
});
// E10-05.1. The same client and the same Office reads; the edit panel differs from the picker
// only in starting from a citation the document already holds.
Object.assign(citationEditPorts, {
  previewCitation: (body) => client.previewCitation(body),
  openEdit: () => openCitationEdit({ client, readSelected: readSelectedCitationFromOffice }),
  saveEdit: async (input) => {
    const outcome = await saveCitationEdit({
      ...input,
      client,
      readSelected: readSelectedCitationFromOffice,
      writeEdit: writeCitationEditToOffice,
      readDocument: readDocumentForRefreshFromOffice,
      writeRefresh: writeDocumentRefreshToOffice,
    });
    await afterDocumentWrite();
    // The edit sheet closes on success; say what happened where the writer is now looking —
    // after the re-read, whose list repaint would otherwise replace the sentence.
    picker.setStatus(outcome.message);
    return outcome;
  },
});
Object.assign(documentPorts, {
  refreshDocument: (options) => refreshDocumentFromPane({
    ...options,
    client,
    readDocument: readDocumentForRefreshFromOffice,
    writeRefresh: writeDocumentRefreshToOffice,
  }),
  insertBibliography: (options) => insertBibliographyFromPane({
    ...options,
    client,
    readDocument: readDocumentForRefreshFromOffice,
    writeBibliography: writeBibliographyToOffice,
  }),
  // No citations yet: the style chosen is simply the one the first citation is written in.
  setNewDocumentStyle: (styleId) => picker.setStyle({ styleId, locale: "en-US" }),
  onChanged: afterDocumentWrite,
});

// §9.2's "create or import a missing reference" in document mode (ADR-0242 §5). In library mode the
// desktop app adds references, so the surface is withdrawn there.
const addReference = mountAddReferenceSurface(document, {
  lookup: (identifiers) => documentMode.lookup(identifiers),
  addManual: (input) => documentMode.addManual(input),
  onAdded: async () => {
    // Newest first: an empty search lists what was just added at the top.
    const query = document.getElementById("citation-search-query");
    if (query) query.value = "";
    await picker.runQuery();
  },
});

/** Words that differ between the modes. Product copy only; nothing from the library or document. */
function presentMode(mode) {
  const query = document.getElementById("citation-search-query");
  const label = document.querySelector('label[for="citation-search-query"]');
  const description = document.getElementById("bridge-description");
  const searchText = mode === "document" ? "Search this document's references" : "Search your library";
  if (query) query.placeholder = searchText;
  if (label) label.textContent = searchText;
  if (description) {
    description.textContent = mode === "document"
      ? hostWithoutDesktop()
        ? "Word here cannot reach the desktop app, so this pane works from the references saved in the document. Citation styles are downloaded from the Citation Style Language project when first needed; nothing from your document is sent."
        : "Reference Manager is not running, so this pane works from the references saved in the document. Open Reference Manager, then choose Reconnect, to search your library. Citation styles are downloaded from the Citation Style Language project when first needed; nothing from your document is sent."
      : "This pane reads your library from Reference Manager on this computer. Keep it open while you write.";
  }
}

async function connectDesktop() {
  setConnectionBanner(document, "pending");
  const result = document.getElementById("bridge-result");
  result.textContent = "Connecting…";
  try {
    client = desktopClient();
    await refreshSessionState();
    presentMode("library");
    addReference.setAvailable(false);
    await picker.open();
    setConnectionBanner(document, "ready");
    result.textContent = "Connected to your library. Keep Reference Manager open while you write.";
    if (report) { record({ outcome: "reached", target: window.location.origin, detail: "Live library and canonical insertion connected" }); document.getElementById("report").textContent = JSON.stringify(report, null, 2); }
  } catch (error) {
    await enterDocumentMode(error);
  }
}

/**
 * No desktop: work from the document (ADR-0242). Only if that too cannot start — the runtime or the
 * style list could not load — is the pane "not connected", and then it says why.
 */
async function enterDocumentMode(cause) {
  const result = document.getElementById("bridge-result");
  try {
    documentMode ??= await startDocumentMode({ loadRuntime: loadDocumentRuntime, storage });
    client = documentMode.client;
    await refreshSessionState();
    presentMode("document");
    addReference.setAvailable(true, documentMode.disclosures());
    await picker.open();
    setConnectionBanner(document, "document");
    result.textContent = hostWithoutDesktop()
      ? "Working from this document's references."
      : "Working from this document's references. Reference Manager was not reached.";
  } catch (error) {
    const detail = error && error.message ? String(error.message)
      : cause && cause.message ? String(cause.message) : "Open Reference Manager, then choose Reconnect.";
    setConnectionBanner(document, "error");
    result.textContent = detail;
    picker.setStatus(`${detail} Reconnect from the Help tab.`);
    documentSurface.setState({ connected: false });
  }
}

function capture(info) {
  return {
    reportVersion: PROBE_REPORT_VERSION,
    capturedAt: new Date().toISOString(),
    host: String(info && info.host ? info.host : "unknown"),
    platform: String(info && info.platform ? info.platform : "unknown"),
    officeVersion: officeVersion(),
    taskPaneOrigin: window.location.origin,
    requirementSets: PROBED_REQUIREMENT_SETS.map((set) => ({
      name: set.name,
      version: set.version,
      supported: isSetSupported(set),
    })),
    bridge: { outcome: "not-attempted", target: null, detail: null },
  };
}

function officeVersion() {
  const diagnostics = Office.context && Office.context.diagnostics;
  return diagnostics && typeof diagnostics.version === "string" ? diagnostics.version : null;
}

/**
 * A set the host throws on is reported unsupported, not skipped: an entry missing from the report
 * reads as "not measured", and this one was measured — the answer was just an exception.
 */
function isSetSupported(set) {
  try {
    return Office.context.requirements.isSetSupported(set.name, set.version) === true;
  } catch {
    return false;
  }
}

function render(current) {
  const status = document.getElementById("status");
  status.textContent = `${current.host} on ${current.platform}. Your document is unchanged.`;
  setConnectionBanner(document, "pending");

  const host = document.getElementById("host");
  host.replaceChildren();
  for (const [label, value] of [
    ["Host", current.host],
    ["Platform", current.platform],
    ["Office version", current.officeVersion === null ? "not reported" : current.officeVersion],
    ["Task pane origin", current.taskPaneOrigin],
  ]) {
    const term = document.createElement("dt");
    term.textContent = label;
    const definition = document.createElement("dd");
    definition.textContent = value;
    host.append(term, definition);
  }

  const body = document.querySelector("#sets tbody");
  body.replaceChildren();
  for (const entry of current.requirementSets) {
    const row = document.createElement("tr");
    row.append(cell(entry.name), cell(entry.version), cell(entry.supported ? "yes" : "no"));
    row.className = entry.supported ? "supported" : "unsupported";
    body.append(row);
  }

  document.getElementById("report").textContent = JSON.stringify(current, null, 2);
}

function cell(text) {
  const element = document.createElement("td");
  element.textContent = text;
  return element;
}

document.getElementById("copy").addEventListener("click", () => {
  const text = document.getElementById("report").textContent;
  if (navigator.clipboard) {
    navigator.clipboard.writeText(text).catch(() => {
      document.getElementById("report").focus();
    });
  } else {
    document.getElementById("report").focus();
  }
});

document.getElementById("bridge-form").addEventListener("submit", (event) => {
  event.preventDefault();
  void connectDesktop();
});

function record(bridge) {
  report = { ...report, bridge };
}
