/** Office.js adapter for E10-04.4/E10-05. Domain preparation and compensation live in word-addin. */

const REFMGR_CUSTOM_XML_NAMESPACE = "urn:refmgr:doc-schema:1";
const CLUSTER_TAG_PREFIX = "refmgr-cite:";

function officeCall(start) {
  return new Promise((resolve, reject) => {
    start((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
      else reject(new Error(result.error && result.error.code ? String(result.error.code) : "OfficeAsyncError"));
    });
  });
}

export function officeInsertionPort() {
  return {
    async addCustomXml(xml) {
      const part = await officeCall((done) => Office.context.document.customXmlParts.addAsync(xml, done));
      return { id: part.id };
    },
    async insertContentControl(control) {
      const word = globalThis.Word;
      await word.run(async (context) => {
        const selection = context.document.getSelection();
        const parent = selection.parentContentControlOrNullObject;
        parent.load("tag");
        await context.sync();
        if (!parent.isNullObject && parent.tag.startsWith(CLUSTER_TAG_PREFIX)) {
          throw new Error("InsertionInsideCitation");
        }
        // Start does not discard a selected passage. Replace would turn an accidental selection
        // into silent data loss; at a collapsed cursor both locations are identical.
        const range = selection.insertText(control.visibleText, word.InsertLocation.start);
        const contentControl = range.insertContentControl();
        contentControl.tag = control.tag;
        contentControl.title = control.title;
        contentControl.appearance = control.appearance;
        await context.sync();
        // Cursor placement is a separate, non-destructive action after the committed write.
        // A selection failure must never roll back the payload of an already-written citation.
        try { contentControl.getRange("After").select(); await context.sync(); } catch { /* The citation remains valid. */ }
      });
    },
    async deleteCustomXml(id) {
      const part = await officeCall((done) => Office.context.document.customXmlParts.getByIdAsync(id, done));
      if (part !== null) await officeCall((done) => part.deleteAsync(done));
    },
  };
}

/** E14-04.3: ordinary annotation prose followed by one live processor-backed citation control. */
export function officeAnnotationInsertionPort() {
  const insertion = officeInsertionPort();
  return {
    addCustomXml: insertion.addCustomXml,
    deleteCustomXml: insertion.deleteCustomXml,
    async insertAnnotationContent({ proseRuns, control }) {
      const word = globalThis.Word;
      await word.run(async (context) => {
        // Collapse to the selection start: an annotation insert must not replace selected prose.
        let tail = context.document.getSelection().getRange("Start");
        for (const prose of proseRuns) {
          tail = tail.insertText(prose, word.InsertLocation.after);
          // The separator is ordinary manuscript text and stays outside the citation control.
          tail = tail.insertText(" ", word.InsertLocation.after);
        }
        const citationRange = tail.insertText(control.visibleText, word.InsertLocation.after);
        const contentControl = citationRange.insertContentControl();
        contentControl.tag = control.tag;
        contentControl.title = control.title;
        contentControl.appearance = control.appearance;
        await context.sync();
      });
    },
  };
}

async function selectedCitationControl() {
  const word = globalThis.Word;
  return word.run(async (context) => {
    const selection = context.document.getSelection();
    const control = selection.parentContentControlOrNullObject;
    control.load("isNullObject,tag");
    await context.sync();
    if (control.isNullObject) throw new Error("NoCitationControlSelected");
    const range = control.getRange("Content");
    range.load("text");
    await context.sync();
    return { tag: control.tag, visibleText: range.text };
  });
}

async function customXmlCandidates() {
  const parts = await officeCall((done) =>
    Office.context.document.customXmlParts.getByNamespaceAsync(REFMGR_CUSTOM_XML_NAMESPACE, done));
  return Promise.all(parts.map(async (part) => ({
    id: part.id,
    xml: await officeCall((done) => part.getXmlAsync(done)),
  })));
}

/** Raw host evidence for `openCitationForEdit`; it makes no payload choice itself. */
export async function readSelectedCitationFromOffice() {
  const [control, customXmlParts] = await Promise.all([selectedCitationControl(), customXmlCandidates()]);
  return { control, customXmlParts };
}

export function officeCitationEditPort() {
  const insertion = officeInsertionPort();
  return {
    addCustomXml: insertion.addCustomXml,
    deleteCustomXml: insertion.deleteCustomXml,
    async replaceContentControl(expectedTag, descriptor) {
      const word = globalThis.Word;
      await word.run(async (context) => {
        const control = context.document.getSelection().parentContentControlOrNullObject;
        control.load("isNullObject,tag");
        await context.sync();
        if (control.isNullObject || control.tag !== expectedTag) throw new Error("SelectedCitationChanged");
        control.getRange("Content").insertText(descriptor.visibleText, word.InsertLocation.replace);
        control.tag = descriptor.tag;
        control.title = descriptor.title;
        control.appearance = descriptor.appearance;
        await context.sync();
      });
    },
  };
}

/** The task pane supplies the pure, validated plan and old part ids. */
export async function writeInsertionToOffice(plan, stalePartIds = []) {
  const port = officeInsertionPort();
  let added;
  try {
    added = await port.addCustomXml(plan.customXml);
  } catch {
    return { status: "failed", stage: "custom-xml", cleanupRequiredPartIds: [] };
  }
  try {
    await port.insertContentControl(plan.control);
  } catch {
    try {
      await port.deleteCustomXml(added.id);
      return { status: "failed", stage: "content-control", cleanupRequiredPartIds: [] };
    } catch {
      return { status: "failed", stage: "content-control", cleanupRequiredPartIds: [added.id] };
    }
  }
  const failed = [];
  for (const id of stalePartIds) {
    if (id === added.id) continue;
    try { await port.deleteCustomXml(id); } catch { failed.push(id); }
  }
  return failed.length === 0
    ? { status: "written", customXmlPartId: added.id }
    : { status: "written-with-cleanup-required", customXmlPartId: added.id, stalePartIds: failed };
}

/** The task pane supplies the composed E14-04.3 plan; no annotation id crosses into Word. */
export async function writeAnnotationInsertionToOffice(plan, stalePartIds = []) {
  const port = officeAnnotationInsertionPort();
  let added;
  try {
    added = await port.addCustomXml(plan.insertion.customXml);
  } catch {
    return { status: "failed", stage: "custom-xml", cleanupRequiredPartIds: [] };
  }
  try {
    await port.insertAnnotationContent({
      proseRuns: plan.proseRuns,
      control: plan.insertion.control,
    });
  } catch {
    try {
      await port.deleteCustomXml(added.id);
      return { status: "failed", stage: "content-control", cleanupRequiredPartIds: [] };
    } catch {
      return { status: "failed", stage: "content-control", cleanupRequiredPartIds: [added.id] };
    }
  }
  const failed = [];
  for (const id of stalePartIds) {
    if (id === added.id) continue;
    try { await port.deleteCustomXml(id); } catch { failed.push(id); }
  }
  return failed.length === 0
    ? { status: "written", customXmlPartId: added.id }
    : { status: "written-with-cleanup-required", customXmlPartId: added.id, stalePartIds: failed };
}

const BIBLIOGRAPHY_TAG = "refmgr-bib";
const BIBLIOGRAPHY_TAG_PREFIX = `${BIBLIOGRAPHY_TAG}:`;

function isBibliographyTag(tag) {
  return tag === BIBLIOGRAPHY_TAG || tag.startsWith(BIBLIOGRAPHY_TAG_PREFIX);
}

/**
 * Read the bibliography as the document actually holds it (§9.4).
 *
 * The paragraph texts are what a person would have edited by hand, so they are read from the
 * control rather than assumed from the payload — the payload cannot know what somebody typed over
 * it. An absent control returns `undefined` rather than an empty list: no bibliography and a
 * bibliography that was emptied are different documents.
 */
export async function readBibliographyFromOffice() {
  const word = globalThis.Word;
  const bibliography = await word.run(async (context) => {
    const controls = context.document.contentControls;
    controls.load("items/tag");
    await context.sync();
    const matches = controls.items.filter((control) => isBibliographyTag(control.tag));
    if (matches.length === 0) return { controlCount: 0, documentEntryTexts: undefined, controls: [] };
    const paragraphs = matches.map((control) => control.paragraphs);
    for (const collection of paragraphs) collection.load("items/text");
    await context.sync();
    const live = matches.map((control, index) => ({
      tag: control.tag,
      entryTexts: paragraphs[index].items.map((paragraph) => paragraph.text),
    }));
    return {
      controlCount: live.length,
      documentEntryTexts: live[0].entryTexts,
      controls: live,
    };
  });
  const customXmlParts = await customXmlCandidates();
  return { ...bibliography, customXmlParts };
}

/**
 * Read every Reference Manager citation in the order Word exposes it, plus the bibliography and
 * every candidate payload. The domain refresh planner, not this adapter, decides what is safe.
 */
export async function readDocumentForRefreshFromOffice() {
  const word = globalThis.Word;
  const controls = await word.run(async (context) => {
    const all = context.document.contentControls;
    all.load("items/tag");
    await context.sync();
    const citations = all.items.filter((control) => control.tag.startsWith(CLUSTER_TAG_PREFIX));
    const ranges = citations.map((control) => control.getRange("Content"));
    for (const range of ranges) range.load("text");
    await context.sync();
    return citations.map((control, index) => ({ tag: control.tag, visibleText: ranges[index].text }));
  });
  const [{ documentEntryTexts, controls: documentBibliographies }, customXmlParts] = await Promise.all([
    readBibliographyFromOffice(),
    customXmlCandidates(),
  ]);
  return { controls, documentEntryTexts, documentBibliographies, customXmlParts };
}

/**
 * Raw, read-only evidence for E10-10. The domain scanner chooses no payload and performs no repair.
 * Text boxes have no Office.js API, so object inspection is explicitly unmeasured rather than
 * reported clear. Citation and bibliography text leave this adapter only for in-memory hashing and
 * comparison; the structural diagnostic report never returns them (§23).
 */
export async function readDocumentDiagnosticsFromOffice(capabilities) {
  const [citationEvidence, bibliography] = await Promise.all([
    readCitationOccurrencesFromOffice(capabilities),
    readBibliographyFromOffice(),
  ]);
  return {
    occurrences: citationEvidence.occurrences,
    notes: citationEvidence.notes,
    noteNumberSource: citationEvidence.noteNumberSource,
    customXmlParts: bibliography.customXmlParts,
    bibliography: {
      controlCount: bibliography.controlCount,
      entryTexts: bibliography.documentEntryTexts,
    },
    unsupportedObjectInspection: "unmeasured",
  };
}

export function officeBibliographyPort() {
  const insertion = officeInsertionPort();
  return {
    addCustomXml: insertion.addCustomXml,
    deleteCustomXml: insertion.deleteCustomXml,
    async writeBibliographyControl(control) {
      const word = globalThis.Word;
      await word.run(async (context) => {
        const existing = context.document.contentControls.getByTag(BIBLIOGRAPHY_TAG);
        existing.load("items/tag");
        await context.sync();

        const body = control.entries.map((entry) => entry.text).join("\r");
        let target;
        if (existing.items.length > 0) {
          target = existing.items[0];
          target.getRange("Content").insertText(body, word.InsertLocation.replace);
        } else {
          const range = context.document.getSelection().insertText(body, word.InsertLocation.start);
          target = range.insertContentControl();
          target.tag = control.tag;
          target.title = control.title;
          target.appearance = control.appearance;
        }
        // A null paragraphFormat means the document already has formatting the author may have
        // chosen; §9.4 says preserve it, so nothing here touches a paragraph property at all.
        if (control.paragraphFormat !== null) {
          const paragraphs = target.paragraphs;
          paragraphs.load("items");
          await context.sync();
          for (const paragraph of paragraphs.items) {
            paragraph.style = control.paragraphFormat.styleName;
            paragraph.leftIndent = control.paragraphFormat.leftIndentTwips / 20;
            paragraph.firstLineIndent = -control.paragraphFormat.hangingIndentTwips / 20;
            paragraph.lineSpacing = control.paragraphFormat.lineTwips / 20;
            paragraph.spaceAfter = control.paragraphFormat.afterTwips / 20;
          }
        }
        await context.sync();
      });
    },
  };
}

/** One preflighted Word.run batch for every citation update and the existing bibliography. */
export function officeDocumentRefreshPort() {
  const insertion = officeInsertionPort();
  return {
    addCustomXml: insertion.addCustomXml,
    deleteCustomXml: insertion.deleteCustomXml,
    async writeDocumentControls(plan) {
      const word = globalThis.Word;
      await word.run(async (context) => {
        const all = context.document.contentControls;
        all.load("items/tag");
        await context.sync();

        const liveCitationTags = all.items
          .filter((control) => control.tag.startsWith(CLUSTER_TAG_PREFIX))
          .map((control) => control.tag);
        if (liveCitationTags.length !== plan.expectedCitationTags.length ||
            liveCitationTags.some((tag, index) => tag !== plan.expectedCitationTags[index])) {
          throw new Error("CitationWalkChanged");
        }

        const targets = [];
        for (const write of plan.citations) {
          const matches = all.items.filter((control) => control.tag === write.expectedTag);
          if (matches.length !== 1) throw new Error("CitationControlChanged");
          const range = matches[0].getRange("Content");
          range.load("text");
          targets.push({ write, control: matches[0], range });
        }

        const liveBibliographyTags = all.items.filter((control) => isBibliographyTag(control.tag)).map((control) => control.tag);
        if (liveBibliographyTags.length !== plan.expectedBibliographyTags.length ||
            liveBibliographyTags.some((tag, index) => tag !== plan.expectedBibliographyTags[index])) {
          throw new Error("BibliographyWalkChanged");
        }
        const bibliographyTargets = [];
        for (const write of plan.bibliographies) {
          const matches = all.items.filter((control) => control.tag === write.expectedTag);
          if (matches.length !== 1) throw new Error("BibliographyControlChanged");
          bibliographyTargets.push({ write, control: matches[0] });
        }
        await context.sync();

        for (const target of targets) {
          if (target.range.text !== target.write.control.visibleText) {
            target.range.insertText(target.write.control.visibleText, word.InsertLocation.replace);
          }
          target.control.tag = target.write.control.tag;
          target.control.title = target.write.control.title;
          target.control.appearance = target.write.control.appearance;
        }
        for (const target of bibliographyTargets) {
          const descriptor = target.write.control;
          target.control.getRange("Content").insertText(
            descriptor.entries.map((entry) => entry.text).join("\r"),
            word.InsertLocation.replace,
          );
          target.control.tag = descriptor.tag;
          target.control.title = descriptor.title;
          target.control.appearance = descriptor.appearance;
          // Refresh descriptors always carry null here (ADR-0109). Keep this guard explicit so a
          // later style-change plan can reuse the port without silently widening refresh writes.
          if (descriptor.paragraphFormat !== null) {
            const paragraphs = target.control.paragraphs;
            paragraphs.load("items");
            await context.sync();
            for (const paragraph of paragraphs.items) {
              paragraph.style = descriptor.paragraphFormat.styleName;
              paragraph.leftIndent = descriptor.paragraphFormat.leftIndentTwips / 20;
              paragraph.firstLineIndent = -descriptor.paragraphFormat.hangingIndentTwips / 20;
              paragraph.lineSpacing = descriptor.paragraphFormat.lineTwips / 20;
              paragraph.spaceAfter = descriptor.paragraphFormat.afterTwips / 20;
            }
          }
        }
        await context.sync();
      });
    },
  };
}

/** Commit an E10-07 plan with payload-first compensation and explicit stale cleanup. */
export async function writeDocumentRefreshToOffice(plan, stalePartIds = []) {
  const port = officeDocumentRefreshPort();
  let added;
  try {
    added = await port.addCustomXml(plan.customXml);
  } catch {
    return { status: "failed", stage: "custom-xml", cleanupRequiredPartIds: [] };
  }
  try {
    await port.writeDocumentControls(plan);
  } catch {
    try {
      await port.deleteCustomXml(added.id);
      return { status: "failed", stage: "document-controls", cleanupRequiredPartIds: [] };
    } catch {
      return { status: "failed", stage: "document-controls", cleanupRequiredPartIds: [added.id] };
    }
  }
  const failed = [];
  for (const id of stalePartIds) {
    if (id === added.id) continue;
    try { await port.deleteCustomXml(id); } catch { failed.push(id); }
  }
  return failed.length === 0
    ? { status: "written", customXmlPartId: added.id }
    : { status: "written-with-cleanup-required", customXmlPartId: added.id, stalePartIds: failed };
}

/**
 * Read what a preferences change needs and nothing more: the payloads (§9.1, E10-08).
 *
 * A style change reads the whole document through `readDocumentForRefreshFromOffice`, because it
 * rewrites every citation. Changing the update mode or locking the bibliography rewrites nothing,
 * so this deliberately never walks the content controls — a read that cannot see the manuscript
 * cannot damage it.
 */
export async function readDocumentPreferencesFromOffice() {
  return { customXmlParts: await customXmlCandidates() };
}

/**
 * Commit an E10-08 preferences plan: one custom XML part, no content control, no text.
 *
 * There is no compensation stage because there is no second write to compensate. If the payload
 * fails to land the document is exactly as it was, still describing itself with the old settings.
 */
export async function writeDocumentPreferencesToOffice(plan, stalePartIds = []) {
  const port = officeInsertionPort();
  let added;
  try {
    added = await port.addCustomXml(plan.customXml);
  } catch {
    return { status: "failed", stage: "custom-xml" };
  }
  const failed = [];
  for (const id of stalePartIds) {
    if (id === added.id) continue;
    try { await port.deleteCustomXml(id); } catch { failed.push(id); }
  }
  return failed.length === 0
    ? { status: "written", customXmlPartId: added.id }
    : { status: "written-with-cleanup-required", customXmlPartId: added.id, stalePartIds: failed };
}

/** Commit an E10-06 bibliography plan with the same payload-first compensation. */
export async function writeBibliographyToOffice(plan, stalePartIds = []) {
  const port = officeBibliographyPort();
  let added;
  try {
    added = await port.addCustomXml(plan.customXml);
  } catch {
    return { status: "failed", stage: "custom-xml", cleanupRequiredPartIds: [] };
  }
  try {
    await port.writeBibliographyControl(plan.control);
  } catch {
    try {
      await port.deleteCustomXml(added.id);
      return { status: "failed", stage: "content-control", cleanupRequiredPartIds: [] };
    } catch {
      return { status: "failed", stage: "content-control", cleanupRequiredPartIds: [added.id] };
    }
  }
  const failed = [];
  for (const id of stalePartIds) {
    if (id === added.id) continue;
    try { await port.deleteCustomXml(id); } catch { failed.push(id); }
  }
  return failed.length === 0
    ? { status: "written", customXmlPartId: added.id }
    : { status: "written-with-cleanup-required", customXmlPartId: added.id, stalePartIds: failed };
}

const NOTE_FEATURE_ID = "footnotes";

/**
 * Whether this host has been **measured** able to reach footnotes and endnotes (WordApi 1.5).
 *
 * The manifest loads the add-in at WordApi 1.1 on purpose (ADR-0101), so everything above that
 * floor is a run-time question. `unknown` is not permission: a probe that never asked is not a
 * host that said yes, and attempting a note operation on a host without the API throws in front of
 * the author in the middle of their manuscript.
 */
function noteApiAvailable(capabilities) {
  const verdicts = Array.isArray(capabilities) ? capabilities : [];
  const verdict = verdicts.find((entry) => entry && entry.featureId === NOTE_FEATURE_ID);
  return verdict !== undefined && verdict.availability === "available";
}

/**
 * Read every citation occurrence with the story it sits in, and every note that holds one.
 *
 * **Office.js exposes no note number.** `Word.NoteItem` carries the note's body and its reference
 * range, not the number Word prints. The number reported here is therefore the item's position in
 * `body.footnotes` / `body.endnotes`, which is Word's own document order — correct for a document
 * that numbers continuously from 1, and wrong for one that restarts numbering per section or uses
 * custom marks. That is why the result carries `noteNumberSource`: the domain refuses a missing
 * number rather than synthesising one, and this is the one place a number is derived, so it says so
 * (risk R-067).
 */
export async function readCitationOccurrencesFromOffice(capabilities) {
  const word = globalThis.Word;
  const notesReadable = noteApiAvailable(capabilities);
  return word.run(async (context) => {
    const all = context.document.contentControls;
    all.load("items/tag");
    await context.sync();

    const citations = all.items.filter((control) => control.tag.startsWith(CLUSTER_TAG_PREFIX));
    const ranges = citations.map((control) => control.getRange("Content"));
    for (const range of ranges) range.load("text");

    let footnotes = null;
    let endnotes = null;
    if (notesReadable) {
      footnotes = context.document.body.footnotes;
      endnotes = context.document.body.endnotes;
      footnotes.load("items");
      endnotes.load("items");
    }
    await context.sync();

    const occurrences = citations.map((control, index) => ({
      occurrenceId: `cc-${index}`,
      tag: control.tag,
      visibleText: ranges[index].text,
      story: "body",
      container: "text",
      revision: "current",
    }));

    const notes = [];
    if (notesReadable) {
      for (const [story, collection] of [["footnote", footnotes], ["endnote", endnotes]]) {
        for (const [position, item] of collection.items.entries()) {
          const controls = item.body.contentControls;
          controls.load("items/tag");
          const noteText = item.body.getRange("Content");
          noteText.load("text");
          await context.sync();

          const inside = controls.items.filter((control) => control.tag.startsWith(CLUSTER_TAG_PREFIX));
          if (inside.length === 0) continue;
          const insideRanges = inside.map((control) => control.getRange("Content"));
          for (const range of insideRanges) range.load("text");
          await context.sync();

          const noteNumber = position + 1;
          const insideIds = [];
          for (const control of inside) {
            const occurrence = occurrences.find((entry) => entry.tag === control.tag);
            if (occurrence === undefined) continue;
            occurrence.story = story;
            occurrence.noteNumber = noteNumber;
            insideIds.push(occurrence.occurrenceId);
          }
          // Whether the author wrote anything in this note besides its citations, measured by
          // length so that no note text ever leaves this function (§23, invariant 6).
          const citationLength = insideRanges.reduce((total, range) => total + range.text.trim().length, 0);
          notes.push({
            story,
            noteNumber,
            occurrenceIds: insideIds,
            hasOtherContent: noteText.text.trim().length > citationLength,
          });
        }
      }
    }

    return {
      occurrences,
      notes,
      notesReadable,
      noteNumberSource: notesReadable ? "collection-order" : null,
    };
  });
}

/**
 * Apply an E10-09.2 relocation: move citations into or out of notes, in one preflighted batch.
 *
 * Every move is checked against the tag the plan expects **before** anything is written, for the
 * reason E10-07 gives: half a relocation is a manuscript with some citations in footnotes and some
 * in the text, and no record of which. A note is deleted only when the plan says it held citations
 * and nothing else.
 */
export async function writeNoteRelocationToOffice(plan, capabilities) {
  if (!noteApiAvailable(capabilities)) {
    return { status: "failed", stage: "capability" };
  }
  const word = globalThis.Word;
  try {
    await word.run(async (context) => {
      const all = context.document.contentControls;
      all.load("items/tag");
      await context.sync();

      const targets = [];
      for (const move of plan.moves) {
        const matches = all.items.filter((control) => control.tag === move.control.tag);
        if (matches.length !== 1) throw new Error("CitationControlChanged");
        targets.push({ move, control: matches[0] });
      }

      // The note collection is loaded up front for both directions: moving out of notes needs each
      // note's reference range in the body, and removing an emptied note needs the same items.
      const outward = plan.direction === "note-to-in-text";
      const usesEndnotes = plan.moves.some((move) => move.from === "endnote") || plan.noteStory === "endnote";
      const collection = usesEndnotes ? context.document.body.endnotes : context.document.body.footnotes;
      if (outward || plan.removeNoteNumbers.length > 0) collection.load("items");
      await context.sync();

      for (const target of targets) {
        if (target.move.to === "body") {
          // Out of a note: the destination is the note's *reference* in the body, not anywhere
          // inside the note. The note itself is removed below and only when the plan named it —
          // a note the author also wrote in is theirs to keep.
          const item = collection.items[target.move.fromNoteNumber - 1];
          if (item === undefined) throw new Error("NoteItemChanged");
          const range = item.reference.insertText(target.move.control.visibleText, word.InsertLocation.after);
          const inserted = range.insertContentControl();
          inserted.tag = target.move.control.tag;
          inserted.title = target.move.control.title;
          inserted.appearance = target.move.control.appearance;
          target.control.delete(false);
          continue;
        }
        const note = target.control.getRange("Content").insertFootnote(target.move.control.visibleText);
        const inserted = note.body.getRange("Content").insertContentControl();
        inserted.tag = target.move.control.tag;
        inserted.title = target.move.control.title;
        inserted.appearance = target.move.control.appearance;
        // `false` removes the control *and* the text it held; `true` would leave the citation
        // behind in the body as loose text, once in the note and once out of it.
        target.control.delete(false);
      }
      await context.sync();

      // Highest number first, so removing one does not shift the index of the next.
      for (const number of [...plan.removeNoteNumbers].sort((a, b) => b - a)) {
        const item = collection.items[number - 1];
        if (item !== undefined) item.delete();
      }
      await context.sync();
    });
  } catch {
    return { status: "failed", stage: "relocation" };
  }
  return { status: "relocated", requiresRefresh: true };
}

/** Commit an E10-05 plan without ever replacing text outside the selected citation control. */
export async function writeCitationEditToOffice(expectedTag, prepared) {
  const port = officeCitationEditPort();
  let added;
  try {
    added = await port.addCustomXml(prepared.plan.customXml);
  } catch {
    return { status: "failed", stage: "custom-xml", cleanupRequiredPartIds: [] };
  }
  try {
    await port.replaceContentControl(expectedTag, prepared.plan.control);
  } catch {
    try {
      await port.deleteCustomXml(added.id);
      return { status: "failed", stage: "content-control", cleanupRequiredPartIds: [] };
    } catch {
      return { status: "failed", stage: "content-control", cleanupRequiredPartIds: [added.id] };
    }
  }
  const failed = [];
  for (const id of prepared.stalePartIds) {
    if (id === added.id) continue;
    try { await port.deleteCustomXml(id); } catch { failed.push(id); }
  }
  return failed.length === 0
    ? { status: "written", customXmlPartId: added.id }
    : { status: "written-with-cleanup-required", customXmlPartId: added.id, stalePartIds: failed };
}

/**
 * E10-11: the document's own bytes, for the backup a repair may not proceed without.
 *
 * `getFileAsync` hands back a compressed `.docx` in slices and the slices must be reassembled in
 * order — Office does not promise the callbacks arrive in order, so they are placed by index rather
 * than pushed. The file handle is closed in a `finally`: an unclosed handle is Word's own document
 * held open, and leaking one on the failure path is how a user ends up unable to save.
 *
 * **This is only the reading half.** An Office task pane has no filesystem, so the copy is persisted
 * by the desktop side through the ADR-0008 bridge, which does not have that route yet (E10-11.1).
 * Until it does, `persistBackup` has nothing to call and `commitDocumentRepair` refuses — which is
 * the correct outcome, not a gap: no verified backup means no write.
 */
export async function readDocumentBytesFromOffice(sliceSize = 4194304) {
  const file = await officeCall((done) =>
    Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize }, done));
  try {
    const slices = new Array(file.sliceCount);
    let total = 0;
    for (let index = 0; index < file.sliceCount; index += 1) {
      const slice = await officeCall((done) => file.getSliceAsync(index, done));
      const bytes = slice.data instanceof Uint8Array ? slice.data : new Uint8Array(slice.data);
      slices[slice.index === undefined ? index : slice.index] = bytes;
      total += bytes.length;
    }
    const document = new Uint8Array(total);
    let offset = 0;
    for (const bytes of slices) {
      if (bytes === undefined) throw new Error("DocumentSliceMissing");
      document.set(bytes, offset);
      offset += bytes.length;
    }
    return document;
  } finally {
    try { await officeCall((done) => file.closeAsync(done)); } catch { /* the read already happened */ }
  }
}

/** `cc-3` → 3. The occurrence ids `readCitationOccurrencesFromOffice` mints are positional. */
function occurrencePosition(occurrenceId) {
  const match = /^cc-(\d+)$/.exec(String(occurrenceId));
  if (match === null) throw new Error("UnknownOccurrenceId");
  return Number(match[1]);
}

function bibliographyPosition(occurrenceId) {
  const match = /^bib-(\d+)$/.exec(String(occurrenceId));
  if (match === null) throw new Error("UnknownBibliographyOccurrenceId");
  return Number(match[1]);
}

/**
 * Read the structural evidence E10-12's planner needs. The document walk is explicitly
 * `unmeasured`: Office.js has no text-box collection, so this adapter cannot claim whole-document
 * coverage. Reference and citation text is never logged or rendered by the unlink surface.
 */
export async function readDocumentUnlinkEvidenceFromOffice(capabilities) {
  const [citationEvidence, customXmlParts] = await Promise.all([
    readCitationOccurrencesFromOffice(capabilities),
    customXmlCandidates(),
  ]);
  const word = globalThis.Word;
  const bibliographyControls = await word.run(async (context) => {
    const controls = context.document.contentControls.getByTag(BIBLIOGRAPHY_TAG);
    controls.load("items/tag");
    await context.sync();
    return controls.items.map((control, index) => ({
      occurrenceId: `bib-${index}`,
      tag: control.tag,
    }));
  });
  return {
    occurrences: citationEvidence.occurrences,
    bibliographyControls,
    customXmlPartIds: customXmlParts.map((part) => part.id),
    walkCoverage: "unmeasured",
  };
}

/**
 * The Word half of E10-12's unlink port. Every collection is preflighted before a delete is
 * queued, and `delete(true)` removes only the wrapper: the author's visible text and formatting
 * stay in place. Backup persistence is the same verified bridge store used by repair.
 */
export function officeDocumentUnlinkPort(backupStore) {
  const insertion = officeInsertionPort();

  async function removeControls(kind, writes, expectedTags) {
    const word = globalThis.Word;
    await word.run(async (context) => {
      const collection = kind === "citation"
        ? context.document.contentControls
        : context.document.contentControls.getByTag(BIBLIOGRAPHY_TAG);
      collection.load("items/tag");
      await context.sync();
      const live = kind === "citation"
        ? collection.items.filter((control) => control.tag.startsWith(CLUSTER_TAG_PREFIX))
        : collection.items;
      if (live.length !== expectedTags.length ||
          live.some((control, index) => control.tag !== expectedTags[index])) {
        throw new Error(kind === "citation" ? "CitationWalkChanged" : "BibliographyWalkChanged");
      }
      const targets = writes.map((write) => {
        const position = kind === "citation"
          ? occurrencePosition(write.occurrenceId)
          : bibliographyPosition(write.occurrenceId);
        const control = live[position];
        if (control === undefined || control.tag !== write.expectedTag) throw new Error("ControlTagChanged");
        return control;
      });
      for (const control of targets) control.delete(true);
      await context.sync();
    });
  }

  return {
    readDocumentBytes: () => readDocumentBytesFromOffice(),
    persistBackup: (bytes) => backupStore.persistBackup(bytes),
    readBackBackup: (location) => backupStore.readBackBackup(location),
    removeCitationControls: (writes, expectedTags) => removeControls("citation", writes, expectedTags),
    removeBibliographyControls: (writes, expectedTags) => removeControls("bibliography", writes, expectedTags),
    deleteCustomXml: insertion.deleteCustomXml,
  };
}

/**
 * The Word half of the repair port. Backup persistence is deliberately not defaulted to a no-op:
 * a port whose `persistBackup` silently succeeded would let a repair write with no copy behind it,
 * which is the one failure this whole task exists to make impossible.
 */
export function officeDocumentRepairPort(backupStore) {
  const insertion = officeInsertionPort();
  return {
    readDocumentBytes: () => readDocumentBytesFromOffice(),
    persistBackup: (bytes) => backupStore.persistBackup(bytes),
    readBackBackup: (location) => backupStore.readBackBackup(location),
    addCustomXml: insertion.addCustomXml,
    deleteCustomXml: insertion.deleteCustomXml,
    async writeCitationTags(writes, expectedTags) {
      const word = globalThis.Word;
      await word.run(async (context) => {
        const all = context.document.contentControls;
        all.load("items/tag");
        await context.sync();

        const live = all.items.filter((control) => control.tag.startsWith(CLUSTER_TAG_PREFIX));
        if (live.length !== expectedTags.length ||
            live.some((control, index) => control.tag !== expectedTags[index])) {
          throw new Error("CitationWalkChanged");
        }
        for (const write of writes) {
          // Addressed by position, never by tag value. The two controls a re-key exists to separate
          // carry the *same* tag — that is the damage — so a tag lookup would find the original and
          // leave the copy, quietly repairing the wrong one.
          const position = occurrencePosition(write.occurrenceId);
          const control = live[position];
          if (control === undefined || control.tag !== write.expectedTag) throw new Error("ControlTagChanged");
          // Only the tag changes. A re-key gives a copied citation its own identity; its text is
          // the author's document and is never rewritten here.
          control.tag = write.control.tag;
        }
        await context.sync();
      });
    },
  };
}
