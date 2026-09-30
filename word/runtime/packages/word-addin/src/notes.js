/**
 * Footnote and endnote citations: the note-number walk, and in-text ⇄ note relocation
 * (SPEC §9.5, §9.1; task E10-09.2; ADR-0113).
 *
 * E10-09.1 reconciled *which* citations exist. This module answers the question that only matters
 * for note styles: **where each one sits, and what number Word gave the note it sits in.**
 *
 * Two things make this different from every other Word surface built so far.
 *
 * **Word owns note numbering, and nothing here may invent it.** A note number decides `ibid`,
 * `ibid-with-locator` and subsequent-note short forms, so a wrong number does not produce a
 * visibly broken citation — it produces a plausible "Ibid." pointing at the wrong work.
 * `resolveNoteIndexes` in `@refmgr/citation` already refuses a backwards sequence and warns when it
 * has to synthesise a number from document position; this module makes sure a *synthesised* number
 * never reaches it in the first place, because a synthesised note number in a real Word document is
 * not a fallback, it is a fabrication.
 *
 * **A relocation cannot predict the numbers it will produce.** Moving citations into footnotes
 * renumbers every note after the first one, interleaved with whatever notes the author already
 * wrote, and no code outside Word can say what the result is. So relocation is deliberately a
 * **two-phase operation**: move the citations, let Word number the notes, then read the walk back
 * and refresh. The plan says so in its type (`requiresRefreshAfterMove: true`) and clears every
 * stored note number rather than carrying a guess forward.
 */
import { wordCitationControl, } from "@refmgr/doc-schema";
/** A story that can hold a citation but can never host a note of its own. */
const NOTE_HOSTILE_STORIES = ["header", "footer", "text-box"];
export function isNoteStory(story) {
    return story === "footnote" || story === "endnote";
}
function refuseWalk(code, message) {
    return { ok: false, refusal: { code, message } };
}
/**
 * Turn the host's note evidence into the numbers a render may use, or refuse.
 *
 * Every refusal here describes a document the host could not actually be looking at. That is
 * deliberate: the alternative to refusing is handing citeproc a sequence it will happily render
 * from, and the result of a wrong sequence is a correct-looking short form.
 */
export function readNoteWalk(input) {
    const noteByOccurrence = new Map();
    const occurrenceIds = new Set(input.occurrences.map((occurrence) => occurrence.occurrenceId));
    for (const note of input.notes) {
        for (const occurrenceId of note.occurrenceIds) {
            if (!occurrenceIds.has(occurrenceId)) {
                return refuseWalk("note-membership-mismatch", `${note.story} ${note.noteNumber} names occurrence ${occurrenceId}, which is not in the citation walk`);
            }
            if (noteByOccurrence.has(occurrenceId)) {
                return refuseWalk("note-membership-mismatch", `occurrence ${occurrenceId} is reported inside more than one note`);
            }
            noteByOccurrence.set(occurrenceId, note);
        }
    }
    const placements = [];
    const noteNumbers = [];
    const inNoteClusterIds = [];
    const outsideNoteClusterIds = [];
    const notices = [];
    const stories = new Set();
    for (const occurrence of input.occurrences) {
        const clusterId = input.clusterIdByOccurrenceId.get(occurrence.occurrenceId);
        if (clusterId === undefined) {
            return refuseWalk("note-membership-mismatch", `occurrence ${occurrence.occurrenceId} has no cluster; reconcile the walk before reading note numbers`);
        }
        const note = noteByOccurrence.get(occurrence.occurrenceId);
        if (!isNoteStory(occurrence.story)) {
            // A number outside a note is not a harmless extra field: it is the host telling us it does
            // not know where this citation is, and a number is exactly what would be believed.
            if (occurrence.noteNumber !== undefined) {
                return refuseWalk("note-number-outside-note", `occurrence ${occurrence.occurrenceId} is in the ${occurrence.story} story but carries a note number`);
            }
            if (note !== undefined) {
                return refuseWalk("note-membership-mismatch", `occurrence ${occurrence.occurrenceId} is reported inside ${note.story} ${note.noteNumber} but its story is ${occurrence.story}`);
            }
            if (input.citationFormat === "note") {
                return refuseWalk("citation-outside-note", `cluster ${clusterId} is in the ${occurrence.story} story, but this document's style renders citations as notes; relocate it with planNoteRelocation`);
            }
            placements.push({ clusterId, occurrenceId: occurrence.occurrenceId, story: occurrence.story });
            outsideNoteClusterIds.push(clusterId);
            continue;
        }
        if (occurrence.noteNumber === undefined) {
            return refuseWalk("note-number-missing", `occurrence ${occurrence.occurrenceId} is in a ${occurrence.story} the host reported no number for`);
        }
        if (note !== undefined && (note.story !== occurrence.story || note.noteNumber !== occurrence.noteNumber)) {
            return refuseWalk("note-membership-mismatch", `occurrence ${occurrence.occurrenceId} reports ${occurrence.story} ${occurrence.noteNumber} but is listed under ${note.story} ${note.noteNumber}`);
        }
        stories.add(occurrence.story);
        if (stories.size > 1) {
            return refuseWalk("mixed-note-stories", "citations sit in both footnotes and endnotes; a single note sequence cannot describe both");
        }
        placements.push({
            clusterId,
            occurrenceId: occurrence.occurrenceId,
            story: occurrence.story,
            noteNumber: occurrence.noteNumber,
        });
        noteNumbers.push({ clusterId, noteNumber: occurrence.noteNumber });
        inNoteClusterIds.push(clusterId);
        if (input.citationFormat === "in-text") {
            notices.push({
                code: "citation-inside-author-note",
                clusterId,
                message: `cluster ${clusterId} sits in ${occurrence.story} ${occurrence.noteNumber}; this document's style renders citations in text, so the number is preserved and not used`,
            });
        }
    }
    // One footnote may hold several citations, so equal numbers are fine. Requiring the sequence to
    // be non-decreasing is the whole check: it also forces citations that share a number to be
    // adjacent, because a number could only reappear later by first going down.
    let previous;
    for (const entry of noteNumbers) {
        if (previous !== undefined && entry.noteNumber < previous) {
            return refuseWalk("note-number-out-of-order", `note number ${entry.noteNumber} follows note number ${previous} in document order`);
        }
        previous = entry.noteNumber;
    }
    const noteStory = [...stories][0] ?? null;
    return {
        ok: true,
        walk: Object.freeze({
            noteStory,
            placements: Object.freeze(placements),
            noteNumbers: Object.freeze(noteNumbers.map((entry) => Object.freeze({ ...entry }))),
            inNoteClusterIds: Object.freeze(inNoteClusterIds),
            outsideNoteClusterIds: Object.freeze(outsideNoteClusterIds),
            notices: Object.freeze(notices),
        }),
    };
}
function refuseRelocation(code, message) {
    return { ok: false, refusal: { code, message } };
}
/** The feature id in `WORD_ADDIN_FEATURES` that gates every note operation. */
export const NOTE_FEATURE_ID = "footnotes";
/**
 * Plan the structural half of a class change.
 *
 * Nothing here renders or renumbers. It says which citation moves where, which notes may go with
 * it, and which must not — and it refuses outright where Word cannot do the thing at all, rather
 * than emitting a move the host will fail half-way through.
 */
export function planNoteRelocation(input) {
    if (input.from === input.to) {
        return refuseRelocation("no-class-change", `this document already renders ${input.from} citations`);
    }
    const verdict = input.capabilities.find((entry) => entry.featureId === NOTE_FEATURE_ID);
    if (verdict === undefined || verdict.availability === "unknown") {
        return refuseRelocation("host-capability-unproven", `this host has not been measured for ${NOTE_FEATURE_ID}; run the capability probe before relocating citations`);
    }
    if (verdict.availability === "unavailable") {
        return refuseRelocation("host-capability-missing", `this host does not support ${NOTE_FEATURE_ID}, so citations cannot be moved into or out of notes`);
    }
    const direction = input.to === "note" ? "in-text-to-note" : "note-to-in-text";
    if (direction === "in-text-to-note" && input.noteStory === undefined) {
        return refuseRelocation("note-story-required", "relocating into notes requires a footnote or endnote story");
    }
    // The walk is read against the class the document has *now*: a note-class document being made
    // in-text must already satisfy the note rules, and an in-text one must not be assumed to.
    const walk = readNoteWalk({
        occurrences: input.occurrences,
        notes: input.notes,
        citationFormat: input.from,
        clusterIdByOccurrenceId: input.clusterIdByOccurrenceId,
    });
    if (!walk.ok) {
        return refuseRelocation("note-walk-refused", `the note walk refused: ${walk.refusal.code} — ${walk.refusal.message}`);
    }
    const payloadIds = new Set(input.part.clusters.map((cluster) => cluster.clusterId));
    const walkedIds = new Set(walk.walk.placements.map((placement) => placement.clusterId));
    for (const clusterId of payloadIds) {
        if (!walkedIds.has(clusterId)) {
            return refuseRelocation("relocation-coverage-incomplete", `cluster ${clusterId} is in the payload but not in the live walk`);
        }
    }
    for (const clusterId of walkedIds) {
        if (!payloadIds.has(clusterId)) {
            return refuseRelocation("relocation-coverage-incomplete", `cluster ${clusterId} is in the live walk but not in the payload`);
        }
    }
    const clusterById = new Map(input.part.clusters.map((cluster) => [cluster.clusterId, cluster]));
    const noteByNumber = new Map(input.notes.map((note) => [note.noteNumber, note]));
    const moves = [];
    const notices = [...walk.walk.notices];
    const emptied = new Set();
    const retained = new Set();
    const clearedNoteNumberClusterIds = [];
    for (const placement of walk.walk.placements) {
        const cluster = clusterById.get(placement.clusterId);
        if (cluster === undefined) {
            return refuseRelocation("relocation-coverage-incomplete", `cluster ${placement.clusterId} is in the live walk but not in the payload`);
        }
        if (cluster.noteNumber !== undefined)
            clearedNoteNumberClusterIds.push(cluster.clusterId);
        if (direction === "in-text-to-note") {
            if (isNoteStory(placement.story))
                continue; // already in a note the author wrote
            if (NOTE_HOSTILE_STORIES.includes(placement.story)) {
                return refuseRelocation("note-unsupported-story", `cluster ${placement.clusterId} is in the ${placement.story} story, where Word cannot place a note`);
            }
            moves.push({
                clusterId: placement.clusterId,
                occurrenceId: placement.occurrenceId,
                from: placement.story,
                to: input.noteStory,
                // The stored note number would travel with the descriptor and become a claim about a
                // sequence that has not happened yet.
                control: wordCitationControl(withoutNoteNumber(cluster)),
            });
            continue;
        }
        if (!isNoteStory(placement.story))
            continue; // already in the body
        const note = placement.noteNumber === undefined ? undefined : noteByNumber.get(placement.noteNumber);
        moves.push({
            clusterId: placement.clusterId,
            occurrenceId: placement.occurrenceId,
            from: placement.story,
            to: "body",
            control: wordCitationControl(withoutNoteNumber(cluster)),
            ...(placement.noteNumber === undefined ? {} : { fromNoteNumber: placement.noteNumber }),
        });
        if (placement.noteNumber === undefined)
            continue;
        if (note !== undefined && note.hasOtherContent) {
            retained.add(placement.noteNumber);
            notices.push({
                code: "note-retained-with-author-content",
                clusterId: placement.clusterId,
                message: `${placement.story} ${placement.noteNumber} holds content besides its citations and is kept`,
            });
        }
        else {
            emptied.add(placement.noteNumber);
        }
    }
    if (clearedNoteNumberClusterIds.length > 0) {
        notices.push({
            code: "note-numbers-cleared-pending-refresh",
            message: `${clearedNoteNumberClusterIds.length} stored note number(s) were cleared; Word assigns the new ones`,
        });
    }
    notices.push({
        code: "render-provisional-until-refresh",
        message: "citation text rendered before this move used note numbers Word has not yet assigned; refresh after the move",
    });
    return {
        ok: true,
        plan: Object.freeze({
            direction,
            noteStory: direction === "in-text-to-note" ? input.noteStory : null,
            moves: Object.freeze(moves),
            // A note that both holds author content and is listed for removal is a contradiction; the
            // author's content wins, so the retained set is subtracted rather than merged.
            removeNoteNumbers: Object.freeze([...emptied].filter((number) => !retained.has(number)).sort((a, b) => a - b)),
            retainedNoteNumbers: Object.freeze([...retained].sort((a, b) => a - b)),
            clearedNoteNumberClusterIds: Object.freeze(clearedNoteNumberClusterIds),
            requiresRefreshAfterMove: true,
            notices: Object.freeze(notices),
        }),
    };
}
function withoutNoteNumber(cluster) {
    if (cluster.noteNumber === undefined)
        return cluster;
    const { noteNumber: _dropped, ...rest } = cluster;
    return rest;
}
/**
 * Strip stored note numbers from a payload a relocation is about to move.
 *
 * Separate from the plan so that the payload write and the structural move cannot disagree: the
 * plan names the clusters, and this is the only thing that edits them.
 */
export function clearNoteNumbers(part, clusterIds) {
    const targets = new Set(clusterIds);
    return {
        ...part,
        clusters: part.clusters.map((cluster) => (targets.has(cluster.clusterId) ? withoutNoteNumber(cluster) : cluster)),
    };
}
