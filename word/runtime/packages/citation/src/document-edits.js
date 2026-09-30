/**
 * Citation insertion and deletion — SPEC §8.2 and §8.3, task E08-07.
 *
 * Inserting or deleting a citation in a numeric document changes the text of citations the
 * user never touched: put a new reference at the top of a Vancouver manuscript and every
 * `[n]` below it moves up by one. The renumbering itself is not implemented here and must
 * never be — citation numbers, `1–3` range collapsing and every other positional rule are
 * produced by the CSL processor under the style's own `<citation>` rules (invariant 4).
 * What this module owns is the part the processor cannot know:
 *
 * **Which citations in the document now show the wrong text.** The host has to rewrite exactly
 * those content controls and no others. Rewriting too few leaves stale numbers in the
 * manuscript — the single most damaging failure a citation manager has, because the document
 * still looks finished. Rewriting everything is safe but destroys Track Changes review and, at
 * §21's thousand-citation scale, is the difference between an edit and a pause.
 *
 * So an edit is *planned*, not applied: the whole document is re-rendered through one processor
 * (`CitationClusterRenderer`, which already guarantees a walk is collectively correct rather
 * than individually plausible), the result is compared against the text the document records
 * itself as showing, and the caller is handed the difference. Three outcomes are kept apart
 * rather than merged, because they oblige the host differently:
 *
 * - `text-changed` — we know what it showed and it is now wrong;
 * - `no-recorded-text` — the document never recorded what this citation shows, so nobody can
 *   say whether it changed. Folding this into "unchanged" would leave a citation unwritten on
 *   the strength of a comparison that never happened;
 * - `inserted` — new, and has no previous text by definition.
 *
 * Two things are deliberately *not* done. A citation carrying a manual edit (§9.3) is never
 * proposed for overwriting, however wrong its number now is; it is reported separately so the
 * §9.4 warning can be raised, because silently restoring our text over the user's writing is
 * data loss. And deleting the last citation of a work never deletes its snapshot: §9.4 lets a
 * bibliography entry be pinned or uncited, so the snapshot is retained and the work is reported
 * as uncited instead.
 */
import { fingerprint, } from "@refmgr/doc-schema";
import { CitationClusterRenderer, resolveDocumentOrder, } from "./clusters.js";
export class DocumentEditError extends Error {
    code;
    /** Present when the fault is attributable to one cluster. Never carries reference text. */
    clusterId;
    constructor(code, message, clusterId) {
        super(message);
        this.name = "DocumentEditError";
        this.code = code;
        if (clusterId !== undefined)
            this.clusterId = clusterId;
    }
}
/**
 * Plans citation insertions and deletions against a portable document part.
 *
 * Stateless: each plan re-renders the document from scratch. That is the honest cost of
 * correctness at this stage — citeproc's own incremental path needs the processor's registry
 * kept alive between edits, which is a §21 performance question E10-07 owns, and a wrong
 * incremental answer is indistinguishable from a right one until a reader spots the number.
 */
export class CitationDocumentEditor {
    #renderer;
    #hash;
    #styleRef;
    constructor(options) {
        this.#renderer = new CitationClusterRenderer(options);
        this.#hash = options.hashRenderedText ?? ((text) => fingerprint(text));
        this.#styleRef = options.styleRef;
    }
    /**
     * Re-render a document exactly as the host currently orders it and plan only the text that
     * differs from the payload's last processor output.
     *
     * This deliberately uses the same full-document path as insert/remove. A refresh is the
     * correctness baseline against which any future incremental processor cache must be compared;
     * keeping a second implementation here would make a fast wrong citation look trustworthy.
     */
    refresh(part, options = {}) {
        const ordered = resolveDocumentOrder(part, options.documentOrder);
        return this.#plan(part, ordered, ordered, part.snapshots, [], -1);
    }
    /** Plans the insertion of one citation at a position in the document. */
    insert(part, insertion, options = {}) {
        const before = resolveDocumentOrder(part, options.documentOrder);
        if (!Number.isInteger(insertion.index) || insertion.index < 0 || insertion.index > before.length) {
            throw new DocumentEditError("position-out-of-range", `Cannot insert at position ${insertion.index}; the document holds ${before.length} citation(s)`, insertion.cluster.clusterId);
        }
        if (before.some((cluster) => cluster.clusterId === insertion.cluster.clusterId)) {
            throw new DocumentEditError("duplicate-cluster-id", `Cluster id "${insertion.cluster.clusterId}" is already in the document; cluster ids are never reused`, insertion.cluster.clusterId);
        }
        if (insertion.cluster.items.length === 0) {
            throw new DocumentEditError("empty-insertion", `Cluster "${insertion.cluster.clusterId}" cites no items`, insertion.cluster.clusterId);
        }
        const snapshots = mergeSnapshots(part, insertion);
        const after = [...before];
        after.splice(insertion.index, 0, insertion.cluster);
        return this.#plan(part, before, after, snapshots, [], insertion.index);
    }
    /** Plans the removal of one or more citations, as one document edit. */
    remove(part, deletion, options = {}) {
        const before = resolveDocumentOrder(part, options.documentOrder);
        if (deletion.clusterIds.length === 0) {
            throw new DocumentEditError("invalid-deletion", "A deletion must name at least one cluster");
        }
        const removing = new Set();
        for (const clusterId of deletion.clusterIds) {
            if (removing.has(clusterId)) {
                throw new DocumentEditError("invalid-deletion", `Cluster "${clusterId}" is named twice in the same deletion`, clusterId);
            }
            if (!before.some((cluster) => cluster.clusterId === clusterId)) {
                throw new DocumentEditError("unknown-cluster", `Cluster "${clusterId}" is not in the document; the host's view and the document part disagree (§9.5)`, clusterId);
            }
            removing.add(clusterId);
        }
        const after = before.filter((cluster) => !removing.has(cluster.clusterId));
        // The earliest removed position is where renumbering starts; note numbers below it stand.
        const firstRemoved = before.findIndex((cluster) => removing.has(cluster.clusterId));
        return this.#plan(part, before, after, part.snapshots, [...removing], firstRemoved);
    }
    #plan(part, before, after, snapshots, removed, editIndex) {
        const previousById = new Map(before.map((cluster) => [cluster.clusterId, cluster]));
        // `order` is derived positional data, so the plan's part states the positions it rendered.
        const ordered = after.map((cluster, index) => ({ ...cluster, order: index }));
        const rendered = this.#renderer.render({ ...part, clusters: ordered, snapshots });
        const updates = [];
        const blockedByManualEdit = [];
        const unchanged = [];
        const warnings = [];
        const written = new Map();
        rendered.clusters.forEach((cluster, position) => {
            const previous = previousById.get(cluster.clusterId);
            const previousText = previous?.renderedText;
            const reason = previous === undefined
                ? "inserted"
                : previousText === undefined
                    ? "no-recorded-text"
                    : previousText === cluster.text
                        ? undefined
                        : "text-changed";
            if (reason === undefined) {
                unchanged.push(cluster.clusterId);
                return;
            }
            const update = {
                clusterId: cluster.clusterId,
                text: cluster.text,
                ...(previousText === undefined ? {} : { previousText }),
                reason,
                position,
            };
            // A manual edit is the user's own writing. It blocks the write even when the citation is
            // now demonstrably wrong: what to do about that is the §9.4 conversation, not our call.
            if (previous !== undefined && previous.manualEdit.state !== "none") {
                blockedByManualEdit.push(update);
                warnings.push({
                    code: "manual-edit-blocks-update",
                    clusterId: cluster.clusterId,
                    message: `This citation needs new text but carries a manual edit the user has ${previous.manualEdit.state === "kept" ? "chosen to keep" : "not yet resolved"}; it was not overwritten (§9.4)`,
                });
                return;
            }
            updates.push(update);
            written.set(cluster.clusterId, cluster.text);
        });
        const cited = new Set();
        for (const cluster of after)
            for (const item of cluster.items)
                cited.add(item.itemId);
        const uncitedItemIds = Object.keys(snapshots)
            .filter((itemId) => !cited.has(itemId))
            .sort();
        if (removed.length > 0 && uncitedItemIds.length > 0) {
            warnings.push({
                code: "snapshot-now-uncited",
                message: `${uncitedItemIds.length} item snapshot(s) are no longer cited; they are retained because §9.4 allows a bibliography entry to be pinned or uncited`,
            });
        }
        // Note numbers come from the host's document, never from us: Word renumbers footnotes on
        // its own, and a note number we invented would decide `ibid` and subsequent-note position
        // from a sequence the document does not actually have.
        if (rendered.noteStyle && editIndex >= 0 && editIndex < after.length) {
            warnings.push({
                code: "note-numbers-stale",
                message: `The style renders notes and ${after.length - editIndex} citation(s) follow the edit; re-report their note numbers from the document and render again before trusting ibid or subsequent-note forms`,
            });
        }
        const nextClusters = ordered.map((cluster) => {
            const text = written.get(cluster.clusterId);
            if (text === undefined)
                return cluster;
            return {
                ...cluster,
                renderedText: text,
                renderedHash: this.#hash(text),
                ...(this.#styleRef === undefined ? {} : { renderedWith: this.#styleRef }),
            };
        });
        return {
            part: { ...part, clusters: nextClusters, snapshots },
            rendered,
            updates,
            unchanged,
            removed,
            blockedByManualEdit,
            uncitedItemIds,
            bibliographyChanged: rendered.bibliographyChanged,
            warnings,
        };
    }
}
/**
 * Merges the snapshots an insertion supplies into the document's own.
 *
 * A supplied snapshot that contradicts the stored one is refused rather than preferred. The
 * snapshot is shared by every citation of that work, so accepting a fresher copy while
 * inserting one citation would silently re-render citations elsewhere in the manuscript;
 * updating metadata is §14's side-by-side review, not a side effect of adding a citation
 * (invariant 3).
 */
function mergeSnapshots(part, insertion) {
    const merged = { ...part.snapshots };
    for (const snapshot of insertion.snapshots ?? []) {
        if (snapshot.csl.id !== snapshot.itemId) {
            throw new DocumentEditError("malformed-snapshot", `Snapshot for item "${snapshot.itemId}" carries CSL id "${snapshot.csl.id}"`, insertion.cluster.clusterId);
        }
        const existing = merged[snapshot.itemId];
        if (existing !== undefined && fingerprint(existing) !== fingerprint(snapshot)) {
            throw new DocumentEditError("conflicting-snapshot", `The document already holds a different snapshot for item "${snapshot.itemId}"; changing it would re-render citations elsewhere in the document (§14)`, insertion.cluster.clusterId);
        }
        merged[snapshot.itemId] = snapshot;
    }
    for (const item of insertion.cluster.items) {
        if (merged[item.itemId] === undefined) {
            throw new DocumentEditError("missing-snapshot", `Item "${item.itemId}" has no document-local snapshot; supply one so the document stays editable without the library (§9.3)`, insertion.cluster.clusterId);
        }
    }
    return merged;
}
