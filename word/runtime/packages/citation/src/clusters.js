/**
 * The citation cluster model — SPEC §8.2, task E08-03.
 *
 * A cluster is one in-document citation object: the Word content control, the Google Docs
 * bookmark, the `[@key, p. 34]` span. `packages/doc-schema` owns what a cluster *is* on disk
 * (ADR-0015); this module owns what a cluster *means to the citation processor* — how its
 * items, locators, affixes and author options are projected into processor input, and how a
 * whole document's worth of clusters is rendered in order.
 *
 * Three rules shape everything here.
 *
 * **Invariant 4 — no formatting outside the processor.** Nothing in this file punctuates,
 * abbreviates, collapses, numbers or sorts anything. Citation numbers, `1–3` range collapsing,
 * `ibid.`, year-suffix disambiguation and within-cluster ordering are all decided by the style
 * and produced by citeproc; we supply well-formed input and return its strings verbatim.
 * In particular a cluster's items are handed to the processor **in the author's stated order**
 * and never re-sorted: `<sort>` inside the CSL `<citation>` element is where citation sorting
 * lives, and sorting here would be hand-rolled formatting exactly as surely as re-punctuating
 * would be.
 *
 * **Ordering is the thing that silently goes wrong.** Every position-sensitive feature in
 * §8.2 — numeric assignment, range collapsing, `ibid`, subsequent-note short forms,
 * same-author-same-year disambiguation — is a function of where a cluster sits in the
 * document, and every one of them fails *plausibly* rather than loudly when the order is
 * wrong: the citation still renders, it just points at the wrong work. So document order is
 * treated as load-bearing input rather than a detail, and a contradiction about it is refused
 * rather than resolved (invariant 3).
 *
 * **Refuse rather than emit a placeholder.** Several malformed clusters do not make citeproc
 * throw; they make it emit `[NO_PRINTED_FORM]` or quietly drop what the user asked for. Those
 * cases are checked here, with the observed citeproc behaviour recorded next to each check,
 * because a citation that renders wrongly is worse than one that refuses to render.
 */
import { CSL_LOCATOR_LABELS, DOC_SCHEMA_VERSION, } from "@refmgr/doc-schema";
import { CiteprocJsAdapter } from "./adapters/citeproc-js/index.js";
import { validateCslStyle } from "./styles.js";
const LOCATOR_LABELS = new Set(CSL_LOCATOR_LABELS);
export class ClusterModelError extends Error {
    code;
    /** Present when the fault is attributable to one cluster. Never carries reference text. */
    clusterId;
    constructor(code, message, clusterId) {
        super(message);
        this.name = "ClusterModelError";
        this.code = code;
        if (clusterId !== undefined)
            this.clusterId = clusterId;
    }
}
/**
 * Projects one document cluster item into processor input.
 *
 * Exported because the same projection is needed by citation preview (§9.2) and by numeric
 * renumbering (E08-07), and because it is the single place where the doc-schema vocabulary
 * (`locator: {label, value}`, `suppressAuthor`) meets citeproc's (`locator`/`label`,
 * `suppress-author`). Affixes cross unchanged: a prefix is the user's own writing and
 * trimming or re-punctuating it would be both formatting and data loss.
 */
export function toProcessorItem(item, clusterId) {
    if (item.suppressAuthor === true && item.authorOnly === true) {
        throw new ClusterModelError("conflicting-author-options", `Cluster item "${item.itemId}" sets both suppressAuthor and authorOnly; citeproc renders that as [NO_PRINTED_FORM]`, clusterId);
    }
    const projected = { id: item.itemId };
    if (item.locator !== undefined) {
        if (item.locator.value.trim() === "") {
            throw new ClusterModelError("empty-locator", `Cluster item "${item.itemId}" has a "${item.locator.label}" locator with no value; citeproc would drop the locator without reporting it`, clusterId);
        }
        // Verbatim: "12–15" keeps its en dash. Range collapsing and page-range formatting are
        // style rules (`page-range-format`), and E08-07 owns citation-range collapsing.
        projected.locator = item.locator.value;
        projected.label = item.locator.label;
    }
    if (item.prefix !== undefined)
        projected.prefix = item.prefix;
    if (item.suffix !== undefined)
        projected.suffix = item.suffix;
    if (item.suppressAuthor === true)
        projected["suppress-author"] = true;
    if (item.authorOnly === true)
        projected["author-only"] = true;
    return projected;
}
/**
 * Projects one document cluster into a processor cluster at a given note index.
 *
 * Item order is preserved exactly. A multi-item cluster may legitimately name the same work
 * twice — two page ranges of one book in one parenthesis — so repeats are passed through
 * rather than deduplicated.
 */
export function toProcessorCluster(cluster, noteIndex) {
    if (cluster.items.length === 0) {
        throw new ClusterModelError("empty-cluster", `Cluster "${cluster.clusterId}" cites no items; citeproc renders that as [NO_PRINTED_FORM]`, cluster.clusterId);
    }
    return {
        citationID: cluster.clusterId,
        citationItems: cluster.items.map((item) => toProcessorItem(item, cluster.clusterId)),
        properties: { noteIndex },
    };
}
/**
 * Renders every citation in a document by walking its clusters in order through one CSL
 * processor.
 *
 * The walk is not an implementation detail. citeproc re-renders *earlier* clusters when a
 * later one changes their meaning — inserting a citation renumbers the ones after it, and a
 * second Smith 2020 turns the first into "Smith 2020a" — so updates are accumulated across
 * the whole document and the last text the processor gives for a cluster is the one that
 * counts. Rendering clusters independently would produce a document that is individually
 * plausible and collectively wrong.
 */
export class CitationClusterRenderer {
    #styleXml;
    #localeXml;
    #locale;
    #outputFormat;
    #createProcessor;
    constructor(options) {
        this.#styleXml = options.styleXml;
        this.#localeXml = options.localeXml;
        this.#locale = options.locale;
        this.#outputFormat = options.outputFormat;
        this.#createProcessor =
            options.createProcessor ?? ((processorOptions) => new CiteprocJsAdapter(processorOptions));
    }
    render(part, options = {}) {
        if (part.schemaVersion > DOC_SCHEMA_VERSION) {
            throw new ClusterModelError("unsupported-schema-version", `Document schema version ${part.schemaVersion} is newer than supported version ${DOC_SCHEMA_VERSION}; treat the document as read-only`);
        }
        const warnings = [...unknownLocatorLabels(part)];
        const ordered = resolveDocumentOrder(part, options.documentOrder, warnings);
        // The style, not the document's stored preference, decides whether citations are notes:
        // the preference is a declaration that a style change (§9.1) leaves behind, and citeproc
        // will render notes from a note style regardless of what the payload says.
        const styleValidation = validateCslStyle(this.#styleXml);
        const noteStyle = styleValidation.metadata?.className === "note";
        const declared = part.preferences.citationFormat;
        if (declared !== undefined && declared !== (noteStyle ? "note" : "in-text")) {
            warnings.push({
                code: "citation-format-mismatch",
                message: `The document declares citationFormat "${declared}" but the style renders ${noteStyle ? "notes" : "in-text citations"}; the style was used`,
            });
        }
        const noteIndexes = resolveNoteIndexes(ordered, noteStyle, warnings);
        const items = [];
        for (const snapshot of Object.values(part.snapshots)) {
            items.push(snapshot.csl);
        }
        const known = new Set(Object.keys(part.snapshots));
        for (const cluster of ordered) {
            for (const item of cluster.items) {
                if (!known.has(item.itemId)) {
                    throw new ClusterModelError("missing-snapshot", `Cluster "${cluster.clusterId}" cites item "${item.itemId}", which has no document-local snapshot; the document would not be editable without the library (§9.3)`, cluster.clusterId);
                }
            }
            if (cluster.manualEdit.state === "kept" || cluster.manualEdit.state === "pending") {
                warnings.push({
                    code: cluster.manualEdit.state === "kept" ? "manual-edit-kept" : "manual-edit-pending",
                    clusterId: cluster.clusterId,
                    message: cluster.manualEdit.state === "kept"
                        ? "This citation carries a manual edit the user chose to keep; rendered text must not replace it without the §9.4 warning"
                        : "This citation carries an undecided manual edit; the user must resolve it before it is replaced",
                });
            }
        }
        const processorOptions = {
            styleXml: this.#styleXml,
            localeXml: this.#localeXml,
            items,
            ...(this.#locale === undefined ? {} : { locale: this.#locale }),
            ...(this.#outputFormat === undefined ? {} : { outputFormat: this.#outputFormat }),
        };
        const processor = this.#createProcessor(processorOptions);
        const text = new Map();
        const citationErrors = [];
        const processed = [];
        let bibliographyChanged = false;
        ordered.forEach((cluster, index) => {
            const noteIndex = noteIndexes[index];
            const result = processor.processCitationCluster(toProcessorCluster(cluster, noteIndex), [...processed], []);
            // Later clusters re-render earlier ones; the newest text for an id always wins.
            for (const update of result.citationUpdates)
                text.set(update.citationID, update.text);
            citationErrors.push(...result.citationErrors);
            if (result.bibliographyChanged)
                bibliographyChanged = true;
            processed.push({ citationID: cluster.clusterId, noteIndex });
        });
        const clusters = ordered.map((cluster, index) => {
            const rendered = text.get(cluster.clusterId);
            if (rendered === undefined) {
                throw new ClusterModelError("unrendered-cluster", `The CSL processor returned no text for cluster "${cluster.clusterId}"`, cluster.clusterId);
            }
            return {
                clusterId: cluster.clusterId,
                text: rendered,
                ...(noteStyle ? { noteNumber: noteIndexes[index] } : {}),
                itemIds: cluster.items.map((item) => item.itemId),
                manualEdit: cluster.manualEdit.state,
            };
        });
        return {
            processor: processor.processor,
            processorVersion: processor.processorVersion,
            noteStyle,
            clusters,
            warnings,
            citationErrors,
            bibliographyChanged,
        };
    }
}
/**
 * Resolves the order the clusters are rendered in.
 *
 * With an explicit walk, the host has looked at the document and its answer is authoritative;
 * stored `order` values are last-known and are expected to be stale, so a disagreement is
 * reported and the walk is used. Without one, the payload's array is the only claim of order
 * that exists — so a stored `order` that contradicts it is a *second, conflicting* claim with
 * nothing to arbitrate between them, and it is refused rather than silently picked
 * (invariant 3).
 *
 * Exported because an edit (E08-07) must resolve the same order before it can say what
 * "insert at position 3" means, and two implementations of this rule would drift apart
 * exactly where it matters least visibly.
 */
export function resolveDocumentOrder(part, documentOrder, warnings = []) {
    const byId = new Map();
    for (const cluster of part.clusters) {
        if (byId.has(cluster.clusterId)) {
            throw new ClusterModelError("duplicate-cluster-id", `Cluster id "${cluster.clusterId}" appears more than once; processor output could not be attributed to a position`, cluster.clusterId);
        }
        byId.set(cluster.clusterId, cluster);
    }
    if (documentOrder === undefined) {
        part.clusters.forEach((cluster, index) => {
            if (cluster.order !== undefined && cluster.order !== index) {
                throw new ClusterModelError("inconsistent-order", `Cluster "${cluster.clusterId}" records order ${cluster.order} but is at position ${index}; supply documentOrder from the document itself rather than letting citation numbering be guessed`, cluster.clusterId);
            }
        });
        return part.clusters;
    }
    if (documentOrder.length !== byId.size) {
        throw new ClusterModelError("order-mismatch", `The supplied document order lists ${documentOrder.length} clusters but the document part holds ${byId.size}`);
    }
    const seen = new Set();
    const ordered = [];
    documentOrder.forEach((clusterId, index) => {
        const cluster = byId.get(clusterId);
        if (cluster === undefined) {
            throw new ClusterModelError("order-mismatch", `The supplied document order names cluster "${clusterId}", which is not in the document part`, clusterId);
        }
        if (seen.has(clusterId)) {
            throw new ClusterModelError("order-mismatch", `The supplied document order names cluster "${clusterId}" more than once`, clusterId);
        }
        seen.add(clusterId);
        if (cluster.order !== undefined && cluster.order !== index) {
            warnings.push({
                code: "stale-order",
                clusterId,
                message: `Stored order ${cluster.order} disagrees with the document walk position ${index}; the walk was used`,
            });
        }
        ordered.push(cluster);
    });
    return ordered;
}
/**
 * Works out the note index handed to the processor for each cluster, in document order.
 *
 * For a note style these are the footnote/endnote numbers the host reports, and they decide
 * `ibid`, `ibid-with-locator` and subsequent-note short forms. A backwards sequence makes
 * citeproc emit a console warning and then render anyway — meaning a stale note number turns
 * into a wrong "Ibid." that nobody is told about — so it is refused here instead. Clusters may
 * share a number, because one footnote can hold several citations.
 */
function resolveNoteIndexes(clusters, noteStyle, warnings) {
    if (!noteStyle) {
        const carrying = clusters.filter((cluster) => cluster.noteNumber !== undefined);
        if (carrying.length > 0) {
            warnings.push({
                code: "note-number-ignored",
                message: `${carrying.length} cluster(s) carry note numbers but the style renders in-text citations; the numbers were preserved and not used`,
            });
        }
        return clusters.map(() => 0);
    }
    const indexes = [];
    let previous = 0;
    clusters.forEach((cluster) => {
        let noteNumber = cluster.noteNumber;
        if (noteNumber === undefined) {
            // The document position is the only evidence available. Say so rather than imply the
            // host reported a footnote number it never reported.
            noteNumber = previous + 1;
            warnings.push({
                code: "synthesised-note-number",
                clusterId: cluster.clusterId,
                message: `No note number was supplied; note ${noteNumber} was derived from document position`,
            });
        }
        else if (noteNumber < previous) {
            throw new ClusterModelError("non-monotonic-note-number", `Cluster "${cluster.clusterId}" has note number ${noteNumber} after note number ${previous}; ibid and subsequent-note position would be decided from a stale sequence`, cluster.clusterId);
        }
        indexes.push(noteNumber);
        previous = noteNumber;
    });
    return indexes;
}
/**
 * Reports locator labels a document uses that CSL 1.0.2 does not define.
 *
 * Separate from rendering because it is a diagnostic, not a failure: citeproc accepts an
 * unknown label and prints the locator value **without any label**, so the citation looks
 * fine and quietly says less than the author meant. `packages/doc-schema` warns about the
 * same thing structurally; this reports it for a rendering pass.
 */
export function unknownLocatorLabels(part) {
    const warnings = [];
    for (const cluster of part.clusters) {
        for (const item of cluster.items) {
            if (item.locator !== undefined && !LOCATOR_LABELS.has(item.locator.label)) {
                warnings.push({
                    code: "unknown-locator-label",
                    clusterId: cluster.clusterId,
                    message: `Locator label "${item.locator.label}" is not a CSL 1.0.2 label; citeproc will print the locator value with no label`,
                });
            }
        }
    }
    return warnings;
}
