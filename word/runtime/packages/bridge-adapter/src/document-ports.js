/**
 * A Word document's own references, as the bridge's library and citation ports (SPEC §9.3;
 * ADR-0242 decision 2).
 *
 * §9.3 requires every citation to carry "a document-local snapshot of essential reference metadata"
 * so that it stays editable when the desktop is closed, the user is offline, or the document is on
 * another computer. Word on iPad and on the web never has the desktop. This composition makes the
 * snapshots the library: the Word session handler (`word-session.ts`) runs over these ports exactly
 * as it runs over the user's library, with the same refusals, the same CSL processor and the same
 * planners. Nothing here renders or plans a citation itself.
 *
 * What a document can supply, and what it cannot:
 *
 * - **Every reference the document cites** (its snapshots), plus any the writer has added in the
 *   pane. Both are CSL-JSON already — a snapshot is `toCslItem` of the item when it was cited — so
 *   the citation port renders them as they stand, with no second projection.
 * - **No collections.** A document has none; a search scoped to one finds nothing.
 * - **Identifiers only through an importer the pane supplies** (ADR-0242 decision 5). Without one,
 *   each identifier comes back unresolved with a reason, never silently dropped.
 * - **No collaborator review.** Comparing a document with a library needs a library.
 */
import { fromCslItem } from "@refmgr/core";
import { BridgeRequestError, } from "@refmgr/bridge";
import { readDocumentPart } from "@refmgr/doc-schema";
import { payloadText } from "../../word-addin/src/editing.js";
import { CitationBridgeAdapter } from "./citation.js";
import { styleSummaries, toSearchHit } from "./search-hit.js";
/**
 * The references a document's citation data holds, read from Word's evidence (`customXmlParts`).
 *
 * Deliberately tolerant: a part that cannot be read contributes nothing rather than failing the
 * listing. Listing is not the place to refuse — every operation that *writes* re-reads the document
 * through the session handler, which refuses an unreadable or ambiguous payload by name.
 */
export function documentReferences(evidence) {
    if (typeof evidence !== "object" || evidence === null)
        return [];
    const parts = evidence.customXmlParts;
    if (!Array.isArray(parts))
        return [];
    const found = [];
    for (const candidate of parts) {
        const xml = typeof candidate === "object" && candidate !== null ? candidate.xml : undefined;
        if (typeof xml !== "string")
            continue;
        const json = payloadText(xml);
        const read = json === null ? null : readDocumentPart(json);
        if (read?.status !== "ok")
            continue;
        for (const snapshot of Object.values(read.part.snapshots)) {
            found.push({ itemId: snapshot.itemId, csl: snapshot.csl, addedAt: snapshot.snapshotAt });
        }
    }
    return mergeDocumentReferences(found);
}
/**
 * The styles and locales a document's citation data names — what must be installed before the
 * session can render it. Read from every readable part, deduplicated, in the order found.
 */
export function documentRequirements(evidence) {
    const styleIds = [];
    const locales = [];
    const parts = typeof evidence === "object" && evidence !== null ? evidence.customXmlParts : undefined;
    if (!Array.isArray(parts))
        return { styleIds, locales };
    for (const candidate of parts) {
        const xml = typeof candidate === "object" && candidate !== null ? candidate.xml : undefined;
        if (typeof xml !== "string")
            continue;
        const json = payloadText(xml);
        const read = json === null ? null : readDocumentPart(json);
        if (read?.status !== "ok")
            continue;
        if (!styleIds.includes(read.part.style.id))
            styleIds.push(read.part.style.id);
        if (!locales.includes(read.part.locale))
            locales.push(read.part.locale);
    }
    return { styleIds, locales };
}
/**
 * One list from several, first occurrence of an id winning. A document's own snapshot is listed
 * before anything added in the pane, so a reference the document already cites is never shadowed
 * by a different copy of itself.
 */
export function mergeDocumentReferences(...lists) {
    const seen = new Set();
    const merged = [];
    for (const list of lists) {
        for (const reference of list) {
            if (seen.has(reference.itemId))
                continue;
            seen.add(reference.itemId);
            merged.push(reference);
        }
    }
    return merged;
}
/**
 * A picker row for a reference held only as CSL-JSON.
 *
 * Built through the canonical projection so the row reads exactly as the library's would. The
 * projection validates, and a snapshot written by another tool may not satisfy it; that reference is
 * still citable (the processor renders the CSL as it stands), so its row falls back to the fields a
 * picker can show without a projection rather than disappearing from the list.
 */
function hitFor(reference) {
    try {
        return toSearchHit(fromCslItem(reference.csl, { libraryId: "document", createdAt: reference.addedAt }));
    }
    catch {
        const csl = reference.csl;
        return {
            itemId: reference.itemId,
            itemType: typeof csl["type"] === "string" ? csl["type"] : "document",
            title: typeof csl["title"] === "string" ? csl["title"] : "",
            creatorSummary: "",
        };
    }
}
/** Case- and accent-insensitive, so "Muller" finds "Müller" as the desktop's search does. */
function folded(text) {
    return text.normalize("NFKD").replace(/\p{M}+/gu, "").toLocaleLowerCase("en");
}
function searchText(reference, hit) {
    const pmid = reference.csl["PMID"];
    return folded([hit.title, hit.creatorSummary, hit.year, hit.containerTitle, hit.doi, typeof pmid === "string" ? pmid : undefined]
        .filter((part) => typeof part === "string" && part.length > 0)
        .join(" "));
}
class DocumentLibraryPort {
    #options;
    constructor(options) {
        this.#options = options;
    }
    #find(itemId) {
        return this.#options.references().find((reference) => reference.itemId === itemId);
    }
    async search(query) {
        // A document has no collections, so a search scoped to one has nothing in it.
        if (query.collectionId !== undefined)
            return [];
        const terms = folded(query.text).split(/\s+/u).filter((term) => term.length > 0);
        const matches = this.#options
            .references()
            .map((reference) => ({ reference, hit: hitFor(reference) }))
            .filter(({ reference, hit }) => {
            const text = searchText(reference, hit);
            return terms.every((term) => text.includes(term));
        });
        // §9.2: references the document already cites first, each group in the order it was listed.
        const preferred = new Set(query.preferItemIds);
        const ranked = [
            ...matches.filter(({ reference }) => preferred.has(reference.itemId)),
            ...matches.filter(({ reference }) => !preferred.has(reference.itemId)),
        ];
        return ranked.slice(0, query.limit).map(({ hit }) => hit);
    }
    async listRecent(query) {
        return [...this.#options.references()]
            .sort((left, right) => left.addedAt < right.addedAt ? 1 : left.addedAt > right.addedAt ? -1 : left.itemId.localeCompare(right.itemId))
            .slice(0, query.limit)
            .map(hitFor);
    }
    async listCollections() {
        return [];
    }
    async getItem(itemId) {
        const reference = this.#find(itemId);
        return reference === undefined ? null : { ...hitFor(reference), csl: reference.csl };
    }
    async listStyles() {
        return styleSummaries(this.#options.styles);
    }
    async importIdentifiers(request) {
        if (this.#options.importIdentifiers !== undefined)
            return this.#options.importIdentifiers(request);
        return {
            imported: [],
            unresolved: request.identifiers.map((identifier) => ({
                identifier,
                code: "unavailable",
                reason: "this pane has no identifier lookup configured",
                retryable: false,
            })),
        };
    }
    async reviewCollaboratorReferences() {
        // A caller's request this composition cannot answer, stated as such — not the listener's
        // unexplained 500 an unconfigured library adapter produces.
        throw new BridgeRequestError(400, "unavailable", "Comparing this document with a library needs Reference Manager on this computer.");
    }
}
/** The library and citation ports over a document's references (ADR-0242). */
export function createDocumentBridgePorts(options) {
    return {
        library: new DocumentLibraryPort(options),
        citation: new CitationBridgeAdapter({
            cslItems: (itemId) => options.references().find((reference) => reference.itemId === itemId)?.csl,
            styles: options.styles,
            createProcessor: options.createProcessor,
        }),
    };
}
