import { cloneJson, jsonEquals, processorCslType } from "@refmgr/core";
import citeproc from "citeproc";
export class CitationConfigurationError extends Error {
    constructor(message) {
        super(message);
        this.name = "CitationConfigurationError";
    }
}
export class CitationInputConflictError extends Error {
    itemId;
    constructor(itemId) {
        super(`Citation input contains conflicting values for item "${itemId}"`);
        this.name = "CitationInputConflictError";
        this.itemId = itemId;
    }
}
export class UnknownCitationItemError extends Error {
    itemId;
    constructor(itemId) {
        super(`Citation item "${itemId}" was not supplied to the processor`);
        this.name = "UnknownCitationItemError";
        this.itemId = itemId;
    }
}
/**
 * A caller supplied a position test variable on the document path, where it cannot be honoured.
 *
 * Refused rather than dropped (E08-10). `processCitationCluster` derives position from document
 * state and citeproc overwrites whatever it was handed, so silently accepting the field would
 * hand back a full first citation to a caller who had explicitly asked for `Ibid.` — a wrong
 * citation that looks entirely plausible, which is the failure mode §8.3 exists to prevent.
 */
export class CitationPositionConflictError extends Error {
    itemId;
    field;
    constructor(itemId, field) {
        super(`Citation item "${itemId}" sets "${field}", which processCitationCluster derives from ` +
            `document state and would overwrite; supply it only to makeCitationCluster`);
        this.name = "CitationPositionConflictError";
        this.itemId = itemId;
        this.field = field;
    }
}
function cloneItem(item) {
    return cloneJson(item);
}
function withProcessorType(item) {
    const type = processorCslType(item.type);
    return type === item.type ? cloneItem(item) : { ...cloneItem(item), type };
}
/**
 * CSL's position vocabulary in citeproc-js's integer encoding.
 *
 * The numbers are read off `CSL.POSITION_*` and are part of citeproc-js's input contract for
 * `makeCitationCluster`, which copies every key of a raw citation item straight through. They are
 * confined to this adapter: nothing outside it should know that `ibid` is `2`.
 */
const CITEPROC_POSITION = {
    first: 0,
    subsequent: 1,
    ibid: 2,
    "ibid-with-locator": 3,
    "container-subsequent": 4,
};
/** The three CSL position test variables (`CSL.POSITION_TEST_VARS`). */
const POSITION_FIELDS = ["position", "near-note", "first-reference-note-number"];
/**
 * Projects a citation item into citeproc's input shape.
 *
 * `honourPosition` says whether this call site can act on the position test variables.
 * `makeCitationCluster` can, because the caller is the only source of position for a cluster
 * rendered with no document around it; `processCitationCluster` cannot, and refuses instead.
 */
function cloneCitationItem(item, honourPosition) {
    const copy = { id: item.id };
    if (item.locator !== undefined)
        copy.locator = item.locator;
    if (item.label !== undefined)
        copy.label = item.label;
    if (item.prefix !== undefined)
        copy.prefix = item.prefix;
    if (item.suffix !== undefined)
        copy.suffix = item.suffix;
    if (item["suppress-author"] !== undefined)
        copy["suppress-author"] = item["suppress-author"];
    if (item["author-only"] !== undefined)
        copy["author-only"] = item["author-only"];
    if (!honourPosition) {
        for (const field of POSITION_FIELDS) {
            if (item[field] !== undefined)
                throw new CitationPositionConflictError(item.id, field);
        }
        return copy;
    }
    if (item.position !== undefined) {
        const encoded = CITEPROC_POSITION[item.position];
        if (encoded === undefined) {
            throw new CitationConfigurationError(`Citation item "${item.id}" has unknown position "${String(item.position)}"`);
        }
        copy.position = encoded;
    }
    if (item["near-note"] !== undefined)
        copy["near-note"] = item["near-note"];
    if (item["first-reference-note-number"] !== undefined) {
        copy["first-reference-note-number"] = item["first-reference-note-number"];
    }
    return copy;
}
function citationPositions(positions) {
    return positions.map((position) => [position.citationID, position.noteIndex]);
}
function errorList(value) {
    if (!Array.isArray(value))
        return [];
    return value.filter((entry) => typeof entry === "string");
}
/** citeproc-js integration. No citeproc source is modified or re-exported. */
export class CiteprocJsAdapter {
    processor = "citeproc-js";
    processorVersion = citeproc.PROCESSOR_VERSION ?? "unknown";
    items = new Map();
    engine;
    constructor(options) {
        this.updateItemMap(options.items);
        try {
            this.engine = new citeproc.Engine({
                retrieveItem: (id) => {
                    const item = this.items.get(id);
                    if (item === undefined)
                        throw new UnknownCitationItemError(id);
                    return item;
                },
                retrieveLocale: () => options.localeXml,
            }, options.styleXml, options.locale ?? "en-US");
            this.engine.setOutputFormat(options.outputFormat ?? "html");
            this.engine.updateItems([...this.items.keys()]);
        }
        catch (error) {
            if (error instanceof UnknownCitationItemError)
                throw error;
            throw new CitationConfigurationError("citeproc-js rejected the supplied style or locale");
        }
    }
    updateItems(items) {
        this.updateItemMap(items);
        this.engine.updateItems(items.map((item) => item.id));
    }
    /**
     * Renders one cluster without registering it in the document.
     *
     * citeproc's own `makeCitationCluster` calls `retrieveItem` for every id and does not check the
     * result: an unknown id makes it fail deep inside the processor with a JSON syntax error that
     * names nothing useful. The ids are therefore checked here first, so an unknown item fails the
     * same way it does on the `processCitationCluster` path.
     *
     * An empty cluster returns citeproc's literal `[NO_PRINTED_FORM]`. That is deliberate at this
     * layer: the adapter is the processor's boundary and reports what the processor said, while
     * the cluster model (`clusters.ts`) is where an empty cluster is refused before it gets here.
     *
     * This is the call that honours an item's `position`, `near-note` and
     * `first-reference-note-number` (E08-10). A cluster rendered in isolation has no preceding
     * cite to compare against, so a style asking `position="ibid"` can only be answered by the
     * caller. citeproc's own `makeCitationCluster` copies every key of a raw item through to the
     * cite, which is what makes the hand-off possible.
     */
    makeCitationCluster(citationItems) {
        for (const item of citationItems) {
            if (!this.items.has(item.id))
                throw new UnknownCitationItemError(item.id);
        }
        return this.engine.makeCitationCluster(citationItems.map((item) => cloneCitationItem(item, true)));
    }
    /**
     * Renders a cluster in its document, letting citeproc derive position from `noteIndex` and the
     * clusters before and after it. That derivation is why `Ibid.` and `subsequent` work on this
     * path without the caller saying anything — and why an item-level position is refused here
     * rather than accepted and overwritten (E08-10).
     */
    processCitationCluster(cluster, citationsPre = [], citationsPost = []) {
        for (const item of cluster.citationItems) {
            if (!this.items.has(item.id))
                throw new UnknownCitationItemError(item.id);
        }
        const properties = {
            citationID: cluster.citationID,
        };
        if (cluster.properties?.noteIndex !== undefined)
            properties.noteIndex = cluster.properties.noteIndex;
        const processorCluster = {
            citationID: cluster.citationID,
            citationItems: cluster.citationItems.map((item) => cloneCitationItem(item, false)),
            properties,
        };
        const response = this.engine.processCitationCluster(processorCluster, citationPositions(citationsPre), citationPositions(citationsPost));
        const knownPositions = [...citationsPre, { citationID: cluster.citationID, noteIndex: cluster.properties?.noteIndex ?? 0 }, ...citationsPost];
        const citationUpdates = response[1].map(([index, text]) => {
            const position = knownPositions[index];
            if (position === undefined) {
                throw new CitationConfigurationError("citeproc-js returned an unknown citation position");
            }
            const update = {
                citationID: position.citationID,
                text,
            };
            if (position.noteIndex !== undefined)
                update.noteIndex = position.noteIndex;
            return update;
        });
        return {
            citationUpdates,
            bibliographyChanged: response[0].bibchange === true,
            citationErrors: errorList(response[0].citation_errors),
        };
    }
    /**
     * `itemIds` selects which items appear; the processor still decides their order, because
     * bibliography order is a `<sort>` rule in the style.
     *
     * citeproc's own parameter is a *bibsection filter object*, not an id list, so the ids are
     * translated into one. Handing citeproc the raw array — which this adapter used to do — is
     * accepted and silently ignored: an array has no `include` key, so no filter is applied and
     * every item comes back. A caller asking for one reference would have been given all of them.
     *
     * **`itemIds` is a request, and its absence is a different request** (E08-11). §9.4 says a
     * bibliography includes what is cited, *removes what is no longer cited*, and supports uncited
     * entries **when requested**. So an explicit id list is registered through citeproc's uncited-item
     * API — that is the caller requesting them, and a bibsection filter can remove an entry but never
     * add one, so without this an explicitly pinned uncited work would be silently dropped. Omitting
     * the list asks the opposite question, "what does this document cite?", and registering every
     * supplied item there answered it wrongly: a work whose last citation had been deleted stayed in
     * the bibliography for as long as the processor still held its data, which is §9.4's "remove
     * records that are no longer cited" failing in the direction a reader cannot detect.
     */
    makeBibliography(itemIds) {
        const ids = itemIds === undefined ? [...this.items.keys()] : [...itemIds];
        for (const id of ids) {
            if (!this.items.has(id))
                throw new UnknownCitationItemError(id);
        }
        if (itemIds !== undefined)
            this.engine.updateUncitedItems(ids);
        const result = this.engine.makeBibliography(itemIds === undefined
            ? undefined
            : { include: ids.map((id) => ({ field: "id", value: id })) });
        if (result === false) {
            return { entries: [], metadata: {}, bibliographyErrors: [] };
        }
        const [metadata, entries] = result;
        return {
            entries: [...entries],
            metadata,
            bibliographyErrors: errorList(metadata.bibliography_errors),
        };
    }
    updateItemMap(items) {
        for (const item of items) {
            const mapped = withProcessorType(item);
            const existing = this.items.get(item.id);
            if (existing !== undefined && !jsonEquals(existing, mapped)) {
                throw new CitationInputConflictError(item.id);
            }
            this.items.set(item.id, mapped);
        }
    }
}
