/**
 * Bibliography rendering for the plain-text export (SPEC §5.3, §8.1; task E04-06.2).
 *
 * Invariant 4 says citation output comes only from the CSL processor. This module is where that
 * rule is made operational for a bibliography: it projects canonical items to CSL-JSON, hands
 * them to a `CslProcessor`, and returns the processor's strings **and the processor's ordering**
 * without touching either. Sorting a bibliography is a style rule (`<sort>` in the CSL
 * `<bibliography>` element), so re-ordering entries here would be hand-rolled formatting just as
 * surely as rewriting punctuation would be.
 *
 * The processor is injected rather than constructed, so the boundary stays processor-neutral in
 * the sense ADR-0056 set up; the citeproc-js adapter is only the default.
 */
import { toCslItem } from "@refmgr/core";
import { CiteprocJsAdapter } from "./adapters/citeproc-js/index.js";
export class BibliographyRenderError extends Error {
    code;
    constructor(code, message) {
        super(message);
        this.name = "BibliographyRenderError";
        this.code = code;
    }
}
function entryIds(metadata) {
    const raw = metadata.entry_ids;
    if (!Array.isArray(raw)) {
        throw new BibliographyRenderError("entry-id-mismatch", "The CSL processor did not report which items its bibliography entries came from");
    }
    return raw.map((group) => {
        const ids = Array.isArray(group) ? group : [group];
        if (!ids.every((id) => typeof id === "string")) {
            throw new BibliographyRenderError("entry-id-mismatch", "The CSL processor reported a bibliography entry id that is not a string");
        }
        return ids;
    });
}
/**
 * Renders a bibliography by asking a CSL processor for one.
 *
 * The only transformation applied to processor output is stripping trailing line breaks, which
 * citeproc's text output appends as an entry terminator. Nothing else is added, removed,
 * re-punctuated or re-ordered.
 */
export class CslBibliographyRenderer {
    #styleXml;
    #localeXml;
    #locale;
    #createProcessor;
    constructor(options) {
        this.#styleXml = options.styleXml;
        this.#localeXml = options.localeXml;
        this.#locale = options.locale;
        this.#createProcessor =
            options.createProcessor ?? ((processorOptions) => new CiteprocJsAdapter(processorOptions));
    }
    render(items) {
        const projected = [];
        const seen = new Set();
        for (const item of items) {
            if (seen.has(item.id)) {
                throw new BibliographyRenderError("duplicate-item-id", `Cannot render a bibliography: item id "${item.id}" was supplied more than once`);
            }
            seen.add(item.id);
            projected.push(toCslItem(item));
        }
        return this.#renderProjected(projected);
    }
    /**
     * The same render from CSL-JSON that has already been projected — a document's §9.3 snapshots,
     * which are exactly `toCslItem` of the item at the moment it was cited (ADR-0242). Converting a
     * snapshot back into a `ReferenceItem` to call {@link render} would run the projection a second
     * time over data it was not written for; a snapshot is the processor input already.
     */
    renderCsl(items) {
        const seen = new Set();
        for (const item of items) {
            if (seen.has(item.id)) {
                throw new BibliographyRenderError("duplicate-item-id", `Cannot render a bibliography: item id "${item.id}" was supplied more than once`);
            }
            seen.add(item.id);
        }
        return this.#renderProjected([...items]);
    }
    #renderProjected(projected) {
        const processorOptions = {
            styleXml: this.#styleXml,
            localeXml: this.#localeXml,
            items: projected,
            outputFormat: "text",
            ...(this.#locale === undefined ? {} : { locale: this.#locale }),
        };
        const processor = this.#createProcessor(processorOptions);
        const bibliography = processor.makeBibliography(projected.map((item) => item.id));
        if (projected.length > 0 && bibliography.entries.length === 0) {
            throw new BibliographyRenderError("bibliography-unsupported", "The CSL style produced no bibliography for the supplied items");
        }
        const ids = entryIds(bibliography.metadata);
        if (ids.length !== bibliography.entries.length) {
            throw new BibliographyRenderError("entry-id-mismatch", "The CSL processor returned a different number of bibliography entries and entry ids");
        }
        return {
            processor: processor.processor,
            processorVersion: processor.processorVersion,
            entries: bibliography.entries.map((text, index) => ({
                itemIds: [...ids[index]],
                text: text.replace(/[\r\n]+$/u, ""),
            })),
            bibliographyErrors: bibliography.bibliographyErrors,
        };
    }
}
