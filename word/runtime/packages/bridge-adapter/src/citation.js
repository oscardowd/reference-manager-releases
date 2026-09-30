/**
 * The citation half of the bridge's ports, backed by the real CSL processor (SPEC §9.2, §9.3,
 * §8.1; ADR-0006, ADR-0008; task E10-04.1).
 *
 * Invariant 4 — citation output comes only from the CSL processor — is satisfied here by there
 * being no other way for text to reach a response: every string returned came out of citeproc,
 * and this file's own code is a projection (`toProcessorItem`, from `@refmgr/citation`, which is
 * exported precisely so §9.2's preview and the document path share one vocabulary map) plus the
 * decisions below about what to refuse.
 *
 * ## What it refuses, and why refusing is the point
 *
 * **A style substitution.** `CslStyleRegistry.resolveStyle` will fall back to a configured style
 * when the requested one is not installed, and report `usedFallback`. For a preview in Word that
 * is the wrong answer at the wrong moment: the user asked for AMA, the pane would show Chicago,
 * and the only signal is a boolean nobody rendered. So a fallback is converted into a refusal
 * naming the style that is missing. A *dependent* style is not a substitution — it carries no
 * formatting of its own and CSL says its independent parent renders it — so that resolves
 * normally and the response reports both ids.
 *
 * **A language substitution.** A locale falling back `en-GB` → `en-US` changes spelling; falling
 * back `de-DE` → `en-US` changes "u. a." into "et al." and every term in the citation. The first
 * is reported as a warning, the second is refused.
 *
 * **An item the library does not have.** citeproc's `makeCitationCluster` calls `retrieveItem`
 * for every id without checking the result and fails deep inside itself; the adapter in
 * `@refmgr/citation` already guards that, and the ids are checked here first anyway so the
 * caller is told which id, in a 400 rather than a 500.
 *
 * ## What it is honest about
 *
 * Both `previewCitation` and `formatCitation` render the cluster **in isolation**, because the
 * request carries no document: `BridgeCitationRequest` is a style, a locale and some items.
 * For an author-date style that is the real answer. For a numeric style every isolated cluster
 * numbers from one, and for a note style `Ibid.` and `supra` cannot be derived at all — so when
 * the style's own `citation-format` says the document decides, the response carries an
 * `isolated-render` warning rather than a plausible wrong number. Document-aware rendering is
 * E10-06/E10-07, which own `<sort>` order, renumbering and note position.
 */
import { toCslItem } from "@refmgr/core";
import { CiteprocJsAdapter, CslBibliographyRenderer, bibliographyLayout, citationSortsItems, toProcessorItem, } from "@refmgr/citation";
import { BridgeRequestError } from "@refmgr/bridge";
/**
 * The `citation-format` values whose citations are decided by the surrounding document rather
 * than by the cluster alone.
 *
 * Transcribed from CSL 1.0.2's `citation-format` vocabulary (`author-date`, `author`, `numeric`,
 * `label`, `note`). `numeric` numbers in document order and `note` derives `Ibid.`/`supra` from
 * the notes before it; the other three are functions of the cited item and the style.
 */
export const DOCUMENT_DEPENDENT_CITATION_FORMATS = ["numeric", "note"];
function citationFormatOf(style) {
    return style.metadata.categories.find((category) => category.citationFormat !== undefined)
        ?.citationFormat;
}
/** The bridge's own refusal type, so a caller-caused failure is a 400 rather than a silent 500. */
function refuse(code, message) {
    throw new BridgeRequestError(400, code, message);
}
/**
 * Run a registry lookup, converting its `StyleUnavailableError` into a bridge refusal.
 *
 * A separate function rather than an inline `try`/`catch` because a `catch` block that ends in a
 * `never` call still leaves the assignment indefinite to the type checker, and the alternative —
 * a `let` with a non-null assertion — would hide a real "not assigned" bug later.
 */
function attempt(read, onFailure) {
    try {
        return read();
    }
    catch {
        onFailure();
    }
}
export class CitationBridgeAdapter {
    #cslItem;
    #styles;
    #createProcessor;
    constructor(options) {
        this.#cslItem = "cslItems" in options ? options.cslItems : libraryCslItems(options.libraryId, options.items);
        this.#styles = options.styles;
        this.#createProcessor = options.createProcessor;
    }
    async previewCitation(request) {
        return this.#renderCluster(request);
    }
    /**
     * Identical to `previewCitation` today, and a test asserts they agree.
     *
     * The two routes are not redundant: `/citation/format` is the call whose output is *written*,
     * and when E10-06 gives it the document it will renumber the clusters around it while
     * `/citation/preview` still may not. Making them one implementation until that difference is
     * real is honest; making the preview quietly different would be a bug nobody could see.
     */
    async formatCitation(request) {
        return this.#renderCluster(request);
    }
    async formatBibliography(request) {
        const { style, locale, warnings } = this.#resolve(request.styleId, request.locale);
        const items = request.itemIds.map((itemId) => this.#requireItem(itemId));
        const renderer = new CslBibliographyRenderer({
            styleXml: style.style.xml,
            localeXml: locale.locale.xml,
            locale: locale.effectiveLocaleId,
            ...(this.#createProcessor === undefined ? {} : { createProcessor: this.#createProcessor }),
        });
        const rendered = renderer.renderCsl(items);
        // §9.4: the layout the style asks for, reported in CSL's vocabulary. It is read from the
        // style that was actually resolved, not the one that was requested, so a dependent style
        // cannot be described by its parent's indentation.
        const layout = bibliographyLayout(style.style.xml);
        return {
            entries: rendered.entries.map((entry) => ({ itemIds: entry.itemIds, text: entry.text })),
            context: this.#context(request.styleId, request.locale, style, locale),
            layout: {
                present: layout.present,
                hangingIndent: layout.hangingIndent,
                ...(layout.lineSpacing === undefined ? {} : { lineSpacing: layout.lineSpacing }),
                ...(layout.entrySpacing === undefined ? {} : { entrySpacing: layout.entrySpacing }),
                ...(layout.secondFieldAlign === undefined ? {} : { secondFieldAlign: layout.secondFieldAlign }),
                unrecognised: layout.unrecognised,
            },
            warnings,
        };
    }
    #renderCluster(request) {
        const { style, locale, warnings } = this.#resolve(request.styleId, request.locale);
        const projected = [];
        const seen = new Set();
        for (const item of request.items) {
            if (seen.has(item.itemId))
                continue;
            seen.add(item.itemId);
            projected.push(this.#requireItem(item.itemId));
        }
        const processor = this.#processor({
            styleXml: style.style.xml,
            localeXml: locale.locale.xml,
            items: projected,
            locale: locale.effectiveLocaleId,
            outputFormat: "text",
        });
        const makeCluster = processor.makeCitationCluster?.bind(processor);
        if (makeCluster === undefined) {
            // A processor that can only render a document in order is a valid `CslProcessor` (the
            // capability is optional); it just cannot answer this question, and saying so beats
            // substituting the document path and returning a number derived from nothing.
            refuse("preview-unsupported", "the configured CSL processor cannot render a cluster in isolation");
        }
        const format = citationFormatOf(style.style);
        if (format !== undefined && DOCUMENT_DEPENDENT_CITATION_FORMATS.includes(format)) {
            warnings.push({
                code: "isolated-render",
                message: `style "${style.effectiveStyleId}" declares citation-format "${format}", which is decided by the document; this cluster was rendered on its own`,
            });
        }
        if (request.items.length > 1 && citationSortsItems(style.style.xml)) {
            warnings.push({
                code: "style-sorted-order",
                message: `style "${style.effectiveStyleId}" sorts items inside a citation; changing draft order may not change printed order`,
            });
        }
        return {
            text: makeCluster(request.items.map((item) => toProcessorItem(toClusterItem(item)))),
            context: this.#context(request.styleId, request.locale, style, locale),
            warnings,
        };
    }
    #processor(options) {
        return this.#createProcessor === undefined
            ? new CiteprocJsAdapter(options)
            : this.#createProcessor(options);
    }
    /**
     * Resolve style and locale, refusing every substitution that changes what the citation says.
     *
     * `resolveStyle` is asked without a fallback id, but a registry may carry one of its own, so the
     * answer is checked rather than the request trusted.
     */
    #resolve(styleId, localeId) {
        const style = attempt(() => this.#styles.resolveStyle(styleId), () => refuse("unknown-style", `no CSL style is installed for "${styleId}"`));
        if (style.usedFallback) {
            refuse("unknown-style", `no CSL style is installed for "${styleId}" (${style.fallbackReason ?? "unavailable"}); rendering it in "${style.effectiveStyleId}" would change the citation format`);
        }
        const locale = attempt(() => this.#styles.resolveLocale(localeId), () => refuse("unknown-locale", `no CSL locale is installed for "${localeId}"`));
        const warnings = [];
        if (locale.usedFallback) {
            if (locale.fallbackReason !== "language-match") {
                refuse("unknown-locale", `no CSL locale is installed for "${localeId}"; rendering it in "${locale.effectiveLocaleId}" would change the citation's language`);
            }
            warnings.push({
                code: "locale-region-substituted",
                message: `locale "${localeId}" is not installed; "${locale.effectiveLocaleId}" was used`,
            });
        }
        return { style, locale, warnings };
    }
    #context(requestedStyleId, requestedLocale, style, locale) {
        const format = citationFormatOf(style.style);
        const updated = style.style.metadata.updated;
        return {
            requestedStyleId,
            effectiveStyleId: style.effectiveStyleId,
            // Absent, not substituted. See `BridgeRenderContext`: the `version` attribute is the CSL
            // schema version, and reporting it as the style's version would be a wrong answer written
            // into every document this product produces.
            ...(updated === undefined ? {} : { styleUpdated: updated }),
            cslVersion: style.style.metadata.version,
            styleRevision: style.style.revision,
            ...(format === undefined ? {} : { citationFormat: format }),
            requestedLocale,
            effectiveLocale: locale.effectiveLocaleId,
        };
    }
    #requireItem(itemId) {
        const item = this.#cslItem(itemId);
        if (item === undefined)
            refuse("unknown-item", `no item "${itemId}" in this library`);
        return item;
    }
}
/** Library-scoped and trash-excluding, exactly as the library port is — for the same reasons. */
function libraryCslItems(libraryId, items) {
    return (itemId) => {
        const item = items.findById(itemId);
        return item === undefined || item.libraryId !== libraryId ? undefined : toCslItem(item);
    };
}
/**
 * The bridge's request vocabulary is the document's (`locator: {label, value}`, `suppressAuthor`),
 * so a request item **is** a `ClusterItem` and `toProcessorItem` is the one place that maps it to
 * citeproc's. Restating that map here is how the preview and the document would drift apart.
 */
function toClusterItem(item) {
    return {
        itemId: item.itemId,
        ...(item.locator === undefined ? {} : { locator: { ...item.locator } }),
        ...(item.prefix === undefined ? {} : { prefix: item.prefix }),
        ...(item.suffix === undefined ? {} : { suffix: item.suffix }),
        ...(item.suppressAuthor === undefined ? {} : { suppressAuthor: item.suppressAuthor }),
        ...(item.authorOnly === undefined ? {} : { authorOnly: item.authorOnly }),
    };
}
