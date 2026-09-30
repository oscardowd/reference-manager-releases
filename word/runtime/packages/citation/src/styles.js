import { XMLParser, XMLValidator } from "fast-xml-parser";
import { CiteprocJsAdapter } from "./adapters/citeproc-js/index.js";
const CSL_NAMESPACE = "http://purl.org/net/xbiblio/csl";
/** The `collapse` values CSL 1.0.2 defines on `<citation>`. */
const CSL_COLLAPSE_MODES = new Set([
    "citation-number",
    "year",
    "year-suffix",
    "year-suffix-ranged",
]);
/**
 * The two modes that can print a range (`1–3`, `2020a–c`). `year` merges the years of one
 * author's works and `year-suffix` joins suffixes with a delimiter; neither ever ranges.
 */
const RANGE_COLLAPSE_MODES = new Set(["citation-number", "year-suffix-ranged"]);
const XML_PARSE_OPTIONS = {
    attributeNamePrefix: "@_",
    ignoreAttributes: false,
    parseTagValue: false,
    processEntities: false,
    removeNSPrefix: true,
    trimValues: false,
};
const xmlParser = new XMLParser(XML_PARSE_OPTIONS);
/** The values CSL 1.0.2 defines for `<bibliography second-field-align>`. */
const CSL_SECOND_FIELD_ALIGN = new Set(["flush", "margin"]);
export class StyleRegistryError extends Error {
    constructor(message) {
        super(message);
        this.name = "StyleRegistryError";
    }
}
export class StyleValidationError extends StyleRegistryError {
    issues;
    constructor(kind, issues) {
        super(`${kind} XML failed CSL validation (${issues.length} issue${issues.length === 1 ? "" : "s"})`);
        this.name = "StyleValidationError";
        this.issues = issues;
    }
}
export class StyleConflictError extends StyleRegistryError {
    id;
    constructor(id, kind = "style") {
        super(`A different ${kind} is already installed for "${id}"`);
        this.name = "StyleConflictError";
        this.id = id;
    }
}
export class StyleUnavailableError extends StyleRegistryError {
    id;
    reason;
    constructor(id, reason) {
        super(`No usable CSL ${reason === "missing-locale" ? "locale" : "style"} is installed for "${id}"`);
        this.name = "StyleUnavailableError";
        this.id = id;
        this.reason = reason;
    }
}
export class StyleUpdateError extends StyleRegistryError {
    constructor(message) {
        super(message);
        this.name = "StyleUpdateError";
    }
}
function asObject(value) {
    return value !== null && typeof value === "object" && !Array.isArray(value)
        ? value
        : undefined;
}
function asArray(value) {
    return value === undefined ? [] : Array.isArray(value) ? value : [value];
}
function childObjects(parent, key) {
    return asArray(parent[key]).map(asObject).filter((value) => value !== undefined);
}
function firstObject(parent, key) {
    return childObjects(parent, key)[0];
}
function attribute(parent, name) {
    const candidates = [`@_${name}`, `@_xml:${name}`, `@_${name.split(":").at(-1) ?? name}`];
    for (const key of candidates) {
        const value = parent[key];
        if (typeof value === "string" && value.trim().length > 0)
            return value.trim();
    }
    return undefined;
}
/**
 * An attribute exactly as the style wrote it.
 *
 * `attribute` trims, which is right for a title or an id and wrong for a delimiter: CSL's
 * `after-collapse-delimiter="; "` carries a trailing space that *is* the separator, and
 * reporting it as `";"` misstates the style to anyone who displays or compares it.
 */
function rawAttribute(parent, name) {
    const value = parent[`@_${name}`];
    return typeof value === "string" ? value : undefined;
}
function text(value) {
    if (typeof value === "string" || typeof value === "number") {
        const result = String(value).trim();
        return result.length > 0 ? result : undefined;
    }
    const object = asObject(value);
    if (object === undefined)
        return undefined;
    const direct = object["#text"];
    if (typeof direct === "string" || typeof direct === "number") {
        const result = String(direct).trim();
        return result.length > 0 ? result : undefined;
    }
    return undefined;
}
function childText(parent, key) {
    return text(parent[key]);
}
function parsedRoot(xml) {
    const issues = [];
    if (/<!(?:DOCTYPE|ENTITY)\b/iu.test(xml)) {
        issues.push({
            code: "xml-entity-declaration",
            path: "$",
            message: "XML declarations and entities are not accepted for installed CSL data",
            severity: "error",
        });
        return { issues };
    }
    const wellFormed = XMLValidator.validate(xml);
    if (wellFormed !== true) {
        issues.push({ code: "malformed-xml", path: "$", message: "XML is not well-formed", severity: "error" });
        return { issues };
    }
    try {
        const parsed = asObject(xmlParser.parse(xml));
        if (parsed === undefined) {
            issues.push({ code: "missing-root", path: "$", message: "XML has no root element", severity: "error" });
            return { issues };
        }
        const names = Object.keys(parsed).filter((key) => key !== "?xml");
        const rootName = names[0];
        const root = rootName === undefined ? undefined : asObject(parsed[rootName]);
        if (root === undefined || rootName === undefined) {
            issues.push({ code: "missing-root", path: "$", message: "XML has no root element", severity: "error" });
            return { issues };
        }
        return { root, rootName, issues };
    }
    catch {
        issues.push({ code: "parse-failed", path: "$", message: "XML could not be parsed", severity: "error" });
        return { issues };
    }
}
function parseLinks(info) {
    return childObjects(info, "link").flatMap((link) => {
        const rel = attribute(link, "rel");
        const href = attribute(link, "href");
        if (rel === undefined || href === undefined)
            return [];
        const citeprocVersion = attribute(link, "citeproc");
        const display = text(link);
        const result = {
            rel,
            href,
            ...(citeprocVersion === undefined ? {} : { citeprocVersion }),
            ...(display === undefined ? {} : { display }),
        };
        return [result];
    });
}
function styleMetadata(root, issues) {
    if (attribute(root, "xmlns") !== undefined && attribute(root, "xmlns") !== CSL_NAMESPACE) {
        issues.push({ code: "wrong-namespace", path: "/style/@xmlns", message: "style uses the wrong CSL namespace", severity: "error" });
    }
    const version = attribute(root, "version");
    if (version === undefined)
        issues.push({ code: "missing-version", path: "/style/@version", message: "style version is required", severity: "error" });
    const className = attribute(root, "class");
    if (className !== "in-text" && className !== "note") {
        issues.push({ code: "invalid-class", path: "/style/@class", message: "style class must be in-text or note", severity: "error" });
    }
    const info = firstObject(root, "info");
    if (info === undefined) {
        issues.push({ code: "missing-info", path: "/style/info", message: "style info is required", severity: "error" });
        return undefined;
    }
    const title = childText(info, "title");
    const id = childText(info, "id");
    if (title === undefined)
        issues.push({ code: "missing-title", path: "/style/info/title", message: "style title is required", severity: "error" });
    if (id === undefined)
        issues.push({ code: "missing-id", path: "/style/info/id", message: "style id is required", severity: "error" });
    const citation = firstObject(root, "citation");
    if (citation === undefined || firstObject(citation, "layout") === undefined) {
        issues.push({ code: "missing-citation-layout", path: "/style/citation/layout", message: "citation layout is required", severity: "error" });
    }
    const bibliography = firstObject(root, "bibliography");
    if (bibliography !== undefined && firstObject(bibliography, "layout") === undefined) {
        issues.push({ code: "missing-bibliography-layout", path: "/style/bibliography/layout", message: "bibliography layout is required when bibliography is present", severity: "error" });
    }
    const links = parseLinks(info);
    const parentLink = links.find((link) => link.rel === "independent-parent");
    const selfLink = links.find((link) => link.rel === "self");
    const categories = childObjects(info, "category").map((category) => {
        const citationFormat = attribute(category, "citation-format");
        const field = attribute(category, "field");
        return {
            ...(citationFormat === undefined ? {} : { citationFormat }),
            ...(field === undefined ? {} : { field }),
        };
    });
    const issns = asArray(info.issn).map((entry) => text(entry)).filter((value) => value !== undefined);
    const summary = childText(info, "summary");
    const updated = childText(info, "updated");
    const defaultLocale = attribute(root, "default-locale");
    const searchTerms = [title, id, summary, selfLink?.href, parentLink?.href, ...issns, ...categories.flatMap((category) => [category.citationFormat, category.field])]
        .filter((value) => value !== undefined);
    if (issues.some((issue) => issue.severity === "error") || id === undefined || title === undefined || version === undefined || className === undefined)
        return undefined;
    return {
        id,
        title,
        version,
        ...(updated === undefined ? {} : { updated }),
        ...(summary === undefined ? {} : { summary }),
        ...(defaultLocale === undefined ? {} : { defaultLocale }),
        className: className,
        kind: parentLink === undefined ? "independent" : "dependent",
        ...(parentLink === undefined ? {} : { independentParentId: parentLink.href }),
        ...(selfLink === undefined ? {} : { selfUrl: selfLink.href }),
        links,
        categories,
        issns,
        searchTerms,
    };
}
function localeMetadata(root, issues) {
    const language = attribute(root, "xml:lang") ?? attribute(root, "lang");
    const version = attribute(root, "version");
    if (language === undefined)
        issues.push({ code: "missing-language", path: "/locale/@xml:lang", message: "locale language is required", severity: "error" });
    if (version === undefined)
        issues.push({ code: "missing-version", path: "/locale/@version", message: "locale version is required", severity: "error" });
    if (firstObject(root, "terms") === undefined)
        issues.push({ code: "missing-terms", path: "/locale/terms", message: "locale terms are required", severity: "warning" });
    const title = childText(root, "title");
    const updated = childText(root, "updated");
    if (issues.some((issue) => issue.severity === "error") || language === undefined || version === undefined)
        return undefined;
    return {
        id: language,
        version,
        ...(title === undefined ? {} : { title }),
        ...(updated === undefined ? {} : { updated }),
        searchTerms: [language, title, updated].filter((value) => value !== undefined),
    };
}
export function validateCslStyle(xml) {
    const parsed = parsedRoot(xml);
    const issues = [...parsed.issues];
    if (parsed.rootName !== "style") {
        issues.push({ code: "wrong-root", path: "$", message: "CSL style XML must have a style root", severity: "error" });
        return { valid: false, issues };
    }
    const metadata = parsed.root === undefined ? undefined : styleMetadata(parsed.root, issues);
    return { valid: metadata !== undefined && !issues.some((issue) => issue.severity === "error"), ...(metadata === undefined ? {} : { metadata }), issues };
}
/**
 * Reports what the style's `<citation>` element says about collapsing (§8.2, task E08-07).
 *
 * This reads the style; it does not collapse anything. Producing `[1–3]` from `[1,2,3]` is
 * citeproc's job under a `collapse` rule (invariant 4), and nothing in this package may do it.
 * What a host *does* need to know is whether asking is even meaningful, because two questions
 * a user will ask have answers that live only in the style:
 *
 * - "why didn't my three citations collapse into a range?" — the style may declare no
 *   `collapse` at all, or declare `year-suffix`, which joins suffixes without ever printing a
 *   range;
 * - "how do I get a range?" — by putting the works in **one cluster**. Collapsing is a
 *   within-cluster rule: three adjacent single-item clusters render `[1][2][3]` and no style
 *   setting changes that, which is why merging citations is a document edit the host must make
 *   rather than a formatting option it can switch on. Verified against citeproc-js 2.4.63.
 *
 * Throws rather than reporting "no collapsing" for XML it could not read: a style that cannot
 * be parsed and a style that declares nothing are different facts, and only the second is a
 * statement about the style.
 */
export function citationCollapsing(xml) {
    const validation = validateCslStyle(xml);
    if (!validation.valid)
        throw new StyleValidationError("style", validation.issues);
    const parsed = parsedRoot(xml);
    const citation = parsed.root === undefined ? undefined : firstObject(parsed.root, "citation");
    // `collapse` is an enumerated token, so it is trimmed; the delimiter is verbatim data.
    const collapse = citation === undefined ? undefined : attribute(citation, "collapse");
    const afterCollapseDelimiter = citation === undefined ? undefined : rawAttribute(citation, "after-collapse-delimiter");
    return {
        ...(collapse === undefined ? {} : { collapse }),
        ...(afterCollapseDelimiter === undefined ? {} : { afterCollapseDelimiter }),
        collapsesRanges: collapse !== undefined && RANGE_COLLAPSE_MODES.has(collapse),
        recognisedMode: collapse === undefined || CSL_COLLAPSE_MODES.has(collapse),
    };
}
/**
 * Read the bibliography layout a style declares (§9.4 "hanging indents and line spacing").
 *
 * CSL puts these on the `<bibliography>` element, and they are the only statement anyone has
 * about how a style's bibliography should sit on the page — the processor renders entry *text*
 * and says nothing about indentation. So the host has three possible situations and they must
 * stay distinguishable: the style asks for a hanging indent, the style asks for none, or there
 * is no bibliography at all. Collapsing the third into the second would have a host silently
 * write an empty bibliography for a style that has none.
 *
 * Like `citationCollapsing`, this throws for XML it could not read rather than reporting
 * "declares nothing": a style that fails to parse and a style that declares no layout are
 * different facts and only the second is a statement about the style.
 */
export function bibliographyLayout(xml) {
    const validation = validateCslStyle(xml);
    if (!validation.valid)
        throw new StyleValidationError("style", validation.issues);
    const parsed = parsedRoot(xml);
    const bibliography = parsed.root === undefined ? undefined : firstObject(parsed.root, "bibliography");
    if (bibliography === undefined) {
        return { present: false, hangingIndent: false, unrecognised: [] };
    }
    const unrecognised = [];
    const spacing = (name, minimum) => {
        const raw = attribute(bibliography, name);
        if (raw === undefined)
            return undefined;
        // CSL 1.0.2: `line-spacing` is a positive integer, `entry-spacing` a non-negative one —
        // `entry-spacing="0"` is a real declaration meaning entries run on with no gap, and is not
        // the same as declaring nothing. `1.5` is a plausible thing for a person to write and is not
        // what the schema allows, so it is reported rather than rounded into something else.
        if (!/^\d+$/u.test(raw) || Number(raw) < minimum) {
            unrecognised.push(`${name}=${raw}`);
            return undefined;
        }
        return Number(raw);
    };
    const lineSpacing = spacing("line-spacing", 1);
    const entrySpacing = spacing("entry-spacing", 0);
    const hangingRaw = attribute(bibliography, "hanging-indent");
    if (hangingRaw !== undefined && hangingRaw !== "true" && hangingRaw !== "false") {
        unrecognised.push(`hanging-indent=${hangingRaw}`);
    }
    const secondFieldAlign = attribute(bibliography, "second-field-align");
    if (secondFieldAlign !== undefined && !CSL_SECOND_FIELD_ALIGN.has(secondFieldAlign)) {
        unrecognised.push(`second-field-align=${secondFieldAlign}`);
    }
    return {
        present: true,
        hangingIndent: hangingRaw === "true",
        ...(lineSpacing === undefined ? {} : { lineSpacing }),
        ...(entrySpacing === undefined ? {} : { entrySpacing }),
        ...(secondFieldAlign === undefined || !CSL_SECOND_FIELD_ALIGN.has(secondFieldAlign)
            ? {}
            : { secondFieldAlign }),
        unrecognised,
    };
}
/** Whether the style, rather than the author's draft order, sorts items inside one citation. */
export function citationSortsItems(xml) {
    const validation = validateCslStyle(xml);
    if (!validation.valid)
        throw new StyleValidationError("style", validation.issues);
    const parsed = parsedRoot(xml);
    const citation = parsed.root === undefined ? undefined : firstObject(parsed.root, "citation");
    return citation !== undefined && firstObject(citation, "sort") !== undefined;
}
export function validateCslLocale(xml) {
    const parsed = parsedRoot(xml);
    const issues = [...parsed.issues];
    if (parsed.rootName !== "locale") {
        issues.push({ code: "wrong-root", path: "$", message: "CSL locale XML must have a locale root", severity: "error" });
        return { valid: false, issues };
    }
    const metadata = parsed.root === undefined ? undefined : localeMetadata(parsed.root, issues);
    return { valid: metadata !== undefined && !issues.some((issue) => issue.severity === "error"), ...(metadata === undefined ? {} : { metadata }), issues };
}
function assertStyle(xml) {
    const result = validateCslStyle(xml);
    if (!result.valid || result.metadata === undefined)
        throw new StyleValidationError("style", result.issues);
    return result.metadata;
}
function assertLocale(xml) {
    const result = validateCslLocale(xml);
    if (!result.valid || result.metadata === undefined)
        throw new StyleValidationError("locale", result.issues);
    return result.metadata;
}
function localeKey(id) {
    return id.trim().toLowerCase();
}
function styleSearchScore(style, query, options) {
    if (options.origin !== undefined && style.origin !== options.origin)
        return 0;
    if (options.citationFormat !== undefined && !style.metadata.categories.some((category) => category.citationFormat === options.citationFormat))
        return 0;
    if (options.field !== undefined && !style.metadata.categories.some((category) => category.field === options.field))
        return 0;
    const needle = query.trim().toLocaleLowerCase();
    if (needle.length === 0)
        return 1;
    const terms = style.metadata.searchTerms.map((term) => term.toLocaleLowerCase());
    let score = 0;
    for (const term of terms) {
        if (term === needle)
            score = Math.max(score, 100);
        else if (term.startsWith(needle))
            score = Math.max(score, 75);
        else if (term.includes(needle))
            score = Math.max(score, 50);
    }
    return score;
}
export class CslStyleRegistry {
    styles = new Map();
    styleAliases = new Map();
    locales = new Map();
    fallbackStyleId;
    fallbackLocaleId;
    constructor(options = {}) {
        this.fallbackStyleId = options.fallbackStyleId;
        this.fallbackLocaleId = options.fallbackLocaleId;
        for (const style of options.styles ?? [])
            this.putStyle(style);
        for (const locale of options.locales ?? [])
            this.putLocale(locale);
    }
    installStyle(xml, options = {}) {
        const metadata = assertStyle(xml);
        const existing = this.findStyle(metadata.id);
        if (existing !== undefined && existing.xml !== xml)
            throw new StyleConflictError(metadata.id);
        if (existing !== undefined)
            return existing;
        const installed = { metadata, xml, origin: options.origin ?? "user", revision: 1, ...(options.etag === undefined ? {} : { etag: options.etag }) };
        this.putStyle(installed);
        return installed;
    }
    installLocale(xml, options = {}) {
        const metadata = assertLocale(xml);
        const existing = this.locales.get(localeKey(metadata.id));
        if (existing !== undefined && existing.xml !== xml)
            throw new StyleConflictError(metadata.id, "locale");
        if (existing !== undefined)
            return existing;
        const installed = { metadata, xml, origin: options.origin ?? "user", revision: 1, ...(options.etag === undefined ? {} : { etag: options.etag }) };
        this.putLocale(installed);
        return installed;
    }
    editStyle(id, xml) {
        const current = this.requireStyle(id);
        const metadata = assertStyle(xml);
        if (metadata.id !== current.metadata.id)
            throw new StyleUpdateError("A custom style edit must retain the installed style id");
        const edited = { metadata, xml, origin: "custom", revision: current.revision + 1 };
        this.replaceStyle(current, edited);
        return edited;
    }
    async updateStyle(id, source) {
        const current = this.requireStyle(id);
        if (current.origin === "custom")
            throw new StyleUpdateError(`Custom style "${id}" requires an explicit edit before remote update`);
        const response = await source.fetchStyle({ id: current.metadata.id, currentVersion: current.metadata.version, ...(current.etag === undefined ? {} : { etag: current.etag }) });
        if (response.status === "not-modified") {
            const unchanged = { ...current, ...(response.etag === undefined ? {} : { etag: response.etag }) };
            this.replaceStyle(current, unchanged);
            return { status: "not-modified", style: unchanged };
        }
        const metadata = assertStyle(response.xml);
        if (metadata.id !== current.metadata.id)
            throw new StyleUpdateError("A style update changed its identity");
        const updated = { metadata, xml: response.xml, origin: current.origin, revision: current.revision + 1, ...(response.etag === undefined ? {} : { etag: response.etag }) };
        this.replaceStyle(current, updated);
        return { status: "updated", style: updated };
    }
    async updateLocale(id, source) {
        const current = this.requireLocale(id);
        if (current.origin === "custom")
            throw new StyleUpdateError(`Custom locale "${id}" requires an explicit edit before remote update`);
        const response = await source.fetchLocale({ id: current.metadata.id, currentVersion: current.metadata.version, ...(current.etag === undefined ? {} : { etag: current.etag }) });
        if (response.status === "not-modified") {
            const unchanged = { ...current, ...(response.etag === undefined ? {} : { etag: response.etag }) };
            this.putLocale(unchanged);
            return { status: "not-modified", locale: unchanged };
        }
        const metadata = assertLocale(response.xml);
        if (localeKey(metadata.id) !== localeKey(current.metadata.id))
            throw new StyleUpdateError("A locale update changed its identity");
        const updated = { metadata, xml: response.xml, origin: current.origin, revision: current.revision + 1, ...(response.etag === undefined ? {} : { etag: response.etag }) };
        this.putLocale(updated);
        return { status: "updated", locale: updated };
    }
    listStyles() {
        return [...this.styles.values()];
    }
    listLocales() {
        return [...this.locales.values()];
    }
    searchStyles(query, options = {}) {
        return this.listStyles()
            .map((style) => ({ style, score: styleSearchScore(style, query, options) }))
            .filter((result) => result.score > 0)
            .sort((left, right) => right.score - left.score || left.style.metadata.title.localeCompare(right.style.metadata.title));
    }
    resolveStyle(id, options = {}) {
        const requested = id;
        const fallbackId = options.fallbackStyleId ?? this.fallbackStyleId;
        const result = this.resolveStyleInternal(id, new Set());
        if (result.style !== undefined)
            return { requestedStyleId: requested, effectiveStyleId: result.style.metadata.id, style: result.style, usedFallback: false };
        if (fallbackId !== undefined && fallbackId !== id) {
            const fallback = this.resolveStyleInternal(fallbackId, new Set());
            if (fallback.style !== undefined) {
                return { requestedStyleId: requested, effectiveStyleId: fallback.style.metadata.id, style: fallback.style, usedFallback: true, fallbackReason: result.reason };
            }
        }
        throw new StyleUnavailableError(id, result.reason);
    }
    resolveLocale(id, options = {}) {
        const requested = id;
        const exact = this.locales.get(localeKey(id));
        if (exact !== undefined)
            return { requestedLocaleId: requested, effectiveLocaleId: exact.metadata.id, locale: exact, usedFallback: false };
        const language = localeKey(id).split("-")[0];
        const languageMatch = this.listLocales().find((locale) => localeKey(locale.metadata.id).split("-")[0] === language);
        if (languageMatch !== undefined)
            return { requestedLocaleId: requested, effectiveLocaleId: languageMatch.metadata.id, locale: languageMatch, usedFallback: true, fallbackReason: "language-match" };
        const fallbackId = options.fallbackLocaleId ?? this.fallbackLocaleId;
        const fallback = fallbackId === undefined ? undefined : this.locales.get(localeKey(fallbackId));
        if (fallback !== undefined)
            return { requestedLocaleId: requested, effectiveLocaleId: fallback.metadata.id, locale: fallback, usedFallback: true, fallbackReason: "configured-fallback" };
        throw new StyleUnavailableError(id, "missing-locale");
    }
    previewStyle(styleId, items, options = {}) {
        const style = this.resolveStyle(styleId);
        const localeId = options.localeId ?? style.style.metadata.defaultLocale ?? "en-US";
        const locale = this.resolveLocale(localeId, options.fallbackLocaleId === undefined ? {} : { fallbackLocaleId: options.fallbackLocaleId });
        const processor = new CiteprocJsAdapter({
            styleXml: style.style.xml,
            localeXml: locale.locale.xml,
            items,
            locale: locale.effectiveLocaleId,
            ...(options.outputFormat === undefined ? {} : { outputFormat: options.outputFormat }),
        });
        const citation = options.citation === undefined ? undefined : processor.processCitationCluster(options.citation, options.citationsPre, options.citationsPost);
        return { style, locale, ...(citation === undefined ? {} : { citation }), bibliography: processor.makeBibliography(options.bibliographyItemIds) };
    }
    findStyle(id) {
        const direct = this.styles.get(id);
        if (direct !== undefined)
            return direct;
        const canonical = this.styleAliases.get(id);
        return canonical === undefined ? undefined : this.styles.get(canonical);
    }
    requireStyle(id) {
        const style = this.findStyle(id);
        if (style === undefined)
            throw new StyleUnavailableError(id, "missing-style");
        return style;
    }
    requireLocale(id) {
        const locale = this.locales.get(localeKey(id));
        if (locale === undefined)
            throw new StyleUnavailableError(id, "missing-locale");
        return locale;
    }
    resolveStyleInternal(id, seen) {
        const style = this.findStyle(id);
        if (style === undefined)
            return { reason: "missing-style" };
        if (style.metadata.kind === "independent")
            return { style, reason: "missing-style" };
        if (seen.has(style.metadata.id))
            return { reason: "parent-cycle" };
        seen.add(style.metadata.id);
        if (style.metadata.independentParentId === undefined)
            return { reason: "missing-independent-parent" };
        const parent = this.resolveStyleInternal(style.metadata.independentParentId, seen);
        return parent.style === undefined ? { reason: parent.reason === "missing-style" ? "missing-independent-parent" : parent.reason } : parent;
    }
    putStyle(style) {
        if (style.metadata.selfUrl !== undefined) {
            const aliasOwner = this.styleAliases.get(style.metadata.selfUrl);
            if (aliasOwner !== undefined && aliasOwner !== style.metadata.id)
                throw new StyleConflictError(style.metadata.selfUrl);
        }
        this.styles.set(style.metadata.id, style);
        if (style.metadata.selfUrl !== undefined)
            this.styleAliases.set(style.metadata.selfUrl, style.metadata.id);
    }
    replaceStyle(previous, next) {
        if (previous.metadata.selfUrl !== undefined && previous.metadata.selfUrl !== next.metadata.selfUrl)
            this.styleAliases.delete(previous.metadata.selfUrl);
        this.putStyle(next);
    }
    putLocale(locale) {
        this.locales.set(localeKey(locale.metadata.id), locale);
    }
}
