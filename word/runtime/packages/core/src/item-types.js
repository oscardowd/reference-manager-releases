function requireKey(value, label) {
    if (!/^[A-Za-z][A-Za-z0-9_-]*$/.test(value)) {
        throw new Error(`${label} must be a non-empty portable key`);
    }
}
/** Data-driven registry for the semantic reference types supported by the product. */
export class ItemTypeRegistry {
    #byKey;
    #definitions;
    constructor(definitions) {
        const byKey = new Map();
        const frozenDefinitions = [];
        const keys = new Set();
        for (const raw of definitions) {
            requireKey(raw.key, "item type key");
            if (keys.has(raw.key))
                throw new Error(`duplicate item type key "${raw.key}"`);
            keys.add(raw.key);
        }
        const aliasesSeen = new Set();
        for (const raw of definitions) {
            if (raw.label.trim().length === 0)
                throw new Error(`item type "${raw.key}" needs a label`);
            requireKey(raw.cslType, "CSL type");
            const aliases = [...(raw.aliases ?? [])];
            for (const alias of aliases) {
                requireKey(alias, `alias for item type "${raw.key}"`);
                if (alias === raw.key || keys.has(alias) || aliasesSeen.has(alias)) {
                    throw new Error(`duplicate item type alias "${alias}"`);
                }
                aliasesSeen.add(alias);
            }
            const definition = Object.freeze({
                ...raw,
                requiredFields: Object.freeze([...raw.requiredFields]),
                ...(raw.aliases === undefined ? {} : { aliases: Object.freeze(aliases) }),
            });
            frozenDefinitions.push(definition);
            byKey.set(definition.key, definition);
            for (const alias of aliases) {
                byKey.set(alias, definition);
            }
        }
        this.#definitions = Object.freeze(frozenDefinitions);
        this.#byKey = byKey;
    }
    get(key) {
        return this.#byKey.get(key);
    }
    has(key) {
        return this.#byKey.has(key);
    }
    definitions() {
        return this.#definitions;
    }
    with(definition) {
        return new ItemTypeRegistry([...this.#definitions, definition]);
    }
}
/**
 * The 25 semantic types named by SPEC §4.1. Their keys are stable application values; `cslType`
 * records the closest standard CSL type without collapsing distinct product types such as a
 * podcast and a video into the same stored value.
 */
export const DEFAULT_ITEM_TYPE_DEFINITIONS = [
    {
        key: "article-journal",
        label: "Journal article",
        cslType: "article-journal",
        requiredFields: ["title", "containerTitle", "issued"],
    },
    {
        key: "preprint",
        label: "Preprint",
        cslType: "article",
        requiredFields: ["title", "issued"],
        aliases: ["article"],
    },
    { key: "book", label: "Book", cslType: "book", requiredFields: ["title", "issued"] },
    {
        key: "book-section",
        label: "Book section",
        cslType: "chapter",
        requiredFields: ["title", "containerTitle", "issued"],
        aliases: ["chapter"],
    },
    {
        key: "conference-paper",
        label: "Conference paper",
        cslType: "paper-conference",
        requiredFields: ["title", "containerTitle", "issued"],
        aliases: ["paper-conference"],
    },
    {
        key: "conference-presentation",
        label: "Conference presentation",
        cslType: "speech",
        requiredFields: ["title", "issued"],
        aliases: ["speech"],
    },
    { key: "thesis", label: "Thesis", cslType: "thesis", requiredFields: ["title", "issued"] },
    { key: "report", label: "Report", cslType: "report", requiredFields: ["title", "issued"] },
    { key: "dataset", label: "Dataset", cslType: "dataset", requiredFields: ["title", "issued"] },
    { key: "software", label: "Software", cslType: "software", requiredFields: ["title", "issued"] },
    {
        key: "webpage",
        label: "Webpage",
        cslType: "webpage",
        requiredFields: ["title", "field:url"],
    },
    { key: "patent", label: "Patent", cslType: "patent", requiredFields: ["title", "issued"] },
    { key: "standard", label: "Standard", cslType: "standard", requiredFields: ["title", "issued"] },
    {
        key: "clinical-trial",
        label: "Clinical trial",
        cslType: "report",
        requiredFields: ["title", "issued"],
    },
    {
        key: "legislation",
        label: "Legislation",
        cslType: "legislation",
        requiredFields: ["title", "issued"],
    },
    {
        key: "case",
        label: "Case",
        cslType: "legal_case",
        requiredFields: ["title", "issued"],
        aliases: ["legal_case"],
    },
    {
        key: "article-newspaper",
        label: "Newspaper article",
        cslType: "article-newspaper",
        requiredFields: ["title", "containerTitle", "issued"],
    },
    {
        key: "article-magazine",
        label: "Magazine article",
        cslType: "article-magazine",
        requiredFields: ["title", "containerTitle", "issued"],
    },
    {
        key: "podcast",
        label: "Podcast",
        cslType: "broadcast",
        requiredFields: ["title"],
        aliases: ["broadcast"],
    },
    {
        key: "video",
        label: "Video",
        cslType: "motion_picture",
        requiredFields: ["title"],
        aliases: ["motion_picture"],
    },
    {
        key: "audio-recording",
        label: "Audio recording",
        cslType: "song",
        requiredFields: ["title"],
        aliases: ["song"],
    },
    { key: "presentation", label: "Presentation", cslType: "speech", requiredFields: ["title"] },
    { key: "manuscript", label: "Manuscript", cslType: "manuscript", requiredFields: ["title"] },
    {
        key: "personal-communication",
        label: "Personal communication",
        cslType: "personal_communication",
        requiredFields: ["title"],
        aliases: ["personal_communication"],
    },
    { key: "document", label: "Generic document", cslType: "document", requiredFields: ["title"] },
];
export const DEFAULT_ITEM_TYPE_REGISTRY = new ItemTypeRegistry(DEFAULT_ITEM_TYPE_DEFINITIONS);
/**
 * Map a stored or imported type spelling onto the registry's own key. Unknown values pass
 * through so validation can still refuse them.
 */
export function canonicalItemType(type, registry = DEFAULT_ITEM_TYPE_REGISTRY) {
    return registry.get(type)?.key ?? type;
}
/**
 * CSL type the citation processor should see. Semantic keys such as `book-section` and `podcast`
 * stay in the model; citeproc only understands the standard hint.
 */
export function processorCslType(type, registry = DEFAULT_ITEM_TYPE_REGISTRY) {
    return registry.get(type)?.cslType ?? type;
}
/** Backwards-friendly names for callers that treat the registry as the reference-type catalog. */
export const REFERENCE_ITEM_TYPE_DEFINITIONS = DEFAULT_ITEM_TYPE_DEFINITIONS;
export const REFERENCE_ITEM_TYPES = DEFAULT_ITEM_TYPE_REGISTRY;
