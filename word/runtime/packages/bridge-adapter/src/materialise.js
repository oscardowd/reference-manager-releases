/**
 * Turning a reconciled metadata record into a canonical `ReferenceItem` (SPEC §5.1, §4.4;
 * task E10-04.2).
 *
 * E04-08.2 retrieves candidates, E04-09 reconciles them, and `resolveReconciliation` says which
 * values may be written. Nothing before this module has ever produced a *record*: the pipeline
 * stopped at a map of canonical paths to resolved values, and the step from that map to a row in
 * the library did not exist. This is that step, and it is the only place in the product where a
 * reference is created out of what a metadata provider said.
 *
 * Two rules decide everything here.
 *
 * - **Nothing is dropped (invariant 2).** A canonical path this build does not recognise, a field
 *   key that is not in the registry, a value whose type the registry refuses — each is preserved
 *   under {@link UNMAPPED_NAMESPACE} rather than discarded. The raw provider responses are already
 *   kept under their own source namespaces by the reconciler; this namespace is for the values
 *   that *were* canonical and still could not be stored canonically, which is a different loss and
 *   needs its own place to be found later.
 * - **`validateReferenceItem` is the only authority on what is storable.** A second opinion about
 *   field types living here would drift from the registry, and the failure mode of that drift is
 *   an import that writes a value the rest of the product cannot read. So the item is built, then
 *   validated, and any field the validator rejects is *moved* to the unmapped namespace and the
 *   item re-validated — the rejection itself decides what is unmappable.
 *
 * It performs no I/O, generates no ids of its own and reads no clock: everything variable is
 * passed in, so a test can pin the whole output.
 */
import { DEFAULT_FIELD_REGISTRY, DEFAULT_ITEM_TYPE_REGISTRY, validateReferenceItem, } from "@refmgr/core";
/**
 * Where a canonical value that could not be stored canonically goes.
 *
 * Namespaced like a source (invariant 2's `extensions[<namespace>]`) and named for what it is:
 * these values came from this build's import pipeline, not from a provider called `refmgr`.
 */
export const UNMAPPED_NAMESPACE = "refmgr:unmapped-import";
/** The `fields.` prefix the reconciler uses for sparse registered fields. */
const FIELD_PREFIX = "fields.";
/** Personal-name and identity components a provider may supply for a creator. */
const CREATOR_PARTS = [
    "family",
    "given",
    "suffix",
    "droppingParticle",
    "nonDroppingParticle",
    "literal",
    "orcid",
    "affiliation",
];
function isJsonObject(value) {
    return typeof value === "object" && value !== null && !Array.isArray(value);
}
function nonEmptyString(value) {
    return typeof value === "string" && value.trim().length > 0 ? value : undefined;
}
function wholeYear(value) {
    return typeof value === "number" && Number.isInteger(value) && value >= 0 ? value : undefined;
}
/** `/fields/3/value` → 3. Anything else → `undefined`. */
function fieldIssueIndex(path) {
    const match = /^\/fields\/(\d+)(?:\/|$)/u.exec(path);
    return match === null ? undefined : Number(match[1]);
}
/**
 * Creators, renumbered per role from zero.
 *
 * The reconciler carries whatever ordinals the sources used, and a provider that numbers its
 * authors from one — or a reconciled list that took an editor's ordinals from a different source —
 * produces a gap the canonical model refuses. Renumbering preserves *order*, which is the thing a
 * citation depends on; it does not reorder, and it does not invent a creator.
 *
 * A structurally impossible list (not an array, or an entry that is not an object) is not
 * repaired: the whole value is reported unmappable and kept verbatim.
 */
function creatorsFrom(value) {
    if (!Array.isArray(value))
        return undefined;
    const nextOrdinal = new Map();
    const creators = [];
    for (const entry of value) {
        if (!isJsonObject(entry))
            return undefined;
        const role = nonEmptyString(entry.role) ?? "author";
        const ordinal = nextOrdinal.get(role) ?? 0;
        nextOrdinal.set(role, ordinal + 1);
        const parts = {};
        for (const key of CREATOR_PARTS) {
            const text = nonEmptyString(entry[key]);
            if (text !== undefined)
                parts[key] = text;
        }
        creators.push({ ordinal, role, ...parts });
    }
    return creators;
}
/**
 * Build a canonical item from resolved values, or say why it cannot be built.
 *
 * Failure is reserved for what makes a record *unstorable* — an unregistered item type, a
 * creator list the model cannot represent. Missing recommended metadata is deliberately not a
 * failure: `validateReferenceItem` treats "a journal article with no journal" as advisory (§4.4,
 * §14), and this module does not get to hold a stricter opinion than the model it writes into.
 */
export function materialiseReference(input) {
    const registry = input.registry ?? DEFAULT_FIELD_REGISTRY;
    const itemTypes = input.itemTypes ?? DEFAULT_ITEM_TYPE_REGISTRY;
    const unmapped = {};
    const unmappedPaths = [];
    const keep = (path, value) => {
        unmapped[path] = value;
        unmappedPaths.push(path);
    };
    // The type is the one canonical value with no fallback: an unregistered type is refused rather
    // than replaced, because every substitute is a claim about what kind of work this is.
    const rawType = input.fields["type"]?.value;
    const definition = typeof rawType === "string" ? itemTypes.get(rawType) : undefined;
    if (definition === undefined) {
        return {
            ok: false,
            issues: [
                {
                    path: "/type",
                    code: "unknown-item-type",
                    message: rawType === undefined
                        ? "the retrieved metadata does not say what kind of work this is"
                        : "the retrieved metadata names an item type this build does not know",
                },
            ],
        };
    }
    const item = {
        id: input.itemId,
        libraryId: input.libraryId,
        // The registry's own key, so an alias a provider used ("chapter") is stored as the canonical
        // type ("book-section") rather than as a second spelling of it.
        type: definition.key,
        createdAt: input.createdAt,
        updatedAt: input.createdAt,
        version: 1,
        fields: [],
        creators: [],
        extensions: {},
    };
    for (const [path, resolved] of Object.entries(input.fields)) {
        const value = resolved.value;
        if (path === "type")
            continue;
        if (path === "creators") {
            const creators = creatorsFrom(value);
            if (creators === undefined)
                keep(path, value);
            else
                item.creators = [...creators];
            continue;
        }
        if (path === "issuedYear") {
            const year = wholeYear(value);
            if (year === undefined)
                keep(path, value);
            else
                item.issuedYear = year;
            continue;
        }
        if (path === "title" || path === "containerTitle" || path === "issuedRaw" || path === "doi" || path === "pmid" || path === "citationKey") {
            const text = nonEmptyString(value);
            if (text === undefined)
                keep(path, value);
            else
                item[path] = text;
            continue;
        }
        if (!path.startsWith(FIELD_PREFIX)) {
            // A canonical path this build has never heard of. Kept, never guessed at.
            keep(path, value);
            continue;
        }
        item.fields.push({
            key: path.slice(FIELD_PREFIX.length),
            value,
            ordinal: 0,
            ...(resolved.provenance === undefined ? {} : { provenance: resolved.provenance }),
        });
    }
    // The validator decides which sparse fields are storable. Anything it rejects moves to the
    // unmapped namespace and the item is re-checked. The loop terminates because every pass either
    // removes at least one field or returns.
    let validation = validateReferenceItem(item, registry, itemTypes);
    while (!validation.ok) {
        const rejected = new Set(validation.issues
            .map((issue) => fieldIssueIndex(issue.path))
            .filter((index) => index !== undefined));
        if (rejected.size === 0) {
            return { ok: false, issues: validation.issues };
        }
        const remaining = [];
        item.fields.forEach((field, index) => {
            if (rejected.has(index))
                keep(`${FIELD_PREFIX}${field.key}`, field.value);
            else
                remaining.push(field);
        });
        item.fields = remaining;
        validation = validateReferenceItem(item, registry, itemTypes);
    }
    item.extensions = {
        ...(input.extensions ?? {}),
        ...(unmappedPaths.length === 0 ? {} : { [UNMAPPED_NAMESPACE]: unmapped }),
    };
    return { ok: true, item, unmapped: unmappedPaths };
}
