/**
 * Adding a reference without the desktop (SPEC §9.2 "Allow creating or importing a missing reference
 * without leaving Word"; ADR-0242 decision 5).
 *
 * Document mode has no library to import into, so a reference the writer adds is kept on the pane's
 * shelf and becomes the citation's §9.3 snapshot when cited. Both ways in reuse the desktop's own
 * pipeline rather than a pane-only copy of it:
 *
 * - **By identifier:** `IdentifierBridgeImporter` — recognition, retrieval, reconciliation and the
 *   canonical materialisation — over a target that looks in the document and the shelf instead of a
 *   repository. A DOI or PMID the document already cites comes back as that reference, not a copy.
 * - **By hand:** the fields a writer types go through the same `materialiseReference` a metadata
 *   source's answer does, so a hand-made reference is validated exactly as an imported one is.
 *
 * Either way the result is CSL-JSON projected by `toCslItem`: the same bytes a library item would
 * have produced, so a document cannot tell whether its snapshot came from a library or a pane.
 */
import { fromCslItem, generateUuidV7, toCslItem } from "@refmgr/core";
import { RETRIEVAL_DISCLOSURES, createMetadataSources, } from "@refmgr/identifiers";
import { IdentifierBridgeImporter } from "./import.js";
import { materialiseReference } from "./materialise.js";
/**
 * The library id an in-pane reference carries while it is a `ReferenceItem`. It never reaches a
 * document: document mode records no `sourceLibraryId` on a snapshot (ADR-0242 decision 5).
 */
export const DOCUMENT_LIBRARY_ID = "document";
/**
 * The metadata sources a browser can ask. arXiv's API sends no CORS headers, so the pane cannot read
 * its answers; an arXiv id is reported as having no source here rather than as a network failure.
 */
export const DOCUMENT_MODE_SOURCES = Object.freeze(["crossref", "datacite", "pubmed", "openlibrary", "clinicaltrials"]);
/** Who learns what when the pane looks an identifier up — for the sentence shown before it does. */
export function documentModeDisclosures() {
    return RETRIEVAL_DISCLOSURES.filter((entry) => DOCUMENT_MODE_SOURCES.includes(entry.source)).map((entry) => ({
        host: entry.host,
        kinds: entry.kinds,
    }));
}
function asItem(reference) {
    try {
        return fromCslItem(reference.csl, { libraryId: DOCUMENT_LIBRARY_ID, createdAt: reference.addedAt });
    }
    catch {
        return undefined;
    }
}
function cslString(reference, key) {
    const value = reference.csl[key];
    return typeof value === "string" ? value : undefined;
}
export function createDocumentImporter(options) {
    const find = (matches) => {
        const reference = options.references().find(matches);
        return reference === undefined ? undefined : asItem(reference);
    };
    const target = {
        // DOIs compare case-insensitively (the DOI Handbook, §2.4); PMIDs are digits.
        findByDoi: (_libraryId, doi) => find((reference) => cslString(reference, "DOI")?.toLowerCase() === doi.toLowerCase()),
        findByPmid: (_libraryId, pmid) => find((reference) => cslString(reference, "PMID") === pmid),
        insert: (item) => {
            options.add({ itemId: item.id, csl: toCslItem(item), addedAt: item.createdAt });
            return item;
        },
    };
    const importer = new IdentifierBridgeImporter({
        libraryId: DOCUMENT_LIBRARY_ID,
        items: target,
        retrieval: {
            transport: options.transport,
            sources: createMetadataSources().filter((source) => DOCUMENT_MODE_SOURCES.includes(source.id)),
        },
        ...(options.now === undefined ? {} : { now: options.now }),
        newId: options.newId ?? generateUuidV7,
    });
    return (request) => importer.import(request);
}
function text(value) {
    const trimmed = value?.trim();
    return trimmed === undefined || trimmed.length === 0 ? undefined : trimmed;
}
/**
 * A reference from what the writer typed, validated by the same materialisation an imported one
 * goes through. The fields are the ones a metadata source supplies, under the same canonical paths,
 * with the writer as their provenance.
 */
export function manualReference(input, options = {}) {
    const now = (options.now ?? (() => new Date().toISOString()))();
    const provenance = { source: "user", retrievedAt: now, confidence: 1 };
    const field = (value) => ({ value, provenance });
    const title = text(input.title);
    if (title === undefined)
        return { ok: false, message: "A reference needs a title." };
    const year = text(input.year);
    if (year !== undefined && !/^\d{1,4}$/u.test(year))
        return { ok: false, message: "The year must be a number, such as 2024." };
    const creators = input.authors
        .map((entry) => entry.trim())
        .filter((entry) => entry.length > 0)
        .map((entry) => {
        const comma = entry.indexOf(",");
        if (comma < 0)
            return { role: "author", literal: entry };
        const family = entry.slice(0, comma).trim();
        const given = entry.slice(comma + 1).trim();
        return given.length === 0 ? { role: "author", family } : { role: "author", family, given };
    });
    const fields = { type: field(input.type), title: field(title) };
    if (creators.length > 0)
        fields["creators"] = field(creators);
    if (year !== undefined)
        fields["issuedYear"] = field(Number(year));
    const optional = [
        ["containerTitle", input.containerTitle],
        ["doi", input.doi],
        ["fields.publisher", input.publisher],
        ["fields.volume", input.volume],
        ["fields.issue", input.issue],
        ["fields.pages", input.pages],
        ["fields.url", input.url],
    ];
    for (const [path, value] of optional) {
        const present = text(value);
        if (present !== undefined)
            fields[path] = field(present);
    }
    const material = materialiseReference({
        itemId: (options.newId ?? generateUuidV7)(),
        libraryId: DOCUMENT_LIBRARY_ID,
        createdAt: now,
        fields,
    });
    if (!material.ok) {
        return { ok: false, message: material.issues[0]?.message ?? "Those details cannot be saved as a reference." };
    }
    const csl = toCslItem(material.item);
    return { ok: true, reference: { itemId: material.item.id, csl, addedAt: now } };
}
