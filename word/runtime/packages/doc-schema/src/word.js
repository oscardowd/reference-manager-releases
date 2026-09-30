/** Word transport descriptors shared by the Office.js add-in and the OOXML evidence harness. */
import { BIBLIOGRAPHY_TAG, bibliographyTag, clusterTag, DOCUMENT_SCOPE_ID } from "./canonical.js";
import { writeDocumentPart } from "./read.js";
export const WORD_CUSTOM_XML_NAMESPACE = "urn:refmgr:doc-schema:1";
function escapeXml(value) {
    return value
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&apos;");
}
/** The exact descriptor both Office.js and the OOXML harness write for a cluster. */
export function wordCitationControl(cluster) {
    const visibleText = cluster.manualEdit.state !== "none" && cluster.manualEdit.text !== undefined
        ? cluster.manualEdit.text
        : cluster.renderedText;
    if (visibleText === undefined || visibleText.length === 0) {
        throw new Error(`cluster ${cluster.clusterId} has no visible processor output`);
    }
    return Object.freeze({
        tag: clusterTag(cluster.clusterId, cluster),
        title: "Citation",
        appearance: "BoundingBox",
        visibleText,
    });
}
/**
 * An ADR-0007 bibliography control, as both Office.js and the OOXML harness write it.
 *
 * ADR-0007 left this tag unfingerprinted because there was one bibliography per document.
 * R-072 is that assumption meeting a thesis with per-chapter references: with nothing in the
 * tag, a refresh could not tell Chapter 1's control from Chapter 2's and wrote the whole
 * document into both. A control now carries its **scope** — which records it lists — and the
 * fingerprint covers that scope, so it moves when the author changes what the chapter covers
 * and stays put through ordinary citation edits (ADR-0123).
 *
 * `scope` is optional and its absence means the whole document, which is both the version-1
 * shape and the ordinary single-bibliography manuscript.
 */
export function wordBibliographyControl(entries, paragraphFormat, scope) {
    for (const entry of entries) {
        if (entry.text.length === 0) {
            throw new Error(`bibliography entry for ${entry.itemIds.join(", ") || "no item"} has no processor output`);
        }
        if (entry.itemIds.length === 0) {
            throw new Error("a bibliography entry was not attributed to any item");
        }
    }
    const scopeId = scope?.scopeId ?? DOCUMENT_SCOPE_ID;
    if (scope !== undefined && scope.mode === "items" && (scope.itemIds ?? []).length === 0) {
        throw new Error(`bibliography scope ${scopeId} lists no items; it would render as an empty bibliography`);
    }
    return Object.freeze({
        tag: scope === undefined ? BIBLIOGRAPHY_TAG : bibliographyTag(scope),
        scopeId,
        title: "Bibliography",
        appearance: "BoundingBox",
        entries: Object.freeze(entries.map((entry) => Object.freeze({ itemIds: Object.freeze([...entry.itemIds]), text: entry.text }))),
        paragraphFormat: paragraphFormat === null ? null : Object.freeze({ ...paragraphFormat }),
    });
}
/** XML handed to `Office.context.document.customXmlParts.addAsync`. */
export function wordCustomXml(part) {
    const written = writeDocumentPart(part);
    if (written.status !== "ok") {
        const issues = written.issues.map((issue) => `${issue.path} ${issue.code}`).join(", ");
        throw new Error(`refusing to write the Word custom XML part: ${written.message}${issues ? ` (${issues})` : ""}`);
    }
    return (`<refmgr:document xmlns:refmgr="${WORD_CUSTOM_XML_NAMESPACE}">` +
        `<refmgr:payload contentType="application/json">${escapeXml(written.json)}</refmgr:payload>` +
        `</refmgr:document>`);
}
