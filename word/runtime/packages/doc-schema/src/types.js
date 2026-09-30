/**
 * Portable document citation schema — ADR-0015, SPEC §9.3 and §10.
 *
 * This is the payload every writing integration transports (Word custom XML part per
 * ADR-0007, Google Docs, LibreOffice, temporary citation syntax). The document must remain
 * fully editable with no library present, so every cited item carries a CSL-JSON snapshot.
 *
 * Version 2. Any change to these shapes bumps DOC_SCHEMA_VERSION and ships a payload
 * migration + test (ADR-0010).
 *
 * Version 2 added `BibliographyState.scopes` (ADR-0123, R-072): a document may hold more than
 * one bibliography control, and version 1 had nothing that could tell them apart.
 */
export const DOC_SCHEMA_VERSION = 2;
/**
 * CSL locator labels (CSL 1.0.2). Unknown labels are preserved but flagged by validation as
 * warnings so a newer CSL vocabulary does not brick older documents.
 */
export const CSL_LOCATOR_LABELS = [
    "act",
    "appendix",
    "article-locator",
    "book",
    "canon",
    "chapter",
    "column",
    "elocation",
    "equation",
    "figure",
    "folio",
    "issue",
    "line",
    "note",
    "opus",
    "page",
    "paragraph",
    "part",
    "rule",
    "scene",
    "section",
    "sub-verbo",
    "supplement",
    "table",
    "timestamp",
    "title-locator",
    "verse",
    "version",
    "volume",
];
