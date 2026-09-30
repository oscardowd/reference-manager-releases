/**
 * One picker row for a reference (SPEC §9.2) — pure, so every port that lists references can share
 * it: the desktop library (`library.ts`) and a document's own snapshots (`document-ports.ts`,
 * ADR-0242). A reference must look the same to the add-in whether it was searched for, created by
 * `/import`, or read out of the document it is already cited in.
 *
 * Nothing here formats a citation. `creatorSummary` is a **picker label**, and the citation port is
 * `citation.ts`.
 */
/** How many creator names a picker row shows before it says "et al.". */
export const PICKER_CREATOR_LIMIT = 3;
/**
 * A picker row's creator line — "Ó Súilleabháin A, Example B" — for a human choosing between
 * search results.
 *
 * **This is a list label and never citation text.** It is not produced by the CSL processor, it is
 * never written into a document, and nothing in `citation.ts` calls it: invariant 4 governs what
 * a citation says, and a search result is not a citation. Keeping the two in separate modules and
 * naming this one `picker…` is the cheapest way to keep a later edit from confusing them.
 */
export function pickerCreatorLabel(creators) {
    if (creators.length === 0)
        return "";
    // Authors when there are any; otherwise whichever role sorts first, so an edited volume shows
    // its editors rather than an empty line.
    const roles = [...new Set(creators.map((creator) => creator.role))].sort();
    const role = roles.includes("author") ? "author" : roles[0];
    const chosen = creators
        .filter((creator) => creator.role === role)
        .slice()
        .sort((left, right) => left.ordinal - right.ordinal);
    const names = chosen.slice(0, PICKER_CREATOR_LIMIT).map((creator) => {
        if (creator.literal !== undefined && creator.literal.length > 0)
            return creator.literal;
        const family = [creator.nonDroppingParticle, creator.family].filter(Boolean).join(" ").trim();
        const initials = (creator.given ?? "")
            .split(/[\s.-]+/u)
            .filter((part) => part.length > 0)
            .map((part) => [...part][0] ?? "")
            .join("");
        return [family, initials].filter((part) => part.length > 0).join(" ");
    });
    const label = names.filter((name) => name.length > 0).join(", ");
    return chosen.length > PICKER_CREATOR_LIMIT && label.length > 0 ? `${label} et al.` : label;
}
/** The four-digit year a picker row shows, or nothing — never a guess at a partial date. */
export function pickerYear(item) {
    if (item.issuedYear !== undefined)
        return String(item.issuedYear);
    const match = /\b(\d{4})\b/u.exec(item.issuedRaw ?? "");
    return match?.[1];
}
/**
 * One picker row. Exported because `/import` returns rows for records it has just written, and a
 * reference must look the same to the add-in whether it was searched for or created.
 */
export function toSearchHit(item) {
    return {
        itemId: item.id,
        itemType: item.type,
        title: item.title ?? "",
        creatorSummary: pickerCreatorLabel(item.creators),
        year: pickerYear(item),
        containerTitle: item.containerTitle,
        doi: item.doi,
    };
}
/** The styles a picker may offer, as `/styles` lists them, whatever registry backs the port. */
export function styleSummaries(styles) {
    return styles
        .listStyles()
        .map((style) => ({
        id: style.metadata.id,
        title: style.metadata.title,
        // A style may declare several categories and only some carry `citation-format`; the first
        // that does is the style's answer. `undefined` means the style did not say, which is not
        // the same as "in-text" and must not be filled in here.
        citationFormat: style.metadata.categories.find((category) => category.citationFormat !== undefined)
            ?.citationFormat,
    }))
        .sort((left, right) => left.title.localeCompare(right.title) || left.id.localeCompare(right.id));
}
