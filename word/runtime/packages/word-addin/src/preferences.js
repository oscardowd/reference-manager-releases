/**
 * Change the citation style and the document preferences (SPEC §9.1, task E10-08; ADR-0111).
 *
 * Two commands, one module, because they are the same decision at two sizes. `updateMode` and the
 * bibliography lock change how the document behaves and rewrite nothing the reader can see, so they
 * are a payload-only write. Style and locale change **every citation in the manuscript**, so they
 * are a full re-render — and this module refuses to do one without processor output that was
 * actually made with the new style.
 *
 * Nothing here formats or renumbers anything. The re-render is E10-07's `prepareDocumentRefresh`
 * over `CitationDocumentEditor`'s output (invariant 4); a second style-aware rendering path would
 * be a second account of numbering, and the wrong one would be indistinguishable from the right one
 * until a reader noticed.
 *
 * The through-line of the refusals is §14/invariant 3: a style change must never *substitute*.
 * A user who asked for AMA and silently got Chicago has a manuscript that is wrong in a way no
 * error message ever mentioned.
 */
import { wordCustomXml, } from "@refmgr/doc-schema";
import {} from "./bibliography.js";
import { clearNoteNumbers } from "./notes.js";
import { prepareDocumentRefresh, } from "./refresh.js";
function refuse(code, message) {
    return { ok: false, refusal: { code, message } };
}
/** BCP-47's primary language subtag. `en-GB` and `en-US` share one; `de-DE` does not. */
export function localeLanguage(tag) {
    return tag.split("-")[0].toLowerCase();
}
/**
 * Set the document options that change no visible text (§9.1 "Document preferences").
 *
 * Style and locale are deliberately **refused** here rather than quietly accepted: applying either
 * without re-rendering would leave a payload that claims one style and a manuscript that shows
 * another, which is the exact state document diagnostics (§9.5) exist to find. Pinning and uncited
 * entries are absent for the same reason in reverse — they change what the bibliography *lists*, so
 * they belong with a regeneration (`setBibliographyOptions` plus `prepareBibliography`).
 */
export function applyDocumentPreferences(input) {
    const { part, request } = input;
    if (request.style !== undefined && request.style.id !== part.style.id) {
        return refuse("requires-rerender", "changing the citation style re-renders every citation; use prepareStyleChange");
    }
    if (request.locale !== undefined && request.locale !== part.locale) {
        return refuse("requires-rerender", "changing the locale re-renders every citation; use prepareStyleChange");
    }
    const changes = [];
    const preferences = { ...part.preferences };
    if (request.updateMode !== undefined && request.updateMode !== part.preferences.updateMode) {
        changes.push({
            field: "update-mode",
            ...(part.preferences.updateMode === undefined ? {} : { from: part.preferences.updateMode }),
            to: request.updateMode,
        });
        preferences.updateMode = request.updateMode;
    }
    const bibliography = part.bibliography;
    const lockedNow = bibliography?.locked ?? false;
    const lockedNext = request.bibliographyLocked ?? lockedNow;
    if (request.bibliographyLocked !== undefined && request.bibliographyLocked !== lockedNow) {
        changes.push({ field: "bibliography-locked", from: String(lockedNow), to: String(lockedNext) });
    }
    if (changes.length === 0) {
        return refuse("no-change-requested", "the request asks for nothing this document does not already record");
    }
    const updated = {
        ...part,
        generator: input.generator,
        preferences,
        ...(bibliography === undefined && request.bibliographyLocked === undefined
            ? {}
            : { bibliography: { present: bibliography?.present ?? false, ...bibliography, locked: lockedNext } }),
    };
    try {
        return {
            ok: true,
            plan: Object.freeze({
                part: updated,
                customXml: wordCustomXml(updated),
                changes: Object.freeze(changes),
                notices: Object.freeze([]),
            }),
        };
    }
    catch (cause) {
        return refuse("invalid-document-part", cause instanceof Error ? cause.message : "the updated document part is invalid");
    }
}
/**
 * Plan a citation-style or locale change over the whole document.
 *
 * The order of the checks is the point. Substitution and class changes are decided **before**
 * anything is rendered into a plan, because both produce a document that looks finished and cites
 * in a style nobody chose. Only then is the render accepted, and only if it was made with the style
 * being switched to — a plan built from the old style's output would write the old citations back
 * under a payload claiming the new style, which is worse than not changing the style at all.
 */
export function prepareStyleChange(input) {
    const { part, request, context } = input;
    // The render context wins over the request for version and hash: the request is what a person
    // picked from a list, and the context is what the processor actually read (§9.3).
    const targetStyle = {
        ...(request.style ?? part.style),
        ...(context.styleUpdated === undefined ? {} : { version: context.styleUpdated }),
        ...(context.styleHash === undefined ? {} : { hash: context.styleHash }),
    };
    const targetLocale = request.locale ?? part.locale;
    const styleChanged = targetStyle.id !== part.style.id;
    const localeChanged = targetLocale !== part.locale;
    if (!styleChanged && !localeChanged) {
        return refuse("no-change-requested", "neither the style nor the locale differs from the document's own");
    }
    // §9.3/invariant 3: the registry may fall back, but this document may not inherit the fallback
    // as though the user had chosen it.
    if (context.requestedStyleId !== targetStyle.id) {
        return refuse("render-style-mismatch", `the render was requested for style ${context.requestedStyleId}, not ${targetStyle.id}`);
    }
    if (context.effectiveStyleId !== context.requestedStyleId) {
        return refuse("style-substituted", `style ${context.requestedStyleId} resolved to ${context.effectiveStyleId}; choose a style that is installed`);
    }
    if (context.requestedLocale !== targetLocale) {
        return refuse("render-style-mismatch", `the render was requested for locale ${context.requestedLocale}, not ${targetLocale}`);
    }
    const notices = [];
    if (context.effectiveLocale !== context.requestedLocale) {
        if (localeLanguage(context.effectiveLocale) !== localeLanguage(context.requestedLocale)) {
            return refuse("locale-substituted", `locale ${context.requestedLocale} resolved to ${context.effectiveLocale}, a different language`);
        }
        notices.push({
            code: "locale-region-substituted",
            message: `locale ${context.requestedLocale} is not installed; ${context.effectiveLocale} was used`,
        });
    }
    // E08-03: the style decides note-ness, and `preferences.citationFormat` is only what the last
    // style change left behind. In-text ⇄ note is not a re-render — it moves citations into or out of
    // Word footnotes, which is §9.5 structural work (E10-09), so it is refused rather than attempted.
    const currentClass = input.currentClassName ?? part.preferences.citationFormat;
    if (currentClass === undefined && part.clusters.length > 0) {
        return refuse("unknown-citation-class", "this document records no citation format, so a change to or from note style cannot be ruled out");
    }
    const relocation = input.noteRelocation;
    const classChanged = currentClass !== undefined && currentClass !== context.className;
    if (classChanged && relocation === undefined) {
        return refuse("citation-class-change-unsupported", `changing from ${currentClass} to ${context.className} citations moves them into or out of Word notes; supply a plan from planNoteRelocation (E10-09.2)`);
    }
    if (relocation !== undefined) {
        if (!classChanged) {
            return refuse("note-relocation-mismatch", "a note relocation was supplied, but this change does not alter the citation class");
        }
        const expected = context.className === "note" ? "in-text-to-note" : "note-to-in-text";
        if (relocation.direction !== expected) {
            return refuse("note-relocation-mismatch", `the relocation moves citations ${relocation.direction}, but the style change is ${expected}`);
        }
        const known = new Set(part.clusters.map((cluster) => cluster.clusterId));
        for (const move of relocation.moves) {
            if (!known.has(move.clusterId)) {
                return refuse("note-relocation-mismatch", `the relocation moves cluster ${move.clusterId}, which this document does not contain`);
            }
        }
    }
    const hasBibliography = part.bibliography?.present === true && part.bibliography.locked !== true;
    if (hasBibliography && (input.renderedBibliography?.layout.present === false ||
        input.scopedBibliographies?.some((entry) => entry.rendered.layout.present === false) === true)) {
        return refuse("style-has-no-bibliography", `style ${targetStyle.id} declares no bibliography; this document has one`);
    }
    // The render must be the new style's output, not the old style's re-labelled.
    if (input.citationPlan.part.style.id !== targetStyle.id || input.citationPlan.part.locale !== targetLocale) {
        return refuse("render-style-mismatch", "the citation plan was not produced with the style and locale being switched to");
    }
    for (const cluster of input.citationPlan.part.clusters) {
        if (cluster.renderedWith !== undefined && cluster.renderedWith.id !== targetStyle.id) {
            return refuse("render-style-mismatch", `cluster ${cluster.clusterId} was rendered with style ${cluster.renderedWith.id}`);
        }
    }
    const preferences = {
        ...part.preferences,
        citationFormat: context.className,
        ...(request.updateMode === undefined ? {} : { updateMode: request.updateMode }),
    };
    // A relocation clears note numbers on the **render**, not on the before-state: the payload the
    // plan writes and the control tags derived from it then describe one document, while the live
    // controls are still matched against the fingerprints Word actually holds. Clearing the
    // before-state instead would make every control look like it had been edited by hand.
    //
    // Every cluster is cleared, not only the ones with a stored number: a note-style render assigns
    // a number to each cluster, and before a move that has not happened, all of them are guesses.
    const before = {
        ...part,
        style: targetStyle,
        locale: targetLocale,
        preferences,
    };
    const rendered = relocation === undefined
        ? input.citationPlan
        : {
            ...input.citationPlan,
            part: clearNoteNumbers(input.citationPlan.part, input.citationPlan.part.clusters.map((cluster) => cluster.clusterId)),
        };
    const refreshed = prepareDocumentRefresh({
        part: before,
        controls: input.controls,
        citationPlan: rendered,
        // A style change is always somebody pressing a button; `on-demand` does not block it.
        trigger: "explicit",
        generator: input.generator,
        // §9.4: the *new* style's layout is the one the document should now carry, and this is the
        // explicit request E10-06 kept `reapplyLayout` for.
        reapplyBibliographyLayout: true,
        ...(input.acceptManualReplacementIds === undefined ? {} : { acceptManualReplacementIds: input.acceptManualReplacementIds }),
        ...(input.renderedBibliography === undefined ? {} : { renderedBibliography: input.renderedBibliography }),
        ...(input.documentBibliographyEntryTexts === undefined ? {} : { documentBibliographyEntryTexts: input.documentBibliographyEntryTexts }),
        ...(input.acceptBibliographyManualReplacement === undefined ? {} : { acceptBibliographyManualReplacement: input.acceptBibliographyManualReplacement }),
        ...(input.scopedBibliographies === undefined ? {} : { scopedBibliographies: input.scopedBibliographies }),
        ...(input.documentBibliographies === undefined ? {} : { documentBibliographies: input.documentBibliographies }),
        ...(input.hangingIndentTwips === undefined ? {} : { hangingIndentTwips: input.hangingIndentTwips }),
    });
    if (!refreshed.ok) {
        return refuse("refresh-refused", `the re-render refused: ${refreshed.refusal.code} — ${refreshed.refusal.message}`);
    }
    // The payload states the style, whatever the caller's renderer recorded on the part it handed in.
    const part_ = {
        ...refreshed.plan.part,
        style: targetStyle,
        locale: targetLocale,
        preferences: { ...refreshed.plan.part.preferences, ...preferences },
    };
    const changes = [];
    if (styleChanged)
        changes.push({ field: "style", from: part.style.id, to: targetStyle.id });
    if (localeChanged)
        changes.push({ field: "locale", from: part.locale, to: targetLocale });
    if (request.updateMode !== undefined && request.updateMode !== part.preferences.updateMode) {
        changes.push({
            field: "update-mode",
            ...(part.preferences.updateMode === undefined ? {} : { from: part.preferences.updateMode }),
            to: request.updateMode,
        });
    }
    if (part.preferences.citationFormat !== context.className) {
        changes.push({
            field: "citation-format",
            ...(part.preferences.citationFormat === undefined ? {} : { from: part.preferences.citationFormat }),
            to: context.className,
        });
        notices.push({
            code: "citation-format-recorded",
            message: `the document now records ${context.className} citations, as the style declares`,
        });
    }
    if (targetStyle.version !== undefined || targetStyle.hash !== undefined) {
        notices.push({
            code: "style-version-recorded",
            message: "the style's declared version was recorded with the document (§9.3)",
        });
    }
    if (refreshed.plan.bibliographies.length > 0) {
        notices.push({
            code: "bibliography-layout-reapplied",
            message: "the bibliography now carries the new style's paragraph layout, replacing the document's",
        });
    }
    if (relocation !== undefined) {
        notices.push({
            code: "citation-class-relocated",
            message: `${relocation.moves.length} citation(s) move ${relocation.direction}${relocation.noteStory === null ? "" : ` into ${relocation.noteStory}s`}`,
        });
        notices.push({
            code: "note-render-provisional",
            message: "this citation text was rendered before Word numbered the notes; refresh the document after applying the moves",
        });
    }
    for (const clusterId of refreshed.plan.heldManualClusterIds) {
        notices.push({
            code: "manual-edit-retains-old-style",
            clusterId,
            message: `cluster ${clusterId} keeps its manual text, which was written under style ${part.style.id}`,
        });
    }
    try {
        return {
            ok: true,
            plan: Object.freeze({
                refresh: Object.freeze({ ...refreshed.plan, part: part_, customXml: wordCustomXml(part_) }),
                changes: Object.freeze(changes),
                notices: Object.freeze(notices),
                ...(relocation === undefined ? {} : { noteRelocation: relocation }),
            }),
        };
    }
    catch (cause) {
        return refuse("invalid-document-part", cause instanceof Error ? cause.message : "the restyled document part is invalid");
    }
}
/**
 * Payload first, stale payloads last (ADR-0107) — with no middle stage, because a preferences
 * change writes no content control at all. Failing before the new part lands leaves the document
 * exactly as it was, which is why there is nothing here to compensate.
 */
export async function commitDocumentPreferences(port, plan, stalePartIds = []) {
    let added;
    try {
        added = await port.addCustomXml(plan.customXml);
    }
    catch {
        return { status: "failed", stage: "custom-xml" };
    }
    const failed = [];
    for (const id of stalePartIds) {
        if (id === added.id)
            continue;
        try {
            await port.deleteCustomXml(id);
        }
        catch {
            failed.push(id);
        }
    }
    return failed.length === 0
        ? { status: "written", customXmlPartId: added.id }
        : { status: "written-with-cleanup-required", customXmlPartId: added.id, stalePartIds: failed };
}
