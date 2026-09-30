/**
 * Insert and regenerate the Word bibliography (SPEC §9.4, task E10-06; ADR-0007, ADR-0109).
 *
 * §9.4 is eight sentences and seven of them are about *not* destroying something: keep the
 * author's paragraph styles, keep pinned records, keep an explicitly locked bibliography, warn
 * before overwriting manual changes, change nothing around it. So this module is written as a
 * planner that says no, in the same shape as insertion and editing (ADR-0107): a pure function
 * produces the exact control and payload bytes, and a separate commit writes the payload first
 * and compensates. Nothing here formats a citation — every visible string is processor output
 * that arrived over the bridge (invariant 4).
 */
import { BIBLIOGRAPHY_TAG, DOCUMENT_SCOPE_ID, fingerprint, parseBibliographyTag, wordBibliographyControl, wordCustomXml, } from "@refmgr/doc-schema";
/** The word-processor unit: 1/20 pt, 1440 to the inch. `w:ind` and `w:spacing` both take it. */
export const TWIPS_PER_INCH = 1440;
/** `w:spacing/@w:line` for single spacing with `w:lineRule="auto"`, at any font size. */
export const SINGLE_LINE_TWIPS = 240;
/** Half an inch: the hanging indent every major style guide specifies, and Word's own default. */
export const DEFAULT_HANGING_INDENT_TWIPS = 720;
/** Word's built-in paragraph style for a bibliography, so the author can restyle it themselves. */
export const BIBLIOGRAPHY_STYLE_NAME = "Bibliography";
function refuse(code, message) {
    return { ok: false, refusal: { code, message } };
}
/**
 * Which records the bibliography must list, and which it drops.
 *
 * Cited records come first, in document order, because that is the order a reader would expect a
 * numeric style to number them in; the processor may then sort them however the style says, and
 * this order is never imposed on its output. Pinned and explicitly uncited records are added
 * after (§9.4 "unless manually pinned", "support uncited entries when requested").
 *
 * `removed` is the §9.4 removal, reported rather than performed: it lists records that still have
 * an offline snapshot but are no longer required. **Their snapshots are kept.** A snapshot is a
 * document's only offline copy of a record, and dropping one because a citation was deleted would
 * make an earlier version of the same document unreadable offline.
 */
export function bibliographyItemIds(part, scopeId = DOCUMENT_SCOPE_ID) {
    const ordered = [...part.clusters].sort((left, right) => (left.order ?? 0) - (right.order ?? 0));
    const cited = [];
    for (const cluster of ordered) {
        for (const item of cluster.items) {
            if (!cited.includes(item.itemId))
                cited.push(item.itemId);
        }
    }
    const pinned = [...new Set(part.bibliography?.pinnedItemIds ?? [])];
    const uncited = [...new Set(part.bibliography?.uncitedItemIds ?? [])];
    const documentRequired = [...cited];
    for (const id of [...pinned, ...uncited]) {
        if (!documentRequired.includes(id))
            documentRequired.push(id);
    }
    // A scope with mode "items" lists exactly what it declares, in the document order of the
    // records it names. Nothing is added to it here: a chapter bibliography that quietly grew a
    // record because it was cited somewhere else is R-072 wearing a smaller hat.
    const scope = findScope(part, scopeId);
    const required = scope !== undefined && scope.mode === "items"
        ? [...new Set(scope.itemIds ?? [])].sort((left, right) => orderIndex(documentRequired, left) - orderIndex(documentRequired, right))
        : documentRequired;
    const removed = Object.keys(part.snapshots)
        .filter((id) => !required.includes(id))
        .sort();
    return { required, cited, pinned, uncited, removed };
}
function orderIndex(order, id) {
    const index = order.indexOf(id);
    return index === -1 ? Number.MAX_SAFE_INTEGER : index;
}
function findScope(part, scopeId) {
    return (part.bibliography?.scopes ?? []).find((scope) => scope.scopeId === scopeId);
}
/**
 * The one place CSL's unit-free layout becomes a Word measurement (§9.4).
 *
 * CSL says "hanging-indent: true" and nothing about how far; Word needs a number. Half an inch is
 * what APA, AMA, Chicago and MLA all specify and what Word's own Bibliography style uses, so it
 * is the default rather than a guess dressed up as one — and it is a parameter, because a user
 * who changes it in Word must be able to keep that change (see `preserveExistingFormat`).
 */
export function bibliographyParagraphFormat(layout, options = {}) {
    const hanging = layout.hangingIndent ? (options.hangingIndentTwips ?? DEFAULT_HANGING_INDENT_TWIPS) : 0;
    return {
        styleName: options.styleName ?? BIBLIOGRAPHY_STYLE_NAME,
        leftIndentTwips: hanging,
        hangingIndentTwips: hanging,
        lineTwips: SINGLE_LINE_TWIPS * (layout.lineSpacing ?? 1),
        afterTwips: SINGLE_LINE_TWIPS * (layout.entrySpacing ?? 1),
    };
}
/**
 * The hash recorded in `bibliography.renderedHash`, and the one recomputed from the document.
 *
 * It covers the entry texts and nothing else, because that is exactly what §9.4 means by "manual
 * bibliography changes": re-running the same style over the same records must not read as an edit,
 * and a reader who retyped an author's initials must.
 */
export function bibliographyRenderedHash(entryTexts) {
    return fingerprint(entryTexts);
}
/**
 * Plan the bibliography without touching Word.
 *
 * The refusals are the feature. Regenerating a bibliography is the one operation in this product
 * that rewrites a block of the author's manuscript wholesale, and every way it can go wrong is
 * silent: a dropped record looks like a record that was never cited, an overwritten manual fix
 * looks like the fix was never made, and a re-applied indent looks like Word did it.
 */
export function prepareBibliography(input) {
    const { part, rendered } = input;
    const state = part.bibliography;
    if (state?.locked === true) {
        return refuse("bibliography-locked", "the document bibliography is locked; unlock it before regenerating");
    }
    if (input.trigger === "automatic" && part.preferences.updateMode === "on-demand") {
        return refuse("update-on-demand-only", "this document updates its bibliography on demand only");
    }
    if (rendered.context.requestedStyleId !== part.style.id || rendered.context.requestedLocale !== part.locale) {
        return refuse("render-mismatch", "the rendered bibliography does not match the document style and locale");
    }
    if (!rendered.layout.present) {
        return refuse("style-has-no-bibliography", `style ${rendered.context.effectiveStyleId} declares no bibliography; this document cannot have one`);
    }
    // §9.4 / R-072: which bibliography is this? Everything below writes one control, so the
    // question has to be settled before a single entry is chosen — writing a plan made for "the"
    // bibliography into a thesis that has six is the defect itself.
    const scopes = part.bibliography?.scopes ?? [];
    const scopeId = input.scopeId ?? (scopes.length === 1 ? scopes[0].scopeId : DOCUMENT_SCOPE_ID);
    if (input.scopeId === undefined && scopes.length > 1) {
        return refuse("scope-required", `this document declares ${scopes.length} bibliographies; name which one is being regenerated`);
    }
    if (scopes.length > 0 && findScope(part, scopeId) === undefined) {
        return refuse("unknown-bibliography-scope", `this document declares no bibliography scope ${scopeId}`);
    }
    const liveTags = input.documentBibliographyTags;
    if (liveTags !== undefined && liveTags.length > 1) {
        const parsed = liveTags.map((tag) => parseBibliographyTag(tag));
        const unscoped = parsed.filter((tag) => tag !== null && tag.shortFingerprint === null).length;
        const unrecognised = parsed.filter((tag) => tag === null).length;
        if (unscoped > 0 || unrecognised > 0 || parsed.length > scopes.length) {
            return refuse("multiple-unscoped-bibliographies", `the document holds ${liveTags.length} bibliography controls and declares ${scopes.length} scopes; regenerating would write one bibliography into all of them`);
        }
    }
    const { required, removed } = bibliographyItemIds(part, scopeId);
    for (const itemId of required) {
        if (part.snapshots[itemId] === undefined) {
            return refuse("missing-snapshot", `record ${itemId} has no offline snapshot in this document`);
        }
    }
    const entries = rendered.entries.map((entry) => ({
        itemIds: [...entry.itemIds],
        text: entry.text,
    }));
    const rendered_ids = new Set(entries.flatMap((entry) => entry.itemIds));
    for (const itemId of required) {
        if (!rendered_ids.has(itemId)) {
            return refuse("missing-entry", `record ${itemId} must be listed but the render has no entry for it`);
        }
    }
    for (const itemId of rendered_ids) {
        if (!required.includes(itemId)) {
            return refuse("unexpected-entry", `the render lists record ${itemId}, which this document does not cite or pin`);
        }
    }
    if (required.length > 0 && entries.length === 0) {
        return refuse("empty-render", "the CSL processor returned no bibliography entries for the records requested");
    }
    const notices = [];
    // §9.4 "warn before overwriting manual bibliography changes". Three states, not two: the
    // document matches what we wrote, the document differs from it, or **nothing records what we
    // wrote** — an older document, or one whose payload was rebuilt. The third is not "unchanged":
    // nobody can say whether it was edited, and quietly overwriting it would be the same failure as
    // overwriting a known edit, minus the evidence.
    // A scoped control's provenance is its own scope's hash, not the document's: Chapter 2's
    // hash says nothing about whether somebody retyped Chapter 1.
    const scope = findScope(part, scopeId);
    const documentTexts = input.documentEntryTexts;
    const knownHash = scope !== undefined ? scope.renderedHash : state?.renderedHash;
    if (documentTexts !== undefined) {
        const documentHash = bibliographyRenderedHash([...documentTexts]);
        if (knownHash === undefined) {
            notices.push({
                code: "unattributed-bibliography",
                message: "the document has a bibliography but records no hash of one we wrote",
            });
            if (input.acceptManualReplacement !== true) {
                return refuse("manual-edit-unresolved", "the document bibliography has no recorded provenance; replacing it needs an explicit decision");
            }
        }
        else if (knownHash !== documentHash && input.acceptManualReplacement !== true) {
            return refuse("manual-edit-unresolved", "the visible bibliography differs from the one last written; keep or replace it explicitly");
        }
    }
    // R-072's honest cost: a scope lists what it declares, so a record cited in the manuscript
    // that no scope lists appears in no bibliography at all. That must be said out loud — the
    // alternative is appending it to whichever bibliography happened to be regenerated, which is
    // how a chapter acquires another chapter's references in the first place.
    if (scopes.length > 0) {
        const listed = new Set(scopes.flatMap((entry) => (entry.mode === "items" ? (entry.itemIds ?? []) : bibliographyItemIds(part).required)));
        const unlisted = bibliographyItemIds(part).cited.filter((itemId) => !listed.has(itemId));
        if (unlisted.length > 0) {
            notices.push({
                code: "unlisted-items",
                itemIds: unlisted,
                message: `${unlisted.length} record(s) are cited in this document but listed by no bibliography; add them to a scope or they appear in none`,
            });
        }
    }
    if (removed.length > 0) {
        notices.push({
            code: "entries-removed",
            itemIds: removed,
            message: `${removed.length} record(s) are no longer cited or pinned and will not be listed; their snapshots are kept`,
        });
    }
    if (rendered.layout.unrecognised.length > 0) {
        notices.push({
            code: "layout-unrecognised",
            message: `the style declares layout this build does not apply: ${rendered.layout.unrecognised.join(", ")}`,
        });
    }
    if (rendered.layout.secondFieldAlign !== undefined) {
        notices.push({
            code: "second-field-align-unsupported",
            message: `the style asks for second-field-align="${rendered.layout.secondFieldAlign}", which this build does not express in Word`,
        });
    }
    // §9.4 "preserve Word paragraph styles". Formatting is written when the bibliography is being
    // created, or when the caller explicitly asks for the style's layout again (E10-08's style
    // change). Otherwise the plan carries no formatting at all, which is the only way to preserve
    // formatting that a user may have changed in Word since we wrote it.
    const existing = state?.present === true && documentTexts !== undefined;
    const applyLayout = !existing || input.reapplyLayout === true;
    const paragraphFormat = applyLayout
        ? bibliographyParagraphFormat(rendered.layout, {
            ...(input.hangingIndentTwips === undefined ? {} : { hangingIndentTwips: input.hangingIndentTwips }),
        })
        : null;
    if (!applyLayout) {
        notices.push({
            code: "layout-preserved",
            message: "the bibliography keeps the paragraph formatting already in the document",
        });
    }
    const bibliography = {
        ...state,
        present: true,
        ...(state?.pinnedItemIds === undefined ? {} : { pinnedItemIds: [...state.pinnedItemIds] }),
        ...(state?.uncitedItemIds === undefined ? {} : { uncitedItemIds: [...state.uncitedItemIds] }),
        // The document-level hash means "the whole document's bibliography", so a chapter refresh
        // leaves it exactly as it was rather than claiming the document was rewritten.
        ...(scope === undefined || scope.scopeId === DOCUMENT_SCOPE_ID
            ? { renderedHash: bibliographyRenderedHash(entries.map((entry) => entry.text)) }
            : {}),
        ...(scopes.length === 0
            ? {}
            : {
                scopes: scopes.map((entry) => entry.scopeId === scopeId
                    ? { ...entry, renderedHash: bibliographyRenderedHash(entries.map((written) => written.text)) }
                    : entry),
            }),
        extensions: {
            ...state?.extensions,
            "refmgr:bibliography": {
                entryItemIds: entries.map((entry) => [...entry.itemIds]),
                effectiveStyleId: rendered.context.effectiveStyleId,
                effectiveLocale: rendered.context.effectiveLocale,
                ...(paragraphFormat === null ? {} : { paragraphFormat: { ...paragraphFormat } }),
            },
        },
    };
    const updated = { ...part, generator: input.generator, bibliography };
    try {
        const control = wordBibliographyControl(entries, paragraphFormat, scope);
        return {
            ok: true,
            plan: Object.freeze({
                scopeId,
                part: updated,
                customXml: wordCustomXml(updated),
                control,
                requiredItemIds: Object.freeze([...required]),
                removedItemIds: Object.freeze([...removed]),
                warnings: Object.freeze([...rendered.warnings]),
                notices: Object.freeze(notices),
            }),
        };
    }
    catch (cause) {
        return refuse("invalid-document-part", cause instanceof Error ? cause.message : "the bibliography part is invalid");
    }
}
/**
 * Set §9.4's per-document switches: pinning, uncited entries, and locked/manual mode.
 *
 * Separate from regeneration on purpose. Pinning a record is a decision about the document, and
 * making it a side effect of a refresh would mean the only way to record it was to rewrite the
 * manuscript. Pinning an unknown record is refused rather than stored, because a pin that names
 * nothing quietly stops being a pin.
 */
export function setBibliographyOptions(part, options) {
    for (const itemId of [...(options.pinnedItemIds ?? []), ...(options.uncitedItemIds ?? [])]) {
        if (part.snapshots[itemId] === undefined) {
            return {
                ok: false,
                refusal: { code: "missing-snapshot", message: `record ${itemId} has no offline snapshot in this document` },
            };
        }
    }
    const state = part.bibliography;
    const bibliography = {
        present: state?.present ?? false,
        ...state,
        ...(options.pinnedItemIds === undefined ? {} : { pinnedItemIds: [...new Set(options.pinnedItemIds)] }),
        ...(options.uncitedItemIds === undefined ? {} : { uncitedItemIds: [...new Set(options.uncitedItemIds)] }),
        ...(options.locked === undefined ? {} : { locked: options.locked }),
    };
    return {
        ok: true,
        part: {
            ...part,
            bibliography,
            ...(options.updateMode === undefined
                ? {}
                : { preferences: { ...part.preferences, updateMode: options.updateMode } }),
        },
    };
}
/**
 * Declare which bibliographies this document has, and what each one lists (§9.4, R-072).
 *
 * Separate from regeneration for the same reason pinning is: deciding that Chapter 2 lists
 * these eleven records is a statement about the manuscript, and making it a side effect of a
 * refresh would mean the only way to record it was to rewrite the bibliography.
 *
 * The refusals are the point. A duplicate scope id makes two controls indistinguishable, which
 * is the state R-072 describes; a scope naming a record with no snapshot is a bibliography that
 * cannot be rendered on a machine with no library (§9.3); and re-declaring the reserved
 * `document` scope as an item list would quietly change what every existing bare `refmgr-bib`
 * control in the document means.
 */
export function setBibliographyScopes(part, scopes) {
    const seen = new Set();
    for (const scope of scopes) {
        if (scope.scopeId.length === 0) {
            return { ok: false, refusal: { code: "unknown-bibliography-scope", message: "a bibliography scope must have an id" } };
        }
        if (seen.has(scope.scopeId)) {
            return {
                ok: false,
                refusal: {
                    code: "multiple-unscoped-bibliographies",
                    message: `bibliography scope ${scope.scopeId} is declared twice; two controls claiming one scope cannot be told apart`,
                },
            };
        }
        seen.add(scope.scopeId);
        if (scope.scopeId === DOCUMENT_SCOPE_ID && scope.mode !== "document") {
            return {
                ok: false,
                refusal: {
                    code: "unknown-bibliography-scope",
                    message: `the reserved scope ${DOCUMENT_SCOPE_ID} always covers the whole document; give a chapter bibliography its own id`,
                },
            };
        }
        if (scope.mode === "items" && (scope.itemIds ?? []).length === 0) {
            return {
                ok: false,
                refusal: { code: "empty-render", message: `bibliography scope ${scope.scopeId} lists no records` },
            };
        }
        for (const itemId of scope.itemIds ?? []) {
            if (part.snapshots[itemId] === undefined) {
                return {
                    ok: false,
                    refusal: { code: "missing-snapshot", message: `record ${itemId} has no offline snapshot in this document` },
                };
            }
        }
    }
    const state = part.bibliography;
    return {
        ok: true,
        part: {
            ...part,
            bibliography: {
                present: state?.present ?? false,
                ...state,
                scopes: scopes.map((scope) => ({ ...scope, ...(scope.itemIds === undefined ? {} : { itemIds: [...scope.itemIds] }) })),
            },
        },
    };
}
/**
 * Payload first, control second, old parts last — the same compensated order as insert and edit
 * (ADR-0107), for the same reason: Office exposes custom XML and content controls through
 * different API surfaces, so no transaction spans both.
 */
export async function commitBibliography(port, plan, stalePartIds = []) {
    let added;
    try {
        added = await port.addCustomXml(plan.customXml);
    }
    catch {
        return { status: "failed", stage: "custom-xml", cleanupRequiredPartIds: [] };
    }
    try {
        await port.writeBibliographyControl(plan.control);
    }
    catch {
        try {
            await port.deleteCustomXml(added.id);
            return { status: "failed", stage: "content-control", cleanupRequiredPartIds: [] };
        }
        catch {
            return { status: "failed", stage: "content-control", cleanupRequiredPartIds: [added.id] };
        }
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
export { BIBLIOGRAPHY_TAG };
