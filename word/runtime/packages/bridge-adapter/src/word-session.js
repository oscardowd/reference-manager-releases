/**
 * The live Word session: every operation the task pane asks of a citation runtime (ADR-0241).
 *
 * Transport has already authenticated the caller. This module depends only on the bridge's port
 * interfaces and a style registry — never on a database — because it runs in two places
 * (ADR-0242): in the desktop over the user's library (`word-session-library.ts`), and inside the
 * task pane itself over the document's own §9.3 snapshots (`document-ports.ts`) when no desktop is
 * reachable. One handler, two compositions: a document looks the same whichever mode wrote it.
 */
import { createPortRouter, BRIDGE_ROUTES, BridgeRequestError } from "@refmgr/bridge";
import { CitationClusterRenderer, CitationDocumentEditor } from "@refmgr/citation";
import { DOCUMENT_SCOPE_ID, parseClusterTag, readDocumentPart, wordCitationControl } from "@refmgr/doc-schema";
import { bibliographyItemIds, prepareBibliography } from "../../word-addin/src/bibliography.js";
import { citationRequest } from "../../word-addin/src/draft.js";
import { openCitationForEdit, payloadText, prepareCitationEdit, } from "../../word-addin/src/editing.js";
import { prepareInsertion } from "../../word-addin/src/insertion.js";
import { prepareStyleChange } from "../../word-addin/src/preferences.js";
import { bibliographyRefreshRequests, prepareDocumentRefresh, } from "../../word-addin/src/refresh.js";
const READ_PATHS = ["/recent", "/search", "/collections", "/styles", "/citation/preview"];
function record(value) {
    if (typeof value !== "object" || value === null || Array.isArray(value))
        throw new BridgeRequestError(400, "invalid-request", "Invalid Word request.");
    return value;
}
function refuse(message) { throw new BridgeRequestError(400, "word-refused", message); }
/** Never select a payload by array order, and never discard an unreadable payload. */
export function readWordState(raw) {
    const evidence = record(raw);
    const candidates = evidence["customXmlParts"];
    const controls = evidence["controls"];
    if (!Array.isArray(candidates) || !Array.isArray(controls))
        refuse("The document could not be read. Reopen the pane and try again.");
    if (candidates.length > 1)
        refuse("This document has conflicting citation data. Check the document before inserting.");
    if (candidates.length === 0) {
        if (controls.length !== 0)
            refuse("Citation controls have no matching document data. Check the document before inserting.");
        return { part: null, stalePartIds: [] };
    }
    const candidate = record(candidates[0]);
    if (typeof candidate["xml"] !== "string" || typeof candidate["id"] !== "string")
        refuse("The document citation data is unreadable.");
    const json = payloadText(candidate["xml"]);
    const read = json === null ? null : readDocumentPart(json);
    if (read?.status !== "ok")
        refuse("The document citation data is unreadable or from an unsupported version.");
    const expected = new Map(read.part.clusters.map((cluster) => {
        const control = wordCitationControl(cluster);
        return [control.tag, control.visibleText];
    }));
    if (controls.length !== expected.size)
        refuse("Some citations were deleted or changed. Choose Update document on the Document tab, then insert.");
    for (const value of controls) {
        const control = record(value);
        const tag = String(control["tag"]);
        if (!expected.has(tag) || expected.get(tag) !== control["visibleText"])
            refuse("A citation was changed by hand. Choose Update document on the Document tab to review it, then insert.");
        expected.delete(tag);
    }
    return { part: read.part, stalePartIds: [candidate["id"]] };
}
/** The generator recorded on every payload this runtime writes. */
const GENERATOR = "refmgr-word/0.1";
/**
 * Fixed sentences for the planners' refusal codes. The planners' own messages name cluster ids and
 * structural states for logs and tests; a writer in Word is told what happened and what to do.
 * Never a reference title or citation text (§23).
 */
const DOCUMENT_REFUSAL_COPY = Object.freeze({
    "bibliography-locked": "This document's bibliography is locked. Unlock it before updating it.",
    "style-has-no-bibliography": "This citation style has no bibliography.",
    "missing-entry": "A cited reference could not be listed in the bibliography. Check that it is still in your library.",
    "missing-snapshot": "A cited reference has no saved copy in this document. Re-insert that citation, then try again.",
    "empty-render": "The bibliography came back empty. Check the citation style.",
    "multiple-unscoped-bibliographies": "This document has more than one bibliography. Keep one, then try again.",
    "scope-required": "This document has more than one bibliography. Keep one, then try again.",
    "duplicate-control": "A citation appears twice, probably from copy and paste. Delete the copy, then try again.",
    "invalid-control": "A citation in the document is not one Reference Manager wrote. Check the document, then try again.",
    "control-payload-mismatch": "A citation no longer matches this document's citation data. Check the document, then try again.",
    "citation-class-change-unsupported": "Switching between in-text and footnote styles is not supported yet.",
    "unknown-citation-class": "This document's citation type is unknown, so its style cannot be changed safely.",
    "style-substituted": "That citation style is not fully installed in Reference Manager.",
    "locale-substituted": "This document's language is not installed in Reference Manager.",
    "no-change-requested": "The document already uses that style.",
});
function documentRefusal(code, fallback) {
    refuse(DOCUMENT_REFUSAL_COPY[code] ?? fallback);
}
/** `prepareStyleChange` carries a refresh refusal as "the re-render refused: <code> — …". */
export function refreshCode(message) {
    return /^the re-render refused: ([a-z-]+)/u.exec(message)?.[1] ?? "";
}
/**
 * A refresh refused inside one of its bibliographies: "bibliography scope <id> refresh refused:
 * <code>", or a statement about the live controls. The bibliography's own code has copy; the
 * control statements mean the document's bibliography is not the one its data describes —
 * deleted or duplicated by hand, as a rule.
 */
function bibliographyRefusal(message) {
    const code = /refresh refused: ([a-z-]+)$/u.exec(message)?.[1];
    if (code === "manual-edit-unresolved") {
        refuse("The bibliography was edited by hand. Choose Update again and confirm replacing it.");
    }
    if (code !== undefined)
        documentRefusal(code, "The bibliography could not be updated. Check the document, then try again.");
    refuse("The document's bibliography is missing or duplicated. Delete any extra copy, or insert the bibliography again.");
}
function stringList(value, message) {
    if (!Array.isArray(value) || value.some((entry) => typeof entry !== "string"))
        refuse(message);
    return value;
}
/**
 * The whole live document, for an operation that rewrites it (refresh, style, bibliography).
 *
 * Unlike {@link readWordState} this does not require every control to show the text the payload
 * last wrote: a hand-edited citation is exactly what refresh exists to find, hold and report. It
 * still refuses what no planner can resolve — two payloads, an unreadable one, controls with none.
 */
export function readLiveDocument(raw) {
    const evidence = record(raw);
    const candidates = evidence["customXmlParts"];
    const controls = evidence["controls"];
    if (!Array.isArray(candidates) || !Array.isArray(controls))
        refuse("The document could not be read. Reopen the pane and try again.");
    if (candidates.length > 1)
        refuse("This document has conflicting citation data. Check the document before updating it.");
    if (candidates.length === 0) {
        if (controls.length !== 0)
            refuse("Citation controls have no matching document data. Check the document before updating it.");
        return null;
    }
    const candidate = record(candidates[0]);
    if (typeof candidate["xml"] !== "string" || typeof candidate["id"] !== "string")
        refuse("The document citation data is unreadable.");
    const json = payloadText(candidate["xml"]);
    const read = json === null ? null : readDocumentPart(json);
    if (read?.status !== "ok")
        refuse("The document citation data is unreadable or from an unsupported version.");
    const live = controls.map((value) => {
        const control = record(value);
        if (typeof control["tag"] !== "string" || typeof control["visibleText"] !== "string")
            refuse("The document could not be read. Reopen the pane and try again.");
        return { tag: control["tag"], visibleText: control["visibleText"] };
    });
    const entryTexts = evidence["documentEntryTexts"];
    const bibliographies = evidence["documentBibliographies"];
    return {
        part: read.part,
        stalePartIds: [candidate["id"]],
        controls: live,
        ...(entryTexts === undefined || entryTexts === null
            ? {}
            : { documentEntryTexts: stringList(entryTexts, "The bibliography could not be read. Reopen the pane and try again.") }),
        ...(bibliographies === undefined || bibliographies === null
            ? {}
            : {
                documentBibliographies: (Array.isArray(bibliographies) ? bibliographies : refuse("The bibliography could not be read.")).map((value) => {
                    const control = record(value);
                    if (typeof control["tag"] !== "string")
                        refuse("The bibliography could not be read. Reopen the pane and try again.");
                    return { tag: control["tag"], entryTexts: stringList(control["entryTexts"], "The bibliography could not be read. Reopen the pane and try again.") };
                }),
            }),
    };
}
/**
 * Citations the payload records but Word no longer shows — deleted by the writer, as a rule.
 *
 * They are **not** dropped silently (invariant 3). Word's body walk cannot see a citation inside a
 * text box, so "missing" can also mean "unreachable"; the writer is asked, with the count, and only
 * the ids they confirmed are removed. Their snapshots are kept, as §9.3 requires.
 */
export function missingCitationIds(live) {
    const shown = new Set(live.controls.map((control) => parseClusterTag(control.tag)?.clusterId).filter((id) => id !== undefined));
    return live.part.clusters.filter((cluster) => !shown.has(cluster.clusterId)).map((cluster) => cluster.clusterId);
}
/**
 * The payload with its clusters in the order Word shows them. The payload's own `order` is where
 * each citation was when it was written; a numeric bibliography is numbered by first citation,
 * so it is asked of the document as it stands. Clusters Word does not show keep their relative
 * order after the ones it does (the caller has already asked about any such deletion).
 */
export function inDocumentOrder(part, controls) {
    const position = new Map();
    controls.forEach((control, index) => {
        const id = parseClusterTag(control.tag)?.clusterId;
        if (id !== undefined && !position.has(id))
            position.set(id, index);
    });
    const ranked = part.clusters
        .map((cluster, index) => ({ cluster, rank: position.get(cluster.clusterId) ?? controls.length + index }))
        .sort((left, right) => left.rank - right.rank);
    return { ...part, clusters: ranked.map(({ cluster }, order) => ({ ...cluster, order })) };
}
export function createWordSession(options) {
    const styles = options.styles;
    const skippedStyles = options.skippedStyleFiles ?? 0;
    const ports = options.ports;
    const newId = options.newId ?? (() => globalThis.crypto.randomUUID());
    const now = options.now ?? (() => new Date().toISOString());
    const withSource = (item) => options.sourceLibraryId === undefined ? item : { ...item, sourceLibraryId: options.sourceLibraryId };
    // The session answers only the library and citation routes, so the other ports stay absent and
    // their routes fail closed, exactly as a host without them would; the document-backup upload is
    // not a session route at all, so the port router — without it — is the whole of what it needs.
    const router = createPortRouter({ library: ports.library, citation: ports.citation });
    /**
     * The selected citation, as the pane read it out of Word.
     *
     * `visibleText` is what Word actually shows, not what the payload claims — that difference is
     * the whole of manual-edit detection (§9.3), so it is carried rather than re-derived here.
     */
    function selectedControl(raw) {
        const control = record(raw);
        if (typeof control["tag"] !== "string" || typeof control["visibleText"] !== "string") {
            refuse("No citation is selected. Click a citation in the document, then try again.");
        }
        return { tag: control["tag"], visibleText: control["visibleText"] };
    }
    /**
     * Every custom XML part the pane found, unfiltered.
     *
     * Deliberately **not** `readWordState`, which refuses more than one payload because an insertion
     * has to choose which document state it extends. Editing does not have to choose: E10-05's
     * `openCitationForEdit` already reconciles several payloads against the selected control's
     * fingerprint, reports the stale ones and refuses an ambiguous set with a reason. Re-refusing
     * here would throw that reconciliation away and tell the user less.
     */
    function customXmlCandidates(raw) {
        const evidence = record(raw);
        const candidates = evidence["customXmlParts"];
        if (!Array.isArray(candidates)) {
            refuse("The document could not be read. Reopen the pane and try again.");
        }
        return candidates.map((value) => {
            const candidate = record(value);
            if (typeof candidate["id"] !== "string" || typeof candidate["xml"] !== "string") {
                refuse("The document citation data is unreadable.");
            }
            return { id: candidate["id"], xml: candidate["xml"] };
        });
    }
    /**
     * A label for one cited reference, for the panel's "currently editing" line.
     *
     * The library is asked first because it holds the current record; a reference the library no
     * longer has is **not** an error — §9.3's snapshot is exactly what keeps an already-cited,
     * since-deleted work editable — so the entry comes back with a null reference and the pane says
     * so rather than refusing the edit.
     */
    async function citedReferences(draft) {
        return Promise.all(draft.entries.map(async (entry) => ({
            entryId: entry.entryId,
            itemId: entry.itemId,
            reference: await ports.library.getItem(entry.itemId),
        })));
    }
    /** Re-open the edit session from the evidence in the request. Stateless on purpose — see below. */
    function openSession(body) {
        const session = openCitationForEdit(selectedControl(body["control"]), customXmlCandidates(body["document"]));
        if (!session.ok)
            refuse(session.refusal.message);
        return session.session;
    }
    async function editOperation(path, body) {
        // **Re-opened from the document on every call, never cached between them.** A session held on
        // this side would go stale the moment the user edited the document in another way, and the
        // first thing the stale copy would do is write a payload built from a state that no longer
        // exists. Re-reading costs one parse and makes "the document changed underneath" a refusal
        // rather than a silent overwrite.
        const session = openSession(body);
        const styleId = session.part.style.id;
        const locale = session.part.locale;
        if (path === "/edit/open") {
            return {
                status: 200,
                body: {
                    clusterId: session.clusterId,
                    draft: session.draft,
                    manualEditDetected: session.manualEditDetected,
                    notices: session.notices,
                    styleId,
                    locale,
                    // What Word shows right now. The panel opens on the citation as it stands, not on a
                    // re-render of it — re-rendering before the user has changed anything would show them a
                    // citation that differs from their document and call it the current one.
                    visibleText: selectedControl(body["control"]).visibleText,
                    references: await citedReferences(session.draft),
                },
            };
        }
        const draft = record(body["draft"]);
        if (!Array.isArray(draft.entries) || draft.entries.length === 0 || draft.entries.length > 50) {
            refuse("Invalid citation draft.");
        }
        const built = citationRequest(draft, { styleId, locale });
        if (!built.ok)
            refuse(built.refusal.message);
        const items = await Promise.all(built.request.items.map(async ({ itemId }) => {
            const item = await ports.library.getItem(itemId);
            return item === null ? null : withSource(item);
        }));
        // **Opening such a citation works; saving it does not, and that is said plainly.**
        // `prepareCitationEdit` would manage — it falls back to the document's own snapshots, which is
        // what §9.3 keeps them for — but the *render* cannot: `/citation/preview` resolves every cite
        // through the library, so a since-deleted reference has no text to write. Letting it reach the
        // renderer would surface `no item "…" in this library`, which reads as a bug rather than as
        // the state it is. Rendering from the snapshot instead is a change to the shared citation
        // adapter and is E10-05.2, not an aside here.
        if (items.some((item) => item === null)) {
            refuse("A reference in this citation is no longer in your library, so its text cannot be re-rendered. " +
                "Restore the reference in Reference Manager, then edit the citation again.");
        }
        const rendered = await ports.citation.formatCitation(built.request);
        // The same anti-drift check insertion makes: the text the user approved is the text that gets
        // written, or nothing is written. A preview that has moved since they read it is a different
        // citation from the one they pressed Save on.
        if (body["previewText"] !== rendered.text) {
            refuse("The preview changed. Review the refreshed preview and save again.");
        }
        const prepared = prepareCitationEdit({
            session,
            draft,
            rendered,
            items: items.filter((item) => item !== null),
            generator: "refmgr-word/0.1",
            snapshotAt: now(),
            acceptManualEditReplacement: body["acceptManualEditReplacement"] === true,
        });
        if (!prepared.ok)
            refuse(prepared.refusal.message);
        return {
            status: 200,
            body: {
                plan: prepared.plan,
                stalePartIds: prepared.stalePartIds,
                expectedTag: selectedControl(body["control"]).tag,
                // A numeric style's text depends on position, which a single-cite render cannot know: the
                // pane writes the edit and then re-renders the document, as it does after an insertion.
                refreshRequired: needsDocumentRender(prepared.plan.part, rendered.warnings),
            },
        };
    }
    /**
     * Whether text written from a single-cite render would stand in this document. The processor's
     * own warning is not enough: a numeric style that omits its `citation-format` category raises
     * none, and still numbers by position. So the whole payload is rendered in document context and
     * compared — any difference, or any error, means the document must be re-rendered after the write.
     */
    function needsDocumentRender(part, warnings) {
        if (warnings.some((warning) => warning.code === "isolated-render"))
            return true;
        const resolved = resolveExact(part.style.id, part.locale);
        const contextual = new CitationClusterRenderer({
            styleXml: resolved.style.style.xml,
            localeXml: resolved.locale.locale.xml,
            locale: part.locale,
            outputFormat: "text",
        }).render(part);
        return contextual.citationErrors.length > 0 || contextual.clusters.some((cluster) => cluster.text !== part.clusters.find((entry) => entry.clusterId === cluster.clusterId)?.renderedText);
    }
    /** Resolve a style and locale exactly, refusing any substitution (§9.3, invariant 3). */
    function resolveExact(styleId, locale) {
        let style;
        let resolvedLocale;
        try {
            style = styles.resolveStyle(styleId);
            resolvedLocale = styles.resolveLocale(locale);
        }
        catch {
            refuse("That citation style is not installed in Reference Manager.");
        }
        if (style.usedFallback)
            refuse("That citation style is not installed in Reference Manager.");
        if (resolvedLocale.usedFallback && resolvedLocale.fallbackReason !== "language-match") {
            refuse("This document's language is not installed in Reference Manager.");
        }
        return { style, locale: resolvedLocale };
    }
    /**
     * The live document with any citations the writer has confirmed deleted taken out of the payload.
     * Unconfirmed deletions come back as a 409 carrying the ids to confirm, never as a guess.
     */
    function confirmedDocument(live, body) {
        const missing = missingCitationIds(live);
        if (missing.length === 0)
            return { live, removed: [] };
        const accepted = body["acceptRemovedClusterIds"] === undefined
            ? []
            : stringList(body["acceptRemovedClusterIds"], "Invalid confirmation.");
        if (!missing.every((id) => accepted.includes(id))) {
            return {
                status: 409,
                body: {
                    error: "confirm-removed-citations",
                    removedClusterIds: missing,
                    message: missing.length === 1
                        ? "One citation is no longer in the document. Remove it from the bibliography too?"
                        : `${String(missing.length)} citations are no longer in the document. Remove them from the bibliography too?`,
                },
            };
        }
        const kept = new Set(missing);
        return {
            live: { ...live, part: { ...live.part, clusters: live.part.clusters.filter((cluster) => !kept.has(cluster.clusterId)) } },
            removed: missing,
        };
    }
    /** One processor bibliography render, from the library's current records. */
    async function renderBibliography(styleId, locale, itemIds) {
        for (const itemId of itemIds) {
            if ((await ports.library.getItem(itemId)) === null) {
                refuse("A cited reference is no longer in your library, so the bibliography cannot be written. Restore it in Reference Manager, then try again.");
            }
        }
        return ports.citation.formatBibliography({ styleId, locale, itemIds });
    }
    /**
     * Update every citation and the bibliography, optionally in a new style (§9.1 Refresh document,
     * Change citation style). Planning only: the pane writes the plan through its own compensated
     * Office batch, which re-checks the whole citation walk before touching anything.
     */
    async function refreshOperation(body) {
        const read = readLiveDocument(body["document"]);
        if (read === null)
            refuse("This document has no Reference Manager citations to update yet.");
        const confirmed = confirmedDocument(read, body);
        if ("status" in confirmed)
            return confirmed;
        const { live, removed } = confirmed;
        const part = live.part;
        const targetStyleId = typeof body["styleId"] === "string" && body["styleId"].length > 0 ? body["styleId"] : part.style.id;
        const changingStyle = targetStyleId !== part.style.id;
        const resolved = resolveExact(targetStyleId, part.locale);
        const workingPart = changingStyle ? { ...part, style: { id: targetStyleId } } : part;
        const documentOrder = live.controls.map((control) => parseClusterTag(control.tag)?.clusterId);
        if (documentOrder.some((id) => id === undefined))
            documentRefusal("invalid-control", "A citation in the document is not one Reference Manager wrote.");
        const editor = new CitationDocumentEditor({
            styleXml: resolved.style.style.xml,
            localeXml: resolved.locale.locale.xml,
            locale: resolved.locale.effectiveLocaleId,
            outputFormat: "text",
            styleRef: { id: targetStyleId },
        });
        let citationPlan;
        try {
            citationPlan = editor.refresh(workingPart, { documentOrder: documentOrder });
        }
        catch {
            refuse("The citations could not be re-rendered in this style. Check the document, then try again.");
        }
        // One processor render per bibliography scope the document declares (none when it has no
        // bibliography, or a locked one) — asked of the *re-rendered* part, whose clusters are in
        // Word's order. The payload's own order is insertion order, and a numeric style numbers its
        // bibliography in the order the records are handed over: rendering from the payload listed a
        // reference cited third as fourth, beside in-text numbers that were right.
        const requests = bibliographyRefreshRequests(citationPlan.part);
        const renders = [];
        for (const request of requests) {
            renders.push({ scopeId: request.scopeId, rendered: await renderBibliography(targetStyleId, part.locale, request.itemIds) });
        }
        const single = requests.length === 1 && requests[0].scopeId === DOCUMENT_SCOPE_ID && workingPart.bibliography?.scopes === undefined;
        const bibliographyInput = {
            ...(single ? { renderedBibliography: renders[0].rendered } : renders.length > 0 ? { scopedBibliographies: renders } : {}),
            ...(live.documentEntryTexts === undefined ? {} : { documentBibliographyEntryTexts: live.documentEntryTexts }),
            ...(live.documentBibliographies === undefined ? {} : { documentBibliographies: live.documentBibliographies }),
            ...(body["acceptBibliographyReplacement"] === true ? { acceptBibliographyManualReplacement: true } : {}),
        };
        const acceptManual = body["acceptManualReplacementIds"] === undefined
            ? {}
            : { acceptManualReplacementIds: stringList(body["acceptManualReplacementIds"], "Invalid confirmation.") };
        let plan;
        let notices;
        if (changingStyle) {
            const metadata = resolved.style.style.metadata;
            let currentClassName;
            try {
                const current = styles.resolveStyle(part.style.id);
                if (!current.usedFallback)
                    currentClassName = current.style.metadata.className;
            }
            catch {
                currentClassName = undefined;
            }
            const context = {
                requestedStyleId: targetStyleId,
                effectiveStyleId: resolved.style.effectiveStyleId,
                requestedLocale: part.locale,
                effectiveLocale: resolved.locale.effectiveLocaleId,
                className: metadata.className,
                ...(metadata.updated === undefined ? {} : { styleUpdated: metadata.updated }),
            };
            const outcome = prepareStyleChange({
                part,
                request: { style: { id: targetStyleId } },
                context,
                controls: live.controls,
                citationPlan,
                generator: GENERATOR,
                ...(currentClassName === undefined ? {} : { currentClassName }),
                ...acceptManual,
                ...bibliographyInput,
            });
            if (!outcome.ok) {
                if (outcome.refusal.code === "refresh-refused")
                    documentRefusal(refreshCode(outcome.refusal.message), "The document could not be updated in this style.");
                documentRefusal(outcome.refusal.code, "The document could not be changed to that style.");
            }
            plan = outcome.plan.refresh;
            notices = [...outcome.plan.refresh.notices, ...outcome.plan.notices];
        }
        else {
            const outcome = prepareDocumentRefresh({
                part,
                controls: live.controls,
                citationPlan,
                trigger: "explicit",
                generator: GENERATOR,
                ...acceptManual,
                ...bibliographyInput,
            });
            if (!outcome.ok) {
                if (outcome.refusal.code === "bibliography-refused")
                    bibliographyRefusal(outcome.refusal.message);
                documentRefusal(outcome.refusal.code, "The document could not be updated. Check it, then try again.");
            }
            plan = outcome.plan;
            notices = [...outcome.plan.notices];
        }
        return {
            status: 200,
            body: {
                plan,
                stalePartIds: live.stalePartIds,
                summary: {
                    styleChanged: changingStyle,
                    updatedCitations: plan.citations.length,
                    unchangedCitations: plan.unchangedClusterIds.length,
                    heldManualClusterIds: plan.heldManualClusterIds,
                    removedCitations: removed.length,
                    bibliographyUpdated: plan.bibliographies.length > 0,
                    noticeCodes: [...new Set(notices.map((notice) => notice.bibliographyCode ?? notice.code))],
                },
            },
        };
    }
    /** §9.1 Insert bibliography: plan the bibliography for the cited records, at the cursor. */
    async function bibliographyOperation(body) {
        const read = readLiveDocument(body["document"]);
        if (read === null)
            refuse("Insert a citation before adding a bibliography.");
        const confirmed = confirmedDocument(read, body);
        if ("status" in confirmed)
            return confirmed;
        const { live } = confirmed;
        const part = inDocumentOrder(live.part, live.controls);
        const { required } = bibliographyItemIds(part);
        if (required.length === 0)
            refuse("This document has no citations to list in a bibliography.");
        resolveExact(part.style.id, part.locale);
        const rendered = await renderBibliography(part.style.id, part.locale, required);
        const prepared = prepareBibliography({
            part,
            rendered,
            trigger: "explicit",
            generator: GENERATOR,
            ...(live.documentEntryTexts === undefined ? {} : { documentEntryTexts: live.documentEntryTexts }),
            ...(live.documentBibliographies === undefined ? {} : { documentBibliographyTags: live.documentBibliographies.map((control) => control.tag) }),
            ...(body["acceptBibliographyReplacement"] === true ? { acceptManualReplacement: true } : {}),
        });
        if (!prepared.ok) {
            if (prepared.refusal.code === "manual-edit-unresolved") {
                return {
                    status: 409,
                    body: {
                        error: "confirm-bibliography-replacement",
                        message: "The bibliography was edited by hand. Replace it with a freshly formatted one?",
                    },
                };
            }
            documentRefusal(prepared.refusal.code, "The bibliography could not be prepared. Check the document, then try again.");
        }
        return {
            status: 200,
            body: {
                plan: prepared.plan,
                stalePartIds: live.stalePartIds,
                summary: {
                    entries: prepared.plan.control.entries.length,
                    replacing: live.documentEntryTexts !== undefined,
                    noticeCodes: [...new Set(prepared.plan.notices.map((notice) => notice.code))],
                },
            },
        };
    }
    /**
     * The state the pane opens on. Connecting must not fail because of the document: a citation that
     * was deleted or typed over is something the Document tab can resolve, so it is reported as
     * `needsUpdate` rather than refused. Only a document whose citation data cannot be read at all is
     * a `problem`, and even then the library and the styles are still returned.
     */
    async function sessionOperation(body) {
        const installed = await ports.library.listStyles();
        const preferred = installed.find((style) => style.id === "http://www.zotero.org/styles/apa") ?? installed[0];
        if (preferred === undefined)
            refuse("No citation styles are installed in Reference Manager.");
        let part = null;
        let needsUpdate = false;
        let problem;
        try {
            part = readWordState(body["document"]).part;
        }
        catch (error) {
            if (!(error instanceof BridgeRequestError))
                throw error;
            try {
                part = readLiveDocument(body["document"])?.part ?? null;
                needsUpdate = true;
            }
            catch (inner) {
                if (!(inner instanceof BridgeRequestError))
                    throw inner;
                problem = inner.message;
            }
        }
        return {
            status: 200,
            body: {
                styleId: part?.style.id ?? preferred.id,
                locale: part?.locale ?? "en-US",
                styles: installed,
                ...(skippedStyles === 0 ? {} : { skippedStyleFiles: skippedStyles }),
                citedItemIds: Object.keys(part?.snapshots ?? {}),
                // What the Document tab states before the writer asks for anything.
                document: {
                    citations: part?.clusters.length ?? 0,
                    bibliography: part?.bibliography?.present === true,
                    bibliographyLocked: part?.bibliography?.locked === true,
                    needsUpdate,
                    ...(problem === undefined ? {} : { problem }),
                },
            },
        };
    }
    return async (raw) => {
        try {
            const request = record(raw);
            if (Object.keys(request).some((key) => !["path", "query", "body"].includes(key)))
                refuse("Unknown Word request field.");
            const path = request["path"];
            if (typeof path !== "string")
                refuse("Missing Word operation.");
            if (READ_PATHS.includes(path)) {
                const route = BRIDGE_ROUTES.find((entry) => entry.path === path);
                const query = request["query"] === undefined ? {} : record(request["query"]);
                if (Object.values(query).some((value) => typeof value !== "string"))
                    refuse("Invalid query.");
                return await router({ route: { ...route, params: {} }, query: query, body: request["body"] === undefined ? "" : JSON.stringify(request["body"]), tokenFingerprint: null }) ?? { status: 404, body: { error: "unavailable" } };
            }
            if (path === "/edit/open" || path === "/edit/prepare") {
                return await editOperation(path, record(request["body"]));
            }
            if (path === "/refresh/prepare")
                return await refreshOperation(record(request["body"]));
            if (path === "/bibliography/prepare")
                return await bibliographyOperation(record(request["body"]));
            if (path !== "/session" && path !== "/insertion/prepare")
                return { status: 404, body: { error: "unknown-operation" } };
            const body = record(request["body"]);
            if (path === "/session")
                return await sessionOperation(body);
            const state = readWordState(body["document"]);
            const draft = record(body["draft"]);
            if (!Array.isArray(draft.entries) || draft.entries.length > 50)
                refuse("Invalid citation draft.");
            const styleId = typeof body["styleId"] === "string" ? body["styleId"] : "";
            const locale = typeof body["locale"] === "string" ? body["locale"] : "";
            const built = citationRequest(draft, { styleId, locale });
            if (!built.ok)
                refuse(built.refusal.message);
            const rendered = await ports.citation.formatCitation(built.request);
            // A footnote style places each citation in a Word note, which this insertion does not do.
            // A numeric style is different: its number depends on where the citation lands, so it is
            // inserted with the isolated render and the whole document is renumbered straight after
            // (`refreshRequired`) — the payload and the visible text agree at every commit point.
            if (resolveExact(styleId, locale).style.style.metadata.className === "note") {
                refuse("Footnote citation styles are not supported for inserting yet. Choose an in-text style in the Document tab.");
            }
            if (body["previewText"] !== rendered.text)
                refuse("The preview changed. Review the refreshed preview and insert again.");
            const documentOrdered = rendered.warnings.some((warning) => warning.code === "isolated-render");
            const items = await Promise.all(built.request.items.map(async ({ itemId }) => {
                const item = await ports.library.getItem(itemId);
                if (item === null)
                    refuse("A selected reference is no longer available in this library.");
                return withSource(item);
            }));
            const prepared = prepareInsertion({ draft, styleId, locale, rendered, items, currentPart: state.part, docId: state.part?.docId ?? newId(), clusterId: newId(), generator: "refmgr-word/0.1", snapshotAt: now() });
            if (!prepared.ok)
                refuse(prepared.refusal.message);
            // Even author-date styles can disambiguate earlier citations ("2020a"). Never *leave* isolated
            // text in place when a whole-document render disagrees with it: the pane must follow this
            // write with a refresh, which renders every citation in Word's real order.
            const contextual = new CitationClusterRenderer({
                styleXml: styles.resolveStyle(styleId).style.xml,
                localeXml: styles.resolveLocale(locale).locale.xml, locale, outputFormat: "text",
            }).render(prepared.plan.part);
            if (contextual.citationErrors.length > 0) {
                refuse("This citation could not be formatted in the context of the document. Check the document, then try again.");
            }
            const refreshRequired = documentOrdered || contextual.clusters.some((cluster) => cluster.text !== prepared.plan.part.clusters.find((entry) => entry.clusterId === cluster.clusterId)?.renderedText);
            return { status: 200, body: { plan: prepared.plan, stalePartIds: state.stalePartIds, refreshRequired } };
        }
        catch (error) {
            if (error instanceof BridgeRequestError)
                return { status: error.status, body: { error: error.code, message: error.message } };
            return { status: 400, body: { error: "invalid-word-request", message: "The citation request could not be prepared. Check the selected references and style." } };
        }
    };
}
