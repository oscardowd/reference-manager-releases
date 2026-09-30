/** Whole-document Word refresh (SPEC §9.1, §21; task E10-07; ADR-0007/0059/0109). */
import { canonicalJson, DOCUMENT_SCOPE_ID, fingerprint, parseBibliographyTag, parseClusterTag, shortFingerprint, wordCitationControl, wordCustomXml, } from "@refmgr/doc-schema";
import { prepareBibliography, bibliographyItemIds, } from "./bibliography.js";
function refuse(code, message) {
    return { ok: false, refusal: { code, message } };
}
function expectedVisibleText(cluster) {
    return cluster.manualEdit.state !== "none" && cluster.manualEdit.text !== undefined
        ? cluster.manualEdit.text
        : cluster.renderedText;
}
/** The exact CSL requests a task pane must make before planning an atomic document refresh. */
export function bibliographyRefreshRequests(part) {
    if (part.bibliography?.present !== true || part.bibliography.locked === true)
        return [];
    const scopes = part.bibliography.scopes ?? [{ scopeId: DOCUMENT_SCOPE_ID, mode: "document" }];
    return Object.freeze(scopes.map((scope) => Object.freeze({
        scopeId: scope.scopeId,
        itemIds: Object.freeze([...bibliographyItemIds(part, scope.scopeId).required]),
    })));
}
function sameDocumentShape(before, plan) {
    if (plan.part.docId !== before.docId || plan.removed.length !== 0)
        return false;
    const beforeIds = [...before.clusters.map((cluster) => cluster.clusterId)].sort();
    const afterIds = [...plan.part.clusters.map((cluster) => cluster.clusterId)].sort();
    return canonicalJson(beforeIds) === canonicalJson(afterIds) &&
        canonicalJson(before.snapshots) === canonicalJson(plan.part.snapshots);
}
/**
 * Combine a processor-backed full-document render with Word's live controls.
 *
 * The processor plan says what citations should say; the controls say what they currently say.
 * Neither is allowed to stand in for the other. A newly detected manual edit is recorded as
 * `pending` and held verbatim unless its cluster id appears in the explicit replacement set.
 */
export function prepareDocumentRefresh(input) {
    if (input.trigger === "automatic" && input.part.preferences.updateMode === "on-demand") {
        return refuse("update-on-demand-only", "this document permits refresh only after an explicit request");
    }
    if (!sameDocumentShape(input.part, input.citationPlan)) {
        return refuse("render-plan-mismatch", "the citation plan does not describe the same document and snapshots");
    }
    const beforeById = new Map(input.part.clusters.map((cluster) => [cluster.clusterId, cluster]));
    const plannedById = new Map(input.citationPlan.part.clusters.map((cluster) => [cluster.clusterId, cluster]));
    const renderedById = new Map(input.citationPlan.rendered.clusters.map((cluster) => [cluster.clusterId, cluster]));
    const accepted = new Set(input.acceptManualReplacementIds ?? []);
    const seen = new Set();
    const ordered = [];
    const writes = [];
    const unchanged = [];
    const held = [];
    const notices = input.citationPlan.rendered.warnings.map((warning) => ({
        code: "citation-render-warning",
        message: warning.message,
        ...(warning.clusterId === undefined ? {} : { clusterId: warning.clusterId }),
    }));
    for (const [position, live] of input.controls.entries()) {
        const parsed = parseClusterTag(live.tag);
        if (parsed === null)
            return refuse("invalid-control", `control at position ${position} is not a Reference Manager citation`);
        if (seen.has(parsed.clusterId))
            return refuse("duplicate-control", `cluster ${parsed.clusterId} appears in more than one live control`);
        seen.add(parsed.clusterId);
        const before = beforeById.get(parsed.clusterId);
        const planned = plannedById.get(parsed.clusterId);
        const rendered = renderedById.get(parsed.clusterId);
        if (before === undefined || planned === undefined || rendered === undefined) {
            return refuse("control-payload-mismatch", `cluster ${parsed.clusterId} is not present in every refresh input`);
        }
        if (shortFingerprint(before) !== parsed.shortFingerprint) {
            return refuse("control-payload-mismatch", `cluster ${parsed.clusterId} does not match the live control fingerprint`);
        }
        const recorded = expectedVisibleText(before);
        const manual = before.manualEdit.state !== "none" || recorded === undefined || live.visibleText !== recorded;
        let next;
        if (manual && !accepted.has(parsed.clusterId)) {
            const state = before.manualEdit.state === "kept" && before.manualEdit.text === live.visibleText ? "kept" : "pending";
            const { renderedText: _plannedText, renderedHash: _plannedHash, renderedWith: _plannedStyle, ...plannedStructural } = planned;
            next = {
                ...plannedStructural,
                order: position,
                // These fields describe what was actually last written, so a held citation must not
                // advance them to processor output that Word never received.
                ...(before.renderedText === undefined ? {} : { renderedText: before.renderedText }),
                ...(before.renderedHash === undefined ? {} : { renderedHash: before.renderedHash }),
                ...(before.renderedWith === undefined ? {} : { renderedWith: before.renderedWith }),
                manualEdit: { state, text: live.visibleText },
            };
            held.push(parsed.clusterId);
            notices.push({ code: "manual-edit-held", clusterId: parsed.clusterId, message: `cluster ${parsed.clusterId} has live text that was not overwritten` });
        }
        else {
            next = {
                ...planned,
                order: position,
                renderedText: rendered.text,
                renderedHash: fingerprint(rendered.text),
                manualEdit: { state: "none" },
            };
            if (manual)
                notices.push({ code: "manual-edit-replaced", clusterId: parsed.clusterId, message: `cluster ${parsed.clusterId} was replaced after an explicit decision` });
        }
        ordered.push(next);
        let descriptor;
        try {
            descriptor = wordCitationControl(next);
        }
        catch (cause) {
            return refuse("invalid-document-part", cause instanceof Error ? cause.message : `cluster ${parsed.clusterId} has no writable output`);
        }
        if (descriptor.tag === live.tag && descriptor.visibleText === live.visibleText)
            unchanged.push(parsed.clusterId);
        else
            writes.push({ clusterId: parsed.clusterId, expectedTag: live.tag, control: descriptor });
    }
    if (seen.size !== beforeById.size) {
        return refuse("control-payload-mismatch", `the live document has ${seen.size} citation controls but the payload has ${beforeById.size} clusters`);
    }
    for (const id of accepted) {
        if (!seen.has(id))
            return refuse("control-payload-mismatch", `manual replacement names unknown cluster ${id}`);
    }
    let part = { ...input.citationPlan.part, generator: input.generator, clusters: ordered };
    const bibliographies = [];
    const expectedBibliographyTags = input.documentBibliographies?.map((control) => control.tag) ??
        (input.documentBibliographyEntryTexts === undefined ? [] : ["refmgr-bib"]);
    if (part.bibliography?.present === true) {
        if (part.bibliography.locked === true) {
            notices.push({ code: "bibliography-locked", message: "the locked bibliography was left unchanged" });
        }
        else {
            const requests = bibliographyRefreshRequests(part);
            const supplied = input.scopedBibliographies ?? (input.renderedBibliography === undefined ? [] : [{
                    scopeId: requests[0]?.scopeId ?? DOCUMENT_SCOPE_ID,
                    rendered: input.renderedBibliography,
                    ...(input.documentBibliographyEntryTexts === undefined ? {} : { documentEntryTexts: input.documentBibliographyEntryTexts }),
                    ...(input.acceptBibliographyManualReplacement === undefined ? {} : { acceptManualReplacement: input.acceptBibliographyManualReplacement }),
                }]);
            const byScope = new Map();
            for (const rendered of supplied) {
                if (byScope.has(rendered.scopeId)) {
                    return refuse("bibliography-refused", `bibliography scope ${rendered.scopeId} was rendered more than once`);
                }
                byScope.set(rendered.scopeId, rendered);
            }
            if (supplied.length !== requests.length || requests.some((request) => !byScope.has(request.scopeId))) {
                return refuse("bibliography-render-required", `an unlocked bibliography requires one processor render for each of its ${requests.length} scope(s)`);
            }
            const liveByScope = new Map();
            for (const control of input.documentBibliographies ?? []) {
                const parsed = parseBibliographyTag(control.tag);
                if (parsed === null || liveByScope.has(parsed.scopeId)) {
                    return refuse("bibliography-refused", "the live bibliography controls do not identify one unique control per scope");
                }
                liveByScope.set(parsed.scopeId, control);
            }
            if (input.documentBibliographies === undefined && input.documentBibliographyEntryTexts !== undefined) {
                liveByScope.set(DOCUMENT_SCOPE_ID, {
                    tag: "refmgr-bib",
                    entryTexts: input.documentBibliographyEntryTexts,
                });
            }
            if (input.documentBibliographies !== undefined &&
                (liveByScope.size !== requests.length || requests.some((request) => !liveByScope.has(request.scopeId)))) {
                return refuse("bibliography-refused", "the live bibliography controls do not match the declared scopes");
            }
            for (const request of requests) {
                const rendered = byScope.get(request.scopeId);
                const live = liveByScope.get(request.scopeId);
                const prepared = prepareBibliography({
                    part,
                    rendered: rendered.rendered,
                    scopeId: request.scopeId,
                    trigger: input.trigger,
                    generator: input.generator,
                    documentBibliographyTags: expectedBibliographyTags,
                    ...(live === undefined && rendered.documentEntryTexts === undefined
                        ? {}
                        : { documentEntryTexts: live?.entryTexts ?? rendered.documentEntryTexts }),
                    ...(rendered.acceptManualReplacement === undefined ? {} : { acceptManualReplacement: rendered.acceptManualReplacement }),
                    ...(input.reapplyBibliographyLayout === undefined ? {} : { reapplyLayout: input.reapplyBibliographyLayout }),
                    ...(input.hangingIndentTwips === undefined ? {} : { hangingIndentTwips: input.hangingIndentTwips }),
                });
                if (!prepared.ok) {
                    return refuse("bibliography-refused", `bibliography scope ${request.scopeId} refresh refused: ${prepared.refusal.code}`);
                }
                part = prepared.plan.part;
                bibliographies.push({
                    scopeId: request.scopeId,
                    expectedTag: live?.tag ?? prepared.plan.control.tag,
                    control: prepared.plan.control,
                });
                notices.push(...prepared.plan.notices.map((notice) => ({
                    code: "bibliography-notice",
                    bibliographyCode: notice.code,
                    message: notice.message,
                })));
            }
        }
    }
    try {
        return {
            ok: true,
            plan: Object.freeze({
                part,
                customXml: wordCustomXml(part),
                citations: Object.freeze(writes),
                expectedCitationTags: Object.freeze(input.controls.map((control) => control.tag)),
                unchangedClusterIds: Object.freeze(unchanged),
                heldManualClusterIds: Object.freeze(held),
                bibliographies: Object.freeze(bibliographies),
                expectedBibliographyTags: Object.freeze(expectedBibliographyTags),
                notices: Object.freeze(notices),
            }),
        };
    }
    catch (cause) {
        return refuse("invalid-document-part", cause instanceof Error ? cause.message : "the refreshed document part is invalid");
    }
}
/** Payload first, one Word control batch second, stale payloads last (ADR-0107). */
export async function commitDocumentRefresh(port, plan, stalePartIds = []) {
    let added;
    try {
        added = await port.addCustomXml(plan.customXml);
    }
    catch {
        return { status: "failed", stage: "custom-xml", cleanupRequiredPartIds: [] };
    }
    try {
        await port.writeDocumentControls(plan);
    }
    catch {
        try {
            await port.deleteCustomXml(added.id);
            return { status: "failed", stage: "document-controls", cleanupRequiredPartIds: [] };
        }
        catch {
            return { status: "failed", stage: "document-controls", cleanupRequiredPartIds: [added.id] };
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
