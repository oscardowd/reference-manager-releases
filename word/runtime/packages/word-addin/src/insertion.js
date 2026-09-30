/**
 * Prepare and commit a newly inserted Word citation (SPEC §9.2–9.3, ADR-0007).
 *
 * The preparation half is pure and produces the exact content-control descriptor and custom XML
 * bytes the host must write. The commit half acknowledges an Office.js constraint: custom XML and
 * Word content controls are different API surfaces, so no Office transaction spans both. It adds
 * the new payload first, inserts the control second, and compensates if the second write fails.
 * Cleanup failures are returned as data; they are never silently called success.
 */
import { DOC_SCHEMA_VERSION, fingerprint, wordCitationControl, wordCustomXml, } from "@refmgr/doc-schema";
import { citationRequest } from "./draft.js";
function refuse(code, message) {
    return { ok: false, refusal: { code, message } };
}
function clusterItems(draft) {
    return draft.entries.map((entry) => ({
        itemId: entry.itemId,
        ...(entry.locator === undefined ? {} : { locator: { ...entry.locator } }),
        ...(entry.prefix === undefined ? {} : { prefix: entry.prefix }),
        ...(entry.suffix === undefined ? {} : { suffix: entry.suffix }),
        ...(entry.authorMode === "suppress-author" ? { suppressAuthor: true } : {}),
        ...(entry.authorMode === "author-only" ? { authorOnly: true } : {}),
    }));
}
/** Build the document payload and control without touching Word. */
export function prepareInsertion(input) {
    const request = citationRequest(input.draft, { styleId: input.styleId, locale: input.locale });
    if (!request.ok)
        return refuse("invalid-draft", request.refusal.message);
    if (input.rendered.text.length === 0) {
        return refuse("empty-render", "the CSL processor returned no visible citation text");
    }
    const context = input.rendered.context;
    if (context.requestedStyleId !== input.styleId || context.requestedLocale !== input.locale) {
        return refuse("render-mismatch", "the rendered citation does not match the requested style and locale");
    }
    const current = input.currentPart;
    if (current !== null && (current.style.id !== input.styleId || current.locale !== input.locale)) {
        return refuse("document-preference-conflict", `the document uses style ${current.style.id} and locale ${current.locale}; change document preferences explicitly before insertion`);
    }
    if (current !== null && current.clusters.some((cluster) => cluster.clusterId === input.clusterId)) {
        return refuse("invalid-document-part", `cluster id ${input.clusterId} is already present in the document`);
    }
    const details = new Map(input.items.map((item) => [item.itemId, item]));
    const snapshots = { ...(current?.snapshots ?? {}) };
    for (const itemId of new Set(input.draft.entries.map((entry) => entry.itemId))) {
        const detail = details.get(itemId);
        if (detail === undefined)
            return refuse("missing-item", `no item detail was supplied for ${itemId}`);
        if (detail.mergedFrom !== undefined) {
            return refuse("merged-item", `item ${detail.mergedFrom} resolves to ${itemId}; insertion needs an explicit merge decision`);
        }
        if (detail.csl["id"] !== itemId || typeof detail.csl["type"] !== "string") {
            return refuse("invalid-snapshot", `item ${itemId} does not carry a matching CSL id and type`);
        }
        const contentHash = fingerprint(detail.csl);
        const existing = snapshots[itemId];
        if (existing !== undefined && existing.contentHash !== contentHash) {
            return refuse("snapshot-conflict", `item ${itemId} differs from the document-local snapshot; refresh or resolve the version before insertion`);
        }
        snapshots[itemId] = existing ?? {
            itemId,
            ...(detail.sourceLibraryId === undefined ? {} : { sourceLibraryId: detail.sourceLibraryId }),
            snapshotAt: input.snapshotAt,
            contentHash,
            csl: detail.csl,
        };
    }
    const cluster = {
        clusterId: input.clusterId,
        items: clusterItems(input.draft),
        order: current?.clusters.length ?? 0,
        renderedText: input.rendered.text,
        renderedHash: fingerprint(input.rendered.text),
        renderedWith: {
            id: context.requestedStyleId,
            ...(context.styleUpdated === undefined ? {} : { version: context.styleUpdated }),
        },
        manualEdit: { state: "none" },
        extensions: {
            "refmgr:render-context": {
                effectiveStyleId: context.effectiveStyleId,
                cslVersion: context.cslVersion,
                styleRevision: context.styleRevision,
                effectiveLocale: context.effectiveLocale,
            },
        },
    };
    const part = {
        schemaVersion: DOC_SCHEMA_VERSION,
        docId: current?.docId ?? input.docId,
        generator: input.generator,
        style: current?.style ?? {
            id: context.requestedStyleId,
            ...(context.styleUpdated === undefined ? {} : { version: context.styleUpdated }),
        },
        locale: current?.locale ?? input.locale,
        preferences: current?.preferences ?? {
            citationFormat: context.citationFormat === "note" ? "note" : "in-text",
            updateMode: "automatic",
        },
        clusters: [...(current?.clusters ?? []), cluster],
        ...(current?.bibliography === undefined ? {} : { bibliography: current.bibliography }),
        snapshots,
        ...(current?.extensions === undefined ? {} : { extensions: current.extensions }),
    };
    try {
        return {
            ok: true,
            plan: Object.freeze({
                part,
                customXml: wordCustomXml(part),
                control: wordCitationControl(cluster),
                warnings: Object.freeze([...input.rendered.warnings]),
            }),
        };
    }
    catch (cause) {
        return refuse("invalid-document-part", cause instanceof Error ? cause.message : "document part is invalid");
    }
}
/** Commit a prepared plan with explicit compensation and explicit stale-part reporting. */
export async function commitInsertion(port, plan, stalePartIds = []) {
    let added;
    try {
        added = await port.addCustomXml(plan.customXml);
    }
    catch {
        return { status: "failed", stage: "custom-xml", cleanupRequiredPartIds: [] };
    }
    try {
        await port.insertContentControl(plan.control);
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
    const failedCleanup = [];
    for (const id of stalePartIds) {
        if (id === added.id)
            continue;
        try {
            await port.deleteCustomXml(id);
        }
        catch {
            failedCleanup.push(id);
        }
    }
    return failedCleanup.length === 0
        ? { status: "written", customXmlPartId: added.id }
        : { status: "written-with-cleanup-required", customXmlPartId: added.id, stalePartIds: failedCleanup };
}
