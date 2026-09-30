/** Edit an existing Word citation (SPEC §9.2, task E10-05; ADR-0007/0107). */
import { XMLParser, XMLValidator } from "fast-xml-parser";
import { canonicalJson, fingerprint, parseClusterTag, readDocumentPart, shortFingerprint, wordCitationControl, wordCustomXml, } from "@refmgr/doc-schema";
import { prepareInsertion, } from "./insertion.js";
const xmlParser = new XMLParser({
    attributeNamePrefix: "@_",
    ignoreAttributes: false,
    parseTagValue: false,
    processEntities: true,
    removeNSPrefix: true,
    trimValues: false,
});
export function payloadText(xml) {
    if (/<!(?:DOCTYPE|ENTITY)\b/iu.test(xml) || XMLValidator.validate(xml) !== true)
        return null;
    try {
        const parsed = xmlParser.parse(xml);
        const document = parsed["document"];
        const payload = document?.["payload"];
        if (typeof payload === "string")
            return payload;
        if (payload !== null && typeof payload === "object") {
            const text = payload["#text"];
            return typeof text === "string" ? text : null;
        }
        return null;
    }
    catch {
        return null;
    }
}
function draftFromCluster(cluster) {
    const entries = [];
    for (const [index, item] of cluster.items.entries()) {
        if (item.suppressAuthor === true && item.authorOnly === true) {
            return {
                code: "unsupported-author-state",
                message: `cluster ${cluster.clusterId} item ${index} carries both author modes and cannot be edited without an explicit decision`,
            };
        }
        entries.push({
            entryId: `entry-${index + 1}`,
            itemId: item.itemId,
            ...(item.locator === undefined ? {} : { locator: { ...item.locator } }),
            ...(item.prefix === undefined ? {} : { prefix: item.prefix }),
            ...(item.suffix === undefined ? {} : { suffix: item.suffix }),
            authorMode: item.suppressAuthor === true ? "suppress-author" : item.authorOnly === true ? "author-only" : "normal",
        });
    }
    return { entries, nextOrdinal: entries.length + 1 };
}
function expectedVisibleText(cluster) {
    return cluster.manualEdit.state !== "none" && cluster.manualEdit.text !== undefined
        ? cluster.manualEdit.text
        : cluster.renderedText;
}
function stateWithoutTarget(part, targetClusterId) {
    const clusters = part.clusters.filter((cluster) => cluster.clusterId !== targetClusterId);
    const retainedItems = new Set(clusters.flatMap((cluster) => cluster.items.map((item) => item.itemId)));
    for (const id of part.bibliography?.pinnedItemIds ?? [])
        retainedItems.add(id);
    for (const id of part.bibliography?.uncitedItemIds ?? [])
        retainedItems.add(id);
    const snapshots = Object.fromEntries(Object.entries(part.snapshots).filter(([id]) => retainedItems.has(id)));
    const { generator: _writer, ...state } = part;
    return canonicalJson({ ...state, clusters, snapshots });
}
function hasForbiddenAuthorPair(raw, clusterId) {
    if (raw === null || typeof raw !== "object")
        return false;
    const clusters = raw["clusters"];
    if (!Array.isArray(clusters))
        return false;
    for (const cluster of clusters) {
        if (cluster === null || typeof cluster !== "object")
            continue;
        const value = cluster;
        if (value["clusterId"] !== clusterId || !Array.isArray(value["items"]))
            continue;
        return value["items"].some((item) => item !== null && typeof item === "object" &&
            item["suppressAuthor"] === true &&
            item["authorOnly"] === true);
    }
    return false;
}
/**
 * Select the payload identified by the selected control's cluster fingerprint.
 *
 * Cleanup failures can leave several namespace-matching custom XML parts (ADR-0107). The control
 * fingerprint is the only durable evidence of which cluster version Word currently displays, so
 * a disagreeing pair is refused rather than resolved by array order or part id.
 */
export function openCitationForEdit(control, candidates) {
    const tag = parseClusterTag(control.tag);
    if (tag === null) {
        return { ok: false, refusal: { code: "invalid-control", message: "the selected control is not a Reference Manager citation" } };
    }
    const ids = new Set();
    for (const candidate of candidates) {
        if (ids.has(candidate.id)) {
            return { ok: false, refusal: { code: "duplicate-part-id", message: `custom XML part id ${candidate.id} is repeated` } };
        }
        ids.add(candidate.id);
    }
    const readableParts = [];
    const notices = [];
    for (const candidate of candidates) {
        const json = payloadText(candidate.xml);
        const read = json === null ? null : readDocumentPart(json);
        if (read !== null && read.status === "unreadable" && hasForbiddenAuthorPair(read.raw, tag.clusterId)) {
            return {
                ok: false,
                refusal: {
                    code: "unsupported-author-state",
                    message: `cluster ${tag.clusterId} carries both author modes and cannot be edited without an explicit decision`,
                },
            };
        }
        if (read === null || read.status !== "ok") {
            notices.push({ code: "unreadable-custom-xml", partId: candidate.id, message: `custom XML part ${candidate.id} is not writable by this build` });
            continue;
        }
        const cluster = read.part.clusters.find((entry) => entry.clusterId === tag.clusterId);
        readableParts.push({ id: candidate.id, part: read.part, ...(cluster === undefined ? {} : { cluster }) });
    }
    const readable = readableParts.filter((entry) => entry.cluster !== undefined);
    if (readable.length === 0) {
        return { ok: false, refusal: { code: "cluster-not-found", message: `no readable custom XML part contains cluster ${tag.clusterId}` } };
    }
    const fingerprintMatches = readable.filter((entry) => shortFingerprint(entry.cluster) === tag.shortFingerprint);
    if (fingerprintMatches.length === 0) {
        return {
            ok: false,
            refusal: { code: "control-payload-mismatch", message: `cluster ${tag.clusterId} exists, but no payload matches the selected control fingerprint` },
        };
    }
    let selected = fingerprintMatches[0];
    if (fingerprintMatches.length > 1) {
        const canonicalParts = new Set(fingerprintMatches.map((entry) => canonicalJson(entry.part)));
        if (canonicalParts.size !== 1) {
            return {
                ok: false,
                refusal: { code: "ambiguous-payload", message: `more than one different payload matches cluster ${tag.clusterId}` },
            };
        }
        selected = [...fingerprintMatches].sort((left, right) => left.id.localeCompare(right.id))[0];
    }
    const stalePartIds = [];
    for (const other of readableParts) {
        if (other.id === selected.id)
            continue;
        if (stateWithoutTarget(other.part, tag.clusterId) !== stateWithoutTarget(selected.part, tag.clusterId)) {
            return {
                ok: false,
                refusal: {
                    code: "ambiguous-payload",
                    message: `custom XML part ${other.id} contains document changes outside cluster ${tag.clusterId}`,
                },
            };
        }
        stalePartIds.push(other.id);
    }
    for (const id of stalePartIds) {
        if (!notices.some((notice) => notice.partId === id)) {
            notices.push({ code: "stale-custom-xml", partId: id, message: `custom XML part ${id} is redundant or stale` });
        }
    }
    const draft = draftFromCluster(selected.cluster);
    if ("code" in draft)
        return { ok: false, refusal: draft };
    const manualEditDetected = selected.cluster.manualEdit.state !== "none" || control.visibleText !== expectedVisibleText(selected.cluster);
    if (manualEditDetected) {
        notices.push({ code: "manual-edit-detected", message: `cluster ${tag.clusterId} visible text differs from its stored formatting state` });
    }
    return {
        ok: true,
        session: {
            clusterId: tag.clusterId,
            sourcePartId: selected.id,
            stalePartIds,
            part: selected.part,
            originalCluster: selected.cluster,
            draft,
            manualEditDetected,
            notices,
        },
    };
}
/** Replace one stable cluster in place, preserving document order and unknown extension fields. */
export function prepareCitationEdit(input) {
    if (input.session.manualEditDetected && input.acceptManualEditReplacement !== true) {
        return { ok: false, refusal: { code: "manual-edit-unresolved", message: `cluster ${input.session.clusterId} has a visible edit that must be accepted or kept explicitly` } };
    }
    const index = input.session.part.clusters.findIndex((cluster) => cluster.clusterId === input.session.clusterId);
    if (index === -1) {
        return { ok: false, refusal: { code: "invalid-edit", message: `cluster ${input.session.clusterId} is absent from the selected payload` } };
    }
    const existingDetails = Object.values(input.session.part.snapshots).map((snapshot) => ({
        itemId: snapshot.itemId,
        csl: snapshot.csl,
        ...(snapshot.sourceLibraryId === undefined ? {} : { sourceLibraryId: snapshot.sourceLibraryId }),
    }));
    const supplied = new Map(existingDetails.map((detail) => [detail.itemId, detail]));
    for (const detail of input.items ?? [])
        supplied.set(detail.itemId, detail);
    // Old documents may carry a legacy/non-canonical contentHash. For the insertion validator's
    // live-vs-document comparison, normalise only the temporary comparison copy; the original
    // snapshots and timestamps are restored below and are never rewritten just because a cite moved.
    const comparisonSnapshots = Object.fromEntries(Object.entries(input.session.part.snapshots).map(([id, snapshot]) => [
        id,
        { ...snapshot, contentHash: fingerprint(snapshot.csl) },
    ]));
    const withoutTarget = {
        ...input.session.part,
        clusters: input.session.part.clusters.filter((cluster) => cluster.clusterId !== input.session.clusterId),
        snapshots: comparisonSnapshots,
    };
    const prepared = prepareInsertion({
        draft: input.draft,
        styleId: input.session.part.style.id,
        locale: input.session.part.locale,
        rendered: input.rendered,
        items: [...supplied.values()],
        currentPart: withoutTarget,
        docId: input.session.part.docId,
        clusterId: input.session.clusterId,
        generator: input.generator,
        snapshotAt: input.snapshotAt,
    });
    if (!prepared.ok)
        return { ok: false, refusal: { code: "invalid-edit", message: prepared.refusal.message } };
    const generated = prepared.plan.part.clusters.at(-1);
    const replacement = {
        ...input.session.originalCluster,
        ...generated,
        order: input.session.originalCluster.order ?? index,
        ...(input.session.originalCluster.noteNumber === undefined ? {} : { noteNumber: input.session.originalCluster.noteNumber }),
        manualEdit: { state: "none" },
        extensions: { ...input.session.originalCluster.extensions, ...generated.extensions },
    };
    const clusters = [...input.session.part.clusters];
    clusters[index] = replacement;
    const newSnapshots = Object.fromEntries(Object.entries(prepared.plan.part.snapshots).filter(([id]) => input.session.part.snapshots[id] === undefined));
    const part = {
        ...prepared.plan.part,
        generator: input.generator,
        clusters,
        snapshots: { ...input.session.part.snapshots, ...newSnapshots },
    };
    try {
        return {
            ok: true,
            plan: Object.freeze({
                part,
                customXml: wordCustomXml(part),
                control: wordCitationControl(replacement),
                warnings: prepared.plan.warnings,
            }),
            stalePartIds: [input.session.sourcePartId, ...input.session.stalePartIds],
        };
    }
    catch (cause) {
        return { ok: false, refusal: { code: "invalid-edit", message: cause instanceof Error ? cause.message : "edited document part is invalid" } };
    }
}
/** Payload first, selected control second, then old parts; the same compensation order as insert. */
export async function commitCitationEdit(port, expectedTag, prepared) {
    let added;
    try {
        added = await port.addCustomXml(prepared.plan.customXml);
    }
    catch {
        return { status: "failed", stage: "custom-xml", cleanupRequiredPartIds: [] };
    }
    try {
        await port.replaceContentControl(expectedTag, prepared.plan.control);
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
    for (const id of prepared.stalePartIds) {
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
