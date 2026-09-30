/**
 * `/import`: §9.2 "create or import a missing reference without leaving Word" (task E10-04.2;
 * SPEC §5.1, §9.2; ADR-0008).
 *
 * This is the bridge's **one write**. Everything else the add-in can ask for reads the library;
 * this route creates records in it from something the user pasted and something a metadata
 * provider said, which makes it the route where a mistake is durable. Three consequences shape it.
 *
 * - **The library is asked before the network is.** A DOI already in the library resolves from the
 *   library, with `created: false` and no request to anybody. §5.5 states the rule for the browser
 *   connector — "detect an existing library record before creating a duplicate" — and the reason
 *   is not politeness to Crossref: a second record for a work the manuscript already cites is how
 *   one paper acquires two entries in a bibliography.
 * - **A disagreement is refused, not resolved.** `resolveReconciliation` returns nothing writable
 *   while a conflict is undecided (invariant 3, §5.1.4), and the side-by-side comparison that
 *   decides one is a desktop-application surface. So the identifier comes back `source-conflict`
 *   with the disputed field names, rather than a record built from whichever source happened to
 *   rank first.
 * - **Every identifier is accounted for.** Each pasted entry either produces a reference or
 *   appears in `unresolved` with a code and a reason. "We imported 7 of your 9" is only acceptable
 *   if the product can say which two and why — the rule `recognizeIdentifiers` was written to, and
 *   the same rule applies once the network is involved.
 *
 * The `identifier` echoed in a failure is the user's own text and is display-only: it must never
 * reach a log (invariant 6, §23). Nothing here formats a citation.
 */
import { generateUuidV7 } from "@refmgr/core";
import { recognizeIdentifiers, reconcileBatch, resolveReconciliation, retrieveMetadata, } from "@refmgr/identifiers";
import { materialiseReference } from "./materialise.js";
import { toSearchHit } from "./search-hit.js";
function identifierKey(identifier) {
    return `${identifier.kind}:${identifier.key}`;
}
/** Why a pasted line yielded no identifier, in words that do not repeat the line back. */
function recognitionReason(problems) {
    return problems.some((problem) => problem.reason === "isbn-check-digit")
        ? "this looks like an ISBN but its check digit does not verify, so it names a different book"
        : "no DOI, PMID, PMCID, ISBN, arXiv id, trial id or URL could be read from this text";
}
/**
 * Collapse a provider's failures for one identifier into the single thing the add-in should do
 * next. A retryable failure outranks a definite one: "could not reach a source" and "no source has
 * this record" lead to different next actions, and reporting the second when the first happened
 * would tell the user their DOI is wrong when their network is.
 */
function failureFor(pending, failures) {
    const mine = failures.filter((failure) => identifierKey(failure.identifier) === identifierKey(pending.identifier));
    const chosen = mine.find((failure) => failure.retryable) ?? mine[0];
    if (chosen === undefined) {
        // No candidate and no failure: the sources answered, and none of them had this record.
        return { identifier: pending.display, code: "not-found", reason: "no metadata source had a record for this identifier", retryable: false };
    }
    const code = (() => {
        switch (chosen.code) {
            case "offline":
                return "offline";
            case "not-found":
                return "not-found";
            case "unsupported":
                return "unsupported";
            default:
                // `http`, `network`, `invalid-response` and `cancelled` are all "the lookup did not
                // complete". The distinction between them is a diagnostic, and a diagnostic that names
                // which provider answered how does not belong in a response leaving the bridge.
                return "network";
        }
    })();
    const reason = (() => {
        switch (code) {
            case "offline":
                return "metadata retrieval is switched off in this session";
            case "not-found":
                return "no metadata source had a record for this identifier";
            case "unsupported":
                return "this build has no metadata source for that kind of identifier";
            default:
                return "a metadata source could not be reached, or answered with something unusable";
        }
    })();
    return { identifier: pending.display, code, reason, retryable: chosen.retryable };
}
/** §5.1.6's retraction, correction and duplicate warnings, carried to the add-in unchanged. */
function toImportWarnings(warnings) {
    return warnings.map((warning) => ({ code: warning.code, message: warning.message }));
}
/**
 * Recognition → retrieval → reconciliation → a record in the library.
 *
 * Constructed by the host and handed to `LibraryBridgeAdapter`; the adapter itself knows only that
 * an importer exists, so a build that has not configured one still answers `/import` honestly
 * (see `unavailableImporter`).
 */
export class IdentifierBridgeImporter {
    #libraryId;
    #items;
    #retrieval;
    #now;
    #newId;
    #registry;
    #itemTypes;
    #actor;
    constructor(options) {
        this.#libraryId = options.libraryId;
        this.#items = options.items;
        this.#retrieval = options.retrieval ?? {};
        this.#now = options.now ?? (() => new Date().toISOString());
        this.#newId = options.newId ?? generateUuidV7;
        this.#registry = options.registry;
        this.#itemTypes = options.itemTypes;
        this.#actor = options.actor;
    }
    async import(request) {
        const imported = [];
        const unresolved = [];
        const pending = [];
        const seen = new Set();
        for (const entry of request.identifiers) {
            // Recognition runs per pasted entry rather than over the joined text, so a failure can be
            // reported beside the thing the caller sent rather than beside a line number in a blob we
            // assembled. An entry holding several identifiers still yields all of them.
            const batch = recognizeIdentifiers(entry);
            if (batch.identifiers.length === 0) {
                unresolved.push({ identifier: entry, code: "unrecognised", reason: recognitionReason(batch.problems), retryable: false });
                continue;
            }
            for (const identifier of batch.identifiers) {
                const key = identifierKey(identifier);
                // The same identifier pasted twice is one reference, reported once. It is not dropped
                // silently: the outcome for that identifier is in the result, exactly once.
                if (seen.has(key))
                    continue;
                seen.add(key);
                const existing = this.#findExisting(identifier);
                if (existing !== undefined) {
                    imported.push({ item: toSearchHit(existing), created: false, warnings: [] });
                    continue;
                }
                pending.push({ identifier, display: identifier.matchedText });
            }
        }
        if (pending.length > 0) {
            const retrieved = await retrieveMetadata(pending.map((entry) => entry.identifier), this.#retrieval);
            const reconciled = reconcileBatch(retrieved);
            // Every record here was reconciled from candidates this method retrieved *by identifier*, so
            // every one of them has one. The filter is how that stays true rather than an assumption:
            // a record with none could not be matched back to the line the user pasted, and dropping it
            // silently would report the paste unresolved — which is what `pending` already does below,
            // and is the honest outcome for a record this importer cannot attribute.
            const byIdentifier = new Map(reconciled.records
                .filter((record) => record.identifier !== undefined)
                .map((record) => [identifierKey(record.identifier), record]));
            for (const entry of pending) {
                const record = byIdentifier.get(identifierKey(entry.identifier));
                if (record === undefined) {
                    unresolved.push(failureFor(entry, retrieved.failures));
                    continue;
                }
                const outcome = this.#create(entry, record);
                if ("failure" in outcome)
                    unresolved.push(outcome.failure);
                else
                    imported.push(outcome.reference);
            }
        }
        return { imported, unresolved };
    }
    /**
     * The library's own record for this identifier, if it has one.
     *
     * DOI and PMID only, because those are the two identifiers the schema indexes on `items`
     * (`items_by_library_and_doi`, `items_by_library_and_pmid`). An ISBN or an arXiv id lives in a
     * sparse field with no such index, so a second record *can* be created for one of those from
     * here — a real gap, recorded in the status document rather than papered over, and the thing
     * E07's duplicate detection exists to catch afterwards.
     */
    #findExisting(identifier) {
        if (identifier.kind === "doi")
            return this.#items.findByDoi(this.#libraryId, identifier.value);
        if (identifier.kind === "pmid")
            return this.#items.findByPmid(this.#libraryId, identifier.value);
        return undefined;
    }
    /** The same DOI/PMID check, asked of a record that has been built but not yet written. */
    #findStored(item) {
        const byDoi = item.doi === undefined ? undefined : this.#items.findByDoi(this.#libraryId, item.doi);
        if (byDoi !== undefined)
            return byDoi;
        return item.pmid === undefined ? undefined : this.#items.findByPmid(this.#libraryId, item.pmid);
    }
    #create(entry, record) {
        // No decisions are passed, and none can be: deciding a conflict needs the comparison §5.1.4
        // describes, and there is no such surface in a task pane. An undecided conflict therefore
        // writes nothing at all — `fields` is empty whenever `status` is `unresolved`.
        const resolved = resolveReconciliation(record);
        if (resolved.status !== "resolved") {
            return {
                failure: {
                    identifier: entry.display,
                    code: "source-conflict",
                    // Field *names*, never the disputed values: the values are reference metadata.
                    reason: `the metadata sources disagree about ${resolved.pending.join(", ")}; resolve this reference in the desktop application`,
                    retryable: false,
                },
            };
        }
        const material = materialiseReference({
            itemId: this.#newId(),
            libraryId: this.#libraryId,
            createdAt: this.#now(),
            fields: resolved.fields,
            extensions: record.extensions,
            ...(this.#registry === undefined ? {} : { registry: this.#registry }),
            ...(this.#itemTypes === undefined ? {} : { itemTypes: this.#itemTypes }),
        });
        if (!material.ok) {
            return {
                failure: {
                    identifier: entry.display,
                    code: "incomplete-metadata",
                    reason: material.issues[0]?.message ?? "the retrieved metadata cannot be stored as a reference",
                    retryable: false,
                },
            };
        }
        // The library is asked a second time, now with the identifiers the *metadata* carries rather
        // than the one the user pasted. Pasting a paper's PMID when its DOI is already in the library
        // is the ordinary way this happens, and the first check could not have known: it had a PMID
        // and the stored record has only a DOI. The check also sees records written earlier in this
        // same batch, so pasting both identifiers for one paper creates one reference, not two.
        const already = this.#findStored(material.item);
        if (already !== undefined) {
            return { reference: { item: toSearchHit(already), created: false, warnings: toImportWarnings(record.warnings) } };
        }
        // A repository failure is not a per-identifier outcome — it means the library itself is not
        // writable — so it is left to propagate to the listener's 500 path, which says nothing about it.
        const stored = this.#items.insert(material.item, this.#actor === undefined ? {} : { actor: this.#actor });
        return { reference: { item: toSearchHit(stored), created: true, warnings: toImportWarnings(record.warnings) } };
    }
}
