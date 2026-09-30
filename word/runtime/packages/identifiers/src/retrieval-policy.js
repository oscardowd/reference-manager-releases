/**
 * What may leave this machine when a user pastes an identifier (SPEC §5.1, §20; task E04-15.1).
 *
 * E04-08.2 can retrieve metadata and deliberately refuses to until a host says so: `retrieveMetadata`
 * contacts nothing unless a transport is supplied or the built-in fetch is explicitly enabled. That
 * default is right, and it left the product with a §5.1 importer nobody can use — ADR-0144 recorded
 * the consequence honestly and E04-15 exists to close it. This module is the half of that answer the
 * domain owns: **the statement of what is disclosed, and the single place a stored setting becomes
 * retrieval options.**
 *
 * Three rules shape it.
 *
 * 1. **The disclosure is derived from the request, not written beside it.** {@link RETRIEVAL_DISCLOSURES}
 *    names, per provider, the host contacted, the identifier kinds it is asked about, and whether a
 *    contact e-mail reaches it — and `retrieval-policy.test.ts` drives every shipped source through a
 *    recording transport and asserts the table against the URL and headers the source actually built.
 *    A disclosure table nobody checks is a privacy notice; a privacy notice that has drifted from the
 *    code is worse than none, because it is believed. Adding a source without a disclosure entry, or
 *    changing a host, fails the suite.
 *
 * 2. **Reading a setting can fail closed and can never fail open.** {@link parseRetrievalSettings}
 *    takes untrusted JSON — a file on disk that a person, another program or a corrupted write may
 *    have touched — and every path through it that is not a complete, well-formed, explicitly-enabled
 *    setting yields {@link RETRIEVAL_SETTINGS_OFF}. There is deliberately no "default on" branch and
 *    no partial recovery that leaves retrieval enabled: a settings file that says something this build
 *    does not understand is not permission to contact anyone.
 *
 * 3. **The contact e-mail is a separate disclosure with a separate opt-in.** Enabling a provider says
 *    "ask them about this identifier". Supplying an e-mail address says "and tell them who I am", which
 *    is personal data about the user rather than about the paper, and is a different decision. It is
 *    carried only when it is present and well-formed, and it reaches only the providers whose published
 *    convention it is (Crossref's polite pool, and DataCite, which follows the same practice) — not
 *    every provider a user happens to have enabled.
 *
 * This module performs no I/O, reads no clock and logs nothing. It holds an e-mail address, which is
 * personal data (§20), so nothing here may print or record one: the settings it returns are meant to
 * be resolved into a request and rendered into a sentence, never logged (§23, X-03).
 */
import { IDENTIFIER_KINDS } from "./recognize.js";
import { METADATA_SOURCE_IDS, createMetadataSources, } from "./retrieval.js";
/**
 * The providers this build can contact, and what each is told.
 *
 * Ordered as `METADATA_SOURCE_IDS` is, and complete with respect to it — a source with no entry here
 * is a source whose disclosure nobody wrote, which is the failure this table exists to make loud.
 */
export const RETRIEVAL_DISCLOSURES = Object.freeze([
    Object.freeze({
        source: "crossref",
        host: "api.crossref.org",
        kinds: Object.freeze(["doi"]),
        receivesContactEmail: true,
    }),
    Object.freeze({
        source: "datacite",
        host: "api.datacite.org",
        kinds: Object.freeze(["doi"]),
        receivesContactEmail: true,
    }),
    Object.freeze({
        source: "pubmed",
        host: "eutils.ncbi.nlm.nih.gov",
        kinds: Object.freeze(["pmid", "pmcid"]),
        receivesContactEmail: false,
    }),
    Object.freeze({
        source: "openlibrary",
        host: "openlibrary.org",
        kinds: Object.freeze(["isbn"]),
        receivesContactEmail: false,
    }),
    Object.freeze({
        source: "arxiv",
        host: "export.arxiv.org",
        kinds: Object.freeze(["arxiv"]),
        receivesContactEmail: false,
    }),
    Object.freeze({
        source: "clinicaltrials",
        host: "clinicaltrials.gov",
        kinds: Object.freeze(["nct"]),
        receivesContactEmail: false,
    }),
]);
/** The disclosure for one provider, or `undefined` for an id this build does not ship. */
export function retrievalDisclosure(source) {
    return RETRIEVAL_DISCLOSURES.find((entry) => entry.source === source);
}
/** The providers that receive a contact e-mail when one is supplied. */
export const CONTACT_EMAIL_SOURCES = Object.freeze(RETRIEVAL_DISCLOSURES.filter((entry) => entry.receivesContactEmail).map((entry) => entry.source));
/**
 * Everything off.
 *
 * The value a build starts from, the value an unreadable settings file resolves to, and the value the
 * shipped shell has behaved as since ADR-0144 — stated once so that "off" is a thing code can return
 * rather than a shape three call sites each rebuild slightly differently.
 */
export const RETRIEVAL_SETTINGS_OFF = Object.freeze({
    enabled: false,
    sources: Object.freeze([]),
});
/** A pragmatic address check: one `@`, something either side, a dotted domain, no whitespace. */
const CONTACT_EMAIL_PATTERN = /^[^\s@]+@[^\s@.]+(?:\.[^\s@.]+)+$/u;
function isMetadataSourceId(value) {
    return typeof value === "string" && METADATA_SOURCE_IDS.includes(value);
}
/**
 * Read a stored setting.
 *
 * The contract is one sentence: **the result is never more permissive than the input clearly asked
 * for.** A value that is not an object, an `enabled` that is not a boolean, a `sources` that is not an
 * array — each of those is a file this build cannot read, and an unreadable file is not consent, so
 * each returns {@link RETRIEVAL_SETTINGS_OFF} with the reason.
 *
 * Two problems are *narrowing* rather than fatal, and they are the interesting cases. An unknown source
 * id is dropped and the rest of the choice is honoured: it is what a setting written by a newer build
 * looks like, and refusing the whole file would silently turn off the providers the user did choose.
 * An unusable contact e-mail is dropped for the opposite reason — it cannot be honoured *at all*, since
 * a malformed address in a `User-Agent` header is still disclosed and still useless — and dropping it
 * leaves the providers enabled but contacted anonymously, which is the safer half of the user's intent.
 */
export function parseRetrievalSettings(value) {
    const problems = [];
    if (typeof value !== "object" || value === null || Array.isArray(value)) {
        return { settings: RETRIEVAL_SETTINGS_OFF, problems: [{ code: "not-an-object" }] };
    }
    const raw = value;
    const enabledValue = raw["enabled"];
    if (enabledValue !== undefined && typeof enabledValue !== "boolean") {
        return { settings: RETRIEVAL_SETTINGS_OFF, problems: [{ code: "invalid-enabled" }] };
    }
    const enabled = enabledValue === true;
    const sourcesValue = raw["sources"];
    if (sourcesValue !== undefined && !Array.isArray(sourcesValue)) {
        return { settings: RETRIEVAL_SETTINGS_OFF, problems: [{ code: "invalid-sources" }] };
    }
    const chosen = new Set();
    for (const entry of Array.isArray(sourcesValue) ? sourcesValue : []) {
        if (isMetadataSourceId(entry)) {
            chosen.add(entry);
            continue;
        }
        problems.push({ code: "unknown-source", ...(typeof entry === "string" ? { source: entry } : {}) });
    }
    // Emitted in the canonical order rather than the file's, so two settings that enable the same
    // providers are the same value however they were written.
    const sources = METADATA_SOURCE_IDS.filter((id) => chosen.has(id));
    const emailValue = raw["contactEmail"];
    let contactEmail;
    if (emailValue !== undefined && emailValue !== null && emailValue !== "") {
        const trimmed = typeof emailValue === "string" ? emailValue.trim() : "";
        if (CONTACT_EMAIL_PATTERN.test(trimmed)) {
            contactEmail = trimmed;
        }
        else {
            problems.push({ code: "invalid-contact-email" });
        }
    }
    return {
        settings: Object.freeze({
            enabled,
            sources: Object.freeze(sources),
            ...(contactEmail === undefined ? {} : { contactEmail }),
        }),
        problems: Object.freeze(problems),
    };
}
// ---------------------------------------------------------------------------------------------
// What the setting does
// ---------------------------------------------------------------------------------------------
/**
 * The identifier kinds a setting can actually answer for.
 *
 * The honest counterpart to enabling providers one at a time: with only Crossref on, a pasted ISBN has
 * nowhere to go, and the user should be told that rather than watching it come back unresolved beside
 * a DOI that worked. Empty while retrieval is off.
 */
export function retrievableKinds(settings) {
    if (!settings.enabled)
        return Object.freeze([]);
    const kinds = new Set();
    for (const source of settings.sources) {
        for (const kind of retrievalDisclosure(source)?.kinds ?? [])
            kinds.add(kind);
    }
    return Object.freeze(IDENTIFIER_KINDS.filter((kind) => kinds.has(kind)));
}
/**
 * The disclosures a setting actually makes, in canonical order.
 *
 * `receivesContactEmail` is narrowed to what *this* setting discloses: a provider that would take an
 * address is reported as not receiving one when the user has supplied none. The surface renders this
 * rather than {@link RETRIEVAL_DISCLOSURES}, so what a user reads is what their own setting does.
 */
export function activeDisclosures(settings) {
    if (!settings.enabled)
        return Object.freeze([]);
    const chosen = new Set(settings.sources);
    const hasEmail = settings.contactEmail !== undefined;
    return Object.freeze(RETRIEVAL_DISCLOSURES.filter((entry) => chosen.has(entry.source)).map((entry) => Object.freeze({ ...entry, receivesContactEmail: entry.receivesContactEmail && hasEmail })));
}
/**
 * Turn a setting into the options `retrieveMetadata` takes.
 *
 * Three properties, and each is a refusal rather than a configuration.
 *
 * **Off means `networkPolicy: "disabled"` and every source still registered.** Registered, because the
 * failure a user should see for an identifier nobody looked up is `offline` — retryable, "retrieval is
 * disabled" — and not `unsupported`, which means "no source exists for this kind" and is not something
 * turning the setting on would fix.
 *
 * **On means only the chosen sources are offered at all.** A disabled provider is absent from the list
 * rather than filtered inside the loop, so there is no code path on which it can be contacted: the
 * request that would disclose to it is never constructed.
 *
 * **The contact e-mail is passed only with the providers that receive it.** `contactEmailSources` is
 * what makes {@link RetrievalDisclosure.receivesContactEmail} a fact rather than a claim.
 */
export function resolveRetrievalOptions(settings, input = {}) {
    const all = input.sources ?? createMetadataSources();
    const shared = {
        ...(input.transport === undefined ? {} : { transport: input.transport }),
        ...(input.cache === undefined ? {} : { cache: input.cache }),
        ...(input.signal === undefined ? {} : { signal: input.signal }),
        ...(input.userAgent === undefined ? {} : { userAgent: input.userAgent }),
        ...(input.now === undefined ? {} : { now: input.now }),
    };
    if (!settings.enabled) {
        return Object.freeze({ ...shared, sources: all, networkPolicy: "disabled" });
    }
    const chosen = new Set(settings.sources);
    return Object.freeze({
        ...shared,
        sources: Object.freeze(all.filter((source) => chosen.has(source.id))),
        networkPolicy: "enabled",
        ...(settings.contactEmail === undefined
            ? {}
            : { contactEmail: settings.contactEmail, contactEmailSources: CONTACT_EMAIL_SOURCES }),
    });
}
