/**
 * Metadata retrieval for recognised identifiers (SPEC §5.1; task E04-08.2).
 *
 * This layer deliberately stops at provider candidates. It never merges candidates or overwrites
 * a library item; field-by-field reconciliation is E04-09. Network access is also explicit: a
 * caller may inject a transport for tests or a host may opt in to the built-in fetch transport.
 * A failed lookup is returned as a structured failure so an offline import remains recognised and
 * can be retried later.
 */
export const METADATA_SOURCE_IDS = [
    "crossref",
    "datacite",
    "pubmed",
    "openlibrary",
    "arxiv",
    "clinicaltrials",
];
function isRecord(value) {
    return typeof value === "object" && value !== null && !Array.isArray(value);
}
function asString(value) {
    return typeof value === "string" && value.trim() !== "" ? value.trim() : undefined;
}
/** Provider abstracts are often JATS/XML fragments; keep the raw response in extensions but make
 * the canonical field readable in the desktop detail pane. */
function readableMetadataText(value) {
    const text = asString(value);
    if (text === undefined)
        return undefined;
    const readable = text
        .replace(/<!\[CDATA\[/gu, "")
        .replace(/\]\]>/gu, "")
        .replace(/<[^>]+>/gu, " ")
        .replace(/&amp;/gu, "&")
        .replace(/&lt;/gu, "<")
        .replace(/&gt;/gu, ">")
        .replace(/&quot;/gu, '"')
        .replace(/&apos;/gu, "'")
        .replace(/\s+/gu, " ")
        .trim();
    return readable === "" ? undefined : readable;
}
function asNumber(value) {
    return typeof value === "number" && Number.isFinite(value) ? value : undefined;
}
function firstString(value) {
    if (!Array.isArray(value))
        return asString(value);
    return value.map(asString).find((entry) => entry !== undefined);
}
/**
 * A canonical `issuedYear` is a non-negative safe integer, and nothing else.
 *
 * `Number.isSafeInteger` rather than `Number.isInteger` because a provider year outside the
 * safe range has already lost digits to `JSON.parse` before this function sees it, so promoting
 * it would publish a number no one supplied (the house rule; cf. `field-registry.ts`).
 *
 * A `YYYY`-prefixed date string is accepted because every provider that returns a full date also
 * returns it as the first four characters of that date, and three adapters already sliced it by
 * hand. Anything else is not a year: it is left out of the canonical field and stays readable in
 * `issuedRaw` and in the source-namespaced raw response, so nothing is lost (invariant 2).
 */
function integerYear(value) {
    const number = asNumber(value)
        ?? (typeof value === "string" && /^\d{4}(?:$|[^\d])/u.test(value) ? Number(value.slice(0, 4)) : undefined);
    return number !== undefined && Number.isSafeInteger(number) && number >= 0 ? number : undefined;
}
function json(value) {
    if (value === null || typeof value === "string" || typeof value === "boolean")
        return value;
    if (typeof value === "number" && Number.isFinite(value))
        return value;
    if (Array.isArray(value)) {
        const values = value.map(json);
        return values.every((entry) => entry !== undefined) ? values : undefined;
    }
    if (isRecord(value)) {
        const output = {};
        for (const [key, entry] of Object.entries(value)) {
            const serialised = json(entry);
            if (serialised === undefined)
                return undefined;
            output[key] = serialised;
        }
        return output;
    }
    return undefined;
}
function creators(value) {
    if (!Array.isArray(value))
        return undefined;
    const result = [];
    for (const [ordinal, raw] of value.entries()) {
        if (!isRecord(raw))
            continue;
        const family = asString(raw.family ?? raw.lastName ?? raw.surname);
        const given = asString(raw.given ?? raw.firstName);
        // An ORCID is not a name. It used to be the last fallback here, which gave a DataCite
        // creator carrying only an identifier an author literally called
        // "https://orcid.org/0000-0002-1825-0097" — a value that then renders in a citation. A
        // creator this provider did not name is skipped below; the row itself survives verbatim in
        // the source-namespaced raw response (invariant 2), and the reviewer raises the missing
        // required field rather than the product inventing one (invariant 3).
        const literal = asString(raw.literal ?? raw.name);
        if (family === undefined && given === undefined && literal === undefined)
            continue;
        // Crossref spells it `ORCID`; DataCite and the others spell it `orcid`. Reading only one
        // spelling dropped the identifier from every Crossref creator.
        const orcid = asString(raw.orcid ?? raw.ORCID);
        result.push({
            ordinal,
            role: "author",
            ...(family === undefined ? {} : { family }),
            ...(given === undefined ? {} : { given }),
            ...(literal !== undefined && family === undefined && given === undefined ? { literal } : {}),
            ...(orcid === undefined ? {} : { orcid }),
        });
    }
    return result.length === 0 ? undefined : result;
}
/**
 * The one place a canonical field path is published, and therefore the one place its canonical
 * type is enforced.
 *
 * `issuedYear` is `number` on `ReferenceItem` and `validateReferenceItem` requires a non-negative
 * integer, but two adapters published `published.slice(0, 4)` — the *string* `"2024"`. Nothing
 * downstream accepts a string year: `materialiseReference` sends it to the unmapped extension,
 * the desktop import review drops it from the record summary, `workKey` stops matching an arXiv
 * preprint to its published DOI, and `reviewImportRecord` reports an ordinary 2024 preprint as a
 * publication year "outside the range this library treats as plausible". Every one of those is a
 * consequence of a type, not of the data, so the type is fixed here rather than re-checked at
 * each of the four call sites.
 *
 * A value that cannot be a year is left out of the canonical field rather than coerced. It stays
 * readable in `issuedRaw` and in the complete source-namespaced response, which is what makes
 * this a projection rule and not data loss (invariant 2).
 */
function canonicalFieldValue(path, value) {
    if (path === "issuedYear")
        return value === undefined ? undefined : integerYear(value);
    return json(value);
}
function candidate(source, identifier, sourceId, retrievedAt, confidence, fields, rawResponse) {
    const cleanFields = {};
    for (const [path, value] of Object.entries(fields)) {
        const serialised = canonicalFieldValue(path, value);
        if (serialised !== undefined)
            cleanFields[path] = serialised;
    }
    const provenance = {};
    for (const path of Object.keys(cleanFields)) {
        provenance[path] = { source, retrievedAt, confidence };
    }
    const raw = json(rawResponse);
    const rawNamespace = isRecord(raw)
        ? raw
        : { rawRecord: raw ?? null };
    return {
        source,
        sourceId,
        identifier,
        retrievedAt,
        confidence,
        fields: cleanFields,
        provenance,
        extensions: { [source]: rawNamespace },
    };
}
/**
 * The headers one source is given.
 *
 * `sourceId` is what makes the contact e-mail a per-provider disclosure (E04-15.1): with
 * `contactEmailSources` omitted every source that is asked sees the address, which is what this
 * function has always done; supplied, the others are asked anonymously.
 */
function headers(options, sourceId) {
    const userAgent = options.userAgent?.trim() || "ReferenceManager/0.1";
    const email = options.contactEmail?.trim();
    const discloses = email !== undefined &&
        email !== "" &&
        (options.contactEmailSources === undefined || options.contactEmailSources.includes(sourceId));
    return {
        Accept: "application/json, application/atom+xml, application/xml;q=0.9, text/xml;q=0.8",
        "User-Agent": discloses ? `${userAgent} (mailto:${email})` : userAgent,
    };
}
function encoded(value) {
    return encodeURIComponent(value);
}
async function bodyJson(response) {
    if (response.json === undefined)
        throw new SourceResponseError("invalid-response");
    return response.json();
}
async function bodyText(response) {
    if (response.text === undefined)
        throw new SourceResponseError("invalid-response");
    return response.text();
}
function responseFailure(response, source) {
    if (response.status === 404)
        return "not-found";
    return response.status >= 500 || response.status === 429 ? "network" : source === "arxiv" ? "invalid-response" : "http";
}
function sourceHeaders(context) {
    return context.headers;
}
class JsonSource {
    async lookup(identifier, context) {
        const request = {
            url: this.urlFor(identifier),
            headers: sourceHeaders(context),
            ...(context.signal === undefined ? {} : { signal: context.signal }),
        };
        const response = await context.transport.request(request);
        if (response.status < 200 || response.status >= 300) {
            const failure = responseFailure(response, this.id);
            throw new SourceResponseError(failure, response.status);
        }
        const payload = await bodyJson(response);
        return this.parse(identifier, payload, context.retrievedAt);
    }
}
export class SourceResponseError extends Error {
    code;
    status;
    constructor(code, status) {
        super(`metadata source response was ${code}`);
        this.name = "SourceResponseError";
        this.code = code;
        if (status !== undefined)
            this.status = status;
    }
}
export class CrossrefSource extends JsonSource {
    id = "crossref";
    supportedKinds = ["doi"];
    urlFor(identifier) {
        return `https://api.crossref.org/works/${encoded(identifier.value)}`;
    }
    parse(identifier, payload, retrievedAt) {
        const message = isRecord(payload) && isRecord(payload.message) ? payload.message : undefined;
        if (message === undefined)
            throw new SourceResponseError("invalid-response");
        const dateParts = isRecord(message.published) ? message.published["date-parts"] : undefined;
        const authorRows = Array.isArray(message.author) ? message.author : [];
        return candidate(this.id, identifier, asString(message.DOI) ?? identifier.value, retrievedAt, 0.9, {
            type: message.type === "book-chapter" ? "chapter" : message.type === "book" ? "book" : "article-journal",
            title: firstString(message.title),
            containerTitle: firstString(message["container-title"]),
            issuedYear: Array.isArray(dateParts) && Array.isArray(dateParts[0]) ? integerYear(dateParts[0][0]) : undefined,
            doi: asString(message.DOI) ?? identifier.value,
            creators: creators(authorRows),
            "fields.volume": asString(message.volume),
            "fields.issue": asString(message.issue),
            "fields.pages": asString(message.page),
            "fields.publisher": asString(message.publisher),
            "fields.url": asString(message.URL),
            "fields.abstract": readableMetadataText(message.abstract),
        }, payload);
    }
}
export class DataCiteSource extends JsonSource {
    id = "datacite";
    supportedKinds = ["doi"];
    urlFor(identifier) {
        return `https://api.datacite.org/dois/${encoded(identifier.value)}`;
    }
    parse(identifier, payload, retrievedAt) {
        const data = isRecord(payload) && isRecord(payload.data) ? payload.data : undefined;
        const attributes = data !== undefined && isRecord(data.attributes) ? data.attributes : undefined;
        if (attributes === undefined)
            throw new SourceResponseError("invalid-response");
        const titles = Array.isArray(attributes.titles) ? attributes.titles : [];
        const container = isRecord(attributes.container) ? attributes.container : undefined;
        const creatorsData = Array.isArray(attributes.creators)
            ? attributes.creators.map((entry) => isRecord(entry) ? {
                family: entry.familyName,
                given: entry.givenName,
                name: entry.name,
                orcid: entry.nameIdentifiers && Array.isArray(entry.nameIdentifiers)
                    ? entry.nameIdentifiers.find((id) => isRecord(id) && id.nameIdentifierScheme === "ORCID")?.nameIdentifier
                    : undefined,
            } : {})
            : [];
        const descriptions = Array.isArray(attributes.descriptions) ? attributes.descriptions : [];
        const abstract = descriptions.find((entry) => isRecord(entry) && entry.descriptionType === "Abstract");
        return candidate(this.id, identifier, asString(data?.id) ?? identifier.value, retrievedAt, 0.85, {
            type: attributes.types && isRecord(attributes.types) && attributes.types.resourceTypeGeneral === "Book" ? "book" : "article-journal",
            title: titles.length > 0 && isRecord(titles[0]) ? titles[0].title : undefined,
            containerTitle: container?.title,
            issuedYear: integerYear(attributes.publicationYear),
            doi: asString(attributes.doi) ?? identifier.value,
            creators: creators(creatorsData),
            "fields.publisher": asString(attributes.publisher),
            "fields.url": asString(attributes.url),
            "fields.abstract": isRecord(abstract) ? readableMetadataText(abstract.description) : undefined,
        }, payload);
    }
}
export class PubMedSource extends JsonSource {
    id = "pubmed";
    supportedKinds = ["pmid", "pmcid"];
    urlFor(identifier) {
        const db = identifier.kind === "pmcid" ? "pmc" : "pubmed";
        return `https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esummary.fcgi?db=${db}&id=${encoded(identifier.value)}&retmode=json`;
    }
    parse(identifier, payload, retrievedAt) {
        const result = isRecord(payload) && isRecord(payload.result) ? payload.result : undefined;
        if (result === undefined)
            throw new SourceResponseError("invalid-response");
        const row = result[identifier.value];
        if (!isRecord(row))
            return undefined;
        // E-utilities returns a row containing only the requested uid and an error for
        // unknown PMIDs/PMCIDs. Treat that as a not-found response instead of projecting
        // an empty "Untitled reference" into the import review.
        if (asString(row.error) !== undefined)
            return undefined;
        const articleIds = Array.isArray(row.articleids) ? row.articleids : [];
        const doi = articleIds.find((entry) => isRecord(entry) && entry.idtype === "doi");
        const authors = Array.isArray(row.authors) ? row.authors.map((entry) => isRecord(entry) ? { name: entry.name } : {}) : [];
        const pubdate = asString(row.pubdate);
        return candidate(this.id, identifier, identifier.value, retrievedAt, 0.95, {
            type: "article-journal",
            title: asString(row.title),
            containerTitle: asString(row.fulljournalname ?? row.source),
            issuedYear: pubdate,
            issuedRaw: pubdate,
            pmid: identifier.kind === "pmid" ? identifier.value : undefined,
            "fields.pmcid": identifier.kind === "pmcid" ? identifier.value : undefined,
            doi: isRecord(doi) ? asString(doi.value) : undefined,
            creators: creators(authors),
            "fields.volume": asString(row.volume),
            "fields.issue": asString(row.issue),
            "fields.pages": asString(row.pages),
        }, payload);
    }
}
export class OpenLibrarySource extends JsonSource {
    id = "openlibrary";
    supportedKinds = ["isbn"];
    urlFor(identifier) {
        return `https://openlibrary.org/isbn/${encoded(identifier.value)}.json`;
    }
    parse(identifier, payload, retrievedAt) {
        if (!isRecord(payload))
            throw new SourceResponseError("invalid-response");
        // `/isbn/{isbn}.json` returns author *keys* — `{ key: "/authors/OL34184A" }` — and only the
        // works endpoint returns names. Reading `key` as a name put "/authors/OL34184A" in the author
        // field of every book imported by ISBN, and that is what a citation would then print. A book
        // whose authors this response only points at is imported with no creators, the pointers
        // survive in `extensions.openlibrary`, and the reviewer asks the user (invariants 2 and 3).
        const authorRows = Array.isArray(payload.authors)
            ? payload.authors.map((entry) => (isRecord(entry) ? { name: entry.name } : {}))
            : [];
        return candidate(this.id, identifier, asString(payload.key) ?? identifier.value, retrievedAt, 0.75, {
            type: "book",
            title: asString(payload.title),
            issuedYear: asString(payload.publish_date)?.match(/\b(\d{4})\b/u)?.[1] === undefined
                ? undefined
                : integerYear(asString(payload.publish_date)?.match(/\b(\d{4})\b/u)?.[1]),
            issuedRaw: asString(payload.publish_date),
            creators: creators(authorRows),
            "fields.subtitle": asString(payload.subtitle),
            "fields.publisher": firstString(payload.publishers),
            "fields.pages": asNumber(payload.number_of_pages)?.toString(),
            "fields.isbn": identifier.value,
        }, payload);
    }
}
function xmlInner(xml, tag) {
    const pattern = new RegExp(`<${tag}(?:\\s[^>]*)?>([\\s\\S]*?)</${tag}>`, "iu");
    const match = pattern.exec(xml);
    if (match === null)
        return undefined;
    return match[1];
}
function xmlText(xml, tag) {
    return xmlInner(xml, tag)?.replace(/<!\[CDATA\[([\s\S]*?)\]\]>/gu, "$1").replace(/<[^>]+>/gu, "").trim();
}
function xmlEntries(xml, tag) {
    const pattern = new RegExp(`<${tag}(?:\\s[^>]*)?>([\\s\\S]*?)</${tag}>`, "giu");
    return [...xml.matchAll(pattern)].map((match) => match[1]?.replace(/<!\[CDATA\[([\s\S]*?)\]\]>/gu, "$1").replace(/<[^>]+>/gu, "").trim()).filter((value) => value !== undefined && value !== "");
}
export class ArxivSource {
    id = "arxiv";
    supportedKinds = ["arxiv"];
    urlFor(identifier) {
        return `https://export.arxiv.org/api/query?id_list=${encoded(identifier.value)}`;
    }
    async lookup(identifier, context) {
        const request = {
            url: this.urlFor(identifier),
            headers: sourceHeaders(context),
            ...(context.signal === undefined ? {} : { signal: context.signal }),
        };
        const response = await context.transport.request(request);
        if (response.status < 200 || response.status >= 300) {
            const failure = responseFailure(response, this.id);
            throw new SourceResponseError(failure, response.status);
        }
        const xml = await bodyText(response);
        if (/<(?:!DOCTYPE|!ENTITY)\b/iu.test(xml))
            throw new SourceResponseError("invalid-response");
        const entry = xmlInner(xml, "entry");
        if (entry === undefined)
            return undefined;
        const authors = [...entry.matchAll(/<author\b[^>]*>([\s\S]*?)<\/author>/giu)].map((match) => ({
            name: xmlText(match[1] ?? "", "name"),
        }));
        const published = xmlText(entry, "published");
        return candidate(this.id, identifier, xmlText(entry, "id") ?? identifier.value, context.retrievedAt, 0.8, {
            type: "article-journal",
            title: xmlText(entry, "title")?.replace(/\s+/gu, " "),
            issuedYear: published,
            issuedRaw: published,
            creators: creators(authors),
            "fields.arxivId": identifier.value,
            "fields.abstract": xmlText(entry, "summary")?.replace(/\s+/gu, " "),
            "fields.keywords": xmlEntries(entry, "category"),
            "fields.url": xmlText(entry, "id"),
        }, xml);
    }
}
export class ClinicalTrialsSource extends JsonSource {
    id = "clinicaltrials";
    supportedKinds = ["nct"];
    urlFor(identifier) {
        return `https://clinicaltrials.gov/api/v2/studies/${encoded(identifier.value)}`;
    }
    parse(identifier, payload, retrievedAt) {
        if (!isRecord(payload) || !isRecord(payload.protocolSection))
            throw new SourceResponseError("invalid-response");
        const protocol = payload.protocolSection;
        const identification = isRecord(protocol.identificationModule) ? protocol.identificationModule : {};
        const status = isRecord(protocol.statusModule) ? protocol.statusModule : {};
        const design = isRecord(protocol.designModule) ? protocol.designModule : {};
        const sponsor = isRecord(protocol.sponsorCollaboratorsModule) ? protocol.sponsorCollaboratorsModule : {};
        const start = asString(status.studyFirstSubmitDate) ?? asString(status.studyFirstSubmitQcDate);
        return candidate(this.id, identifier, identifier.value, retrievedAt, 0.9, {
            type: "article-journal",
            title: asString(identification.officialTitle) ?? asString(identification.briefTitle),
            issuedYear: start,
            issuedRaw: start,
            creators: asString(sponsor.leadSponsor) === undefined ? undefined : [{ ordinal: 0, role: "author", literal: asString(sponsor.leadSponsor) }],
            "fields.clinicalTrialId": identifier.value,
            "fields.publicationStatus": asString(status.overallStatus),
            "fields.extra": Array.isArray(protocol.conditionsModule && isRecord(protocol.conditionsModule) ? protocol.conditionsModule.conditions : undefined)
                ? protocol.conditionsModule.conditions
                : undefined,
            "fields.url": `https://clinicaltrials.gov/study/${identifier.value}`,
            "fields.publisher": asString(sponsor.leadSponsor),
            "fields.keywords": Array.isArray(design.phases) ? design.phases : undefined,
        }, payload);
    }
}
export function createMetadataSources() {
    return [new CrossrefSource(), new DataCiteSource(), new PubMedSource(), new OpenLibrarySource(), new ArxivSource(), new ClinicalTrialsSource()];
}
export class FetchMetadataTransport {
    async request(request) {
        const response = await fetch(request.url, {
            ...(request.headers === undefined ? {} : { headers: request.headers }),
            ...(request.signal === undefined ? {} : { signal: request.signal }),
        });
        return {
            status: response.status,
            json: () => response.json(),
            text: () => response.text(),
        };
    }
}
function cacheKey(source, identifier) {
    return `${source.id}:${identifier.kind}:${identifier.key}`;
}
function failure(source, identifier, code, message, retryable, status) {
    return { source: source.id, identifier, code, message, retryable, ...(status === undefined ? {} : { status }) };
}
function failureFromError(source, identifier, error) {
    if (error instanceof SourceResponseError) {
        return failure(source, identifier, error.code, `provider ${source.id} returned ${error.code}`, error.code === "network", error.status);
    }
    if (error instanceof Error && error.name === "AbortError") {
        return failure(source, identifier, "cancelled", "metadata lookup was cancelled", true);
    }
    return failure(source, identifier, "network", "metadata provider could not be reached", true);
}
/**
 * Retrieve independent provider candidates for a recognised batch. Results are stable in batch
 * and source order. No candidate is merged, and no input line is turned into a record here.
 */
export async function retrieveMetadata(batch, options = {}) {
    const identifiers = "identifiers" in batch ? batch.identifiers : batch;
    const sources = options.sources ?? createMetadataSources();
    const transport = options.transport ?? (options.networkPolicy === "enabled" ? new FetchMetadataTransport() : undefined);
    const enabled = transport !== undefined && options.networkPolicy !== "disabled";
    const now = options.now ?? (() => new Date().toISOString());
    const records = [];
    const failures = [];
    let cacheHits = 0;
    for (const identifier of identifiers) {
        let matchedSource = false;
        for (const source of sources) {
            if (!source.supportedKinds.includes(identifier.kind))
                continue;
            matchedSource = true;
            if (!enabled || transport === undefined) {
                failures.push(failure(source, identifier, "offline", "network retrieval is disabled; identifier remains retryable", true));
                continue;
            }
            const key = cacheKey(source, identifier);
            const cached = await options.cache?.get(key);
            if (cached !== undefined) {
                records.push(cached);
                cacheHits += 1;
                continue;
            }
            try {
                const context = {
                    transport,
                    headers: headers(options, source.id),
                    retrievedAt: now(),
                    ...(options.signal === undefined ? {} : { signal: options.signal }),
                };
                const record = await source.lookup(identifier, context);
                if (record === undefined) {
                    failures.push(failure(source, identifier, "not-found", "provider returned no matching record", false));
                }
                else {
                    records.push(record);
                    await options.cache?.set(key, record);
                }
            }
            catch (error) {
                failures.push(failureFromError(source, identifier, error));
            }
        }
        if (!matchedSource) {
            failures.push({
                source: "retrieval",
                identifier,
                code: "unsupported",
                message: `no metadata source is registered for ${identifier.kind}`,
                retryable: false,
            });
        }
    }
    return { records, failures, cacheHits };
}
