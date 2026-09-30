/**
 * Stable and human-facing identifiers (task E01-04; SPEC §2; ADR-0009).
 *
 * UUIDv7 is the only authoritative identity. Citation keys are deliberately derived suggestions:
 * callers may edit them, imports may preserve a source key, and no lookup should depend on them.
 */
const MAX_UUID_TIMESTAMP = 0xffffffffffff;
const UUID_V7 = /^[0-9a-f]{8}-[0-9a-f]{4}-7[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i;
function defaultRandomFill(bytes) {
    globalThis.crypto.getRandomValues(bytes);
}
function formatUuid(bytes) {
    const hex = Array.from(bytes, (byte) => byte.toString(16).padStart(2, "0")).join("");
    return `${hex.slice(0, 8)}-${hex.slice(8, 12)}-${hex.slice(12, 16)}-${hex.slice(16, 20)}-${hex.slice(20)}`;
}
/**
 * Create a process-local monotonic UUIDv7 generator.
 *
 * The 48-bit millisecond timestamp is followed by a 12-bit sequence. The sequence is seeded
 * randomly for a new millisecond and incremented for another id in the same millisecond. This
 * preserves UUIDv7's natural byte/lexical ordering even during a burst, and a backwards wall
 * clock cannot make a newly generated id sort before one this generator already returned.
 */
export function createUuidV7Generator(options = {}) {
    const now = options.now ?? Date.now;
    const randomFill = options.randomFill ?? defaultRandomFill;
    let lastTimestamp = -1;
    let sequence = -1;
    return () => {
        const observed = now();
        if (!Number.isSafeInteger(observed) || observed < 0 || observed > MAX_UUID_TIMESTAMP) {
            throw new RangeError("UUIDv7 time must be a non-negative integer Unix millisecond value.");
        }
        const bytes = new Uint8Array(16);
        randomFill(bytes);
        let timestamp = Math.max(observed, lastTimestamp);
        if (timestamp > lastTimestamp) {
            sequence = ((bytes[6] & 0x0f) << 8) | bytes[7];
        }
        else if (sequence < 0x0fff) {
            sequence += 1;
        }
        else {
            timestamp += 1;
            if (timestamp > MAX_UUID_TIMESTAMP) {
                throw new RangeError("UUIDv7 timestamp space is exhausted.");
            }
            sequence = 0;
        }
        lastTimestamp = timestamp;
        let remaining = timestamp;
        for (let index = 5; index >= 0; index -= 1) {
            bytes[index] = remaining % 256;
            remaining = Math.floor(remaining / 256);
        }
        bytes[6] = 0x70 | (sequence >> 8);
        bytes[7] = sequence & 0xff;
        bytes[8] = (bytes[8] & 0x3f) | 0x80;
        return formatUuid(bytes);
    };
}
const defaultUuidV7Generator = createUuidV7Generator();
/** Generate one UUIDv7 using the process-local monotonic generator. */
export function generateUuidV7() {
    return defaultUuidV7Generator();
}
/** Read the embedded Unix millisecond timestamp, refusing non-UUIDv7 input. */
export function uuidV7Timestamp(id) {
    if (!UUID_V7.test(id))
        throw new TypeError("Expected a canonical UUIDv7 string.");
    return Number.parseInt(id.replaceAll("-", "").slice(0, 12), 16);
}
const TITLE_STOP_WORDS = new Set([
    "a",
    "an",
    "and",
    "for",
    "from",
    "in",
    "of",
    "on",
    "the",
    "to",
    "with",
]);
function keyPart(value) {
    return value
        .normalize("NFKD")
        .replace(/\p{Mark}/gu, "")
        .toLocaleLowerCase("en")
        .replace(/[^a-z0-9]+/g, "");
}
function firstAuthor(item) {
    return item.creators
        .filter((creator) => creator.role === "author")
        .sort((left, right) => left.ordinal - right.ordinal)[0];
}
function titleWord(title) {
    if (title === undefined)
        return "untitled";
    const words = title
        .normalize("NFKD")
        .replace(/\p{Mark}/gu, "")
        .toLocaleLowerCase("en")
        .match(/[\p{Letter}\p{Number}]+/gu) ?? [];
    const chosen = words.find((word) => !TITLE_STOP_WORDS.has(word)) ?? words[0];
    return chosen === undefined ? "untitled" : keyPart(chosen) || "untitled";
}
function alphaSuffix(index) {
    let remaining = index;
    let suffix = "";
    do {
        suffix = String.fromCharCode(97 + (remaining % 26)) + suffix;
        remaining = Math.floor(remaining / 26) - 1;
    } while (remaining >= 0);
    return suffix;
}
/**
 * Suggest a readable citation key such as `smith2024effects`.
 *
 * The first author, year and first meaningful title word are stable inputs. Missing or
 * non-Latin components receive explicit fallbacks, and collisions gain deterministic alphabetic
 * suffixes. The function never mutates the item and the result remains non-authoritative.
 */
export function generateCitationKey(item, options = {}) {
    const author = firstAuthor(item);
    const authorPart = keyPart(author?.family ?? author?.literal ?? "") || "anon";
    const yearPart = item.issuedYear === undefined ? "nd" : String(item.issuedYear);
    const base = `${authorPart}${yearPart}${titleWord(item.title)}`;
    const reserved = new Set(Array.from(options.existingKeys ?? [], (key) => key.toLocaleLowerCase("en")));
    if (!reserved.has(base.toLocaleLowerCase("en")))
        return base;
    for (let index = 0;; index += 1) {
        const candidate = `${base}${alphaSuffix(index)}`;
        if (!reserved.has(candidate.toLocaleLowerCase("en")))
            return candidate;
    }
}
