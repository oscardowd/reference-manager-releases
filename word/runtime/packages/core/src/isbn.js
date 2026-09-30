/**
 * ISBN normalisation and check-digit verification (SPEC §5.1, §6.5; tasks E04-08.1, E07-01).
 *
 * This lives in `core` beside the other identifier formats because two packages need it and
 * neither may import the other: `identifiers` recognises a pasted ISBN, and duplicate detection
 * — which `db`'s merge engine consumes — compares two library records by it. `@refmgr/identifiers`
 * re-exports the same functions, so its public surface is unchanged.
 *
 * The check digit is verified rather than assumed. A mistyped ISBN is the one identifier class
 * where a single wrong character silently names a different, real book, so recognising an
 * ISBN-shaped string without checking it would import the wrong record and look successful.
 * `recognize.ts` therefore reports a failed check digit as an explicit reason rather than
 * dropping the input or accepting it.
 */
/** Digits, hyphens and spaces, with `X` allowed only as the ISBN-10 check character. */
const ISBN_CHARACTERS = /^[0-9Xx-‐‑‒–— ]+$/u;
/** Strip the separators an ISBN may be written with and upper-case a trailing `x`. */
export function normalizeIsbn(value) {
    return value.replace(/[-‐‑‒–— ]/gu, "").toUpperCase();
}
/** True when `value` uses only characters an ISBN may contain. Length is not checked. */
export function hasIsbnCharacters(value) {
    return ISBN_CHARACTERS.test(value);
}
function digitValues(compact, allowX) {
    const values = [];
    for (let index = 0; index < compact.length; index += 1) {
        const character = compact[index];
        if (character >= "0" && character <= "9") {
            values.push(character.charCodeAt(0) - 48);
            continue;
        }
        // `X` is the value ten, and only ever the last character of an ISBN-10.
        if (allowX && character === "X" && index === compact.length - 1) {
            values.push(10);
            continue;
        }
        return undefined;
    }
    return values;
}
/** True when `value` is ten characters with a valid mod-11 check digit. */
export function isValidIsbn10(value) {
    const compact = normalizeIsbn(value);
    if (compact.length !== 10)
        return false;
    const digits = digitValues(compact, true);
    if (digits === undefined)
        return false;
    let sum = 0;
    for (let index = 0; index < 10; index += 1) {
        sum += digits[index] * (10 - index);
    }
    return sum % 11 === 0;
}
/** True when `value` is thirteen digits in a real Bookland prefix with a valid mod-10 check digit. */
export function isValidIsbn13(value) {
    const compact = normalizeIsbn(value);
    if (compact.length !== 13)
        return false;
    if (!compact.startsWith("978") && !compact.startsWith("979"))
        return false;
    const digits = digitValues(compact, false);
    if (digits === undefined)
        return false;
    let sum = 0;
    for (let index = 0; index < 13; index += 1) {
        sum += digits[index] * (index % 2 === 0 ? 1 : 3);
    }
    return sum % 10 === 0;
}
/** True when `value` is a valid ISBN in either form. */
export function isValidIsbn(value) {
    return isValidIsbn10(value) || isValidIsbn13(value);
}
/**
 * Widen a valid ISBN-10 to its ISBN-13 equivalent. Returns `undefined` for anything else.
 *
 * The two forms name the same book, so comparison happens in ISBN-13 space; the form the user
 * actually typed is what gets stored, because that is what is printed on their copy.
 */
export function isbn10ToIsbn13(value) {
    if (!isValidIsbn10(value))
        return undefined;
    const body = `978${normalizeIsbn(value).slice(0, 9)}`;
    let sum = 0;
    for (let index = 0; index < 12; index += 1) {
        sum += (body[index].charCodeAt(0) - 48) * (index % 2 === 0 ? 1 : 3);
    }
    return `${body}${(10 - (sum % 10)) % 10}`;
}
/** The ISBN-13 form of any valid ISBN, used only for comparison. */
export function isbnComparisonKey(value) {
    const compact = normalizeIsbn(value);
    if (isValidIsbn13(compact))
        return compact;
    return isbn10ToIsbn13(compact);
}
