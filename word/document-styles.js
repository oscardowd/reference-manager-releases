/**
 * CSL styles and locales for document mode (ADR-0242 decision 4).
 *
 * Without the desktop the pane has no installed styles. It fetches the ones a request needs from
 * the official CSL repositories, at the pinned revisions `csl-catalog.json` records, and installs a
 * file only if its bytes hash to the SHA-256 committed there — a changed, truncated or substituted
 * file is refused, never used. What was fetched is kept in the pane's storage under its hash, so a
 * style needs the network once per device, and is re-checked every time it is read back.
 *
 * A style the catalog does not list cannot be obtained here, and a request for one is refused by
 * name. It is never swapped for another style (invariant 3, §9.3).
 *
 * The request reveals the style's file name to GitHub, which serves the CSL repositories; nothing
 * about the document or its references is sent. The pane says so before the first download.
 */

/** New documents start in APA, as the desktop's session does when it has APA installed. */
export const DEFAULT_STYLE_ID = "http://www.zotero.org/styles/apa";
export const DEFAULT_LOCALE = "en-US";

/**
 * How long a style, locale or catalog request may take before the pane gives up on it. Fifteen
 * seconds, the figure the pane's desktop client already uses: chosen, not measured, and generous for
 * a file of tens of kilobytes. A server that accepts and never answers must not hold the pane.
 */
export const STYLE_FETCH_TIMEOUT_MS = 15_000;

/**
 * One request with a deadline. `fetchImpl` is injected by tests; the page's own `fetch` otherwise.
 * An abandoned request rejects, and the caller reports it like any other failure to download.
 */
export async function boundedFetch(injected, url, init = {}, timeoutMs = STYLE_FETCH_TIMEOUT_MS) {
  const fetchImpl = injected ?? ((target, options) => globalThis.fetch(target, options));
  const controller = new AbortController();
  const timer = globalThis.setTimeout(() => controller.abort(), timeoutMs);
  try {
    return await fetchImpl(url, { ...init, signal: controller.signal });
  } finally {
    globalThis.clearTimeout(timer);
  }
}

export class StyleUnavailableError extends Error {
  constructor(message) {
    super(message);
    this.name = "StyleUnavailableError";
    this.code = "style-unavailable";
  }
}

/** Lowercase hex SHA-256 of UTF-8 text, through the browser's own digest. */
export async function sha256Hex(text, subtle = globalThis.crypto && globalThis.crypto.subtle) {
  if (!subtle || typeof subtle.digest !== "function") throw new StyleUnavailableError("This pane cannot check downloaded styles here.");
  const digest = new Uint8Array(await subtle.digest("SHA-256", new globalThis.TextEncoder().encode(text)));
  let hex = "";
  for (const byte of digest) hex += byte.toString(16).padStart(2, "0");
  return hex;
}

function hasStyle(registry, styleId) {
  return registry.listStyles().some((style) => style.metadata.id === styleId);
}

function hasLocale(registry, lang) {
  return registry.listLocales().some((locale) => locale.metadata.id === lang);
}

/** The catalog locale for a language tag: the exact one, or one of the same language. */
export function catalogLocaleFor(catalog, lang) {
  const entries = catalog.locales.entries;
  const exact = entries.find((entry) => entry.lang === lang);
  if (exact) return exact;
  const language = String(lang).split("-")[0].toLowerCase();
  return entries.find((entry) => entry.lang.split("-")[0].toLowerCase() === language) ?? null;
}

export function createStyleSource({ catalog, storage, fetchImpl, subtle, timeoutMs = STYLE_FETCH_TIMEOUT_MS } = {}) {
  const styles = new Map(catalog.styles.entries.map((entry) => [entry.id, entry]));

  async function obtain(entry, base, what) {
    const key = `refmgr:csl:${entry.sha256}`;
    const cached = await storage.get(key);
    if (typeof cached === "string" && (await sha256Hex(cached, subtle)) === entry.sha256) return cached;
    let response;
    try {
      response = await boundedFetch(fetchImpl, `${base}${entry.file}`, { credentials: "omit", referrerPolicy: "no-referrer", cache: "no-cache" }, timeoutMs);
    } catch {
      throw new StyleUnavailableError(`${what} could not be downloaded. Connect to the internet once, or open Reference Manager on this computer.`);
    }
    if (!response || !response.ok) {
      throw new StyleUnavailableError(`${what} could not be downloaded. Try again later, or open Reference Manager on this computer.`);
    }
    const xml = await response.text();
    if ((await sha256Hex(xml, subtle)) !== entry.sha256) {
      throw new StyleUnavailableError(`${what} did not match the published version, so it was not used.`);
    }
    await storage.set(key, xml);
    return xml;
  }

  async function ensureStyle(registry, styleId) {
    if (hasStyle(registry, styleId)) return;
    const entry = styles.get(styleId);
    if (!entry) {
      throw new StyleUnavailableError("This document's citation style is only available with Reference Manager on this computer. Open Reference Manager, then choose Reconnect.");
    }
    registry.installStyle(await obtain(entry, catalog.styles.base, `The ${entry.title} style`), { origin: "user" });
  }

  async function ensureLocale(registry, lang) {
    const entry = catalogLocaleFor(catalog, lang);
    // A language the catalog lacks is not guessed at: the session refuses it by name.
    if (!entry || hasLocale(registry, entry.lang)) return;
    registry.installLocale(await obtain(entry, catalog.locales.base, "This document's language"), { origin: "user" });
  }

  return Object.freeze({
    /** The styles a document may be switched to in document mode, as `/styles` lists them. */
    catalogStyles() {
      return catalog.styles.entries.map((entry) => ({
        id: entry.id,
        title: entry.title,
        ...(entry.citationFormat === undefined ? {} : { citationFormat: entry.citationFormat }),
      }));
    },
    ensureStyle,
    ensureLocale,
    /**
     * Install what one session request will render with: the style and locale it names, and the
     * ones the document's own citation data names. A request that names none — the opening
     * `/session` of a new document — gets the default, so there is always a style to start in.
     */
    async prepare(registry, payload, requirements) {
      const body = payload && typeof payload.body === "object" && payload.body !== null ? payload.body : {};
      const styleIds = new Set(requirements.styleIds);
      const locales = new Set(requirements.locales);
      if (typeof body.styleId === "string" && body.styleId.length > 0) styleIds.add(body.styleId);
      if (typeof body.locale === "string" && body.locale.length > 0) locales.add(body.locale);
      if (styleIds.size === 0) styleIds.add(DEFAULT_STYLE_ID);
      if (locales.size === 0) locales.add(DEFAULT_LOCALE);
      for (const styleId of styleIds) await ensureStyle(registry, styleId);
      for (const lang of locales) await ensureLocale(registry, lang);
    },
  });
}
