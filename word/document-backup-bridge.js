/**
 * Browser-side adapter for the document-backup routes (SPEC §9.5; ADR-0115/0116; E10-11.2).
 *
 * The task pane owns transport, not backup policy. It declares the document length, follows the
 * slice size the desktop bridge returns, commits, then reads the copy back slice by slice. The
 * domain layer performs the final byte comparison before allowing a repair to write.
 *
 * **Every request here is bounded (R-098), and the deadline is per request rather than per
 * operation.** A refused connection fails in microseconds; what this guards against is a loopback
 * port that **accepts and then holds** — another program on it, a wedged desktop process, a proxy
 * that stalls — which an unbounded `fetch` waits out for as long as the browser will, with the
 * repair the user asked for waiting behind it. One deadline over the whole operation could not
 * work here: a store-and-verify is `2 + 2n` requests (ADR-0116), 34 of them at the 64 MiB ceiling,
 * so a number short enough to mean anything for one slice would abandon a large document, and one
 * long enough for a large document would let a stalled peer hold a small one for just as long.
 * The bound on a whole operation is therefore the sum, and it is stated rather than hidden: at the
 * ceiling, 33 × {@link DOCUMENT_BACKUP_SLICE_TIMEOUT_MS} + {@link DOCUMENT_BACKUP_COMMIT_TIMEOUT_MS}
 * — under eight minutes against a peer answering every request one millisecond inside its
 * deadline, where an unbounded client waited indefinitely.
 *
 * **Only `commit` writes, so only an aborted `commit` leaves an unknown outcome.** Decision 4 of
 * the route note is that nothing is written until the whole document has arrived: `begin` and
 * `slice` accumulate in the desktop's memory and `read` is a pure read, so a deadline on any of
 * them ends a wait and decides nothing. `commit` persists the copy and runs the retention sweep,
 * and past its deadline the file may exist, may not, and may be `.partial` — this client does not
 * know which. The repair is refused either way and the document is untouched, but saying "nothing
 * was changed" would assert something about the user's backup folder that was never established,
 * which is the same reclassification this rule exists to stop, one layer out.
 */

const TOKEN_SHAPE = /^[A-Za-z0-9_-]{43}$/;
const BACKUP_ID_SHAPE = /^[A-Za-z0-9._-]{1,128}$/;
const CANONICAL_BASE64 = /^[A-Za-z0-9+/]*={0,2}$/;
const BASE64_CHUNK_BYTES = 32768;

const GENERIC_REFUSAL = "Reference Manager could not create and verify a backup, so nothing was changed.";
const TOO_LARGE_REFUSAL =
  "This document is too large for automatic backup. Save a copy yourself first, then try the repair again.";
/**
 * The sentence for the one failure whose outcome this client does not know.
 *
 * It says the two things separately, because they are separately true: the document was not
 * changed (the repair is refused before it writes), and a copy may or may not have been saved.
 * {@link GENERIC_REFUSAL} would claim the second, and would send a user to an empty folder or
 * leave them unaware of a file that is there.
 */
const UNKNOWN_OUTCOME_REFUSAL =
  "Reference Manager stopped waiting while the backup was being saved. Your document was not changed, "
  + "but a copy may or may not have been saved — check your document backups before trying the repair again.";

/** The one route that writes; see the module note. */
const COMMIT_PATH = "/document-backup/commit";

/**
 * How long one slice-sized exchange may take: `begin`, one `slice`, one read-back.
 *
 * Sized from what those requests do, not from the number next door. The largest of them carries a
 * 4 MiB slice as base64 over loopback — `DOCUMENT_BACKUP_SLICE_BYTES`, about 5.6 MiB of JSON — and
 * the read-back adds a 4 MiB disk read on the far side; `begin` carries an integer. Ten seconds is
 * an order of magnitude past any of that, which is the point: this deadline separates *held* from
 * *slow*, not slow from fast. It lands on the annotation search's number by arriving at it, not by
 * inheriting it, and is named separately so that changing one cannot silently move the other.
 */
export const DOCUMENT_BACKUP_SLICE_TIMEOUT_MS = 10_000;

/**
 * How long `commit` may take — the only request that touches the disk on the far side.
 *
 * `commit` concatenates the whole document, writes it under `.partial`, links it into place and
 * then runs the retention sweep over the backup directory. At `MAX_DOCUMENT_BACKUP_BYTES` that is
 * 64 MiB of writing, and the backup folder may sit on a slow or network-backed volume: at a
 * deliberately pessimistic 1 MiB/s floor the write alone is about 64 s, with the sweep after it.
 * **That floor is chosen, not measured** — nothing here has ever timed a real one (R-011) — so it
 * is set low on purpose and 120 s is it with room, still far short of the minutes an unbounded
 * request would wait. It is twelve times the slice deadline because it does an unrelated amount of
 * work, not because a write deserves a longer number.
 */
export const DOCUMENT_BACKUP_COMMIT_TIMEOUT_MS = 120_000;

/**
 * The error every failure here becomes.
 *
 * `outcomeUnknown` is the write half of the classification: true only when the request that never
 * came back was the `commit`. A caller that treated the other aborts as undecided would send a
 * user looking through their backups after a `begin` that wrote nothing.
 */
export class DocumentBackupBridgeError extends Error {
  constructor(code, status = 0, outcomeUnknown = false) {
    const unknown = outcomeUnknown === true;
    super(unknown ? UNKNOWN_OUTCOME_REFUSAL : code === "document-too-large" ? TOO_LARGE_REFUSAL : GENERIC_REFUSAL);
    this.name = "DocumentBackupBridgeError";
    this.code = code;
    this.status = status;
    this.outcomeUnknown = unknown;
    this.userMessage = this.message;
  }
}

/** A fixed sentence safe to put in the task pane; never includes a response body or credential. */
export function formatDocumentBackupRefusal(error) {
  return error instanceof DocumentBackupBridgeError ? error.userMessage : GENERIC_REFUSAL;
}

function bridgeBaseUrl(value) {
  let parsed;
  try {
    parsed = new URL(value);
  } catch {
    throw new DocumentBackupBridgeError("invalid-bridge-config");
  }
  if (
    parsed.protocol !== "http:" ||
    parsed.hostname !== "127.0.0.1" ||
    parsed.port === "" ||
    parsed.username !== "" ||
    parsed.password !== "" ||
    parsed.pathname !== "/" ||
    parsed.search !== "" ||
    parsed.hash !== ""
  ) {
    throw new DocumentBackupBridgeError("invalid-bridge-config");
  }
  return parsed.origin;
}

function positiveInteger(value, field) {
  if (!Number.isSafeInteger(value) || value < 1) throw new DocumentBackupBridgeError(`invalid-${field}-response`);
  return value;
}

function canonicalBase64ToBytes(value) {
  if (typeof value !== "string" || value.length % 4 !== 0 || !CANONICAL_BASE64.test(value)) {
    throw new DocumentBackupBridgeError("invalid-base64-response");
  }
  let binary;
  try {
    binary = globalThis.atob(value);
  } catch {
    throw new DocumentBackupBridgeError("invalid-base64-response");
  }
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) bytes[index] = binary.charCodeAt(index);
  if (bytesToCanonicalBase64(bytes) !== value) throw new DocumentBackupBridgeError("invalid-base64-response");
  return bytes;
}

function bytesToCanonicalBase64(bytes) {
  const chunks = [];
  for (let offset = 0; offset < bytes.length; offset += BASE64_CHUNK_BYTES) {
    chunks.push(String.fromCharCode(...bytes.subarray(offset, offset + BASE64_CHUNK_BYTES)));
  }
  return globalThis.btoa(chunks.join(""));
}

/** A configured deadline, or the shipped one. Present-but-unusable is refused, never substituted. */
function deadline(value, fallback) {
  if (value === undefined) return fallback;
  if (!Number.isSafeInteger(value) || value <= 0) throw new DocumentBackupBridgeError("invalid-timeout-config");
  return value;
}

function objectResponse(value) {
  if (typeof value !== "object" || value === null || Array.isArray(value)) {
    throw new DocumentBackupBridgeError("invalid-json-response");
  }
  return value;
}

/**
 * Build the persistence/read-back half of `WordDocumentBackupPort`.
 *
 * `baseUrl` and `token` come from the desktop composition layer. The token is carried only in the
 * Authorization header: putting it in a URL would copy it into history, logs and Referer headers.
 */
export function createDocumentBackupBridgeClient(options) {
  const baseUrl = bridgeBaseUrl(options && options.baseUrl);
  const token = options && options.token;
  if (typeof token !== "string" || !TOKEN_SHAPE.test(token)) {
    throw new DocumentBackupBridgeError("invalid-bridge-config");
  }
  const fetchImpl = options.fetchImpl ?? globalThis.fetch.bind(globalThis);
  const onFailure = typeof options.onFailure === "function" ? options.onFailure : () => {};
  // Both exist so a test can state a deadline instead of waiting one out; the shipped task pane
  // passes neither. A present-but-unusable value is refused rather than quietly replaced by the
  // default — a caller that wrote a deadline down meant it.
  const sliceTimeoutMs = deadline(options.sliceTimeoutMs, DOCUMENT_BACKUP_SLICE_TIMEOUT_MS);
  const commitTimeoutMs = deadline(options.commitTimeoutMs, DOCUMENT_BACKUP_COMMIT_TIMEOUT_MS);

  async function request(path, init) {
    // `commit` is matched exactly: it is the only route called without a query string, and the
    // read-back's path carries one, so a prefix test would be looser than it needs to be.
    const writing = path === COMMIT_PATH;
    const controller = new AbortController();
    // Through `globalThis` for the same reason `fetch` is: this module is loaded by a task pane
    // where the timers hang off `window`, and by tests where there is no `window` at all.
    const timer = globalThis.setTimeout(
      () => controller.abort(),
      writing ? commitTimeoutMs : sliceTimeoutMs,
    );
    try {
      let response;
      try {
        response = await fetchImpl(`${baseUrl}${path}`, {
          ...init,
          mode: "cors",
          credentials: "omit",
          cache: "no-store",
          referrerPolicy: "no-referrer",
          signal: controller.signal,
          headers: {
            Accept: "application/json",
            Authorization: `Bearer ${token}`,
            ...(init.body === undefined ? {} : { "Content-Type": "application/json" }),
          },
        });
      } catch {
        // Our own deadline, a refused connection and a browser that declined the request are
        // indistinguishable here and mean the same thing: nobody answered. It must land here and
        // not on `bridge-refused`, which says the desktop application decided something it never
        // got the chance to decide. **`writing` alone decides `outcomeUnknown`, not `aborted`** —
        // a connection that failed after the body went out leaves the commit exactly as undecided
        // as our own deadline does, and this side cannot tell the two apart.
        throw new DocumentBackupBridgeError("bridge-unreachable", 0, writing);
      }

      let json;
      try {
        json = objectResponse(await response.json());
      } catch (cause) {
        // The deadline is still running here on purpose: a peer that answers 200 and then holds
        // the body open would hang exactly where the unbounded request did. The body is also where
        // the two causes separate — an abort is still nobody answering, while anything else is a
        // peer that answered with something this client cannot read, which decides the write.
        if (controller.signal.aborted) {
          throw new DocumentBackupBridgeError("bridge-unreachable", 0, writing);
        }
        const error = cause instanceof DocumentBackupBridgeError
          ? cause
          : new DocumentBackupBridgeError("invalid-json-response", response.status);
        throw error;
      }
      if (!response.ok) {
        const code = typeof json.error === "string" ? json.error : "bridge-refused";
        const error = new DocumentBackupBridgeError(code, response.status);
        throw error;
      }
      return json;
    } finally {
      // A timer left pending would abort a signal nothing is waiting on, and in the task pane
      // would hold a callback alive for two minutes after every backup the user ever took.
      globalThis.clearTimeout(timer);
    }
  }

  async function persistBackup(bytes) {
    if (!(bytes instanceof Uint8Array) || bytes.length === 0) {
      throw new DocumentBackupBridgeError("empty-document");
    }
    const begun = await request("/document-backup/begin", {
      method: "POST",
      body: JSON.stringify({ byteLength: bytes.length }),
    });
    const uploadId = typeof begun.uploadId === "string" ? begun.uploadId : "";
    const sliceBytes = positiveInteger(begun.sliceBytes, "slice-size");
    const sliceCount = positiveInteger(begun.sliceCount, "slice-count");
    if (uploadId.length === 0 || sliceCount !== Math.ceil(bytes.length / sliceBytes)) {
      throw new DocumentBackupBridgeError("invalid-begin-response");
    }

    for (let index = 0; index < sliceCount; index += 1) {
      const slice = bytes.subarray(index * sliceBytes, Math.min(bytes.length, (index + 1) * sliceBytes));
      await request("/document-backup/slice", {
        method: "POST",
        body: JSON.stringify({ uploadId, index, bytesBase64: bytesToCanonicalBase64(slice) }),
      });
    }

    const committed = await request("/document-backup/commit", {
      method: "POST",
      body: JSON.stringify({ uploadId }),
    });
    const location = typeof committed.backupId === "string" ? committed.backupId : "";
    if (!BACKUP_ID_SHAPE.test(location) || committed.byteLength !== bytes.length) {
      throw new DocumentBackupBridgeError("invalid-commit-response");
    }
    return { location };
  }

  async function readBackBackup(location) {
    if (typeof location !== "string" || !BACKUP_ID_SHAPE.test(location)) {
      throw new DocumentBackupBridgeError("invalid-backup-id");
    }
    const slices = [];
    let expectedSliceCount = null;
    let expectedByteLength = null;
    let received = 0;

    for (let index = 0; expectedSliceCount === null || index < expectedSliceCount; index += 1) {
      const result = await request(
        `/document-backup/read?id=${encodeURIComponent(location)}&slice=${index}`,
        { method: "GET" },
      );
      const sliceCount = positiveInteger(result.sliceCount, "slice-count");
      const byteLength = positiveInteger(result.byteLength, "byte-length");
      if (
        result.backupId !== location ||
        result.index !== index ||
        (expectedSliceCount !== null && sliceCount !== expectedSliceCount) ||
        (expectedByteLength !== null && byteLength !== expectedByteLength)
      ) {
        throw new DocumentBackupBridgeError("inconsistent-read-response");
      }
      expectedSliceCount = sliceCount;
      expectedByteLength = byteLength;
      const bytes = canonicalBase64ToBytes(result.bytesBase64);
      received += bytes.length;
      if (received > byteLength) throw new DocumentBackupBridgeError("invalid-read-length");
      slices.push(bytes);
    }

    if (expectedByteLength === null || received !== expectedByteLength) {
      throw new DocumentBackupBridgeError("invalid-read-length");
    }
    const document = new Uint8Array(received);
    let offset = 0;
    for (const slice of slices) {
      document.set(slice, offset);
      offset += slice.length;
    }
    return document;
  }

  async function reportFailure(operation) {
    try {
      return await operation();
    } catch (cause) {
      const error = cause instanceof DocumentBackupBridgeError
        ? cause
        : new DocumentBackupBridgeError("unexpected-client-failure");
      onFailure(error);
      throw error;
    }
  }

  return Object.freeze({
    persistBackup: (bytes) => reportFailure(() => persistBackup(bytes)),
    readBackBackup: (location) => reportFailure(() => readBackBackup(location)),
  });
}
