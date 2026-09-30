/**
 * SHA-256 (FIPS 180-4) of a string's UTF-8 bytes, synchronously and without a platform module
 * (ADR-0242 decision 3).
 *
 * Why this exists rather than `node:crypto`: the fingerprints in `canonical.ts` are computed deep
 * inside synchronous planners (`prepareInsertion`, `prepareDocumentRefresh`, the tag of every
 * content control), and those planners now also run in the Word task pane, where the browser's
 * only digest (`crypto.subtle.digest`) is asynchronous. Making every planner asynchronous to reach
 * it would be a far larger change than a hash function.
 *
 * **The output is a contract, not a detail.** Every Word document this product has written carries
 * these digests in its content-control tags and its payload's `contentHash` / `renderedHash`. A
 * digest that differed from the one `node:crypto` produced would make every existing document read
 * as changed. `sha256.test.ts` therefore checks the NIST vectors and holds this function equal to
 * `node:crypto` over generated inputs, including lone surrogates, which both must encode as U+FFFD.
 *
 * Not for secrets. Nothing here is constant-time and nothing needs to be: a fingerprint is a
 * public identity of public data.
 */
const ROUND_CONSTANTS = new Uint32Array([
    0x428a2f98, 0x71374491, 0xb5c0fbcf, 0xe9b5dba5, 0x3956c25b, 0x59f111f1, 0x923f82a4, 0xab1c5ed5,
    0xd807aa98, 0x12835b01, 0x243185be, 0x550c7dc3, 0x72be5d74, 0x80deb1fe, 0x9bdc06a7, 0xc19bf174,
    0xe49b69c1, 0xefbe4786, 0x0fc19dc6, 0x240ca1cc, 0x2de92c6f, 0x4a7484aa, 0x5cb0a9dc, 0x76f988da,
    0x983e5152, 0xa831c66d, 0xb00327c8, 0xbf597fc7, 0xc6e00bf3, 0xd5a79147, 0x06ca6351, 0x14292967,
    0x27b70a85, 0x2e1b2138, 0x4d2c6dfc, 0x53380d13, 0x650a7354, 0x766a0abb, 0x81c2c92e, 0x92722c85,
    0xa2bfe8a1, 0xa81a664b, 0xc24b8b70, 0xc76c51a3, 0xd192e819, 0xd6990624, 0xf40e3585, 0x106aa070,
    0x19a4c116, 0x1e376c08, 0x2748774c, 0x34b0bcb5, 0x391c0cb3, 0x4ed8aa4a, 0x5b9cca4f, 0x682e6ff3,
    0x748f82ee, 0x78a5636f, 0x84c87814, 0x8cc70208, 0x90befffa, 0xa4506ceb, 0xbef9a3f7, 0xc67178f2,
]);
const INITIAL_STATE = [
    0x6a09e667, 0xbb67ae85, 0x3c6ef372, 0xa54ff53a, 0x510e527f, 0x9b05688c, 0x1f83d9ab, 0x5be0cd19,
];
const encoder = new TextEncoder();
function rotateRight(value, bits) {
    return (value >>> bits) | (value << (32 - bits));
}
/** The digest of `bytes`, as 32 bytes. */
export function sha256Bytes(bytes) {
    // Padding: a 1 bit, zeros, then the message length in bits as a 64-bit big-endian integer, so
    // the total is a multiple of 64 bytes.
    const bitLength = bytes.length * 8;
    const paddedLength = Math.ceil((bytes.length + 9) / 64) * 64;
    const message = new Uint8Array(paddedLength);
    message.set(bytes);
    message[bytes.length] = 0x80;
    const view = new DataView(message.buffer);
    // A string long enough to overflow 2^32 bits would be half a gigabyte of payload; the high word
    // is still written correctly rather than assumed zero.
    view.setUint32(paddedLength - 8, Math.floor(bitLength / 0x1_0000_0000));
    view.setUint32(paddedLength - 4, bitLength >>> 0);
    const state = [...INITIAL_STATE];
    const schedule = new Uint32Array(64);
    for (let offset = 0; offset < paddedLength; offset += 64) {
        for (let t = 0; t < 16; t += 1)
            schedule[t] = view.getUint32(offset + t * 4);
        for (let t = 16; t < 64; t += 1) {
            const w15 = schedule[t - 15];
            const w2 = schedule[t - 2];
            const s0 = rotateRight(w15, 7) ^ rotateRight(w15, 18) ^ (w15 >>> 3);
            const s1 = rotateRight(w2, 17) ^ rotateRight(w2, 19) ^ (w2 >>> 10);
            schedule[t] = (schedule[t - 16] + s0 + schedule[t - 7] + s1) >>> 0;
        }
        let [a, b, c, d, e, f, g, h] = state;
        for (let t = 0; t < 64; t += 1) {
            const sum1 = rotateRight(e, 6) ^ rotateRight(e, 11) ^ rotateRight(e, 25);
            const choice = (e & f) ^ (~e & g);
            const temp1 = (h + sum1 + choice + ROUND_CONSTANTS[t] + schedule[t]) >>> 0;
            const sum0 = rotateRight(a, 2) ^ rotateRight(a, 13) ^ rotateRight(a, 22);
            const majority = (a & b) ^ (a & c) ^ (b & c);
            const temp2 = (sum0 + majority) >>> 0;
            h = g;
            g = f;
            f = e;
            e = (d + temp1) >>> 0;
            d = c;
            c = b;
            b = a;
            a = (temp1 + temp2) >>> 0;
        }
        state[0] = (state[0] + a) >>> 0;
        state[1] = (state[1] + b) >>> 0;
        state[2] = (state[2] + c) >>> 0;
        state[3] = (state[3] + d) >>> 0;
        state[4] = (state[4] + e) >>> 0;
        state[5] = (state[5] + f) >>> 0;
        state[6] = (state[6] + g) >>> 0;
        state[7] = (state[7] + h) >>> 0;
    }
    const digest = new Uint8Array(32);
    const out = new DataView(digest.buffer);
    state.forEach((word, index) => out.setUint32(index * 4, word));
    return digest;
}
/**
 * The lowercase hex digest of `text` encoded as UTF-8 — what
 * `createHash("sha256").update(text, "utf8").digest("hex")` returns.
 */
export function sha256Hex(text) {
    const digest = sha256Bytes(encoder.encode(text));
    let hex = "";
    for (const byte of digest)
        hex += byte.toString(16).padStart(2, "0");
    return hex;
}
