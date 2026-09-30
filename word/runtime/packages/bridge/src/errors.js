/**
 * The one error a route handler may throw to answer a caller.
 *
 * It lives in its own module so that both `routes.ts` and the document-backup surface can raise it
 * without importing each other. `routes.ts` re-exports it, so the package's public surface is
 * unchanged.
 */
export class BridgeRequestError extends Error {
    status;
    code;
    constructor(status, code, message) {
        super(message);
        this.status = status;
        this.code = code;
        this.name = "BridgeRequestError";
    }
}
/** Refuse a request. Messages name the rule, never the value (ADR-0099). */
export function badRequest(code, message) {
    throw new BridgeRequestError(400, code, message);
}
