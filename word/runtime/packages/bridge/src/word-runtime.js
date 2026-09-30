/**
 * The `@refmgr/bridge` surface the Word task pane's runtime reaches (ADR-0242 decision 3).
 *
 * The router and its route table, and nothing that listens, authenticates or touches a file: the
 * package index also exports the loopback server, the token and the handshake writer, which need
 * `node:http`, `node:crypto` and `node:fs` and have no meaning inside a task pane.
 */
export { BridgeRequestError, createPortRouter } from "./routes.js";
export { BRIDGE_ROUTES } from "./route-table.js";
