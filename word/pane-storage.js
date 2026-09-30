/**
 * Where the task pane keeps what it may reuse without the desktop (ADR-0242): the CSL styles it
 * downloaded and the references the writer added in the pane.
 *
 * Office offers `OfficeRuntime.storage` to add-ins with a shared runtime (Windows, Mac, the web);
 * Word on iPad has no shared runtime, so `localStorage` is used there. Either may be missing or
 * refuse — a private window, a full quota, a host that blocks storage — and **storage failing
 * never breaks the pane**: it falls back to memory for the rest of the session and says which store
 * it is using, so the Help tab can tell the writer that nothing will be remembered.
 *
 * Nothing stored here is a secret or a document. It lives only on this device, in this add-in's own
 * storage, and is never sent anywhere.
 */

function memoryStore() {
  const values = new Map();
  return {
    kind: "memory",
    async get(key) { return values.has(key) ? values.get(key) : null; },
    async set(key, value) { values.set(key, String(value)); },
    async remove(key) { values.delete(key); },
  };
}

function officeStore(storage, memory) {
  return {
    kind: "office",
    async get(key) {
      try {
        const value = await storage.getItem(key);
        return typeof value === "string" ? value : memory.get(key);
      } catch {
        return memory.get(key);
      }
    },
    async set(key, value) {
      try { await storage.setItem(key, String(value)); } catch { await memory.set(key, value); }
    },
    async remove(key) {
      await memory.remove(key);
      try { await storage.removeItem(key); } catch { /* already gone from memory; nothing else to do */ }
    },
  };
}

function localStore(local, memory) {
  return {
    kind: "local",
    async get(key) {
      try {
        const value = local.getItem(key);
        return value === null ? memory.get(key) : value;
      } catch {
        return memory.get(key);
      }
    },
    async set(key, value) {
      try { local.setItem(key, String(value)); } catch { await memory.set(key, value); }
    },
    async remove(key) {
      await memory.remove(key);
      try { local.removeItem(key); } catch { /* already gone from memory; nothing else to do */ }
    },
  };
}

/** `localStorage`, if this page may use it at all. Merely reading the property can throw. */
function availableLocalStorage() {
  try {
    const local = globalThis.localStorage;
    if (!local) return null;
    const probe = "refmgr:probe";
    local.setItem(probe, "1");
    local.removeItem(probe);
    return local;
  } catch {
    return null;
  }
}

export function createPaneStorage({ officeRuntime = globalThis.OfficeRuntime, local = availableLocalStorage() } = {}) {
  const memory = memoryStore();
  const storage = officeRuntime && officeRuntime.storage;
  if (storage && typeof storage.getItem === "function" && typeof storage.setItem === "function") {
    return officeStore(storage, memory);
  }
  if (local) return localStore(local, memory);
  return memory;
}
