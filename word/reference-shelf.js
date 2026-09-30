/**
 * References the writer added in the pane without the desktop (ADR-0242 decision 5).
 *
 * In document mode the document's own §9.3 snapshots are the library. A reference it does not cite
 * yet — looked up by DOI or PMID, or typed in — has to live somewhere between being added and being
 * cited, and should still be there for the next document on this device. It lives here: in the
 * pane's own storage on this device, never sent anywhere. Once cited, the document carries its own
 * copy, and removing it from the shelf does not touch any document.
 *
 * Every entry is CSL-JSON with an id, a type and the moment it was added. Anything else read back
 * from storage — a truncated write, another version's shape — is dropped on load rather than
 * offered as a reference.
 */

const SHELF_KEY = "refmgr:shelf:v1";
/** A shelf, not a library: past this the writer wants the desktop's library. */
export const SHELF_LIMIT = 500;

function wellFormed(entry) {
  return entry !== null && typeof entry === "object" &&
    typeof entry.itemId === "string" && entry.itemId.length > 0 &&
    typeof entry.addedAt === "string" &&
    entry.csl !== null && typeof entry.csl === "object" && !Array.isArray(entry.csl) &&
    entry.csl.id === entry.itemId && typeof entry.csl.type === "string";
}

export async function createReferenceShelf(storage) {
  let entries = [];
  try {
    const stored = JSON.parse((await storage.get(SHELF_KEY)) ?? "[]");
    if (Array.isArray(stored)) entries = stored.filter(wellFormed);
  } catch {
    entries = [];
  }

  async function save() {
    await storage.set(SHELF_KEY, JSON.stringify(entries));
  }

  return Object.freeze({
    /** Newest first. */
    list() {
      return entries;
    },
    /** Add or replace one reference by id; the newest copy wins and moves to the front. */
    async add(reference) {
      if (!wellFormed(reference)) throw new Error("That reference is incomplete, so it was not added.");
      entries = [reference, ...entries.filter((entry) => entry.itemId !== reference.itemId)].slice(0, SHELF_LIMIT);
      await save();
      return reference;
    },
    async remove(itemId) {
      entries = entries.filter((entry) => entry.itemId !== itemId);
      await save();
    },
  });
}
