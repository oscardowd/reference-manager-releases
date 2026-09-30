/**
 * The legacy function file (SPEC §9.1; ADR-0101).
 *
 * The manifest makes `taskpane.html` the function file — the add-in runs one shared runtime — so
 * the ribbon's functions are registered in `taskpane.js`, where they can open the pane on the
 * workflow they name. This file is kept for a host that loads `commands.html` without a shared
 * runtime: there each function opens the pane and nothing more, because a ribbon button must never
 * write to the document on its own (§9.4 makes regeneration a decision the pane asks about).
 *
 * `apps/word-taskpane/src/addin.test.ts` holds both files to registering exactly the ribbon's
 * functions, so neither can declare a button the other leaves empty.
 */

/** Open the pane and complete the ribbon event, whatever happened. */
function openPane(event) {
  Promise.resolve()
    .then(() => Office.addin.showAsTaskpane())
    .catch(() => undefined)
    .finally(() => event.completed());
}

Office.onReady(() => {
  Office.actions.associate("insertCitation", openPane);
  Office.actions.associate("editCitation", openPane);
  Office.actions.associate("insertBibliography", openPane);
  Office.actions.associate("refreshDocument", openPane);
  Office.actions.associate("openCitationStyle", openPane);
});
