/**
 * Task-pane navigation and small presentation state (E10-19).
 *
 * The document workflow modules own their domain events. This module only makes the page's
 * navigation predictable: three tabs with the ARIA tab pattern, a workflow link (or a ribbon
 * command) that opens the tab and any disclosure holding its target and moves focus to the first
 * useful control, and the connection indicator. No library, citation or document decision belongs
 * here.
 */

function closestDetails(element) {
  let current = element;
  while (current !== null && current !== undefined) {
    if (current.tagName === "DETAILS") return current;
    current = current.parentElement;
  }
  return null;
}

function closestTabPanel(element) {
  let current = element;
  while (current !== null && current !== undefined) {
    if (typeof current.getAttribute === "function" && current.getAttribute("role") === "tabpanel") return current;
    current = current.parentElement;
  }
  return null;
}

function focusElement(element) {
  if (!element || typeof element.focus !== "function") return;
  try {
    element.focus({ preventScroll: true });
  } catch {
    element.focus();
  }
}

function prefersReducedMotion(document) {
  const view = document.defaultView;
  if (!view || typeof view.matchMedia !== "function") return false;
  return view.matchMedia("(prefers-reduced-motion: reduce)").matches === true;
}

/** Every tab, in document order. */
function tabsOf(document) {
  return typeof document.querySelectorAll === "function" ? [...document.querySelectorAll('[role="tab"]')] : [];
}

/**
 * Show one tab's panel and hide the others, with the ARIA marking and the roving tab stop that the
 * tab pattern requires. Returns false when the id names no tab, so a caller can tell a real switch
 * from a typo.
 */
export function selectTaskPaneTab(document, tabId, { focus = false } = {}) {
  const tabs = tabsOf(document);
  const chosen = tabs.find((tab) => tab.id === tabId);
  if (!chosen) return false;
  for (const tab of tabs) {
    const selected = tab === chosen;
    tab.setAttribute("aria-selected", selected ? "true" : "false");
    tab.tabIndex = selected ? 0 : -1;
    const panel = document.getElementById(tab.getAttribute("aria-controls"));
    if (panel) panel.hidden = !selected;
  }
  if (focus) focusElement(chosen);
  return true;
}

/** Click to select; Left/Right/Home/End move between tabs, as the ARIA tab pattern specifies. */
export function setupTaskPaneTabs(document) {
  const tabs = tabsOf(document);
  for (const [index, tab] of tabs.entries()) {
    tab.addEventListener("click", () => selectTaskPaneTab(document, tab.id));
    tab.addEventListener("keydown", (event) => {
      const last = tabs.length - 1;
      const next =
        event.key === "ArrowRight" ? (index === last ? 0 : index + 1)
          : event.key === "ArrowLeft" ? (index === 0 ? last : index - 1)
            : event.key === "Home" ? 0
              : event.key === "End" ? last
                : null;
      if (next === null) return;
      event.preventDefault();
      selectTaskPaneTab(document, tabs[next].id, { focus: true });
    });
  }
}

/**
 * Opens, reveals and focuses one task-pane workflow. Page navigation, the Document tab and ribbon
 * commands all use this path, so the operating-system motion preference cannot drift between them.
 * The tab holding the target is selected first: a workflow on a hidden tab is not revealed by
 * scrolling to it.
 */
export function focusTaskPaneWorkflow(document, targetId, focusId) {
  const target = typeof targetId === "string" ? document.getElementById(targetId) : null;
  if (!target) return false;

  const panel = closestTabPanel(target);
  if (panel && typeof panel.getAttribute === "function") {
    const tab = tabsOf(document).find((candidate) => candidate.getAttribute("aria-controls") === panel.id);
    if (tab) selectTaskPaneTab(document, tab.id);
  }

  let details = closestDetails(target);
  while (details) {
    details.open = true;
    details = closestDetails(details.parentElement);
  }
  target.scrollIntoView({
    behavior: prefersReducedMotion(document) ? "auto" : "smooth",
    block: "start",
  });

  const focusTarget = typeof focusId === "string" ? document.getElementById(focusId) : null;
  focusElement(focusTarget);
  return true;
}

export function setupTaskPaneNavigation(document) {
  const links = [...document.querySelectorAll("[data-task-target]")];
  for (const link of links) {
    link.addEventListener("click", (event) => {
      if (focusTaskPaneWorkflow(document, link.dataset.taskTarget, link.dataset.focusTarget)) {
        event.preventDefault();
      }
    });
  }
}

/**
 * The header's connection pill. It is a few words wide, so it says only the state; what went wrong
 * and what to do about it belong to the caller's status lines, where there is room to say it.
 * States: `pending`, `ready` (the library), `document` (the document's own references), `error`.
 */
export function setConnectionBanner(document, state) {
  const banner = document.getElementById("connection-status");
  if (!banner) return;
  const message = banner.querySelector("span:last-child");
  banner.classList.remove("connection-banner-pending", "connection-banner-ready", "connection-banner-document", "connection-banner-error");
  if (state === "ready") {
    banner.classList.add("connection-banner-ready");
    if (message) message.textContent = "Connected";
  } else if (state === "document") {
    // Working, but from the document's own references rather than the library (ADR-0242). A state
    // of its own, not "Not connected": the pane can do nearly everything, and saying it cannot
    // would send the writer looking for a fault.
    banner.classList.add("connection-banner-document");
    if (message) message.textContent = "This document";
  } else if (state === "error") {
    banner.classList.add("connection-banner-error");
    if (message) message.textContent = "Not connected";
  } else {
    banner.classList.add("connection-banner-pending");
    if (message) message.textContent = "Connecting…";
  }
}
