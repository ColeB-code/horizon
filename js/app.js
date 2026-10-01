import { initAuth, getActiveAccount, signIn, signOut, getAccessToken } from "./auth.js";
import { getCalendars, getEventDescription, getHighlightEvents } from "./graph.js";

const DEFAULT_LENGTH_MONTHS = 6;

const loginBtn = document.getElementById("loginBtn");
const signOutBtn = document.getElementById("signOutBtn");
const themeToggleBtn = document.getElementById("themeToggleBtn");
const userNameEl = document.getElementById("userName");
const statusEl = document.getElementById("statusMessage");
const progressBarEl = document.getElementById("progressBar");
const progressBarFillEl = document.getElementById("progressBarFill");
const dashboardEl = document.getElementById("dashboard");
const currentDateEl = document.getElementById("currentDate");
const eventGroupsEl = document.getElementById("eventGroups");
const calendarListEl = document.getElementById("calendarList");
const skippedCalendarListEl = document.getElementById("skippedCalendarList");
const categoryListEl = document.getElementById("categoryList");
const tagListEl = document.getElementById("tagList");
const loadSummaryEl = document.getElementById("loadSummary");
const cacheVersionEl = document.getElementById("cacheVersion");

const THEME_STORAGE_KEY = "horizon-theme";
const INLINE_IMAGE_TYPES = new Set(["image/gif", "image/jpeg", "image/png", "image/webp"]);
const SHOW_IMAGE_DIAGNOSTICS = false;
const DESCRIPTION_NON_WHITESPACE = /[^\s\u00a0\u200b\u200c\u200d\u2060\ufeff]/u;
const DESCRIPTION_TRAILING_WHITESPACE = /[\s\u00a0\u200b\u200c\u200d\u2060\ufeff]+$/u;

function getCurrentTheme() {
  const stored = localStorage.getItem(THEME_STORAGE_KEY);
  if (stored) return stored;
  return window.matchMedia("(prefers-color-scheme: dark)").matches ? "dark" : "light";
}

function applyTheme(theme) {
  document.documentElement.setAttribute("data-theme", theme);
  localStorage.setItem(THEME_STORAGE_KEY, theme);
}

function toggleTheme() {
  const next = getCurrentTheme() === "dark" ? "light" : "dark";
  applyTheme(next);
}

function setStatus(message) {
  statusEl.textContent = message || "";
}

async function renderCacheVersion() {
  try {
    const cacheNames = "caches" in window ? await window.caches.keys() : [];
    const versionNumbers = cacheNames
      .map((name) => /^family-horizon-v(\d+)$/.exec(name))
      .filter(Boolean)
      .map((match) => Number(match[1]));
    const installedVersion = versionNumbers.length > 0 ? Math.max(...versionNumbers) : null;
    if (installedVersion !== null) {
      cacheVersionEl.textContent = `v${installedVersion}`;
    }

    const sourceVersionUrl = new URL("app-version.json", window.location.href);
    sourceVersionUrl.searchParams.set("_", Date.now().toString());
    const response = await fetch(sourceVersionUrl, { cache: "no-store" });
    if (!response.ok) throw new Error(`Version request failed (${response.status})`);

    const sourceVersion = Number((await response.json()).version);
    if (Number.isInteger(sourceVersion)) {
      cacheVersionEl.textContent = `${installedVersion ?? "?"}.${sourceVersion}`;
    }
  } catch (error) {
    console.warn("Could not compare installed and source versions:", error);
  }
}

function showProgress({ completed, total }) {
  progressBarEl.hidden = false;
  progressBarFillEl.style.width = `${Math.round((completed / total) * 100)}%`;
}

function hideProgress() {
  progressBarEl.hidden = true;
  progressBarFillEl.style.width = "0%";
}

function showSignedInUI(account) {
  loginBtn.hidden = true;
  signOutBtn.hidden = false;
  userNameEl.textContent = account.name?.trim().split(/\s+/)[0] || "Signed in";
  dashboardEl.hidden = false;
}

function showSignedOutUI() {
  loginBtn.hidden = false;
  signOutBtn.hidden = true;
  userNameEl.textContent = "";
  dashboardEl.hidden = true;
}

function formatMonthKey(date) {
  return date.toLocaleDateString(undefined, { month: "long", year: "numeric" });
}

function addMonths(date, months) {
  const result = new Date(date);
  const day = result.getDate();
  result.setDate(1);
  result.setMonth(result.getMonth() + months);
  const lastDay = new Date(result.getFullYear(), result.getMonth() + 1, 0).getDate();
  result.setDate(Math.min(day, lastDay));
  return result;
}

function renderCurrentDate() {
  currentDateEl.textContent = new Intl.DateTimeFormat(undefined, { dateStyle: "full" }).format(new Date());
}

function normalizeContentId(value) {
  const contentId = value.replace(/^cid:/i, "").replace(/^<|>$/g, "").trim();
  try {
    return decodeURIComponent(contentId).toLowerCase();
  } catch {
    return contentId.toLowerCase();
  }
}

function hasVisibleDescriptionContent(node) {
  if (node.nodeType === Node.TEXT_NODE) {
    return DESCRIPTION_NON_WHITESPACE.test(node.textContent);
  }
  if (node.nodeType !== Node.ELEMENT_NODE) return false;
  if (node.matches("img, hr, video, audio, canvas, svg")) return true;
  return Array.from(node.childNodes).some(hasVisibleDescriptionContent);
}

function appendImageDiagnostics(target, diagnostics) {
  if (!SHOW_IMAGE_DIAGNOSTICS) return;

  const graphDiagnostics = diagnostics.graph;
  const hasImageEvidence = graphDiagnostics && (
    graphDiagnostics.cidReferenceCount > 0 ||
    graphDiagnostics.imageMarkup?.length > 0 ||
    graphDiagnostics.attachmentList?.some((attachment) => attachment.contentType?.toLowerCase().startsWith("image/"))
  );
  if (!hasImageEvidence) return;

  const panel = document.createElement("details");
  panel.className = "image-diagnostics";
  panel.open = true;

  const heading = document.createElement("summary");
  heading.textContent = "Image diagnostics";

  const report = document.createElement("pre");
  report.textContent = JSON.stringify(diagnostics, null, 2);
  panel.append(heading, report);
  target.appendChild(panel);
}

function trimTrailingDescriptionWhitespace(node) {
  while (node.lastChild) {
    const lastChild = node.lastChild;
    if (lastChild.nodeType === Node.TEXT_NODE) {
      const trimmedText = lastChild.textContent.replace(DESCRIPTION_TRAILING_WHITESPACE, "");
      if (trimmedText) {
        lastChild.textContent = trimmedText;
        break;
      }
      lastChild.remove();
      continue;
    }

    if (lastChild.nodeType === Node.ELEMENT_NODE) {
      trimTrailingDescriptionWhitespace(lastChild);
      if (hasVisibleDescriptionContent(lastChild)) break;
    }
    lastChild.remove();
  }
}

function renderEventDescription(eventContent, target) {
  const body = eventContent?.body;
  if (!body?.content) {
    target.textContent = "No description provided.";
    appendImageDiagnostics(target, { graph: eventContent.imageDiagnostics, renderedImages: [] });
    return;
  }

  if (body.contentType?.toLowerCase() !== "html") {
    target.classList.add("event-description-plain-text");
    target.textContent = body.content.trimEnd();
    appendImageDiagnostics(target, { graph: eventContent.imageDiagnostics, renderedImages: [] });
    return;
  }

  target.classList.remove("event-description-plain-text");

  const sanitizedContent = window.DOMPurify.sanitize(body.content, {
    USE_PROFILES: { html: true },
    RETURN_DOM_FRAGMENT: true,
    FORBID_TAGS: ["embed", "form", "iframe", "object", "style"]
  });
  trimTrailingDescriptionWhitespace(sanitizedContent);

  for (const element of sanitizedContent.querySelectorAll("[style], [color], [bgcolor], [background]")) {
    element.removeAttribute("color");
    element.removeAttribute("bgcolor");
    element.removeAttribute("background");
    element.style.removeProperty("color");
    element.style.removeProperty("background");
    element.style.removeProperty("background-color");
    element.style.removeProperty("background-image");
    if (!element.getAttribute("style")?.trim()) element.removeAttribute("style");
  }

  const inlineAttachments = new Map(
    (eventContent.inlineAttachments || []).map((attachment) => [
      normalizeContentId(attachment.contentId || ""),
      attachment
    ])
  );
  const imageRenderDiagnostics = [];

  for (const image of sanitizedContent.querySelectorAll("img[src]")) {
    const source = image.getAttribute("src");
    if (!/^cid:/i.test(source)) {
      imageRenderDiagnostics.push({
        sourceScheme: /^[a-z][a-z\d+.-]*:/i.exec(source)?.[0] || "relative",
        outcome: "not a CID image; left unchanged"
      });
      continue;
    }

    const contentId = normalizeContentId(source);
    const attachment = inlineAttachments.get(contentId);
    const contentType = attachment?.contentType?.toLowerCase().split(";")[0].trim();
    const diagnostic = {
      contentId,
      matchedAttachment: Boolean(attachment),
      attachmentId: attachment?.id,
      attachmentName: attachment?.name,
      contentType,
      contentBytesLength: attachment?.contentBytes?.length || 0,
      altTextLength: image.getAttribute("alt")?.length || 0
    };

    if (!attachment) {
      diagnostic.outcome = "dropped: no matching attachment contentId";
      imageRenderDiagnostics.push(diagnostic);
      image.remove();
      continue;
    }

    if (!attachment.contentBytes) {
      diagnostic.outcome = "dropped: attachment has no contentBytes";
      imageRenderDiagnostics.push(diagnostic);
      image.remove();
      continue;
    }

    if (!INLINE_IMAGE_TYPES.has(contentType)) {
      diagnostic.outcome = "dropped: unsupported image MIME type";
      imageRenderDiagnostics.push(diagnostic);
      image.remove();
      continue;
    }

    image.setAttribute("src", `data:${contentType};base64,${attachment.contentBytes}`);
    diagnostic.outcome = "rendered";
    imageRenderDiagnostics.push(diagnostic);
  }

  target.replaceChildren(sanitizedContent);
  appendImageDiagnostics(target, {
    graph: eventContent.imageDiagnostics,
    renderedImages: imageRenderDiagnostics
  });
}

function createCopyIcon() {
  const svgNamespace = "http://www.w3.org/2000/svg";
  const icon = document.createElementNS(svgNamespace, "svg");
  icon.setAttribute("viewBox", "0 0 24 24");
  icon.setAttribute("aria-hidden", "true");
  icon.setAttribute("fill", "none");
  icon.setAttribute("stroke", "currentColor");
  icon.setAttribute("stroke-width", "1.8");
  icon.setAttribute("stroke-linecap", "round");
  icon.setAttribute("stroke-linejoin", "round");

  const rearPage = document.createElementNS(svgNamespace, "rect");
  rearPage.setAttribute("x", "8");
  rearPage.setAttribute("y", "8");
  rearPage.setAttribute("width", "14");
  rearPage.setAttribute("height", "14");
  rearPage.setAttribute("rx", "2");

  const frontPage = document.createElementNS(svgNamespace, "path");
  frontPage.setAttribute("d", "M16 8V5a2 2 0 0 0-2-2H5a2 2 0 0 0-2 2v9a2 2 0 0 0 2 2h3");
  icon.append(rearPage, frontPage);
  return icon;
}

function createExpandIcon() {
  const svgNamespace = "http://www.w3.org/2000/svg";
  const icon = document.createElementNS(svgNamespace, "svg");
  icon.setAttribute("viewBox", "0 0 24 24");
  icon.setAttribute("aria-hidden", "true");
  icon.setAttribute("fill", "none");
  icon.setAttribute("stroke", "currentColor");
  icon.setAttribute("stroke-width", "2");
  icon.setAttribute("stroke-linecap", "round");
  icon.setAttribute("stroke-linejoin", "round");

  const chevron = document.createElementNS(svgNamespace, "path");
  chevron.setAttribute("d", "m6 9 6 6 6-6");
  icon.append(chevron);
  return icon;
}

function formatEventDate(date) {
  return date.toLocaleDateString(undefined, { month: "short", day: "numeric" });
}

function formatEventTime(date) {
  return date.toLocaleTimeString(undefined, { hour: "numeric", minute: "2-digit" });
}

function isSameCalendarDay(a, b) {
  return a.toDateString() === b.toDateString();
}

// Builds the date (and time, and multi-day range) string shown on each event card.
function formatEventDateTime(event) {
  if (event.isAllDay) {
    // Graph's "end" for all-day events is exclusive (the day after the last day), so step it back one.
    const lastDay = new Date(event.end);
    lastDay.setDate(lastDay.getDate() - 1);

    return isSameCalendarDay(event.start, lastDay)
      ? formatEventDate(event.start)
      : `${formatEventDate(event.start)} \u2013 ${formatEventDate(lastDay)}`;
  }

  if (isSameCalendarDay(event.start, event.end)) {
    return `${formatEventDate(event.start)}, ${formatEventTime(event.start)}`;
  }

  return `${formatEventDate(event.start)} ${formatEventTime(event.start)} \u2013 ${formatEventDate(event.end)} ${formatEventTime(event.end)}`;
}

function renderEvents(events, lengthMonths, token) {
  if (events.length === 0) {
    const emptyState = document.createElement("p");
    emptyState.className = "empty-state";
    emptyState.textContent = `No Horizon events in the next ${lengthMonths} month${lengthMonths === 1 ? "" : "s"}.`;
    eventGroupsEl.replaceChildren(emptyState);
    return;
  }

  const groupsFragment = document.createDocumentFragment();
  const groups = new Map();
  let eventIndex = 0;
  for (const event of events) {
    const key = formatMonthKey(event.start);
    if (!groups.has(key)) {
      groups.set(key, []);
    }
    groups.get(key).push(event);
  }

  for (const [monthLabel, monthEvents] of groups) {
    const groupEl = document.createElement("div");
    groupEl.className = "month-group";

    const heading = document.createElement("h3");
    heading.textContent = monthLabel;
    groupEl.appendChild(heading);

    for (const event of monthEvents) {
      const card = document.createElement("div");
      card.className = "event-card";

      const metaParts = [event.calendarName];
      if (event.categories && event.categories.length > 0) {
        metaParts.push(event.categories.join(", "));
      }

      const dateEl = document.createElement("span");
      dateEl.className = "event-date";
      dateEl.textContent = formatEventDateTime(event);

      const subjectEl = document.createElement("span");
      subjectEl.className = "event-subject";
      subjectEl.textContent = event.subject;

      const metaEl = document.createElement("span");
      metaEl.className = "event-meta";
      metaEl.textContent = metaParts.join(" \u00b7 ");

      const descriptionEl = document.createElement("div");
      descriptionEl.className = "event-description";
      descriptionEl.id = `event-description-${eventIndex++}`;
      descriptionEl.hidden = true;

      const eventHeading = document.createElement("div");
      eventHeading.className = "event-heading";

      const eventSummary = document.createElement("div");
      eventSummary.className = "event-summary";
      eventSummary.append(dateEl, subjectEl, metaEl);

      const toggleButton = document.createElement("button");
      toggleButton.type = "button";
      toggleButton.className = "event-toggle-button";
      toggleButton.title = "Expand description";
      toggleButton.setAttribute("aria-label", "Expand description");
      toggleButton.setAttribute("aria-expanded", "false");
      toggleButton.setAttribute("aria-controls", descriptionEl.id);
      toggleButton.append(createExpandIcon());
      eventHeading.append(eventSummary, toggleButton);

      let descriptionLoaded = false;
      let descriptionLoading = false;
      toggleButton.addEventListener("click", async () => {
        const isExpanded = toggleButton.getAttribute("aria-expanded") === "true";
        toggleButton.setAttribute("aria-expanded", String(!isExpanded));
        toggleButton.title = isExpanded ? "Expand description" : "Collapse description";
        toggleButton.setAttribute("aria-label", toggleButton.title);
        descriptionEl.hidden = isExpanded;
        if (isExpanded || descriptionLoaded || descriptionLoading) return;

        descriptionLoading = true;
        descriptionEl.textContent = "Loading description...";
        try {
          let eventContent;
          try {
            eventContent = await getEventDescription(token, event.calendarId, event.id);
          } catch (error) {
            console.error("Could not fetch event description:", error);
            const graphError = /Graph request failed \((\d{3})\):\s*(.*)/.exec(error.message);
            descriptionEl.textContent = graphError
              ? `Outlook returned an error (${graphError[1]}): ${graphError[2]}`
              : "Could not reach Outlook to load the description.";
            return;
          }

          try {
            renderEventDescription(eventContent, descriptionEl);
            descriptionLoaded = true;
          } catch (error) {
            console.error("Could not render event description:", error);
            descriptionEl.textContent = "Could not display the formatted description.";
            appendImageDiagnostics(descriptionEl, {
              stage: "description rendering",
              graph: eventContent.imageDiagnostics,
              error: { name: error.name, message: error.message }
            });
          }
        } finally {
          descriptionLoading = false;
        }
      });

      card.append(eventHeading);

      if (event.location) {
        const locationRow = document.createElement("div");
        locationRow.className = "event-location-row";

        const locationEl = document.createElement("span");
        locationEl.className = "event-location";
        locationEl.textContent = event.location;

        const copyLocationButton = document.createElement("button");
        copyLocationButton.type = "button";
        copyLocationButton.className = "copy-location-button";
        copyLocationButton.title = "Copy location";
        copyLocationButton.setAttribute("aria-label", "Copy location");
        copyLocationButton.append(createCopyIcon());
        copyLocationButton.addEventListener("click", async () => {
          try {
            await navigator.clipboard.writeText(event.location);
            copyLocationButton.setAttribute("aria-label", "Location copied to clipboard");
          } catch (error) {
            console.error("Could not copy event location:", error);
            copyLocationButton.setAttribute("aria-label", "Could not copy location");
          }
          window.setTimeout(() => {
            copyLocationButton.setAttribute("aria-label", "Copy location");
          }, 1500);
        });

        locationRow.append(locationEl, copyLocationButton);
        card.append(locationRow);
      }

      card.append(descriptionEl);

      groupEl.appendChild(card);
    }

    groupsFragment.appendChild(groupEl);
  }
  eventGroupsEl.replaceChildren(groupsFragment);
}

// Renders a flat list of strings as <li> items into the given <ul>/<ol> element,
// showing a placeholder row when there's nothing to list.
function renderList(listEl, items) {
  const isEmpty = items.length === 0;
  const listItems = isEmpty ? ["(none)"] : items;
  const listElements = listItems.map((item) => {
    const listItem = document.createElement("li");
    if (isEmpty) listItem.className = "empty-item";
    listItem.textContent = item;
    return listItem;
  });
  listEl.replaceChildren(...listElements);
}

// Exclude calendars by name via ?exclude=Name1,Name2 in the URL (no hardcoded names in source).
function getExcludedCalendarNames() {
  const exclude = new URLSearchParams(window.location.search).get("exclude");
  if (!exclude) return [];
  return exclude.split(",").map((name) => name.trim().toLowerCase()).filter(Boolean);
}

// Filter by Outlook category via ?category=Name1,Name2 in the URL (case-insensitive, empty by default).
function getCategoryNames() {
  const category = new URLSearchParams(window.location.search).get("category");
  if (!category) return [];
  return category.split(",").map((name) => name.trim()).filter(Boolean);
}

// Filter by #hashtag in subject/location/description via ?tags=word1,word2 (case-insensitive,
// defaults to "horizon"). Hashtags can't contain spaces, so whitespace is stripped from each tag.
function getTagNames() {
  const tags = new URLSearchParams(window.location.search).get("tags");
  if (!tags) return ["horizon"];
  return tags
    .split(",")
    .map((tag) => tag.replace(/\s+/g, ""))
    .filter(Boolean);
}

// Override how many months ahead to look via ?length=N in the URL (any positive integer, defaults to 6).
function getLengthMonths() {
  const length = Number(new URLSearchParams(window.location.search).get("length"));
  return Number.isInteger(length) && length > 0 ? length : DEFAULT_LENGTH_MONTHS;
}

async function loadDashboard() {
  const startTime = performance.now();
  setStatus("Loading your Horizon events...");

  const token = await getAccessToken();
  const allCalendars = await getCalendars(token);
  const excludedNames = getExcludedCalendarNames();

  // Single pass split instead of filtering the full list twice.
  const calendars = [];
  const skippedCalendarNames = [];
  for (const calendar of allCalendars) {
    if (excludedNames.includes(calendar.name.trim().toLowerCase())) {
      skippedCalendarNames.push(calendar.name);
    } else {
      calendars.push(calendar);
    }
  }
  renderList(calendarListEl, calendars.map((calendar) => calendar.name));
  renderList(skippedCalendarListEl, skippedCalendarNames);

  const categoryNames = getCategoryNames();
  renderList(categoryListEl, categoryNames);

  const tagNames = getTagNames();
  renderList(tagListEl, tagNames.map((tag) => `#${tag}`));

  const start = new Date();
  const lengthMonths = getLengthMonths();
  const end = addMonths(start, lengthMonths);

  const events = await getHighlightEvents(
    token,
    calendars,
    start.toISOString(),
    end.toISOString(),
    showProgress,
    categoryNames,
    tagNames
  );

  setStatus("Building your dashboard...");
  progressBarFillEl.style.width = "100%";
  await new Promise((resolve) => requestAnimationFrame(resolve));

  renderEvents(events, lengthMonths, token);
  hideProgress();
  setStatus("");

  const elapsedSeconds = ((performance.now() - startTime) / 1000).toFixed(1);
  loadSummaryEl.textContent = `${lengthMonths} month${lengthMonths === 1 ? "" : "s"} loaded in ${elapsedSeconds} seconds`;
}

async function handleLogin() {
  setStatus("Signing in...");
  try {
    const account = await signIn();
    showSignedInUI(account);
    await loadDashboard();
  } catch (error) {
    console.error(error);
    setStatus("Sign-in failed. Please try again.");
  }
}

async function handleSignOut() {
  await signOut();
  showSignedOutUI();
  setStatus("");
}

async function main() {
  loginBtn.addEventListener("click", handleLogin);
  signOutBtn.addEventListener("click", handleSignOut);
  themeToggleBtn.addEventListener("click", toggleTheme);
  renderCurrentDate();
  renderCacheVersion();

  if ("serviceWorker" in navigator) {
    navigator.serviceWorker.addEventListener("controllerchange", renderCacheVersion);
    navigator.serviceWorker.register("service-worker.js")
      .then(renderCacheVersion)
      .catch((error) => {
        console.error("Service worker registration failed:", error);
      });
  }

  try {
    await initAuth();
  } catch (error) {
    console.error(error);
    setStatus("Failed to initialize sign-in. Check the browser console for details.");
    return;
  }

  const account = getActiveAccount();
  if (account) {
    showSignedInUI(account);
    try {
      await loadDashboard();
    } catch (error) {
      console.error(error);
      setStatus("Could not load events. Try signing in again.");
      showSignedOutUI();
    }
  } else {
    showSignedOutUI();
  }

}

main();
