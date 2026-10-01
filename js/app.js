import { initAuth, getActiveAccount, signIn, signOut, getAccessToken } from "./auth.js";
import { getCalendars, getHighlightEvents } from "./graph.js";

const DEFAULT_LENGTH_MONTHS = 6;

const loginBtn = document.getElementById("loginBtn");
const signOutBtn = document.getElementById("signOutBtn");
const themeToggleBtn = document.getElementById("themeToggleBtn");
const userNameEl = document.getElementById("userName");
const statusEl = document.getElementById("statusMessage");
const progressBarEl = document.getElementById("progressBar");
const progressBarFillEl = document.getElementById("progressBarFill");
const dashboardEl = document.getElementById("dashboard");
const eventGroupsEl = document.getElementById("eventGroups");
const calendarListEl = document.getElementById("calendarList");
const skippedCalendarListEl = document.getElementById("skippedCalendarList");
const categoryListEl = document.getElementById("categoryList");
const tagListEl = document.getElementById("tagList");
const processingTimeEl = document.getElementById("processingTime");

const THEME_STORAGE_KEY = "horizon-theme";

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
  userNameEl.textContent = account.username;
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

function renderEvents(events, lengthMonths) {
  eventGroupsEl.innerHTML = "";

  if (events.length === 0) {
    eventGroupsEl.innerHTML = `<p class="empty-state">No Horizon events in the next ${lengthMonths} month${lengthMonths === 1 ? "" : "s"}.</p>`;
    return;
  }

  const groups = new Map();
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
      if (event.location) {
        metaParts.push(event.location);
      }

      card.innerHTML = `
        <div class="event-date">${formatEventDateTime(event)}</div>
        <div class="event-subject">${event.subject}</div>
        <div class="event-meta">${metaParts.join(" &middot; ")}</div>
      `;

      groupEl.appendChild(card);
    }

    eventGroupsEl.appendChild(groupEl);
  }
}

// Renders a flat list of strings as <li> items into the given <ul>/<ol> element,
// showing a placeholder row when there's nothing to list.
function renderList(listEl, items) {
  if (items.length === 0) {
    listEl.innerHTML = '<li class="empty-item">(none)</li>';
    return;
  }
  listEl.innerHTML = items.map((item) => `<li>${item}</li>`).join("");
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
  const end = new Date();
  const lengthMonths = getLengthMonths();
  end.setMonth(end.getMonth() + lengthMonths);

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

  renderEvents(events, lengthMonths);
  hideProgress();
  setStatus("");

  const elapsedSeconds = ((performance.now() - startTime) / 1000).toFixed(1);
  processingTimeEl.textContent = `${lengthMonths} month${lengthMonths === 1 ? "" : "s"} loaded in ${elapsedSeconds}s`;
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

  if ("serviceWorker" in navigator) {
    navigator.serviceWorker.register("service-worker.js").catch((error) => {
      console.error("Service worker registration failed:", error);
    });
  }
}

main();
