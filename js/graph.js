// Microsoft Graph calls: discover calendars and pull Horizon-tagged events.

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";
const DEFAULT_TAG = "horizon";
// Fetching all calendars fully in parallel triggers Graph's 429 rate limiting;
// fetching one at a time is safe but slow. A small worker pool balances both.
// Using 2 (vs. an odd number like 3) means a batch of calendars divides more
// evenly, so the last request rarely ends up running alone with no overlap
// to mask its natural duration, which otherwise looks like a stall.
const CONCURRENCY_LIMIT = 2;

async function graphGet(token, url, extraHeaders = {}, retriesLeft = 3) {
  const response = await fetch(url, {
    headers: {
      Authorization: `Bearer ${token}`,
      ...extraHeaders
    }
  });

  if (response.status === 429 && retriesLeft > 0) {
    const retryAfterSeconds = Number(response.headers.get("Retry-After")) || 2;
    await new Promise((resolve) => setTimeout(resolve, retryAfterSeconds * 1000));
    return graphGet(token, url, extraHeaders, retriesLeft - 1);
  }

  if (!response.ok) {
    throw new Error(`Graph request failed (${response.status}): ${url}`);
  }

  return response.json();
}

async function graphGetAllPages(token, initialUrl, extraHeaders = {}) {
  const items = [];
  let url = initialUrl;

  while (url) {
    const page = await graphGet(token, url, extraHeaders);
    items.push(...(page.value || []));
    url = page["@odata.nextLink"] || null;
  }

  return items;
}

export async function getCalendars(token) {
  const url = `${GRAPH_BASE}/me/calendars?$select=id,name`;
  return graphGetAllPages(token, url);
}

function hasHighlightCategory(categories, categoryNames) {
  if (categoryNames.length === 0) return false;
  const eventCategories = (categories || []).map((category) => category.trim().toLowerCase());
  return categoryNames.some((name) => eventCategories.includes(name));
}

function escapeRegExp(text) {
  return text.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

// Matches a literal "#tag" token (word-bounded so "#horizon" doesn't match "#horizons")
// anywhere in the subject, location, or a short preview of the description.
function hasHighlightTag(event, tagNames) {
  if (tagNames.length === 0) return false;

  const searchableText = [
    event.subject,
    event.location && event.location.displayName,
    event.bodyPreview
  ]
    .filter(Boolean)
    .join(" ")
    .toLowerCase();

  return tagNames.some((tag) => new RegExp(`#${escapeRegExp(tag)}\\b`).test(searchableText));
}

// Removes only the searched-for "#tag" tokens from display text, leaving any other hashtags intact.
function stripMatchedTags(text, tagNames) {
  if (!text || tagNames.length === 0) return text;

  let result = text;
  for (const tag of tagNames) {
    result = result.replace(new RegExp(`#${escapeRegExp(tag)}\\b`, "gi"), "");
  }
  return result.replace(/\s{2,}/g, " ").trim();
}

export async function getHighlightEvents(token, calendars, startISO, endISO, onProgress, categoryNames = [], tagNames = [DEFAULT_TAG]) {
  const timeZone = Intl.DateTimeFormat().resolvedOptions().timeZone;
  const headers = { Prefer: `outlook.timezone="${timeZone}"` };
  const normalizedCategories = categoryNames.map((name) => name.trim().toLowerCase());
  const normalizedTags = tagNames.map((tag) => tag.trim().toLowerCase());

  const eventsByCalendar = new Array(calendars.length);
  let nextIndex = 0;
  let completed = 0;

  async function worker() {
    while (nextIndex < calendars.length) {
      const currentIndex = nextIndex++;
      const calendar = calendars[currentIndex];

      const url =
        `${GRAPH_BASE}/me/calendars/${calendar.id}/calendarView` +
        `?startDateTime=${encodeURIComponent(startISO)}` +
        `&endDateTime=${encodeURIComponent(endISO)}` +
        `&$select=subject,start,end,categories,location,isAllDay,bodyPreview` +
        `&$top=100`;

      const events = await graphGetAllPages(token, url, headers);

      eventsByCalendar[currentIndex] = events
        .filter(
          (event) =>
            hasHighlightCategory(event.categories, normalizedCategories) ||
            hasHighlightTag(event, normalizedTags)
        )
        .map((event) => ({
          subject: stripMatchedTags(event.subject, normalizedTags),
          start: new Date(event.start.dateTime),
          end: new Date(event.end.dateTime),
          isAllDay: event.isAllDay,
          categories: event.categories,
          location: stripMatchedTags(event.location && event.location.displayName, normalizedTags),
          calendarName: calendar.name
        }));

      completed++;
      if (onProgress) {
        onProgress({ completed, total: calendars.length });
      }
    }
  }

  const workers = Array.from({ length: Math.min(CONCURRENCY_LIMIT, calendars.length) }, worker);
  await Promise.all(workers);

  return eventsByCalendar.flat().sort((a, b) => a.start - b.start);
}
