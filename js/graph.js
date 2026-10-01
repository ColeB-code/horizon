// Microsoft Graph calls: discover calendars and pull Horizon-tagged events.

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";
const DEFAULT_TAG = "horizon";
// Fetching all calendars fully in parallel triggers Graph's 429 rate limiting;
// fetching one at a time is safe but slow. A small worker pool balances both.
// Using 2 (vs. an odd number like 3) means a batch of calendars divides more
// evenly, so the last request rarely ends up running alone with no overlap
// to mask its natural duration, which otherwise looks like a stall.
const CONCURRENCY_LIMIT = 2;

function graphErrorDetails(error) {
  return {
    name: error.name,
    message: error.message,
    status: error.status,
    code: error.code,
    requestId: error.requestId,
    clientRequestId: error.clientRequestId,
    requestPath: error.requestPath
  };
}

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
    const errorResponse = await response.json().catch(() => null);
    const graphError = errorResponse?.error;
    const errorMessage = graphError?.message || response.statusText;
    const error = new Error(`Graph request failed (${response.status}): ${errorMessage}`);
    error.status = response.status;
    error.code = graphError?.code;
    error.requestId = response.headers.get("request-id");
    error.clientRequestId = response.headers.get("client-request-id");
    const requestUrl = new URL(url);
    error.requestPath = `${requestUrl.pathname}${requestUrl.search.replace(/contentBytes/gi, "contentBytes")}`;
    throw error;
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

function summarizeImageSource(source) {
  const value = source.trim();
  if (/^cid:/i.test(value)) {
    return { scheme: "cid", contentId: value.replace(/^cid:/i, "").replace(/^<|>$/g, "") };
  }
  if (/^data:/i.test(value)) {
    return { scheme: "data", contentType: value.slice(5).split(/[;,]/)[0] || "unknown" };
  }

  try {
    const url = new URL(value, "https://horizon-image-diagnostics.invalid");
    return {
      scheme: url.protocol.slice(0, -1),
      host: url.host,
      path: url.pathname,
      hasQuery: Boolean(url.search)
    };
  } catch {
    return { scheme: "unparsed", characterCount: value.length };
  }
}

function inspectImageMarkup(content) {
  const body = new DOMParser().parseFromString(content || "", "text/html").body;
  const candidates = Array.from(body.querySelectorAll("*")).filter((element) => {
    const tagName = element.tagName.toLowerCase();
    const style = element.getAttribute("style") || "";
    return tagName === "img" || tagName.includes("imagedata") ||
      element.hasAttribute("background") || /url\s*\(/i.test(style);
  });

  return candidates.map((element) => {
    const sources = [];
    for (const attribute of ["src", "srcset", "href", "data", "background"]) {
      const value = element.getAttribute(attribute);
      if (!value) continue;
      const values = attribute === "srcset"
        ? value.split(",").map((candidate) => candidate.trim().split(/\s+/)[0]).filter(Boolean)
        : [value];
      sources.push(...values.map((source) => ({ attribute, ...summarizeImageSource(source) })));
    }

    const style = element.getAttribute("style") || "";
    for (const match of style.matchAll(/url\(\s*(['"]?)(.*?)\1\s*\)/gi)) {
      sources.push({ attribute: "style-url", ...summarizeImageSource(match[2]) });
    }

    return {
      tag: element.tagName.toLowerCase(),
      sources,
      width: element.getAttribute("width"),
      height: element.getAttribute("height"),
      hasAlt: element.hasAttribute("alt"),
      altTextLength: element.getAttribute("alt")?.length || 0
    };
  });
}

export async function getEventDescription(token, calendarId, eventId) {
  const eventUrl =
    `${GRAPH_BASE}/me/calendars/${encodeURIComponent(calendarId)}/events/${encodeURIComponent(eventId)}`;
  const event = await graphGet(token, `${eventUrl}?$select=body`);
  const inlineAttachments = [];
  const imageDiagnostics = {
    eventId,
    calendarId,
    bodyContentType: event.body?.contentType,
    bodyCharacterCount: event.body?.content?.length || 0,
    cidReferenceCount: (event.body?.content?.match(/\bcid:/gi) || []).length,
    imageMarkup: inspectImageMarkup(event.body?.content || ""),
    attachmentListQueried: false,
    attachmentListCount: null,
    attachmentList: [],
    imageFetches: [],
    attachmentListError: null
  };

  try {
    imageDiagnostics.attachmentListQueried = true;
    const attachmentsUrl =
      `${eventUrl}/attachments?$select=id,name,contentType,isInline,size`;
    const attachments = await graphGetAllPages(token, attachmentsUrl);
    imageDiagnostics.attachmentListCount = attachments.length;
    imageDiagnostics.attachmentList = attachments.map((attachment) => ({
      id: attachment.id,
      name: attachment.name,
      contentType: attachment.contentType,
      isInline: attachment.isInline,
      size: attachment.size
    }));
    const inlineImages = attachments.filter(
      (attachment) => attachment.isInline && attachment.contentType?.toLowerCase().startsWith("image/")
    );

    if (imageDiagnostics.cidReferenceCount === 0) {
      imageDiagnostics.imageFetches = inlineImages.map((attachment) => ({
        id: attachment.id,
        name: attachment.name,
        contentType: attachment.contentType,
        isInline: attachment.isInline,
        size: attachment.size,
        outcome: "not fetched: body has no CID image reference"
      }));
    } else {
      const imageResults = await Promise.allSettled(inlineImages.map(async (attachment) => {
        const attachmentUrl =
          `${eventUrl}/attachments/${encodeURIComponent(attachment.id)}`;
        return { ...attachment, ...await graphGet(token, attachmentUrl) };
      }));

      imageResults.forEach((result, index) => {
        const listedAttachment = inlineImages[index];
        if (result.status === "fulfilled") {
          const attachment = result.value;
          inlineAttachments.push(attachment);
          imageDiagnostics.imageFetches.push({
            id: attachment.id,
            name: attachment.name,
            listedContentType: listedAttachment.contentType,
            returnedContentType: attachment.contentType,
            isInline: attachment.isInline,
            size: attachment.size,
            contentId: attachment.contentId,
            contentBytesLength: attachment.contentBytes?.length || 0,
            outcome: "fetched"
          });
        } else {
          const failure = graphErrorDetails(result.reason);
          imageDiagnostics.imageFetches.push({
            id: listedAttachment.id,
            name: listedAttachment.name,
            listedContentType: listedAttachment.contentType,
            isInline: listedAttachment.isInline,
            size: listedAttachment.size,
            outcome: "fetch failed",
            error: failure
          });
        }
      });
    }
  } catch (error) {
    console.warn("Could not load event attachments:", error);
    imageDiagnostics.attachmentListError = graphErrorDetails(error);
  }

  return { body: event.body, inlineAttachments, imageDiagnostics };
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
        `&$select=id,subject,start,end,categories,location,isAllDay,bodyPreview` +
        `&$top=100`;

      const events = await graphGetAllPages(token, url, headers);

      eventsByCalendar[currentIndex] = events
        .filter(
          (event) =>
            hasHighlightCategory(event.categories, normalizedCategories) ||
            hasHighlightTag(event, normalizedTags)
        )
        .map((event) => ({
          id: event.id,
          calendarId: calendar.id,
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
