# Family Horizon

A personal dashboard that answers one question: **"What important things are coming up over the next several months?"**

This is not a calendar replacement. Outlook stays the source of truth for all event details. Family Horizon just pulls out the events your family has flagged as important and lists them in one simple, chronological view.

## How it works

An event counts as a "Horizon" event if either of these is true:

1. It has the Outlook category **`Horizon`** (any case: `horizon`, `HORIZON`, etc. all work), **or**
2. Its subject, location, or description preview contains the hashtag **`#horizon`** (case-insensitive). This is useful for calendars where categories can't be applied (see Known limitations below), since hashtags are just plain text anyone can type. Microsoft Graph supplies a short preview rather than the complete description.

Open the dashboard, sign in with your Microsoft account, and it discovers every calendar you have access to, pulls matching events over the next 6 months (configurable, see below), and lists them grouped by month.

Select an event to expand its full formatted description. Descriptions and inline images load when opened; select the copy icon beside a location to copy it to the clipboard.

Routine/recurring events (practices, ordinary appointments) are left untagged and simply won't appear, only the "yes, family needs to know about this" events show up.

## Running it locally

This is a static site (HTML/CSS/vanilla JS), no build step, no backend. From this folder, run:

```powershell
.\serve.ps1
```

Then open `http://localhost:5500` in your browser. `Ctrl+C` in the terminal to stop the server.

The port must be `5500` to match the redirect URI registered in Azure for this app's sign-in.

## URL options

These are query string parameters you can bookmark, so you don't have to re-type them:

| Parameter  | Example                              | Effect                                                                 |
|------------|---------------------------------------|-------------------------------------------------------------------------|
| `exclude`  | `?exclude=Anniversaries,US Holidays` | Skips the named calendar(s) entirely (case-insensitive, comma-separated) |
| `category` | `?category=Horizon,Major`             | Also matches events tagged with any of the listed Outlook categories (case-insensitive, comma-separated, empty/off by default) |
| `tags`     | `?tags=horizon,important`             | Matches events with any of the listed `#hashtags` in subject/location/description preview (case-insensitive, comma-separated, spaces stripped from each tag, defaults to `horizon`) |
| `length`   | `?length=3`                           | How many months ahead to look (any positive integer, defaults to 6)     |

Example combining all four:

```
http://localhost:5500/?exclude=Anniversaries,US Holidays&category=Horizon&tags=horizon&length=3
```

The `exclude` parameter exists because some calendars (read-only system calendars, shared calendars you don't own, etc.) can't be tagged with a category or edited at all. Excluding them by name keeps the dashboard fast without losing any real Horizon events, since those calendars could never contribute any.

## Updating the code without seeing a hard refresh

This app installs itself as a Progressive Web App, which means a service worker caches its files for offline/fast loading. Whenever the code changes, the cache name in `service-worker.js` (`CACHE_NAME`) is bumped so browsers pick up the new version automatically on the next load. If you're actively developing and changes don't seem to show up, do a hard refresh (`Ctrl+F5`) or unregister the service worker via DevTools → Application → Service Workers.

The load summary shows the installed cache version and current source version as `installed.current` (for example, `54.55`). The version file bypasses the service worker cache so it can reveal stale app caches.

## File structure

```
index.html          Page shell, sign-in button, dashboard container
css/styles.css       All styling, including light/dark theme variables
js/auth.js           MSAL config and sign-in/token logic
js/graph.js          Microsoft Graph calls: calendar discovery + event fetching/filtering
js/app.js            UI wiring: rendering, progress bar, URL option parsing
app-version.json     Current source version marker, fetched outside the service worker cache
manifest.json        PWA metadata (name, icons, display mode)
service-worker.js    Offline/app-shell caching
assets/              PWA icons
serve.ps1            Zero-dependency local static file server (PowerShell)
```

## Known limitations

- The Outlook "Family" group calendar (a special Microsoft consumer calendar type) can't have categories applied to its events through any client. Adding a `#horizon` hashtag to the subject, location, or description preview works around this, as long as the calendar itself allows editing at all. Otherwise, tag events on a regular shared calendar instead, or use `?exclude=` to skip it.
- Calendars with many recurring series (e.g., weekly practices) take longer to fetch, since Microsoft Graph expands every recurring instance within the date range before this app filters for a match.
