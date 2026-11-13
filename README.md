# 1AV Project Health Monitor

A dashboard for the 1AV Power Apps and Power Automate projects, hosted on GitHub Pages.

- The page contains **no data**. After a user signs in with their company Microsoft account, it calls a Power Automate flow and shows what the flow returns.
- It refreshes every 15 minutes, and there is a **Refresh now** button.
- New projects and new flows appear automatically (see *Auto-discovery* below).

```
GitHub Pages (index.html)
   │  1. user signs in (Microsoft / Okta SSO)
   │  2. POST + Bearer token
   ▼
Power Automate "API" flow  (When an HTTP request is received, Any user in my tenant)
   │  3. reads latest snapshot (fast, < 2 s)
   ▼
SharePoint list "MonitorSnapshot"  ◄── Power Automate "Collector" flow, every 15 min
                                        reads Projects config → finds flows by prefix
                                        → counts runs/failures → reads app activity
```

## Files

| File | Purpose |
|---|---|
| `index.html` | Page layout |
| `assets/app.js` | Sign-in, data loading, rendering |
| `assets/styles.css` | 1AV colors, Noto Sans |
| `assets/vendor/msal-browser.min.js` | Microsoft sign-in library v3.30.0, included so there's no CDN dependency |
| `config.js` | Client ID, tenant ID, API flow URL |
| `data/sample.json` | Demo data. Also the reference for the JSON format the API flow must return |

## Step 1: Publish on GitHub Pages (demo mode)

1. Create a repo, for example `1av-project-monitor`, and upload all the files.
2. Go to **Settings → Pages → Source: Deploy from a branch → main / root**.
3. Open `https://<your-account>.github.io/1av-project-monitor/`. It runs in demo mode while `apiUrl` in `config.js` is empty.

You can open `?demo=1` at any time to see sample data.

## Step 2: App registration (Entra ID)

Configure your existing app registration (the client ID):

1. **Authentication → Add a platform → Single-page application**. Set the redirect URI to the exact Pages URL, with the trailing slash, for example `https://<account>.github.io/1av-project-monitor/`.
2. **API permissions → Add → APIs my organization uses → Power Automate** (Flow Service). Add the **Delegated → User** permission, then grant admin consent if your tenant requires it.
3. Put the **client ID** in `config.js`. Do **not** create or commit a client secret; a single-page app doesn't use one.

## Step 3: Power Automate (next phase)

### SharePoint lists

**MonitorProjects** holds the configuration, one row per project. Add a row and the project appears on the dashboard.

| Column | Example |
|---|---|
| Title (project name) | Baggage Tracking |
| Key | bts |
| FlowPrefix | 1AVBTS |
| SiteUrl | https://cebupacificairph.sharepoint.com/sites/1AVITAutomation_BaggageDeliverySystem |
| ActivityList | the list the app writes to, used for "app last used" |
| Active | Yes |

**MonitorSnapshot** holds a single row. The `Json` column (multi-line text) contains the latest JSON.

**MonitorErrors** is optional. Flows write to it from a catch scope when they fail.

### Collector flow (Recurrence, every 15 minutes)

1. Get the items from **MonitorProjects** where `Active = Yes`.
2. List the flows this account can see. Keep those whose display name starts with `FlowPrefix`. This is the **auto-discovery** step: any new `1AVBTS…` flow is included on the next run.
3. For each flow, get the runs from the last 24 hours with *HTTP with Microsoft Entra ID (preauthorized)*. Count succeeded and failed runs per hour, and keep the latest error.
4. Read app activity from each project's `ActivityList`.
5. Build the JSON in the same format as `data/sample.json` and save it to **MonitorSnapshot**.

### API flow

1. Trigger: **When an HTTP request is received**, with **Who can trigger the flow = Any user in my tenant**.
2. Get the **MonitorSnapshot** item.
3. Return it with a **Response** action: status `200`, body = `Json`, and these headers:
   - `Content-Type: application/json`
   - `Access-Control-Allow-Origin: https://<account>.github.io`
4. Copy the trigger URL into `config.js` → `apiUrl`.

## JSON format

The source of truth is `data/sample.json`.

```json
{
  "generatedAt": "2026-11-13T12:24:00Z",
  "windowHours": 24,
  "services": [{ "name": "Outlook", "status": "ok|advisory|degraded|down", "note": "" }],
  "projects": [{
    "key": "bts", "name": "Baggage Tracking", "prefix": "1AVBTS",
    "status": "healthy|warning|failing|stale",
    "uptimePct": 91.4, "runs": 163, "failed": 14,
    "hourly": [[ok, failed], "... 24 buckets, oldest first"],
    "flows": [{ "name": "", "runs": 0, "failed": 0, "lastRun": "ISO", "lastStatus": "", "lastError": "" }],
    "apps":  [{ "name": "", "lastUsed": "ISO", "activeHours": 17 }],
    "checks":[{ "label": "", "value": "", "status": "healthy|warning|failing" }],
    "summary": "one line"
  }],
  "incidents": [{ "time": "ISO", "project": "", "source": "", "message": "", "severity": "failing|warning" }]
}
```

## Privacy

- The GitHub Pages site is public, but it holds **no data**. Data is returned only to signed-in users in the 1AV / Cebu Pacific tenant.
- Keep the repo free of real snapshots. `data/sample.json` is sanitized demo data.
- Even with sign-in in place, keep staff names and exact locations out of the collector JSON.
