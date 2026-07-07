# argotrack (beartime, Argonauts kiosk)

Attendance kiosk for FRC team 8728 Argonauts. It combines the beartime
Google-Sheets backend with a PIN-keypad kiosk frontend, plus session typing
(Meeting / Outreach / Competition) and a live "who's signed in" panel.

Deployed separately from the original beartime app: **argotrack.filipkin.com**
(Coolify), off the `argotrack` branch. The `main` branch and
`beartime.app.filipkin.com` are untouched.

## Endpoints

| Route | Purpose |
|-------|---------|
| `GET /punch?pin=&sessionType=&eventName=` | Toggle a user in/out. Records a raw `Sessions` row; on sign-out also adds per-day hours to `Log`. `/login` is a back-compat alias. |
| `GET /signedin` | `{ success, mentors:[], students:[] }` — kiosk polls this every 30s. |
| `GET /stats?pin=` | Mentor-only attendance stats (unchanged from beartime). |
| `GET /health` | Pings the sheet. 200 when healthy, 503 when the connection is wedged (Coolify uses this to auto-restart). |

## Spreadsheet layout (3 tabs, order matters — read by index)

0. **Users** — `pin, fname, lname, email, type, gender, login, logout, hours, total, loggedin, sessionType, sessionName`. `type` = `MENTOR` or `STUDENT`. This is the roster you edit.
1. **Log** — date-pivot: `pin, fname, lname` then one column per day (hours). Feeds `/stats`.
2. **Sessions** — raw append log: `timestamp, pin, name, type, event(IN/OUT), sessionType, eventName, hours`.

## The "stops working after ~a week" fix

The old app was restarted every 24h by a cron to work around it silently dying.
Root causes were fire-and-forget sheet saves (unhandled rejections) and a
connection cached once at startup with no recovery. This version:

- awaits every save / `saveUpdatedCells`,
- caches the connection but **clears it on failure** so the next request reconnects,
- wraps each sheet op in `withReconnect` (one automatic retry),
- exposes `/health` that actually probes the sheet, and
- runs a self-watchdog that `process.exit(1)`s after 15 min of being unhealthy
  so Docker restarts a fresh process — only when actually broken, not on a timer.

No external restart cron is needed anymore.

## Provisioning a new sheet

Service accounts can't create Drive files, so create a blank Google Sheet in a
human Drive, share it (Editor) with the service account, then run the
provisioning script (kept outside the repo) with `SPREADSHEET_ID` + a Coolify
token to create the three tabs.
