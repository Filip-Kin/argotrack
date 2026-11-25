# argotrack (Argonauts attendance)

Attendance kiosk, student portal, coach dashboard and volunteer sign-ups for
FRC team 8728 Argonauts, at **argotrack.filipkin.com**.

It is the server version of the coach's Google Apps Script app and runs on the
**same spreadsheet, with the same layout**. Both can be used at the same time:
who is signed in is always read from the last `Log` row for each PIN.

## Deploy

Push to `main`. Coolify builds the `Dockerfile` and deploys. `GET /version`
reports the running commit.

## Spreadsheet

The spreadsheet is the only data store. Its layout belongs to the coach. The
app finds tabs by name and columns by position, and never renames, reorders or
reformats a tab.

| Tab | Columns | App access |
|-----|---------|------------|
| `Roster` | Name, PIN, Role, Student Type, Department, then the coach's own columns ("Drive Team Eligible" is found by its header) | read; rollover sets Student Type |
| `Log` | Timestamp, Student Name, PIN, Event (IN/OUT), Session Duration (minutes), Session Type, Event Name, IP Address | read; append one row per sign-in or sign-out |
| `Summary` | formulas over `Log` and `Roster` | never written (rollover clears it, as the coach's script does) |
| `Seasons` | Season Name, Start Date, End Date, Min Hours New / Veteran / Captain | read |
| `Helpers` | Season, Total Meetings, Total Available Minutes | read |
| `Checklist` | row 1 deadline group, row 2 requirement, row 3 detail, row 4 consequence, row 5+ one student per row (A = PIN, B = name, C+ = status) | read; rollover clears statuses |
| `VolEvents`, `VolSlots` | events and their slots, edited by the coach | read |
| `VolSignups` | Timestamp, Signup ID, Event ID, Slot ID, Person Name, Type, Student Name, Student PIN, Email, Note, Status | read; append sign-ups; set Status to Cancelled |
| `Attendance` | PIN, Name, one column per day | **added by this app**, see below |

PINs are stored both as text (`0716`) and as numbers (`716`). `normPin` pads
both to four digits before comparing. New `Log` rows store the PIN as a number,
like the existing rows, because the `Summary` and `Checklist` formulas match on it.

### The `Attendance` tab

The one tab the app adds, at the end of the spreadsheet. It is a PIN-by-day
grid of hours, rebuilt from `Roster` and `Log` each time the mentor stats page
is opened. A mentor types `E` in an empty cell to mark an excused absence; `E`
marks are the only thing kept from the tab itself.

### Access the service account needs

1. Editor on the spreadsheet.
2. The coach's sheet locks whole tabs. The service account must be added to the
   lock on `Log` (sign-ins), and on `Roster`, `Checklist` and `Summary` for the
   year rollover. Only the owner of the sheet can do that:
   Data, Protect sheets and ranges, pick the tab, Change permissions.

## Endpoints

| Route | Purpose |
|-------|---------|
| `GET /punch?pin=&sessionType=&eventName=` | Sign a person in or out. `/login` is an alias. The sign-out-all PIN signs everyone out. |
| `GET /signedin` | Names currently signed in. The kiosk polls it. |
| `GET /portal?pin=` | Student portal: season hours, attendance, checklist, captain eligibility. |
| `GET /volunteer/events?pin=` | Upcoming volunteer events, slots and the student's own sign-ups. |
| `GET /volunteer/signup?pin=&slotId=&signStudent=&parentName=&parentEmail=&studentEmail=&note=` | Sign up the student, a parent, or both. |
| `GET /volunteer/cancel?pin=&signupId=` | Cancel one of the student's own sign-ups. |
| `GET /stats?pin=` | Mentor PIN only: attendance by person and month, hours by session type. |
| `GET /mentor/roster?mentorPin=` | Mentor PIN only: roster with in/out status. |
| `GET /report`, `/roster`, `/seasons`, `/checklist`, `/signoutall`, `/signout`, `/volunteer/roster`, `/rollover`, `/rollover/status` | Coach password (`?password=`). |
| `GET /health` | Reads one cell. 200 when the sheet is reachable, 503 when not. |
| `GET /version` | Running commit and uptime. |

## Environment

`CLIENT_EMAIL`, `API_KEY` (service account), `SPREADSHEET_ID`, `COACH_PASSWORD`,
`SIGNOUT_ALL_PIN` (default 9999), `TIMEZONE` (fallback only; the spreadsheet's
own timezone is used), `PORT`.

`API_KEY` is stored as a single line. The app turns any run of backslashes
before an `n` into a real newline.

## Health

`/health` probes the sheet. A watchdog exits the process after 15 minutes
without a healthy probe, and Docker starts a fresh one. Reads retry once;
writes never retry, so a sign-in can never be recorded twice.
