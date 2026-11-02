# Basecamp → Google Sheets Sync

A Google Apps Script that pulls your **Basecamp to-dos** into a **Google Sheet** automatically:
tasks assigned to you, and tasks where someone **@mentions you**. Rows are colour-coded by due date,
refresh every 5 minutes, and a fresh mention brings the task back to the bottom of the sheet with the
exact time you were tagged.

It uses the official Basecamp API and OAuth 2.0. No scraping, no passwords.

> Each person runs their **own copy** (own sheet, own script, own Basecamp authorization). Nothing is shared between copies.

---

## Features

- **Direct assignments**: everything assigned to you, including completed ones, via `/my/assignments.json`.
- **@mention detection across the whole account**: comments and to-do descriptions, matched by your
  Basecamp `attachable_sgid` (so typing your name without a real mention never matches).
- **Re-mention bubbles up**: a new mention on a task already in the sheet removes the old row and re-adds
  it at the bottom. Your manual *Working* / *Pending* is kept; a manual *Done* is reset (the assigner
  wants more work).
- **Deep links**: for comment mentions, column D holds the full URL including `#__recording_<id>`, so one
  click lands on the exact comment.
- **Colour by urgency** plus a manual **My Status** dropdown (see below).
- **Activity Time** column: date + time of the assignment, or of the latest mention.
- **Auto-jump**: opening or refreshing the sheet selects the row with the latest Activity Time.
- **Duplicate-safe**: every row is matched by the Basecamp to-do ID; running sync 100 times never duplicates.
- **Your manual rows are never touched** (rows without a Basecamp to-do ID are ignored).
- **Safe by design**: errors go to a `Sync Log` sheet, a failed API call never wipes sheet data,
  overlapping runs are blocked with a script lock, rate limits (HTTP 429) are retried.

## How it works

```
Basecamp API ──► Apps Script (every 5 min) ──► Google Sheet
  /my/assignments.json            │             ├─ rows, colours, links
  /projects/recordings.json       │             ├─ Activity Time
  /todos/{id}.json                │             └─ "Sync Log" sheet
                                  ▼
                 Script Properties (tokens, IDs, checkpoint)
```

Each sync run:

1. Self-heals columns D/O (raw URL vs. "Open Task" link).
2. Fetches your assignments (active + completed).
3. Scans **new** comments and to-dos since the last checkpoint for a mention of you
   (`MENTION_SCAN_LAST_UPDATED_AT`; the first run looks back 30 days).
4. Adds new rows, updates existing ones, moves re-mentioned ones to the bottom.
5. Re-verifies rows that nothing else refreshed (e.g. a task someone completed without commenting).
6. Recomputes every row's colours from its stored due date.

## Sheet layout

| Col | Header | Notes |
|---|---|---|
| A | Assigned Date | Basecamp creation date of the to-do |
| B | Due Date | `dd/MM/yyyy` |
| C | Project Name | Domain-style names are shortened (`example.com.au` → `example`) |
| D | Basecamp Task URL (raw) | Full URL; includes `#__recording_…` for comment mentions |
| E | Work Title | To-do title |
| F | Website Link | First non-Basecamp/non-Google URL in the description, else the project name if it looks like a domain; blank otherwise (never invented) |
| G | Added By | Creator of the to-do |
| H | My Status | Dropdown you control: `Done`, `Pending`, `Working`, `DONE By Assigner` |
| I | Activity Time | Assignment time, or time of the latest mention |
| J–N | Helper columns | Basecamp Todo ID, Project ID, Status, Last Updated, Last Synced (hidden on a fresh install) |
| O | Task ID | Clickable "Open Task" link to the URL in column D |

Column order is defined once in the `COL` constant at the top of `src/Code.gs`.
If you add or move columns, update `COL` to match.

## Colours

Highest priority first:

| Condition | Row colour | Text |
|---|---|---|
| Basecamp says complete (or `DONE By Assigner`) | forest green `#274E13` | white |
| My Status = **Done** | bright green `#96fe0e` | black |
| My Status = **Working** | yellow `#e3ea1d` | red once due/overdue |
| Overdue | dark red `#990000` | white |
| Due today | red `#CC0000` | white |
| 1–2 days left | `#E06666` | black |
| 3–4 days left | `#F9CB9C` | black |
| 5–7 days left | `#FCE5CD` | black |
| More than 7 days, or no due date | light blue `#D9EAF7` | black |

Change them in the `COLOR` and `FONT` constants.

## Requirements

- A Google account and a Google Sheet (a blank one is fine).
- A Basecamp account with API access (Basecamp 3/4 API, `3.basecampapi.com`).
- The sheet's locale should be **day-first** (India, UK, Australia, …): *File → Settings → Locale*.
  Dates are written as `dd/MM/yyyy`, and a US locale would misread them.

## Setup (about 15 minutes)

1. **Create the sheet.** New Google Sheet. Leave row 1 empty; headers are created on the first sync.
2. **Add the code.** *Extensions → Apps Script*. Delete the placeholder, paste `src/Code.gs`, save.
   (Or push it with [clasp](#optional-deploy-with-clasp).) Also open *Project Settings* and tick
   **Show "appsscript.json" manifest file**, then paste in `src/appsscript.json` (set `timeZone` to yours).
3. **Set your timezone** in `Code.gs`: `const TIMEZONE = 'Asia/Kolkata';` (any [IANA name](https://en.wikipedia.org/wiki/List_of_tz_database_time_zones)).
   Also set your contact email in `BASECAMP_USER_AGENT` (Basecamp asks API clients to identify themselves).
4. **Deploy a Web App** (it receives the OAuth redirect):
   *Deploy → New deployment → Web app → Execute as: **Me** → Who has access: **Anyone** → Deploy.*
   Copy the `/exec` URL.
5. **Register a Basecamp integration** at <https://launchpad.37signals.com/integrations>.
   Set its **Redirect URI** to the exact `/exec` URL from step 4. Note the Client ID and Secret.
6. **Add Script Properties** (*Project Settings → Script Properties*):

   | Property | Value |
   |---|---|
   | `BASECAMP_CLIENT_ID` | from step 5 |
   | `BASECAMP_CLIENT_SECRET` | from step 5 |
   | `BASECAMP_REDIRECT_URI` | the exact `/exec` URL from step 4 |

7. **Authorize.** Reload the sheet. In the new **Basecamp Sync** menu: *Setup Basecamp (Step 1: Authorize)*,
   click the link, approve in Basecamp. The account ID and tokens are saved automatically.
8. **Find your IDs.** *Basecamp Sync → Find My Basecamp User ID*. This saves your person ID and the
   mention-detection ID (`MY_ATTACHABLE_SGID`).
9. **First sync.** *Basecamp Sync → Sync Now*.
10. **Automate.** *Basecamp Sync → Create 5-Minute Auto-Sync Trigger*.

Check anything at any time with *Basecamp Sync → Show Configuration Status*.

### Script Properties reference

| Property | Set by | Purpose |
|---|---|---|
| `BASECAMP_CLIENT_ID`, `BASECAMP_CLIENT_SECRET` | you | OAuth integration credentials |
| `BASECAMP_REDIRECT_URI` | you | Must equal the Web App `/exec` URL **and** the Launchpad redirect URI |
| `BASECAMP_ACCOUNT_ID` | script | Detected after authorizing |
| `BASECAMP_ACCESS_TOKEN`, `BASECAMP_REFRESH_TOKEN`, `BASECAMP_TOKEN_EXPIRES_AT` | script | Tokens, refreshed automatically |
| `MY_BASECAMP_PERSON_ID`, `MY_ATTACHABLE_SGID` | script | Your identity (for assignments and mentions) |
| `MENTION_SCAN_LAST_UPDATED_AT` | script | Incremental scan checkpoint |
| `SHEET_NAME` | optional | Tab to use; default is the first tab |
| `EXTRACT_WEBSITE_FROM_PROJECT_NAME` | optional | `false` to stop deriving a website from a domain-like project name |
| `MENTION_SCAN_TIME_BUDGET_SECONDS` | optional | Default `240` |
| `RECONCILE_TIME_BUDGET_SECONDS` | optional | Default `90` |
| `WATCHED_TL_PERSON_IDS` | legacy | No longer used for scanning; kept only for the name-lookup helper |

Keep the two budgets together comfortably under Apps Script's 6-minute limit.

## Menu reference

| Item | What it does |
|---|---|
| Sync Now | Full sync immediately |
| Refresh Colours | Recolours every row from stored data, no API calls |
| Jump to Latest Activity | Selects the row with the newest Activity Time |
| Setup Basecamp (Step 1: Authorize) | Starts OAuth |
| Find My Basecamp User ID | Saves your person ID and mention ID |
| Find a Person's ID by Name | Looks up a colleague's Basecamp ID |
| Test Basecamp Connection | Confirms the token works |
| One-Time: Insert Activity Time Column | Only for sheets created before that column existed |
| Fix Task Link Columns | Repairs D (raw URL) and O ("Open Task") |
| Show Configuration Status | Shows what is set and flags a redirect-URI mismatch |
| Create 5-Minute / Hourly Auto-Sync Trigger, Remove Auto-Sync Triggers | Manage the schedule |

## Day-to-day use

- Set **My Status** in column H. The row recolours instantly.
- When the assigner completes a task in Basecamp, the row turns forest green and H becomes
  `DONE By Assigner`. If it is reopened, that label clears itself.
- Use *Open Task* (column O) or the URL in column D to go to the task.
- Add your own notes in rows or columns the script doesn't manage. Rows with no Basecamp to-do ID are left alone.

## Updating the code

Always update the **same** deployment:
*Deploy → Manage deployments → ✏️ → Version: New version → Deploy.*

Do **not** click *New deployment* for routine changes: it creates a new `/exec` URL and breaks the
redirect URI you registered.

## Adding a teammate

Every person needs their own copy: their own sheet, script, Launchpad integration (recommended, so
redirect URIs never clash), Script Properties, and authorization. Repeat the setup above with their Basecamp login.

## Optional: deploy with clasp

```bash
npm install -g @google/clasp
clasp login
cp .clasp.json.example .clasp.json      # then put your Script ID in it
clasp push                              # uploads src/Code.gs and src/appsscript.json
```

The Script ID is under *Project Settings* in the Apps Script editor. `.clasp.json` is git-ignored.

## Known limitations

- **To-dos only.** Card Table cards and messages are not synced.
- **Polling, not webhooks.** Changes appear within about 5 minutes.
- **Mentions** are detected from the checkpoint onward (first run: last 30 days). Anchored links lost before
  version 1.0.0 cannot be recovered; the next mention restores them.
- A re-mention **deletes and re-creates** the row, so don't rely on fixed row numbers.
- Google quotas apply (6-minute execution limit, URL fetch quota). Very large accounts may need several runs
  to catch up; the Sync Log says when a time budget was hit.
- "Added By" depends on Basecamp returning a `creator` on the to-do; it stays blank if missing.
- The OAuth callback has no `state` check; see [SECURITY.md](SECURITY.md).

## Troubleshooting

| Symptom | Fix |
|---|---|
| **Basecamp Sync** menu missing | Open *Extensions → Apps Script*: there must be one file with `onOpen` (no duplicate copies of the code). Run `onOpen` manually and read the red error in the Execution log. |
| Authorization page errors or hangs | `BASECAMP_REDIRECT_URI`, the Web App `/exec` URL and the Launchpad Redirect URI must be identical. *Show Configuration Status* flags a mismatch. |
| `401 Unauthorized` | Token revoked or expired: run *Setup Basecamp* again. |
| Mentions never appear | Run *Find My Basecamp User ID* again; `MY_ATTACHABLE_SGID` must be set. |
| Dates look wrong / US-style | Set the sheet locale to a day-first region (*File → Settings*). |
| Jump-to-latest does nothing | Make sure column I contains dates; check the Execution log for `onOpen`. |
| Colours stale | Run *Refresh Colours*, or wait for the next sync. |
| Anything else | Read the **Sync Log** sheet (timestamp, action, task ID, result, error). Tokens are never written there. |

## Project structure

```
.
├── README.md
├── LICENSE
├── SECURITY.md
├── CHANGELOG.md
├── .gitignore
├── .claspignore
├── .clasp.json.example
└── src/
    ├── Code.gs            # the whole script
    └── appsscript.json    # Apps Script manifest
```

## Security

Credentials live only in Script Properties, never in the repository. Read [SECURITY.md](SECURITY.md)
before publishing your fork.

## Contributing

Issues and pull requests are welcome. Keep secrets out of commits, and note user-visible changes in `CHANGELOG.md`.

## License

[MIT](LICENSE). Not affiliated with or endorsed by Basecamp / 37signals.
