# Security

## Where secrets live

This project never stores credentials in code or in the repository. Everything
sensitive is kept in **Apps Script → Project Settings → Script Properties**,
which is private to the Google account that owns the script:

- `BASECAMP_CLIENT_ID`, `BASECAMP_CLIENT_SECRET`
- `BASECAMP_ACCESS_TOKEN`, `BASECAMP_REFRESH_TOKEN` (written automatically after you authorize)

The script never logs tokens, and the **Sync Log** sheet never records credentials.
Authentication is OAuth 2.0 only - the script never asks for, or sees, a Basecamp or Google password.

## Before you push to Git

- Do not paste a client secret, token, or `.clasprc.json` into any file you commit.
- `.gitignore` already excludes `.clasp.json`, `.clasprc.json`, `.env*` and common credential file names.
- If a secret is ever committed, treat it as leaked: regenerate the client secret at
  <https://launchpad.37signals.com/integrations>, then rewrite Git history or delete the repository.

## Data exposure

Anyone with edit access to the Google Sheet can open **Extensions -> Apps Script** and read Script
Properties. Only share the spreadsheet with people you would trust with that Basecamp access.
The Web App is deployed as "Anyone" so Basecamp's OAuth redirect can reach it; its only job is to
receive the one-time OAuth `code` and exchange it for tokens.

**Known limitation:** the OAuth callback does not validate a `state` value. Anyone who knows your
`/exec` URL *and* your integration's client ID could authorize their own Basecamp account and cause
the script to store their tokens in place of yours (your sheet would then stop syncing, or sync the
wrong account). Treat the `/exec` URL as private, do not publish it, and if the sheet ever starts
showing someone else's tasks, re-run **Setup Basecamp** and rotate the client secret. Adding a
`state` check is a welcome contribution once it is confirmed that Basecamp's Launchpad echoes it back.

## Reporting a vulnerability

Please open a private security advisory on the repository (or contact the maintainer directly)
rather than a public issue.
