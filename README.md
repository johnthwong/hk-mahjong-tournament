# HK Mahjong Tournament

A Google Apps Script web app for running Hong Kong mahjong tournaments
(pairings, scorekeeping, faan table, standings). The code lives in this repo and
is pushed to Apps Script with [`clasp`](https://github.com/google/clasp); the app
runs on Google's servers and reads/writes Google Sheets in your Drive.

- `Code.gs` — server-side logic (`doGet` is the web-app entry point).
- `index.html` — player portal. `admin.html` — admin portal (`?portal=admin`).
- `appsscript.json` — manifest (must keep its `webapp` block).
- `MASTER Tournament Sheet.xlsx` — seed spreadsheet you upload to Drive.

## How it's structured (important)

The Apps Script project is **container-bound** to a "master" Google Sheet — i.e.
the script lives *inside* that sheet (Extensions → Apps Script), not as a
standalone project. The code relies on `SpreadsheetApp.getActiveSpreadsheet()`, so
a standalone script won't work. Each **tournament is its own Google Sheet** created
by the app in the same Drive folder as the master; the master just holds settings
and a pointer to the active tournament.

## Setup

### 1. Prerequisites
- **Node.js 18+** and npm.
- Install clasp globally:
  ```bash
  npm install -g @google/clasp
  ```

### 2. Authenticate clasp
```bash
clasp login
```
This opens a browser to authorize clasp against your Google account (the account
that will *own* the tournament data). Creates `~/.clasprc.json`.

### 3. Create the master spreadsheet + its bound script
1. Upload **`MASTER Tournament Sheet.xlsx`** to Google Drive (ideally into a
   dedicated folder — the app lists tournaments by scanning the master's folder).
2. Open it and **File → Save as Google Sheets** (it must be a native Google Sheet,
   not `.xlsx`).
3. In that Sheet: **Extensions → Apps Script**. This creates a script **bound** to
   the sheet. In the editor, open **Project Settings (⚙) → IDs** and copy the
   **Script ID**.

### 4. Link this folder to that script
`.clasp.json` is not committed (it's per-user). Create it in the repo root:
```json
{ "scriptId": "PASTE_SCRIPT_ID_HERE", "rootDir": "." }
```

### 5. Push the code
```bash
clasp push -f
```
Uploads `Code.gs`, `index.html`, `admin.html`, and `appsscript.json` to the bound
script. (`-f` skips the manifest-overwrite prompt.)

> If `clasp create-script` ever rewrites `appsscript.json`, make sure it keeps:
> ```json
> "webapp": { "executeAs": "USER_DEPLOYING", "access": "ANYONE_ANONYMOUS" }
> ```
> Without the `webapp` block, deployment fails with "No web app entry point found".

### 6. Authorize the script's scopes
The first run needs you to grant Sheets/Drive permissions:
```bash
clasp open-script
```
In the editor, pick any function and **Run** once. Approve the OAuth prompt
(**Advanced → Go to project (unsafe)** is normal for your own script → Allow).

### 7. Deploy and open
```bash
clasp create-deployment     # creates a versioned (@N) web-app deployment
clasp open-web-app          # pick the @N deployment to open it
```
- The **player portal** is the deployment's `/exec` URL.
- The **admin portal** is the same URL with `?portal=admin`.
- Use the `@HEAD` deployment while developing (it always serves the latest
  `clasp push`); cut a new `@N` deployment for a stable shareable link.

### 8. First run
On a fresh master, the `Settings` tab's `Active_Tournament_ID` may be empty or
stale. Either is fine — the app falls back to the master sheet. Open the admin
portal → **Select Tournament File → Create New Tournament** to make a real
tournament (it prompts for a name, which becomes the Drive file name).

## Admin PIN

The admin screen can be gated by a PIN (Settings → **Admin PIN**). The PIN and a
signing secret are stored in **Script Properties** (owner-only; not in the Sheet or
the page). A correct PIN issues a 30-day token kept in the browser's
`localStorage`, so a known browser isn't re-prompted. Leave the PIN blank to
disable the gate. **Recovery:** clear the `ADMIN_PIN` property in the Apps Script
editor (Project Settings → Script Properties) if you ever lock yourself out.

> Note: the PIN currently gates the **UI only** (a deterrent); it is not yet
> enforced per server function. See `wishlist.md`.

## Everyday workflow
- Edit code locally → `clasp push -f` → refresh the `@HEAD` web-app URL.
- To open the deployed app in a browser: `clasp open-web-app`.
- See `.claude/architecture.md` for how the data model, pairings, subs, faan
  scoring, and the PIN gate work.
