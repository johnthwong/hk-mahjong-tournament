# Data model & tournament architecture

Reference notes for how this app stores and switches tournament data. Not a
skill — this is project background to read before touching `Code.gs` data flow.

## Platform

Google Apps Script web app, managed locally with `clasp`.

- `Code.gs` — server-side code; `doGet(e)` (line ~5) is the web-app entry point.
- `index.html` — player portal. `admin.html` — admin portal (`?portal=admin`).
- `appsscript.json` — manifest. Must contain a `webapp` block, or deployments
  fail with "No web app entry point found":
  ```json
  "webapp": { "executeAs": "USER_DEPLOYING", "access": "ANYONE_ANONYMOUS" }
  ```
- `.clasp.json` is **not** committed (local-only); a fresh clone has none. It must
  point at the **container-bound** script (see below).

### Deploy workflow

```
clasp push                 # upload Code.gs + HTML + manifest
clasp create-deployment    # cut a versioned (@N) deployment
clasp open-web-app         # open a deployment in the browser
```

- `@HEAD` always serves the latest pushed code — use it while developing.
- `@N` versioned deployments are frozen snapshots; they do **not** update on push.

## The master / tournament split

This is the core of the design.

- **Master sheet** — a permanent hub. The Apps Script project is **container-bound**
  to it, so all `SpreadsheetApp.getActiveSpreadsheet()` calls resolve to the master.
  Because it's bound, the script *must* live inside this sheet (Extensions → Apps
  Script), not as a standalone script — a standalone script has no active
  spreadsheet and throws "Sorry, unable to open the file at this time."
  The master holds the `Settings` tab and a pointer to the active tournament. It is
  never overwritten by tournament operations.
- **Tournament sheets** — each tournament is its **own separate Google Sheet**,
  created on demand and stored in the **same Drive folder** as the master.

### How the active tournament is resolved

`getDataSS()` (line ~54):
1. Reads the master's `Settings` tab for the `Active_Tournament_ID` row.
2. If set, `openById(targetId)` returns that tournament sheet (cached in
   `_cachedDataSS`).
3. If empty/missing/unopenable, **falls back to the master sheet itself.**

Implication: if `Active_Tournament_ID` points at a sheet you can't access (e.g. a
stale ID from another owner after a reclone), clear the cell — the app then reads
and writes directly in the master until you create a real tournament.

### Create New Tournament — `startNewTournament()` (line ~235)

1. `SpreadsheetApp.create()` makes a brand-new sheet, moved into the master's
   parent Drive folder.
2. Adds the standard tabs: `Players`, `Settings`, `Penalties_List`, `Scores`,
   `Penalties`, `Pairings`, `Leaderboard`.
3. Repoints the master's `Active_Tournament_ID` / `Active_Tournament_Name` at the
   new file.

**The previous tournament is never deleted, trashed, or modified** — only the
pointer moves. Old tournaments stay archived in the folder.

### Listing & switching

- `getTournamentList()` (line ~209) scans the master's **parent Drive folder** and
  returns every Google Sheet except the master. So any Sheet dropped in that folder
  appears as a "tournament" — keep the folder clean (don't leave the master in My
  Drive root alongside unrelated Sheets).
- `switchTournament(fileId)` (line ~225) just rewrites `Active_Tournament_ID`.
  Switching is non-destructive and fully reversible.
- There is **no in-app delete** — remove old tournament sheets from Drive manually.

## Schema lives in two places

Tab and column names are hardcoded as string literals in `Code.gs` (e.g.
`"Players"`, `"Settings"`, `"Player ID"`, `"Name"`, `"ARA ID"`,
`"Active_Tournament_ID"`, `"Penalties_List"`). Renaming a column or tab requires
updating **both** the sheet and every literal in the code (and any in
`index.html` / `admin.html`).

### Players sheet

Columns: `Player ID`, `Name`, `Checked In`, `ARA ID`.

- **Player IDs must be text shaped like `P1`** (a `P` followed by digits), or
  `SUB1` for substitutes. The UI strips the prefix with
  `String(id).replace(/^P/i,'')` and compares IDs as strings. Bulk import
  (`addPlayersBulk`) auto-generates correct `P`-prefixed IDs (`"P" + nextNum`); the
  bad-data case is usually leftover numeric IDs read from the master fallback.
  `index.html` validates ID format on load (`/^(?:P|SUB)\d+$/i`) and throws a clear
  message otherwise.
- Bulk import format: one player per line, `Name - ARA ID` (ARA ID optional). The
  parser splits on `' - '`, so names must not contain ` - `.

### Substitutes (SUB players)

When the active player count isn't a multiple of 4, the app can auto-add filler
players to complete the last table.

- Auto-added subs get **ID `SUB<n>` and name `SUBSTITUTE <n>`** (e.g. `SUB1` /
  `SUBSTITUTE 1`), created in both `generateNextRound` and the round-1 repair flow.
  Numbering is taken from the max existing `SUB<n>` ID so removing a sub can't cause
  an ID collision.
- **SUB detection must use the ID prefix or the name prefix consistently.** Name
  checks use `name.toUpperCase().startsWith("SUB")` (matches `SUBSTITUTE`); ID checks
  use `id.startsWith("SUB")`. Earlier auto-subs were created with `P<n>` IDs and
  `SUB <n>` names, so the ID-based skip in `countRepeats` (lines ~909/915) silently
  never matched — fixed by giving subs `SUB<n>` IDs.
- **The "add subs?" prompt fires on every round**, not just round 1
  (`admin.html`): if the active (non-`[DNF]`) count isn't divisible by 4 it offers to
  add the needed subs. The server (`generateNextRound`, `addSubs` arg) only pads when
  the client passes `addSubs = true`.
- **Known limitation (not yet changed): subs are still full participants in scoring,
  standings, and Swiss seeding.** `getStandingsData` and the Swiss `ranked` sort
  include every non-`[DNF]` player with no SUB check, so a sub accrues points, earns
  a rank, and can be seeded into a top bucket — displacing a real contestant. Making
  subs neutral fillers would require filtering them out of standings ranking and the
  seeding sort. The `countRepeats` fix above only stops them being counted as repeat
  opponents.

## Swiss pairing — how ties at a bracket boundary resolve

`generateNextRound` (Swiss branch, line ~758) sorts all players by score
descending (`ranked.sort((a,b) => b.pts - a.pts)`), then `sliceIntoSwissBuckets`
(line ~952) cuts that sorted array into **contiguous index slices** — bucket 0 is
the top positions, bucket 1 the next, etc. Bracket membership is purely positional
in the sorted array.

When two players tie on score at a bucket boundary (one slot left in the higher
bracket), the comparator returns 0, and JS's `Array.prototype.sort` is **stable**,
so tied players keep their relative order from the source `players` array (Players
sheet / check-in order). Therefore:

- The tied player listed **earlier in the pool** takes the **last slot in the
  higher bracket**; the other drops to the start of the next bucket.
- This is **list order, not a competitive tiebreaker** — no head-to-head, fewest
  byes, or secondary stat is applied. (The `Tiebreaker_Rule` setting governs in-game
  score *splitting*, not pairing order.)
- Seating **within** a bucket is randomized (`.sort(() => Math.random() - 0.5)`),
  but that shuffle never moves a player across a bucket boundary.

## Swap / substitute players (mid-round seat edits)

`swapPairings(round, p1Id, p2Id, force)` (`Code.gs`) edits seats in an existing
round's `Pairings` rows. The admin UI offers **one all-players picker per side**
(`renderSwapPlayerLists` in `admin.html`); each option is annotated with the
player's seat for that round (`Table N`) or `(bench)`. Tables are **derived** from
the pairings, not chosen by the user.

The backend decides the operation from who is seated:

- **Both seated → swap** — exchange the two players' seats (the original behavior).
- **One seated + one benched → replace** — write the benched player's ID into the
  seated player's cell; the seated player is left unseated (benched). This is how you
  bring an unseated substitute into a seat vacated mid-round. It does **not**
  auto-DNF the outgoing player — the success message reminds you to mark them DNF if
  they withdrew.
- **Neither seated → error.**

Guards (raise a warning the caller can confirm past with `force=true`):

- **Repeat opponents** — if the incoming player has already faced anyone at the
  target table (`getHistoryMatrix(round)`).
- **Already-scored table** — scores live in the `Scores` sheet keyed to the seat's
  player ID *at submission time*, so editing a seat after scoring does not move the
  score. The warning tells you to re-enter the score for the incoming player.

Note: substitutes added through the normal Add Player / Bulk Import UI get `P<n>`
IDs (only the auto-add path uses `SUB<n>`), so a manually-added bench player swaps
fine but isn't recognized by `SUB`-aware logic.
