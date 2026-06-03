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

Subs are a **filler pool**, not permanent participants. The number actually matched
each round is `subsNeeded = (4 - realActive % 4) % 4` — just enough to round the
**real (non-sub) active players** up to a multiple of 4. `isSubPlayer(p)` (a player
whose ID or name starts with `SUB`) is the single detection helper.

- **`generateNextRound`** splits active (non-`[DNF]`) players into reals and subs,
  computes `subsNeeded`, then matches `reals.concat(subs.slice(0, subsNeeded))`.
  **Excess subs are benched (sliced off), not matched.** So if real players drop to a
  multiple of 4 (e.g. 12), all subs are benched automatically. If subs are *short*
  and the caller passed `addSubs=true`, it creates the missing ones.
- Auto-added subs get **ID `SUB<n>` and name `SUBSTITUTE <n>`** (e.g. `SUB1` /
  `SUBSTITUTE 1`), created in both `generateNextRound` and the round-1 repair flow.
  Numbering is taken from the max existing `SUB<n>` ID so removing a sub can't cause
  an ID collision. (Subs added via the normal Add Player / Bulk Import UI get `P<n>`
  IDs and aren't recognized as subs.)
- **`getPairingState`** sizes the bucket preview as `ceil(realActive / 4) * 4` —
  excluding `[DNF]` players and not inflating for excess subs — so the preview matches
  what will actually be paired.
- **The add-subs prompt (`admin.html`) fires only when short on subs**, on any round:
  it compares `subsNeeded` to existing subs and asks to add the difference. When there
  are enough or too many subs, it just confirms generation and the server benches the
  extras.
- **Regenerating a round** (see below) is how you re-apply this to an *already
  generated* round — e.g. drop now-unneeded subs after a mid-round DNF.
- **Known limitation (not changed): matched subs are still full participants in
  scoring, standings, and Swiss seeding.** `getStandingsData` and the Swiss `ranked`
  sort include every matched player with no SUB check, so a matched sub accrues points,
  earns a rank, and can be seeded into a top bucket. The filler-pool logic only
  controls *how many* subs are matched, not whether a matched sub affects standings.

### Regenerating the latest round

`deleteRoundAndScores(roundNum)` (`Code.gs`) clears a round's `Pairings` rows **and
its `Scores` rows**, then the client re-runs generation for that round. Exposed two
ways in `admin.html`:

- **Conflict modal "Redo Pairings"** (`redoPairings`) — when generation hits
  unavoidable repeat matchups.
- **"Regenerate Latest Round" card** (`actionRegenerateLatest`, after the swap card)
  — deletes the latest round (`nextRound - 1`) and re-pairs current active players.
  Deleting a round doesn't touch the roster, so the sub decision is recomputed from
  `DATA.players`; it then reuses the redo path (`actionGenerateRound(true)`). **Scores
  for that round are cleared** — the confirm dialog says so.

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

## Hidden admin UI (penalties) + PIN gate

- **Penalties are hidden, not removed.** The **Penalties nav tab** and the
  **Penalty Configuration** block on the Settings tab are both set to
  `display:none` in `admin.html` (no HK-specific penalties for now). All the
  penalty code (`addPenalty`, `deletePenalty`, the penalty form, `newRuleset`,
  standings deductions) is intact — un-hide the two elements to restore it.
- **Admin PIN gate (UI-level deterrent).** A PIN stored in Script Properties
  (`ADMIN_PIN`, signed with `ADMIN_TOKEN_SECRET`) gates the admin *screen*: on load
  `getAdminGate(token)` decides whether to show the UI or a PIN overlay;
  `verifyAdminPin` returns a 30-day signed token kept in `localStorage` so a known
  browser isn't re-prompted. Set/clear the PIN from the Settings tab (`setAdminPin`);
  if no PIN is set the gate is open. **This only hides the UI — individual server
  functions are not token-checked**, so it's bypassable via DevTools. The
  server-side enforcement ("solid") is on the wishlist.

## HK mahjong scoring model

The scorer was converted from Riichi to **Hong Kong** rules (branch
`hk-mahjong-scoring`). Key differences from the original engine:

- **Net-points, zero-sum entry.** Each player's entered score IS their net result
  for the game; the four scores must sum to 0 (`actionSubmitScore` / `checkSum`).
  `saveScore` stores the value as-is — no Uma bonus, no starting stack, no
  `x1000`/`÷1000` scaling. The legacy `Scores` "Leftover" column is kept for shape
  and written as `0`. Penalty points are likewise stored/read literally (the old
  thousand-scaling was removed from `addPenalty` and `getStandingsData`).
- **No Uma / no Riichi sticks.** Those settings, inputs, and the placement-bonus
  logic are gone. The tiebreaker setting was reworded (it no longer affects points
  with Uma removed — it only records rank order).

### Faan table

Per-tournament settings (`getFullSettings` / `saveTournamentSettings`):
`Faan_Min` (default 3), `Faan_Max` (default 13, **capped at 13**, and min must be
`< max` — enforced client- and server-side), `Faan_Min_Points` (default 8),
`Faan_Scaling` (`half` | `full` | `custom`), `Self_Pick_Multiplier` (default 1.5),
`False_Win_Points`.

`computeFaanTable(s)` builds rows `{faan, points, selfPick}`. **The progression is
anchored at the minimum faan** (worth `Faan_Min_Points`) and grows from there:
- **half-spicy** (default): faan ≤ 4 pure-double every faan (canonical low region
  1,2,4,8,16); from faan 5 up it doubles every 2 faan with the in-between faan at
  1.5x the previous. The doubling phase anchors at `max(minFaan, 4)`, so if the base
  lands on an odd faan its even neighbours are 1.5x (and vice versa) — clean integers
  from any base, and `min=3` reproduces the canonical table (3:8, 4:16, 5:24, …).
- **full-spicy**: every faan doubles (`minPts * 2^offset`).
- **custom**: literal points list (comma/space separated, Min→Max).
- `selfPick` = `points x Self_Pick_Multiplier` (winner's total on a self-draw,
  split 3 ways).

`getFaanTable()` exposes it. The **player portal** has a third view (Standings →
Pairings → Faan Table) showing Faan/Points/Self-pick. The **admin settings** screen
shows a **live preview** (`renderFaanPreview` mirrors the server computation) that
updates as parameters change.

### False Win

`applyFalseWin(round, table, offenderId)` is a table-wide penalty: the offender is
deducted `3*V` and each of the other three players is credited `+V` (zero-sum),
where `V` = `False_Win_Points` or, if blank, the points at the maximum faan. Written
as `Penalties` rows (positive "Points Deducted" reduces score, negative credits it).
Admin uses the **Apply False Win** button in the Penalties tab (reuses the
round/table/player selectors).

### Theme

`jade` is a player-portal theme (CSS class in `index.html`, option in admin
settings) using base `#1ccc81` with `#cc1c68`/`#1c68cc` accents — joining the
existing `default` and `cherry-blossom` themes.
