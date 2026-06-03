# hk-mahjong-tournament

Google Apps Script web app for running Hong Kong mahjong tournaments, managed
locally with `clasp`.

See [`.claude/architecture.md`](.claude/architecture.md) for the data model: the
master/tournament-sheet split, how the active tournament is resolved and switched,
the deploy workflow, and schema/player-ID rules. Read it before changing data flow
in `Code.gs`.
