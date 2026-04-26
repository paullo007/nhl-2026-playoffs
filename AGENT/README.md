# Scheduler Agent — 2026 NHL Playoffs

This folder is reserved for the **scheduled remote agent** that refreshes the workbook with new game results on a recurring basis.

## What goes here (after `/schedule` is run)

When you run `/schedule` in Claude Code to create the routine, the harness will store:
- **Routine config** — the cron schedule, prompt, model, and isolation settings.
- **Run history** — logs of each scheduled execution (timestamps, what changed).
- **Generated artifacts** — any side files the agent produces (e.g., a daily diff log, a "what changed" markdown summary).

You can manually drop additional files here too:
- `prompt.md` — the prompt template the agent runs (so you can tweak it without re-running `/schedule`).
- `notes.md` — your own observations about what the agent does well / poorly.

## Reusing this agent for another project

The scheduler-agent pattern (sports brackets, project trackers, KPI dashboards — anything where you want a daily/weekly refresh) generalizes to:

1. **Trigger:** cron schedule via `/schedule`, or manual via `/loop`.
2. **Inputs:** an existing workbook + a list of authoritative URLs to pull from.
3. **Action:** fetch data → diff against current state → update only the changed cells → save as a new versioned file.
4. **Output:** new `vX.Y.xlsx` + a short "what changed" note.

The reusable scripts are in `../scripts/` — start by copying those into a new project folder and adapting the `updates` dict and source URLs.

## Linked resources

- Workbook (current): `../2026 NHL Playoffs_v1.11.xlsx`
- Past update scripts: `../scripts/`
- Authoritative data source: https://www.nhl.com/playoffs/2026/bracket
