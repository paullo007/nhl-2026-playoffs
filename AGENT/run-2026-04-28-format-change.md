# Format Change — 2026-04-28

## What changed

**Bracket series-score badge format updated:**

| Old | New |
|---|---|
| `COL 4 / LAK 0` | `COL 4 - 0 LAK` |
| `DAL 2 / MIN 2` | `DAL 2 - 2 MIN` |
| `VGK 1 / UTA 2` | `VGK 1 - 2 UTA` |
| `EDM 1 / ANA 2` | `EDM 1 - 3 ANA` |
| `BUF 2 / BOS 1` | `BUF 3 - 1 BOS` |
| `TBL 1 / MTL 2` | `TBL 2 - 2 MTL` |
| `CAR 4 / OTT 0` | `CAR 4 - 0 OTT` |
| `PIT 1 / PHI 3` | `PIT 1 - 3 PHI` |

## Format rule (now standard)

`<TEAM_LEFT> <SCORE_LEFT> - <SCORE_RIGHT> <TEAM_RIGHT>`

- Team positions follow the bracket's vertical seed order (top seed left), NOT the current lead.
- Single hyphen with single spaces around it: `N - M`.
- Three-letter team abbreviations.

## Convention going forward

**Each iteration uses the previous version's format as the authoritative standard.** When the agent processes v1.X → v1.X+1, it inspects the existing bracket badge format in v1.X and reproduces it exactly — no silent format drift.

## File version

- Previous: `v1.13.xlsx` (used `/` separator)
- Current: `v1.14.xlsx` (uses `-` separator)
- Local copy: `~/Documents/2026 NHL Playoffs_LATEST.xlsx`

## Code changes

- Cells B7, B15, B23, B31 (West) and P7, P15, P23, P31 (East) updated.
- Cloud agent prompt updated (step 8) to specify new format and the "carry-forward convention".
- Format-change script saved to `scripts/format_change.py` for reference.
