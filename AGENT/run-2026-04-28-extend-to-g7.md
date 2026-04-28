# Schedule Extension — 2026-04-28

## What changed

**Added 18 rows to the "By Dates" sheet** for Games 5/6/7 of all 6 in-progress series. Skipped CAR/OTT and COL/LAK (already swept).

| Match # | Series | Game | Date Local | Status |
|---|---|---|---|---|
| 33 | PIT/PHI | G5 | Mon Apr 27, 7:00 PM ET | ✅ Filled (PIT 3-2) |
| 34 | BUF/BOS | G5 | Tue Apr 28, 7:30 PM ET | scheduled |
| 35 | DAL/MIN | G5 | Tue Apr 28, 7:00 PM CT | scheduled |
| 36 | EDM/ANA | G5 | Tue Apr 28, 8:00 PM MT | scheduled |
| 37 | PIT/PHI | G6 | Wed Apr 29, 7:30 PM ET | scheduled |
| 38 | TBL/MTL | G5 | Wed Apr 29, 7:00 PM ET | scheduled |
| 39 | VGK/UTA | G5 | Wed Apr 29, 7:00 PM PT | scheduled |
| 40 | EDM/ANA | G6 | Thu Apr 30 TBD | scheduled (if necessary) |
| 41 | DAL/MIN | G6 | Thu Apr 30 TBD | scheduled (if necessary) |
| 42 | BUF/BOS | G6 | Fri May 1 TBD | scheduled (if necessary) |
| 43 | TBL/MTL | G6 | Fri May 1 TBD | scheduled (if necessary) |
| 44 | VGK/UTA | G6 | Fri May 1 TBD | scheduled (if necessary) |
| 45 | PIT/PHI | G7 | Sat May 2 TBD | scheduled (if necessary) |
| 46 | EDM/ANA | G7 | Sat May 2 TBD | scheduled (if necessary) |
| 47 | DAL/MIN | G7 | Sat May 2 TBD | scheduled (if necessary) |
| 48 | BUF/BOS | G7 | Sun May 3 TBD | scheduled (if necessary) |
| 49 | TBL/MTL | G7 | Sun May 3 TBD | scheduled (if necessary) |
| 50 | VGK/UTA | G7 | Sun May 3 TBD | scheduled (if necessary) |

## PIT/PHI G5 result (added today)

**Pittsburgh 3, Philadelphia 2** (Mon Apr 27)
- PIT: E. Soderblom, C. Dewar, K. Letang (GW)
- PHI: A. Bump, T. Sanheim
- Series: Philadelphia leads 3-2

Verified against NHL.com recap.

## Bracket updates

- **P31** badge: `PIT 1 - 3 PHI` → `PIT 2 - 3 PHI`

## Sweep series (no Games 5–7 added)

- **CAR/OTT**: Carolina won 4-0 (G4 was the clincher)
- **COL/LAK**: Colorado won 4-0 (G4 was the clincher)

## Format conventions preserved

All new rows follow the v1.15 conventions exactly:
- Match # column: integer
- Round column: "Round 1"
- Series column: "<Higher Seed Team> vs <Lower Seed Team>"
- Game column: "Game N"
- Matchup column: "<Away Team> vs <Home Team> (at <Home Team>)"
- Date & Time (Local): `<TIME> <TZ>, <DAY>, <MMM><DD>` or `TBD, <DAY>, <MMM><DD>`
- Date & Time (SGT): same pattern, +12h (ET), +13h (CT), +14h (MT), +15h (PT)
- Arena: `<Venue Name>, <City>, <ST>`
- Result + Goal Scorers: empty for unplayed games (agent will fill on next run)

Home rotation pattern (best-of-7): G1/G2 at higher seed, G3/G4 at lower seed, G5 at higher seed, G6 at lower seed, G7 at higher seed.

## Saved as

`2026 NHL Playoffs_v1.16.xlsx`

## Reference script

`scripts/extend_to_g7.py` (saved for reuse)
