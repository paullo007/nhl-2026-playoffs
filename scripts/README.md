# Update Scripts — 2026 NHL Playoffs Workbook

Three Python scripts that built up `2026 NHL Playoffs_v1.8.xlsx → v1.11.xlsx`. They share a reusable pattern: **copy a source workbook, mutate cells via openpyxl, save as a new version.**

## Files

| Script | Input → Output | What it does |
|---|---|---|
| `update_nhl.py` | v1.8 → v1.9 | Fills Result + Goal Scorers (cols I, J) for played games on the "By Dates" sheet; updates series-status text on the "Bracket" sheet. |
| `correct_nhl.py` | v1.9 → v1.10 | Replaces every Goal Scorers row with verified scorers (cross-checked against NHL.com recaps), adds GW/EN/PP/SH tags. |
| `redesign_bracket.py` | v1.10 → v1.11 | Rebuilds the "Bracket" sheet: swaps West/East to match NHL.com (West left, East right), uses compact `COL 3 / LAK 0` series scores, applies fill colors for advanced/leading/eliminated. |

## Run

```bash
python3 update_nhl.py        # produces v1.9
python3 correct_nhl.py       # produces v1.10
python3 redesign_bracket.py  # produces v1.11
```

Requires `openpyxl` (`pip install openpyxl`).

## Reusable patterns

These scripts are templates for any "fetch data → update Excel" workflow (sports brackets, project trackers, dashboards). The reusable shapes:

**1. Copy-then-mutate** — never edit the source file in place:
```python
import shutil, openpyxl
shutil.copyfile(src, dst)
wb = openpyxl.load_workbook(dst)
# ... mutate ...
wb.save(dst)
```

**2. Row-keyed update dict** — keep data separate from logic:
```python
updates = {
    17: ('Result text', 'Scorer text'),
    18: (...),
}
for row, (result, scorers) in updates.items():
    ws.cell(row=row, column=9).value = result
    ws.cell(row=row, column=10).value = scorers
```

**3. Style preservation** — write `value` but reuse `cell.fill`/`cell.font` from the source so layout/colors don't break. To set fills explicitly, use `PatternFill('solid', fgColor='RRGGBB')`.

**4. Coordinate map for layout swaps** — when redesigning a sheet, build a `(coord, value, fill, font)` list and iterate. Easier to review than scattered cell writes.

## Adapting to a new project

1. Replace `src` / `dst` paths.
2. Replace the `updates` dict with your data.
3. Adjust column indices (`column=9, 10` here for Result/Scorers — change for your sheet).
4. For `redesign_bracket.py`-style layouts: inspect cell coordinates in the target template first (`ws.merged_cells`, `ws.column_dimensions`, then read existing styles).

## Data verification habit

`correct_nhl.py` exists because the first pass had inaccuracies. Lesson: when scraping/searching for data, hit the **authoritative source** (NHL.com game recaps, not aggregator articles), and tag the source URL alongside each data point so future-you can re-verify.
