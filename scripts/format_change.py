import openpyxl
import shutil

REPO = '/Users/paullo/01_PLO/02_CLAUDE CODE/02_NHL 2026 Playoffs'
src = f'{REPO}/2026 NHL Playoffs_v1.13.xlsx'
dst = f'{REPO}/2026 NHL Playoffs_v1.14.xlsx'
shutil.copyfile(src, dst)

wb = openpyxl.load_workbook(dst)
b = wb['Bracket']

# New format: "TEAM_LEFT N - M TEAM_RIGHT"
# Position is based on bracket layout (top seed left), not lead.
new_badges = {
    # LEFT (Western)
    'B7':  'COL 4 - 0 LAK',
    'B15': 'DAL 2 - 2 MIN',
    'B23': 'VGK 1 - 2 UTA',
    'B31': 'EDM 1 - 3 ANA',
    # RIGHT (Eastern)
    'P7':  'BUF 3 - 1 BOS',
    'P15': 'TBL 2 - 2 MTL',
    'P23': 'CAR 4 - 0 OTT',
    'P31': 'PIT 1 - 3 PHI',
}

for coord, val in new_badges.items():
    b[coord].value = val

wb.save(dst)
print(f'Saved: {dst}')
print('New badges:')
for coord, val in new_badges.items():
    print(f'  {coord}: {val}')
