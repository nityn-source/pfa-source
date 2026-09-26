# ASB roster importer (Phase 1, step 1)

Reads the **Roster 2026** Google Sheet and turns it into clean calendar events, then
writes a review report. It **only reads** the sheet. It never changes the sheet,
and it never sends anything anywhere.

What you get (in the folder `roster_import_out/`):

| File | What it is |
|---|---|
| `report.html` | **The review report. Open this one.** Works on a phone. |
| `issues.csv` | Every cell the importer could not resolve, with the sheet cell (e.g. `'Sep'!B21:C21`) |
| `events.csv` | One row per event: date, start, end, type, title, coaches, athletes, sheet cells |
| `names.csv` | Every spelling of every name, and what it was matched to |
| `memberships.csv` | Who is in which PT group / squad, from the side-tables |
| `staff_status.csv` | Per coach per day: hours, or WO / L / SL |
| `roster_import.json` | Everything above in the app's data model |

> These output files contain athletes' names. Don't upload them to GitHub or share them
> outside ASB staff. The `.gitignore` in this folder blocks them from being committed.

---

## Option A: ask Claude to run it (easiest)

In a Claude Code session on this repository, paste:

```
Run the ASB roster importer in asb/roster_importer for 1 Jun to 30 Sep 2026.
Download Roster 2026 (file id 1kKNIbgWmPDmW0E_4gLRW1fEqhMX3VSZIiDKxKeIBo3c) as .xlsx
with the Google Drive connector (read-only), run import_roster.py on it, and send me
report.html. Don't write to the sheet and don't commit the output.
```

## Option B: run it yourself on a Mac

1. Open **Roster 2026** in Google Sheets → **File → Download → Microsoft Excel (.xlsx)**.
   It lands in your Downloads folder as `Roster 2026.xlsx`.
2. Open **Terminal** (press ⌘-Space, type `Terminal`, press Enter).
3. Copy and paste this line, then press Enter. You only need to do it once:
   ```
   python3 -m pip install --user openpyxl
   ```
   (If macOS asks to install "command line developer tools", click **Install**, wait,
   and paste the line again.)
4. Go to the importer folder. Replace the path with where this repository is on your Mac:
   ```
   cd ~/pfa-source/asb/roster_importer
   ```
5. Run the importer:
   ```
   python3 import_roster.py ~/Downloads/"Roster 2026.xlsx" --from 2026-06-01 --to 2026-09-30
   ```
   It prints something like `Imported 1628 events over 119 days` and where the report is.
6. Open the report:
   ```
   open roster_import_out/report.html
   ```

To import a different period, change the two dates in step 5.

---

## Fixing what the report finds

There are two ways to fix an issue:

1. **Fix the sheet** (best for "fix" items). Each issue names the exact cell, e.g.
   `'Sep'!B21:C21`. The most common fix is to put one session per line and the
   matching coach on the same line in the coach column. Then run the importer again.
2. **Confirm a spelling** (for name variants). Copy `aliases.example.csv` to
   `aliases.csv` in this folder. Keep only the lines you agree with, and add your own,
   one per line:
   ```
   raw,canonical,role
   Marioo,Mario,coach
   Asha K,Asha Kumar,athlete
   ```
   Then add `--aliases aliases.csv` to the command in step 5:
   ```
   python3 import_roster.py ~/Downloads/"Roster 2026.xlsx" --aliases aliases.csv
   ```
   `aliases.csv` holds athletes' names, so it is kept out of GitHub by `.gitignore`.

The importer never merges two spellings on its own. It lists likely matches in the
report, and a person confirms them in `aliases.csv`.

## How it reads the sheet

- **Weeks.** It finds every row with dates in columns B, D, F, H, J, L, N. Blocks
  don't have to start on a Monday. If a header has a copy-paste mistake (for example,
  last week's dates), it corrects it from the week before and reports the cell.
- **Duplicates.** If the same day appears in two tabs, it uses the month tab
  (e.g. `June` over `June 26`) and reports the other copy.
- **Cells.** It splits each half-hour cell into sessions. Sessions can be separated
  by new lines, by 3+ spaces or by " / ". It then pairs each session with the coach on
  the same line. When the lines don't pair up, it attaches every coach to every
  session in that cell, marks the event "check coaches", and lists the cell as an issue.
- **Events.** Back-to-back half-hours with the same session and the same coaches
  become one event.
- **Coaches.** Honorifics (sir, di, didi, ma'am) are ignored. The known spellings are
  in `COACHES` at the top of `import_roster.py`. "All" / "All coaches" becomes the
  "All coaches" group.
- **Hours rows.** Each coach's hours or WO / L / SL are read per day. The weekly total
  is checked against the days. The report also flags a coach who is marked off but
  named on the roster that day.
- **Side-tables.** The PT group and squad lists to the right of the grid become the
  membership seed. Each table counts as a snapshot for the week it sits in.
- **Not imported yet:** every tab that isn't a month tab, such as `S&C`, `Physio`,
  `junr elite`, the individual-athlete plans and the scratch tabs. They use different
  layouts, so they are left for a later step. The report lists them.

## For developers

- Python 3.9+ and `openpyxl`. There are no other dependencies.
- Tests use a small made-up workbook: `python3 -m unittest test_import_roster.py`
- The vocabulary lives at the top of `import_roster.py`: `COACHES`,
  `GROUP_PATTERNS` and `ACTIVITIES`. Add new session names or coaches there.
