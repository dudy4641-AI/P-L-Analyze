# P&L Analyze

Builds a P&L analysis workbook from a monthly vendor-expense file.

## Input
An Excel file where every sheet named `MM.YY` (e.g. `01.26`) is one month and columns A-D are:
`סעיף | ספק | סכום | Entity`. Other sheets are ignored. Amounts are in USD.

## Run
```bash
python build_pnl.py input.xlsx output/PL_Analysis.xlsx
```
Formulas are written without cached values: open the file in Excel (it recalculates on load).

## Output sheets
| Sheet | Contents |
|---|---|
| Dashboard | Filters (dept, budget item, entity, period type + period), KPIs incl. previous-period change, top-10 vendors / items, 4 charts |
| P&L | Department → budget item, by month / quarter / year, entity filter, collapsible groups |
| QoQ | Quarter-over-quarter $ and % change, red = increase, green = decrease |
| Vendors | Vendor mix for a selected dept + item + entity: top 30, share, cumulative share, QoQ % |
| Vendor Detail | Full Dept → Item → Vendor tree by quarter (outline levels 1-3) |
| Data | Unified rows from all month sheets |
| Mapping | Original item → department + budget item (editable) |

## Normalisation
- Department = text before the first `-`; budget item = the rest, trimmed.
- `Rencruitment`→`Recruitment`, `Confrences & Events`/`Conferences`→`Conferences & Events`,
  `Salaries Inc` / `Salaries Ltd` / `Salaries-Ltd` → `Salaries` (entity stays in its own column).
- Vendor names differing only by case/spaces are merged; blank vendor → `(ללא ספק)`.
Edit `ITEM_FIXES` / `DEPT_ORDER` in `build_pnl.py` to change these rules.
