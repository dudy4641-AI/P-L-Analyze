#!/usr/bin/env python3
"""
build_pnl.py - בניית קובץ ניתוח P&L מתוך קובץ הוצאות לפי ספקים.

קלט : קובץ Excel שבו כל לשונית בשם MM.YY (למשל 01.26) היא חודש הוצאה,
       ובטורים A-D: סעיף | ספק | סכום | Entity
פלט : קובץ Excel עם הלשוניות:
       Dashboard      - דשבורד עם פילטרים (מחלקה, סעיף, יישות, תקופה) וגרפים
       P&L            - דוח לפי מחלקה -> סעיף תקציבי, חודשים / רבעונים / שנים
       QoQ            - שינויים רבעון מול רבעון ($ ו-%)
       Vendors        - הרכב ספקים דינמי לסעיף נבחר
       Vendor Detail  - מחלקה -> סעיף -> ספק (מקובץ, ניתן לכיווץ)
       Data           - טבלת נתונים מאוחדת
       Mapping        - מיפוי סעיף מקורי -> מחלקה + סעיף תקציבי
       Lists, Calc    - רשימות ועזרי חישוב (לא לעריכה)

שימוש:
    python build_pnl.py input.xlsx [output.xlsx]
"""
import re
import sys
from collections import Counter, defaultdict, OrderedDict

import openpyxl
from openpyxl.chart import BarChart, DoughnutChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

# ----------------------------------------------------------------- settings
ALL = "הכל"
NO_VENDOR = "(ללא ספק)"
DEPT_ORDER = ["Cogs", "R&D", "Sales", "Marketing", "G&A"]
# fixes applied to the budget-item part (after the department prefix), lower-case keys
ITEM_FIXES = {
    "rencruitment": "Recruitment",
    "confrences & events": "Conferences & Events",
    "conferences": "Conferences & Events",
}
DEPT_COLORS = ["2A78D6", "EB6834", "1BAF7A", "EDA100", "E87BA4", "008300", "7F7F7F"]
SINGLE = "2A78D6"
CUR = '$#,##0;($#,##0);"-"'
PCT = '0.0%;(0.0%);"-"'
EXTRA_ROWS = 5000  # room to append rows in Data without breaking formulas

FONT = "Arial"
F_BASE = Font(name=FONT, size=10)
F_BOLD = Font(name=FONT, size=10, bold=True)
F_HDR = Font(name=FONT, size=10, bold=True, color="FFFFFF")
F_TITLE = Font(name=FONT, size=16, bold=True, color="1F3864")
F_SUB = Font(name=FONT, size=9, italic=True, color="666666")
F_GREY = Font(name=FONT, size=8, color="A0A0A0")
FILL_HDR = PatternFill("solid", fgColor="1F3864")
FILL_DEPT = PatternFill("solid", fgColor="D9E2F3")
FILL_TOT = PatternFill("solid", fgColor="B4C6E7")
FILL_INPUT = PatternFill("solid", fgColor="FFFF00")
FILL_KPI = PatternFill("solid", fgColor="F2F2F2")
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)


# ----------------------------------------------------------------- helpers
def clean(s):
    return re.sub(r"\s+", " ", str(s)).strip() if s is not None else ""


def split_item(raw):
    """'G&A-Salaries Ltd' -> ('G&A', 'Salaries')"""
    s = clean(raw)
    if "-" in s:
        dept, item = s.split("-", 1)
    else:
        dept, item = s, s
    dept, item = dept.strip(), item.strip(" -")
    for d in DEPT_ORDER:  # normalise dept casing
        if d.lower() == dept.lower():
            dept = d
    low = item.lower()
    if re.match(r"^salaries[\s-]*(inc|ltd)?$", low):
        item = "Salaries"
    item = ITEM_FIXES.get(low, item)
    return dept, item


def month_key(sheet_name):
    mm, yy = sheet_name.split(".")
    return f"20{yy}-{mm}", f"20{yy}-Q{(int(mm) - 1) // 3 + 1}", f"20{yy}"


def style_header(ws, row, c1, c2, height=30):
    for c in range(c1, c2 + 1):
        cell = ws.cell(row, c)
        cell.font, cell.fill, cell.alignment, cell.border = F_HDR, FILL_HDR, CENTER, BOX
    ws.row_dimensions[row].height = height


def set_font_all(ws):
    for row in ws.iter_rows():
        for c in row:
            if c.value is not None and (c.font is None or c.font.name != FONT):
                c.font = Font(name=FONT, size=c.font.size or 10, bold=c.font.bold,
                              italic=c.font.italic, color=c.font.color)


def add_name(wb, name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)


def add_list_validation(ws, cell, formula, prompt):
    dv = DataValidation(type="list", formula1=formula, allow_blank=False,
                        showDropDown=False)
    dv.promptTitle, dv.prompt = "בחירה", prompt
    dv.showInputMessage = True
    ws.add_data_validation(dv)
    dv.add(cell)


# ----------------------------------------------------------------- read
def read_source(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    rows, skipped = [], []
    for ws in wb.worksheets:
        if not re.fullmatch(r"\d\d\.\d\d", ws.title.strip()):
            continue
        month, q, y = month_key(ws.title.strip())
        for i, r in enumerate(ws.iter_rows(min_row=2, max_col=4, values_only=True), 2):
            item, vendor, amt, ent = r
            if item is None and vendor is None and amt is None:
                continue
            if not isinstance(amt, (int, float)):
                skipped.append((ws.title, i, r))
                continue
            rows.append(dict(sheet=ws.title.strip(), month=month, q=q, y=y,
                             item_raw=item if item is not None else "",
                             vendor_raw=vendor, amt=float(amt), ent=clean(ent)))
    if not rows:
        raise ValueError("לא נמצאו לשוניות חודשיות בשם MM.YY (למשל 01.26) עם נתונים בטורים A-D")
    rows.sort(key=lambda d: d["month"])
    return rows, skipped


def vendor_names(rows):
    """unify vendor spellings (case / spaces) -> display name + numeric id"""
    spell = defaultdict(Counter)
    for r in rows:
        v = clean(r["vendor_raw"]) or NO_VENDOR
        spell[v.lower()][v] += 1
    disp = {}
    for k, cnt in spell.items():
        disp[k] = sorted(cnt.items(), key=lambda kv: (-kv[1], kv[0].isupper(), kv[0]))[0][0]
    names = sorted(set(disp.values()), key=str.lower)
    vid = {n: i + 1 for i, n in enumerate(names)}
    for r in rows:
        r["vendor"] = disp[(clean(r["vendor_raw"]) or NO_VENDOR).lower()]
        r["vid"] = vid[r["vendor"]]
    return names, vid


def load_dataframe(src):
    """the same normalised rows the workbook uses, as a pandas DataFrame"""
    import pandas as pd
    rows, skipped = read_source(src)
    vendor_names(rows)
    cache = {}
    for r in rows:
        if r["item_raw"] not in cache:
            cache[r["item_raw"]] = split_item(r["item_raw"])
        r["dept"], r["item"] = cache[r["item_raw"]]
    df = pd.DataFrame(rows).rename(columns={"q": "quarter", "y": "year", "ent": "entity",
                                            "amt": "amount"})
    return df[["sheet", "month", "quarter", "year", "dept", "item", "vendor", "entity",
               "amount", "item_raw", "vendor_raw"]], skipped


# ----------------------------------------------------------------- build
def build(src, out):
    """src / out: a path or a file-like object (e.g. io.BytesIO)"""
    rows, skipped = read_source(src)
    vnames, vid = vendor_names(rows)
    mapping = OrderedDict()
    for r in rows:
        if r["item_raw"] not in mapping:
            mapping[r["item_raw"]] = split_item(r["item_raw"])
        r["dept"], r["item"] = mapping[r["item_raw"]]

    depts = [d for d in DEPT_ORDER if any(m[0] == d for m in mapping.values())]
    depts += sorted({m[0] for m in mapping.values()} - set(depts))
    dept_items = OrderedDict((d, sorted({m[1] for m in mapping.values() if m[0] == d}))
                             for d in depts)
    items = sorted({m[1] for m in mapping.values()})
    months = sorted({r["month"] for r in rows})
    quarters = sorted({r["q"] for r in rows})
    years = sorted({r["y"] for r in rows})
    ents = sorted({r["ent"] for r in rows})
    pairs = [(d, i) for d in depts for i in dept_items[d]]

    wb = openpyxl.Workbook()
    wsD = wb.active
    wsD.title = "Dashboard"
    wsP = wb.create_sheet("P&L")
    wsQ = wb.create_sheet("QoQ")
    wsV = wb.create_sheet("Vendors")
    wsVD = wb.create_sheet("Vendor Detail")
    wsData = wb.create_sheet("Data")
    wsM = wb.create_sheet("Mapping")
    wsL = wb.create_sheet("Lists")
    wsC = wb.create_sheet("Calc")

    # ---------------------------------------------------------- Mapping
    wsM.append(["סעיף מקורי (מהקובץ)", "מחלקה", "סעיף תקציבי"])
    style_header(wsM, 1, 1, 3)
    for raw, (d, i) in mapping.items():
        wsM.append([raw, d, i])
        r = wsM.max_row
        for c in (2, 3):
            wsM.cell(r, c).fill = FILL_INPUT
    nmap = wsM.max_row
    wsM["E1"] = "הסבר"
    wsM["E1"].font = F_BOLD
    wsM["E2"] = ("המחלקה והסעיף בלשונית Data נשלפים מכאן. ניתן לשנות את התאים הצהובים. "
                 "סעיף חדש שלא קיים בדוח יופיע רק לאחר הרצה מחדש של build_pnl.py.")
    wsM["E3"] = ("נרמול שבוצע: הסרת רווחים מיותרים, Rencruitment->Recruitment, "
                 "Confrences/Conferences->Conferences & Events, Salaries Inc/Ltd->Salaries "
                 "(היישות נשמרת בעמודה נפרדת).")
    for c in ("E2", "E3"):
        wsM[c].alignment = Alignment(wrap_text=True, vertical="top")
    wsM.column_dimensions["A"].width = 36
    wsM.column_dimensions["B"].width = 14
    wsM.column_dimensions["C"].width = 26
    wsM.column_dimensions["E"].width = 70
    wsM.freeze_panes = "A2"

    # ---------------------------------------------------------- Data
    hdr = ["לשונית מקור", "חודש", "רבעון", "שנה", "מחלקה", "סעיף תקציבי", "ספק",
           "Vendor ID", "יישות", "סכום ($)", "סעיף מקורי", "ספק מקורי"]
    wsData.append(hdr)
    style_header(wsData, 1, 1, len(hdr))
    for n, r in enumerate(rows, 2):
        mref = f"Mapping!$A$2:$A${nmap}"
        wsData.append([
            r["sheet"], r["month"], r["q"], r["y"],
            f'=IFERROR(INDEX(Mapping!$B$2:$B${nmap},MATCH(K{n},{mref},0)),"לא ממופה")',
            f'=IFERROR(INDEX(Mapping!$C$2:$C${nmap},MATCH(K{n},{mref},0)),"לא ממופה")',
            r["vendor"], r["vid"], r["ent"], r["amt"], r["item_raw"],
            clean(r["vendor_raw"]) if r["vendor_raw"] is not None else "",
        ])
        wsData.cell(n, 10).number_format = '#,##0.00'
    last = len(rows) + 1 + EXTRA_ROWS
    for col, w in zip("ABCDEFGHIJKL", [10, 9, 9, 7, 11, 24, 32, 9, 7, 12, 30, 32]):
        wsData.column_dimensions[col].width = w
    wsData.freeze_panes = "A2"
    wsData.auto_filter.ref = f"A1:L{len(rows) + 1}"
    for name, col in [("D_Month", "B"), ("D_Q", "C"), ("D_Year", "D"), ("D_Dept", "E"),
                      ("D_Item", "F"), ("D_Vendor", "G"), ("D_VID", "H"),
                      ("D_Ent", "I"), ("D_Amt", "J")]:
        add_name(wb, name, f"Data!${col}$2:${col}${last}")
    add_name(wb, "D_Per", f"Data!$B$2:$D${last}")

    # ---------------------------------------------------------- Lists
    lists = OrderedDict([
        ("L_Dept", [ALL] + depts), ("L_Item", [ALL] + items), ("L_Ent", [ALL] + ents),
        ("L_Type", [ALL, "חודש", "רבעון", "שנה"]), ("L_All", [ALL]),
        ("L_Month", months), ("L_Quarter", quarters), ("L_Year", years),
        ("L_TypeName", ["L_All", "L_Month", "L_Quarter", "L_Year"]),
    ])
    for c, (name, vals) in enumerate(lists.items(), 1):
        wsL.cell(1, c, name).font = F_BOLD
        for r, v in enumerate(vals, 2):
            wsL.cell(r, c, v)
        add_name(wb, name, f"Lists!${L(c)}$2:${L(c)}${len(vals) + 1}")
        wsL.column_dimensions[L(c)].width = 22

    # ---------------------------------------------------------- Calc (filters)
    # Dashboard selectors live at Dashboard!C4:C8
    wsC["A1"] = "עזרי חישוב - לא לעריכה"
    wsC["A1"].font = F_BOLD
    calc_cells = [
        ("F_Dept", "קריטריון מחלקה", f'=IF(Dashboard!$C$4="{ALL}","*",Dashboard!$C$4)'),
        ("F_Item", "קריטריון סעיף", f'=IF(Dashboard!$C$5="{ALL}","*",Dashboard!$C$5)'),
        ("F_Ent", "קריטריון יישות", f'=IF(Dashboard!$C$6="{ALL}","*",Dashboard!$C$6)'),
        ("F_TypeN", "סוג תקופה (1-4)", '=IFERROR(MATCH(Dashboard!$C$7,L_Type,0),1)'),
        ("F_PCol", "עמודת תקופה ב-Data", '=MAX(1,F_TypeN-1)'),
        ("F_Per", "קריטריון תקופה",
         f'=IF(OR(F_TypeN=1,Dashboard!$C$8="{ALL}",Dashboard!$C$8=""),"*",Dashboard!$C$8)'),
        ("F_ListName", "שם רשימת תקופות", '=INDEX(L_TypeName,F_TypeN)'),
        ("F_Pos", "מיקום התקופה ברשימה",
         '=IF(F_TypeN=1,0,IFERROR(MATCH(Dashboard!$C$8,CHOOSE(F_PCol,L_Month,L_Quarter,L_Year),0),0))'),
        ("F_Prev", "תקופה קודמת",
         '=IF(F_Pos>1,INDEX(CHOOSE(F_PCol,L_Month,L_Quarter,L_Year),F_Pos-1),"")'),
    ]
    for r, (name, label, f) in enumerate(calc_cells, 3):
        wsC.cell(r, 1, label)
        wsC.cell(r, 2, f)
        add_name(wb, name, f"Calc!$B${r}")
    wsC.column_dimensions["A"].width = 24
    wsC.column_dimensions["B"].width = 16

    base = "D_Amt,D_Dept,F_Dept,D_Item,F_Item,D_Ent,F_Ent"
    per = "INDEX(D_Per,0,F_PCol),F_Per"

    # monthly trend by dept (filters: dept/item/entity)  rows 15..
    r0 = 15
    wsC.cell(r0 - 1, 1, "מגמה חודשית לפי מחלקה").font = F_BOLD
    wsC.cell(r0, 1, "חודש")
    for j, d in enumerate(depts, 2):
        wsC.cell(r0, j, d)
    for i, m in enumerate(months, r0 + 1):
        wsC.cell(i, 1, m)
        for j, d in enumerate(depts, 2):
            wsC.cell(i, j, f'=IF(OR(F_Dept="*",F_Dept={L(j)}${r0}),'
                           f'SUMIFS(D_Amt,D_Dept,{L(j)}${r0},D_Item,F_Item,D_Ent,F_Ent,D_Month,$A{i}),0)')
            wsC.cell(i, j).number_format = CUR
    trend = (r0, r0 + len(months))

    # quarterly by dept
    r1 = trend[1] + 3
    wsC.cell(r1 - 1, 1, "רבעונים לפי מחלקה").font = F_BOLD
    wsC.cell(r1, 1, "רבעון")
    for j, d in enumerate(depts, 2):
        wsC.cell(r1, j, d)
    for i, q in enumerate(quarters, r1 + 1):
        wsC.cell(i, 1, q)
        for j, d in enumerate(depts, 2):
            wsC.cell(i, j, f'=IF(OR(F_Dept="*",F_Dept={L(j)}${r1}),'
                           f'SUMIFS(D_Amt,D_Dept,{L(j)}${r1},D_Item,F_Item,D_Ent,F_Ent,D_Q,$A{i}),0)')
            wsC.cell(i, j).number_format = CUR
    qtab = (r1, r1 + len(quarters))

    # dept share for selected period
    r2 = qtab[1] + 3
    wsC.cell(r2 - 1, 1, "התפלגות מחלקות בתקופה הנבחרת").font = F_BOLD
    wsC.cell(r2, 1, "מחלקה")
    wsC.cell(r2, 2, "סכום")
    for i, d in enumerate(depts, r2 + 1):
        wsC.cell(i, 1, d)
        wsC.cell(i, 2, f'=IF(OR(F_Dept="*",F_Dept=$A{i}),SUMIFS(D_Amt,D_Dept,$A{i},'
                       f'D_Item,F_Item,D_Ent,F_Ent,{per}),0)')
        wsC.cell(i, 2).number_format = CUR
    share = (r2, r2 + len(depts))

    # dept/item pairs for the selected period -> top items
    r3 = share[1] + 3
    wsC.cell(r3 - 1, 1, "סעיפים (מחלקה - סעיף) בתקופה הנבחרת").font = F_BOLD
    for j, h in enumerate(["מחלקה", "סעיף", "תווית", "סכום", "מפתח דירוג"], 1):
        wsC.cell(r3, j, h)
    for i, (d, it) in enumerate(pairs, r3 + 1):
        wsC.cell(i, 1, d)
        wsC.cell(i, 2, it)
        wsC.cell(i, 3, f'=A{i}&" - "&B{i}')
        wsC.cell(i, 4, f'=IF(AND(OR(F_Dept="*",F_Dept=A{i}),OR(F_Item="*",F_Item=B{i})),'
                       f'SUMIFS(D_Amt,D_Dept,A{i},D_Item,B{i},D_Ent,F_Ent,{per}),0)')
        wsC.cell(i, 5, f'=IF(D{i}>0,D{i}+ROW()/10000000,-ROW())')
    itab = (r3 + 1, r3 + len(pairs))

    # vendor block (columns H..) : dashboard filters + Vendors-sheet filters
    vc = 8  # column H
    heads = ["Vendor ID", "ספק", "דשבורד: סכום", "דשבורד: מפתח",
             "Vendors: סה\"כ", "Vendors: מפתח"] + [f"Vendors: {q}" for q in quarters]
    wsC.cell(1, vc, "חישובי ספקים").font = F_BOLD
    for j, h in enumerate(heads):
        wsC.cell(2, vc + j, h)
    vbase = "D_Amt,D_VID,{id},D_Dept,V_Dept,D_Item,V_Item,D_Ent,V_Ent"
    for i, name in enumerate(vnames, 3):
        idc = f"${L(vc)}{i}"
        wsC.cell(i, vc, vid[name])
        wsC.cell(i, vc + 1, name)
        wsC.cell(i, vc + 2, f'=SUMIFS({base},D_VID,{idc},{per})')
        wsC.cell(i, vc + 3, f'=IF({L(vc + 2)}{i}>0,{L(vc + 2)}{i}+ROW()/10000000,-ROW())')
        wsC.cell(i, vc + 4, f'=SUMIFS({vbase.format(id=idc)})')
        wsC.cell(i, vc + 5, f'=IF({L(vc + 4)}{i}>0,{L(vc + 4)}{i}+ROW()/10000000,-ROW())')
        for k, q in enumerate(quarters):
            wsC.cell(i, vc + 6 + k, f'=SUMIFS({vbase.format(id=idc)},D_Q,"{q}")')
    vlast = 2 + len(vnames)
    add_name(wb, "C_VName", f"Calc!${L(vc + 1)}$3:${L(vc + 1)}${vlast}")
    add_name(wb, "C_DashAmt", f"Calc!${L(vc + 2)}$3:${L(vc + 2)}${vlast}")
    add_name(wb, "C_DashKey", f"Calc!${L(vc + 3)}$3:${L(vc + 3)}${vlast}")
    add_name(wb, "C_VTot", f"Calc!${L(vc + 4)}$3:${L(vc + 4)}${vlast}")
    add_name(wb, "C_VKey", f"Calc!${L(vc + 5)}$3:${L(vc + 5)}${vlast}")
    add_name(wb, "C_VQ", f"Calc!${L(vc + 6)}$3:${L(vc + 5 + len(quarters))}${vlast}")
    add_name(wb, "C_ILabel", f"Calc!$C${itab[0]}:$C${itab[1]}")
    add_name(wb, "C_IAmt", f"Calc!$D${itab[0]}:$D${itab[1]}")
    add_name(wb, "C_IKey", f"Calc!$E${itab[0]}:$E${itab[1]}")

    # ---------------------------------------------------------- Dashboard
    ws = wsD
    ws.sheet_view.rightToLeft = True
    ws.sheet_view.showGridLines = False
    ws["B1"] = "דשבורד הוצאות P&L"
    ws["B1"].font = F_TITLE
    ws["B2"] = (f"נתונים: {months[0]} עד {months[-1]} | כל הסכומים בדולרים ($) | "
                "התאים הצהובים הם פילטרים לבחירה")
    ws["B2"].font = F_SUB
    filt = [("מחלקה", "=L_Dept", "בחר מחלקה או הכל"),
            ("סעיף תקציבי", "=L_Item", "בחר סעיף תקציבי או הכל"),
            ("יישות", "=L_Ent", "בחר יישות או הכל"),
            ("סוג תקופה", "=L_Type", "הכל / חודש / רבעון / שנה"),
            ("תקופה", "=INDIRECT(Calc!$B$9)", "בחר ערך לפי סוג התקופה")]
    for r, (lab, f, prompt) in enumerate(filt, 4):
        ws.cell(r, 2, lab).font = F_BOLD
        c = ws.cell(r, 3, ALL)
        c.fill, c.border, c.font = FILL_INPUT, BOX, F_BOLD
        c.alignment = Alignment(horizontal="center")
        add_list_validation(ws, f"C{r}", f, prompt)
    ws["C7"] = "רבעון"
    ws["C8"] = quarters[-1]
    ws["B10"] = (f'=IF(AND(F_TypeN>1,Dashboard!$C$8<>"{ALL}",F_Pos=0),'
                 '"⚠ הערך בשדה התקופה אינו מתאים לסוג התקופה - בחר מחדש","")')
    ws["B10"].font = Font(name=FONT, size=9, bold=True, color="C00000")
    ws["B9"] = "* פילטר התקופה משפיע על המדדים, ההתפלגות והדירוגים. גרפי המגמה מציגים את כל התקופות."
    ws["B9"].font = F_SUB

    # KPI tiles to the right of the filters. total = H5, previous total = K5
    kpis = [
        ("H4", "H5", "סה\"כ הוצאות בתקופה", f"=SUMIFS({base},{per})", CUR),
        ("I4:J4", "I5:J5", "תקופה קודמת", '=IF(F_Prev="","-",F_Prev)', "@"),
        ("K4", "K5", "סה\"כ תקופה קודמת",
         f'=IF(F_Prev="","-",SUMIFS({base},INDEX(D_Per,0,F_PCol),F_Prev))', CUR),
        ("L4", "L5", "שינוי ($)", '=IF(F_Prev="","-",H5-K5)', CUR),
        ("M4", "M5", "שינוי (%)", '=IF(OR(F_Prev="",N(K5)=0),"-",H5/K5-1)', PCT),
        ("H7", "H8", "ספקים עם הוצאה", '=COUNTIF(C_DashAmt,">0")', "#,##0"),
        ("I7:J7", "I8:J8", "מספר תנועות",
         f'=COUNTIFS(D_Dept,F_Dept,D_Item,F_Item,D_Ent,F_Ent,{per})', "#,##0"),
    ]
    for lab_rng, val_rng, lab, f, fmt in kpis:
        c1 = ws[lab_rng.split(":")[0]]
        c1.value, c1.font, c1.fill, c1.alignment = lab, Font(name=FONT, size=9, color="595959"), FILL_KPI, CENTER
        c2 = ws[val_rng.split(":")[0]]
        c2.value, c2.font, c2.fill, c2.alignment = f, Font(name=FONT, size=14, bold=True, color="1F3864"), FILL_KPI, CENTER
        c2.number_format = fmt
        for rng in (lab_rng, val_rng):
            if ":" in rng:
                ws.merge_cells(rng)
    for rr in (5, 8):
        ws.row_dimensions[rr].height = 28
    # conditional colour on change (expense up = red)
    for rng in ("L5", "M5"):
        ws.conditional_formatting.add(rng, CellIsRule(operator="greaterThan", formula=["0"],
                                      font=Font(name=FONT, size=14, bold=True, color="C00000")))
        ws.conditional_formatting.add(rng, CellIsRule(operator="lessThan", formula=["0"],
                                      font=Font(name=FONT, size=14, bold=True, color="008000")))

    # top 10 vendors table (rows 15..)  columns P..R  and top 10 items  T..V
    def top_table(col0, title, keyname, labname, amtname, total_ref, n=10, row0=12):
        ws.cell(row0, col0, title).font = F_BOLD
        for j, h in enumerate(["#", "שם", "סכום", "% מהסה\"כ"]):
            ws.cell(row0 + 1, col0 + j, h)
        style_header(ws, row0 + 1, col0, col0 + 3, 20)
        for k in range(1, n + 1):
            r = row0 + 1 + k
            lc, ac = L(col0 + 1), L(col0 + 2)
            ws.cell(r, col0, k)
            ws.cell(r, col0 + 1, f'=IF(LARGE({keyname},{k})>0,INDEX({labname},MATCH(LARGE({keyname},{k}),{keyname},0)),"")')
            ws.cell(r, col0 + 2, f'=IF({lc}{r}="","",INDEX({amtname},MATCH(LARGE({keyname},{k}),{keyname},0)))')
            ws.cell(r, col0 + 3, f'=IF(OR({lc}{r}="",N({total_ref})=0),"",{ac}{r}/{total_ref})')
            ws.cell(r, col0 + 2).number_format = CUR
            ws.cell(r, col0 + 3).number_format = PCT
            for j in range(4):
                ws.cell(r, col0 + j).border = BOX
                ws.cell(r, col0 + j).font = F_BASE
        return row0 + 2, row0 + 1 + n

    tv = top_table(2, "10 הספקים המובילים (לפי הפילטרים)", "C_DashKey", "C_VName", "C_DashAmt", "$H$5")
    ti = top_table(7, "10 הסעיפים המובילים (מחלקה - סעיף)", "C_IKey", "C_ILabel", "C_IAmt", "$H$5")
    for col, w in zip("ABCDEFGHIJKLMN", [2, 14, 26, 14, 10, 3, 5, 30, 14, 10, 14, 14, 14, 14]):
        ws.column_dimensions[col].width = w

    # charts
    def bar(title, ytitle=None, horizontal=False, stacked=False, w=17, h=8.5):
        ch = BarChart()
        ch.type = "bar" if horizontal else "col"
        ch.title = title
        ch.style = 10
        ch.width, ch.height = w, h
        ch.y_axis.numFmt = '$#,##0'
        ch.y_axis.majorGridlines.spPr = None
        ch.gapWidth = 60
        if stacked:
            ch.grouping, ch.overlap = "stacked", 100
        ch.legend.position = "b"
        ch.y_axis.delete = False
        ch.x_axis.delete = False
        return ch

    def color_series(ch, colors):
        for s, col in zip(ch.series, colors):
            s.graphicalProperties.solidFill = col
            s.graphicalProperties.line.solidFill = "FFFFFF"

    ch = bar("מגמה חודשית לפי מחלקה ($)", stacked=True)
    ch.add_data(Reference(wsC, min_col=2, max_col=1 + len(depts), min_row=trend[0], max_row=trend[1]), titles_from_data=True)
    ch.set_categories(Reference(wsC, min_col=1, min_row=trend[0] + 1, max_row=trend[1]))
    color_series(ch, DEPT_COLORS)
    ch.width = 13.5
    ws.add_chart(ch, "B26")

    ch = bar("השוואה רבעונית לפי מחלקה ($)")
    ch.add_data(Reference(wsC, min_col=2, max_col=1 + len(depts), min_row=qtab[0], max_row=qtab[1]), titles_from_data=True)
    ch.set_categories(Reference(wsC, min_col=1, min_row=qtab[0] + 1, max_row=qtab[1]))
    color_series(ch, DEPT_COLORS)
    ch.width = 19
    ws.add_chart(ch, "H26")

    dn = DoughnutChart()
    dn.title = "התפלגות לפי מחלקה - תקופה נבחרת"
    dn.width, dn.height = 13.5, 8.5
    dn.add_data(Reference(wsC, min_col=2, min_row=share[0], max_row=share[1]), titles_from_data=True)
    dn.set_categories(Reference(wsC, min_col=1, min_row=share[0] + 1, max_row=share[1]))
    dn.dataLabels = DataLabelList()
    dn.dataLabels.showPercent = True
    dn.dataLabels.showVal = dn.dataLabels.showCatName = dn.dataLabels.showSerName = False
    dn.dataLabels.showLeaderLines = False
    from openpyxl.chart.series import DataPoint
    for k, col in enumerate(DEPT_COLORS[:len(depts)]):
        pt = DataPoint(idx=k)
        pt.graphicalProperties.solidFill = col
        pt.graphicalProperties.line.solidFill = "FFFFFF"
        dn.series[0].dPt.append(pt)
    dn.legend.position = "r"
    ws.add_chart(dn, "B44")

    ch = bar("10 הספקים המובילים - תקופה נבחרת ($)", horizontal=True)
    ch.add_data(Reference(ws, min_col=4, min_row=tv[0], max_row=tv[1]), titles_from_data=False)
    ch.set_categories(Reference(ws, min_col=3, min_row=tv[0], max_row=tv[1]))
    ch.x_axis.scaling.orientation = "maxMin"
    ch.legend = None
    color_series(ch, [SINGLE])
    ch.width = 19
    ws.add_chart(ch, "H44")
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    # ---------------------------------------------------------- P&L
    def pnl_sheet(ws, title, cols, entity_cell="$C$3"):
        """cols: list of (header, kind, value) kind in month/q/year/total"""
        ws.sheet_view.rightToLeft = True
        ws.sheet_view.showGridLines = False
        ws["A1"] = title
        ws["A1"].font = F_TITLE
        ws["A3"] = "יישות:"
        ws["A3"].font = F_BOLD
        ws["C3"] = ALL
        ws["C3"].fill, ws["C3"].border, ws["C3"].font = FILL_INPUT, BOX, F_BOLD
        add_list_validation(ws, "C3", "=L_Ent", "בחר יישות או הכל")
        ws["D3"] = "כל הסכומים בדולרים ($). ניתן לכווץ/להרחיב מחלקות עם כפתורי ה- +/- בצד."
        ws["D3"].font = F_SUB
        ent = f'IF({entity_cell}="{ALL}","*",{entity_cell})'
        hr = 5
        ws.cell(hr, 1, "מחלקה")
        ws.cell(hr, 2, "סעיף תקציבי")
        for j, (h, _, _) in enumerate(cols, 3):
            ws.cell(hr, j, h)
        style_header(ws, hr, 1, 2 + len(cols))
        r = hr + 1
        dept_rows = []
        for d in depts:
            dr = r
            dept_rows.append(dr)
            ws.cell(dr, 1, d).font = F_BOLD
            ws.cell(dr, 2, f"סה\"כ {d}").font = F_BOLD
            its = dept_items[d]
            for k, it in enumerate(its, 1):
                ir = dr + k
                ws.cell(ir, 1, d).font = F_GREY
                ws.cell(ir, 2, it).font = F_BASE
                ws.cell(ir, 2).alignment = Alignment(indent=1)
                ws.row_dimensions[ir].outlineLevel = 1
                for j, (h, kind, val) in enumerate(cols, 3):
                    ws.cell(ir, j, cell_formula(kind, val, f"$A{ir}", f"$B{ir}", ent, ws, ir, j, cols))
                    ws.cell(ir, j).font = F_BASE
            for j, (h, kind, val) in enumerate(cols, 3):
                c = ws.cell(dr, j)
                if kind in ("dpct", "dabs"):
                    c.value = cell_formula(kind, val, None, None, ent, ws, dr, j, cols)
                elif kind == "share":
                    c.value = cell_formula(kind, val, None, None, ent, ws, dr, j, cols)
                else:
                    c.value = f"=SUM({L(j)}{dr + 1}:{L(j)}{dr + len(its)})"
                c.font = F_BOLD
            for j in range(1, 3 + len(cols)):
                ws.cell(dr, j).fill = FILL_DEPT
                ws.cell(dr, j).border = BOX
            r = dr + len(its) + 1
        tr = r
        ws.cell(tr, 1, "סה\"כ").font = F_BOLD
        ws.cell(tr, 2, "סה\"כ הוצאות").font = F_BOLD
        for j, (h, kind, val) in enumerate(cols, 3):
            if kind in ("dpct", "dabs", "share"):
                ws.cell(tr, j, cell_formula(kind, val, None, None, ent, ws, tr, j, cols))
            else:
                ws.cell(tr, j, "=" + "+".join(f"{L(j)}{x}" for x in dept_rows))
        for j in range(1, 3 + len(cols)):
            ws.cell(tr, j).fill, ws.cell(tr, j).border = FILL_TOT, BOX
            ws.cell(tr, j).font = F_BOLD
        for rr in range(hr + 1, tr + 1):
            for j, (h, kind, val) in enumerate(cols, 3):
                ws.cell(rr, j).number_format = PCT if kind in ("dpct", "share") else CUR
        ws.column_dimensions["A"].width = 11
        ws.column_dimensions["B"].width = 26
        for j in range(3, 3 + len(cols)):
            ws.column_dimensions[L(j)].width = 12.5
        ws.freeze_panes = ws.cell(hr + 1, 3)
        ws.sheet_properties.outlinePr.summaryBelow = False
        ws._total_row = tr
        return hr, tr

    def cell_formula(kind, val, dref, iref, ent, ws, r, j, cols):
        crit = f"D_Amt,D_Dept,{dref},D_Item,{iref},D_Ent,{ent}" if dref else None
        if kind == "month":
            return f'=SUMIFS({crit},D_Month,"{val}")'
        if kind == "q":
            return f'=SUMIFS({crit},D_Q,"{val}")'
        if kind == "year":
            return f'=SUMIFS({crit},D_Year,"{val}")'
        if kind == "dabs":  # val = (cur_offset, prev_offset) column indexes
            c, p = val
            return f"={L(c)}{r}-{L(p)}{r}"
        if kind == "dpct":
            c, p = val
            return f'=IF({L(p)}{r}=0,IF({L(c)}{r}=0,0,"חדש"),{L(c)}{r}/{L(p)}{r}-1)'
        if kind == "share":  # share of grand total for column val
            return f"=IF(${L(val)}$TOT=0,0,{L(val)}{r}/${L(val)}$TOT)"
        raise ValueError(kind)

    pcols = [(m, "month", m) for m in months] + [(q, "q", q) for q in quarters] + \
            [(f"סה\"כ {y}", "year", y) for y in years]
    tot_col = 2 + len(pcols)
    pcols.append(("% מסה\"כ", "share", tot_col))
    hr, tr = pnl_sheet(wsP, "דוח P&L - הוצאות לפי מחלקה וסעיף תקציבי ($)", pcols)
    # fix the TOT placeholder
    for row in wsP.iter_rows(min_row=hr + 1, max_row=tr):
        for c in row:
            if isinstance(c.value, str) and "$TOT" in c.value:
                c.value = c.value.replace("$TOT", f"${tr}")
    # visual separators between month/quarter/year blocks
    for j in (3 + len(months), 3 + len(months) + len(quarters)):
        for rr in range(hr, tr + 1):
            c = wsP.cell(rr, j)
            c.border = Border(left=c.border.left, right=Side(style="medium", color="1F3864"),
                              top=c.border.top, bottom=c.border.bottom)

    # ---------------------------------------------------------- QoQ
    qcols = [(q, "q", q) for q in quarters]
    for k in range(1, len(quarters)):
        cur, prev = 3 + k, 2 + k
        qcols.append((f"שינוי $ {quarters[k][-2:]} מול {quarters[k - 1][-2:]}", "dabs", (cur, prev)))
        qcols.append((f"שינוי % {quarters[k][-2:]} מול {quarters[k - 1][-2:]}", "dpct", (cur, prev)))
    hrq, trq = pnl_sheet(wsQ, "ניתוח רבעון מול רבעון (QoQ)", qcols)
    wsQ["D4"] = "עלייה בהוצאה מסומנת באדום, ירידה בירוק. \"חדש\" = לא הייתה הוצאה ברבעון הקודם."
    wsQ["D4"].font = F_SUB
    red = PatternFill("solid", fgColor="F8CBAD")
    green = PatternFill("solid", fgColor="C6EFCE")
    for j, (h, kind, val) in enumerate(qcols, 3):
        if kind in ("dpct", "dabs"):
            rng = f"{L(j)}{hrq + 1}:{L(j)}{trq}"
            wsQ.conditional_formatting.add(rng, CellIsRule(operator="greaterThan", formula=["0"], fill=red))
            wsQ.conditional_formatting.add(rng, CellIsRule(operator="lessThan", formula=["0"], fill=green))
    # QoQ chart - dept totals per quarter
    dept_rows_q = [r for r in range(hrq + 1, trq) if wsQ.cell(r, 2).value and str(wsQ.cell(r, 2).value).startswith("סה\"כ")]
    # helper table to the right for chart
    hc = 4 + len(qcols)
    wsQ.cell(hrq, hc, "מחלקה")
    for k, q in enumerate(quarters, 1):
        wsQ.cell(hrq, hc + k, q)
    style_header(wsQ, hrq, hc, hc + len(quarters))
    for n, r in enumerate(dept_rows_q, 1):
        wsQ.cell(hrq + n, hc, f"=A{r}")
        for k in range(1, len(quarters) + 1):
            wsQ.cell(hrq + n, hc + k, f"={L(2 + k)}{r}").number_format = CUR
    ch = bar("הוצאות לפי מחלקה ורבעון ($)")
    ch.add_data(Reference(wsQ, min_col=hc + 1, max_col=hc + len(quarters), min_row=hrq, max_row=hrq + len(dept_rows_q)), titles_from_data=True)
    ch.set_categories(Reference(wsQ, min_col=hc, min_row=hrq + 1, max_row=hrq + len(dept_rows_q)))
    qcolors = ["9EC5F4", "5598E7", "2A78D6", "1C5CAB", "104281", "0D366B"][-len(quarters):] if len(quarters) <= 6 else None
    if qcolors:
        color_series(ch, qcolors)
    wsQ.add_chart(ch, f"{L(hc)}{hrq + len(dept_rows_q) + 3}")
    for j in range(hc, hc + len(quarters) + 1):
        wsQ.column_dimensions[L(j)].width = 13

    # ---------------------------------------------------------- Vendors (dynamic)
    ws = wsV
    ws.sheet_view.rightToLeft = True
    ws.sheet_view.showGridLines = False
    ws["B1"] = "הרכב ספקים לפי סעיף"
    ws["B1"].font = F_TITLE
    ws["B2"] = "בחר מחלקה, סעיף ויישות (תאים צהובים). הטבלה מציגה את 30 הספקים הגדולים ממוינים לפי סכום. סכומים ב-$."
    ws["B2"].font = F_SUB
    for r, (lab, f, val) in enumerate([("מחלקה", "=L_Dept", depts[0]),
                                       ("סעיף תקציבי", "=L_Item", dept_items[depts[0]][0]),
                                       ("יישות", "=L_Ent", ALL)], 4):
        ws.cell(r, 2, lab).font = F_BOLD
        c = ws.cell(r, 3, val)
        c.fill, c.border, c.font = FILL_INPUT, BOX, F_BOLD
        add_list_validation(ws, f"C{r}", f, "בחר ערך")
    # criteria names on Calc
    for k, (name, src_cell) in enumerate([("V_Dept", "C4"), ("V_Item", "C5"), ("V_Ent", "C6")]):
        rr = 3 + len(calc_cells) + k
        wsC.cell(rr, 1, f"קריטריון Vendors {src_cell}")
        wsC.cell(rr, 2, f'=IF(Vendors!${src_cell[0]}${src_cell[1:]}="{ALL}","*",Vendors!${src_cell[0]}${src_cell[1:]})')
        add_name(wb, name, f"Calc!$B${rr}")
    ws["E4"] = "סה\"כ סעיף:"
    ws["E4"].font = F_BOLD
    ws["F4"] = "=SUM(C_VTot)"
    ws["F4"].number_format = CUR
    ws["F4"].font = Font(name=FONT, size=12, bold=True, color="1F3864")
    ws["E5"] = "מספר ספקים:"
    ws["E5"].font = F_BOLD
    ws["F5"] = '=COUNTIF(C_VTot,">0")'
    ws["E6"] = "חלק 5 הגדולים:"
    ws["E6"].font = F_BOLD
    hdrs = ["#", "ספק"] + quarters + ["סה\"כ", "% מהסעיף", "% מצטבר"]
    if len(quarters) > 1:
        hdrs.append(f"שינוי % {quarters[-1][-2:]} מול {quarters[-2][-2:]}")
    vr = 8
    for j, h in enumerate(hdrs, 2):
        ws.cell(vr, j, h)
    style_header(ws, vr, 2, 1 + len(hdrs))
    NV = 30
    nq = len(quarters)
    for k in range(1, NV + 1):
        r = vr + k
        m = f"MATCH(LARGE(C_VKey,{k}),C_VKey,0)"
        ws.cell(r, 2, k)
        ws.cell(r, 3, f'=IF(LARGE(C_VKey,{k})>0,INDEX(C_VName,{m}),"")')
        for q in range(nq):
            ws.cell(r, 4 + q, f'=IF($C{r}="","",INDEX(C_VQ,{m},{q + 1}))').number_format = CUR
        tc = 4 + nq
        ws.cell(r, tc, f'=IF($C{r}="","",INDEX(C_VTot,{m}))').number_format = CUR
        ws.cell(r, tc + 1, f'=IF(OR($C{r}="",$F$4=0),"",{L(tc)}{r}/$F$4)').number_format = PCT
        ws.cell(r, tc + 2, f'=IF($C{r}="","",SUM(${L(tc)}${vr + 1}:{L(tc)}{r})/$F$4)').number_format = PCT
        if nq > 1:
            a, b = L(3 + nq), L(2 + nq)
            ws.cell(r, tc + 3, f'=IF($C{r}="","",IF(N({b}{r})=0,IF(N({a}{r})=0,"","חדש"),{a}{r}/{b}{r}-1))').number_format = PCT
        for j in range(2, 2 + len(hdrs)):
            ws.cell(r, j).border = BOX
            ws.cell(r, j).font = F_BASE
    orow = vr + NV + 1
    tc = 4 + nq
    ws.cell(orow, 3, "שאר הספקים").font = F_BOLD
    for q in range(nq):
        col = L(4 + q)
        ws.cell(orow, 4 + q, f"=SUM(INDEX(C_VQ,0,{q + 1}))-SUM({col}{vr + 1}:{col}{vr + NV})").number_format = CUR
    ws.cell(orow, tc, f"=$F$4-SUM({L(tc)}{vr + 1}:{L(tc)}{vr + NV})").number_format = CUR
    ws.cell(orow + 1, 3, "סה\"כ").font = F_BOLD
    for j in range(4, tc + 1):
        ws.cell(orow + 1, j, f"=SUM({L(j)}{vr + 1}:{L(j)}{orow})").number_format = CUR
    ws.cell(orow + 1, tc + 1, '=IF($F$4=0,"",1)').number_format = PCT
    for rr in (orow, orow + 1):
        for j in range(2, 2 + len(hdrs)):
            ws.cell(rr, j).fill, ws.cell(rr, j).border = FILL_DEPT, BOX
            ws.cell(rr, j).font = F_BOLD
    ws["F6"] = f'=IF($F$4=0,"-",SUM({L(tc)}{vr + 1}:{L(tc)}{vr + 5})/$F$4)'
    ws["F6"].number_format = PCT
    if nq > 1:
        rng = f"{L(tc + 3)}{vr + 1}:{L(tc + 3)}{vr + NV}"
        ws.conditional_formatting.add(rng, CellIsRule(operator="greaterThan", formula=["0"], fill=red))
        ws.conditional_formatting.add(rng, CellIsRule(operator="lessThan", formula=["0"], fill=green))
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 5
    ws.column_dimensions["C"].width = 32
    for j in range(4, 2 + len(hdrs)):
        ws.column_dimensions[L(j)].width = 13
    ws.column_dimensions["E"].width = 15
    ch = bar("10 הספקים המובילים בסעיף ($)", horizontal=True, w=16, h=10)
    ch.add_data(Reference(ws, min_col=tc, min_row=vr + 1, max_row=vr + 10), titles_from_data=False)
    ch.set_categories(Reference(ws, min_col=3, min_row=vr + 1, max_row=vr + 10))
    ch.x_axis.scaling.orientation = "maxMin"
    ch.legend = None
    color_series(ch, [SINGLE])
    ws.add_chart(ch, f"{L(3 + len(hdrs))}4")
    ch = bar("ספקים מובילים לפי רבעון ($)", stacked=True, w=16, h=10)
    ch.add_data(Reference(ws, min_col=4, max_col=3 + nq, min_row=vr, max_row=vr + 10), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=3, min_row=vr + 1, max_row=vr + 10))
    if qcolors:
        color_series(ch, qcolors)
    ch.type = "bar"
    ch.x_axis.scaling.orientation = "maxMin"
    ws.add_chart(ch, f"{L(3 + len(hdrs))}25")
    ws.freeze_panes = f"A{vr + 1}"

    # ---------------------------------------------------------- Vendor Detail (static grouped)
    ws = wsVD
    ws.sheet_view.rightToLeft = True
    ws.sheet_view.showGridLines = False
    ws["A1"] = "פירוט ספקים: מחלקה > סעיף > ספק ($)"
    ws["A1"].font = F_TITLE
    ws["A3"] = "יישות:"
    ws["A3"].font = F_BOLD
    ws["C3"] = ALL
    ws["C3"].fill, ws["C3"].border, ws["C3"].font = FILL_INPUT, BOX, F_BOLD
    add_list_validation(ws, "C3", "=L_Ent", "בחר יישות או הכל")
    ws["D3"] = "רמה 1 = מחלקה, רמה 2 = סעיף, רמה 3 = ספקים. השתמש בכפתורי 1/2/3 בצד כדי לכווץ."
    ws["D3"].font = F_SUB
    ent = f'IF($C$3="{ALL}","*",$C$3)'
    hdr = ["מחלקה", "סעיף", "ספק"] + quarters + ["סה\"כ", "% מהסעיף"]
    if nq > 1:
        hdr.append(f"שינוי % {quarters[-1][-2:]} מול {quarters[-2][-2:]}")
    hr = 5
    for j, h in enumerate(hdr, 1):
        ws.cell(hr, j, h)
    style_header(ws, hr, 1, len(hdr))
    totals = defaultdict(float)
    for r in rows:
        totals[(r["dept"], r["item"], r["vendor"])] += r["amt"]
    r = hr + 1
    tcol = 4 + nq
    pc = tcol + 1
    for d in depts:
        dr = r
        ws.cell(dr, 1, d)
        ws.cell(dr, 2, f"סה\"כ {d}")
        item_rows = []
        r += 1
        for it in dept_items[d]:
            ir = r
            item_rows.append(ir)
            ws.cell(ir, 1, d).font = F_GREY
            ws.cell(ir, 2, it)
            ws.row_dimensions[ir].outlineLevel = 1
            vs = sorted([v for (dd, ii, v) in totals if dd == d and ii == it],
                        key=lambda v: -totals[(d, it, v)])
            r += 1
            for v in vs:
                ws.cell(r, 1, d).font = F_GREY
                ws.cell(r, 2, it).font = F_GREY
                ws.cell(r, 3, v)
                ws.cell(r, 3).alignment = Alignment(indent=1)
                ws.row_dimensions[r].outlineLevel = 2
                for q in range(nq):
                    ws.cell(r, 4 + q, f'=SUMIFS(D_Amt,D_Dept,$A{r},D_Item,$B{r},D_VID,{vid[v]},D_Ent,{ent},D_Q,"{quarters[q]}")')
                ws.cell(r, tcol, f"=SUM({L(4)}{r}:{L(3 + nq)}{r})")
                ws.cell(r, pc, f"=IF(${L(tcol)}${ir}=0,0,{L(tcol)}{r}/${L(tcol)}${ir})")
                r += 1
            for j in range(4, tcol + 1):
                ws.cell(ir, j, f"=SUM({L(j)}{ir + 1}:{L(j)}{r - 1})" if vs else 0)
            ws.cell(ir, pc, f"=IF(${L(tcol)}${dr}=0,0,{L(tcol)}{ir}/${L(tcol)}${dr})")
            for j in range(1, len(hdr) + 1):
                ws.cell(ir, j).fill = FILL_KPI
                ws.cell(ir, j).font = F_BOLD if j > 1 else F_GREY
        for j in range(4, tcol + 1):
            ws.cell(dr, j, "=" + "+".join(f"{L(j)}{x}" for x in item_rows))
        ws.cell(dr, pc, "")
        for j in range(1, len(hdr) + 1):
            ws.cell(dr, j).fill, ws.cell(dr, j).border = FILL_DEPT, BOX
            ws.cell(dr, j).font = F_BOLD
    vd_last = r - 1
    for rr in range(hr + 1, vd_last + 1):
        if nq > 1:
            a, b = L(3 + nq), L(2 + nq)
            if ws.cell(rr, pc).value != "":
                ws.cell(rr, pc + 1, f'=IF({b}{rr}=0,IF({a}{rr}=0,"","חדש"),{a}{rr}/{b}{rr}-1)')
            ws.cell(rr, pc + 1).number_format = PCT
        for j in range(4, tcol + 1):
            ws.cell(rr, j).number_format = CUR
        ws.cell(rr, pc).number_format = PCT
    if nq > 1:
        rng = f"{L(pc + 1)}{hr + 1}:{L(pc + 1)}{vd_last}"
        ws.conditional_formatting.add(rng, CellIsRule(operator="greaterThan", formula=["0"], fill=red))
        ws.conditional_formatting.add(rng, CellIsRule(operator="lessThan", formula=["0"], fill=green))
    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 24
    ws.column_dimensions["C"].width = 34
    for j in range(4, len(hdr) + 1):
        ws.column_dimensions[L(j)].width = 13
    ws.freeze_panes = "D6"
    ws.sheet_properties.outlinePr.summaryBelow = False
    ws.auto_filter.ref = f"A{hr}:{L(len(hdr))}{vd_last}"

    # ---------------------------------------------------------- finish
    for wsx in wb.worksheets:
        set_font_all(wsx)
    wsL.sheet_state = "hidden"
    wsC.sheet_properties.tabColor = "A6A6A6"
    wsL.sheet_properties.tabColor = "A6A6A6"
    wsD.sheet_properties.tabColor = "1F3864"
    wb.active = 0
    wb.calculation.fullCalcOnLoad = True
    wb.save(out)
    return dict(rows=len(rows), skipped=skipped, vendors=len(vnames), months=months,
                total=sum(r["amt"] for r in rows), mapping=mapping)


if __name__ == "__main__":
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    src = sys.argv[1]
    out = sys.argv[2] if len(sys.argv) > 2 else "PL_Analysis.xlsx"
    info = build(src, out)
    print(f"rows={info['rows']} vendors={info['vendors']} months={info['months'][0]}..{info['months'][-1]} "
          f"total=${info['total']:,.2f}")
    for s in info["skipped"]:
        print("skipped (non-numeric amount):", s)
    print("saved", out)
