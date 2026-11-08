"""
אפליקציית Streamlit לניתוח P&L לפי ספקים.

מעלים את קובץ הספקים (לשוניות חודשיות בשם MM.YY, טורים A-D: סעיף | ספק | סכום | Entity),
ומקבלים:
  * ניתוח אינטראקטיבי בדפדפן (דשבורד, P&L, רבעון מול רבעון, הרכב ספקים)
  * הורדה של קובץ Excel מלא (אותו קובץ ש-build_pnl.py מייצר)

הרצה:
    streamlit run app.py
"""
import io

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from build_pnl import ALL, DEPT_ORDER, build, load_dataframe

st.set_page_config(page_title="ניתוח P&L לפי ספקים", page_icon="📊", layout="wide")

# fixed colour per department: colour follows the department, not its rank
DEPT_COLORS = ["#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4", "#008300", "#7f7f7f"]
Q_COLORS = ["#9ec5f4", "#5598e7", "#2a78d6", "#1c5cab", "#104281", "#0d366b"]
INCREASE, DECREASE = "#f8cbad", "#c6efce"

st.markdown(
    """
    <style>
    .stApp, .stMarkdown, .stTabs, section[data-testid="stSidebar"] { direction: rtl; text-align: right; }
    div[data-testid="stMetricValue"] { direction: ltr; text-align: right; }
    .stDataFrame, .js-plotly-plot { direction: ltr; }
    </style>
    """,
    unsafe_allow_html=True,
)


# ----------------------------------------------------------------- helpers
def money(x):
    if x is None or pd.isna(x):
        return "-"
    return f"(${abs(x):,.0f})" if x < 0 else f"${x:,.0f}"


def pct(x):
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return "-"
    if isinstance(x, str):
        return x
    if x in (float("inf"), float("-inf")):
        return "חדש"
    return f"{x:+.1%}"


def pct_change(cur, prev):
    """float always (keeps the column numeric): inf = new spend, nan = nothing in both"""
    if prev == 0:
        return float("nan") if cur == 0 else (float("inf") if cur > 0 else float("-inf"))
    return cur / prev - 1


def color_change(v):
    if isinstance(v, (int, float)) and not pd.isna(v):
        if v > 0:
            return f"background-color: {INCREASE}"
        if v < 0:
            return f"background-color: {DECREASE}"
    return ""


@st.cache_data(show_spinner="קורא את הקובץ...")
def load(data: bytes):
    df, skipped = load_dataframe(io.BytesIO(data))
    return df, skipped


@st.cache_data(show_spinner="בונה את קובץ ה-Excel...")
def build_excel(data: bytes) -> bytes:
    out = io.BytesIO()
    build(io.BytesIO(data), out)
    return out.getvalue()


def dept_list(df):
    found = df["dept"].unique().tolist()
    return [d for d in DEPT_ORDER if d in found] + sorted(set(found) - set(DEPT_ORDER))


def dept_color_map(depts):
    return {d: DEPT_COLORS[i % len(DEPT_COLORS)] for i, d in enumerate(depts)}


def pnl_table(df, depts, columns):
    """rows: dept subtotal then its items; columns: list of (label, mask-column, value)"""
    out = []
    for d in depts:
        sub = df[df["dept"] == d]
        rec = {"מחלקה": d, "סעיף תקציבי": f'סה"כ {d}'}
        for lab, col, val in columns:
            rec[lab] = sub.loc[sub[col] == val, "amount"].sum() if col else sub["amount"].sum()
        out.append(rec)
        for it in sorted(sub["item"].unique()):
            s2 = sub[sub["item"] == it]
            rec = {"מחלקה": d, "סעיף תקציבי": "    " + it}
            for lab, col, val in columns:
                rec[lab] = s2.loc[s2[col] == val, "amount"].sum() if col else s2["amount"].sum()
            out.append(rec)
    rec = {"מחלקה": 'סה"כ', "סעיף תקציבי": 'סה"כ הוצאות'}
    for lab, col, val in columns:
        rec[lab] = df.loc[df[col] == val, "amount"].sum() if col else df["amount"].sum()
    out.append(rec)
    return pd.DataFrame(out)


def bold_subtotals(row):
    s = str(row["סעיף תקציבי"])
    if s.startswith('סה"כ'):
        return ["font-weight: bold; background-color: #d9e2f3"] * len(row)
    return [""] * len(row)


# ----------------------------------------------------------------- upload
st.title("📊 ניתוח P&L לפי ספקים")

with st.sidebar:
    st.header("העלאת קובץ")
    up = st.file_uploader("קובץ ספקים (xlsx)", type=["xlsx", "xlsm"])
    st.caption("כל לשונית בשם MM.YY (למשל 01.26) היא חודש. טורים A-D: "
               "סעיף | ספק | סכום | Entity. לשוניות אחרות לא נקראות. הסכומים בדולרים.")

if up is None:
    st.info("העלה את קובץ הספקים בסרגל הצד כדי לקבל את הניתוח ואת קובץ ה-Excel.")
    st.stop()

data = up.getvalue()
try:
    df, skipped = load(data)
except ValueError as e:
    st.error(str(e))
    st.stop()
except Exception as e:  # noqa: BLE001
    st.error(f"לא הצלחתי לקרוא את הקובץ: {e}")
    st.stop()

depts = dept_list(df)
cmap = dept_color_map(depts)
months = sorted(df["month"].unique())
quarters = sorted(df["quarter"].unique())
years = sorted(df["year"].unique())

with st.sidebar:
    st.divider()
    st.subheader("קובץ Excel מלא")
    try:
        xlsx = build_excel(data)
        base = up.name.rsplit(".", 1)[0]
        st.download_button("⬇️ הורדת דוח P&L (Excel)", xlsx, file_name=f"{base}_PL_Analysis.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           type="primary", width="stretch")
        st.caption("הקובץ כולל דשבורד עם פילטרים, P&L, QoQ והרכב ספקים. "
                   "הנוסחאות מחושבות כשפותחים אותו ב-Excel.")
    except Exception as e:  # noqa: BLE001
        st.error(f"יצירת קובץ ה-Excel נכשלה: {e}")
    st.divider()
    st.caption(f"{len(df):,} תנועות | {df['vendor'].nunique():,} ספקים | {months[0]} עד {months[-1]}")

if skipped:
    with st.expander(f"⚠️ {len(skipped)} שורות לא נקראו (סכום לא מספרי)"):
        st.dataframe(pd.DataFrame(skipped, columns=["לשונית", "שורה", "ערכים"]), hide_index=True)

tab_dash, tab_pnl, tab_qoq, tab_vend, tab_data = st.tabs(
    ["דשבורד", "P&L", "רבעון מול רבעון", "הרכב ספקים", "נתונים"])

# ----------------------------------------------------------------- dashboard
with tab_dash:
    c1, c2, c3, c4, c5 = st.columns([1.2, 1.6, 0.8, 0.8, 1])
    sel_dept = c1.multiselect("מחלקה", depts, placeholder=ALL)
    items_avail = sorted(df[df["dept"].isin(sel_dept)]["item"].unique()) if sel_dept \
        else sorted(df["item"].unique())
    sel_item = c2.multiselect("סעיף תקציבי", items_avail, placeholder=ALL)
    sel_ent = c3.selectbox("יישות", [ALL] + sorted(df["entity"].unique()))
    ptype = c4.selectbox("סוג תקופה", [ALL, "חודש", "רבעון", "שנה"], index=2)
    plist = {"חודש": months, "רבעון": quarters, "שנה": years}.get(ptype, [])
    pcol = {"חודש": "month", "רבעון": "quarter", "שנה": "year"}.get(ptype)
    pval = c5.selectbox("תקופה", plist, index=len(plist) - 1) if plist else None

    f = df
    if sel_dept:
        f = f[f["dept"].isin(sel_dept)]
    if sel_item:
        f = f[f["item"].isin(sel_item)]
    if sel_ent != ALL:
        f = f[f["entity"] == sel_ent]
    fp = f[f[pcol] == pval] if pcol else f
    prev = plist[plist.index(pval) - 1] if pcol and plist.index(pval) > 0 else None
    fprev = f[f[pcol] == prev] if prev else None

    total = fp["amount"].sum()
    k = st.columns(5)
    if prev:
        ptot = fprev["amount"].sum()
        chg = pct_change(total, ptot)
        k[0].metric("סה\"כ הוצאות בתקופה", money(total),
                    delta=pct(chg) if not pd.isna(chg) else None, delta_color="inverse")
        k[1].metric(f"תקופה קודמת ({prev})", money(ptot))
        k[2].metric("שינוי ($)", money(total - ptot))
    else:
        k[0].metric("סה\"כ הוצאות בתקופה", money(total))
        k[1].metric("תקופה קודמת", "-")
        k[2].metric("שינוי ($)", "-")
    vsum = fp.groupby("vendor")["amount"].sum()
    k[3].metric("ספקים עם הוצאה", f"{(vsum > 0).sum():,}")
    k[4].metric("מספר תנועות", f"{len(fp):,}")
    st.caption("פילטר התקופה משפיע על המדדים, ההתפלגות והדירוגים. גרפי המגמה מציגים את כל התקופות. "
               "החץ במדד הראשון: אדום = עלייה בהוצאה.")

    g1, g2 = st.columns(2)
    m = f.groupby(["month", "dept"], as_index=False)["amount"].sum()
    fig = px.bar(m, x="month", y="amount", color="dept", color_discrete_map=cmap,
                 category_orders={"dept": depts, "month": months},
                 labels={"month": "חודש", "amount": "סכום ($)", "dept": "מחלקה"},
                 title="מגמה חודשית לפי מחלקה")
    fig.update_traces(marker_line_color="white", marker_line_width=1,
                      hovertemplate="%{x}<br>%{fullData.name}: $%{y:,.0f}<extra></extra>")
    fig.update_layout(yaxis_tickformat="$,.0f", legend_title_text="", bargap=0.35)
    g1.plotly_chart(fig, width="stretch")

    qd = f.groupby(["quarter", "dept"], as_index=False)["amount"].sum()
    fig = px.bar(qd, x="quarter", y="amount", color="dept", barmode="group", color_discrete_map=cmap,
                 category_orders={"dept": depts, "quarter": quarters},
                 labels={"quarter": "רבעון", "amount": "סכום ($)", "dept": "מחלקה"},
                 title="השוואה רבעונית לפי מחלקה")
    fig.update_traces(hovertemplate="%{x}<br>%{fullData.name}: $%{y:,.0f}<extra></extra>")
    fig.update_layout(yaxis_tickformat="$,.0f", legend_title_text="")
    g2.plotly_chart(fig, width="stretch")

    g3, g4 = st.columns(2)
    sd = fp.groupby("dept", as_index=False)["amount"].sum()
    sd = sd[sd["amount"] > 0]
    fig = go.Figure(go.Pie(labels=sd["dept"], values=sd["amount"], hole=0.55, sort=False,
                           marker=dict(colors=[cmap[d] for d in sd["dept"]],
                                       line=dict(color="white", width=2)),
                           hovertemplate="%{label}: $%{value:,.0f} (%{percent})<extra></extra>"))
    fig.update_layout(title="התפלגות לפי מחלקה - תקופה נבחרת")
    g3.plotly_chart(fig, width="stretch")

    top = vsum.sort_values(ascending=False).head(10)[::-1]
    fig = go.Figure(go.Bar(x=top.values, y=top.index, orientation="h", marker_color=DEPT_COLORS[0],
                           hovertemplate="%{y}: $%{x:,.0f}<extra></extra>"))
    fig.update_layout(title="10 הספקים המובילים - תקופה נבחרת", xaxis_tickformat="$,.0f")
    g4.plotly_chart(fig, width="stretch")

    t1, t2 = st.columns(2)
    tv = vsum.sort_values(ascending=False).head(10).reset_index()
    tv.columns = ["ספק", "סכום"]
    tv["% מהסה\"כ"] = tv["סכום"] / total if total else 0
    t1.markdown("**10 הספקים המובילים**")
    t1.dataframe(tv.style.format({"סכום": money, "% מהסה\"כ": "{:.1%}"}), hide_index=True,
                 width="stretch")
    ti = fp.groupby(["dept", "item"])["amount"].sum().sort_values(ascending=False).head(10).reset_index()
    ti["סעיף"] = ti["dept"] + " - " + ti["item"]
    ti = ti[["סעיף", "amount"]].rename(columns={"amount": "סכום"})
    ti["% מהסה\"כ"] = ti["סכום"] / total if total else 0
    t2.markdown("**10 הסעיפים המובילים**")
    t2.dataframe(ti.style.format({"סכום": money, "% מהסה\"כ": "{:.1%}"}), hide_index=True,
                 width="stretch")

# ----------------------------------------------------------------- P&L
with tab_pnl:
    ent = st.selectbox("יישות", [ALL] + sorted(df["entity"].unique()), key="pnl_ent")
    d2 = df if ent == ALL else df[df["entity"] == ent]
    cols = [(m_, "month", m_) for m_ in months] + [(q, "quarter", q) for q in quarters] + \
           [(f'סה"כ {y}', "year", y) for y in years]
    t = pnl_table(d2, depts, cols)
    tot_lab = f'סה"כ {years[-1]}'
    grand = t[tot_lab].iloc[-1]
    t['% מסה"כ'] = t[tot_lab] / grand if grand else 0
    num = [c for c in t.columns if c not in ("מחלקה", "סעיף תקציבי")]
    st.dataframe(
        t.style.apply(bold_subtotals, axis=1)
        .format({c: ("{:.1%}" if c == '% מסה"כ' else money) for c in num}),
        hide_index=True, width="stretch", height=min(38 * (len(t) + 1), 900))

# ----------------------------------------------------------------- QoQ
with tab_qoq:
    ent = st.selectbox("יישות", [ALL] + sorted(df["entity"].unique()), key="qoq_ent")
    d2 = df if ent == ALL else df[df["entity"] == ent]
    if len(quarters) < 2:
        st.info("צריך לפחות שני רבעונים כדי להשוות.")
    else:
        t = pnl_table(d2, depts, [(q, "quarter", q) for q in quarters])
        chg_cols = []
        for i in range(1, len(quarters)):
            a, b = quarters[i], quarters[i - 1]
            la, lb = a[-2:], b[-2:]
            t[f"שינוי $ {la} מול {lb}"] = t[a] - t[b]
            t[f"שינוי % {la} מול {lb}"] = [pct_change(x, y) for x, y in zip(t[a], t[b])]
            chg_cols += [f"שינוי $ {la} מול {lb}", f"שינוי % {la} מול {lb}"]
        st.caption('עלייה בהוצאה מסומנת באדום, ירידה בירוק. "חדש" = לא הייתה הוצאה ברבעון הקודם.')
        fmt = {q: money for q in quarters}
        fmt.update({c: (money if "$" in c else pct) for c in chg_cols})
        st.dataframe(
            t.style.apply(bold_subtotals, axis=1).map(color_change, subset=chg_cols).format(fmt),
            hide_index=True, width="stretch", height=min(38 * (len(t) + 1), 900))

        dq = d2.groupby(["dept", "quarter"], as_index=False)["amount"].sum()
        fig = px.bar(dq, x="dept", y="amount", color="quarter", barmode="group",
                     category_orders={"dept": depts, "quarter": quarters},
                     color_discrete_sequence=Q_COLORS[-len(quarters):] if len(quarters) <= 6 else None,
                     labels={"dept": "מחלקה", "amount": "סכום ($)", "quarter": "רבעון"},
                     title="הוצאות לפי מחלקה ורבעון")
        fig.update_traces(hovertemplate="%{x} %{fullData.name}: $%{y:,.0f}<extra></extra>")
        fig.update_layout(yaxis_tickformat="$,.0f", legend_title_text="")
        st.plotly_chart(fig, width="stretch")

# ----------------------------------------------------------------- vendors
with tab_vend:
    v1, v2, v3 = st.columns(3)
    vd = v1.selectbox("מחלקה", depts, key="v_dept")
    vi = v2.selectbox("סעיף תקציבי", sorted(df[df["dept"] == vd]["item"].unique()), key="v_item")
    ve = v3.selectbox("יישות", [ALL] + sorted(df["entity"].unique()), key="v_ent")
    s = df[(df["dept"] == vd) & (df["item"] == vi)]
    if ve != ALL:
        s = s[s["entity"] == ve]
    if s.empty:
        st.info("אין הוצאות לבחירה הזו.")
    else:
        pv = s.pivot_table(index="vendor", columns="quarter", values="amount", aggfunc="sum",
                           fill_value=0).reindex(columns=quarters, fill_value=0)
        pv['סה"כ'] = pv.sum(axis=1)
        pv = pv.sort_values('סה"כ', ascending=False)
        tot = pv['סה"כ'].sum()
        pv["% מהסעיף"] = pv['סה"כ'] / tot if tot else 0
        pv["% מצטבר"] = pv["% מהסעיף"].cumsum()
        if len(quarters) > 1:
            a, b = quarters[-1], quarters[-2]
            lab = f"שינוי % {a[-2:]} מול {b[-2:]}"
            pv[lab] = [pct_change(x, y) for x, y in zip(pv[a], pv[b])]
        else:
            lab = None
        m1, m2, m3 = st.columns(3)
        m1.metric("סה\"כ סעיף", money(tot))
        n_pos = int((pv['סה"כ'] > 0).sum())
        m2.metric("ספקים עם הוצאה", f"{n_pos:,}")
        m3.metric("חלק 5 הגדולים", f"{pv['% מהסעיף'].head(5).sum():.1%}")
        fmt = {c: money for c in quarters + ['סה"כ']}
        fmt.update({"% מהסעיף": "{:.1%}", "% מצטבר": "{:.1%}"})
        if lab:
            fmt[lab] = pct
        tbl = pv.reset_index().rename(columns={"vendor": "ספק"})
        tbl.columns = [str(c) for c in tbl.columns]
        sty = tbl.style.format(fmt)
        if lab:
            sty = sty.map(color_change, subset=[lab])
        st.dataframe(sty, hide_index=True, width="stretch")

        top = pv.head(10)
        c1, c2 = st.columns(2)
        fig = go.Figure(go.Bar(x=top['סה"כ'][::-1], y=top.index[::-1], orientation="h",
                               marker_color=DEPT_COLORS[0],
                               hovertemplate="%{y}: $%{x:,.0f}<extra></extra>"))
        fig.update_layout(title="10 הספקים המובילים בסעיף", xaxis_tickformat="$,.0f")
        c1.plotly_chart(fig, width="stretch")
        long = top[quarters].reset_index().melt(id_vars="vendor", var_name="quarter", value_name="amount")
        fig = px.bar(long, x="amount", y="vendor", color="quarter", orientation="h",
                     category_orders={"vendor": list(top.index), "quarter": quarters},
                     color_discrete_sequence=Q_COLORS[-len(quarters):] if len(quarters) <= 6 else None,
                     labels={"amount": "סכום ($)", "vendor": "", "quarter": "רבעון"},
                     title="ספקים מובילים לפי רבעון")
        fig.update_traces(marker_line_color="white", marker_line_width=1,
                          hovertemplate="%{y} %{fullData.name}: $%{x:,.0f}<extra></extra>")
        fig.update_layout(xaxis_tickformat="$,.0f", legend_title_text="")
        c2.plotly_chart(fig, width="stretch")

# ----------------------------------------------------------------- data
with tab_data:
    show = df.rename(columns={"sheet": "לשונית", "month": "חודש", "quarter": "רבעון", "year": "שנה",
                              "dept": "מחלקה", "item": "סעיף תקציבי", "vendor": "ספק",
                              "entity": "יישות", "amount": "סכום ($)", "item_raw": "סעיף מקורי",
                              "vendor_raw": "ספק מקורי"})
    st.dataframe(show, hide_index=True, width="stretch", height=600)
    st.download_button("⬇️ הורדת הנתונים המאוחדים (CSV)", show.to_csv(index=False).encode("utf-8-sig"),
                       file_name="pl_data.csv", mime="text/csv")
