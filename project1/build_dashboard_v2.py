#!/usr/bin/env python3
"""
build_dashboard.py  -  Online Retail II  |  Sales Performance Dashboard (Excel, forest-green sidebar theme)
==============================================================================

Builds a professional, fully interactive Excel dashboard from the UCI
"Online Retail II" dataset (online_retail_II.csv).

What you get (same structure as the reference workbook)
-------------------------------------------------------
Dashboard     the report: 3 drop-down filters in a left sidebar,
              8 KPI cards on top, 5 charts, top-10 customers, RFM segment table and a
              weekday x hour heat-map. Everything is a LIVE FORMULA - change a
              yellow cell and the whole page recalculates (no macros).
Notes         how to use, definitions, data-cleaning log (counts are computed
              by this script), assumptions.
Calc          the calculation engine feeding the dashboard.
Customers     one row per customer incl. RFM scoring formulas.
Fact_Monthly / Fact_Product / Fact_Time   pre-aggregated data model.
Lists         drop-down sources.

Usage
-----
    pip install pandas numpy xlsxwriter
    python build_dashboard.py --csv online_retail_II.csv --out Online_Retail_Dashboard.xlsx

Open the result in Microsoft Excel (formulas calculate on load).
"""
import argparse
import re
import sys

import numpy as np
import pandas as pd
import xlsxwriter
from xlsxwriter.utility import xl_col_to_name

# --------------------------------------------------------------------------- #
# Configuration
# --------------------------------------------------------------------------- #
FONT = "Arial"
# Forest green + cream palette
FOREST, LEAF, SAGE, MOSS = "#1F3D2B", "#2D5A3D", "#6B9080", "#A4C3B2"
GOLD, TERRA, RUST, TAUPE = "#C9A227", "#B5651D", "#A44A3F", "#9C8F7A"
CREAM, CARD, SAND, INK, MUTED, FILTER = "#F6F1E7", "#FFFDF8", "#E3D9C2", "#2F3E34", "#7A8B7F", "#F3E3A6"
# aliases used by the Calc / Notes sheets
NAVY, TEAL, PURPLE, GREEN, ORANGE, GREY, BLUE = FOREST, LEAF, SAGE, LEAF, GOLD, MUTED, SAGE
BG, LINE, TXT, YELLOW = CREAM, SAND, INK, FILTER

TOP_PRODUCTS = 500                  # products kept in the data model
HOUR_MIN, HOUR_MAX = 6, 20          # heat-map hours (others folded into the edges)
OVERLAP_START, OVERLAP_END = "2010-12-01", "2010-12-10"   # period present in both year-sheets
HOME_COUNTRY = "United Kingdom"
HOME_SHORT = "UK"                  # label used in the markets caption

# Stock codes that are not merchandise (postage, fees, adjustments, vouchers ...)
NON_MERCH = re.compile(
    r"^(POST|DOT|M|C2|C3|D|S|B|CRUK|PADS|BANK CHARGES|ADJUST2?|AMAZONFEE|SP\d+|TEST.*|GIFT.*|DCGS.*)$")

SEGMENTS = ["Champions", "Loyal", "Recent / New", "At Risk", "Hibernating / Lost"]
SEG_COLORS = ["#3E8E5A", SAGE, GOLD, TERRA, TAUPE]
METRICS = ["Net Revenue", "Gross Sales", "Returns", "Orders", "Units Sold", "Avg Order Value"]
DAYS = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]


# --------------------------------------------------------------------------- #
# 1. Load + clean + aggregate
# --------------------------------------------------------------------------- #
def load_and_clean(csv_path):
    raw = pd.read_csv(csv_path, dtype={"Invoice": str, "StockCode": str},
                      parse_dates=["InvoiceDate"])
    log = {"rows_source": len(raw)}
    d = raw.copy()
    d["Invoice"] = d["Invoice"].str.strip().str.upper()
    d["StockCode"] = d["StockCode"].str.strip().str.upper()
    d["Description"] = d["Description"].fillna("").str.strip()

    # a) overlap between the two year-sheets -> keep once
    in_window = (d.InvoiceDate >= OVERLAP_START) & (d.InvoiceDate < OVERLAP_END)
    dup = d.duplicated(keep="first") & in_window
    log["overlap"] = int(dup.sum())
    d = d[~dup]

    # b) non-merchandise lines and bad-debt adjustment invoices ('A...')
    bad = d.StockCode.str.match(NON_MERCH) | d.Invoice.str.startswith("A")
    log["non_merch"] = int(bad.sum())
    d = d[~bad]

    # c) zero / negative price
    log["bad_price"] = int((d.Price <= 0).sum())
    d = d[d.Price > 0]

    # d) negative quantity on a normal (non-cancel) invoice
    cancel = d.Invoice.str.startswith("C")
    odd = (~cancel) & (d.Quantity < 0)
    log["other"] = int(odd.sum())
    d = d[~odd].copy()

    d["cancel"] = d.Invoice.str.startswith("C")
    d["rev"] = d.Quantity * d.Price
    d["gross"] = np.where(~d.cancel, d.rev, 0.0)
    d["ret"] = np.where(d.cancel, -d.rev, 0.0)
    d["units"] = np.where(~d.cancel, d.Quantity, 0)
    d["Month"] = d.InvoiceDate.dt.to_period("M").dt.to_timestamp()
    d["Year"] = d.InvoiceDate.dt.year
    d["Hour"] = d.InvoiceDate.dt.hour.clip(HOUR_MIN, HOUR_MAX)
    d["Weekday"] = d.InvoiceDate.dt.dayofweek + 1            # Mon = 1 ... Sun = 7
    log["clean_rows"] = len(d)
    log["no_cust_rows"] = int(d["Customer ID"].isna().sum())
    log["no_cust_share"] = float(d.loc[d["Customer ID"].isna(), "rev"].sum() / d.rev.sum())
    return d, log


def build_facts(d):
    f = {}
    # monthly x country
    g = d.groupby(["Month", "Country"])[["gross", "ret", "units"]].sum()
    g["orders"] = d[~d.cancel].groupby(["Month", "Country"]).Invoice.nunique()
    g = g.fillna(0).reset_index()
    g["Year"], g["MonthNum"] = g.Month.dt.year, g.Month.dt.month
    f["monthly"] = g[["Month", "Year", "MonthNum", "Country", "gross", "ret", "orders", "units"]]

    # product x year x country (top N by lifetime net revenue)
    top = d.groupby("StockCode").rev.sum().nlargest(TOP_PRODUCTS).index
    p = d[d.StockCode.isin(top)]
    desc = (p[p.Description != ""].groupby("StockCode").Description
            .agg(lambda s: s.value_counts().index[0]))
    prod = (p.groupby(["StockCode", "Year", "Country"])
            .agg(net=("rev", "sum"), units=("Quantity", "sum")).reset_index())
    prod["Description"] = prod.StockCode.map(desc).fillna(prod.StockCode).str.title()
    f["product"] = prod[["StockCode", "Description", "Year", "Country", "net", "units"]]
    order = d[d.StockCode.isin(top)].groupby("StockCode").rev.sum().sort_values(ascending=False).index
    plist = prod.drop_duplicates("StockCode").set_index("StockCode").loc[order, "Description"]
    f["plist"] = list(zip(plist.index, plist.values))

    # weekday x hour x country x year (gross sales)
    t = (d[~d.cancel].groupby(["Year", "Country", "Weekday", "Hour"]).gross.sum().reset_index())
    f["time"] = t

    # customers
    c = d[d["Customer ID"].notna()].copy()
    c["cid"] = c["Customer ID"].astype(int)
    cc = c.groupby(["cid", "Country"]).rev.sum().reset_index()
    cc = cc.sort_values(["cid", "rev"], ascending=[True, False]).drop_duplicates("cid")
    cust = cc.set_index("cid")[["Country"]]
    nc = c[~c.cancel]
    cust["first"] = nc.groupby("cid").InvoiceDate.min().dt.normalize()
    cust["last"] = nc.groupby("cid").InvoiceDate.max().dt.normalize()
    cust["first"] = cust["first"].fillna(c.groupby("cid").InvoiceDate.min().dt.normalize())
    cust["last"] = cust["last"].fillna(c.groupby("cid").InvoiceDate.max().dt.normalize())
    cust["orders"] = nc.groupby("cid").Invoice.nunique()
    cust["orders"] = cust["orders"].fillna(0).astype(int)
    cust["net"] = c.groupby("cid").rev.sum()
    years = sorted(d.Year.unique())
    ry = c.pivot_table(index="cid", columns="Year", values="rev", aggfunc="sum").reindex(columns=years).fillna(0)
    oy = nc.pivot_table(index="cid", columns="Year", values="Invoice", aggfunc="nunique").reindex(columns=years).fillna(0)
    ry, oy = ry.reindex(cust.index).fillna(0), oy.reindex(cust.index).fillna(0)
    for y in years:
        cust[f"rev{y}"] = ry[y]
    for y in years:
        cust[f"ord{y}"] = oy[y].astype(int)
    f["customers"] = cust.sort_values("net", ascending=False).reset_index()
    f["years"] = years
    f["countries"] = sorted(d.Country.unique())
    return f


# --------------------------------------------------------------------------- #
# 2. Workbook
# --------------------------------------------------------------------------- #
def build_workbook(path, d, log, f):
    wb = xlsxwriter.Workbook(path, {"default_date_format": "dd mmm yyyy"})
    wb.set_calc_mode("auto")
    cache = {}

    def F(**kw):
        base = {"font_name": FONT, "font_size": 9, "valign": "vcenter"}
        base.update(kw)
        key = tuple(sorted(base.items()))
        if key not in cache:
            cache[key] = wb.add_format(base)
        return cache[key]

    # worksheets in tab order
    ws_d = wb.add_worksheet("Dashboard")
    ws_n = wb.add_worksheet("Notes")
    ws_c = wb.add_worksheet("Calc")
    ws_u = wb.add_worksheet("Customers")
    ws_m = wb.add_worksheet("Fact_Monthly")
    ws_p = wb.add_worksheet("Fact_Product")
    ws_t = wb.add_worksheet("Fact_Time")
    ws_l = wb.add_worksheet("Lists")
    for w, col in [(ws_d, NAVY), (ws_n, GREY), (ws_c, TEAL), (ws_u, GREY),
                   (ws_m, GREY), (ws_p, GREY), (ws_t, GREY), (ws_l, GREY)]:
        w.set_tab_color(col)

    years, countries = f["years"], f["countries"]
    minY, maxY = years[0], years[-1]
    first_dt, last_dt = d.InvoiceDate.min(), d.InvoiceDate.max()
    last_full_month = last_dt.month if last_dt.day >= last_dt.days_in_month else last_dt.month - 1
    last_full_month = max(last_full_month, 1)

    hdr = F(bold=True, font_color="#FFFFFF", bg_color=NAVY, align="center", border=1, border_color=NAVY)

    def write_table(ws, df, fmts=None, widths=None):
        for j, col in enumerate(df.columns):
            ws.write(0, j, col, hdr)
        fmts = fmts or {}
        for j, col in enumerate(df.columns):
            fm = fmts.get(col)
            vals = df[col].tolist()
            for i, v in enumerate(vals, start=1):
                if isinstance(v, pd.Timestamp):
                    ws.write_datetime(i, j, v.to_pydatetime(), fm)
                elif isinstance(v, str):
                    ws.write_string(i, j, v, fm)
                else:
                    ws.write_number(i, j, float(v), fm)
        for j, w_ in enumerate(widths or []):
            ws.set_column(j, j, w_)
        ws.freeze_panes(1, 0)

    # ---------------------------------------------------------------- facts --
    fm_df = f["monthly"].rename(columns={"gross": "Gross", "ret": "Returns", "orders": "Orders", "units": "Units"})
    write_table(ws_m, fm_df, {"Month": F(num_format="mmm yyyy"), "Gross": F(num_format="#,##0.00"),
                              "Returns": F(num_format="#,##0.00"), "Orders": F(num_format="#,##0"),
                              "Units": F(num_format="#,##0")}, [11, 7, 10, 20, 14, 14, 10, 12])
    nFM = len(fm_df)
    fp_df = f["product"].rename(columns={"StockCode": "Code", "net": "Net revenue", "units": "Units"})
    write_table(ws_p, fp_df, {"Net revenue": F(num_format="#,##0.00"), "Units": F(num_format="#,##0")},
                [12, 38, 7, 20, 14, 10])
    nFP = len(fp_df)
    ft_df = f["time"].rename(columns={"gross": "Gross sales"})
    write_table(ws_t, ft_df, {"Gross sales": F(num_format="#,##0.00")}, [7, 20, 9, 7, 14])
    nFT = len(ft_df)

    def rng(sheet, col, n):
        return f"{sheet}!${col}$2:${col}${n + 1}"

    FM = {c: rng("Fact_Monthly", c, nFM) for c in "ABCDEFGH"}
    FP = {c: rng("Fact_Product", c, nFP) for c in "ABCDEF"}
    FT = {c: rng("Fact_Time", c, nFT) for c in "ABCDE"}

    # ------------------------------------------------------------- customers --
    cust = f["customers"]
    nU = len(cust)
    nY = len(years)
    cols = ["Customer ID", "Country", "First order", "Last order", "Orders", "Net revenue"]
    cols += [f"Rev {y}" for y in years] + [f"Orders {y}" for y in years]
    cdf = cust.copy()
    cdf.columns = cols
    base = len(cols)                      # first formula column index
    L = {k: xl_col_to_name(i) for i, k in enumerate(["id", "cty", "first", "last", "orders", "net"])}
    revc0, ordc0 = 6, 6 + nY
    L.update({"rev0": xl_col_to_name(revc0), "revN": xl_col_to_name(revc0 + nY - 1),
              "ord0": xl_col_to_name(ordc0), "ordN": xl_col_to_name(ordc0 + nY - 1)})
    for i, k in enumerate(["recency", "R", "F", "seg", "selrev", "selord", "key"]):
        L[k] = xl_col_to_name(base + i)
    write_table(ws_u, cdf, {"First order": F(num_format="dd mmm yyyy"), "Last order": F(num_format="dd mmm yyyy"),
                            "Net revenue": F(num_format="#,##0.00"),
                            **{f"Rev {y}": F(num_format="#,##0.00") for y in years}},
                [11, 20, 12, 12, 8, 13] + [12] * nY + [10] * nY)
    for i, (k, t) in enumerate([("recency", "Recency (days)"), ("R", "R score"), ("F", "F score"),
                                ("seg", "Segment"), ("selrev", "Net rev (selected yr)"),
                                ("selord", "Orders (selected yr)"), ("key", "Rank key (helper)")]):
        ws_u.write(0, base + i, t, hdr)
    ws_u.set_column(base, base + 6, 16)
    ws_u.set_column(base + 3, base + 3, 18)
    ws_u.autofilter(0, 0, nU, base + 6)

    def UC(k):                              # absolute Customers range
        return f"Customers!${L[k]}$2:${L[k]}${nU + 1}"

    # Calc sheet geometry -------------------------------------------------------
    N = {}                                  # name -> absolute address on Calc

    def A(name):
        return N[name]

    # fixed dashboard filter cells
    D_CTY, D_YEAR, D_METRIC = "Dashboard!$B$10", "Dashboard!$B$14", "Dashboard!$B$18"

    # ---- customer formulas (need Calc names -> define addresses first) --------
    # Calc layout (col C values) -- addresses decided here, written further below
    calc_rows = {
        "yearSel": 5, "yearLo": 6, "yearHi": 7, "ctyCrit": 8, "ctyLabel": 9, "yrLabel": 10,
        "yearIdx": 11, "metricIdx": 12, "period": 13,
        "snapshot": 16, "lastFull": 17, "y1": 18, "y0": 19, "minY": 20,
        "p25": 23, "p50": 24, "p75": 25, "f1": 28, "f2": 29, "f3": 30,
        "gross": 34, "ret": 35, "net": 36, "orders": 37, "units": 38, "aov": 39, "rr": 40,
        "active": 41, "identNet": 42, "rpc": 43, "ipo": 44,
        "net1": 47, "net0": 48, "ord1": 49, "ord0": 50, "uni1": 51, "uni0": 52,
        "gro1": 53, "gro0": 54, "yoyNet": 55, "yoyOrd": 56, "yoyUni": 57, "yoyAov": 58,
        "ukShare": 61,
    }
    for k, r in calc_rows.items():
        N[k] = f"Calc!$C${r}"

    n_ = len(cdf)
    for i in range(n_):
        r = i + 2
        ws_u.write_formula(i + 1, base + 0, f"={A('snapshot')}-{L['last']}{r}", F(num_format="0"))
        ws_u.write_formula(i + 1, base + 1,
                           f"=IF({L['recency']}{r}<={A('p25')},4,IF({L['recency']}{r}<={A('p50')},3,"
                           f"IF({L['recency']}{r}<={A('p75')},2,1)))", F(num_format="0"))
        ws_u.write_formula(i + 1, base + 2,
                           f"=1+({L['orders']}{r}>={A('f1')})+({L['orders']}{r}>={A('f2')})+({L['orders']}{r}>={A('f3')})",
                           F(num_format="0"))
        R_, F_ = f"{L['R']}{r}", f"{L['F']}{r}"
        ws_u.write_formula(i + 1, base + 3,
                           f'=IF(AND({R_}=4,{F_}>=3),"Champions",IF(AND({R_}=3,{F_}>=3),"Loyal",'
                           f'IF(AND({R_}>=3,{F_}<=2),"Recent / New",IF(AND({R_}<=2,{F_}>=3),"At Risk",'
                           f'"Hibernating / Lost"))))')
        ws_u.write_formula(i + 1, base + 4,
                           f"=IF({A('yearIdx')}=0,{L['net']}{r},INDEX({L['rev0']}{r}:{L['revN']}{r},{A('yearIdx')}))",
                           F(num_format="#,##0.00"))
        ws_u.write_formula(i + 1, base + 5,
                           f"=IF({A('yearIdx')}=0,{L['orders']}{r},INDEX({L['ord0']}{r}:{L['ordN']}{r},{A('yearIdx')}))",
                           F(num_format="0"))
        ws_u.write_formula(i + 1, base + 6,
                           f'=IF(OR({A("ctyCrit")}="*",{L["cty"]}{r}={A("ctyCrit")}),{L["selrev"]}{r},0)+ROW()*0.000000001',
                           F(num_format="#,##0.00"))

    # ----------------------------------------------------------------- Lists --
    ws_l.write(0, 0, "Countries", hdr)
    ws_l.write(1, 0, "All Countries")
    for i, c in enumerate(countries):
        ws_l.write(2 + i, 0, c)
    nCty = len(countries)
    ws_l.write(0, 2, "Years", hdr)
    ws_l.write(1, 2, "All Years")
    for i, y in enumerate(years):
        ws_l.write_number(2 + i, 2, int(y))
    ws_l.write(0, 4, "Trend metric", hdr)
    for i, m in enumerate(METRICS):
        ws_l.write(1 + i, 4, m)
    ws_l.write(0, 6, "Product code", hdr)
    ws_l.write(0, 7, "Product description", hdr)
    for i, (code, desc) in enumerate(f["plist"]):
        ws_l.write_string(1 + i, 6, code)
        ws_l.write_string(1 + i, 7, desc)
    ws_l.set_column(0, 0, 22)
    ws_l.set_column(2, 2, 12)
    ws_l.set_column(4, 4, 18)
    ws_l.set_column(6, 6, 14)
    ws_l.set_column(7, 7, 40)
    nProd = len(f["plist"])

    # ------------------------------------------------------------------ Calc --
    ws_c.set_column("A:A", 2)
    ws_c.set_column("B:B", 52)
    ws_c.set_column("C:J", 17)
    ws_c.hide_gridlines(2)
    h1 = F(bold=True, font_size=12, font_color=NAVY)
    h2 = F(bold=True, font_color="#FFFFFF", bg_color=NAVY)
    lab = F()
    thead = F(bold=True, font_color="#FFFFFF", bg_color=TEAL, align="center")
    inp = F(font_color="#0000FF", bg_color="#FFFFE0", border=1, border_color=LINE)
    num0, num2 = F(num_format="#,##0"), F(num_format="#,##0.00")
    pct1 = F(num_format="0.0%")
    datef = F(num_format="mmm yyyy", align="left")

    def put(row, text, formula, fmt=None):
        ws_c.write(row - 1, 1, text, lab)
        if isinstance(formula, str) and formula.startswith("="):
            ws_c.write_formula(row - 1, 2, formula, fmt or lab)
        else:
            ws_c.write(row - 1, 2, formula, fmt or lab)

    ws_c.write(0, 1, "CALC ENGINE - every dashboard number is a live formula on the Fact_* / Customers sheets. "
                     "Blue cells = editable assumptions.", h1)
    ws_c.merge_range(3, 1, 3, 2, "FILTER STATE (driven by the Dashboard drop-downs)", h2)
    put(5, "Year selected (0 = all years)", f"=IFERROR(VALUE({D_YEAR}),0)")
    put(6, "Year filter - low", f"={A('yearSel')}")
    put(7, "Year filter - high", f"=IF({A('yearSel')}=0,9999,{A('yearSel')})")
    put(8, 'Country criterion ("*" = all)', f'=IF({D_CTY}="All Countries","*",{D_CTY})')
    put(9, "Country label", f'=IF({D_CTY}="All Countries","All countries",{D_CTY})')
    put(10, "Year label", f'=IF({A("yearSel")}=0,"All years",TEXT({A("yearSel")},"0"))')
    put(11, "Year index into Customers year-columns", f"=IF({A('yearSel')}=0,0,{A('yearSel')}-{minY - 1})")
    put(12, "Trend metric index", f"=IFERROR(MATCH({D_METRIC},Lists!$E$2:$E${len(METRICS) + 1},0),1)")
    lbl_all = f"{first_dt:%b %Y} - {last_dt:%b %Y}"
    lbl_min = (f"{first_dt:%b} - Dec {minY} (partial)" if first_dt.month != 1 or first_dt.day != 1
               else f"Jan - Dec {minY}")
    lbl_max = f"Jan - {last_dt:%b} {maxY} (partial)"
    put(13, "Period label (card sub-text)",
        f'=IF({A("yearSel")}=0,"{lbl_all}",IF({A("yearSel")}={minY},"{lbl_min}",'
        f'IF({A("yearSel")}={maxY},"{lbl_max}","Jan - Dec "&{A("yearSel")})))')

    ws_c.merge_range(14, 1, 14, 2, "ASSUMPTIONS", h2)
    put(16, "Snapshot date for recency (day after last order)", f"=MAX({UC('last')})+1", F(num_format="dd mmm yyyy", align="left"))
    put(17, f"Last COMPLETE month of {maxY} (for YoY; hard-coded from the data)", last_full_month, inp)
    put(18, "Latest year in data", f"=MAX({FM['B']})")
    put(19, "Previous year", f"={A('y1')}-1")
    put(20, "Earliest year in data", f"=MIN({FM['B']})")
    ws_c.merge_range(21, 1, 21, 2, "RFM recency cut-offs (days since last order, quartiles of all customers)", h2)
    put(23, "Score 4 if recency <= (25th pct)", f"=PERCENTILE({UC('recency')},0.25)", num0)
    put(24, "Score 3 if recency <= (50th pct)", f"=PERCENTILE({UC('recency')},0.5)", num0)
    put(25, "Score 2 if recency <= (75th pct)", f"=PERCENTILE({UC('recency')},0.75)", num0)
    ws_c.merge_range(26, 1, 26, 2, "RFM frequency cut-offs (number of orders) - editable", h2)
    put(28, "F score +1 at >= orders", 2, inp)
    put(29, "F score +1 again at >= orders", 4, inp)
    put(30, "F score +1 again at >= orders", 8, inp)

    ws_c.merge_range(31, 1, 31, 2, "KPI CALCULATIONS (follow Country + Year filters)", h2)
    crit = (f"{FM['D']},{A('ctyCrit')},{FM['B']},\">=\"&{A('yearLo')},{FM['B']},\"<=\"&{A('yearHi')}")
    put(34, "Gross sales (£)", f"=SUMIFS({FM['E']},{crit})", num2)
    put(35, "Returns / cancellations (£)", f"=SUMIFS({FM['F']},{crit})", num2)
    put(36, "Net revenue (£)", f"={A('gross')}-{A('ret')}", num2)
    put(37, "Orders (unique invoices)", f"=SUMIFS({FM['G']},{crit})", num0)
    put(38, "Units sold", f"=SUMIFS({FM['H']},{crit})", num0)
    put(39, "Avg order value (gross sales / orders)", f"=IFERROR({A('gross')}/{A('orders')},0)", num2)
    put(40, "Return rate (returns / gross sales)", f"=IFERROR({A('ret')}/{A('gross')},0)", pct1)
    put(41, "Active customers (with Customer ID)",
        f'=COUNTIFS({UC("cty")},{A("ctyCrit")},{UC("selrev")},">0")', num0)
    put(42, "Net revenue from identified customers (£)",
        f"=SUMIFS({UC('selrev')},{UC('cty')},{A('ctyCrit')})", num2)
    put(43, "Revenue per active customer (£)", f"=IFERROR({A('identNet')}/{A('active')},0)", num2)
    put(44, "Items (units) per order", f"=IFERROR({A('units')}/{A('orders')},0)", num2)

    ws_c.merge_range(44, 1, 44, 2, "YEAR-OVER-YEAR (latest year vs previous, Jan .. last complete month; Country filter applies)", h2)

    def yoy_sum(col, yr):
        return (f"SUMIFS({FM[col]},{FM['D']},{A('ctyCrit')},{FM['B']},{A(yr)},{FM['C']},\"<=\"&{A('lastFull')})")
    put(47, "Net revenue - latest year (£)", f"={yoy_sum('E', 'y1')}-{yoy_sum('F', 'y1')}", num2)
    put(48, "Net revenue - previous year (£)", f"={yoy_sum('E', 'y0')}-{yoy_sum('F', 'y0')}", num2)
    put(49, "Orders - latest year", f"={yoy_sum('G', 'y1')}", num0)
    put(50, "Orders - previous year", f"={yoy_sum('G', 'y0')}", num0)
    put(51, "Units - latest year", f"={yoy_sum('H', 'y1')}", num0)
    put(52, "Units - previous year", f"={yoy_sum('H', 'y0')}", num0)
    put(53, "Gross sales - latest year (£)", f"={yoy_sum('E', 'y1')}", num2)
    put(54, "Gross sales - previous year (£)", f"={yoy_sum('E', 'y0')}", num2)
    put(55, "Net revenue YoY", f"=IFERROR({A('net1')}/{A('net0')}-1,0)", pct1)
    put(56, "Orders YoY", f"=IFERROR({A('ord1')}/{A('ord0')}-1,0)", pct1)
    put(57, "Units YoY", f"=IFERROR({A('uni1')}/{A('uni0')}-1,0)", pct1)
    put(58, "Avg order value YoY", f"=IFERROR(({A('gro1')}/{A('ord1')})/({A('gro0')}/{A('ord0')})-1,0)", pct1)

    # card sub-text helper ------------------------------------------------------
    mon_txt = f'TEXT(DATE(2000,{A("lastFull")},1),"mmm")'

    def sub(yoy):
        return (f'=IF({A("yearSel")}={A("y1")},TEXT({A(yoy)},"+0.0%;-0.0%")&" vs "&{A("y0")}&" (Jan-"&{mon_txt}&")",'
                f'{A("period")})')
    put(63, "Net revenue sub-text", sub("yoyNet")); N["sub_net"] = "Calc!$C$63"
    put(64, "Orders sub-text", sub("yoyOrd")); N["sub_ord"] = "Calc!$C$64"
    put(65, "AOV sub-text", sub("yoyAov")); N["sub_aov"] = "Calc!$C$65"
    put(66, "Units sub-text", sub("yoyUni")); N["sub_uni"] = "Calc!$C$66"
    put(67, "Return sub-text", f'="Returns value: £"&TEXT({A("ret")},"#,##0")'); N["sub_ret"] = "Calc!$C$67"

    # ------------------ monthly series (follows Country, ignores Year) -----------
    nMon = f["monthly"].Month.nunique()
    r0 = 70
    ws_c.merge_range(r0 - 2, 1, r0 - 2, 9, "MONTHLY SERIES (follows Country filter; ignores Year filter)", h2)
    for j, t in enumerate(["Month", "Gross sales", "Returns", "Net revenue", "Orders", "Units",
                           "Avg order value", "Label", "Selected metric"]):
        ws_c.write(r0 - 1, 1 + j, t, thead)
    for i in range(nMon):
        r = r0 + 1 + i                       # excel row
        ri = r - 1
        ws_c.write_formula(ri, 1, f"=MIN({FM['A']})" if i == 0 else f"=EDATE(B{r - 1},1)", datef)
        mc = f"{FM['A']},$B{r},{FM['D']},{A('ctyCrit')}"
        ws_c.write_formula(ri, 2, f"=SUMIFS({FM['E']},{mc})", num2)
        ws_c.write_formula(ri, 3, f"=SUMIFS({FM['F']},{mc})", num2)
        ws_c.write_formula(ri, 4, f"=C{r}-D{r}", num2)
        ws_c.write_formula(ri, 5, f"=SUMIFS({FM['G']},{mc})", num0)
        ws_c.write_formula(ri, 6, f"=SUMIFS({FM['H']},{mc})", num0)
        ws_c.write_formula(ri, 7, f"=IFERROR(C{r}/F{r},0)", num2)
        ws_c.write_formula(ri, 8, f'=TEXT(B{r},"mmm yy")')
        ws_c.write_formula(ri, 9, f"=CHOOSE({A('metricIdx')},E{r},C{r},D{r},F{r},G{r},H{r})", num2)
    mon_first, mon_last = r0 + 1, r0 + nMon

    # ------------------ calendar-month comparison y0 vs y1 ------------------------
    r1 = mon_last + 4
    ws_c.merge_range(r1 - 2, 1, r1 - 2, 4, "NET REVENUE BY CALENDAR MONTH, previous vs latest year (Country filter applies)", h2)
    ws_c.write(r1 - 1, 1, "Month", thead)
    ws_c.write(r1 - 1, 2, "Month #", thead)
    ws_c.write_formula(r1 - 1, 3, f"={A('y0')}", F(bold=True, font_color="#FFFFFF", bg_color=TEAL, align="center", num_format="0"))
    ws_c.write_formula(r1 - 1, 4, f"={A('y1')}", F(bold=True, font_color="#FFFFFF", bg_color=TEAL, align="center", num_format="0"))
    for m in range(1, 13):
        r = r1 + m
        ws_c.write_formula(r - 1, 1, f'=TEXT(DATE(2000,C{r},1),"mmm")')
        ws_c.write_number(r - 1, 2, m)
        for col, yc in ((3, "$D$"), (4, "$E$")):
            yref = f"{yc}{r1}"
            ws_c.write_formula(
                r - 1, col,
                f"=SUMIFS({FM['E']},{FM['D']},{A('ctyCrit')},{FM['B']},{yref},{FM['C']},$C{r})"
                f"-SUMIFS({FM['F']},{FM['D']},{A('ctyCrit')},{FM['B']},{yref},{FM['C']},$C{r})", num2)
    cal_first, cal_last = r1 + 1, r1 + 12

    # ------------------ international markets -------------------------------------
    r2 = cal_last + 4
    ws_c.merge_range(r2 - 2, 1, r2 - 2, 4, "MARKETS (follow Year filter; Country filter ignored)", h2)
    for j, t in enumerate(["Country", "Net revenue", "Rank key"]):
        ws_c.write(r2 - 1, 1 + j, t, thead)
    ycrit = f"{FM['B']},\">=\"&{A('yearLo')},{FM['B']},\"<=\"&{A('yearHi')}"
    for i in range(nCty):
        r = r2 + 1 + i
        ws_c.write_formula(r - 1, 1, f"=Lists!$A${3 + i}")
        ws_c.write_formula(r - 1, 2, f"=SUMIFS({FM['E']},{FM['D']},B{r},{ycrit})-SUMIFS({FM['F']},{FM['D']},B{r},{ycrit})", num2)
        ws_c.write_formula(r - 1, 3, f'=IF(B{r}="{HOME_COUNTRY}",-1000000000,C{r})+ROW()*0.000000001', num2)
    mk_a, mk_b = r2 + 1, r2 + nCty
    N["ukShare"] = f"Calc!$C${mk_b + 2}"
    ws_c.write(mk_b + 1, 1, f"{HOME_COUNTRY} share of net revenue", lab)
    ws_c.write_formula(mk_b + 1, 2,
                       f'=IFERROR(SUMIFS($C${mk_a}:$C${mk_b},$B${mk_a}:$B${mk_b},"{HOME_COUNTRY}")/SUM($C${mk_a}:$C${mk_b}),0)', pct1)
    r3 = mk_b + 5
    ws_c.merge_range(r3 - 2, 1, r3 - 2, 3, "TOP 10 MARKETS (rows run rank 10 -> 1 so the bar chart shows #1 on top)", h2)
    for j, t in enumerate(["Country", "Net revenue", "Rank key"]):
        ws_c.write(r3 - 1, 1 + j, t, thead)
    for j in range(10):
        r = r3 + 1 + j
        k = 10 - j
        ws_c.write_formula(r - 1, 3, f"=LARGE($D${mk_a}:$D${mk_b},{k})", num2)
        ws_c.write_formula(r - 1, 1, f"=INDEX($B${mk_a}:$B${mk_b},MATCH(D{r},$D${mk_a}:$D${mk_b},0))")
        ws_c.write_formula(r - 1, 2, f"=INDEX($C${mk_a}:$C${mk_b},MATCH(D{r},$D${mk_a}:$D${mk_b},0))", num2)
    tm_a, tm_b = r3 + 1, r3 + 10

    # ------------------ segments -----------------------------------------------------
    r4 = tm_b + 4
    ws_c.merge_range(r4 - 2, 1, r4 - 2, 6, "CUSTOMER SEGMENTS (Country filter only; all-time snapshot)", h2)
    for j, t in enumerate(["Segment", "Customers", "% customers", "Net revenue", "% revenue", "Avg rev / customer"]):
        ws_c.write(r4 - 1, 1 + j, t, thead)
    for i, s in enumerate(SEGMENTS):
        r = r4 + 1 + i
        ws_c.write(r - 1, 1, s)
        ws_c.write_formula(r - 1, 2, f"=COUNTIFS({UC('seg')},B{r},{UC('cty')},{A('ctyCrit')})", num0)
        ws_c.write_formula(r - 1, 3, f"=IFERROR(C{r}/C${r4 + 6},0)", pct1)
        ws_c.write_formula(r - 1, 4, f"=SUMIFS({UC('net')},{UC('seg')},B{r},{UC('cty')},{A('ctyCrit')})", num2)
        ws_c.write_formula(r - 1, 5, f"=IFERROR(E{r}/E${r4 + 6},0)", pct1)
        ws_c.write_formula(r - 1, 6, f"=IFERROR(E{r}/C{r},0)", num2)
    rt = r4 + 6
    ws_c.write(rt - 1, 1, "Total", F(bold=True))
    ws_c.write_formula(rt - 1, 2, f"=SUM(C{r4 + 1}:C{r4 + 5})", num0)
    ws_c.write_formula(rt - 1, 3, f"=SUM(D{r4 + 1}:D{r4 + 5})", pct1)
    ws_c.write_formula(rt - 1, 4, f"=SUM(E{r4 + 1}:E{r4 + 5})", num2)
    ws_c.write_formula(rt - 1, 5, f"=SUM(F{r4 + 1}:F{r4 + 5})", pct1)
    ws_c.write_formula(rt - 1, 6, f"=IFERROR(E{rt}/C{rt},0)", num2)
    seg_a, seg_b = r4 + 1, r4 + 5

    # ------------------ top 10 customers ----------------------------------------------
    r5 = rt + 4
    ws_c.merge_range(r5 - 2, 1, r5 - 2, 3, "TOP 10 CUSTOMERS (Country + Year filters)", h2)
    for j, t in enumerate(["Rank", "Rank key", "Row in Customers"]):
        ws_c.write(r5 - 1, 1 + j, t, thead)
    for k in range(1, 11):
        r = r5 + k
        ws_c.write_number(r - 1, 1, k)
        ws_c.write_formula(r - 1, 2, f"=LARGE({UC('key')},B{r})", num2)
        ws_c.write_formula(r - 1, 3, f"=MATCH(C{r},{UC('key')},0)")
    tc_a = r5 + 1

    # ------------------ products ---------------------------------------------------------
    r6 = tc_a + 14
    ws_c.merge_range(r6 - 2, 1, r6 - 2, 4, f"PRODUCT TABLE (top {nProd} products; follows Country + Year filters)", h2)
    for j, t in enumerate(["Code", "Description", "Net revenue", "Rank key"]):
        ws_c.write(r6 - 1, 1 + j, t, thead)
    pcrit = f"{FP['D']},{A('ctyCrit')},{FP['C']},\">=\"&{A('yearLo')},{FP['C']},\"<=\"&{A('yearHi')}"
    for i in range(nProd):
        r = r6 + 1 + i
        ws_c.write_formula(r - 1, 1, f"=Lists!$G${2 + i}")
        ws_c.write_formula(r - 1, 2, f"=Lists!$H${2 + i}")
        ws_c.write_formula(r - 1, 3, f"=SUMIFS({FP['E']},{FP['A']},B{r},{pcrit})", num2)
        ws_c.write_formula(r - 1, 4, f"=D{r}+ROW()*0.000000001", num2)
    pr_a, pr_b = r6 + 1, r6 + nProd
    r7 = pr_b + 4
    ws_c.merge_range(r7 - 2, 1, r7 - 2, 4, "TOP 10 PRODUCTS (rows run rank 10 -> 1 so the bar chart shows #1 on top)", h2)
    for j, t in enumerate(["Product", "Net revenue", "Rank key"]):
        ws_c.write(r7 - 1, 1 + j, t, thead)
    for j in range(10):
        r = r7 + 1 + j
        k = 10 - j
        ws_c.write_formula(r - 1, 3, f"=LARGE($E${pr_a}:$E${pr_b},{k})", num2)
        ws_c.write_formula(r - 1, 1, f"=LEFT(INDEX($C${pr_a}:$C${pr_b},MATCH(D{r},$E${pr_a}:$E${pr_b},0)),30)")
        ws_c.write_formula(r - 1, 2, f"=INDEX($D${pr_a}:$D${pr_b},MATCH(D{r},$E${pr_a}:$E${pr_b},0))", num2)
    tp_a, tp_b = r7 + 1, r7 + 10

    # ================================================================== DASHBOARD
    # Layout: left sidebar (B:C) with filters  |  main area (E:T) with KPI cards on top
    ws = ws_d
    ws.hide_gridlines(2)
    ws.set_zoom(90)
    ws.set_column("A:A", 2, F(bg_color=CREAM))
    ws.set_column("B:C", 14, F(bg_color=FOREST))
    ws.set_column("D:D", 2, F(bg_color=CREAM))
    ws.set_column("E:T", 10.5, F(bg_color=CREAM))
    ws.set_column("U:U", 2, F(bg_color=CREAM))
    ws.set_column("V:XFD", None, F(bg_color=CREAM))
    LASTROW = 81
    heights = {1: 37.5, 2: 18, 3: 3, 4: 16.5, 5: 21.75, 6: 21.75, 7: 16.5, 8: 7.5, 9: 19.5, 26: 7.5,
               27: 19.5, 46: 7.5, 47: 19.5, 48: 18, 59: 7.5, 60: 18, 67: 26, 68: 7.5, 69: 19.5, 70: 19.5,
               78: 14, 79: 7.5, 80: 16, 81: 7.5}
    for r_ in range(71, 78):
        heights[r_] = 19.5

    def rh(r_):
        return heights.get(r_, 15)
    for r_ in range(1, LASTROW + 1):
        ws.set_row(r_ - 1, rh(r_))

    def px(r1_, r2_):
        return int(sum(rh(x) for x in range(r1_, r2_ + 1)) * 4 / 3)
    COLPX = 78

    # ---- sidebar ---------------------------------------------------------------
    sb = dict(bg_color=FOREST)
    ws.merge_range("B1:C1", "ONLINE RETAIL II", F(bold=True, font_size=13, font_color=CREAM, align="center", **sb))
    ws.merge_range("B2:C2", "Sales Performance", F(font_size=9, font_color=MOSS, align="center", **sb))
    ws.merge_range("B4:C4", "◆  FILTERS", F(bold=True, font_size=9, font_color=GOLD, align="left", indent=1, **sb))
    ws.merge_range("B5:C7", "Use the drop-downs below.  Every KPI, chart and table updates instantly.",
                   F(italic=True, font_size=8, font_color=MOSS, text_wrap=True, valign="top", indent=1, **sb))
    slabel = F(bold=True, font_size=8, font_color=MOSS, align="left", indent=1, **sb)
    sfilter = F(bold=True, font_size=10, font_color=FOREST, bg_color=FILTER, border=1, border_color=GOLD,
                align="center", text_wrap=True)
    ws.merge_range("B9:C9", "COUNTRY", slabel)
    ws.merge_range("B10:C11", "All Countries", sfilter)
    ws.merge_range("B13:C13", "YEAR", slabel)
    ws.merge_range("B14:C15", "All Years", sfilter)
    ws.merge_range("B17:C17", "TREND METRIC", slabel)
    ws.merge_range("B18:C19", "Net Revenue", sfilter)
    ws.data_validation("B10", {"validate": "list", "source": f"=Lists!$A$2:$A${nCty + 2}",
                               "input_title": "Country", "input_message": "Pick a country or All Countries"})
    ws.data_validation("B14", {"validate": "list", "source": f"=Lists!$C$2:$C${nY + 2}"})
    ws.data_validation("B18", {"validate": "list", "source": f"=Lists!$E$2:$E${len(METRICS) + 1}"})

    ws.merge_range("B27:C27", "◆  RFM SEGMENTS", F(bold=True, font_size=9, font_color=GOLD, align="left", indent=1, **sb))
    for i_, (sname, scol) in enumerate(zip(SEGMENTS, SEG_COLORS)):
        ws.write(27 + i_, 1, "■", F(font_size=12, font_color=scol, align="right", **sb))
        ws.write(27 + i_, 2, sname, F(font_size=8, font_color=CREAM, align="left", **sb))
    ws.merge_range("B34:C34", "◆  DATA", F(bold=True, font_size=9, font_color=GOLD, align="left", indent=1, **sb))
    facts = [("PERIOD", f"{first_dt:%d %b %Y} - {last_dt:%d %b %Y}"),
             ("ROWS ANALYSED", f"{log['clean_rows']:,}"),
             ("CURRENCY", "GBP (£)"),
             ("NET REVENUE", "sales minus cancellations")]
    rr_ = 35
    for k_, v_ in facts:
        ws.merge_range(rr_ - 1, 1, rr_ - 1, 2, k_, F(font_size=7, font_color=MOSS, align="left", indent=1, **sb))
        ws.merge_range(rr_, 1, rr_, 2, v_, F(bold=True, font_size=8, font_color=CREAM, align="left", indent=1, **sb))
        rr_ += 2

    # ---- header ----------------------------------------------------------------
    ws.merge_range("E1:T1", "Online Retail II  |  Sales Performance Dashboard",
                   F(bold=True, font_size=20, font_color=FOREST, bg_color=CREAM))
    ws.merge_range("E2:T2", "UK online gift retailer  |  Dec 2009 - Dec 2011  |  Net revenue = sales minus cancellations  |  "
                            "Filters are in the green sidebar on the left".replace(
                                "Dec 2009 - Dec 2011", f"{first_dt:%b %Y} - {last_dt:%b %Y}"),
                   F(font_size=9, font_color=MUTED, bg_color=CREAM))
    ws.merge_range("E3:T3", "", F(bg_color=GOLD))

    # ---- KPI cards -------------------------------------------------------------
    GBP_M = '[>=1000000]\\£#,##0.00,,"M";[>=1000]\\£#,##0.0,"K";\\£#,##0'
    cards = [
        ("NET REVENUE", f"={A('net')}", GBP_M, f"={A('sub_net')}", LEAF),
        ("ORDERS", f"={A('orders')}", "#,##0", f"={A('sub_ord')}", SAGE),
        ("AVG ORDER VALUE", f"={A('aov')}", "\\£#,##0.00", f"={A('sub_aov')}", GOLD),
        ("UNITS SOLD", f"={A('units')}", '[>=1000000]#,##0.00,,"M";#,##0', f"={A('sub_uni')}", TERRA),
        ("ACTIVE CUSTOMERS", f"={A('active')}", "#,##0", '="Buyers with a Customer ID"', FOREST),
        ("RETURN RATE", f"={A('rr')}", "0.0%", f"={A('sub_ret')}", RUST),
        ("REVENUE / CUSTOMER", f"={A('rpc')}", "\\£#,##0", '="Identified customers only"', LEAF),
        ("ITEMS / ORDER", f"={A('ipo')}", "#,##0.0", '="Units sold / orders"', GOLD),
    ]
    for i, (title_, val, nf, subf, colr) in enumerate(cards):
        c0 = 4 + 2 * i
        edge = dict(left=1, right=1, left_color=SAND, right_color=SAND)
        ws.merge_range(3, c0, 3, c0 + 1, title_,
                       F(bold=True, font_size=8, font_color=MUTED, bg_color=CARD, top=5, top_color=colr,
                         align="left", indent=1, **edge))
        vf = F(bold=True, font_size=19, font_color=FOREST, bg_color=CARD, num_format=nf, align="left", indent=1, **edge)
        ws.merge_range(4, c0, 5, c0 + 1, "", vf)
        ws.write_formula(4, c0, val, vf)
        sf = F(font_size=8, font_color=MUTED, bg_color=CARD, align="left", indent=1, bottom=1, bottom_color=SAND, **edge)
        ws.merge_range(6, c0, 6, c0 + 1, "", sf)
        ws.write_formula(6, c0, subf, sf)

    sec = F(bold=True, font_size=11, font_color=FOREST, bg_color=CREAM, bottom=2, bottom_color=GOLD)

    def section(rng_, formula):
        ws.merge_range(rng_, "", sec)
        ws.write_formula(rng_.split(":")[0], formula, sec)

    section("E9:L9", f'="Monthly trend  |  "&{D_METRIC}&"  |  "&{A("ctyLabel")}&"  (all months)"')
    section("M9:T9", f'="Net revenue by calendar month, "&{A("y0")}&" vs "&{A("y1")}&"  |  "&{A("ctyLabel")}'
                     f'&"  ({last_dt:%b %Y} partial)"')
    section("E27:L27", f'="Top 10 international markets  |  "&{A("yrLabel")}&"  |  {HOME_SHORT} = "'
                       f'&TEXT({A("ukShare")},"0%")&" of net revenue"')
    section("M27:T27", f'="Top 10 products by net revenue  |  "&{A("yrLabel")}&"  |  "&{A("ctyLabel")}')
    section("E47:J47", f'="Customer segments (RFM, all-time)  |  "&{A("ctyLabel")}')
    section("K47:T47", f'="Top 10 customers  |  "&{A("yrLabel")}&"  |  "&{A("ctyLabel")}')
    section("E69:T69", f'="When customers buy: gross sales by weekday and hour (£k)  |  "&{A("yrLabel")}&"  |  "&{A("ctyLabel")}')

    # ---- charts ----------------------------------------------------------------
    def style(ch, w, h):
        ch.set_size({"width": w, "height": h})
        ch.set_chartarea({"border": {"color": SAND}, "fill": {"color": CARD}})
        ch.set_plotarea({"fill": {"color": CARD}})
        ch.set_legend({"none": True})
        ch.set_title({"none": True})

    axfont = {"name": FONT, "size": 8, "color": INK}
    grid = {"visible": True, "line": {"color": SAND, "width": 0.75}}
    axline = {"color": SAND}
    KM = '[>=1000000]0.0,,"M";[>=1000]0,"K";0'

    def cr(r1_, c1_, r2_, c2_):
        return ["Calc", r1_ - 1, c1_, r2_ - 1, c2_]

    ch = wb.add_chart({"type": "line"})
    ch.add_series({"name": "Selected metric", "categories": cr(mon_first, 8, mon_last, 8),
                   "values": cr(mon_first, 9, mon_last, 9), "line": {"color": LEAF, "width": 2.5},
                   "marker": {"type": "circle", "size": 5, "fill": {"color": CARD}, "border": {"color": LEAF, "width": 1.5}}})
    ch.set_x_axis({"num_font": {**axfont, "rotation": -45}, "line": axline, "major_tick_mark": "none"})
    ch.set_y_axis({"num_font": axfont, "major_gridlines": grid, "line": {"none": True}, "num_format": KM})
    style(ch, 8 * COLPX - 4, px(10, 25) - 2)
    ws.insert_chart("E10", ch, {"x_offset": 2, "y_offset": 1, "object_position": 1})

    ch = wb.add_chart({"type": "column"})
    for col, colr in ((3, "#B8AD94"), (4, LEAF)):
        ch.add_series({"name": ["Calc", r1 - 1, col], "categories": cr(cal_first, 1, cal_last, 1),
                       "values": cr(cal_first, col, cal_last, col), "fill": {"color": colr}, "gap": 60,
                       "border": {"none": True}})
    ch.set_x_axis({"num_font": axfont, "line": axline, "major_tick_mark": "none"})
    ch.set_y_axis({"num_font": axfont, "major_gridlines": grid, "line": {"none": True}, "num_format": KM})
    style(ch, 8 * COLPX - 4, px(10, 25) - 2)
    ch.set_legend({"position": "top", "font": axfont})
    ws.insert_chart("M10", ch, {"x_offset": 2, "y_offset": 1, "object_position": 1})

    lbl_fmt = '[>=1000000]\\£0.00,,"M";\\£0,"K"'
    for anchor, (a_, b_), colr in (("E28", (tm_a, tm_b), SAGE), ("M28", (tp_a, tp_b), GOLD)):
        ch = wb.add_chart({"type": "bar"})
        ch.add_series({"categories": cr(a_, 1, b_, 1), "values": cr(a_, 2, b_, 2), "fill": {"color": colr},
                       "gap": 45, "border": {"none": True},
                       "data_labels": {"value": True, "num_format": lbl_fmt, "font": {**axfont, "color": FOREST}}})
        ch.set_x_axis({"num_font": axfont, "line": axline, "major_tick_mark": "none"})
        ch.set_y_axis({"num_font": axfont, "major_gridlines": grid, "line": {"none": True},
                       "num_format": '\\£#,##0,"K"'})
        style(ch, 8 * COLPX - 4, px(28, 45) - 2)
        ws.insert_chart(anchor, ch, {"x_offset": 2, "y_offset": 1, "object_position": 1})

    ch = wb.add_chart({"type": "column"})
    ch.add_series({"name": "Customers", "categories": cr(seg_a, 1, seg_b, 1), "values": cr(seg_a, 2, seg_b, 2),
                   "points": [{"fill": {"color": c}} for c in SEG_COLORS], "gap": 50, "border": {"none": True},
                   "data_labels": {"value": True, "num_format": "#,##0", "font": {**axfont, "color": FOREST, "bold": True}}})
    ch.set_x_axis({"num_font": axfont, "line": axline, "major_tick_mark": "none"})
    ch.set_y_axis({"num_font": axfont, "major_gridlines": grid, "line": {"none": True}, "num_format": "#,##0"})
    style(ch, 6 * COLPX - 4, px(48, 66) - 2)
    ws.insert_chart("E48", ch, {"x_offset": 2, "y_offset": 1, "object_position": 1})

    # ---- top 10 customers (K:T) ---------------------------------------------------
    th = F(bold=True, font_color=CREAM, bg_color=FOREST, align="center", border=1, border_color=FOREST)
    ws.write("K48", "#", th)
    ws.write("L48", "Customer", th)
    ws.merge_range("M48:N48", "Country", th)
    ws.merge_range("O48:P48", "Net revenue", th)
    ws.write("Q48", "Orders", th)
    ws.write("R48", "Last order", th)
    ws.merge_range("S48:T48", "Segment", th)
    for k in range(1, 11):
        row = 48 + k
        bg = CARD if k % 2 else "#F4EFE2"
        z = F(bg_color=bg, bottom=1, bottom_color=SAND, font_color=INK)
        zc = F(bg_color=bg, bottom=1, bottom_color=SAND, align="center", font_color=INK)
        zr = F(bg_color=bg, bottom=1, bottom_color=SAND, align="right", num_format="\\£#,##0", font_color=INK)
        zd = F(bg_color=bg, bottom=1, bottom_color=SAND, align="center", num_format="dd mmm yy", font_color=INK)
        key, mrow = f"Calc!$C${tc_a + k - 1}", f"Calc!$D${tc_a + k - 1}"

        def pick(col):
            return f'=IF({key}<1,"",INDEX({UC(col)},{mrow}))'
        ws.write_number(row - 1, 10, k, zc)
        ws.write_formula(row - 1, 11, pick("id"), zc)
        ws.merge_range(row - 1, 12, row - 1, 13, "", z)
        ws.write_formula(row - 1, 12, pick("cty"), z)
        ws.merge_range(row - 1, 14, row - 1, 15, "", zr)
        ws.write_formula(row - 1, 14, pick("selrev"), zr)
        ws.write_formula(row - 1, 16, pick("selord"), zc)
        ws.write_formula(row - 1, 17, pick("last"), zd)
        ws.merge_range(row - 1, 18, row - 1, 19, "", zc)
        ws.write_formula(row - 1, 18, pick("seg"), zc)
    for name, colr in zip(SEGMENTS, SEG_COLORS):
        ws.conditional_format("S49:T58", {"type": "cell", "criteria": "==", "value": f'"{name}"',
                                          "format": wb.add_format({"font_color": colr, "bold": True})})

    # ---- segment table ---------------------------------------------------------------
    ws.merge_range("K60:M60", "Segment", th)
    ws.write("N60", "Customers", th)
    ws.write("O60", "% cust.", th)
    ws.merge_range("P60:Q60", "Net revenue", th)
    ws.write("R60", "% rev.", th)
    ws.merge_range("S60:T60", "Avg rev / cust.", th)
    for i in range(6):
        row = 61 + i
        cr_ = seg_a + i if i < 6 - 1 else rt
        tot = i == 5
        bg = CARD if i % 2 == 0 else "#F4EFE2"
        kw = dict(bg_color=bg, bottom=1, bottom_color=SAND, bold=tot, font_color=INK)
        if tot:
            kw.update(bg_color="#DCE8DF", top=1, top_color=FOREST)
        fl, fc = F(**kw), F(align="center", num_format="#,##0", **kw)
        fp = F(align="center", num_format="0.0%", **kw)
        fr = F(align="right", num_format="\\£#,##0", **kw)
        ws.merge_range(row - 1, 10, row - 1, 12, "", fl)
        ws.write_formula(row - 1, 10, f"=Calc!$B${cr_}", fl)
        ws.write_formula(row - 1, 13, f"=Calc!$C${cr_}", fc)
        ws.write_formula(row - 1, 14, f"=Calc!$D${cr_}", fp)
        ws.merge_range(row - 1, 15, row - 1, 16, "", fr)
        ws.write_formula(row - 1, 15, f"=Calc!$E${cr_}", fr)
        ws.write_formula(row - 1, 17, f"=Calc!$F${cr_}", fp)
        ws.merge_range(row - 1, 18, row - 1, 19, "", fr)
        ws.write_formula(row - 1, 18, f"=Calc!$G${cr_}", fr)
    ws.merge_range("K67:T67", "Champions = bought in the last ~month & 4+ orders  |  Loyal = recent-ish, 4+ orders  |  "
                              "Recent / New = recent, 1-3 orders  |  At Risk = lapsed, 4+ orders  |  Hibernating / Lost = lapsed, 1-3 orders",
                   F(font_size=7, font_color=MUTED, bg_color=CREAM, text_wrap=True, valign="top"))

    # ---- heat-map (E70:T77) ------------------------------------------------------------
    hh = F(bold=True, font_color=CREAM, bg_color=FOREST, align="center", border=1, border_color=FOREST)
    ws.write("E70", "Day \\ Hour", hh)
    nH = HOUR_MAX - HOUR_MIN + 1
    for j in range(nH):
        ws.write_number(69, 5 + j, HOUR_MIN + j, hh)
    hm = F(num_format='[=0]"";[<10000]0.0,;#,##0,', align="center", bg_color=CARD, border=1, border_color=CREAM,
           font_color=INK)
    for i, dname in enumerate(DAYS):
        row = 71 + i
        ws.write(row - 1, 4, dname, F(bold=True, font_color=FOREST, bg_color="#F4EFE2", align="center", border=1, border_color=SAND))
        for j in range(nH):
            col = xl_col_to_name(5 + j)
            ws.write_formula(
                row - 1, 5 + j,
                f"=SUMIFS({FT['E']},{FT['C']},{i + 1},{FT['D']},{col}$70,{FT['B']},{A('ctyCrit')},"
                f"{FT['A']},\">=\"&{A('yearLo')},{FT['A']},\"<=\"&{A('yearHi')})", hm)
    ws.conditional_format(70, 5, 76, 5 + nH - 1, {"type": "3_color_scale", "min_color": "#FFFDF8",
                                                    "mid_color": "#CFE3D6", "max_color": "#5E9C76"})
    ws.merge_range("E78:T78", "Gross sales only (cancellations excluded). Values in £ thousands; blank = no sales. "
                              "Saturday trading is negligible.", F(font_size=7, font_color=MUTED, bg_color=CREAM))
    ws.merge_range("E80:T80", "Data model: Fact_Monthly / Fact_Product / Fact_Time / Customers sheets  |  Logic: Calc sheet  |  "
                              "Definitions & cleaning log: Notes sheet", F(font_size=8, font_color=MUTED, bg_color=CREAM, italic=True))

    ws.freeze_panes(0, 3)
    ws.set_landscape()
    ws.set_paper(9)
    ws.print_area(f"A1:U{LASTROW}")
    ws.fit_to_pages(1, 1)
    ws.set_margins(0.3, 0.3, 0.3, 0.3)

    # ====================================================================== NOTES
    ws = ws_n
    ws.hide_gridlines(2)
    ws.set_column("A:A", 2)
    ws.set_column("B:B", 30)
    ws.set_column("C:C", 120)
    nt = F(bold=True, font_size=14, font_color=NAVY)
    ns = F(bold=True, font_color="#FFFFFF", bg_color=NAVY)
    nl = F(bold=True, font_color=NAVY, valign="top")
    nv = F(text_wrap=True, valign="top")
    row = [0]

    def title(t):
        ws.write(row[0], 1, t, nt)
        row[0] += 2

    def section_n(t):
        ws.merge_range(row[0], 1, row[0], 2, t, ns)
        row[0] += 1

    def item(a, b):
        ws.write(row[0], 1, a, nl)
        ws.write(row[0], 2, b, nv)
        row[0] += 1
    title("Notes: how to use, definitions, data cleaning and assumptions")
    section_n("HOW TO USE")
    item("Filters", "On the Dashboard sheet change the three cream-yellow filter cells in the green left sidebar (Country, Year, Trend metric) using their drop-down arrows. "
                    "All KPIs, charts and tables recalculate instantly. No macros are needed.")
    item("What each filter drives", 'Country: every panel except the "Top 10 international markets" chart. Year: KPIs, markets, products, '
                                    "top customers, heat map. The monthly trend and year-vs-year charts always show all months. "
                                    "Segments follow Country only (all-time snapshot).")
    item("Sheets", "Dashboard (report) | Calc (formulas feeding the dashboard) | Customers (one row per customer incl. RFM formulas) | "
                   "Fact_Monthly, Fact_Product, Fact_Time (pre-aggregated data) | Lists (drop-down sources).")
    section_n("DEFINITIONS")
    for a, b in [
        ("Gross sales", "Quantity x Price on normal (non-cancelled) invoices."),
        ("Returns / cancellations", 'Absolute value of lines on invoices starting with "C" (customer cancellations / credit notes).'),
        ("Net revenue", "Gross sales minus returns. Headline revenue figure."),
        ("Orders", "Unique non-cancelled invoice numbers."),
        ("Avg order value", "Gross sales / orders."),
        ("Return rate", "Returns / gross sales (value-based). Returns are dated when processed, so monthly rates can be noisy."),
        ("Active customers", "Customers with a Customer ID and positive net revenue in the selected country/year."),
        ("RFM segments", "Recency = days from the snapshot date (Calc!C16) to last order, scored 4 (best) to 1 by quartile. "
                         "Frequency = order count scored 1-4 using editable cut-offs in Calc!C28:C30. "
                         "Champions: R=4 & F>=3 | Loyal: R=3 & F>=3 | Recent / New: R>=3 & F<=2 | At Risk: R<=2 & F>=3 | Hibernating / Lost: R<=2 & F<=2."),
        ("Customer country", "Each customer is assigned the country where they generated the most revenue."),
    ]:
        item(a, b)
    section_n("DATA CLEANING LOG (computed by build_dashboard.py)")
    item("Rows in source", f"{log['rows_source']:,}")
    item("Overlap removed", f"{log['overlap']:,} exact duplicate rows inside {OVERLAP_START} .. {OVERLAP_END} (period present in both year sheets) - kept once.")
    item("Non-merchandise removed", f"{log['non_merch']:,} rows: postage, manual, bank charges, Amazon fees, samples, gift vouchers, test codes, "
                                   "bad-debt adjustment invoices ('A').")
    item("Zero / negative price removed", f"{log['bad_price']:,} rows (stock write-offs, damaged goods, free items).")
    item("Other", f"{log['other']:,} non-cancel rows with negative quantity removed.")
    item("Clean rows analysed", f"{log['clean_rows']:,} rows, {first_dt:%d %b %Y} to {last_dt:%d %b %Y}.")
    section_n("ASSUMPTIONS & LIMITATIONS")
    item("Partial months", f"{first_dt:%b %Y} starts on {first_dt:%d %b} and {last_dt:%b %Y} ends on {last_dt:%d %b}, "
                           f"so Year = {minY} and the last point of {maxY} are partial periods.")
    item("YoY comparison", f"Card sub-texts show YoY only for Year = {maxY}, comparing Jan to the last complete month "
                           f"(Calc!C17 = {last_full_month}) of {maxY} against {maxY - 1}.")
    item("Missing Customer ID", f"{log['no_cust_rows']:,} cleaned rows ({log['no_cust_share']:.0%} of net revenue) have no Customer ID. They are "
                                "included in revenue, orders, units, products and the heat map, but excluded from customer counts, RFM and top customers.")
    item("Product universe", f"Product rankings cover the top {TOP_PRODUCTS} products by lifetime net revenue.")
    item("Home market", f"{HOME_COUNTRY} dominates revenue, so the country chart ranks international markets ({HOME_COUNTRY} excluded); its share is shown in the caption.")
    item("Heat-map hours", f"Sales outside {HOUR_MIN}:00-{HOUR_MAX}:59 are folded into the edge hours.")
    item("Refreshing", "To rebuild from a newer export run:  python build_dashboard.py --csv <file.csv> --out <dashboard.xlsx>")

    wb.close()


# --------------------------------------------------------------------------- #
def main():
    ap = argparse.ArgumentParser(description="Build the Online Retail II Excel dashboard.")
    ap.add_argument("--csv", default="online_retail_II.csv", help="path to online_retail_II.csv")
    ap.add_argument("--out", default="Online_Retail_Dashboard.xlsx", help="output workbook")
    a = ap.parse_args()
    print("Loading & cleaning ...")
    d, log = load_and_clean(a.csv)
    print(f"  {log['clean_rows']:,} clean rows")
    print("Aggregating ...")
    f = build_facts(d)
    print("Writing workbook ...")
    build_workbook(a.out, d, log, f)
    print(f"Done -> {a.out}")


if __name__ == "__main__":
    sys.exit(main())
