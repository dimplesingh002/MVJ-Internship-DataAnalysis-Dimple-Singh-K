#!/usr/bin/env python3
"""
build_dashboard_only.py  -  Online Retail II | Sales Performance Dashboard
=========================================================================
Creates ONE Excel sheet ("Dashboard") - no Notes / Calc / Customers / Fact_* / Lists sheets.

Theme  : Forest green + cream
Layout : Left sidebar (filters + legend + data facts)  |  KPI cards on top  |  charts & tables

Because there are no helper sheets, the numbers are calculated in Python and written
to the sheet as values. The "filters" are therefore chosen when you run the script:

    pip install pandas numpy xlsxwriter

    python build_dashboard_only.py --csv online_retail_II.csv
    python build_dashboard_only.py --csv online_retail_II.csv --country Germany --year 2010 --metric Orders
    python build_dashboard_only.py --csv online_retail_II.csv --out MyDashboard.xlsx

Options
    --country  "All Countries" (default) or any country name, e.g. "EIRE"
    --year     "All Years" (default) or a year present in the data, e.g. 2011
    --metric   Net Revenue (default) | Gross Sales | Returns | Orders | Units Sold | Avg Order Value

(Chart source data sits in hidden columns to the right of the dashboard on the same sheet.)
"""
import argparse
import re
import sys

import numpy as np
import pandas as pd
import xlsxwriter
from xlsxwriter.utility import xl_col_to_name

# ------------------------------------------------------------------ palette ---
FONT = "Arial"
FOREST, LEAF, SAGE, MOSS = "#1F3D2B", "#2D5A3D", "#6B9080", "#A4C3B2"
GOLD, TERRA, RUST, TAUPE = "#C9A227", "#B5651D", "#A44A3F", "#9C8F7A"
CREAM, CARD, SAND, INK, MUTED, FILTER = "#F6F1E7", "#FFFDF8", "#E3D9C2", "#2F3E34", "#7A8B7F", "#F3E3A6"
ALT = "#F4EFE2"

# ----------------------------------------------------------- configuration ---
HOME_COUNTRY, HOME_SHORT = "United Kingdom", "UK"
TOP_N = 10
HOUR_MIN, HOUR_MAX = 6, 20
OVERLAP_START, OVERLAP_END = "2010-12-01", "2010-12-10"
NON_MERCH = re.compile(
    r"^(POST|DOT|M|C2|C3|D|S|B|CRUK|PADS|BANK CHARGES|ADJUST2?|AMAZONFEE|SP\d+|TEST.*|GIFT.*|DCGS.*)$")
SEGMENTS = ["Champions", "Loyal", "Recent / New", "At Risk", "Hibernating / Lost"]
SEG_COLORS = ["#3E8E5A", SAGE, GOLD, TERRA, TAUPE]
METRICS = ["Net Revenue", "Gross Sales", "Returns", "Orders", "Units Sold", "Avg Order Value"]
DAYS = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]


# ------------------------------------------------------------ load + clean ---
def load_and_clean(path):
    d = pd.read_csv(path, dtype={"Invoice": str, "StockCode": str}, parse_dates=["InvoiceDate"])
    d["Invoice"] = d["Invoice"].str.strip().str.upper()
    d["StockCode"] = d["StockCode"].str.strip().str.upper()
    d["Description"] = d["Description"].fillna("").str.strip()
    win = (d.InvoiceDate >= OVERLAP_START) & (d.InvoiceDate < OVERLAP_END)
    d = d[~(d.duplicated(keep="first") & win)]                       # overlap between year sheets
    d = d[~(d.StockCode.str.match(NON_MERCH) | d.Invoice.str.startswith("A"))]   # non-merchandise
    d = d[d.Price > 0]                                               # zero / negative price
    cancel = d.Invoice.str.startswith("C")
    d = d[~((~cancel) & (d.Quantity < 0))].copy()                    # odd negative quantities
    d["cancel"] = d.Invoice.str.startswith("C")
    d["rev"] = d.Quantity * d.Price
    d["gross"] = np.where(~d.cancel, d.rev, 0.0)
    d["ret"] = np.where(d.cancel, -d.rev, 0.0)
    d["units"] = np.where(~d.cancel, d.Quantity, 0)
    d["Month"] = d.InvoiceDate.dt.to_period("M").dt.to_timestamp()
    d["Year"] = d.InvoiceDate.dt.year
    d["Hour"] = d.InvoiceDate.dt.hour.clip(HOUR_MIN, HOUR_MAX)
    d["Weekday"] = d.InvoiceDate.dt.dayofweek + 1
    return d


# ---------------------------------------------------------------- analysis ---
def customer_table(d):
    """One row per customer: main country, orders, last order, revenue, RFM segment."""
    c = d[d["Customer ID"].notna()].copy()
    c["cid"] = c["Customer ID"].astype(int)
    cc = c.groupby(["cid", "Country"]).rev.sum().reset_index()
    cc = cc.sort_values(["cid", "rev"], ascending=[True, False]).drop_duplicates("cid")
    t = cc.set_index("cid")[["Country"]]
    nc = c[~c.cancel]
    t["last"] = nc.groupby("cid").InvoiceDate.max().dt.normalize()
    t["last"] = t["last"].fillna(c.groupby("cid").InvoiceDate.max().dt.normalize())
    t["orders"] = nc.groupby("cid").Invoice.nunique().reindex(t.index).fillna(0).astype(int)
    t["net"] = c.groupby("cid").rev.sum()
    ry = c.pivot_table(index="cid", columns="Year", values="rev", aggfunc="sum").reindex(t.index).fillna(0)
    oy = nc.pivot_table(index="cid", columns="Year", values="Invoice", aggfunc="nunique").reindex(t.index).fillna(0)
    t["rev_by_year"] = [ry.loc[i].to_dict() for i in t.index]
    t["ord_by_year"] = [oy.loc[i].to_dict() for i in t.index]
    snapshot = t["last"].max() + pd.Timedelta(days=1)
    rec = (snapshot - t["last"]).dt.days
    p25, p50, p75 = np.percentile(rec, [25, 50, 75])
    t["R"] = np.where(rec <= p25, 4, np.where(rec <= p50, 3, np.where(rec <= p75, 2, 1)))
    t["F"] = 1 + (t.orders >= 2).astype(int) + (t.orders >= 4).astype(int) + (t.orders >= 8).astype(int)
    R, Fq = t["R"], t["F"]
    t["seg"] = np.select(
        [(R == 4) & (Fq >= 3), (R == 3) & (Fq >= 3), (R >= 3) & (Fq <= 2), (R <= 2) & (Fq >= 3)],
        ["Champions", "Loyal", "Recent / New", "At Risk"], default="Hibernating / Lost")
    return t.reset_index()


def analyse(d, country, year, metric):
    years = sorted(d.Year.unique())
    minY, maxY = years[0], years[-1]
    first_dt, last_dt = d.InvoiceDate.min(), d.InvoiceDate.max()
    last_full = last_dt.month if last_dt.day >= last_dt.days_in_month else max(last_dt.month - 1, 1)
    R = {"years": years, "minY": minY, "maxY": maxY, "first": first_dt, "last": last_dt,
         "country": country or "All Countries", "year": year, "metric": metric}
    ctry_lbl = country or "All countries"
    yr_lbl = "All years" if year is None else str(year)
    R["ctry_lbl"], R["yr_lbl"] = ctry_lbl, yr_lbl

    dc = d if country is None else d[d.Country == country]
    dy = dc if year is None else dc[dc.Year == year]

    # --- KPIs
    g, r = dy.gross.sum(), dy.ret.sum()
    orders = dy[~dy.cancel].Invoice.nunique()
    units = int(dy.units.sum())
    cust = customer_table(d)
    cc = cust if country is None else cust[cust.Country == country]
    cc = cc.copy()
    cc["selrev"] = [(rb.get(year, 0.0) if year else n) for rb, n in zip(cc.rev_by_year, cc.net)]
    cc["selord"] = [(ob.get(year, 0) if year else n) for ob, n in zip(cc.ord_by_year, cc.orders)]
    active = int((cc.selrev > 0).sum())
    ident = cc.selrev.sum()
    R.update(gross=g, ret=r, net=g - r, orders=orders, units=units,
             aov=g / orders if orders else 0, rr=r / g if g else 0, active=active,
             rpc=ident / active if active else 0, ipo=units / orders if orders else 0)

    # --- YoY (latest year, Jan..last complete month)
    def part(y):
        x = dc[(dc.Year == y) & (dc.InvoiceDate.dt.month <= last_full)]
        o = x[~x.cancel].Invoice.nunique()
        return x.gross.sum() - x.ret.sum(), o, int(x.units.sum()), x.gross.sum()
    n1, o1, u1, g1 = part(maxY)
    n0, o0, u0, g0 = part(maxY - 1)
    sf = lambda a, b: (a / b - 1) if b else 0
    yoy = {"net": sf(n1, n0), "ord": sf(o1, o0), "uni": sf(u1, u0),
           "aov": sf((g1 / o1) if o1 else 0, (g0 / o0) if o0 else 0)}
    mon = pd.Timestamp(2000, last_full, 1).strftime("%b")
    if year is None:
        period = f"{first_dt:%b %Y} - {last_dt:%b %Y}"
    elif year == minY:
        period = f"{first_dt:%b} - Dec {minY} (partial)" if first_dt.month != 1 else f"Jan - Dec {minY}"
    elif year == maxY:
        period = f"Jan - {last_dt:%b} {maxY} (partial)"
    else:
        period = f"Jan - Dec {year}"

    def sub(k):
        return f"{yoy[k]:+.1%} vs {maxY - 1} (Jan-{mon})" if year == maxY else period
    R["sub"] = {"net": sub("net"), "ord": sub("ord"), "aov": sub("aov"), "uni": sub("uni"),
                "ret": f"Returns value: £{r:,.0f}"}

    # --- monthly trend (country filter only)
    allm = pd.date_range(d.Month.min(), d.Month.max(), freq="MS")
    m = dc.groupby("Month").agg(gross=("gross", "sum"), ret=("ret", "sum"), units=("units", "sum"))
    m["orders"] = dc[~dc.cancel].groupby("Month").Invoice.nunique()
    m = m.reindex(allm).fillna(0)
    m["net"] = m.gross - m.ret
    m["aov"] = np.where(m.orders > 0, m.gross / m.orders.replace(0, np.nan), 0)
    col = {"Net Revenue": "net", "Gross Sales": "gross", "Returns": "ret", "Orders": "orders",
           "Units Sold": "units", "Avg Order Value": "aov"}[metric]
    R["trend"] = (m.index.strftime("%b %y").tolist(), m[col].fillna(0).tolist())

    # --- calendar comparison
    cal = dc.assign(mn=dc.InvoiceDate.dt.month).groupby(["Year", "mn"]).apply(
        lambda x: x.gross.sum() - x.ret.sum(), include_groups=False)
    R["cal"] = ([pd.Timestamp(2000, i, 1).strftime("%b") for i in range(1, 13)],
                [float(cal.get((maxY - 1, i), 0)) for i in range(1, 13)],
                [float(cal.get((maxY, i), 0)) for i in range(1, 13)])

    # --- markets (year filter only, home market excluded)
    dyr = d if year is None else d[d.Year == year]
    mk = (dyr.groupby("Country").rev.sum())
    R["uk_share"] = mk.get(HOME_COUNTRY, 0) / mk.sum() if mk.sum() else 0
    top = mk.drop(HOME_COUNTRY, errors="ignore").nlargest(TOP_N)[::-1]      # rank 10 -> 1
    R["markets"] = (top.index.tolist(), top.values.tolist())

    # --- products (country + year filters)
    desc = (d[d.Description != ""].groupby("StockCode").Description
            .agg(lambda s: s.value_counts().index[0]).str.title())
    pr = dy.groupby("StockCode").rev.sum().nlargest(TOP_N)[::-1]
    R["products"] = ([desc.get(c, c)[:30] for c in pr.index], pr.values.tolist())

    # --- segments (country filter only)
    sg = cc.groupby("seg").agg(n=("net", "size"), rev=("net", "sum")).reindex(SEGMENTS).fillna(0)
    R["seg"] = sg
    # --- top customers
    tc = cc.sort_values("selrev", ascending=False).head(TOP_N)
    tc = tc[tc.selrev >= 1]
    R["topc"] = tc

    # --- heat-map (gross sales, country + year)
    h = dy[~dy.cancel].groupby(["Weekday", "Hour"]).gross.sum()
    R["heat"] = [[float(h.get((wd, hr), 0)) for hr in range(HOUR_MIN, HOUR_MAX + 1)] for wd in range(1, 8)]
    R["rows"] = len(d)
    return R


# --------------------------------------------------------------- workbook ----
def build(path, R):
    wb = xlsxwriter.Workbook(path)
    cache = {}

    def F(**kw):
        base = {"font_name": FONT, "font_size": 9, "valign": "vcenter"}
        base.update(kw)
        key = tuple(sorted(base.items()))
        if key not in cache:
            cache[key] = wb.add_format(base)
        return cache[key]

    ws = wb.add_worksheet("Dashboard")
    ws.set_tab_color(FOREST)
    ws.hide_gridlines(2)
    ws.set_zoom(90)
    ws.set_column("A:A", 2, F(bg_color=CREAM))
    ws.set_column("B:C", 14, F(bg_color=FOREST))
    ws.set_column("D:D", 2, F(bg_color=CREAM))
    ws.set_column("E:T", 10.5, F(bg_color=CREAM))
    ws.set_column("U:U", 2, F(bg_color=CREAM))
    ws.set_column("V:X", 10.5, F(bg_color=CREAM))

    LASTROW = 81
    heights = {1: 37.5, 2: 18, 3: 3, 4: 16.5, 5: 21.75, 6: 21.75, 7: 16.5, 8: 7.5, 9: 19.5, 26: 7.5,
               27: 19.5, 46: 7.5, 47: 19.5, 48: 18, 59: 7.5, 60: 18, 67: 26, 68: 7.5, 69: 19.5, 70: 19.5,
               78: 14, 79: 7.5, 80: 16, 81: 7.5}
    for r_ in range(71, 78):
        heights[r_] = 19.5
    rh = lambda r_: heights.get(r_, 15)
    for r_ in range(1, LASTROW + 1):
        ws.set_row(r_ - 1, rh(r_))
    px = lambda a, b: int(sum(rh(x) for x in range(a, b + 1)) * 4 / 3)
    COLPX = 78

    # ---------------- hidden chart-source data (columns Z onward) ----------------
    C0 = 25                                    # column Z
    tl, tv = R["trend"]
    for i, (a, b) in enumerate(zip(tl, tv)):
        ws.write_string(i, C0, a)
        ws.write_number(i, C0 + 1, float(b))
    cm, c0v, c1v = R["cal"]
    ws.write_string(0, C0 + 3, str(R["maxY"] - 1))
    ws.write_string(0, C0 + 4, str(R["maxY"]))
    for i in range(12):
        ws.write_string(i + 1, C0 + 2, cm[i])
        ws.write_number(i + 1, C0 + 3, c0v[i])
        ws.write_number(i + 1, C0 + 4, c1v[i])
    mn, mv = R["markets"]
    for i, (a, b) in enumerate(zip(mn, mv)):
        ws.write_string(i, C0 + 6, a)
        ws.write_number(i, C0 + 7, float(b))
    pn, pv = R["products"]
    for i, (a, b) in enumerate(zip(pn, pv)):
        ws.write_string(i, C0 + 9, a)
        ws.write_number(i, C0 + 10, float(b))
    for i, s in enumerate(SEGMENTS):
        ws.write_string(i, C0 + 12, s)
        ws.write_number(i, C0 + 13, float(R["seg"].n.iloc[i]))
    ws.set_column(C0, C0 + 14, 12, None, {"hidden": True})

    def ref(r1, c1, r2, c2):
        return ["Dashboard", r1, c1, r2, c2]

    # ---------------- sidebar ----------------
    sb = dict(bg_color=FOREST)
    gold_h = F(bold=True, font_size=9, font_color=GOLD, align="left", indent=1, **sb)
    ws.merge_range("B1:C1", "ONLINE RETAIL II", F(bold=True, font_size=13, font_color=CREAM, align="center", **sb))
    ws.merge_range("B2:C2", "Sales Performance", F(font_size=9, font_color=MOSS, align="center", **sb))
    ws.merge_range("B4:C4", "◆  FILTERS APPLIED", gold_h)
    ws.merge_range("B5:C7", "Set with --country, --year and --metric when running the script.",
                   F(italic=True, font_size=8, font_color=MOSS, text_wrap=True, valign="top", indent=1, **sb))
    slabel = F(bold=True, font_size=8, font_color=MOSS, align="left", indent=1, **sb)
    pill = F(bold=True, font_size=10, font_color=FOREST, bg_color=FILTER, border=1, border_color=GOLD,
             align="center", text_wrap=True)
    ws.merge_range("B9:C9", "COUNTRY", slabel)
    ws.merge_range("B10:C11", R["country"], pill)
    ws.merge_range("B13:C13", "YEAR", slabel)
    ws.merge_range("B14:C15", "All Years" if R["year"] is None else str(R["year"]), pill)
    ws.merge_range("B17:C17", "TREND METRIC", slabel)
    ws.merge_range("B18:C19", R["metric"], pill)

    ws.merge_range("B27:C27", "◆  RFM SEGMENTS", gold_h)
    for i, (s, col) in enumerate(zip(SEGMENTS, SEG_COLORS)):
        ws.write(27 + i, 1, "■", F(font_size=12, font_color=col, align="right", **sb))
        ws.write(27 + i, 2, s, F(font_size=8, font_color=CREAM, align="left", **sb))
    ws.merge_range("B34:C34", "◆  DATA", gold_h)
    facts = [("PERIOD", f"{R['first']:%d %b %Y} - {R['last']:%d %b %Y}"), ("ROWS ANALYSED", f"{R['rows']:,}"),
             ("CURRENCY", "GBP (£)"), ("NET REVENUE", "sales minus cancellations")]
    rr = 35
    for k_, v_ in facts:
        ws.merge_range(rr - 1, 1, rr - 1, 2, k_, F(font_size=7, font_color=MOSS, align="left", indent=1, **sb))
        ws.merge_range(rr, 1, rr, 2, v_, F(bold=True, font_size=8, font_color=CREAM, align="left", indent=1, **sb))
        rr += 2

    # ---------------- header ----------------
    ws.merge_range("E1:T1", "Online Retail II  |  Sales Performance Dashboard",
                   F(bold=True, font_size=20, font_color=FOREST, bg_color=CREAM))
    ws.merge_range("E2:T2", f"UK online gift retailer  |  {R['first']:%b %Y} - {R['last']:%b %Y}  |  "
                            f"Net revenue = sales minus cancellations  |  Filters shown in the green sidebar",
                   F(font_size=9, font_color=MUTED, bg_color=CREAM))
    ws.merge_range("E3:T3", "", F(bg_color=GOLD))

    # ---------------- KPI cards ----------------
    GBP_M = '[>=1000000]\\£#,##0.00,,"M";[>=1000]\\£#,##0.0,"K";\\£#,##0'
    S = R["sub"]
    cards = [
        ("NET REVENUE", R["net"], GBP_M, S["net"], LEAF),
        ("ORDERS", R["orders"], "#,##0", S["ord"], SAGE),
        ("AVG ORDER VALUE", R["aov"], "\\£#,##0.00", S["aov"], GOLD),
        ("UNITS SOLD", R["units"], '[>=1000000]#,##0.00,,"M";#,##0', S["uni"], TERRA),
        ("ACTIVE CUSTOMERS", R["active"], "#,##0", "Buyers with a Customer ID", FOREST),
        ("RETURN RATE", R["rr"], "0.0%", S["ret"], RUST),
        ("REVENUE / CUSTOMER", R["rpc"], "\\£#,##0", "Identified customers only", LEAF),
        ("ITEMS / ORDER", R["ipo"], "#,##0.0", "Units sold / orders", GOLD),
    ]
    for i, (t_, v_, nf, sub_, colr) in enumerate(cards):
        c0 = 4 + 2 * i
        edge = dict(left=1, right=1, left_color=SAND, right_color=SAND)
        ws.merge_range(3, c0, 3, c0 + 1, t_, F(bold=True, font_size=8, font_color=MUTED, bg_color=CARD, top=5,
                                               top_color=colr, align="left", indent=1, **edge))
        vf = F(bold=True, font_size=19, font_color=FOREST, bg_color=CARD, num_format=nf, align="left", indent=1, **edge)
        ws.merge_range(4, c0, 5, c0 + 1, "", vf)
        ws.write_number(4, c0, float(v_), vf)
        sf = F(font_size=8, font_color=MUTED, bg_color=CARD, align="left", indent=1, bottom=1, bottom_color=SAND, **edge)
        ws.merge_range(6, c0, 6, c0 + 1, sub_, sf)

    sec = F(bold=True, font_size=11, font_color=FOREST, bg_color=CREAM, bottom=2, bottom_color=GOLD)
    ctry, yr = R["ctry_lbl"], R["yr_lbl"]
    ws.merge_range("E9:L9", f"Monthly trend  |  {R['metric']}  |  {ctry}  (all months)", sec)
    ws.merge_range("M9:T9", f"Net revenue by calendar month, {R['maxY'] - 1} vs {R['maxY']}  |  {ctry}  "
                            f"({R['last']:%b %Y} partial)", sec)
    ws.merge_range("E27:L27", f"Top 10 international markets  |  {yr}  |  {HOME_SHORT} = {R['uk_share']:.0%} of net revenue", sec)
    ws.merge_range("M27:T27", f"Top 10 products by net revenue  |  {yr}  |  {ctry}", sec)
    ws.merge_range("E47:J47", f"Customer segments (RFM, all-time)  |  {ctry}", sec)
    ws.merge_range("K47:T47", f"Top 10 customers  |  {yr}  |  {ctry}", sec)
    ws.merge_range("E69:T69", f"When customers buy: gross sales by weekday and hour (£k)  |  {yr}  |  {ctry}", sec)

    # ---------------- charts ----------------
    def style(ch, w, h):
        ch.set_size({"width": w, "height": h})
        ch.set_chartarea({"border": {"color": SAND}, "fill": {"color": CARD}})
        ch.set_plotarea({"fill": {"color": CARD}})
        ch.set_legend({"none": True})
        ch.set_title({"none": True})
        ch.show_hidden_data()

    axfont = {"name": FONT, "size": 8, "color": INK}
    grid = {"visible": True, "line": {"color": SAND, "width": 0.75}}
    axline = {"color": SAND}
    KM = '[>=1000000]0.0,,"M";[>=1000]0,"K";0'
    ins = {"x_offset": 2, "y_offset": 1, "object_position": 1}

    ch = wb.add_chart({"type": "line"})
    ch.add_series({"name": "Selected metric", "categories": ref(0, C0, len(tl) - 1, C0),
                   "values": ref(0, C0 + 1, len(tl) - 1, C0 + 1), "line": {"color": LEAF, "width": 2.5},
                   "marker": {"type": "circle", "size": 5, "fill": {"color": CARD}, "border": {"color": LEAF, "width": 1.5}}})
    ch.set_x_axis({"num_font": {**axfont, "rotation": -45}, "line": axline, "major_tick_mark": "none"})
    ch.set_y_axis({"num_font": axfont, "major_gridlines": grid, "line": {"none": True}, "num_format": KM})
    style(ch, 8 * COLPX - 4, px(10, 25) - 2)
    ws.insert_chart("E10", ch, ins)

    ch = wb.add_chart({"type": "column"})
    for col, colr in ((C0 + 3, "#B8AD94"), (C0 + 4, LEAF)):
        ch.add_series({"name": ["Dashboard", 0, col], "categories": ref(1, C0 + 2, 12, C0 + 2),
                       "values": ref(1, col, 12, col), "fill": {"color": colr}, "gap": 60, "border": {"none": True}})
    ch.set_x_axis({"num_font": axfont, "line": axline, "major_tick_mark": "none"})
    ch.set_y_axis({"num_font": axfont, "major_gridlines": grid, "line": {"none": True}, "num_format": KM})
    style(ch, 8 * COLPX - 4, px(10, 25) - 2)
    ch.set_legend({"position": "top", "font": axfont})
    ws.insert_chart("M10", ch, ins)

    lbl_fmt = '[>=1000000]\\£0.00,,"M";\\£0,"K"'
    for anchor, (cc_, n_), colr in (("E28", (C0 + 6, len(mn)), SAGE), ("M28", (C0 + 9, len(pn)), GOLD)):
        ch = wb.add_chart({"type": "bar"})
        ch.add_series({"categories": ref(0, cc_, n_ - 1, cc_), "values": ref(0, cc_ + 1, n_ - 1, cc_ + 1),
                       "fill": {"color": colr}, "gap": 45, "border": {"none": True},
                       "data_labels": {"value": True, "num_format": lbl_fmt, "font": {**axfont, "color": FOREST}}})
        ch.set_x_axis({"num_font": axfont, "line": axline, "major_tick_mark": "none"})
        ch.set_y_axis({"num_font": axfont, "major_gridlines": grid, "line": {"none": True}, "num_format": '\\£#,##0,"K"'})
        style(ch, 8 * COLPX - 4, px(28, 45) - 2)
        ws.insert_chart(anchor, ch, ins)

    ch = wb.add_chart({"type": "column"})
    ch.add_series({"name": "Customers", "categories": ref(0, C0 + 12, 4, C0 + 12), "values": ref(0, C0 + 13, 4, C0 + 13),
                   "points": [{"fill": {"color": c}} for c in SEG_COLORS], "gap": 50, "border": {"none": True},
                   "data_labels": {"value": True, "num_format": "#,##0", "font": {**axfont, "color": FOREST, "bold": True}}})
    ch.set_x_axis({"num_font": axfont, "line": axline, "major_tick_mark": "none"})
    ch.set_y_axis({"num_font": axfont, "major_gridlines": grid, "line": {"none": True}, "num_format": "#,##0"})
    style(ch, 6 * COLPX - 4, px(48, 66) - 2)
    ws.insert_chart("E48", ch, ins)

    # ---------------- top-10 customers ----------------
    th = F(bold=True, font_color=CREAM, bg_color=FOREST, align="center", border=1, border_color=FOREST)
    ws.write("K48", "#", th)
    ws.write("L48", "Customer", th)
    ws.merge_range("M48:N48", "Country", th)
    ws.merge_range("O48:P48", "Net revenue", th)
    ws.write("Q48", "Orders", th)
    ws.write("R48", "Last order", th)
    ws.merge_range("S48:T48", "Segment", th)
    tc = R["topc"].reset_index(drop=True)
    for k in range(1, TOP_N + 1):
        row = 48 + k
        bg = CARD if k % 2 else ALT
        z = F(bg_color=bg, bottom=1, bottom_color=SAND, font_color=INK)
        zc = F(bg_color=bg, bottom=1, bottom_color=SAND, align="center", font_color=INK)
        zr = F(bg_color=bg, bottom=1, bottom_color=SAND, align="right", num_format="\\£#,##0", font_color=INK)
        zd = F(bg_color=bg, bottom=1, bottom_color=SAND, align="center", num_format="dd mmm yy", font_color=INK)
        ws.write_number(row - 1, 10, k, zc)
        ws.merge_range(row - 1, 12, row - 1, 13, "", z)
        ws.merge_range(row - 1, 14, row - 1, 15, "", zr)
        ws.merge_range(row - 1, 18, row - 1, 19, "", zc)
        if k <= len(tc):
            x = tc.iloc[k - 1]
            ws.write_number(row - 1, 11, int(x.cid), zc)
            ws.write_string(row - 1, 12, x.Country, z)
            ws.write_number(row - 1, 14, float(x.selrev), zr)
            ws.write_number(row - 1, 16, int(x.selord), zc)
            ws.write_datetime(row - 1, 17, x["last"].to_pydatetime(), zd)
            ws.write_string(row - 1, 18, x.seg, zc)
        else:
            for cidx in (11, 16, 17):
                ws.write_blank(row - 1, cidx, "", zc)
    for name, colr in zip(SEGMENTS, SEG_COLORS):
        ws.conditional_format("S49:T58", {"type": "cell", "criteria": "==", "value": f'"{name}"',
                                          "format": wb.add_format({"font_color": colr, "bold": True})})

    # ---------------- segment table ----------------
    ws.merge_range("K60:M60", "Segment", th)
    ws.write("N60", "Customers", th)
    ws.write("O60", "% cust.", th)
    ws.merge_range("P60:Q60", "Net revenue", th)
    ws.write("R60", "% rev.", th)
    ws.merge_range("S60:T60", "Avg rev / cust.", th)
    sg = R["seg"]
    tn, tr = sg.n.sum(), sg.rev.sum()
    rows_ = [(s, sg.n[s], sg.rev[s]) for s in SEGMENTS] + [("Total", tn, tr)]
    for i, (s, n_, rv) in enumerate(rows_):
        row = 61 + i
        tot = s == "Total"
        kw = dict(bg_color=CARD if i % 2 == 0 else ALT, bottom=1, bottom_color=SAND, bold=tot, font_color=INK)
        if tot:
            kw.update(bg_color="#DCE8DF", top=1, top_color=FOREST)
        fl = F(**kw)
        fc = F(align="center", num_format="#,##0", **kw)
        fp = F(align="center", num_format="0.0%", **kw)
        fr = F(align="right", num_format="\\£#,##0", **kw)
        ws.merge_range(row - 1, 10, row - 1, 12, s, fl)
        ws.write_number(row - 1, 13, float(n_), fc)
        ws.write_number(row - 1, 14, float(n_ / tn) if tn else 0, fp)
        ws.merge_range(row - 1, 15, row - 1, 16, "", fr)
        ws.write_number(row - 1, 15, float(rv), fr)
        ws.write_number(row - 1, 17, float(rv / tr) if tr else 0, fp)
        ws.merge_range(row - 1, 18, row - 1, 19, "", fr)
        ws.write_number(row - 1, 18, float(rv / n_) if n_ else 0, fr)
    ws.merge_range("K67:T67", "Champions = bought in the last ~month & 4+ orders  |  Loyal = recent-ish, 4+ orders  |  "
                              "Recent / New = recent, 1-3 orders  |  At Risk = lapsed, 4+ orders  |  Hibernating / Lost = lapsed, 1-3 orders",
                   F(font_size=7, font_color=MUTED, bg_color=CREAM, text_wrap=True, valign="top"))

    # ---------------- heat-map ----------------
    hh = F(bold=True, font_color=CREAM, bg_color=FOREST, align="center", border=1, border_color=FOREST)
    ws.write("E70", "Day \\ Hour", hh)
    nH = HOUR_MAX - HOUR_MIN + 1
    for j in range(nH):
        ws.write_number(69, 5 + j, HOUR_MIN + j, hh)
    hm = F(num_format='[=0]"";[<10000]0.0,;#,##0,', align="center", bg_color=CARD, border=1, border_color=CREAM, font_color=INK)
    for i, dname in enumerate(DAYS):
        row = 71 + i
        ws.write(row - 1, 4, dname, F(bold=True, font_color=FOREST, bg_color=ALT, align="center", border=1, border_color=SAND))
        for j in range(nH):
            ws.write_number(row - 1, 5 + j, R["heat"][i][j], hm)
    ws.conditional_format(70, 5, 76, 5 + nH - 1, {"type": "3_color_scale", "min_color": "#FFFDF8",
                                                    "mid_color": "#CFE3D6", "max_color": "#5E9C76"})
    ws.merge_range("E78:T78", "Gross sales only (cancellations excluded). Values in £ thousands; blank = no sales. "
                              "Saturday trading is negligible.", F(font_size=7, font_color=MUTED, bg_color=CREAM))
    ws.merge_range("E80:T80", "Static snapshot generated by build_dashboard_only.py  |  Re-run with --country / --year / --metric "
                              "to change the view", F(font_size=8, font_color=MUTED, bg_color=CREAM, italic=True))

    ws.freeze_panes(0, 3)
    ws.set_landscape()
    ws.set_paper(9)
    ws.print_area(f"A1:U{LASTROW}")
    ws.fit_to_pages(1, 1)
    ws.set_margins(0.3, 0.3, 0.3, 0.3)
    wb.close()


# ------------------------------------------------------------------- main ----
def main():
    ap = argparse.ArgumentParser(description="Single-sheet Online Retail II dashboard (forest green + cream).")
    ap.add_argument("--csv", default="online_retail_II.csv")
    ap.add_argument("--out", default="Online_Retail_Dashboard_Only.xlsx")
    ap.add_argument("--country", default="All Countries")
    ap.add_argument("--year", default="All Years")
    ap.add_argument("--metric", default="Net Revenue", help=" | ".join(METRICS))
    a = ap.parse_args()

    print("Loading & cleaning ...")
    d = load_and_clean(a.csv)
    country = None if a.country.lower() in ("all", "all countries") else a.country
    if country and country not in set(d.Country):
        sys.exit(f"Unknown country '{a.country}'. Examples: {', '.join(sorted(d.Country.unique())[:8])} ...")
    year = None if a.year.lower() in ("all", "all years") else int(a.year)
    if year and year not in set(d.Year):
        sys.exit(f"Year {year} not in data ({', '.join(map(str, sorted(d.Year.unique())))})")
    metric = {m.lower(): m for m in METRICS}.get(a.metric.lower())
    if not metric:
        sys.exit(f"Unknown metric '{a.metric}'. Choose from: {', '.join(METRICS)}")
    print("Analysing ...")
    R = analyse(d, country, year, metric)
    print("Writing workbook ...")
    build(a.out, R)
    print(f"Done -> {a.out}")


if __name__ == "__main__":
    main()
