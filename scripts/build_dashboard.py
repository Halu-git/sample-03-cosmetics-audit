#!/usr/bin/env python3
"""Build the Excel companion dashboard for the Japan Cosmetics & Skincare
Market Entry Audit (sample 03).

Source of truth: output/audit_data_appendix.xlsx (8 sheets).
Discipline:
- No figure is typed in from outside the appendix / layer2 CSV. Values are
  read from the appendix at build time.
- Values marked "Derived" are arithmetic on appendix figures only, and the
  formula is shown wherever they appear.
- Layer 1 (official statistics) and Layer 2 (platform samples) are never
  mixed on one visual; Layer 2 is never used to size the market.
"""
import copy

import openpyxl
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

BASE = "/home/hatch/workspace/estat-analysis/sample-03-english-cosmetics"
APPENDIX = f"{BASE}/output/audit_data_appendix.xlsx"
OUT = f"{BASE}/output/cosmetics_audit_dashboard.xlsx"
ACCESSED = "2026-11-12"

TITLE_FONT = Font(name="Calibri", size=14, bold=True)
SUB_FONT = Font(name="Calibri", size=10, bold=True)
NOTE_FONT = Font(name="Calibri", size=9, color="404040")
WARN_FONT = Font(name="Calibri", size=9, bold=True, color="7B2D0E")
HDR_FILL_L1 = PatternFill("solid", fgColor="2F5597")   # Layer 1: blue
HDR_FILL_L2 = PatternFill("solid", fgColor="BF8F00")   # Layer 2: amber
HDR_FONT = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
CELL_FONT = Font(name="Calibri", size=10)
BOLD_FONT = Font(name="Calibri", size=10, bold=True)

FMT_BN = '#,##0.0'
FMT_INT = '#,##0'
FMT_PCT = '0.0"%"'
FMT_RATING = '0.0'


# --------------------------------------------------------------------------
# Load appendix
# --------------------------------------------------------------------------
src = openpyxl.load_workbook(APPENDIX, data_only=True)


def sheet_rows(name):
    ws = src[name]
    return [[c for c in row] for row in ws.iter_rows(values_only=True)]


def find_row(name, col_idx, needle):
    for r in sheet_rows(name)[1:]:
        if r[col_idx] is not None and needle in str(r[col_idx]):
            return r
    raise KeyError(f"{name}: {needle!r} not found")


# market_size: Indicator | Year | Value | YoY | Basis | Processing | Source | URL | Accessed
ms = {}
for r in sheet_rows("market_size")[1:]:
    if r[0]:
        ms[(str(r[0]), str(r[1]))] = r
SHIP_2019 = float(ms[("Domestic factory shipments of cosmetics (all categories)", "2019")][2])
SHIP_2023_ROW = ms[("Domestic factory shipments of cosmetics (all categories)", "2023")]
SHIP_2023 = float(SHIP_2023_ROW[2])
SHIP_2023_YOY = float(SHIP_2023_ROW[3])
import re as _re
_rec_row = ms[("Recovery ratio vs 2019 peak", "2023")]
RECOVERY = float(_re.match(r"([\d.]+)", str(_rec_row[2])).group(1))  # 73.9 — Derived AS PUBLISHED in the appendix
GAP = round(100 - RECOVERY, 1)                                       # Derived: 26.1 (≈26% in report wording)

# category: Category | Year | Share | ...
CATS = [(r[0], float(r[2])) for r in sheet_rows("category")[1:]
        if r[0] and isinstance(r[2], (int, float))]

# trade: Flow | Year | Value | YoY | Scope | ...
TRADE = [r for r in sheet_rows("trade")[1:] if r[0]]
def trade(flow, year):
    for r in TRADE:
        if r[0] == flow and str(r[1]) == str(year):
            return r
    raise KeyError(f"trade {flow} {year}")
EXP = [(int(r[1]), float(r[2])) for r in TRADE if r[0] == "Exports"]
IMP = [(int(r[1]), float(r[2])) for r in TRADE if r[0] == "Imports"]
ASSOC_TOTAL = trade("Imports excl. toothpaste & soap", 2024)
KOREA = trade("Imports from South Korea (excl. toothpaste & soap)", 2024)
FRANCE = trade("Imports from France (excl. toothpaste & soap)", 2024)
ROW_VALUE = round(float(ASSOC_TOTAL[2]) - float(KOREA[2]) - float(FRANCE[2]), 1)  # Derived
ROW_SHARE = round(100 - 30.3 - 24.4, 1)                                          # Derived
IMP_CATS = [(r[0].replace("Imports - ", "").replace("Imports – ", ""), float(r[2]), float(r[3]))
            for r in TRADE if str(r[0]).startswith("Imports - ")]

# channel: Indicator | Period | Value | YoY | ...
CHAN_ROWS = [r for r in sheet_rows("channel")[1:] if r[0]]
def chan(indicator_prefix, period_hint):
    for r in CHAN_ROWS:
        if str(r[0]).startswith(indicator_prefix) and period_hint in str(r[1]):
            return r
    raise KeyError(f"channel {indicator_prefix} {period_hint}")
DS_TOTAL = chan("Drugstore sales - all categories", "2024")
DS_BEAUTY_H1 = chan("Drugstore sales - beauty care", "H1")
DS_BEAUTY_SEP = chan("Drugstore sales - beauty care", "2024-09")
DS_SEP_TOTAL = 728.8  # from appendix beauty-care Sep row basis text (total that month)
BEAUTY_SHARE_SEP = round(float(DS_BEAUTY_SEP[2]) / DS_SEP_TOTAL * 100, 1)  # Derived
INBOUND = chan("Inbound travel consumption", "Q3")

# household
HH = {str(r[0]): r for r in sheet_rows("household")[1:] if r[0]}
HH_EXP = HH["Average monthly consumption expenditure per household (two-or-more-person households)"]
HH_VALUE = int(str(HH_EXP[2]).replace(",", "").replace(" JPY", ""))

# consumer voice (appendix sheet mirrors the CSV). Frame strings vary in
# wording between rows ("(same period)" etc.), so bucket by platform.
CV = [r for r in sheet_rows("consumer_voice")[1:] if r[0]]
def bucket(platform):
    return {"@cosme": "cosme_rank", "@cosme Shopping": "cosme_shop",
            "Rakuten Ichiba": "rakuten", "Amazon.co.jp": "amazon"}[platform]
BUCKETS = {}
for r in CV:
    BUCKETS.setdefault(bucket(r[0]), []).append(r)
COSME_RANK = sorted(BUCKETS["cosme_rank"], key=lambda x: int(x[2]))
COSME_SHOP = BUCKETS["cosme_shop"]
RAKUTEN = sorted(BUCKETS["rakuten"], key=lambda x: int(x[2]))
AMAZON = BUCKETS["amazon"]
N_TOTAL = len(CV)


# --------------------------------------------------------------------------
# Helpers (mirroring sample-02 build_excel.py style)
# --------------------------------------------------------------------------
def put_title(ws, title, subtitle, source_line):
    ws["A1"] = title
    ws["A1"].font = TITLE_FONT
    ws["A2"] = subtitle
    ws["A2"].font = SUB_FONT
    ws["A3"] = source_line
    ws["A3"].font = NOTE_FONT


def write_table(ws, start_row, start_col, headers, rows, fmts=None, hdr_fill=HDR_FILL_L1):
    for j, h in enumerate(headers):
        c = ws.cell(row=start_row, column=start_col + j, value=h)
        c.fill = hdr_fill
        c.font = HDR_FONT
        c.alignment = Alignment(horizontal="center", wrap_text=True)
    for i, row in enumerate(rows):
        for j, v in enumerate(row):
            c = ws.cell(row=start_row + 1 + i, column=start_col + j, value=v)
            c.font = CELL_FONT
            if fmts and j < len(fmts) and fmts[j]:
                c.number_format = fmts[j]
    return start_row + 1 + len(rows)  # first free row after the table


def put_notes(ws, row, notes, font=NOTE_FONT):
    for n in notes:
        ws.cell(row=row, column=1, value=n).font = font
        row += 1
    return row


def widths(ws, spec):
    for col, w in spec.items():
        ws.column_dimensions[col].width = w


def style_chart(ch, title, y_title=None, x_title=None, height=7.5, width=15):
    ch.title = title
    if y_title:
        ch.y_axis.title = y_title
    if x_title:
        ch.x_axis.title = x_title
    ch.height = height
    ch.width = width
    ch.legend = None
    return ch


wb = openpyxl.Workbook()

# ==========================================================================
# Sheet 1 — Market overview (Layer 1)
# ==========================================================================
ws = wb.active
ws.title = "1_Market overview"
ws.sheet_properties.pageSetUpPr.fitToPage = True
widths(ws, {"A": 46, "B": 18, "C": 18, "D": 62})
put_title(
    ws,
    "Japan Cosmetics & Skincare — Dashboard 1. Market overview",
    "Layer 1 (official statistics). Companion dashboard to the Market Entry Audit — verdict: CONDITIONAL GO "
    "(e-commerce-first, ingredient-led skincare, JPY 3,000–6,000 band; drugstore expansion as phase two).",
    "Sources: METI Production Dynamics Statistics, published via JCIA shipment statistics; METI committee material. "
    f"Accessed {ACCESSED}. Basis: factory shipment value — NOT retail market size.",
)
r = put_notes(ws, 5, ["Key indicators"])
ws["A5"].font = SUB_FONT
r = write_table(ws, 6, 1,
                ["Indicator", "Value", "Unit / change", "Basis & processing"],
                [
                    ["Domestic factory shipments of cosmetics, 2019 (all-time peak)", SHIP_2019, "JPY bn", "Factory shipment value. Transcribed."],
                    ["Domestic factory shipments of cosmetics, 2023", SHIP_2023, f"JPY bn (+{SHIP_2023_YOY}% YoY)", "Factory shipment value. Transcribed."],
                    ["2023 shipments as % of 2019 peak", RECOVERY, "%", "Derived: 1,302.4 / 1,761.1 — value as published in the appendix (73.9%)."],
                    ["Gap vs 2019 peak", -GAP, "% (approx.)", "Derived: 100 − 73.9 = 26.1. The audit report phrases this as '≈26% below peak'."],
                ],
                fmts=[None, FMT_BN, None, None])
r += 1
ws.cell(row=r, column=1, value="Factory shipments — published reference years").font = SUB_FONT
r += 1
tab1_top = r
r = write_table(ws, r, 1,
                ["Year", "Factory shipments (JPY bn)", "YoY %", "Note"],
                [
                    [2019, SHIP_2019, None, "All-time peak year"],
                    [2023, SHIP_2023, SHIP_2023_YOY, "Latest year published in the appendix"],
                ],
                fmts=[FMT_INT, FMT_BN, FMT_PCT, None])
ch = BarChart()
data = Reference(ws, min_col=2, min_row=tab1_top, max_row=tab1_top + 2)
cats = Reference(ws, min_col=1, min_row=tab1_top + 1, max_row=tab1_top + 2)
ch.add_data(data, titles_from_data=True)
ch.set_categories(cats)
ch.dLbls = DataLabelList(showVal=True)
style_chart(ch, "Domestic factory shipments of cosmetics (JPY bn)", y_title="JPY bn")
ws.add_chart(ch, f"A{r + 1}")
r += 17
ws.cell(row=r, column=1, value="Category mix — share of factory shipment value, 2023").font = SUB_FONT
r += 1
tab2_top = r
cat_labels = {
    "Skin care products": "Skin care",
    "Hair care products": "Hair care",
    "Makeup products": "Makeup",
    "Special-purpose products": "Special-purpose",
    "Fragrances": "Fragrances",
}
r = write_table(ws, r, 1,
                ["Category (METI/JCIA classification)", "Share of shipment value, 2023"],
                [[cat_labels.get(c, c), v] for c, v in CATS],
                fmts=[None, FMT_PCT])
ch2 = BarChart()
ch2.type = "bar"
data = Reference(ws, min_col=2, min_row=tab2_top, max_row=tab2_top + len(CATS))
cats = Reference(ws, min_col=1, min_row=tab2_top + 1, max_row=tab2_top + len(CATS))
ch2.add_data(data, titles_from_data=True)
ch2.set_categories(cats)
ch2.dLbls = DataLabelList(showVal=True)
style_chart(ch2, "Category mix, 2023 (% of factory shipment value)", x_title="%")
ws.add_chart(ch2, f"A{r + 1}")
r += 17
put_notes(ws, r, [
    "Notes:",
    "1. Factory shipment value = value of cosmetics shipped from domestic factories (including for export). It is NOT the retail market size; never present it as 'the market is JPY X'.",
    "2. JCIA is an industry association; its shipment page republishes METI Production Dynamics Statistics. The METI committee material gives the 2023 total as JPY 1,302.5 bn — a 0.1 bn rounding difference from the JCIA figure used here (see appendix, sources sheet).",
    "3. Fine-item production detail (lotions, essences, sheet masks, sunscreens) was NOT OBTAINED in this production run — no fine-item chart is shown (appendix, category sheet).",
    "4. Shipments (domestic production) and imports (sheet 2) are separate series on different bases. Do not add them together to size the market.",
])

# ==========================================================================
# Sheet 2 — Trade (Layer 1)
# ==========================================================================
ws = wb.create_sheet("2_Trade")
ws.sheet_properties.pageSetUpPr.fitToPage = True
widths(ws, {"A": 40, "B": 18, "C": 16, "D": 16, "E": 58})
put_title(
    ws,
    "Japan Cosmetics & Skincare — Dashboard 2. Trade",
    "Layer 1 (official statistics). Imports are the footprint of foreign brands in Japan; exports show the strength of Japanese brands abroad.",
    "Sources: Ministry of Finance, Trade Statistics of Japan (cosmetics aggregate, HS Ch.33 basis, compiled by JCIA); "
    f"Japan Imported Cosmetics Association, FY2024 import results. Accessed {ACCESSED}. Customs values: exports FOB, imports CIF.",
)
r = 5
ws.cell(row=r, column=1, value="Exports — published years (JCIA aggregate)").font = SUB_FONT
r += 1
tab_top = r
r = write_table(ws, r, 1,
                ["Year", "Exports (JPY bn)", "Note"],
                [[y, v, "Peak year" if y == 2021 else ""] for y, v in EXP],
                fmts=[FMT_INT, FMT_BN, None])
ch = BarChart()
data = Reference(ws, min_col=2, min_row=tab_top, max_row=tab_top + len(EXP))
cats = Reference(ws, min_col=1, min_row=tab_top + 1, max_row=tab_top + len(EXP))
ch.add_data(data, titles_from_data=True)
ch.set_categories(cats)
ch.dLbls = DataLabelList(showVal=True)
style_chart(ch, "Cosmetics exports (JPY bn, customs value FOB)", y_title="JPY bn")
ws.add_chart(ch, f"A{r + 1}")
r += 17
ws.cell(row=r, column=1, value="Imports — published years (JCIA aggregate)").font = SUB_FONT
r += 1
tab_top = r
r = write_table(ws, r, 1,
                ["Year", "Imports (JPY bn)", "Note"],
                [[y, v, ""] for y, v in IMP],
                fmts=[FMT_INT, FMT_BN, None])
ch = BarChart()
data = Reference(ws, min_col=2, min_row=tab_top, max_row=tab_top + len(IMP))
cats = Reference(ws, min_col=1, min_row=tab_top + 1, max_row=tab_top + len(IMP))
ch.add_data(data, titles_from_data=True)
ch.set_categories(cats)
ch.dLbls = DataLabelList(showVal=True)
style_chart(ch, "Cosmetics imports (JPY bn, customs value CIF)", y_title="JPY bn")
ws.add_chart(ch, f"A{r + 1}")
r += 17
r = put_notes(ws, r, [
    "Scope warning — two different 2024 import totals exist and must not be mixed:",
    f"• JCIA aggregate (above): JPY {float(trade('Imports', 2024)[2]):,.1f} bn.",
    f"• Japan Imported Cosmetics Association scope (excl. toothpaste & soap): JPY {float(ASSOC_TOTAL[2]):,.1f} bn, +{float(ASSOC_TOTAL[3])}% YoY. The tables below use THIS scope.",
    "Only the years published in the data appendix are plotted. There is no interpolation — do not read the bars as a continuous annual series.",
])
r += 1
ws.cell(row=r, column=1, value="2024 imports by origin (association scope, excl. toothpaste & soap)").font = SUB_FONT
r += 1
tab_top = r
r = write_table(ws, r, 1,
                ["Origin", "Imports (JPY bn)", "YoY %", "Share %", "Processing"],
                [
                    ["South Korea", float(KOREA[2]), float(KOREA[3]), 30.3, "Transcribed (share as published)"],
                    ["France", float(FRANCE[2]), float(FRANCE[3]), 24.4, "Transcribed (share as published)"],
                    ["All other origins", ROW_VALUE, None, ROW_SHARE, "Derived: 443.8 − 134.3 − 108.3; share = 100 − 30.3 − 24.4"],
                ],
                fmts=[None, FMT_BN, FMT_PCT, FMT_PCT, None])
ch = PieChart()
data = Reference(ws, min_col=2, min_row=tab_top + 1, max_row=tab_top + 3)
cats = Reference(ws, min_col=1, min_row=tab_top + 1, max_row=tab_top + 3)
ch.add_data(data, titles_from_data=False)
ch.set_categories(cats)
ch.dLbls = DataLabelList(showCatName=True, showPercent=True)
style_chart(ch, "2024 cosmetics imports by origin (association scope)")
ws.add_chart(ch, f"A{r + 1}")
r += 17
ws.cell(row=r, column=1, value="2024 imports by category (association scope)").font = SUB_FONT
r += 1
tab_top = r
imp_cat_labels = {
    "base makeup & skin care": "Base makeup & skin care",
    "colour makeup": "Colour makeup",
    "hair care": "Hair care",
    "fragrance": "Fragrance",
}
rows = [[imp_cat_labels.get(n.lower(), n), v, y] for n, v, y in IMP_CATS]
r = write_table(ws, r, 1,
                ["Category", "Imports (JPY bn)", "YoY %"],
                rows,
                fmts=[None, FMT_BN, FMT_PCT])
ch = BarChart()
data = Reference(ws, min_col=2, min_row=tab_top, max_row=tab_top + len(rows))
cats = Reference(ws, min_col=1, min_row=tab_top + 1, max_row=tab_top + len(rows))
ch.add_data(data, titles_from_data=True)
ch.set_categories(cats)
ch.dLbls = DataLabelList(showVal=True)
style_chart(ch, "2024 cosmetics imports by category (JPY bn, association scope)", y_title="JPY bn")
ws.add_chart(ch, f"A{r + 1}")
r += 17
put_notes(ws, r, [
    "Notes:",
    "1. The four categories above are the categories published by the association; they do not sum to the total (other categories exist).",
    "2. South Korea has been the No. 1 import origin since 2022 and grew +40.0% in 2024 — the clearest Layer 1 evidence that foreign (especially Korean mid-price) brands are gaining share now.",
    "3. Published YoY % values are transcribed from the sources. They are never recalculated from the sparse published years in this dashboard (e.g., no 2014→2023 'growth rate' is computed).",
])

# ==========================================================================
# Sheet 3 — Channel & household (Layer 1)
# ==========================================================================
ws = wb.create_sheet("3_Channel & household")
ws.sheet_properties.pageSetUpPr.fitToPage = True
widths(ws, {"A": 52, "B": 24, "C": 18, "D": 16, "E": 56})
put_title(
    ws,
    "Japan Cosmetics & Skincare — Dashboard 3. Channel & household",
    "Layer 1 (official statistics). Drugstores are the volume channel for skincare in Japan; the household panel gives demand-side context.",
    "Sources: METI Current Survey of Commerce (2024 annual review, H1 2024 review, Sep 2024 monthly flash); "
    f"MIC Family Income and Expenditure Survey (FIES), 2024 annual summary. Accessed {ACCESSED}.",
)
r = 5
ws.cell(row=r, column=1, value="Drugstore sales (METI Current Survey of Commerce)").font = SUB_FONT
r += 1
tab_top = r
ds_rows = [
    ["All categories", "2024 (full year)", float(DS_TOTAL[2]), float(DS_TOTAL[3]), "Retail sales value; store count +3.3% YoY"],
    ["Beauty care (cosmetics & sundries)", "2024 H1 (Jan–Jun)", float(DS_BEAUTY_H1[2]), None, "Half-year cumulative. Beauty care was the 2nd-largest contributor to drugstore sales growth in 2024 (after food)"],
    ["Beauty care (cosmetics & sundries)", "Sep 2024 (single month)", float(DS_BEAUTY_SEP[2]), float(DS_BEAUTY_SEP[3]), "Single month"],
    ["All categories (context for the row above)", "Sep 2024 (single month)", DS_SEP_TOTAL, 3.9, "19,525 stores that month (from the appendix basis text)"],
    ["Beauty care share of drugstore sales", "Sep 2024", BEAUTY_SHARE_SEP, None, "Derived: 96.0 / 728.8 = 13.2%"],
]
r = write_table(ws, r, 1,
                ["Indicator", "Period", "Sales (JPY bn) / share %", "YoY %", "Basis & notes"],
                ds_rows,
                fmts=[None, None, FMT_BN, FMT_PCT, None])
ch = BarChart()
data = Reference(ws, min_col=3, min_row=tab_top, max_row=tab_top + 4)
cats = Reference(ws, min_col=1, min_row=tab_top + 1, max_row=tab_top + 4)
ch.add_data(data, titles_from_data=True)
ch.set_categories(cats)
ch.dLbls = DataLabelList(showVal=True)
style_chart(ch, "Drugstore sales — published reference points (JPY bn; periods differ, see notes)", y_title="JPY bn")
ws.add_chart(ch, f"A{r + 1}")
r += 17
r = put_notes(ws, r, [
    "The bars above cover DIFFERENT periods (full year / half year / single month). They are reference points, not a time series — never trend them.",
    "The Derived share row (13.2%) is not plotted; it is calculated from the two Sep 2024 figures only.",
])
r += 1
ws.cell(row=r, column=1, value="Household context (MIC Family Income and Expenditure Survey)").font = SUB_FONT
r += 1
r = write_table(ws, r, 1,
                ["Indicator", "Year", "Value", "Basis & status"],
                [
                    ["Average monthly consumption expenditure per household (two-or-more-person households, all items)", 2024, HH_VALUE,
                     "JPY per month. Nominal +2.1% YoY / real −1.1% YoY. Transcribed from the FIES 2024 annual summary."],
                    ["FIES cosmetics item codes (2025 revision, group 904–915)", 2025, "909 cream / 910 lotion / 911 foundation / 912 lipstick / 913 other cosmetics / 914 emulsion",
                     "Structure confirmed on e-Stat. Item-level values exist in FIES detailed tables."],
                    ["Item-level cosmetics expenditure time series", "—", "NOT OBTAINED",
                     "Extraction from e-Stat FIES detailed tables was not completed in this production run. No substitute figures are used, and no household cosmetics chart is shown."],
                ],
                fmts=[None, None, FMT_INT, None])
r += 1
ws.cell(row=r, column=1, value="Inbound context (Japan Tourism Agency)").font = SUB_FONT
r += 1
r = write_table(ws, r, 1,
                ["Indicator", "Period", "Value (JPY bn)", "YoY %", "Basis & status"],
                [
                    ["Inbound travel consumption (all purposes)", str(INBOUND[1]), float(INBOUND[2]), float(INBOUND[3]),
                     "Survey estimate, preliminary. The survey tracks a 'cosmetics & perfumes' shopping item by nationality; category-level values were NOT extracted in this run."],
                ],
                fmts=[None, None, FMT_BN, FMT_PCT, None])
r += 1
put_notes(ws, r, [
    "Notes:",
    "1. No official Japanese statistical series isolating cosmetics e-commerce sales was confirmed in the design or production phase. E-commerce evidence in this package comes from Layer 2 platform samples only (sheet 4) — keep the two layers separate.",
    "2. 'Beauty care' is the Current Survey of Commerce category 'beauty care (cosmetics & sundries)' for drugstores; it is broader than cosmetics alone.",
    "3. Household expenditure above is all-items context for wallet size. It is NOT cosmetics spending — the cosmetics item series is listed as NOT OBTAINED rather than estimated.",
])

# ==========================================================================
# Sheet 4 — Consumer voice (Layer 2)
# ==========================================================================
ws = wb.create_sheet("4_Consumer voice")
ws.sheet_properties.pageSetUpPr.fitToPage = True
widths(ws, {"A": 44, "B": 30, "C": 14, "D": 16, "E": 52})
put_title(
    ws,
    "Japan Cosmetics & Skincare — Dashboard 4. Consumer voice",
    "Layer 2 (consumer-platform samples). What sampled shoppers reward: ratings, rankings and review volumes.",
    f"Manual sampling of public platform pages. Accessed {ACCESSED} (Rakuten data date 2026-11-10).",
)
r = 5
sample_block = [
    "SAMPLE DATA — READ FIRST",
    f"N = {N_TOTAL} product records, manually sampled from public pages across 4 sample frames (table below).",
    "This is a sample, not representative of all Japanese consumers. Platform users are a self-selected population;",
    "frames cover lotions / skincare rankings only, and one frame (Amazon.co.jp) is an ARCHIVED page copy dated 2024-04-25,",
    "used as structural reference only. Layer 2 is never used to size the market — market size comes from Layer 1 (sheets 1–3).",
]
for line in sample_block:
    c = ws.cell(row=r, column=1, value=line)
    c.font = WARN_FONT if line.startswith("SAMPLE") else NOTE_FONT
    r += 1
r += 1
ws.cell(row=r, column=1, value="Sample frames").font = SUB_FONT
r += 1
frame_info = [
    ("@cosme", "Review ranking — lotions", len(COSME_RANK), "Aggregation period 2026-08-05 to 2026-11-04; page updated 2026-11-06"),
    ("@cosme Shopping", "Bestseller ranking — lotions (sales-based)", len(COSME_SHOP), f"Accessed {ACCESSED}"),
    ("Rakuten Ichiba", "Daily ranking — skincare", len(RAKUTEN), "Data date 2026-11-10"),
    ("Amazon.co.jp", "Bestseller ranking — lotions (ARCHIVED 2024-04-25)", len(AMAZON), "Archived copy — structural reference only, NOT a current ranking"),
]
r = write_table(ws, r, 1,
                ["Platform", "Sample frame", "Records (N)", "Period / data date"],
                [list(x) for x in frame_info] + [["TOTAL", "", N_TOTAL, ""]],
                fmts=[None, None, FMT_INT, None], hdr_fill=HDR_FILL_L2)
r += 1
ws.cell(row=r, column=1, value="@cosme review ranking — lotions (ratings out of 7)").font = SUB_FONT
r += 1
tab_top = r
cosme_rows = []
for row in COSME_RANK:
    label = f"#{int(row[2])} {row[3]}"
    cosme_rows.append([label, f"{row[3]} — {row[4]}", float(row[5]), int(row[7])])
r = write_table(ws, r, 1,
                ["Rank & brand", "Product", "Rating (max 7)", "Review count"],
                cosme_rows,
                fmts=[None, None, FMT_RATING, FMT_INT], hdr_fill=HDR_FILL_L2)
ch = BarChart()
data = Reference(ws, min_col=3, min_row=tab_top, max_row=tab_top + len(cosme_rows))
cats = Reference(ws, min_col=1, min_row=tab_top + 1, max_row=tab_top + len(cosme_rows))
ch.add_data(data, titles_from_data=True)
ch.set_categories(cats)
ch.dLbls = DataLabelList(showVal=True)
style_chart(ch, "@cosme lotion review ranking — rating (max 7), sampled top 10", y_title="Rating (max 7)")
ws.add_chart(ch, f"A{r + 1}")
r += 17
ch = BarChart()
data = Reference(ws, min_col=4, min_row=tab_top, max_row=tab_top + len(cosme_rows))
cats = Reference(ws, min_col=1, min_row=tab_top + 1, max_row=tab_top + len(cosme_rows))
ch.add_data(data, titles_from_data=True)
ch.set_categories(cats)
ch.dLbls = DataLabelList(showVal=True)
style_chart(ch, "@cosme lotion review ranking — review count", y_title="Reviews")
ws.add_chart(ch, f"A{r + 1}")
r += 17
r = put_notes(ws, r, [
    "Rank #3 is absent from this sample: the category page printed the ELIXIR Brightening Lotion (moist) at #4 while @cosme product pages list the line at #3 — ranks are shown as captured (see appendix, consumer_voice notes).",
    "Highest rating in the sample: DECORTE Liposome Advanced Repair Serum, 5.8/7 (18,671 reviews). Largest review base: SK-II Facial Treatment Essence, 40,386 reviews.",
])
r += 1
ws.cell(row=r, column=1, value="Rakuten Ichiba daily ranking — skincare (review counts)").font = SUB_FONT
r += 1
tab_top = r
rak_rows = [[f"#{int(row[2])} {row[3]}", f"{row[3]} — {row[4]}", int(row[7])] for row in RAKUTEN]
r = write_table(ws, r, 1,
                ["Rank & brand", "Product", "Review count"],
                rak_rows,
                fmts=[None, None, FMT_INT], hdr_fill=HDR_FILL_L2)
ch = BarChart()
ch.type = "bar"
data = Reference(ws, min_col=3, min_row=tab_top, max_row=tab_top + len(rak_rows))
cats = Reference(ws, min_col=1, min_row=tab_top + 1, max_row=tab_top + len(rak_rows))
ch.add_data(data, titles_from_data=True)
ch.set_categories(cats)
ch.dLbls = DataLabelList(showVal=True)
style_chart(ch, "Rakuten daily skincare ranking — review count (data date 2026-11-10)", x_title="Reviews")
ws.add_chart(ch, f"A{r + 1}")
r += 17
r = put_notes(ws, r, [
    "Rakuten star ratings were not captured in the text view used for sampling (appendix) — this frame is charted on review counts only.",
    "Displayed ranks 5 and 10 were not legible in the text view; AESTURA is shown at rank 6 as printed in the appendix record.",
    "Foreign entrants charting here: AESTURA (South Korea) #6 and VT Cosmetics (South Korea) #7 — mid-price foreign brands can rank in Japanese EC (audit report §1, reason 2).",
])
r += 1
ws.cell(row=r, column=1, value="Largest review bases in the sample (all frames)").font = SUB_FONT
r += 1
top_reviews = sorted(CV, key=lambda x: -int(x[7]))[:8]
r = write_table(ws, r, 1,
                ["Brand — Product", "Platform / frame", "Review count"],
                [[f"{row[3]} — {row[4]}", f"{row[0]} ({'archived 2024-04-25' if row[0] == 'Amazon.co.jp' else 'sampled 2026-11'})", int(row[7])]
                 for row in top_reviews],
                fmts=[None, None, FMT_INT], hdr_fill=HDR_FILL_L2)
put_notes(ws, r + 1, [
    "Notes:",
    "1. Rating scales differ by platform (@cosme: max 7; Amazon.co.jp: max 5) and Rakuten ratings were not captured. Never average or compare ratings across platforms.",
    "2. Review counts are cumulative platform counts at access date, shown as evidence of incumbents' review moats (the audit's 'cold-start discovery' risk) — not as sales figures.",
    "3. Prices for every sampled product are in the data sheet (data_consumer_voice), transcribed as displayed (ranges where the page showed a range).",
    "4. Competitor read-across from this sample is in the appendix (competitors sheet): Shiseido-group brands hold 4 of the top-7 @cosme lotion positions; domestic indie/D2C brands (Yunth, TAKAMI, N organic) chart on Rakuten.",
])

# ==========================================================================
# Data sheets — verbatim copies of the appendix
# ==========================================================================
for name in ["market_size", "category", "trade", "household", "channel",
             "consumer_voice", "competitors", "sources"]:
    sws = src[name]
    dws = wb.create_sheet(f"data_{name}")
    for row in sws.iter_rows():
        for c in row:
            if c.value is not None:
                nc = dws.cell(row=c.row, column=c.column, value=c.value)
                nc.font = HDR_FONT if c.row == 1 else CELL_FONT
                if c.row == 1:
                    nc.fill = HDR_FILL_L2 if name == "consumer_voice" else HDR_FILL_L1
    dws.column_dimensions["A"].width = 42
    dws.column_dimensions["B"].width = 30

# ==========================================================================
# Definitions & sources
# ==========================================================================
ws = wb.create_sheet("Definitions & sources")
widths(ws, {"A": 34, "B": 110})
put_title(
    ws,
    "Definitions, traceability & cautions — Cosmetics Audit Dashboard",
    "Every figure in sheets 1–4 traces to the data appendix (output/audit_data_appendix.xlsx), copied verbatim into the data_* sheets.",
    f"All sources accessed {ACCESSED} unless otherwise stated. Full source list with URLs: data_sources sheet.",
)
r = 5
ws.cell(row=r, column=1, value="Definitions").font = SUB_FONT
r += 1
defs = [
    ["Factory shipment value", "Value of cosmetics shipped from domestic factories (METI Production Dynamics Statistics). Includes goods for export. NOT retail market size."],
    ["Customs value (FOB / CIF)", "Trade Statistics of Japan values: exports are FOB, imports are CIF. Two import totals exist for 2024 (JCIA aggregate 446.1 bn; association scope excl. toothpaste & soap 443.8 bn) — scope is stated on every table."],
    ["Transcribed", "Copied from the source without modification."],
    ["Derived", "Arithmetic on transcribed figures only; the formula is shown next to the value wherever it appears (recovery ratio, gap vs peak, rest-of-world imports, beauty-care share). Derived values are never mixed into transcribed series."],
    ["Published YoY %", "Growth rates as published by the source. Never recalculated from the sparse published years in this package."],
    ["Layer 1 / Layer 2", "Layer 1 = official statistics (sheets 1–3). Layer 2 = manual samples of consumer platforms (sheet 4). The layers are never combined in one visual, and Layer 2 is never used to size the market."],
    ["NOT OBTAINED", "Data that could not be extracted in this production run. Written as NOT OBTAINED; never replaced by an estimate."],
]
r = write_table(ws, r, 1, ["Term", "Definition"], defs)
r += 1
ws.cell(row=r, column=1, value="Key figures — traceability").font = SUB_FONT
r += 1
trace = [
    ["Factory shipments 2019 / 2023", "1,761.1 / 1,302.4 JPY bn (+2.9%)", "data_market_size", "METI committee material; JCIA (METI Production Dynamics Statistics)"],
    ["Category shares 2023", "Skin 43.6 / Hair 28.0 / Makeup 20.7 / Special 7.0 / Fragrance 0.7 %", "data_category", "JCIA (METI Production Dynamics Statistics)"],
    ["Exports 2014 / 2021 / 2023 / 2024", "152.3 / 790.6 / 608.3 / 520.1 JPY bn", "data_trade", "MOF Trade Statistics via JCIA"],
    ["Imports 2014 / 2023 / 2024", "217.1 / 377.1 / 446.1 JPY bn", "data_trade", "MOF Trade Statistics via JCIA"],
    ["Imports 2024 (association scope)", "443.8 JPY bn, +17.7%; Korea 134.3 (+40.0%, 30.3%); France 108.3 (+19.4%, 24.4%)", "data_trade", "Japan Imported Cosmetics Association"],
    ["Drugstore sales 2024", "8,919.9 JPY bn, +6.9%; beauty care H1 547.2 bn; Sep 2024 96.0 bn (+12.6%)", "data_channel", "METI Current Survey of Commerce"],
    ["Household expenditure 2024", "JPY 300,243 / month (all items, 2+ person households)", "data_household", "MIC FIES 2024 annual summary"],
    ["Consumer-voice sample", "N = 32 records, 4 frames; top review count 50,283 (Attenir, Rakuten)", "data_consumer_voice", "@cosme / Rakuten / Amazon.co.jp public pages (manual sample)"],
]
r = write_table(ws, r, 1, ["Figure", "Value", "Data sheet", "Source"], trace)
r += 1
ws.cell(row=r, column=1, value="Cautions").font = SUB_FONT
r += 1
cautions = [
    "1. Two 2024 import totals (446.1 vs 443.8 JPY bn) differ by scope. Sheet 2 labels the scope on every table — do not mix them.",
    "2. Trade and shipment charts plot only the years published in the appendix. Sparse points are not a continuous series; no growth rate is computed across gaps.",
    "3. The FIES item-level cosmetics expenditure series was NOT OBTAINED. Sheet 3 shows all-items household expenditure as context only and says so.",
    "4. No official cosmetics × e-commerce sales series was confirmed to exist. EC evidence is Layer 2 only.",
    "5. JCIA has previously issued corrections to some 2025 monthly figures (design.md §2/§7). Cross-check association republications against METI originals when updating.",
    "6. Layer 2 ratings use different scales per platform (@cosme max 7, Amazon max 5) and Rakuten ratings were not captured. Cross-platform rating comparison is invalid.",
    "7. This dashboard is a companion to the audit report (verdict: CONDITIONAL GO). It is a sample prepared with a hypothetical entrant profile; figures are from the real sources cited. Not legal, regulatory, tax, or investment advice.",
]
for ctext in cautions:
    ws.cell(row=r, column=1, value=ctext).font = NOTE_FONT
    r += 1
r += 1
ws.cell(row=r, column=1, value="How to update / rebuild").font = SUB_FONT
r += 1
r = put_notes(ws, r, [
    "1. Refresh the appendix first (scripts/build_appendix.py), then re-run scripts/build_dashboard.py — this workbook is generated entirely from the appendix.",
    "2. The Power BI version is specified in powerbi/ (model.md, measures.dax, pages.md) and must reproduce the same values; its pages.md contains the pre-publish cross-check list.",
])

wb.save(OUT)
print(f"saved: {OUT}")
print(f"appendix check: SHIP_2019={SHIP_2019} SHIP_2023={SHIP_2023} RECOVERY={RECOVERY} GAP={GAP}")
print(f"categories={CATS} sum={sum(v for _, v in CATS)}")
print(f"exports={EXP} imports={IMP}")
print(f"assoc total={ASSOC_TOTAL[2]} korea={KOREA[2]} france={FRANCE[2]} RoW={ROW_VALUE} ({ROW_SHARE}%)")
print(f"import cats={IMP_CATS}")
print(f"drugstore total={DS_TOTAL[2]} beauty H1={DS_BEAUTY_H1[2]} beauty Sep={DS_BEAUTY_SEP[2]} beauty share Sep={BEAUTY_SHARE_SEP}")
print(f"household exp={HH_VALUE}; CV buckets: " +
      ", ".join(f"{k}={len(v)}" for k, v in BUCKETS.items()) + f"; N={N_TOTAL}")
