#!/usr/bin/env python3
"""Build audit_data_appendix.xlsx for Sample 03 (Japan Cosmetics Market Entry Audit).
All figures are transcribed verbatim from the cited public sources; processed/derived
values are flagged in the 'Processing' column. No figures are estimated.
"""
import csv
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

BASE = Path(__file__).resolve().parent.parent
OUT = BASE / "output" / "audit_data_appendix.xlsx"
ACCESSED = "2026-11-12"

wb = Workbook()
hdr_font = Font(bold=True)
wrap = Alignment(wrap_text=True, vertical="top")

def sheet(name, headers, rows, widths):
    ws = wb.create_sheet(name)
    ws.append(headers)
    for c in range(1, len(headers) + 1):
        ws.cell(row=1, column=c).font = hdr_font
    for r in rows:
        ws.append(r)
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = wrap
    ws.freeze_panes = "A2"
    return ws

del wb["Sheet"]

# ---------------- market_size ----------------
sheet("market_size",
      ["Indicator", "Year", "Value (JPY bn)", "YoY %", "Basis / definition", "Processing", "Source", "Source URL", "Accessed"],
      [
        ["Domestic factory shipments of cosmetics (all categories)", 2019, 1761.1, "", "Factory shipment value; all-time peak year. NOT retail market size.", "Transcribed", "METI committee material citing Production Dynamics Statistics (Census of Manufactures basis)", "https://www.meti.go.jp/shingikai/mono_info_service/cosmetic_industry/pdf/001_04_00.pdf", ACCESSED],
        ["Domestic factory shipments of cosmetics (all categories)", 2023, 1302.4, 2.9, "Factory shipment value. NOT retail market size.", "Transcribed", "Japan Cosmetic Industry Association (JCIA) statistics, based on METI Production Dynamics Statistics", "https://jcia.org/user/statistics/shipment", ACCESSED],
        ["Recovery ratio vs 2019 peak", 2023, "73.9% of 2019 peak", "", "Derived: 1302.4 / 1761.1", "Derived (calculation shown in Basis)", "Calculated from the two rows above", "", ACCESSED],
      ], [52, 8, 16, 9, 44, 22, 46, 52, 11])

# ---------------- category ----------------
sheet("category",
      ["Category (METI/JCIA shipment classification)", "Year", "Share of factory shipment value %", "Basis / definition", "Processing", "Source", "Source URL", "Accessed"],
      [
        ["Skin care products", 2023, 43.6, "Share of domestic factory shipment value", "Transcribed", "JCIA statistics (METI Production Dynamics Statistics)", "https://jcia.org/user/statistics/shipment", ACCESSED],
        ["Hair care products", 2023, 28.0, "Share of domestic factory shipment value", "Transcribed", "JCIA statistics (METI Production Dynamics Statistics)", "https://jcia.org/user/statistics/shipment", ACCESSED],
        ["Makeup products", 2023, 20.7, "Share of domestic factory shipment value", "Transcribed", "JCIA statistics (METI Production Dynamics Statistics)", "https://jcia.org/user/statistics/shipment", ACCESSED],
        ["Special-purpose products", 2023, 7.0, "Share of domestic factory shipment value", "Transcribed", "JCIA statistics (METI Production Dynamics Statistics)", "https://jcia.org/user/statistics/shipment", ACCESSED],
        ["Fragrances", 2023, 0.7, "Share of domestic factory shipment value", "Transcribed", "JCIA statistics (METI Production Dynamics Statistics)", "https://jcia.org/user/statistics/shipment", ACCESSED],
        ["Fine-item production detail (lotions, essences, sheet masks, sunscreens etc.)", "—", "NOT OBTAINED in this production run", "e-Stat file-level extraction of Production Dynamics Statistics fine items was not completed; do not infer fine-item trends from the broad categories above.", "Not obtained", "e-Stat (stat code 00550200) - pending extraction", "https://www.e-stat.go.jp/stat-search/files?page=1&layout=dataset&toukei=00550200&toukei_kind=6", ACCESSED],
      ], [46, 8, 20, 52, 16, 44, 52, 11])

# ---------------- trade ----------------
sheet("trade",
      ["Flow", "Year", "Value (JPY bn)", "YoY %", "Scope / basis", "Processing", "Source", "Source URL", "Accessed"],
      [
        ["Exports", 2014, 152.3, "", "Cosmetics aggregate (HS Ch.33 basis as compiled by JCIA); customs value", "Transcribed", "Ministry of Finance, Trade Statistics of Japan (via JCIA compilation)", "https://jcia.org/user/statistics/trade", ACCESSED],
        ["Exports", 2021, 790.6, "", "Peak year; customs value", "Transcribed", "MOF Trade Statistics (via JCIA)", "https://jcia.org/user/statistics/trade", ACCESSED],
        ["Exports", 2023, 608.3, "", "Customs value", "Transcribed", "MOF Trade Statistics (via JCIA)", "https://jcia.org/user/statistics/trade", ACCESSED],
        ["Exports", 2024, 520.1, "", "Customs value", "Transcribed", "MOF Trade Statistics (via JCIA / Cosmetic Science journal table, Oct 2025)", "https://cosmetic-science.net/wp-content/uploads/2025/06/2.pdf", ACCESSED],
        ["Imports", 2014, 217.1, "", "Customs value", "Transcribed", "MOF Trade Statistics (via JCIA)", "https://jcia.org/user/statistics/trade", ACCESSED],
        ["Imports", 2023, 377.1, "", "Customs value", "Transcribed", "MOF Trade Statistics (via JCIA)", "https://jcia.org/user/statistics/trade", ACCESSED],
        ["Imports", 2024, 446.1, "", "Customs value", "Transcribed", "MOF Trade Statistics (via JCIA)", "https://jcia.org/user/statistics/trade", ACCESSED],
        ["Imports excl. toothpaste & soap", 2024, 443.8, 17.7, "Association scope (slightly narrower than the JCIA aggregate above; the two totals are close but not identical in scope)", "Transcribed", "Japan Imported Cosmetics Association, FY2024 import results (released 2025-02-10)", "https://www.this.ne.jp/news/19015/", ACCESSED],
        ["Imports from South Korea (excl. toothpaste & soap)", 2024, 134.3, 40.0, "Share 30.3% - largest origin since 2022", "Transcribed", "Japan Imported Cosmetics Association", "https://www.this.ne.jp/news/19015/", ACCESSED],
        ["Imports from France (excl. toothpaste & soap)", 2024, 108.3, 19.4, "Share 24.4%", "Transcribed", "Japan Imported Cosmetics Association", "https://www.this.ne.jp/news/19015/", ACCESSED],
        ["Imports - base makeup & skin care", 2024, 200.6, 20.1, "Category within association scope", "Transcribed", "Japan Imported Cosmetics Association", "https://www.this.ne.jp/news/19015/", ACCESSED],
        ["Imports - colour makeup", 2024, 73.0, 29.8, "Category within association scope", "Transcribed", "Japan Imported Cosmetics Association", "https://www.this.ne.jp/news/19015/", ACCESSED],
        ["Imports - hair care", 2024, 59.8, 0.7, "Category within association scope", "Transcribed", "Japan Imported Cosmetics Association", "https://www.this.ne.jp/news/19015/", ACCESSED],
        ["Imports - fragrance", 2024, 47.5, 21.1, "Category within association scope", "Transcribed", "Japan Imported Cosmetics Association", "https://www.this.ne.jp/news/19015/", ACCESSED],
      ], [40, 8, 15, 9, 50, 14, 46, 52, 11])

# ---------------- household ----------------
sheet("household",
      ["Indicator", "Year", "Value", "YoY %", "Basis / definition", "Processing", "Source", "Source URL", "Accessed"],
      [
        ["Average monthly consumption expenditure per household (two-or-more-person households)", 2024, "300,243 JPY", "Nominal +2.1 / Real -1.1", "All consumption items; FIES annual average", "Transcribed", "Ministry of Internal Affairs and Communications, Family Income and Expenditure Survey (FIES), 2024 annual summary", "https://www.stat.go.jp/data/kakei/sokuhou/tsuki/pdf/fies_gaikyo2024.pdf", ACCESSED],
        ["Cosmetics item codes available in FIES item classification (2025 revision)", 2025, "909 cosmetic cream / 910 skin lotion / 911 foundation / 912 lipstick / 913 other cosmetics / 914 emulsion (within 904-915 soap & cosmetics group)", "", "Item-level annual expenditure per household by item is published in FIES detailed tables on e-Stat", "Structure confirmed; values NOT extracted in this run", "e-Stat FIES dataset (stat code 00200561)", "https://www.e-stat.go.jp/stat-search/database?layout=dataset&statdisp_id=0004023601", ACCESSED],
        ["Item-level cosmetics expenditure time series", "—", "NOT OBTAINED in this production run", "", "Extraction from e-Stat FIES detailed tables (annual, item classification) is the first update task for a paid engagement; no substitute figures are used in the report.", "Not obtained", "e-Stat FIES detailed tables", "https://www.stat.go.jp/data/kakei/", ACCESSED],
      ], [54, 8, 46, 20, 52, 26, 48, 52, 11])

# ---------------- channel ----------------
sheet("channel",
      ["Indicator", "Period", "Value (JPY bn)", "YoY %", "Basis / definition", "Processing", "Source", "Source URL", "Accessed"],
      [
        ["Drugstore sales - all categories", "2024 (full year)", 8919.9, 6.9, "Retail sales value, all product categories; store count +3.3% YoY", "Transcribed", "METI, Current Survey of Commerce - 2024 annual review", "https://www.meti.go.jp/statistics/toppage/report/archive/kako/20250610_1.html", ACCESSED],
        ["Drugstore sales - beauty care (cosmetics & sundries)", "2024 H1 (Jan-Jun)", 547.2, "", "Retail sales value, half-year cumulative. Beauty care was the 2nd-largest contributor to drugstore sales growth in 2024 (after food).", "Transcribed", "METI, Current Survey of Commerce - H1 2024 review", "https://www.meti.go.jp/statistics/toppage/report/minikeizai/pdf/h2amini168j.pdf", ACCESSED],
        ["Drugstore sales - beauty care (cosmetics & sundries)", "2024-09 (single month)", 96.0, 12.6, "Retail sales value, single month; total drugstore sales that month JPY 728.8 bn (+3.9%), 19,525 stores", "Transcribed", "METI, Current Survey of Commerce - monthly flash, Sep 2024", "https://news.shoninsha.co.jp/statistics/262551", ACCESSED],
        ["Cosmetics x e-commerce dedicated official series", "—", "NOT CONFIRMED / NOT OBTAINED", "", "No official Japanese statistical series isolating cosmetics e-commerce sales was confirmed in the design or production phase. EC evidence in this report comes from Layer 2 platform samples only.", "Not obtained", "Design-phase finding (see design.md section 2)", "", ACCESSED],
        ["Inbound travel consumption (all purposes)", "2026 Q3", 2207.6, 3.2, "Survey estimate, preliminary. The survey tracks a 'cosmetics & perfumes' shopping item by nationality; category-level values were not extracted in this run.", "Transcribed (total only)", "Japan Tourism Agency, Inbound Consumption Trend Survey, Q3 2026 flash", "https://www.mlit.go.jp/kankocho/tokei_hakusyo/gaikokujinshohidoko.html", ACCESSED],
      ], [46, 16, 15, 9, 62, 20, 46, 52, 11])

# ---------------- consumer_voice ----------------
rows = []
with open(BASE / "data" / "layer2_consumer_voice.csv", encoding="utf-8") as f:
    for r in csv.DictReader(f):
        rows.append([r["platform"], r["sample_frame"], r["displayed_rank"], r["brand"], r["product"],
                     r["rating"], r["rating_scale"], r["review_count"], r["price_jpy"], r["accessed"], r["notes"]])
sheet("consumer_voice",
      ["Platform", "Sample frame", "Displayed rank", "Brand", "Product", "Rating", "Rating scale (max)", "Review count", "Price (JPY, tax incl. where shown)", "Accessed", "Notes"],
      rows, [16, 52, 11, 22, 46, 8, 12, 12, 20, 11, 56])

# ---------------- competitors ----------------
sheet("competitors",
      ["Group / brand", "Home base", "Positioning & price evidence (from Layer 2 samples)", "Main channels in Japan", "Layer 2 signal (sampled, accessed 2026-11-12)", "Source", "Accessed"],
      [
        ["Shiseido Group (ELIXIR, IPSA, d program)", "Japan (listed major)", "Mid-price skincare: ELIXIR lotions JPY 3,300; IPSA Aqua JPY 5,500/200ml", "Department stores, drugstores, brand EC, @cosme Shopping", "4 of the top-7 positions in the @cosme lotion review-ranking sample; ELIXIR also #2 in @cosme Shopping bestsellers and #2/#11 on Rakuten daily skincare", "Layer 2 samples in consumer_voice sheet", ACCESSED],
        ["KOSE (DECORTE)", "Japan (listed major)", "Prestige: Liposome Advanced Repair Serum JPY 8,800-17,600", "Department stores, brand EC", "Highest rating in @cosme sample (5.8/7) with 18,671 reviews", "Layer 2 samples in consumer_voice sheet", ACCESSED],
        ["Kao (Curel, SOFINA)", "Japan (listed major)", "Mass / sensitive-skin: Curel spray JPY 990-2,750; SOFINA lotion JPY 1,210-1,430", "Drugstores, mass retail, EC", "Curel is an @cosme Best Cosme Hall of Fame product (9,621 reviews, 5.6/7); SOFINA in @cosme top 6", "Layer 2 samples in consumer_voice sheet", ACCESSED],
        ["P&G (SK-II)", "Global (US)", "Prestige: Facial Treatment Essence JPY 12,650-29,150", "Department stores, duty free, brand EC", "Largest review base in @cosme sample: 40,386 reviews, 5.6/7, rank #2", "Layer 2 samples in consumer_voice sheet", ACCESSED],
        ["Rohto Pharmaceutical (Hada Labo, Melano CC)", "Japan (listed major)", "Mass medicated skincare: lotions from JPY 500-1,500", "Drugstores, EC", "Multiple top positions in Amazon.co.jp lotion bestsellers (archived copy dated 2024-04-25 - structural evidence only)", "Layer 2 samples in consumer_voice sheet", ACCESSED],
        ["Amorepacific (AESTURA) / VT Cosmetics", "South Korea", "Mid-price derma / sheet masks: JPY 3,300-4,560", "EC-led (official Rakuten stores), variety stores", "Both charted in Rakuten daily skincare top 10 (ranks 6 and 7, 2026-11-10 data) - evidence that foreign mid-price brands can chart in Japanese EC rankings", "Layer 2 samples in consumer_voice sheet", ACCESSED],
        ["Domestic indie / D2C (TOUVER, Yunth, N organic, TAKAMI)", "Japan", "Ingredient-led mid price: JPY 2,000-11,000", "Own EC, Rakuten/Amazon official stores, subscription", "Yunth serum: 45,427 Rakuten reviews; TAKAMI Skin Peel #9 on Rakuten; TOUVER #2 in Amazon lotion bestsellers (archived copy)", "Layer 2 samples in consumer_voice sheet", ACCESSED],
        ["Note on company financials", "—", "Segment revenue figures for the groups above are published in securities reports / IR materials but were NOT transcribed in this production run; this matrix uses price, channel and consumer-signal evidence only. No competitor revenue figures are stated anywhere in the report.", "—", "—", "Design rule (design.md section 3): unpublished figures are never filled with estimates", ACCESSED],
      ], [34, 16, 52, 34, 62, 30, 11])

# ---------------- sources ----------------
sheet("sources",
      ["#", "Source", "Publisher / type", "What it supports in this report", "URL", "Accessed"],
      [
        [1, "Production Dynamics Statistics (cosmetics) via JCIA shipment statistics", "METI statistics as published by Japan Cosmetic Industry Association (industry association)", "Factory shipment value 2023 and category shares", "https://jcia.org/user/statistics/shipment", ACCESSED],
        [2, "Situation of the cosmetics manufacturing industry (committee material)", "Ministry of Economy, Trade and Industry (government)", "2019 peak shipment value JPY 1,761.1 bn; 2023 value JPY 1,302.5 bn (rounding differs by 0.1 from JCIA figure)", "https://www.meti.go.jp/shingikai/mono_info_service/cosmetic_industry/pdf/001_04_00.pdf", ACCESSED],
        [3, "Trade Statistics of Japan - cosmetics exports/imports compilation", "Ministry of Finance (government), compiled by JCIA", "Export/import time series 2014-2024", "https://jcia.org/user/statistics/trade", ACCESSED],
        [4, "FY2024 cosmetics import results", "Japan Imported Cosmetics Association (industry association)", "2024 imports by origin and category", "https://www.this.ne.jp/news/19015/", ACCESSED],
        [5, "Current Survey of Commerce - 2024 annual review & H1 review; monthly flash Sep 2024", "METI (government)", "Drugstore total and beauty-care sales", "https://www.meti.go.jp/statistics/toppage/report/archive/kako/20250610_1.html", ACCESSED],
        [6, "Family Income and Expenditure Survey - 2024 annual summary", "Ministry of Internal Affairs and Communications (government)", "Household consumption context; cosmetics item code structure", "https://www.stat.go.jp/data/kakei/sokuhou/tsuki/pdf/fies_gaikyo2024.pdf", ACCESSED],
        [7, "Inbound Consumption Trend Survey", "Japan Tourism Agency (government)", "Inbound consumption total; survey design for cosmetics & perfumes item", "https://www.mlit.go.jp/kankocho/tokei_hakusyo/gaikokujinshohidoko.html", ACCESSED],
        [8, "@cosme review rankings & @cosme Shopping bestseller rankings (lotions)", "Layer 2 platform sample (public pages, manual sampling)", "Consumer-voice sample: ratings, review counts, prices", "https://www.cosme.net/categories/item/902/ranking/", ACCESSED],
        [9, "Rakuten Ichiba daily ranking - skincare", "Layer 2 platform sample (public pages, manual sampling)", "Consumer-voice sample: ranks, review counts, prices (data date 2026-11-10)", "https://ranking.rakuten.co.jp/daily/100944/", ACCESSED],
        [10, "Amazon.co.jp bestseller ranking - lotions", "Layer 2 platform sample. ARCHIVED page copy dated 2024-04-25 - used as structural evidence of the mass price band only, NOT as a current ranking.", "Mass-market price band and brand presence (dated)", "https://www.amazon.co.jp/gp/bestsellers/beauty/170092011", ACCESSED],
        [11, "USD/JPY reference rate 159.7 (2026-11-11)", "OANDA via Finnhub (market data)", "Indicative USD conversions, rounded to JPY 160 = USD 1 in the report text", "https://finnhub.io/?q=%22USD%2FJPY%22", ACCESSED],
        [12, "Channel research for this product format (Upwork audit comparables)", "Internal research file (2026-11-12)", "Existence of a fixed-price market-entry audit product (~USD 150) and typical Japan-research rates USD 18-27/hr on Upwork", "Local file: market-research-2026-11-11/channel_check_english_research.md", ACCESSED],
      ], [5, 52, 40, 58, 56, 11])

wb.save(OUT)
print("saved", OUT)
