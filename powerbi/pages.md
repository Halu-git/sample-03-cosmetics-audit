# Power BI Page Layout — Japan Cosmetics & Skincare Market Entry Audit (Dashboard Companion)

- 4 pages, matching the Excel companion (`cosmetics_audit_dashboard.xlsx`) sheet-for-sheet, with identical values. The Excel file is the reference if any visual disagrees.
- Common: each page carries its source copy block (below) in a text box at top-right. Pages are layer-fixed: pages 1–3 are Layer 1 only, page 4 is Layer 2 only. No report-level slicers; no cross-layer visuals.
- Layer colour convention (same as the report): Layer 1 = blue, Layer 2 = amber. Use it for page titles/accents so the layer is recognisable at a glance.

Source copy blocks (use verbatim):

- Pages 1–3 header line: `Layer 1 — official statistics. Accessed 2026-11-12.`
- Page 1: `Sources: METI Production Dynamics Statistics via JCIA; METI committee material. Basis: factory shipment value — NOT retail market size.`
- Page 2: `Sources: MOF Trade Statistics of Japan via JCIA; Japan Imported Cosmetics Association (FY2024). Customs values: exports FOB, imports CIF. Scope is stated on every visual.`
- Page 3: `Sources: METI Current Survey of Commerce; MIC Family Income and Expenditure Survey (2024).`
- Page 4 header block (mandatory, above all visuals — full text in §Page 4 below).

---

## Page 1: Market overview (Layer 1)

| # | Visual | Binding | Settings |
|---|---|---|---|
| 1-1 | Card ×4 | Shipment Value (JPY bn) @2023; Shipment YoY % (published); Derived — Recovery vs 2019 Peak %; Derived — Gap vs 2019 Peak % | Year context fixed per card (2019/2023). Derived cards carry a "Derived" caption under the value |
| 1-2 | Column chart | X: DimYear (2019, 2023), Y: Shipment Value (JPY bn) | Data labels on. Title: "Domestic factory shipments of cosmetics (JPY bn)" |
| 1-3 | Bar chart (horizontal) | Category Share % (published) by DimCategory, filter CategorySystem = "METI shipment classification" | Descending sort. Title: "Category mix, 2023 (% of factory shipment value)" |

Slicers: none (contexts are fixed; the data has two published years only).
Note text: "Factory shipment value includes goods for export and is not retail market size." / "Fine-item production detail (lotions, essences, sheet masks) was NOT OBTAINED — no fine-item visual is shown." / "Shipments and imports (page 2) are separate series on different bases; never add them."

## Page 2: Trade (Layer 1)

| # | Visual | Binding | Settings |
|---|---|---|---|
| 2-1 | Column chart | FactTradeFlow, filter Flow = Exports; X: DimYear (2014/2021/2023/2024), Y: Trade Value (JPY bn) | Data labels on. Title suffix "(customs value FOB)" |
| 2-2 | Column chart | FactTradeFlow, filter Flow = Imports; X: DimYear (2014/2023/2024), Y: Trade Value (JPY bn) | Title suffix "(customs value CIF) — JCIA aggregate scope" |
| 2-3 | Card | Association-scope 2024 imports total (443.8) + Import YoY % (published) (+17.7) | Caption: "Association scope (excl. toothpaste & soap) — a different scope from visual 2-2" |
| 2-4 | Pie chart | FactImportsByOrigin (2024): Imports by Origin (JPY bn) by DimOrigin | Labels: category + percent. The "All other origins" slice is the Derived row — footnote it |
| 2-5 | Column chart | FactImportsByCategory (2024): Imports by Category (JPY bn) by DimCategory, filter CategorySystem = "Association import categories" | Data labels on; YoY shown in the tooltip via Import Category YoY % (published) |

Slicers: none.
Note text: "Only the years published in the appendix are plotted — no interpolation; do not read the bars as a continuous annual series." / "Two 2024 import totals exist (JCIA aggregate 446.1 bn; association scope 443.8 bn). They differ by scope and must not be mixed." / "The four categories in 2-5 are the categories published by the association; they do not sum to the total."

## Page 3: Channel & household (Layer 1)

| # | Visual | Binding | Settings |
|---|---|---|---|
| 3-1 | Column chart | FactChannel drugstore rows: Channel Sales (JPY bn), categories = Indicator + Period label | Periods differ (full year / H1 / single month) — the axis labels must show the period, and the note below is mandatory. Title: "Drugstore sales — published reference points (JPY bn; periods differ)" |
| 3-2 | Card | Derived — Beauty Care Share of Drugstore Sales (Sep 2024) % (13.2) | Caption: "Derived: 96.0 / 728.8 (Sep 2024 only)" |
| 3-3 | Card | Household Expenditure (JPY) @2024 = 300,243 | Caption: "All items, two-or-more-person households, monthly average (FIES). Context only — NOT cosmetics spending" |
| 3-4 | Text box | Household gap statement (copy below) | — |
| 3-5 | Card | Inbound travel consumption (all purposes), 2026 Q3 = 2,207.6 JPY bn, +3.2% | Caption: "Preliminary survey estimate. The survey's cosmetics & perfumes item by nationality was NOT extracted" |

Copy for 3-4 (verbatim): `FIES cosmetics item codes (2025 revision): 909 cream / 910 lotion / 911 foundation / 912 lipstick / 913 other cosmetics / 914 emulsion. Item-level cosmetics expenditure time series: NOT OBTAINED in this production run — no chart is shown, and no substitute figures are used.`
Note text: "No official cosmetics × e-commerce sales series was confirmed to exist. E-commerce evidence in this package is Layer 2 only (page 4)." / "'Beauty care' is the Current Survey of Commerce drugstore category (cosmetics & sundries) — broader than cosmetics alone."

## Page 4: Consumer voice (Layer 2)

Mandatory header block (verbatim, above all visuals):

> **SAMPLE DATA — READ FIRST.** N = 32 product records, manually sampled from public pages across 4 frames (@cosme review ranking — lotions; @cosme Shopping bestsellers — lotions; Rakuten Ichiba daily ranking — skincare, data date 2026-11-10; Amazon.co.jp bestsellers — lotions, ARCHIVED page copy dated 2024-04-25, structural reference only). Accessed 2026-11-12. This is a sample, not representative of all Japanese consumers. Layer 2 is never used to size the market — market size comes from Layer 1 only (pages 1–3).

| # | Visual | Binding | Settings |
|---|---|---|---|
| 4-1 | Table | Sample frames: Platform / Frame / Sample N / period | Must sum to N = 32 (9 / 4 / 10 / 9) |
| 4-2 | Column chart | @cosme lotion frame: Avg Rating (single platform only) by product (rank order) | Subtitle must state "Rating, max 7". DECORTE 5.8 is the sample maximum |
| 4-3 | Column chart | @cosme lotion frame: Total Reviews in Sample by product | SK-II 40,386 is the frame maximum |
| 4-4 | Bar chart (horizontal) | Rakuten frame: Total Reviews in Sample by product (rank order) | Review counts only — Rakuten ratings were not captured and must not be charted |
| 4-5 | Table | Largest review bases, all frames (top 8) with Platform shown per row | Attenir 50,283 (Rakuten) at top; Amazon rows must carry the "archived 2024-04-25" flag from DimPlatform |

Slicers: Platform/Frame (page-level only). The rating visual must keep a single-platform filter at all times — if the slicer would mix platforms, Avg Rating returns BLANK by design (see measures.dax); do not override.
Note text: "Rating scales differ by platform (@cosme max 7; Amazon.co.jp max 5). Cross-platform rating comparison is invalid." / "Review counts are cumulative platform counts at access date — evidence of incumbents' review moats, not sales figures." / "Rank gaps are as captured (@cosme #3 and Rakuten #5/#10 were not legible/confirmed in sampling — see appendix consumer_voice notes). Ranks are never renumbered."

---

## Verification checklist (cross-check against the Excel companion before publishing)

- [ ] Shipments: 2019 = 1,761.1 / 2023 = 1,302.4 JPY bn, YoY +2.9%; Recovery (Derived) = 73.9%; Gap ≈ 26.1%
- [ ] Category shares 2023: Skin care 43.6 / Hair care 28.0 / Makeup 20.7 / Special-purpose 7.0 / Fragrances 0.7 (sum 100.0)
- [ ] Exports: 2014 = 152.3 / 2021 = 790.6 / 2023 = 608.3 / 2024 = 520.1 JPY bn
- [ ] Imports (JCIA aggregate): 2014 = 217.1 / 2023 = 377.1 / 2024 = 446.1 JPY bn
- [ ] Imports (association scope) 2024 = 443.8 JPY bn (+17.7%); South Korea 134.3 (+40.0%, share 30.3%); France 108.3 (+19.4%, share 24.4%); All other origins (Derived) 201.2 (45.3%)
- [ ] Import categories 2024: Base makeup & skin care 200.6 (+20.1) / Colour makeup 73.0 (+29.8) / Hair care 59.8 (+0.7) / Fragrance 47.5 (+21.1)
- [ ] Drugstores: total 2024 = 8,919.9 JPY bn (+6.9%); beauty care 2024 H1 = 547.2; beauty care Sep 2024 = 96.0 (+12.6%); Derived beauty-care share Sep 2024 = 13.2%
- [ ] Household: all-items monthly expenditure 2024 = JPY 300,243; cosmetics item series shown as NOT OBTAINED (no chart)
- [ ] Inbound: 2026 Q3 total = 2,207.6 JPY bn (+3.2%), labelled preliminary, total only
- [ ] Layer 2: Sample N = 32 (frames 9/4/10/9); @cosme top rating DECORTE 5.8/7; largest review counts: Attenir 50,283 / Yunth 45,427 / SK-II 40,386
- [ ] No visual anywhere combines Layer 1 and Layer 2; every page shows its source copy block; page 4 shows the mandatory sample block
