# Japan Cosmetics & Skincare Market Entry Audit — Sample Report

A go / conditional-go / no-go market entry audit for a foreign beauty brand considering Japan, built on **Japanese government statistics (Layer 1)** plus **structured consumer-voice evidence (Layer 2)** — delivered in English with a full data appendix.

> **Sample report.** Prepared with a hypothetical entrant profile for demonstration; all market figures are from the real public sources cited. Not legal, regulatory, tax, or investment advice.

## Verdict (sample)

**CONDITIONAL GO** — enter as an e-commerce-first, ingredient-led skincare play in the JPY 3,000–6,000 band, with drugstore expansion as phase two. Not as department-store prestige; not leading with colour cosmetics.

Key evidence (all sourced in the report and appendix):
- Domestic factory shipments of cosmetics: JPY 1,302.4 bn in 2023 (+2.9% YoY), still ≈26% below the 2019 peak of JPY 1,761.1 bn (METI Production Dynamics Statistics via JCIA). Skin care is 43.6% of shipment value.
- Cosmetics imports rose 17.7% in 2024; South Korea is the No. 1 origin (+40.0%) — foreign brands are gaining share now (MOF Trade Statistics; Japan Imported Cosmetics Association).
- Drugstore sales grew 6.9% in 2024 to JPY 8,919.9 bn, with beauty care the second-largest growth contributor (METI Current Survey of Commerce).
- Consumer samples (@cosme / Rakuten / Amazon.co.jp, manually sampled 12 Nov 2026) show ingredient-led mid-price products dominating EC rankings — and Korean entrants already charting in Rakuten's daily top 10.

## Dashboard companion (Excel + Power BI)

The audit's data also ships as a dashboard, in two forms built from the same appendix:

- **Excel (finished product)** — `data/cosmetics_audit_dashboard.xlsx`: 4 dashboard sheets (1 Market overview / 2 Trade / 3 Channel & household / 4 Consumer voice) with native charts, plus the appendix data sheets and a definitions & sources sheet. Layer 1 pages use official statistics only; the Layer 2 page is headed by a sample disclaimer (N = 32, manual samples, not representative of all Japanese consumers) and is never used to size the market. Every figure was cross-checked against the appendix; values computed from appendix figures are labelled "Derived" with their formula.
- **Power BI design pack** — `powerbi/`: `model.md` (star schema per layer; deliberately **no** relationships between Layer 1 and Layer 2), `measures.dax` (measure-branching; published YoY/shares are stored columns, never recalculated across the sparse published years; computed measures carry a "Derived —" prefix), `pages.md` (4-page layout spec with the pre-publish cross-check list). Assemble in Power BI Desktop from the appendix/CSVs — the pack reproduces the Excel values exactly.

## Contents

```
report/   japan_cosmetics_market_entry_audit.pdf   (11 pp., incl. sources & methods)
data/     audit_data_appendix.xlsx                 (8 sheets: market_size, category, trade,
                                                     household, channel, consumer_voice,
                                                     competitors, sources)
          cosmetics_audit_dashboard.xlsx           (Excel dashboard companion, 13 sheets)
          layer1_market_trade.csv                  (Layer 1 figures as CSV)
          layer2_consumer_voice.csv                (Layer 2 raw sample records as CSV)
powerbi/  model.md / measures.dax / pages.md       (Power BI design pack for the dashboard)
scripts/  build_appendix.py                        (rebuilds the Excel appendix from the CSVs)
          build_dashboard.py                       (rebuilds the Excel dashboard from the appendix)
          data_notes.md                            (how every source was obtained, incl. failures)
notes/    glossary.md                              (Japanese statistics / retail terms in English)
```

## Method in one paragraph

Layer 1 figures are transcribed from official statistics without modification; every figure carries its source, its basis (e.g., *factory shipment value is not retail market size*) and its access date. Layer 2 data are manual samples of public platform pages — sample frame, size and date are always stated, platform users are not treated as representative of all Japanese consumers, and review data is never used to size the market. Anything that could not be obtained is written as **NOT OBTAINED** (the report's §8 lists every gap: FIES item-level cosmetics series, production fine items, a cosmetics×e-commerce official series — which does not appear to exist — and more). No figure in this repository is an estimate invented to fill a gap.

## Sources (primary)

- METI — Production Dynamics Statistics; Current Survey of Commerce (https://www.meti.go.jp)
- Ministry of Finance — Trade Statistics of Japan
- MIC — Family Income and Expenditure Survey (https://www.stat.go.jp/data/kakei/)
- Japan Tourism Agency — Inbound Consumption Trend Survey (https://www.mlit.go.jp/kankocho/)
- JCIA — shipment/trade compilations of the above (https://jcia.org)
- Layer 2 platforms: @cosme (https://www.cosme.net), Rakuten Ichiba (https://ranking.rakuten.co.jp), Amazon.co.jp

Full source list with URLs and access dates: report, "Sources & methods"; appendix sheet `sources`.

## Disclosure

Research and drafting for this sample were assisted by AI tools. All figures were checked against the cited sources at the access dates stated (12 November 2026). Japanese-language source names are retained in the appendix so any figure can be traced to its original table.
