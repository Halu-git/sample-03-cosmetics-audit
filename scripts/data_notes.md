# Data acquisition notes — Sample 03 (production run, 2026-11-12)

How each source was actually obtained in this production run, including failures.
Access date for all web sources: 2026-11-12 (JST).

## Layer 1 — official statistics

| Source | Method | Result |
|---|---|---|
| METI Production Dynamics Statistics via JCIA shipment page (jcia.org/user/statistics/shipment) | Page fetch | JCIA site was under maintenance/restructure on the access date; the statistics body did not render. Figures used are the ones verified against the official pages during the design phase (2026-11-12): 2023 shipments JPY 1,302.4 bn (+2.9%), category shares 43.6/28.0/20.7/7.0/0.7. Cross-checked against METI committee material (2023: JPY 1,302.5 bn; 2019 peak JPY 1,761.1 bn). |
| METI committee PDF (001_04_00.pdf) | Direct download attempted | Download returned a JavaScript placeholder (site bot protection), not the PDF. Content relied on the design-phase verification and search-index text of the same document. Recorded honestly here. |
| MOF Trade Statistics (cosmetics) | JCIA trade compilation + industry journal table (Cosmetic Science, Oct 2025, explicitly sourced to MOF Trade Statistics) + Japan Imported Cosmetics Association FY2024 release (trade-press report, 2025-02-10) | Obtained: exports/imports 2014–2024 series; 2024 imports by origin (Korea JPY 134.3 bn, France JPY 108.3 bn) and category. Direct e-Stat file extraction of trade tables was NOT performed in this run. |
| METI Current Survey of Commerce | METI 2024 annual review / H1 review (search-index text of official PDFs) and monthly flash reporting | Obtained: drugstore total 2024 JPY 8,919.9 bn (+6.9%); beauty care H1 2024 JPY 547.2 bn (category mapping cross-checked against monthly figures); Sep 2024 beauty care JPY 96.0 bn (+12.6%). Annual category table in the review PDF had one value dropped in text extraction, so the annual beauty-care value is NOT quoted — only the verified H1 and monthly figures are used. |
| MIC Family Income and Expenditure Survey | stat.go.jp 2024 annual summary + e-Stat dataset metadata | Obtained: 2024 average monthly consumption expenditure (two-or-more-person households) JPY 300,243; cosmetics item codes (909–915 group) confirmed. Item-level cosmetics expenditure series NOT extracted in this run (recorded as NOT OBTAINED in report §8 and appendix). |
| JTA Inbound Consumption Trend Survey | Agency page fetch | Obtained: Q3 2026 total inbound travel consumption JPY 2,207.6 bn (+3.2%, preliminary). Cosmetics & perfumes item values by nationality NOT extracted. |
| Economic Census (2021) | Design-phase confirmation only | Not extracted; known constraint (cosmetics merged with pharmaceuticals at small-category level) carried into report §8. |
| USD/JPY reference | Finnhub/OANDA quote in search results | 159.7 JPY/USD on 2026-11-11; report conversions rounded to JPY 160 = USD 1 and labelled indicative. |

## Layer 2 — consumer voice (manual sampling of public pages; no automated scraping)

| Platform | Method | Result |
|---|---|---|
| @cosme review ranking — lotions | Search-index text of cosme.net category ranking page + product/review pages | Obtained top-10 sample: ratings (7-point), review counts, prices; aggregation period 2026-08-05 to 2026-11-04, page updated 2026-11-06. Rank 3 differs between the category page and product-page sidebars (ELIXIR Brightening Lotion line) — recorded as a methods note in the appendix instead of being resolved by assumption. |
| @cosme Shopping bestsellers — lotions | Search-index text of cosme.com ranking page | Obtained ranks 1–4 with @cosme review scores/counts and prices. |
| Rakuten Ichiba daily ranking — skincare | Page fetch (ranking.rakuten.co.jp/daily/100944) | Obtained, data date 2026-11-10 (page updated 2026-11-11): ranks, review counts, prices for the top band. Displayed rank numbers 5 and 10 were not legible in text view; affected rows are recorded with the ranks that were legible and the gap is noted. Star ratings were not captured in text view (review counts only). |
| Amazon.co.jp bestsellers — lotions | Page fetch | The fetched copy was an ARCHIVED page dated 2024-04-25. Used only as structural evidence of the mass price band, explicitly dated everywhere it appears. No current Amazon ranking is claimed. Individual review-text mining (high/low review reading) was NOT performed in this run; complaint/praise themes in report §6 are qualitative, drawn from sampled product positioning and ranking metadata, and labelled as such. |
| SNS signals (X/Instagram/TikTok) | — | NOT collected (per design fallback: downgraded to qualitative observation; no SNS figures quoted anywhere). |

## Compliance notes
- No automated collection (scraping) was used; sampling was manual reading of public pages via browser fetch/search.
- Review texts were summarised thematically; no review text is reproduced at length.
- No personal data, credentials or API keys are included anywhere in the deliverables. The e-Stat API was not used (no app ID involved).
