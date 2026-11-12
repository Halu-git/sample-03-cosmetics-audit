# Power BI Data Model — Japan Cosmetics & Skincare Market Entry Audit (Dashboard Companion)

- Data: `data/audit_data_appendix.xlsx` (8 sheets) + `data/layer1_market_trade.csv` + `data/layer2_consumer_voice.csv`. All figures accessed 2026-11-12.
- Purpose: reproduce the Excel companion (`data/cosmetics_audit_dashboard.xlsx`) in Power BI Desktop with identical values. This file defines the model; `measures.dax` defines the measures; `pages.md` defines the pages.
- Design policy: star schema per layer. **Layer 1 (official statistics) and Layer 2 (consumer-platform samples) are modelled as two separate islands with NO relationships between them.** Pages are layer-fixed; no visual may combine the two layers. This mirrors the report's rule that Layer 2 is never used to size the market.

---

## 1. Table list

### 1.1 Facts — Layer 1 (official statistics)

| Table | Source | Grain | Notes |
|---|---|---|---|
| FactMarketSize | appendix `market_size` | Indicator × Year | Factory shipments 2019 / 2023 + the published Derived recovery row. Keep the appendix's `Processing` flag on every row. |
| FactCategoryShare | appendix `category` | Category × Year (2023) | Shares of factory shipment value (%). The "NOT OBTAINED" fine-item row is documentation only — exclude it at import (filter in Power Query) and keep the exclusion in the query comment. |
| FactTradeFlow | appendix `trade` (JCIA aggregate rows) | Flow (Exports/Imports) × Year | Sparse published years only (Exports 2014/2021/2023/2024; Imports 2014/2023/2024). Do NOT densify. |
| FactImportsByOrigin | appendix `trade` (association scope rows) | Origin × Year (2024) | Scope = association scope (excl. toothpaste & soap). Includes the Derived "All other origins" row, flagged `IsDerived = TRUE`. |
| FactImportsByCategory | appendix `trade` (association scope rows) | Category × Year (2024) | Four categories as published; they do not sum to the total and no "total" row may be added. |
| FactChannel | appendix `channel` | Indicator × Period | Drugstore rows have mixed periods (full year / H1 / single month) — `Period` is a text dimension, never a date. The e-commerce "NOT CONFIRMED" row is documentation only (exclude at import). Inbound total row included as context. One extra row is added at import: the Sep 2024 all-categories drugstore total (JPY 728.8 bn), which the appendix records in the Basis text of the beauty-care Sep row — import it as its own row with `Processing = "Transcribed (from Basis text)"`. It is needed by the Derived beauty-care share measure. |
| FactHousehold | appendix `household` | Indicator × Year | All-items household expenditure (context) + item-code structure row + NOT OBTAINED row (documentation only, exclude the latter two at import; the gap is stated on the page, not in the data). |

### 1.2 Facts — Layer 2 (consumer-platform samples)

| Table | Source | Grain | Notes |
|---|---|---|---|
| FactConsumerVoice | appendix `consumer_voice` = `layer2_consumer_voice.csv` | One row per sampled product record (N = 32) | Carries Platform, SampleFrame, Rank, Brand, Product, Rating, RatingScaleMax, ReviewCount, PriceMinJPY, PriceMaxJPY, Accessed, Notes. Star ratings for Rakuten rows are blank in the source (not captured) — keep them blank; never impute. |
| CompetitorEvidence | appendix `competitors` | One row per group/brand | Standalone evidence table for a matrix visual only. **No measures, no relationships.** The "Note on company financials" row is documentation (exclude at import). No revenue figures exist in this table by design. |

### 1.3 Dimensions

| Table | Built from | Contents |
|---|---|---|
| DimYear | Distinct years appearing in Layer 1 facts | Sparse year list (2014, 2019, 2021, 2023, 2024). NOT a continuous calendar — see §5. |
| DimFlow | FactTradeFlow | Exports / Imports. |
| DimCategory | FactCategoryShare + FactImportsByCategory | Category name + `CategorySystem` (METI shipment classification vs association import categories). The two systems are different classifications — the `CategorySystem` column is mandatory and pages must filter to one system. |
| DimOrigin | FactImportsByOrigin | South Korea / France / All other origins (Derived). |
| DimPlatform | FactConsumerVoice | Platform + SampleFrame + `RatingScaleMax` (7 for @cosme, 5 for Amazon.co.jp, blank for Rakuten) + `IsArchived` flag (Amazon frame = archived 2024-04-25 copy). |
| DimPeriod | FactChannel | Period text + `PeriodType` (FullYear / HalfYear / SingleMonth / Quarter) so visuals can warn instead of trending. |

---

## 2. Column definitions (key facts)

### FactMarketSize
| Column | Type | Notes |
|---|---|---|
| Indicator | Text | As published (e.g., "Domestic factory shipments of cosmetics (all categories)") |
| YearKey | Integer | → DimYear |
| ValueJPYbn | Decimal | JPY bn. The recovery row stores 73.9 in `ValuePct`, not here |
| ValuePct | Decimal | Only for the Derived recovery row (73.9) |
| YoYPctPublished | Decimal | As published (+2.9 for 2023). Never recalculated |
| Basis | Text | Definition text ("Factory shipment value; NOT retail market size") |
| Processing | Text | Transcribed / Derived — drives the derived-flag logic in measures |

### FactTradeFlow
| Column | Type | Notes |
|---|---|---|
| FlowKey | Text | → DimFlow |
| YearKey | Integer | → DimYear (sparse) |
| ValueJPYbn | Decimal | Customs value (exports FOB, imports CIF) |
| Scope | Text | Fixed "JCIA aggregate". Association-scope rows live in their own facts — the two scopes must never be summed or shown in one visual |

### FactImportsByOrigin / FactImportsByCategory
| Column | Type | Notes |
|---|---|---|
| OriginKey / CategoryKey | Text | → DimOrigin / DimCategory |
| ValueJPYbn | Decimal | 2024, association scope |
| YoYPctPublished | Decimal | As published by the association |
| SharePctPublished | Decimal | Origin fact only (30.3 / 24.4 as published) |
| IsDerived | Boolean | TRUE only for "All other origins" |

### FactConsumerVoice
| Column | Type | Notes |
|---|---|---|
| PlatformKey | Text | → DimPlatform |
| SampleFrame | Text | Frame wording as in the source |
| Rank | Integer | Displayed rank as captured. Gaps are real (e.g., @cosme #3, Rakuten #5/#10 not captured) — do not renumber |
| Rating | Decimal | Blank where not captured (all Rakuten rows) |
| RatingScaleMax | Integer | From DimPlatform; shown in every rating visual's subtitle |
| ReviewCount | Integer | Cumulative platform count at access date. Not a sales figure |
| PriceMinJPY / PriceMaxJPY | Integer | Split from the source price text at import ("12650-29150" → 12650 / 29150; single price → both equal). The original text stays in `PriceAsDisplayed` |

---

## 3. Relationships

All single-direction (dimension → fact, one-to-many), **within Layer 1 only, plus DimPlatform → FactConsumerVoice within Layer 2**:

```
DimYear    ──< FactMarketSize / FactCategoryShare / FactTradeFlow / FactImportsByOrigin / FactImportsByCategory / FactHousehold
DimFlow    ──< FactTradeFlow
DimCategory──< FactCategoryShare / FactImportsByCategory   (filter by CategorySystem per page)
DimOrigin  ──< FactImportsByOrigin
DimPeriod  ──< FactChannel          (FactChannel has no DimYear link: periods are mixed grains)
DimPlatform──< FactConsumerVoice
CompetitorEvidence: no relationships (standalone)
```

Rules:
- **No relationship, filter path, or shared slicer may connect Layer 1 and Layer 2.** Report-level slicers are prohibited; slicers are page-level only.
- No bi-directional filters.
- DimYear is sparse: do not mark it as a date table, and do not use time-intelligence functions (DATEADD, SAMEPERIODLASTYEAR, …) anywhere. Published YoY is a stored column.

---

## 4. Import steps (Power Query)

1. Get Data > Excel: load the 8 appendix sheets; Get Data > Text/CSV: load the two CSVs as a cross-check source (the appendix is authoritative; the CSVs must agree with it).
2. Per fact, rename columns to §2, set types, and apply the documented exclusions (NOT OBTAINED / documentation rows) as explicit filter steps with comments — never silent row deletion. FactChannel additionally gains the Sep 2024 total row described in §1.1; FactHousehold's value text ("300,243 JPY") is parsed into the numeric column `ValueJPY`.
3. Add `Processing` / `IsDerived` flags exactly as published in the appendix; do not recompute the appendix's Derived rows (recovery 73.9%) at import — they are data.
4. FactConsumerVoice: split the price text on "–"/"-" into PriceMinJPY / PriceMaxJPY (Text.Select digits), keep the original in `PriceAsDisplayed`; convert blank ratings to null (not 0).
5. Build dimensions as reference queries with duplicates removed (DimYear/DimPeriod are curated lists per §1.3, not generated calendars).
6. Create the relationships in §3, hide fact key columns, register the measures from `measures.dax` in per-fact measure tables.
7. Lay out the 4 pages per `pages.md` and run its verification checklist against the Excel companion before publishing.

## 5. Model caveats (premises)

- **Sparse years.** Trade/shipment facts contain only the years published in the appendix. A line chart implying continuity is a modelling error: use column charts with the year as a category, and the pages.md note text ("published years only — no interpolation").
- **Two import scopes.** JCIA aggregate (446.1 bn, 2024) and association scope (443.8 bn, 2024) coexist in the model in separate facts. Any visual must state its scope; `Scope` must appear in the visual title or subtitle (pages.md copy blocks do this).
- **Ratings are not comparable across platforms.** Different scales (7 vs 5) and one frame lacks ratings entirely. DimPlatform.RatingScaleMax exists so every rating visual can display its scale; cross-platform rating aggregation is prohibited in measures (enforced by making the average-rating measure require a single-platform filter — see measures.dax).
- **Derived vs published.** Anything computed in DAX is named with a `Derived —` prefix and listed in the measures header. Published YoY and published shares are columns, not measures.
