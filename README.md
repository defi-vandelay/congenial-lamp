# congenial-lamp

## this repo and project was created to test idea but was unsuccessful. Do not recommend using.

# Prompt for Excelt Create

Build a new Excel 365 workbook template for **Congenial-lamp v2 account analysis**.

This workbook is for **intelligenct analysis**, not court exhibit use. It must be **macro-free**, **Power Query free**, and **safe for manual paste-as-values workflows**. Do not use VBA, Office Scripts, external connections, linked workbooks, embedded web data, or macros.

## Primary objective

Create a workbook template where downloaded CSV exports from a Congenial-lamp v2 web app can be opened, copied, and pasted **as values** directly into specific workbook sheets. The pasted data must fit cleanly into predefined paste zones with no structural friction.

The workbook must separate:
1. pasted source data
2. lightweight sheet-level helper logic
3. derived analysis sheets
4. visual summary sheets

## General workbook rules

- Excel 365 formulas only
- No macros
- No Power Query
- No external links
- No merged cells in any paste zone
- No formulas inside any canonical paste block
- All paste zones must start at `A1`
- Row 1 of each paste sheet must contain the exact CSV headers in the exact order given below
- All workbook-only helper columns must sit to the right of the canonical pasted CSV columns
- Format pasted data sheets as Excel Tables
- Use structured references where practical
- Freeze top row on all paste sheets
- Use filters on all table headers
- Use consistent professional formatting suitable for internal intelligence work
- Use readable fonts and moderate column widths
- Do not add decorative elements that interfere with data handling

## Required sheet order

Create sheets in this exact order:

1. `README`
2. `Asset_Map`
3. `Event_Log`
4. `Position_Snapshots`
5. `Protocol_Params`
6. `Risk_Summary`
7. `Tx_Final_View`
8. `Liquidation_Analysis`
9. `Charts`
10. `Case_Summary`
11. `QA_Checks`

---

## Sheet-by-sheet instructions

### 1. README

Purpose:
- workbook instructions
- workflow notes
- embedded manifest paste section
- top-level case metadata display

Create a visible instructions area at the top explaining:
- paste `export_manifest.csv` into the manifest block
- paste `event_log.csv` into `Event_Log`
- paste `position_snapshots.csv` into `Position_Snapshots`
- paste `protocol_params.csv` into `Protocol_Params`
- paste `risk_summary.csv` into `Risk_Summary`

Reserve range `A20:P21` for the manifest paste block and convert it into a table named `tblManifest`.

Use these exact headers in row 20:

- `schema_version`
- `export_generated_utc`
- `app_version`
- `protocol`
- `network`
- `chain_id`
- `subject_address`
- `start_block`
- `end_block`
- `start_datetime_utc`
- `end_datetime_utc`
- `observed_markets`
- `observed_assets`
- `event_count`
- `tx_count`

Create a metadata display block near the top with formulas that pull from `tblManifest`:

- Subject Address
- Export Generated UTC
- Observed Markets
- Observed Assets
- Event Count
- Tx Count
- Review Window

Use formulas equivalent to:
- `=IFERROR(INDEX(tblManifest[subject_address],1),"")`
- `=IFERROR(INDEX(tblManifest[export_generated_utc],1),"")`
- `=IFERROR(INDEX(tblManifest[observed_markets],1),"")`
- `=IFERROR(INDEX(tblManifest[observed_assets],1),"")`
- `=IFERROR(INDEX(tblManifest[event_count],1),"")`
- `=IFERROR(INDEX(tblManifest[tx_count],1),"")`
- `=IFERROR(INDEX(tblManifest[start_datetime_utc],1)&" to "&INDEX(tblManifest[end_datetime_utc],1),"")`

### 2. Asset_Map

Create a static reference sheet with a table named `tblAssetMap`.

Use columns:

- `asset`
- `ctoken_symbol`
- `market_contract`
- `underlying_contract`
- `underlying_decimals`
- `asset_type`
- `display_order`
- `notes`

Populate a few example rows for:
- ETH / cETH
- SAI / cSAI
- DAI / cDAI
- USDC / cUSDC
- BAT / cBAT

This sheet is a reference sheet only and does not need a CSV paste block.

### 3. Event_Log

This is a direct paste sheet for `event_log.csv`.

Create the canonical paste block starting at `A1` and convert it into a table named `tblEventLog`.

Use these exact columns in this exact order:

- `event_num`
- `txn_hash`
- `tx_event_order`
- `is_tx_final`
- `datetime_utc`
- `block_number`
- `transaction_index`
- `log_index`
- `trace_index`
- `event_type`
- `market_contract`
- `ctoken_symbol`
- `underlying_contract`
- `asset`
- `counterparty`
- `transfer_direction`
- `amount_ctoken`
- `amount_underlying`
- `amount_usd`
- `affects_supply_position`
- `affects_borrow_position`

Add helper columns to the right, starting immediately after the canonical columns:

- `protocol_action_class`
- `tx_row_count`
- `is_transfer_row`
- `is_liquidation_row`
- `related_party_flag`
- `investigator_notes`

Add formulas for the helper columns only:
- classify collateral movement vs debt movement vs eligibility change
- count rows per tx hash
- flag cToken transfer rows
- flag liquidation rows

Do not put formulas inside the canonical pasted columns.

### 4. Position_Snapshots

This is a direct paste sheet for `position_snapshots.csv`.

Create the canonical paste block starting at `A1` and convert it into a table named `tblPositionSnapshots`.

Use these exact columns in this exact order:

- `event_num`
- `txn_hash`
- `tx_event_order`
- `is_tx_final`
- `datetime_utc`
- `block_number`
- `market_contract`
- `ctoken_symbol`
- `underlying_contract`
- `asset`
- `entered_market`
- `ctoken_balance`
- `supplied_underlying`
- `supplied_value_usd`
- `borrow_balance`
- `borrow_value_usd`
- `oracle_price_usd`
- `exchange_rate`
- `borrow_index`
- `collateral_factor`
- `adj_collateral_value_usd`

Add helper columns to the right:

- `has_supply`
- `has_borrow`
- `contributes_collateral`
- `net_market_equity_usd`
- `snapshot_notes`

Use formulas only in helper columns.

### 5. Protocol_Params

This is a direct paste sheet for `protocol_params.csv`.

Create the canonical paste block starting at `A1` and convert it into a table named `tblProtocolParams`.

Use these exact columns in this exact order:

- `event_num`
- `txn_hash`
- `tx_event_order`
- `is_tx_final`
- `datetime_utc`
- `block_number`
- `market_contract`
- `ctoken_symbol`
- `underlying_contract`
- `asset`
- `oracle_price_usd`
- `exchange_rate`
- `borrow_index`
- `collateral_factor`
- `close_factor`
- `liquidation_incentive`
- `reserve_factor`
- `borrow_rate_per_block`
- `supply_rate_per_block`
- `total_cash`
- `total_borrows`
- `total_reserves`
- `accrual_block_number`

Add helper columns to the right:

- `cf_changed`
- `price_changed`
- `rate_changed`
- `cash_shift_flag`
- `governance_or_market_note`
- `param_notes`

Use formulas only in helper columns.

### 6. Risk_Summary

This is a direct paste sheet for `risk_summary.csv`.

Create the canonical paste block starting at `A1` and convert it into a table named `tblRiskSummary`.

Use these exact columns in this exact order:

- `event_num`
- `txn_hash`
- `tx_event_order`
- `is_tx_final`
- `datetime_utc`
- `block_number`
- `event_type`
- `event_asset`
- `borrow_limit_usd`
- `total_borrow_usd`
- `liquidity_cushion_usd`
- `health_ratio`
- `utilization_pct`
- `liquidation_margin_pct`
- `active_collateral_markets`
- `active_borrow_markets`
- `risk_band`

Add helper columns to the right:

- `risk_direction`
- `material_change_flag`
- `tx_final_flag_check`
- `narrative_seed`
- `review_notes`

Use formulas only in helper columns.

Suggested logic:
- classify whether risk increased, reduced, or was mixed
- flag material movements
- seed a short narrative like `Borrow / DAI / Warning`

### 7. Tx_Final_View

This is a derived sheet only. No manual paste.

Purpose:
- one row per transaction final state
- clean tx-level analysis
- source for charts and summary

Build this sheet from:
- `tblRiskSummary`
- `tblEventLog`

Create columns:

- `event_num`
- `txn_hash`
- `datetime_utc`
- `block_number`
- `event_type`
- `event_asset`
- `borrow_limit_usd`
- `total_borrow_usd`
- `liquidity_cushion_usd`
- `health_ratio`
- `utilization_pct`
- `liquidation_margin_pct`
- `active_collateral_markets`
- `active_borrow_markets`
- `risk_band`
- `tx_event_count`
- `tx_event_types`
- `tx_assets`
- `delta_borrow_limit_usd`
- `delta_total_borrow_usd`
- `delta_liquidity_cushion_usd`
- `delta_health_ratio`
- `delta_utilization_pct`
- `primary_driver`
- `risk_direction`
- `review_priority`

Use dynamic formulas that:
- filter `tblRiskSummary` where `is_tx_final="TRUE"`
- count event rows per tx
- join distinct event types per tx
- join distinct assets per tx
- calculate deltas from prior tx-final row
- classify primary driver
- classify risk direction
- assign review priority

This should be the cleanest investigator-facing sheet in the workbook.

### 8. Liquidation_Analysis

Derived sheet only. No manual paste.

Purpose:
- isolate all liquidation events
- help assess economic implications and related-party theories

Build this sheet from:
- `tblEventLog`
- `tblProtocolParams`
- `tblRiskSummary`

Create columns:

- `event_num`
- `txn_hash`
- `datetime_utc`
- `debt_asset`
- `debt_repaid_underlying`
- `debt_repaid_usd`
- `close_factor`
- `liquidation_incentive`
- `post_borrow_limit_usd`
- `post_total_borrow_usd`
- `post_liquidity_cushion_usd`
- `post_risk_band`
- `within_tx_event_count`
- `related_party_flag`
- `notes`

Use formulas to:
- filter `LiquidateBorrow` rows
- pull related account-level risk metrics by `event_num`
- count event rows within each liquidation tx

Leave `related_party_flag` and `notes` as manual analyst fields.

### 9. Charts

Derived sheet only. No manual paste.

Create at least these charts:

1. **Borrow limit vs total borrow over time**
   - source: `Tx_Final_View`
   - x-axis: `datetime_utc`
   - series: `borrow_limit_usd`, `total_borrow_usd`

2. **Utilization % over time**
   - source: `Tx_Final_View`
   - x-axis: `datetime_utc`
   - series: `utilization_pct`

3. Optional: **Liquidation event values**
   - source: `Liquidation_Analysis`
   - x-axis: `datetime_utc`
   - series: `debt_repaid_usd`

Keep the chart sheet clean and readable.

### 10. Case_Summary

Derived and partly manual.

Purpose:
- top-level case overview
- key metrics
- working narrative
- theory-testing notes

Create a metadata block pulling from `tblManifest`:
- subject address
- export generated utc
- observed markets
- observed assets

Create KPI cells for:
- max total borrow
- lowest liquidity cushion
- highest utilization
- liquidation count
- count of `High Risk` or `LIQUIDATABLE` tx-final states

Create manual commentary sections for:
- likely misuse indicators
- likely incompetence indicators
- liquidation observations
- key tx hashes for review
- unresolved issues

Use formulas to populate the KPI block from `Tx_Final_View` and `Liquidation_Analysis`.

### 11. QA_Checks

Derived sheet only. No manual paste.

Purpose:
- quick integrity check after each paste cycle

Create checks for:

- manifest event count matches `tblEventLog` row count
- manifest tx count matches unique `txn_hash` count
- count of `is_tx_final="TRUE"` in `tblEventLog` matches unique tx count
- `tblRiskSummary` row count matches `tblEventLog` row count
- `tblPositionSnapshots` has rows
- `tblProtocolParams` has rows
- SAI exists when expected
- cToken transfer events exist when expected

Display each result as `PASS`, `FAIL`, or `CHECK`.

Use simple formulas only.

## Formatting instructions

Apply consistent formatting across the workbook:

- professional internal-analyst style
- no bright colors
- use subtle header fill and borders
- freeze top row on all data sheets
- enable filters on all tables
- set widths so hashes, addresses, and timestamps are visible
- apply text format to hash, address, and identifier columns where needed
- use number formats suitable for high precision crypto values
- do not convert hashes or addresses to scientific notation
- do not add conditional formatting that slows down large sheets unless it is minimal and useful

## Formula guidance

Use Excel formulas only. No macros.

Use dynamic array formulas where appropriate, especially on:
- `Tx_Final_View`
- `Liquidation_Analysis`

Use structured references where practical.

Do not place spill formulas inside paste zones.

Keep all heavy logic off the direct paste sheets.

## Build outcome required

The final workbook must work like this:

1. paste manifest CSV into `README` manifest block
2. paste `event_log.csv` into `Event_Log` starting at `A1`
3. paste `position_snapshots.csv` into `Position_Snapshots` starting at `A1`
4. paste `protocol_params.csv` into `Protocol_Params` starting at `A1`
5. paste `risk_summary.csv` into `Risk_Summary` starting at `A1`
6. derived sheets update automatically
7. charts update automatically
8. QA checks show whether the paste cycle is internally consistent

Build the workbook template now according to these instructions.

## After Claude finishes

Use this quick review checklist inside Excel:

1. Confirm each paste sheet starts at `A1` with exact headers.
2. Confirm no formulas exist inside the canonical paste blocks.
3. Confirm helper columns begin immediately to the right of each paste block.
4. Confirm all direct paste blocks are formatted as tables:
   - `tblEventLog`
   - `tblPositionSnapshots`
   - `tblProtocolParams`
   - `tblRiskSummary`
   - `tblManifest`
   - `tblAssetMap`
5. Confirm `Tx_Final_View`, `Liquidation_Analysis`, `Charts`, `Case_Summary`, and `QA_Checks` update automatically after paste.
6. Confirm the workbook remains macro-free.
