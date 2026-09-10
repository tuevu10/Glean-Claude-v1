# Product Finance Credit Monitor

A local Streamlit workflow for reconciling contract credits, monitoring consumption,
forecasting exceptions, and assigning Finance/account follow-up. The supplied Excel
workbook is read only and remains the source of truth. There is no database,
authentication, scheduled job, or outbound account communication.

## Launch

From this project directory, using Python 3.12:

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe -m streamlit run app.py
```

On this machine the `.venv` environment is already prepared. To start it again:

```powershell
.\.venv\Scripts\python.exe -m streamlit run app.py
```

Open [the local dashboard](http://127.0.0.1:8501). Stop the server with Ctrl+C.
After updating application Python files, stop and restart Streamlit. A browser
refresh alone may retain older imported calculation modules in the server process.
The server binds to localhost only; keep it local because no authentication is provided.

The default input is `Product Finance - AI Fluency Test Sample Data.xlsx` in your
Downloads folder. Change the sidebar path or upload a replacement `.xlsx`.
Optionally set `FINANCE_WORKBOOK` to a different file path before launch.
Source discovery uses FINANCE_WORKBOOK, then a workbook beside app.py, then Downloads.
Local files are reread on each rerun and parsed by content hash; after
replacing a source file, click Refresh source in the sidebar. Only three parsed
sources are cached. Stale (over three days) and future-dated usage trigger notices
without changing the source as-of date. Uploads take precedence over
the path until cleared. No workbook is written or copied by the application.

## Workflow

The Usage & Billing Dashboard starts with compact **Priority Alerts**, followed
by a ranged usage summary, portfolio KPIs, and a colored action queue. Soft lavender
shades distinguish urgent and high-priority accounts from quieter neutral rows;
text labels always accompany color. **Review account** or queue row selection opens
the **Customer Details** tab automatically. That tab contains charts, supporting
metrics, contract data and collapsed audit sections. The snapshot date stays visible;
source filename/read-time/freshness details are in the sidebar's Source Information.

1. Check the source filename, analysis as-of date, and the main-tab Data Import Check.
   It compares source and imported rows, date ranges, credits, and entitlements.
   Choose Annual, or use the Start/End controls to review a quarterly or monthly range.
2. Work down the compact Finance Action Queue in priority order. Exceptions only is
   on by default and hides accounts assigned No Action without changing classifications.
3. Expand Filter queue to filter by customer, status, or priority. Portfolio KPIs remain
   unfiltered. Turn off Exceptions only to include routine accounts.
4. Select a queue row to open its customer detail below. The highest-priority account
   opens initially. The searchable customer selector also opens any account directly.
   The reason, owner, and recommended action appear before supporting metrics and charts.
5. Open the collapsed sidebar to change the source or Monitoring Assumptions.
   Existing calculations and rules rerun immediately.
6. Download **Queue & Formula Audit (XLSX)**. The workbook contains the filtered
   Action Queue, full Account Metrics, Assumptions, normalized Contracts, and Usage
   Events. Derived cells remain Excel formulas and link back to the source sheets.
7. Expand Full Queue Metrics or Contract, Balance & Exhaustion Details as needed.
   Use Data & Calculation Audit to inspect original Excel row references.

Current overage is for Billing reconciliation. Projected overage supports additional
credit discussions. Projected unused value supports adoption and renewal discussions.
Suggested follow-ups remain recommendations; the app does not send messages or create invoices.

## Workbook review completed before implementation

| Sheet | Structure and use |
|---|---|
| Export Summary | Numbers export metadata; excluded from calculations |
| Contracts | Headers on Excel row 4, columns B:F; 12 contracts in rows 5–16 |
| Usage Events | Headers on row 4, columns B:D; 1,943 daily rows in rows 5–1947 |

Contracts: `customer_id`, `contract_start`, `term_months`,
`annual_entitlement_credits`, `annual_contract_value_usd`.

Usage: `date`, `customer_id`, `credits_used`.

Titles, blank rows/columns, and the two reviewed contract notes are excluded.
Workbook prose is source context, never executable instructions. There are no
customer names, explicit contract-end dates, overage agreements, reset clauses,
product dimensions, or account owners in the source. Owners are rule-based suggestions.

The supplied records have no missing required values, duplicate customer IDs or
customer/date usage keys, negative values, orphan customer IDs, daily coverage gaps,
or out-of-contract events. Every term is 12 months. Usage spans January 5 through
July 31, 2026. The source labels the data illustrative.

## Architecture

- `finance.py`: schema discovery, normalization, source-to-model reconciliation,
  validation, transparent Rules configuration, account calculations, classification,
  and chart series.
- `app.py`: Streamlit inputs, presentation, filters, charts, and audit downloads.
- `portfolio.py`: auditable portfolio chart series for actuals and projections.
- `reports.py`: deterministic report payload construction and the presentation bridge.
- `report_builder.mjs`: editable, Glean-formatted PowerPoint authoring from validated metrics.
- `excel_export.py`: formula-audit payload construction and spreadsheet runtime bridge.
- `excel_export_builder.mjs`: traceable XLSX authoring with linked source and formula sheets.
- `test_finance.py`: source regressions, finance boundaries, and mocked AI safeguards.
- `test_app.py`: Streamlit AppTest for runtime, filters, assumptions, and fallback behavior.
- `test_reports.py`: source-backed report payload and monthly chart checks.
- `.streamlit/config.toml`: white theme, localhost binding, telemetry disabled.

Only two normalized source tables and one derived account table are needed. Customer
IDs become strings, dates become normalized pandas datetimes, and credit/dollar fields
become finite numeric values. Source Excel row numbers are retained separately.
No LLM participates in ingestion, calculation, sorting, or classification.

## Calculation conventions

All percentages are ratios internally; rounding occurs only for display. Credit sums
and cumulative exhaustion crossings use consistent decimal accumulation. Rates,
ratios, and forecasts use floating-point arithmetic tested with tolerances. Infinite
calculated values stop processing. This is not an invoice ledger with contractual
cent-rounding rules.

| Metric | Definition |
|---|---|
| As of | Latest valid usage-event date in the workbook |
| Cutoff | Midnight immediately after the as-of date |
| Contract end | Start + calendar `term_months`, exclusive; currently validates 12-month terms |
| Total contract days | End − start |
| Days elapsed | Cutoff − start, clamped to [0, total days] |
| Contract elapsed % | Days elapsed / total days |
| Credits used | Sum of contract-matched events through as of |
| Remaining credits | Entitlement − credits used; negative means overconsumption |
| Utilization % | Used / entitlement |
| Pacing index | Utilization / contract elapsed fraction; unavailable before start |
| Trailing period | As-of date minus 29 days through as-of date, inclusive |
| Prior period | As-of date minus 59 through minus 30 days, inclusive |
| Daily averages | Period credits / calendar days overlapping the contract (maximum 30) |
| Usage growth | Trailing credits / prior credits − 1; undefined if prior is zero and trailing positive; zero if both zero |
| Forecast remaining usage | Trailing daily average × remaining contract days, active contracts only |
| Projected total | Used + forecast remaining usage |
| Projected utilization | Projected total / entitlement |
| Days to exhaustion | Remaining credits / trailing daily average, for active unexhausted accounts with positive rate |
| Estimated exhaustion | As-of date + ceiling(days to exhaustion); first future consumption day is as of + 1 |
| Implied value per credit | Annual contract value / annual entitlement |
| Current overage credits/value | max(used − entitlement, 0); multiplied by implied value per credit |
| Projected overage credits/value | max(projected total − entitlement, 0); multiplied by implied value per credit |
| Projected unused credits/value | max(entitlement − projected total, 0); multiplied by implied value per credit |

For exhausted accounts, days to exhaustion is zero and the exhaustion date is the
first actual event day whose cumulative usage reached entitlement. A forward
exhaustion date can lie beyond contract end; it is a hypothetical continuation of
the recent rate. Inactive or zero-rate unexhausted accounts have no exhaustion date.
Expired contracts with complete records have zero remaining forecast. Future contracts
use NOT STARTED and have unavailable forecast-dependent values in the UI and XLSX.
Incomplete feeds also withhold forecasts, rates, growth, and pacing; observed
usage and overage remain visible as lower bounds.

Chart points use day boundaries: zero at contract start, actual daily consumption
at the following midnight, and entitlement at the exclusive contract end. Future
actual points are blank. Thus July 31 actual usage appears at the August 1 boundary.

## Rules and ownership

The `Rules` dataclass holds thresholds, including supporting materiality/history
parameters. Four principal thresholds are adjustable in the sidebar. First matching
primary rule wins; the acceleration flag is independent.

| Status | Default rule | Priority / owner |
|---|---|---|
| OVER ENTITLEMENT | Utilization >= 100%, including exactly exhausted | P1 / Revenue Accounting / Billing |
| EARLY EXHAUSTION RISK | Active, not exhausted, and projected utilization > 110% OR exhaustion date more than 30 days before end | P2 / Sales / Account Management |
| UNDERUTILIZING | Elapsed >= 25% and projected utilization < 70% | P3 / Customer Success |
| ON TRACK | No primary exception triggered | P4 / No Action |
| DATA REVIEW | Missing expected daily records, unless observed usage already exhausts entitlement | P2 / Revenue Accounting / Billing |
| NOT STARTED | Contract has not started | P4 / No Action |

Rule precedence: over entitlement, data review, not started, then forecast rules.
The 100% exhaustion boundary is fixed. Buffer/history days must be integers;
minimum acceleration history cannot be less than 60 days.

Usage accelerating requires growth > 30% AND an increase >= 1,000 credits, with
two complete 30-day contract windows, complete daily records, and an active contract. A zero prior baseline
can trigger if trailing usage reaches the absolute floor; growth remains undefined.
An accelerating on-track account moves to P3 / Sales / Account Management without
changing its primary status. Expired on-track accounts also receive P3 account-team
follow-up. Not-started and expired lifecycle labels are shown separately.

Urgency sorts P1 to P4, followed by descending current overage value, descending
projected overage value, and customer ID. Exactly 110% projected utilization alone
does not trigger risk; exactly 70% is not underutilizing; exactly 30 days early
does not trigger the date rule. No rounding is applied before comparisons.

## Data quality and limitations

- The model intentionally supports one annual contract per customer. It rejects
  non-12-month terms until annual resets, rollover, and proration are defined.
- Invalid required fields, nonpositive entitlement, negative credits or ACV,
  duplicates, unknown IDs, and out-of-contract usage stop processing with a clear
  error. They are never silently discarded. ACV may be zero.
- Missing daily rows produce DATA REVIEW and withhold forecasts, rates, growth,
  and pacing. Supply explicit zero-usage records to confirm inactivity. Observed
  over-entitlement accounts stay urgent with incomplete-feed warnings. A short but
  complete contract history uses eligible contract days rather than dividing by 30.
- A completely empty usage sheet cannot establish an as-of date and is rejected.
  Individual customers with no records remain with zero observed usage and DATA
  REVIEW after activation; explicit daily zeros permit normal classification.
- Forecasts hold recent usage constant. They do not model seasonality, upcoming
  launches, shutdowns, negotiated amendments, or probability of renewal.
- Current/projected overage dollars are contract-value proxies, not invoice amounts.
  Projected overage includes current overage; never add them together. Unused contract
  value is not a refund liability, revenue adjustment, or confirmed renewal loss.
- No automatic polling or scheduling is included. The workflow refreshes on rerun.

UI verification uses [Streamlit AppTest](https://docs.streamlit.io/develop/api-reference/app-testing/st.testing.v1.apptest).

## Monthly PowerPoint report

At the bottom of the Dashboard tab, **Generate Monthly Usage Report** creates a three-slide report for the workbook's latest usage month. The deck includes a cover, a deterministic executive summary, and an editable annual usage chart with portfolio statistics. The issue summary also reflects the open or resolved state of priority alerts in the current browser session.

PowerPoint export uses the local presentation runtime bundled with Codex desktop. When launching elsewhere, set `PPTX_NODE` to a Node.js executable and `ARTIFACT_TOOL_MODULE` to the installed `@oai/artifact-tool/dist/artifact_tool.mjs` file before starting Streamlit.

## Validation against the supplied Excel source

Independently read the original cells with `openpyxl`, summed credits with `Decimal`,
and calculated date differences separately from the application. Trailing period:
July 2–31; prior period: June 2–July 1. Figures below are rounded only for presentation.

| Customer | Contract Excel row | Credits used | Trailing credits | Elapsed / remaining days | Projected total credits | Status |
|---|---:|---:|---:|---:|---:|---|
| CUST-03 | 7 | 180,000.7 | 43,905.8 | 180 / 185 | 450,753.13 | Over entitlement |
| CUST-02 | 6 | 224,399.3 | 12,114.4 | 201 / 164 | 290,624.69 | Early exhaustion risk |
| CUST-04 | 8 | 6,287.2 | 1,195.2 | 153 / 212 | 14,733.28 | Underutilizing |
| CUST-01 | 5 | 129,928.5 | 20,082.3 | 208 / 157 | 235,025.87 | On track |

Examples of independently checked arithmetic:

- CUST-03: (180,000.7 − 120,000) × (60,000 / 120,000) = **$30,000.35** current overage proxy.
- CUST-02: 224,399.3 + (12,114.4 / 30 × 164) = **290,624.6867** projected credits; / 240,000 = **121.0936%**.
- CUST-04: 6,287.2 + (1,195.2 / 30 × 212) = **14,733.28** projected credits; / 60,000 = **24.5555%**.
- CUST-01: 129,928.5 + (20,082.3 / 30 × 157) = **235,025.87** projected credits; / 240,000 = **97.9274%**.

Portfolio reconciliation: **1,752,000 entitlement**, **912,704.1 consumed**, **52.0950%
utilization**, **$30,000.35 current overage proxy**. Default statuses: 1 over, 1 risk,
6 underutilizing, 4 on track. No account passes the default acceleration threshold;
lowering growth to 20% flags CUST-05 (subject to the unchanged absolute floor).

Run regression tests:

```powershell
.\.venv\Scripts\python.exe -m unittest -v
```

The suite covers source reconciliation, exact rule boundaries, calendar windows,
zero usage, short history, missing days, inactive contracts, leap-day dates,
exhaustion dates, input rejection, acceleration, filter interaction, threshold
updates, and mocked AI request/fallback behavior. Source-dependent tests skip if
the original workbook is absent. No paid API request is made by the tests.
