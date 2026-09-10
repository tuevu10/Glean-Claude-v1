# Production-readiness review

Scope: local, manually refreshed Finance review of annual credit contracts. This
does not approve automatic billing, scheduled monitoring, or shared hosting.

## Highest-priority findings

| Priority | Finding | Disposition |
|---|---|---|
| P1 | Fractional sums and cumulative sums disagreed at exact exhaustion, crashing the portfolio | Fixed: consistent decimal accumulation for totals, windows, crossing date, and chart |
| P1 | Sparse activity files could suppress valid forecasts | Fixed: dates without usage events count as zero consumption; source reconciliation still catches import mismatches, duplicates, invalid values, and out-of-contract activity |
| P2 | Numeric dates silently normalized to 1970 | Fixed: numeric dates rejected; require real Excel dates or ISO text |
| P2 | Future contracts exposed zero forecasts and full unused ACV | Fixed: NOT STARTED status and unavailable forecast-dependent values throughout |
| P2 | An old export could be mistaken for a current monitor | Fixed: Refresh source, read timestamp, snapshot type, stale/future-date notices. As-of stays source-driven |
| P2 | AI could invent financial claims or actions | Reduced: summary, financial impact, owner and action are deterministic. Only the explanation is model-authored |
| P2 | Config could redefine exhaustion or allow invalid windows | Fixed: 100% exhaustion; whole-day buffers/history; minimum 60-day acceleration history; bounded elapsed fraction |
| P3 | Calculation scanned all usage for every account | Fixed: group usage once by customer; parsing cache bounded to three sources |
| P3 | Source discovery depended on Downloads | Improved: environment override, workbook beside app.py, Downloads fallback, or upload |
| P3 | Unexpected failures lacked diagnostics | Improved: unexpected exceptions logged locally; invalid inputs suppress results; infinite calculated values fail closed |

## Validation

- Final suite: **22 tests passed**. Running app health endpoint returned `ok`.
- Sample workbook load plus calculation: 0.245 seconds on this run.
- Supplied portfolio reconciles to 1,752,000 entitlement, 912,704.1 consumed,
  and $30,000.35 current estimated overage. Original sample classifications remain unchanged.
- Regression coverage includes fractional exhaustion and chart reconciliation,
  numeric dates, missing feeds, observed overage despite gaps, explicit daily zeros,
  expired complete terms, threshold rejection, and deterministic AI impact/action.
- Refresh test replaces a temporary source with invalid bytes, verifies previous
  results disappear, restores the workbook, and verifies recovery.
- UI coverage includes filters, manual navigation, no-key mode, and AI failure fallback.
- Installed dependencies pass pip check. No paid/live OpenAI request was made.
- Synthetic benchmark on this Windows Python 3.12 environment: 1,000 customers and
  180,000 daily records normalized in 0.309 seconds and calculated in 1.079 seconds.
  One measurement, not an SLA; Excel parsing, browser rendering, and cold startup
  are excluded. Large Excel workbooks remain subject to parsing and memory costs.

## Residual constraints, in priority order

1. **Contract semantics require Finance sign-off.** One 12-month contract per customer.
   Negotiated overage prices, rollover, refunds, proration, amendments and invoice
   rounding are absent. Dollar values remain commercial proxies.
2. **Feed completeness depends on the upstream export.** Every elapsed contract day
   must have a record, including explicit zeros. Sparse event-only exports require
   a separate completeness signal before this conservative rule can be relaxed.
3. **Thresholds are heuristics.** 110% projection, 30-day buffer, 70% underuse, 25%
   elapsed, and 30% growth with a 1,000-credit floor need operational calibration.
   A complete but very short history can still produce a volatile forecast.
4. **Refresh is manual.** No scheduler, upstream watermark, completeness certification,
   or retained history. Read time is not export-generation time. The three-day notice
   is operational guidance and does not change reproducible calculations.
5. **AI explanation needs review.** Structural checks cannot prove qualitative claims.
   Requests fall back safely; live availability/model access are not verified.
6. **Local access only.** No authentication or multi-user authorization. Keep localhost
   binding. Shared deployment needs a separate security and operations review.
7. **Portability and scale remain bounded.** Windows/Python 3.12 tested. Other OS clean
   installs and large Excel imports unverified. Direct dependencies are pinned; a
   transitive lockfile and CI are appropriate for a maintained shared product.

## Maintainability and portability

Rules remain in finance.py; presentation in app.py; optional text in briefs.py.
Statuses and forecast availability are exported. Tests target observable failures.
No database, authentication framework, scheduler, or extra frontend was introduced.

Use Python 3.12 and install requirements in an isolated virtual environment. On
macOS/Linux use `.venv/bin/python` instead of `.venv\Scripts\python.exe` in the README
commands. Set FINANCE_WORKBOOK or place the workbook beside app.py. No absolute
developer-specific path is required by the app. Source-dependent tests skip when the
sample is absent; pure finance tests remain runnable.
