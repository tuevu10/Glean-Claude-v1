"""Deterministic ingestion, forecasting and rules. No UI or AI dependencies.

Dates are day boundaries: [contract_start, contract_end). As-of usage includes
the full day; cutoff is midnight immediately after as_of. Ratios are fractions.
"""
from dataclasses import dataclass, asdict
from collections import Counter
from decimal import Decimal
from io import BytesIO
import math

import numpy as np
import pandas as pd


CONTRACT_FIELDS = ["customer_id", "contract_start", "term_months",
                   "annual_entitlement_credits", "annual_contract_value_usd"]
USAGE_FIELDS = ["date", "customer_id", "credits_used"]
STATUSES = ["OVER ENTITLEMENT", "EARLY EXHAUSTION RISK", "UNDERUTILIZING", "ON TRACK",
            "DATA REVIEW", "NOT STARTED"]


def credit_sum(values):
    """Consistent decimal accumulation avoids binary drift at exhaustion boundaries."""
    return float(sum((Decimal(str(v)) for v in values), Decimal(0)))


@dataclass(frozen=True)
class Rules:
    projected_utilization_threshold: float = 1.10
    early_exhaustion_days: int = 30
    underutilization_threshold: float = 0.70
    usage_acceleration_threshold: float = 0.30
    underutilization_min_elapsed: float = 0.25
    acceleration_min_increase_credits: float = 1000.0
    acceleration_min_history_days: int = 60
    over_entitlement_threshold: float = 1.0

    def __post_init__(self):
        if not all(math.isfinite(float(v)) and v >= 0 for v in asdict(self).values()):
            raise ValueError("Thresholds must be finite and nonnegative.")
        if self.projected_utilization_threshold < 1 or self.underutilization_threshold > 1:
            raise ValueError("Projected threshold must be >= 100%; underutilization <= 100%.")
        if self.over_entitlement_threshold != 1:
            raise ValueError("Over entitlement is a fixed 100% boundary.")
        if not 0 <= self.underutilization_min_elapsed <= 1:
            raise ValueError("Minimum elapsed fraction must be between zero and one.")
        if any(float(v).is_integer() is False for v in
               [self.early_exhaustion_days, self.acceleration_min_history_days]):
            raise ValueError("History and buffer days must be whole days.")
        if self.acceleration_min_history_days < 60:
            raise ValueError("Acceleration requires at least two 30-day windows.")


def _table(book, sheet, fields):
    raw = pd.read_excel(book, sheet_name=sheet, header=None, dtype=object)
    header = None
    for i, row in raw.iterrows():
        labels = [str(v).strip() for v in row if pd.notna(v)]
        if set(fields).issubset(labels):
            header = i
            break
    if header is None:
        raise ValueError(f"{sheet}: required headers not found: {', '.join(fields)}")
    labels = raw.loc[header].map(lambda v: str(v).strip())
    if any((labels == field).sum() != 1 for field in fields):
        raise ValueError(f"{sheet}: duplicate required headers.")
    data = raw.iloc[header + 1:].copy()
    data.columns = labels
    data = data[fields].dropna(how="all").copy()
    # Only these two reviewed source notes may be excluded; malformed rows fail.
    notes = data.iloc[:, 0].astype(str).str.startswith((
        "Notes: credits are billed", "All data is purely illustrative"))
    notes &= data.iloc[:, 1:].isna().all(axis=1)
    data = data.loc[~notes].copy()
    data["source_excel_row"] = data.index + 1
    return data.reset_index(drop=True)


def normalize(contracts, usage):
    c, u = contracts.copy(), usage.copy()
    for name, frame, fields in [("Contracts", c, CONTRACT_FIELDS), ("Usage Events", u, USAGE_FIELDS)]:
        if not set(fields).issubset(frame.columns):
            raise ValueError(f"{name}: missing required fields.")
        if frame[fields].isna().any().any():
            raise ValueError(f"{name}: missing required values. Correct the source workbook.")
        frame["customer_id"] = frame.customer_id.map(
            lambda x: str(int(x)) if isinstance(x, (float, np.floating)) and x.is_integer() else str(x).strip())
        if frame.customer_id.eq("").any():
            raise ValueError(f"{name}: empty customer ID.")
    for frame, col in [(c, "contract_start"), (u, "date")]:
        if frame[col].map(lambda v: isinstance(v, (int, float, np.number))).any():
            raise ValueError(f"{col}: numeric dates are ambiguous. Use Excel date cells or ISO YYYY-MM-DD text.")
        frame[col] = pd.to_datetime(frame[col], errors="coerce")
        if frame[col].isna().any():
            raise ValueError(f"Invalid dates in {col}.")
        frame[col] = frame[col].dt.normalize()
    for frame, cols in [(c, CONTRACT_FIELDS[2:]), (u, ["credits_used"])]:
        for col in cols:
            frame[col] = pd.to_numeric(frame[col], errors="coerce")
            if not np.isfinite(frame[col]).all() or frame[col].lt(0).any():
                raise ValueError(f"{col}: values must be finite, numeric and nonnegative.")
    if c.empty or u.empty:
        raise ValueError("Contracts and Usage Events must each contain records to determine an as-of date.")
    if c.customer_id.duplicated().any():
        raise ValueError("Duplicate contract customer IDs; one contract per customer is required.")
    if u.duplicated(["customer_id", "date"]).any():
        raise ValueError("Duplicate customer/date usage rows; resolve duplicates before loading.")
    if c.annual_entitlement_credits.le(0).any():
        raise ValueError("Entitlement must be positive; zero entitlements cannot support utilization or price ratios.")
    if c.term_months.ne(12).any():
        raise ValueError("This annual-entitlement model supports 12-month contracts only. Confirm reset/proration terms before extending it.")
    unknown = sorted(set(u.customer_id) - set(c.customer_id))
    if unknown:
        raise ValueError(f"Usage customers without contracts: {', '.join(unknown)}")
    c["contract_end"] = c.contract_start.map(lambda d: d + pd.DateOffset(months=12))
    joined = u.merge(c[["customer_id", "contract_start", "contract_end"]], on="customer_id", validate="many_to_one")
    invalid = (joined.date < joined.contract_start) | (joined.date >= joined.contract_end)
    if invalid.any():
        raise ValueError(f"{int(invalid.sum())} usage rows fall outside the associated contract. Correct or allocate them before loading.")
    return c, u.sort_values(["customer_id", "date"]).reset_index(drop=True)


def load_workbook(data: bytes):
    with pd.ExcelFile(BytesIO(data)) as book:
        missing = {"Contracts", "Usage Events"} - set(book.sheet_names)
        if missing:
            raise ValueError(f"Missing sheets: {', '.join(sorted(missing))}")
        contracts = _table(book, "Contracts", CONTRACT_FIELDS)
        usage = _table(book, "Usage Events", USAGE_FIELDS)
    return normalize(contracts, usage)


def _canonical_customer_id(value):
    return str(int(value)) if isinstance(value, (float, np.floating)) and value.is_integer() else str(value).strip()


def _canonical_rows(frame, fields, date_fields=(), numeric_fields=()):
    """Represent source and loaded records with comparable normalized scalar values."""
    rows = []
    for record in frame.to_dict("records"):
        values = []
        for field in fields:
            value = record[field]
            if field == "customer_id":
                value = _canonical_customer_id(value)
            elif field in date_fields:
                value = pd.Timestamp(value).normalize().date().isoformat()
            elif field in numeric_fields:
                value = Decimal(str(value))
            values.append(value)
        rows.append(tuple(values))
    return Counter(rows)


def _decimal_by_customer(frame, value_field):
    totals = {}
    for customer_id, values in frame.groupby("customer_id")[value_field]:
        totals[_canonical_customer_id(customer_id)] = sum(
            (Decimal(str(value)) for value in values), Decimal(0))
    return totals


def reconcile_source_import(data: bytes, contracts, usage, accounts):
    """Verify normalized rows and account rollups against the source workbook."""
    with pd.ExcelFile(BytesIO(data)) as book:
        source_contracts = _table(book, "Contracts", CONTRACT_FIELDS)
        source_usage = _table(book, "Usage Events", USAGE_FIELDS)

    contract_fields = CONTRACT_FIELDS + ["source_excel_row"]
    usage_fields = USAGE_FIELDS + ["source_excel_row"]
    source_contract_rows = _canonical_rows(
        source_contracts, contract_fields, ["contract_start"], CONTRACT_FIELDS[2:])
    loaded_contract_rows = _canonical_rows(
        contracts, contract_fields, ["contract_start"], CONTRACT_FIELDS[2:])
    source_usage_rows = _canonical_rows(
        source_usage, usage_fields, ["date"], ["credits_used"])
    loaded_usage_rows = _canonical_rows(
        usage, usage_fields, ["date"], ["credits_used"])

    source_usage_dates = pd.to_datetime(source_usage.date).dt.normalize()
    source_contract_dates = pd.to_datetime(source_contracts.contract_start).dt.normalize()
    loaded_usage_dates = pd.to_datetime(usage.date).dt.normalize()
    loaded_contract_dates = pd.to_datetime(contracts.contract_start).dt.normalize()

    source_used = _decimal_by_customer(source_usage.assign(
        customer_id=source_usage.customer_id.map(_canonical_customer_id)), "credits_used")
    account_used = {
        _canonical_customer_id(row.customer_id): Decimal(str(row.total_credits_used))
        for row in accounts.itertuples()
    }
    source_entitlement = {
        _canonical_customer_id(row.customer_id): Decimal(str(row.annual_entitlement_credits))
        for row in source_contracts.itertuples()
    }
    account_entitlement = {
        _canonical_customer_id(row.customer_id): Decimal(str(row.annual_entitlement_credits))
        for row in accounts.itertuples()
    }
    usage_customer_mismatches = sorted(
        customer_id for customer_id in set(source_used) | set(account_used)
        if source_used.get(customer_id) != account_used.get(customer_id))
    entitlement_customer_mismatches = sorted(
        customer_id for customer_id in set(source_entitlement) | set(account_entitlement)
        if source_entitlement.get(customer_id) != account_entitlement.get(customer_id))

    checks = [
        {"Check": "Contract row count", "Source": f"{len(source_contracts):,}",
         "Imported / Calculated": f"{len(contracts):,}", "Passed": len(source_contracts) == len(contracts)},
        {"Check": "Usage row count", "Source": f"{len(source_usage):,}",
         "Imported / Calculated": f"{len(usage):,}", "Passed": len(source_usage) == len(usage)},
        {"Check": "Contract start date range",
         "Source": f"{source_contract_dates.min():%b %d, %Y} – {source_contract_dates.max():%b %d, %Y}",
         "Imported / Calculated": f"{loaded_contract_dates.min():%b %d, %Y} – {loaded_contract_dates.max():%b %d, %Y}",
         "Passed": source_contract_dates.min() == loaded_contract_dates.min() and source_contract_dates.max() == loaded_contract_dates.max()},
        {"Check": "Usage event date range",
         "Source": f"{source_usage_dates.min():%b %d, %Y} – {source_usage_dates.max():%b %d, %Y}",
         "Imported / Calculated": f"{loaded_usage_dates.min():%b %d, %Y} – {loaded_usage_dates.max():%b %d, %Y}",
         "Passed": source_usage_dates.min() == loaded_usage_dates.min() and source_usage_dates.max() == loaded_usage_dates.max()},
        {"Check": "Contract field values", "Source": "Workbook contract rows",
         "Imported / Calculated": "All normalized rows match" if source_contract_rows == loaded_contract_rows else "Mismatch found",
         "Passed": source_contract_rows == loaded_contract_rows},
        {"Check": "Usage field values", "Source": "Workbook usage rows",
         "Imported / Calculated": "All normalized rows match" if source_usage_rows == loaded_usage_rows else "Mismatch found",
         "Passed": source_usage_rows == loaded_usage_rows},
        {"Check": "Credits used by customer", "Source": f"{len(source_used):,} customer totals",
         "Imported / Calculated": "All totals reconcile" if not usage_customer_mismatches else ", ".join(usage_customer_mismatches),
         "Passed": not usage_customer_mismatches},
        {"Check": "Entitlement by customer", "Source": f"{len(source_entitlement):,} customer totals",
         "Imported / Calculated": "All totals reconcile" if not entitlement_customer_mismatches else ", ".join(entitlement_customer_mismatches),
         "Passed": not entitlement_customer_mismatches},
    ]
    failed = [check for check in checks if not check["Passed"]]
    return {
        "passed": not failed,
        "failed_checks": len(failed),
        "usage_min": source_usage_dates.min(),
        "usage_max": source_usage_dates.max(),
        "contract_min": source_contract_dates.min(),
        "contract_max": source_contract_dates.max(),
        "usage_rows": len(source_usage),
        "contract_rows": len(source_contracts),
        "checks": checks,
    }


def classify(m, rules):
    """First matching rule wins; lifecycle prevents forecasts on inactive terms."""
    active = m["contract_lifecycle"] == "ACTIVE"
    if m["utilization_pct"] >= rules.over_entitlement_threshold:
        status, priority, owner = STATUSES[0], "P1 · Urgent", "Revenue Accounting / Billing"
        action = "Reconcile credit ledger and verify overage terms; coordinate a top-up or contract amendment."
        reason = "Consumed credits have reached or exceeded entitlement."
    elif m.get("missing_usage_days", 0) > 0:
        return dict(status="DATA REVIEW", priority="P2 · High",
                    recommended_owner="Revenue Accounting / Billing",
                    recommended_action="Verify usage-feed completeness and supply explicit zero-usage days before acting on forecasts.",
                    status_reason="Missing daily records prevent a reliable pacing or forecast classification.")
    elif m["contract_lifecycle"] == "NOT STARTED":
        return dict(status="NOT STARTED", priority="P4 · Routine", recommended_owner="No Action",
                    recommended_action="Wait for contract activation; no usage forecast is available.",
                    status_reason="The contract has not started; forecast-dependent values are unavailable.")
    elif active and (m["projected_utilization_pct"] > rules.projected_utilization_threshold or
                     (pd.notna(m["estimated_exhaustion_date"]) and
                      m["estimated_exhaustion_date"] < m["contract_end"] - pd.Timedelta(days=rules.early_exhaustion_days))):
        status, priority, owner = STATUSES[1], "P2 · High", "Sales / Account Management"
        action = "Review forecast with the account team and discuss additional credits before exhaustion."
        triggers = []
        if m["projected_utilization_pct"] > rules.projected_utilization_threshold:
            triggers.append("Projected utilization exceeds the configured threshold.")
        if pd.notna(m["estimated_exhaustion_date"]) and m["estimated_exhaustion_date"] < m["contract_end"] - pd.Timedelta(days=rules.early_exhaustion_days):
            triggers.append("Estimated exhaustion is earlier than the configured contract-end buffer.")
        reason = " ".join(triggers)
    elif m["contract_elapsed_pct"] >= rules.underutilization_min_elapsed and m["projected_utilization_pct"] < rules.underutilization_threshold:
        status, priority, owner = STATUSES[2], "P3 · Medium", "Customer Success"
        action = "Review adoption barriers, agree a usage plan, and assess renewal implications."
        reason = "Enough contract time has elapsed and projected utilization is below the configured threshold."
    else:
        status, priority, owner = STATUSES[3], "P4 · Routine", "No Action"
        action, reason = "Continue routine monitoring.", "No primary exception rule is triggered."
    if m["contract_lifecycle"] == "NOT STARTED":
        action = "Wait for contract activation; no usage forecast is available."
    elif m["contract_lifecycle"] == "EXPIRED":
        action += " Contract expired: confirm final reconciliation and renewal disposition."
        if status == STATUSES[3]:
            owner, priority = "Sales / Account Management", "P3 · Medium"
    if m["usage_accelerating"]:
        action += " Investigate the recent usage increase and confirm whether it will persist."
        if status == STATUSES[3]:
            owner, priority = "Sales / Account Management", "P3 · Medium"
    return dict(status=status, priority=priority, recommended_owner=owner,
                recommended_action=action, status_reason=reason)


def calculate_accounts(contracts, usage, rules=Rules()):
    """Inputs must be normalized. Includes calendar-day zeros, never forward fills usage."""
    as_of = usage.date.max()
    cutoff = as_of + pd.Timedelta(days=1)
    rows = []
    groups = {cid: frame.sort_values("date") for cid, frame in usage.groupby("customer_id", sort=False)}
    for c in contracts.to_dict("records"):
        start, end = c["contract_start"], c["contract_end"]
        events = groups.get(c["customer_id"], usage.iloc[:0])
        entitlement = float(c["annual_entitlement_credits"])
        total_days = (end - start).days
        elapsed = max(0, min((cutoff - start).days, total_days))
        remaining_days = max(0, total_days - elapsed)
        lifecycle = "NOT STARTED" if cutoff <= start else ("EXPIRED" if cutoff >= end else "ACTIVE")
        used = credit_sum(events.credits_used)
        remaining = entitlement - used
        windows = []
        for offset in (0, 30):
            right = cutoff - pd.Timedelta(days=offset)
            left = right - pd.Timedelta(days=30)
            eligible_days = max(0, (min(right, end) - max(left, start)).days)
            window = events.loc[(events.date >= left) & (events.date < right)]
            credits = credit_sum(window.credits_used)
            windows.append((credits, eligible_days, credits / eligible_days if eligible_days else 0.0))
        (trailing, trailing_days, rate), (prior, prior_days, prior_rate) = windows
        growth = trailing / prior - 1 if prior > 0 else (0.0 if trailing == 0 else np.nan)
        coverage_days = max(0, (min(cutoff, end) - start).days)
        missing_days = max(0, coverage_days - len(events))
        complete_comparison = missing_days == 0 and elapsed >= rules.acceleration_min_history_days and trailing_days == prior_days == 30
        accelerating = bool(lifecycle == "ACTIVE" and complete_comparison and
                            trailing - prior >= rules.acceleration_min_increase_credits and
                            ((prior > 0 and growth > rules.usage_acceleration_threshold) or (prior == 0 and trailing > 0)))
        forecast = rate * remaining_days if lifecycle == "ACTIVE" else 0.0
        projected = used + forecast
        exhaustion = pd.NaT
        if used >= entitlement:
            cumulative = Decimal(0)
            for event in events.itertuples():
                cumulative += Decimal(str(event.credits_used))
                if cumulative >= Decimal(str(entitlement)):
                    exhaustion = event.date
                    break
            days_to_exhaustion = 0.0
        elif lifecycle == "ACTIVE" and rate > 0:
            days_to_exhaustion = remaining / rate
            # First future usage day is as_of + 1; round up to a whole event date.
            try:
                exhaustion = as_of + pd.Timedelta(days=math.ceil(days_to_exhaustion))
            except (OverflowError, ValueError, pd.errors.OutOfBoundsDatetime):
                exhaustion = pd.NaT
        else:
            days_to_exhaustion = np.nan
        price = float(c["annual_contract_value_usd"]) / entitlement
        m = dict(c, as_of_date=as_of, contract_lifecycle=lifecycle,
                 implied_contract_value_per_credit=price, total_credits_used=used,
                 remaining_credits=remaining, utilization_pct=used / entitlement,
                 days_elapsed=elapsed, total_contract_days=total_days,
                 contract_elapsed_pct=elapsed / total_days,
                 pacing_index=(used / entitlement) / (elapsed / total_days) if elapsed else np.nan,
                 trailing_30_day_credits=trailing, prior_30_day_credits=prior,
                 trailing_30_day_average_daily_usage=rate, prior_30_day_average_daily_usage=prior_rate,
                 trailing_window_days=trailing_days, prior_window_days=prior_days,
                 usage_growth_30_day_pct=growth, usage_accelerating=accelerating,
                 growth_note="Unavailable: missing daily records" if missing_days else
                             "New usage from zero baseline" if prior == 0 and trailing > 0 else
                             ("Insufficient comparable history" if not complete_comparison else "Comparable 30-day windows"),
                 forecast_remaining_usage=forecast, projected_total_contract_usage=projected,
                 projected_utilization_pct=projected / entitlement,
                 days_remaining=remaining_days, days_to_exhaustion=days_to_exhaustion,
                 estimated_exhaustion_date=exhaustion, overage_credits=max(used-entitlement, 0),
                 estimated_overage_value=max(used-entitlement, 0)*price,
                 projected_overage_credits=max(projected-entitlement, 0),
                 projected_overage_value=max(projected-entitlement, 0)*price,
                 projected_unused_credits=max(entitlement-projected, 0),
                 projected_unused_contract_value=max(entitlement-projected, 0)*price,
                 source_usage_rows=len(events), missing_usage_days=missing_days,
                 last_usage_date=events.date.max() if len(events) else pd.NaT)
        m["forecast_available"] = lifecycle != "NOT STARTED" and missing_days == 0
        if not m["forecast_available"]:
            for field in ["forecast_remaining_usage", "projected_total_contract_usage",
                          "projected_utilization_pct", "projected_overage_credits", "projected_overage_value",
                          "projected_unused_credits", "projected_unused_contract_value"]:
                m[field] = np.nan
            if used < entitlement:
                m["days_to_exhaustion"], m["estimated_exhaustion_date"] = np.nan, pd.NaT
        if missing_days:
            # Observed credits are lower bounds; rates and growth are not valid evidence.
            for field in ["pacing_index", "trailing_30_day_average_daily_usage",
                          "prior_30_day_average_daily_usage", "usage_growth_30_day_pct"]:
                m[field] = np.nan
        m["data_notes"] = "; ".join(filter(None, [
            f"{missing_days} missing daily records; observed usage is incomplete and forecasts are withheld" if missing_days else "",
            "Short trailing window: rate uses eligible contract days" if 0 < trailing_days < 30 else "",
            "No recent rate; exhaustion date unavailable" if rate == 0 and used < entitlement else "",
            "Forecast unavailable before contract start" if lifecycle == "NOT STARTED" else ""]))
        m.update(classify(m, rules))
        rows.append(m)
    result = pd.DataFrame(rows)
    if np.isinf(result.select_dtypes(include="number").to_numpy()).any():
        raise ValueError("Calculated values overflowed numeric limits. Review source credit and dollar magnitudes.")
    return result.sort_values(
        ["priority", "estimated_overage_value", "projected_overage_value", "customer_id"],
        ascending=[True, False, False, True]).reset_index(drop=True)


def pacing_chart_data(account, usage):
    start, end = account["contract_start"], account["contract_end"]
    days = pd.date_range(start, end, freq="D")
    chart = pd.DataFrame({"date": days})
    chart["Expected linear pacing"] = ((chart.date-start).dt.days / (end-start).days * account["annual_entitlement_credits"])
    events = usage.loc[usage.customer_id.eq(account["customer_id"])].set_index("date").credits_used
    # Plot daily consumption at end-of-day boundary, leaving the future blank.
    daily = events.reindex(days[:-1], fill_value=0)
    cumulative = Decimal(0)
    actual = [0.0]
    for value in daily:
        cumulative += Decimal(str(value))
        actual.append(float(cumulative))
    chart["Actual cumulative usage"] = actual
    chart.loc[chart.date > account["as_of_date"] + pd.Timedelta(days=1), "Actual cumulative usage"] = np.nan
    return chart
