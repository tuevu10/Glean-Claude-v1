"""Auditable period chart series, using validated account rates only."""
import pandas as pd
from finance import credit_sum


def portfolio_chart(accounts, usage, start, end, snapshot):
    dates = pd.date_range(start, end, freq="D")
    cutoff = min(snapshot + pd.Timedelta(days=1), end)
    events = usage.loc[usage.customer_id.isin(accounts.customer_id) &
                       (usage.date >= start) & (usage.date < cutoff)]
    daily = events.groupby("date").credits_used.agg(credit_sum)
    actual = 0.0
    projected = None
    # Future-starting contracts have no usage history yet and should not suppress
    # the projection for accounts active at the snapshot.
    forecast_ok = bool(accounts.loc[
        accounts.contract_start.le(snapshot), "forecast_available"].all())
    snapshot = min(pd.Timestamp(snapshot), end - pd.Timedelta(days=1))
    first_month = pd.Timestamp(start).to_period("M")
    snapshot_month = snapshot.to_period("M")
    last_complete_month = snapshot_month if snapshot.is_month_end else snapshot_month - 1
    historical_months = pd.period_range(first_month, last_complete_month, freq="M")
    monthly_additions = [credit_sum(accounts.loc[
        accounts.contract_start.dt.to_period("M").eq(month) & accounts.contract_start.le(snapshot),
        "annual_entitlement_credits"]) for month in historical_months]
    average_monthly_addition = (credit_sum(monthly_additions) / len(monthly_additions)
                                if len(monthly_additions) else None)
    actual_contracted_at_snapshot = credit_sum(accounts.loc[
        accounts.contract_start.le(snapshot), "annual_entitlement_credits"])
    projection_months = len(pd.period_range(
        (snapshot + pd.Timedelta(days=1)).to_period("M"),
        (end - pd.Timedelta(days=1)).to_period("M"), freq="M")) if snapshot < end - pd.Timedelta(days=1) else 0
    rows = []
    for date in dates:
        if date > start and date <= cutoff:
            actual = credit_sum([actual, daily.get(date-pd.Timedelta(days=1), 0)])
        pacing = credit_sum([
            a.annual_entitlement_credits * max(0, (min(date, a.contract_end) -
                max(start, a.contract_start)).days) / a.total_contract_days
            for a in accounts.itertuples()])
        if date == cutoff:
            projected = actual
        elif date > cutoff and forecast_ok:
            event_date = date - pd.Timedelta(days=1)
            projected = credit_sum([projected, *[
                a.trailing_30_day_average_daily_usage for a in accounts.itertuples()
                if a.contract_start <= event_date < a.contract_end]])
        # Cumulative entitlements activated by this date; retain ended contracts
        # because the usage series also retains their historical consumption.
        contracted = credit_sum(accounts.loc[
            accounts.contract_start.le(min(date, snapshot)), "annual_entitlement_credits"])
        projected_contracted = None
        if average_monthly_addition is not None and date >= snapshot:
            # Draw a straight run-rate forecast across the remaining period.
            fraction = min(1.0, max(0.0, (date - snapshot) / (end - snapshot)))
            projected_contracted = (actual_contracted_at_snapshot +
                                    fraction * projection_months * average_monthly_addition)
        rows.append({"date": date, "Contract pacing": pacing,
                     "Total contracted credits": contracted if date <= snapshot else None,
                     "Projected contracted credits": projected_contracted,
                     "Credits consumed": actual if date <= cutoff else None,
                     "Projected consumption": projected if date >= cutoff and forecast_ok else None})
    result = pd.DataFrame(rows)
    result.attrs["average_monthly_contracted_addition"] = average_monthly_addition
    return result
