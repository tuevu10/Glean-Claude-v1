"""Run with: python -m streamlit run app.py"""
from dataclasses import asdict
from hashlib import sha256
import json
import os
import logging
import base64
from html import escape
from datetime import datetime, timezone
from pathlib import Path

import altair as alt
import pandas as pd
import streamlit as st

from finance import Rules, STATUSES, load_workbook, calculate_accounts, pacing_chart_data
from briefs import deterministic_brief, generate_ai_brief, account_facts


SOURCE_NAME = "Product Finance - AI Fluency Test Sample Data.xlsx"
DEFAULT_SOURCE = Path(os.environ.get("FINANCE_WORKBOOK", str(
    Path(__file__).parent / SOURCE_NAME if (Path(__file__).parent / SOURCE_NAME).exists()
    else Path.home() / "Downloads" / SOURCE_NAME)))

st.set_page_config(page_title="Usage & Billing Dashboard", layout="wide", initial_sidebar_state="collapsed")
st.markdown("""<style>
.block-container {padding-top:4rem; padding-bottom:3rem; max-width:1600px;}
[data-testid="stApp"]::before {content:"";position:fixed;top:0;left:0;right:0;height:6px;background:#4438FF;z-index:999999;pointer-events:none;}
[data-testid="stMetric"] {background:#f5f8f9;border:1px solid #e1e9ec;border-radius:8px;padding:14px;}
[data-testid="stMetricValue"] {font-size:1.65rem;}
[data-testid="stMetricLabel"] {font-size:0.85rem;}
[data-testid="stMetricLabel"] p {white-space:normal; overflow:visible;}
.action-card {border-radius:10px;padding:20px 22px;border:1px solid var(--line);border-left:5px solid var(--accent);background:var(--wash);min-height:245px;}
.action-card .eyebrow {color:var(--accent);font-size:12px;font-weight:750;text-transform:uppercase;letter-spacing:.06em;}
.action-card h3 {margin:8px 0 2px;padding:0;font-size:24px;}
.action-card .signal {font-size:19px;font-weight:650;margin:8px 0 14px;color:var(--accent);}
.action-card .owner {font-size:13px;font-weight:650;margin:12px 0 5px;}
.action-card p {margin:0;font-size:14px;line-height:1.55;}
.snapshot {display:inline-block;border-radius:5px;padding:4px 10px;background:#F2F0FF;color:#574B88;font-size:12px;}
.st-key-kpi_over [data-testid="stMetric"],
.st-key-kpi_risk [data-testid="stMetric"],
.st-key-kpi_value [data-testid="stMetric"] {background:#F1EEE8;border-color:#DAD6CD;border-top:4px solid #DAD6CD;}
[data-baseweb="tab"] {font-weight:650;}
h1 {letter-spacing:-0.035em;} h2 {letter-spacing:-0.02em;}
.dashboard-header {display:flex;align-items:center;justify-content:space-between;gap:24px;margin-bottom:16px;}
.dashboard-header h1 {margin:0;padding:0;font-size:2.25rem;line-height:1.2;min-width:0;overflow-wrap:break-word;}
.dashboard-header img {width:112px;height:auto;flex-shrink:0;}
@media(max-width:700px) {.dashboard-header {align-items:flex-start;gap:16px;} .dashboard-header h1 {font-size:1.6rem;} .dashboard-header img {width:85px;margin-top:6px;}}
</style>""", unsafe_allow_html=True)


@st.cache_data(show_spinner=False, max_entries=3)
def parse_source(data):
    return load_workbook(data)


def fmt(value, kind="number"):
    if pd.isna(value):
        return "N/A"
    if kind == "date":
        return pd.Timestamp(value).strftime("%b %d, %Y")
    if kind == "pct":
        return f"{value:.1%}"
    if kind == "usd":
        return f"${value:,.2f}"
    if kind == "ratio":
        return f"{value:.2f}×"
    return f"{value:,.1f}"


with st.sidebar:
    st.subheader("Source workbook")
    upload = st.file_uploader("Upload a replacement workbook", type=["xlsx"])
    source_path = st.text_input("Local workbook path", str(DEFAULT_SOURCE))
    st.button("Refresh source", help="Rereads the local file. For an upload, remove it and upload the new export.")
    st.caption("Manual refresh only. An uploaded file takes precedence over the local path.")
    st.subheader("Monitoring Assumptions")
    defaults = Rules()
    projected = st.number_input("Projected utilization threshold (%)", 100.0, 500.0, defaults.projected_utilization_threshold * 100, 5.0)
    early = st.number_input("Early exhaustion days", 0, 365, defaults.early_exhaustion_days, 5)
    under = st.number_input("Underutilization threshold (%)", 0.0, 100.0, defaults.underutilization_threshold * 100, 5.0)
    acceleration = st.number_input("Usage acceleration threshold (%)", 0.0, 1000.0, defaults.usage_acceleration_threshold * 100, 5.0)
    st.caption("Acceleration also requires ≥1,000 additional credits and two complete 30-day contract windows. Underutilization starts at 25% elapsed.")
    rules = Rules(projected_utilization_threshold=projected / 100, early_exhaustion_days=early,
                  underutilization_threshold=under / 100, usage_acceleration_threshold=acceleration / 100)

logo_data = base64.b64encode((Path(__file__).parent / "assets" / "glean-logo.svg").read_bytes()).decode("ascii")
st.markdown(f'<div class="dashboard-header"><h1>Usage &amp; Billing Dashboard</h1><img src="data:image/svg+xml;base64,{logo_data}" alt="Glean"></div>', unsafe_allow_html=True)

try:
    raw = upload.getvalue() if upload else Path(source_path).expanduser().read_bytes()
    source_name = upload.name if upload else Path(source_path).name
    contracts, usage = parse_source(raw)
    accounts = calculate_accounts(contracts, usage, rules)
    if "forecast_available" not in accounts.columns:
        raise RuntimeError("Application modules are out of sync. Stop Streamlit and restart it to load the current finance engine.")
except Exception as exc:
    if not isinstance(exc, (ValueError, OSError)):
        logging.getLogger(__name__).exception("Workbook analysis failed")
    st.error(f"Unable to analyze workbook: {exc}")
    st.info("Choose the supplied workbook or upload an Excel file with Contracts and Usage Events sheets. No results are produced from invalid data.")
    st.stop()

as_of = accounts.as_of_date.iloc[0]
source_hash = sha256(raw).hexdigest()
read_at = datetime.now(timezone.utc)
age_days = (pd.Timestamp.now().normalize() - as_of).days
snapshot_label = f"Snapshot · {as_of:%b %d, %Y}" + (" · Refresh recommended" if age_days > 3 else " · Check future date" if age_days < 0 else "")
st.markdown(f'<span class="snapshot">{escape(snapshot_label)}</span>', unsafe_allow_html=True)
with st.sidebar.expander("Source information"):
    st.write(source_name)
    st.write(f"Source read: {read_at:%Y-%m-%d %H:%M:%S} UTC")
    st.write("Uploaded snapshot" if upload else "Local workbook")
    st.write("Manual refresh; no background monitoring.")
    if age_days > 3:
        st.warning(f"Latest usage is {age_days} days old. Expected for historical sample data. Refresh the export for current decisions.")
    elif age_days < 0:
        st.warning("Future-dated usage: verify the source dates.")

issues = accounts.loc[accounts.data_notes.ne(""), ["customer_id", "data_notes"]]
if not issues.empty:
    st.warning(f"{len(issues)} account(s) have data or forecast limitations. Review Data & calculation audit below.")
    st.caption("Observed usage and overage may be lower bounds where daily records are missing. Forecast-dependent values are unavailable for those accounts.")

def open_account(customer_id):
    st.session_state["detail_customer"] = customer_id
    st.session_state["workspace_view"] = "Customer Details"

dashboard_tab, details_tab = st.tabs(["Dashboard", "Customer Details"], key="workspace_view", on_change="rerun")
with dashboard_tab:
    st.subheader("Priority")
    urgent = accounts.loc[accounts.recommended_owner.ne("No Action")].head(2)
    if urgent.empty:
        st.success("No accounts require follow-up under the current rules.")
    else:
        for column, item in zip(st.columns(len(urgent)), urgent.to_dict("records")):
            priority = item["priority"].split(" · ")[0]
            palette = {"P1": ("#6550B5", "#F0EDFF", "#DAD2F4"),
                       "P2": ("#7862AA", "#F7F3FF", "#E5DDF2"),
                       "P3": ("#616184", "#F5F4FA", "#E1DFEB")}
            accent, wash, line = palette.get(priority, ("#5E626A", "#F7F7F8", "#E3E3E7"))
            signal = (fmt(item["estimated_overage_value"], "usd") + " current overage" if item["status"] == "OVER ENTITLEMENT"
                      else fmt(item["projected_utilization_pct"], "pct") + " projected utilization" if item["forecast_available"]
                      else "Validate the usage feed")
            with column:
                st.markdown(
                    f'<div class="action-card" style="--accent:{accent};--wash:{wash};--line:{line}">'
                    f'<div class="eyebrow">{escape(item["priority"])} · {escape(item["status"])}</div>'
                    f'<h3>{escape(item["customer_id"])}</h3>'
                    f'<div class="signal">{escape(signal)}</div>'
                    f'<p>{escape(item["status_reason"])}</p>'
                    f'<div class="owner">{escape(item["recommended_owner"])}</div>'
                    f'<p>{escape(item["recommended_action"])}</p></div>',
                    unsafe_allow_html=True)
                st.button("Review " + item["customer_id"], key="review_" + item["customer_id"],
                          on_click=open_account, args=(item["customer_id"],), width="stretch")

with dashboard_tab:
    kpis = [
        ("Total Contracted Credits", f"{accounts.annual_entitlement_credits.sum():,.0f}"),
        ("Total Credits Consumed", f"{accounts.total_credits_used.sum():,.1f}"),
        ("Portfolio Utilization %", f"{accounts.total_credits_used.sum() / accounts.annual_entitlement_credits.sum():.1%}"),
        ("Customers Over Entitlement", str(accounts.status.eq(STATUSES[0]).sum())),
        ("Early Exhaustion Risk", str(accounts.status.eq(STATUSES[1]).sum())),
        ("Estimated Overage Value", f"${accounts.estimated_overage_value.sum():,.2f}"),
    ]
    st.subheader("Portfolio overview")
    for col, (label, value), card_key in zip(st.columns(3), kpis[3:], ["kpi_over", "kpi_risk", "kpi_value"]):
        with col.container(key=card_key):
            st.metric(label, value)
    with st.expander("Credits & utilization"):
        for col, (label, value) in zip(st.columns(3), kpis[:3]):
            col.metric(label, value)
    st.caption("Portfolio totals include all accounts. Overage dollars are estimates at implied contract value per credit, not invoice amounts.")
    
    st.subheader("Finance Action Queue")
    st.caption("Urgent accounts first. Select a row to open Customer Details.")
    filter_cols = st.columns([1, 1.6])
    exceptions_only = filter_cols[0].toggle("Exceptions only", value=True, key="exceptions_only",
        help="Hides only accounts assigned No Action. Existing priorities and statuses are unchanged.")
    with filter_cols[1].expander("Filter queue"):
        customers = st.multiselect("Customer", sorted(accounts.customer_id), key="filter_customer")
        statuses = st.multiselect("Status", STATUSES, key="filter_status")
        priorities = st.multiselect("Priority", sorted(accounts.priority.unique()), key="filter_priority")
        st.caption("Customer detail can also open any account, including accounts outside these filters.")
    filtered = accounts.copy()
    if exceptions_only:
        filtered = filtered.loc[filtered.recommended_owner.ne("No Action")]
    for field, chosen in [("customer_id", customers), ("status", statuses), ("priority", priorities)]:
        if chosen:
            filtered = filtered.loc[filtered[field].isin(chosen)]
    filtered = filtered.reset_index(drop=True)
    
    QUEUE = {
        "priority": "Priority", "customer_id": "Customer", "status": "Status",
        "usage_accelerating": "Usage Accelerating", "total_credits_used": "Credits Used",
        "annual_entitlement_credits": "Entitlement", "utilization_pct": "Utilization %",
        "contract_elapsed_pct": "Contract Elapsed %", "pacing_index": "Pacing Index",
        "projected_utilization_pct": "Projected Utilization %", "days_to_exhaustion": "Days to Exhaustion",
        "estimated_overage_value": "Estimated Overage $", "recommended_owner": "Recommended Owner",
        "recommended_action": "Recommended Action",
    }
    display = filtered[list(QUEUE)].rename(columns=QUEUE).copy()
    for col in ["Utilization %", "Contract Elapsed %", "Projected Utilization %"]:
        display[col] *= 100
    config = {c: st.column_config.NumberColumn(c, format="%.1f%%") for c in ["Utilization %", "Contract Elapsed %", "Projected Utilization %"]}
    for col in ["Credits Used", "Entitlement", "Days to Exhaustion"]:
        config[col] = st.column_config.NumberColumn(col, format="localized", width="small")
    config.update({"Estimated Overage $": st.column_config.NumberColumn(format="dollar"),
                   "Pacing Index": st.column_config.NumberColumn(format="%.2f"),
                   "Recommended Action": st.column_config.TextColumn(width="large"),
                   "Recommended Owner": st.column_config.TextColumn(width="medium")})
    colors = {STATUSES[0]: "#6550B5", STATUSES[1]: "#7862AA", STATUSES[2]: "#616184", STATUSES[3]: "#5E626A"}
    styled = display.style.map(lambda s: f"color: {colors.get(s, '#20313b')}; font-weight: 600", subset=["Status"])
    # Compact review signals only format existing results; the engine remains authoritative.
    def review_signal(row):
        if row["status"] == "DATA REVIEW":
            return "Data review · missing daily usage records"
        if row["status"] == "NOT STARTED":
            return "Not started · forecast unavailable"
        if row["status"] == "OVER ENTITLEMENT":
            text = f'Over entitlement · {row["utilization_pct"]:.1%} consumed'
        elif row["status"] == "EARLY EXHAUSTION RISK":
            text = f'Early exhaustion risk · {row["projected_utilization_pct"]:.1%} projected'
        elif row["status"] == "UNDERUTILIZING":
            text = f'Underutilizing · {row["projected_utilization_pct"]:.1%} projected'
        else:
            text = f'On track · {row["projected_utilization_pct"]:.1%} projected'
        if row["missing_usage_days"]:
            text += " · Incomplete feed"
        return text + (" · Accelerating" if row["usage_accelerating"] else "")
    
    compact = pd.DataFrame({
        "Priority": filtered.priority.str.split(" · ").str[0],
        "Customer": filtered.customer_id,
        "Why review": filtered.apply(review_signal, axis=1) if len(filtered) else pd.Series(dtype=str),
        "Current overage": filtered.estimated_overage_value,
    })
    queue_key = "queue_" + sha256((source_hash + json.dumps(asdict(rules), sort_keys=True)
                                   + "|".join(filtered.customer_id)).encode()).hexdigest()[:16]
    
    def select_queue_account():
        rows = st.session_state[queue_key]["selection"]["rows"]
        if rows:
            st.session_state["detail_customer"] = filtered.iloc[rows[0]].customer_id
            st.session_state["workspace_view"] = "Customer Details"
    
    priority_palette = {"P1": ("#F0EDFF", "#59439F"), "P2": ("#F7F3FF", "#725A9C"),
                        "P3": ("#F5F4FA", "#616184"), "P4": ("#F7F7F8", "#5E626A")}
    def color_queue(row):
        wash, ink = priority_palette.get(row["Priority"], ("#fff", "#20313b"))
        return [f"background-color:{wash};color:{ink};" + ("font-weight:650;" if col in ["Priority", "Why review"] else "") for col in row.index]
    compact_style = compact.style.apply(color_queue, axis=1)
    st.dataframe(compact_style, hide_index=True, width="stretch", height=min(335, 38 + 36 * max(len(compact), 1)),
        column_config={
            "Priority": st.column_config.TextColumn(width=65),
            "Customer": st.column_config.TextColumn(width=100),
            "Why review": st.column_config.TextColumn(width="large"),
            "Current overage": st.column_config.NumberColumn(format="dollar", width=150,
                help="Current overage only. Full forecast exposure is in the account details."),
        }, on_select=select_queue_account, selection_mode="single-row", key=queue_key)
    st.caption(f"{len(filtered)} of {len(accounts)} accounts shown · P1 urgent / P2 high / P3 medium / P4 routine")
    if filtered.empty:
        st.info("No accounts match. Turn off Exceptions only or clear filters to broaden the queue.")
    with st.expander("Full queue metrics"):
        st.dataframe(styled, hide_index=True, width="stretch", column_config=config)
    
    audit_export = filtered.copy()
    audit_export["source_sha256"] = source_hash
    for key, value in asdict(rules).items():
        audit_export[f"rule_{key}"] = value
    st.download_button("Download queue & calculated metrics (CSV)", audit_export.to_csv(index=False).encode("utf-8-sig"),
                       f"finance_action_queue_{as_of:%Y%m%d}.csv", "text/csv")

with details_tab:
    
    st.subheader("Customer detail")
    if st.session_state.get("detail_customer") not in set(accounts.customer_id):
        st.session_state["detail_customer"] = accounts.iloc[0].customer_id
    selected = st.selectbox("Select customer", accounts.customer_id.tolist(), key="detail_customer",
        help="Search by customer ID. Queue row selection opens the same account here.")
    m = accounts.loc[accounts.customer_id.eq(selected)].iloc[0].to_dict()
    st.markdown(f"**{m['status']}** · {m['priority']} · {m['contract_lifecycle'].title()}" + (" · USAGE ACCELERATING" if m["usage_accelerating"] else ""))
    st.caption(f"Contract: {fmt(m['contract_start'], 'date')} to {fmt(m['contract_end'], 'date')} (end exclusive) · {int(m['term_months'])} months")
    # Put the decision ahead of supporting metrics.
    with st.container(border=True):
        st.markdown("**Why this account is flagged**" if m["status"] != "ON TRACK" or m["usage_accelerating"] else "**Monitoring assessment**")
        st.write(m["status_reason"])
        if m["status"] == "EARLY EXHAUSTION RISK":
            st.caption(f"Estimated exhaustion: {fmt(m['estimated_exhaustion_date'], 'date')} · Contract end: {fmt(m['contract_end'], 'date')}")
        elif m["status"] == "UNDERUTILIZING":
            st.caption(f"Contract elapsed: {fmt(m['contract_elapsed_pct'], 'pct')} · Projected utilization: {fmt(m['projected_utilization_pct'], 'pct')}")
        if m["usage_accelerating"]:
            st.write("Usage is accelerating under the configured growth and materiality rules.")
        st.markdown("**Owner: " + m["recommended_owner"] + "**")
        st.write(m["recommended_action"])
    if m["data_notes"]:
        st.warning(m["data_notes"])
    key_metrics = [
        ("Credits Used", "total_credits_used", "number"),
        ("Utilization", "utilization_pct", "pct"),
        ("Projected Utilization", "projected_utilization_pct", "pct"),
    ]
    for col, (label, key, kind) in zip(st.columns(3), key_metrics):
        col.metric(label, "N/A" if key == "projected_utilization_pct" and m["contract_lifecycle"] == "NOT STARTED" else fmt(m[key], kind))
    with st.expander("Contract, balance & exhaustion details"):
        fields = [
            ("Annual Contract Value", "annual_contract_value_usd", "usd"),
            ("Entitlement", "annual_entitlement_credits", "number"),
            ("Remaining Credits", "remaining_credits", "number"),
            ("Contract Elapsed", "contract_elapsed_pct", "pct"),
            ("Pacing Index", "pacing_index", "ratio"),
            ("Days to Exhaustion", "days_to_exhaustion", "number"),
            ("Exhaustion Date", "estimated_exhaustion_date", "date"),
            ("Current Overage", "estimated_overage_value", "usd"),
        ]
        st.table(pd.DataFrame([{"Metric": label, "Value": fmt(m[key], kind)} for label, key, kind in fields]))
        st.caption("Negative balance means overconsumption. Exhausted accounts show the first actual exhaustion date. Future dates assume the recent rate continues, even beyond contract end.")
    
    left, right = st.columns([1.3, 1])
    with left:
        st.markdown("**Cumulative usage vs. contract pacing**")
        chart_data = pacing_chart_data(m, usage).melt("date", var_name="Series", value_name="Credits")
        chart = alt.Chart(chart_data).mark_line(strokeWidth=2).encode(
            x=alt.X("date:T", title="Date (end-of-day boundaries)"), y=alt.Y("Credits:Q", title="Cumulative credits"),
            color=alt.Color("Series:N", scale=alt.Scale(domain=["Actual cumulative usage", "Expected linear pacing"], range=["#8271DD", "#B6B3C0"]), legend=alt.Legend(orient="bottom")),
            tooltip=[alt.Tooltip("date:T", title="Boundary date"), "Series:N", alt.Tooltip("Credits:Q", format=",.1f")])
        st.altair_chart(chart.properties(height=310), width="stretch")
        st.caption("Actual usage stops at the analysis cutoff; expected pacing continues to the full contract entitlement.")
        if m["missing_usage_days"]:
            st.caption("Incomplete feed: this is observed cumulative usage only, not a complete consumption ledger.")
    with right:
        st.markdown("**Daily usage trend**")
        end_date = min(as_of, m["contract_end"] - pd.Timedelta(days=1))
        days = pd.date_range(m["contract_start"], end_date, freq="D")
        daily = usage.loc[usage.customer_id.eq(selected)].set_index("date").credits_used.reindex(days, fill_value=0)
        trend = daily.rename_axis("date").reset_index(name="credits_used")
        daily_chart = alt.Chart(trend).mark_bar(color="#A397E0").encode(
            x=alt.X("date:T", title="Usage date"), y=alt.Y("credits_used:Q", title="Daily credits"),
            tooltip=[alt.Tooltip("date:T"), alt.Tooltip("credits_used:Q", format=",.1f")])
        st.altair_chart(daily_chart.properties(height=310), width="stretch")
        if m["missing_usage_days"]:
            st.caption("Missing dates are shown with zero-height bars; they are unknown, not confirmed zero usage.")
    
    with st.expander("Recent usage & commercial impact", expanded=False):
        detail_fields = {
            "trailing_30_day_credits": "Trailing 30-day credits", "prior_30_day_credits": "Prior 30-day credits",
            "trailing_30_day_average_daily_usage": "Trailing average daily usage", "prior_30_day_average_daily_usage": "Prior average daily usage",
            "usage_growth_30_day_pct": "30-day usage growth", "forecast_remaining_usage": "Forecast remaining credits",
            "projected_total_contract_usage": "Projected total credits", "projected_overage_value": "Total projected overage proxy",
            "projected_unused_contract_value": "Projected unused contract value", "implied_contract_value_per_credit": "Implied value per credit",
        }
        st.table(pd.DataFrame([{"Metric": label, "Value": fmt(m[key], "pct" if key.endswith("pct") else "usd" if "value" in key else "number")} for key, label in detail_fields.items()]))
        st.caption(f"{m['growth_note']}. Total projected overage includes current overage; do not add the two. Unused contract value is an adoption/renewal indicator, not a refund or revenue-loss estimate.")
    
    
    with st.expander("Data & calculation audit"):
        st.caption(f"Source SHA-256: {source_hash}")
        st.write("Contract and usage source row numbers refer to the original Excel worksheet. All calculations use unrounded values; only display values are rounded.")
        st.json(asdict(rules))
        st.markdown("""- Cutoff = latest usage date + one day. Contract end = start + 12 calendar months (exclusive).
    - Elapsed days = cutoff minus start, clamped to the contract term. Utilization = used / entitlement; pacing = utilization / elapsed fraction.
    - Recent periods = the 30 calendar days ending on the as-of date and the preceding 30 days. Rates divide by eligible contract days, up to 30. Missing daily records withhold rates and forecasts and require data review.
    - Projected total = consumed + recent daily rate × remaining contract days. Inactive terms have no future usage forecast.
    - Overage proxy = max(consumed − entitlement, 0) × ACV / entitlement.
    - Rules are evaluated in order: over entitlement, data review for missing records, not started, early exhaustion risk (active contracts), underutilizing, then on track. Acceleration is independent and requires complete records.
    - Priority order is P1, P2, P3, P4; ties sort by current overage value, projected overage value, then customer ID.
    - Zero baselines with new usage have undefined percentage growth. Acceleration can still trigger after two full windows and the absolute materiality floor.
    """)
        if not issues.empty:
            st.dataframe(issues, hide_index=True, width="stretch")
        st.markdown("**Selected account: full calculated record**")
        st.dataframe(pd.DataFrame([{"Field": k, "Value": str(v)} for k, v in m.items()]), hide_index=True, width="stretch")
        st.markdown("**Source contract**")
        st.dataframe(contracts.loc[contracts.customer_id.eq(selected)], hide_index=True, width="stretch")
        st.markdown("**Source daily usage**")
        st.dataframe(usage.loc[usage.customer_id.eq(selected)], hide_index=True, width="stretch")
        st.download_button("Download monitoring assumptions (JSON)", json.dumps({"as_of": str(as_of.date()), "source_sha256": source_hash, "rules": asdict(rules)}, indent=2), "monitoring_assumptions.json", "application/json")

with dashboard_tab:
    with st.expander("Account interpretation & suggested action", expanded=True):
        st.markdown("**" + selected + "**")
        fallback = deterministic_brief(m)
        brief_key = sha256((json.dumps(account_facts(m), sort_keys=True) + source_hash + json.dumps(asdict(rules), sort_keys=True)).encode()).hexdigest()
        if os.environ.get("OPENAI_API_KEY", "").strip():
            st.caption("Optional: send this customer's validated metrics to OpenAI for a four-bullet draft. All financial figures remain controlled by Python.")
            if st.button("Generate Finance Brief"):
                try:
                    with st.spinner("Drafting from validated facts…"):
                        st.session_state["ai_brief"] = (brief_key, generate_ai_brief(m))
                except Exception:
                    st.session_state.pop("ai_brief", None)
                    st.warning("AI brief was unavailable or failed validation. Showing the deterministic brief.")
            cached = st.session_state.get("ai_brief")
            bullets = cached[1] if cached and cached[0] == brief_key else fallback
            st.caption("AI explanation — review interpretation before sharing. Summary, financial impact, owner and action are deterministic." if bullets is not fallback else "Deterministic brief")
        else:
            st.caption("Deterministic brief · no API key required")
            bullets = fallback
        for bullet in bullets:
            # Dollar signs must remain currency, not paired Markdown math delimiters.
            st.markdown("• " + bullet.replace("$", "\\$"))
