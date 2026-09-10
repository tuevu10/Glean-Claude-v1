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



from finance import Rules, STATUSES, load_workbook, calculate_accounts, pacing_chart_data, credit_sum

from portfolio import portfolio_chart
from reports import build_monthly_report_payload, generate_monthly_report
from excel_export import build_audit_workbook_payload, generate_audit_workbook





SOURCE_NAME = "Product Finance - AI Fluency Test Sample Data.xlsx"

DEFAULT_SOURCE = Path(os.environ.get("FINANCE_WORKBOOK", str(

    Path(__file__).parent / SOURCE_NAME if (Path(__file__).parent / SOURCE_NAME).exists()

    else Path.home() / "Downloads" / SOURCE_NAME)))



st.set_page_config(page_title="Usage & Billing Dashboard", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""<style>

.block-container {padding-top:3.5rem; padding-bottom:3rem; max-width:1600px;}

[data-testid="stApp"]::before {content:"";position:fixed;top:0;left:0;right:0;height:6px;background:#4438FF;z-index:999999;pointer-events:none;}

[data-testid="stMetric"] {background:#f5f8f9;border:1px solid #e1e9ec;border-radius:8px;padding:14px;}

[data-testid="stMetricValue"] {font-size:1.65rem;}

[data-testid="stMetricLabel"] {font-size:0.85rem;}

[data-testid="stMetricLabel"] p {white-space:normal; overflow:visible;}

[class*="st-key-alert_row_"] {background:#F7F4FF;border-left:3px solid #8271DD;border-radius:7px;padding:2px 8px;margin-bottom:0;}
[class*="st-key-alert_row_"]:has([class*="st-key-resolve_done_"]) {background:#FFFFFF;border-left-color:transparent;}
[class*="st-key-alert_row_"]:has([class*="st-key-resolve_done_"]) [data-testid="stExpander"] details {background:#FFFFFF;}
[class*="st-key-alert_row_"] [data-testid="stExpander"] {border:0;background:transparent;}
[class*="st-key-alert_row_"] [data-testid="stExpander"] summary p {white-space:normal;overflow:visible;font-weight:600;}
[class*="st-key-resolve_open_"] button {background:#F0F0F2;color:#666874;border-color:#DDDDE2;}
[class*="st-key-resolve_done_"] button {background:#E7F4EA;color:#247341;border-color:#B8DCC2;}
[class*="st-key-alert_row_"] button p {white-space:normal;overflow:visible;}
.priority-alert {padding:9px 14px;border-left:3px solid #8271DD;border-radius:6px;background:#F7F4FF;font-size:15px;line-height:1.4;} .priority-alert span {color:#626477;font-size:13px;} .alert-status {font-size:12px;color:#6550B5;margin-top:2px;}
.action-card {border-radius:10px;padding:20px 22px;border:1px solid var(--line);border-left:5px solid var(--accent);background:var(--wash);}

.action-card .eyebrow {color:var(--accent);font-size:12px;font-weight:750;text-transform:uppercase;letter-spacing:.06em;}

.resolved-line {padding:12px 16px;background:#F1F7F2;border:1px solid #D5E7D8;border-radius:8px;display:flex;justify-content:space-between;gap:16px;align-items:center;} .resolved-check {color:#287546;font-size:13px;white-space:nowrap;}
.action-card h3 {margin:8px 0 2px;padding:0;font-size:24px;}

.action-card .signal {font-size:19px;font-weight:650;margin:8px 0 14px;color:var(--accent);}

.action-card .owner {font-size:13px;font-weight:650;margin:12px 0 5px;}

.action-card p {margin:0;font-size:14px;line-height:1.55;}

.snapshot {display:inline-block;border-radius:5px;padding:4px 10px;background:#F2F0FF;color:#574B88;font-size:12px;}

.st-key-kpi_over [data-testid="stMetric"],

.st-key-kpi_risk [data-testid="stMetric"],

.st-key-kpi_value [data-testid="stMetric"] {background:#F7F4FF;border-color:#E1D9F0;border-top:4px solid #8271DD;}

[data-baseweb="tab"] {font-weight:650;}

h1 {letter-spacing:-0.035em;} h2 {letter-spacing:-0.02em;}

.dashboard-header {display:flex;align-items:center;justify-content:space-between;gap:16px;margin-bottom:0;}

.dashboard-header h1 {margin:0;padding:0;font-size:1.8rem;line-height:1.2;min-width:0;overflow-wrap:break-word;}

.dashboard-header img {width:90px;height:auto;flex-shrink:0;}

@media(max-width:700px) {.dashboard-header {align-items:flex-start;gap:16px;} .dashboard-header h1 {font-size:1.6rem;} .dashboard-header img {width:85px;margin-top:6px;}}


[data-testid="stMainBlockContainer"] > [data-testid="stVerticalBlock"] {gap:0.5rem;}
[data-testid="stTabPanel"] > [data-testid="stVerticalBlock"] {gap:0.5rem;}
[class*="st-key-alert_row_"] [data-testid="stVerticalBlock"] {gap:0.2rem;}
[data-testid="stMetricValue"] {font-size:1.35rem;}
.st-key-contracted_box {background:#f5f8f9;border:1px solid #e1e9ec;border-radius:8px;}
.st-key-contracted_box [data-testid="stMetric"] {border:0;padding-bottom:0;}
.st-key-contracted_box .contract-dollars {color:#7B7E85;font-style:italic;font-size:12px;padding:0 14px 10px;}

[data-testid="stMainBlockContainer"] {padding-top:3.2rem;padding-left:2rem;padding-right:2rem;}
[data-testid="stMainBlockContainer"] > [data-testid="stVerticalBlock"],
[data-testid="stTabPanel"] > [data-testid="stVerticalBlock"] {gap:0.25rem;}
[data-testid="stTabPanel"] h3, [data-testid="stTabPanel"] h4 {padding-top:0;padding-bottom:0;font-size:1.15rem;}
[data-testid="stMetric"] {padding:8px 10px;min-height:80px;}
[data-testid="stMetricLabel"] p {font-size:12px;line-height:1.25;}
[class*="st-key-alert_row_"] summary {padding:5px 8px;min-height:32px;}
[class*="st-key-alert_row_"] button {min-height:32px;padding:3px 8px;}
.st-key-contracted_box [data-testid="stVerticalBlock"] {gap:0;}
.st-key-contracted_box [data-testid="stMetric"] {min-height:62px;}
.st-key-contracted_box .contract-dollars {padding:0 10px 5px;font-size:11px;}
.st-key-contracted_box {padding-bottom:8px;}

[data-testid="stMainBlockContainer"] > [data-testid="stVerticalBlock"],
[data-testid="stTabPanel"] > [data-testid="stVerticalBlock"] {gap:1rem;}
[data-testid="stTabPanel"] h3, [data-testid="stTabPanel"] h4 {line-height:1.4;padding:8px 0;margin:0;}
.st-key-kpi_over [data-testid="stMetric"],
.st-key-kpi_risk [data-testid="stMetric"],
.st-key-kpi_value [data-testid="stMetric"],
.st-key-summary_used [data-testid="stMetric"],
.st-key-summary_utilization [data-testid="stMetric"],
.st-key-contracted_box {height:160px;box-sizing:border-box;}
.st-key-kpi_over [data-testid="stMetricLabel"],
.st-key-kpi_risk [data-testid="stMetricLabel"],
.st-key-kpi_value [data-testid="stMetricLabel"],
.st-key-summary_used [data-testid="stMetricLabel"],
.st-key-summary_utilization [data-testid="stMetricLabel"],
.st-key-contracted_box [data-testid="stMetricLabel"] {min-height:42px;align-items:flex-start;}
.st-key-contracted_box [data-testid="stMetric"] {height:auto;min-height:0;}
.st-key-contracted_box .contract-dollars {padding-top:8px;}
.st-key-contracted_box {min-height:160px;flex-shrink:0;}
.chart-labels {display:flex;justify-content:center;gap:28px;margin-top:-6px;font-size:13px;}
.chart-labels .contract-label {color:#8E8A96;}
.chart-labels .usage-label {color:#6550B5;}
[role="dialog"] {border-radius:24px;border:1px solid #E3DDF2;box-shadow:0 18px 55px rgba(73,54,122,.20);overflow:hidden;}
[role="dialog"]::before {content:"";display:block;height:5px;background:#6550B5;margin:-1rem -1rem 1rem;}
[role="dialog"] h2 {color:#20313B;letter-spacing:-.02em;}
.intro-eyebrow {color:#6550B5;font-size:12px;font-weight:750;letter-spacing:.08em;text-transform:uppercase;margin-bottom:6px;}
.intro-copy {color:#3E4B53;font-size:16px;line-height:1.65;margin:0 0 12px;}
.intro-detail {background:#F7F4FF;border-radius:10px;padding:12px 14px;color:#574B70;font-size:14px;line-height:1.55;margin-bottom:10px;}
.assumption-label {position:relative;display:inline-block;margin:4px 0 -8px;color:#24343D;font-size:14px;cursor:help;outline:none;}
.assumption-label::after {content:attr(data-help);position:absolute;left:0;top:calc(100% + 7px);width:232px;padding:9px 11px;border:1px solid #DAD6CD;border-radius:8px;background:#FFFFFF;color:#3E4B53;font-size:12px;line-height:1.4;box-shadow:0 7px 20px rgba(45,37,68,.14);opacity:0;visibility:hidden;transform:translateY(-3px);transition:opacity .16s ease,transform .16s ease,visibility .16s;z-index:1000;pointer-events:none;}
.assumption-label:hover::after,.assumption-label:focus::after {opacity:1;visibility:visible;transform:translateY(0);}
.calculation-table {width:100%;border-collapse:separate;border-spacing:0;border:1px solid #DADDE1;border-radius:8px;overflow:visible;font-size:14px;}
.calculation-table th,.calculation-table td {padding:8px 10px;border-bottom:1px solid #E5E7EA;text-align:left;vertical-align:middle;}
.calculation-table th {background:#F7F8FA;color:#52616B;font-weight:500;}
.calculation-table tr:last-child td {border-bottom:0;}
.calculation-table th:first-child,.calculation-table td:first-child {width:62%;}
.calculation-metric {position:relative;display:inline-block;color:#24343D;cursor:help;outline:none;border-bottom:1px dotted #8B7BC4;}
.calculation-metric::after {content:attr(data-help);position:absolute;left:calc(100% + 12px);top:-9px;width:360px;padding:9px 11px;border:1px solid #DAD6CD;border-radius:8px;background:#FFFFFF;color:#3E4B53;font-size:12px;font-weight:400;line-height:1.45;box-shadow:0 7px 20px rgba(45,37,68,.14);opacity:0;visibility:hidden;transform:translateY(-3px);transition:opacity .16s ease,transform .16s ease,visibility .16s;z-index:1000;pointer-events:none;white-space:normal;}
.calculation-metric:hover::after,.calculation-metric:focus::after {opacity:1;visibility:visible;transform:translateY(0);}
.calculation-metric .info-mark {color:#7A858C;font-size:12px;margin-left:3px;}
.customer-kpi-card {box-sizing:border-box;min-height:80px;padding:8px 10px;background:#F5F8F9;border:1px solid #E1E9EC;border-radius:8px;}
.customer-kpi-label {min-height:30px;color:#24343D;font-size:12px;line-height:1.25;}
.customer-kpi-value {color:#24343D;font-size:1.35rem;line-height:1.3;white-space:nowrap;}
.customer-kpi-value .over-entitlement-value {color:#A33A45;font-weight:650;}
</style>""", unsafe_allow_html=True)





@st.cache_data(show_spinner=False, max_entries=3)

def parse_source(data):

    return load_workbook(data)


@st.cache_data(show_spinner=False, max_entries=6)
def create_monthly_report(report_payload_json):
    return generate_monthly_report(json.loads(report_payload_json))


@st.cache_data(show_spinner=False, max_entries=6)
def create_audit_workbook(audit_payload_json):
    return generate_audit_workbook(json.loads(audit_payload_json))


def assumption_label(label, definition):
    """Render a label whose definition appears on hover or keyboard focus."""
    st.markdown(
        f'<div class="assumption-label" tabindex="0" data-help="{escape(definition, quote=True)}">'
        f'{escape(label)}</div>', unsafe_allow_html=True)





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


@st.dialog("Welcome to Usage & Billing Dashboard", width="medium")
def show_intro():
    st.markdown('<div class="intro-eyebrow">Product Finance</div>', unsafe_allow_html=True)
    st.markdown(
        '<p class="intro-copy">Use this dashboard to track customer credit consumption against contracted '
        'entitlements and identify accounts that need immediate attention.</p>',
        unsafe_allow_html=True)
    st.markdown(
        '<div class="intro-detail">Upload a source workbook containing customer contracts and usage to date. '
        'The application validates the data, flags portfolio exceptions, estimates commercial exposure, and '
        'provides account-level details and suggested follow-up actions.</div>',
        unsafe_allow_html=True)
with st.sidebar:

    st.subheader("Source Workbook")

    upload = st.file_uploader("Upload a replacement workbook", type=["xlsx"])

    source_path = st.text_input("Local workbook path", str(DEFAULT_SOURCE))

    st.button("Refresh source", help="Rereads the local file. For an upload, remove it and upload the new export.")

    st.caption("Manual refresh only. An uploaded file takes precedence over the local path.")

    st.subheader("Monitoring Assumptions")

    defaults = Rules()

    assumption_label("Projected utilization threshold (%)",
        "Flags an active account for early exhaustion risk when projected contract usage strictly exceeds this percentage of entitlement.")
    projected = st.number_input("Projected utilization threshold (%)", 100.0, 500.0,
        defaults.projected_utilization_threshold * 100, 5.0, label_visibility="collapsed")

    assumption_label("Early exhaustion days",
        "Flags an active account when its estimated exhaustion date is more than this many days before contract end.")
    early = st.number_input("Early exhaustion days", 0, 365, defaults.early_exhaustion_days, 5,
        label_visibility="collapsed")

    assumption_label("Underutilization threshold (%)",
        "Flags projected usage below this percentage after at least 25% of the contract term has elapsed.")
    under = st.number_input("Underutilization threshold (%)", 0.0, 100.0,
        defaults.underutilization_threshold * 100, 5.0, label_visibility="collapsed")

    assumption_label("Usage acceleration threshold (%)",
        "Flags trailing 30-day average usage when it strictly exceeds the prior 30-day average by this percentage and meets the history and 1,000-credit materiality checks.")
    acceleration = st.number_input("Usage acceleration threshold (%)", 0.0, 1000.0,
        defaults.usage_acceleration_threshold * 100, 5.0, label_visibility="collapsed")

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

if "intro_seen" not in st.session_state:
    # Mark it before opening so dismissing with the X does not make it return
    # on the next widget rerun. A new browser session receives it again.
    st.session_state["intro_seen"] = True
    show_intro()

read_at = datetime.now(timezone.utc)

age_days = (pd.Timestamp.now().normalize() - as_of).days

snapshot_label = f"Snapshot · {as_of:%b %d, %Y}" + (" · Refresh recommended" if age_days > 3 else " · Check future date" if age_days < 0 else "")

with st.sidebar:
    st.caption(snapshot_label)

with st.sidebar.expander("Source Information"):

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

    st.markdown("#### Priority Alerts")


    if "priority_resolutions" not in st.session_state:
        st.session_state["priority_resolutions"] = {}

    def change_resolution(resolution_id):
        resolutions = st.session_state["priority_resolutions"]
        if resolution_id in resolutions:
            resolutions.pop(resolution_id)
        else:
            resolutions[resolution_id] = datetime.now().astimezone().isoformat()

    urgent = accounts.loc[accounts.status.isin(["OVER ENTITLEMENT", "EARLY EXHAUSTION RISK"])].copy()
    # Display existing engine outputs only; resolution never changes financial status.
    def dollar_signal(item):
        if item["status"] == "OVER ENTITLEMENT":
            return item["estimated_overage_value"], "current overage"
        if item["status"] == "UNDERUTILIZING":
            return item["projected_unused_contract_value"], "projected unused contract value"
        return item["projected_overage_value"], "projected overage"

    items = []
    for item in urgent.to_dict("records"):
        amount, label = dollar_signal(item)
        resolution_id = "|".join([source_hash, item["customer_id"], str(item["contract_start"])])
        item.update(display_amount=amount, display_label=label, resolution_id=resolution_id,
                    resolved_at=st.session_state["priority_resolutions"].get(resolution_id))
        items.append(item)
    items.sort(key=lambda item: (bool(item["resolved_at"]),
                                -item["display_amount"] if pd.notna(item["display_amount"]) else float("inf"),
                                item["customer_id"]))
    if not items:
        st.success("No accounts require follow-up under the current rules.")
    for item in items:
        amount_text = fmt(item["display_amount"], "usd")
        row_id = sha256(item["resolution_id"].encode()).hexdigest()[:16]
        with st.container(key="alert_row_" + row_id):
            body, controls = st.columns([5, 1], gap="small", vertical_alignment="top")
            with controls:
                state = "done" if item["resolved_at"] else "open"
                with st.container(key="resolve_" + state + "_" + row_id):
                    st.button("Resolved" if item["resolved_at"] else "Resolve?",
                              icon=":material/check_circle:" if item["resolved_at"] else None,
                              key="resolved_" + row_id, width="stretch",
                              on_click=change_resolution, args=(item["resolution_id"],),
                              help="Click again to reopen. Review marks last for this browser session.")
                if item["resolved_at"]:
                    st.caption(datetime.fromisoformat(item["resolved_at"]).strftime("%b %d, %Y"))
            with body:
                title = f'{item["customer_id"]} · {amount_text} {item["display_label"]}'
                with st.expander(title, expanded=False):
                    st.markdown(f'**{item["status"].title()}**')
                    st.write(item["status_reason"])
                    if item["forecast_available"]:
                        st.caption(f'Projected utilization: {fmt(item["projected_utilization_pct"], "pct")}')
                    st.markdown(f'**{item["recommended_owner"]}**')
                    st.write(item["recommended_action"])
                    st.button("Review " + item["customer_id"], key="review_" + item["customer_id"],
                              on_click=open_account, args=(item["customer_id"],))



with dashboard_tab:

    summary_title, view_col, start_col, end_col = st.columns([2, 1, 1.2, 1.2])
    summary_title.markdown("### Usage Summary")

    view = view_col.selectbox("View", ["Annual", "Quarterly", "Monthly"], key="summary_view")

    years = list(range(min(contracts.contract_start.min().year, as_of.year), max(contracts.contract_end.max().year, as_of.year) + 1))

    if view == "Annual":
        year = start_col.selectbox("Year", years, index=years.index(as_of.year), key="summary_year")
        end_col.empty()
        period_start = pd.Timestamp(year, 1, 1)
        period_end = pd.Timestamp(year + 1, 1, 1)
    elif view == "Quarterly":
        first = pd.Period(contracts.contract_start.min(), freq="Q")
        last = pd.Period(contracts.contract_end.max() - pd.Timedelta(days=1), freq="Q")
        periods = list(pd.period_range(first, last, freq="Q"))
        default_start = pd.Period(f"{as_of.year}Q1", freq="Q")
        default_end = pd.Period(as_of, freq="Q")
        start_q = start_col.selectbox("Start Quarter", periods,
            index=periods.index(default_start) if default_start in periods else 0,
            format_func=lambda p: f"Q{p.quarter} {str(p.year)[-2:]}", key="summary_start_quarter")
        valid_ends = [p for p in periods if p >= start_q]
        end_q = end_col.selectbox("End Quarter", valid_ends,
            index=valid_ends.index(default_end) if default_end in valid_ends else len(valid_ends) - 1,
            format_func=lambda p: f"Q{p.quarter} {str(p.year)[-2:]}", key="summary_end_quarter")
        period_start, period_end = start_q.start_time.normalize(), (end_q + 1).start_time.normalize()
    else:
        first = pd.Period(contracts.contract_start.min(), freq="M")
        last = pd.Period(contracts.contract_end.max() - pd.Timedelta(days=1), freq="M")
        periods = list(pd.period_range(first, last, freq="M"))
        default_start = pd.Period(f"{as_of.year}-01", freq="M")
        default_end = pd.Period(as_of, freq="M")
        start_m = start_col.selectbox("Start Month", periods,
            index=periods.index(default_start) if default_start in periods else 0,
            format_func=lambda p: p.start_time.strftime("%b %y"), key="summary_start_month")
        valid_ends = [p for p in periods if p >= start_m]
        end_m = end_col.selectbox("End Month", valid_ends,
            index=valid_ends.index(default_end) if default_end in valid_ends else len(valid_ends) - 1,
            format_func=lambda p: p.start_time.strftime("%b %y"), key="summary_end_month")
        period_start, period_end = start_m.start_time.normalize(), (end_m + 1).start_time.normalize()



    period_events = usage.loc[(usage.date >= period_start) & (usage.date < period_end)]

    if period_events.empty:

        st.info("No usage records available for this period. Select another period or upload historical data.")

    else:

        snapshot_events = usage.loc[usage.date < period_end]

        snapshot_date = snapshot_events.date.max()

        snapshot_accounts = calculate_accounts(contracts, snapshot_events, rules)

        summary = snapshot_accounts.loc[(snapshot_accounts.contract_start < period_end) &

                                        (snapshot_accounts.contract_end > period_start)]

        entitlement = credit_sum(summary.annual_entitlement_credits)

        consumed = credit_sum(summary.total_credits_used)

        st.caption(f"{period_start:%b %d, %Y} – {period_end - pd.Timedelta(days=1):%b %d, %Y} · "

                   f"Usage available through {snapshot_date:%b %d, %Y}.")

        kpis = [

            ("Customers Over Entitlement", str(summary.status.eq(STATUSES[0]).sum())),

            ("Early Exhaustion Risk", str(summary.status.eq(STATUSES[1]).sum())),

            ("Estimated Overage Value", f"${summary.estimated_overage_value.sum():,.2f}"),

        ]

        st.markdown("**Portfolio Credit Consumption & Projection**")
        chart_data = portfolio_chart(summary, usage, period_start, period_end, snapshot_date)
        chart_data = chart_data.drop(columns=["Contract pacing"])
        chart_long = chart_data.melt("date", var_name="Series", value_name="Credits")
        max_credits = chart_long.Credits.max(skipna=True)
        axis_max = max(200_000, int(max_credits / 200_000 + 1) * 200_000)
        x_axis = alt.X("date:T", title=None,
                       axis=alt.Axis(format="%b %y", tickCount="month", labelAngle=0))
        y_axis = alt.Y("Credits:Q", title="Credits",
                       scale=alt.Scale(domain=[0, axis_max]),
                       axis=alt.Axis(values=list(range(0, axis_max + 1, 200_000)), format="~s"))
        series_domain = ["Total contracted credits", "Projected contracted credits",
                         "Credits consumed", "Projected consumption"]
        series_colors = ["#B8B5BF", "#C9C6CE", "#8271DD", "#8271DD"]
        series_dashes = [[1, 0], [6, 4], [1, 0], [6, 4]]
        contract_data = chart_long.loc[chart_long.Series.eq("Total contracted credits")]
        contract_line = alt.Chart(contract_data).mark_line(
            strokeWidth=3, interpolate="step-after", color="#B8B5BF").encode(
                x=x_axis, y=y_axis,
                tooltip=[alt.Tooltip("date:T", title="Date"),
                         alt.Tooltip("Credits:Q", format=",.0f")])
        contract_projection_data = chart_long.loc[
            chart_long.Series.eq("Projected contracted credits") & chart_long.Credits.notna()]
        contract_projection = alt.Chart(contract_projection_data).mark_line(
            strokeWidth=3, color="#C9C6CE", strokeDash=[6, 4]).encode(
                x=x_axis, y=y_axis,
                tooltip=[alt.Tooltip("date:T", title="Date"),
                         alt.Tooltip("Credits:Q", format=",.0f")])
        usage_data = chart_long.loc[chart_long.Series.isin(
            ["Credits consumed", "Projected consumption"])]
        usage_lines = alt.Chart(usage_data).mark_line(strokeWidth=2.5).encode(
            x=x_axis, y=y_axis,
            color=alt.Color("Series:N", scale=alt.Scale(domain=series_domain, range=series_colors),
                            legend=None),
            strokeDash=alt.StrokeDash("Series:N", scale=alt.Scale(
                domain=series_domain, range=series_dashes), legend=None),
            tooltip=[alt.Tooltip("date:T", title="Date"), "Series:N",
                     alt.Tooltip("Credits:Q", format=",.1f")])
        chart_layers = [contract_line, contract_projection, usage_lines]
        projection_start = snapshot_date + pd.Timedelta(days=1)
        if projection_start < period_end and summary.forecast_available.all():
            region = pd.DataFrame({"start": [projection_start], "end": [period_end],
                                   "middle": [projection_start + (period_end - projection_start) / 2],
                                   "label_y": [axis_max * 0.97]})
            shade = alt.Chart(region).mark_rect(color="#BFC1C5", opacity=0.16).encode(
                x="start:T", x2="end:T")
            annotation = alt.Chart(region).mark_text(
                text="Projection", fontStyle="italic", color="#777B83",
                align="center", baseline="top", dy=8).encode(
                    x="middle:T", y=alt.Y("label_y:Q", scale=alt.Scale(domain=[0, axis_max])))
            chart_layers = [shade, contract_line, contract_projection, usage_lines, annotation]
        chart = alt.layer(*chart_layers).properties(height=340)
        st.altair_chart(chart, width="stretch")
        st.markdown('<div class="chart-labels"><span class="contract-label">&#8226; Total contracted credits</span>'
                    '<span class="usage-label">&#8226; Cumulative Consumption</span></div>',
                    unsafe_allow_html=True)
        average_addition = chart_data.attrs.get("average_monthly_contracted_addition")
        st.caption(
            "**How projections are calculated**  \n"
            f"1. **Total contracted credits:** credits contracted at the snapshot + the average credits added per "
            f"completed month ({fmt(average_addition)} per month) × remaining months.  \n"
            "2. **Cumulative consumption:** credits used at the snapshot + each customer's trailing 30-day average "
            "daily usage × its remaining active contract days.  \n"
            "Dashed lines show projections. Consumption is not projected when daily usage records are incomplete."
        )

        flags, utilization = st.columns(2, gap="medium")
        with flags:
            st.markdown("**Contracts Flags**")
            for col, (label, value), card_key in zip(st.columns(3), kpis, ["kpi_over", "kpi_risk", "kpi_value"]):
                with col.container(key=card_key):
                    st.metric(label, value)
        with utilization:
            st.markdown("**Utilization Statistics**")
            credit_col, used_col, utilization_col = st.columns(3)
            with credit_col.container(key="contracted_box"):
                st.metric("Total Contracted Credits", f"{entitlement:,.0f}")
                st.markdown(f'<div class="contract-dollars">${credit_sum(summary.annual_contract_value_usd):,.2f} contract value</div>',
                            unsafe_allow_html=True)
            with used_col.container(key="summary_used"):
                st.metric("Credits Consumed in Period", f"{credit_sum(period_events.credits_used):,.1f}")
            with utilization_col.container(key="summary_utilization"):
                st.metric("Contract Utilization %", f"{consumed / entitlement:.1%}" if entitlement else "N/A")

    st.subheader("Finance Action Queue")

    st.caption("Urgent accounts first. Select a row to open Customer Details.")

    filter_cols = st.columns([1, 1.6])

    exceptions_only = filter_cols[0].toggle("Exceptions only", value=True, key="exceptions_only",

        help="Hides only accounts assigned No Action. Existing priorities and statuses are unchanged.")

    with filter_cols[1].expander("Filter Queue"):

        customers = st.multiselect("Customer", sorted(accounts.customer_id), key="filter_customer")

        status_help = (
            f"**Over Entitlement:** credits used are at least 100% of entitlement.  \n"
            f"**Early Exhaustion Risk:** active, not already over entitlement, and projected utilization is above "
            f"{projected:.0f}% or estimated exhaustion is more than {early} days before contract end.  \n"
            f"**Underutilizing:** at least 25% of the contract has elapsed and projected utilization is below {under:.0f}%.  \n"
            "**On Track:** no exception rule applies.  \n"
            "**Data Review:** expected daily usage records are missing, so pacing and forecasts are withheld; "
            "an account already over entitlement remains Over Entitlement.  \n"
            "**Not Started:** the analysis date is before the contract begins, so forecast metrics are unavailable."
        )
        statuses = st.multiselect("Status", STATUSES, key="filter_status", help=status_help,
                                  format_func=lambda status: status.title())

        priority_help = (
            "**P1 - Urgent:** over entitlement; Billing should reconcile current exposure.  \n"
            "**P2 - High:** early exhaustion risk or incomplete usage data requiring review.  \n"
            "**P3 - Medium:** underutilizing, usage-accelerating On Track, or expired On Track accounts requiring follow-up.  \n"
            "**P4 - Routine:** On Track or Not Started accounts with no immediate action."
        )
        priorities = st.multiselect("Priority", sorted(accounts.priority.unique()),
                                    key="filter_priority", help=priority_help,
                                    format_func=lambda priority: priority.replace(" · ", " - "))

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

            return f'Data review · {row["utilization_pct"]:.1%} currently consumed · missing daily usage records'

        if row["status"] == "NOT STARTED":

            return f'Not started · {row["utilization_pct"]:.1%} currently consumed · forecast unavailable'

        if row["status"] == "OVER ENTITLEMENT":

            text = f'Over entitlement · {row["utilization_pct"]:.1%} currently consumed'

        elif row["status"] == "EARLY EXHAUSTION RISK":

            text = (f'Early exhaustion risk · {row["utilization_pct"]:.1%} currently consumed · '
                    f'{row["projected_utilization_pct"]:.1%} projected')

        elif row["status"] == "UNDERUTILIZING":

            text = (f'Underutilizing · {row["utilization_pct"]:.1%} currently consumed · '
                    f'{row["projected_utilization_pct"]:.1%} projected')

        else:

            text = (f'On track · {row["utilization_pct"]:.1%} currently consumed · '
                    f'{row["projected_utilization_pct"]:.1%} projected')

        if row["missing_usage_days"]:

            text += " · Incomplete feed"

        return text + (" · Accelerating" if row["usage_accelerating"] else "")

    

    compact = pd.DataFrame({

        "Priority": filtered.priority.str.replace(" · ", " - ", regex=False),

        "Status": filtered.status.str.title(),

        "Customer": filtered.customer_id,

        "Why review": filtered.apply(review_signal, axis=1) if len(filtered) else pd.Series(dtype=str),

        "Current overage": filtered.estimated_overage_value,

        "Pacing Index": filtered.pacing_index,

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

        priority_code = str(row["Priority"]).split()[0]
        wash, ink = priority_palette.get(priority_code, ("#fff", "#20313b"))

        return [f"background-color:{wash};color:{ink};" + ("font-weight:650;" if col in ["Priority", "Status", "Why review"] else "") for col in row.index]

    compact_style = compact.style.apply(color_queue, axis=1)

    st.dataframe(compact_style, hide_index=True, width="stretch", height=min(335, 38 + 36 * max(len(compact), 1)),

        column_config={

            "Priority": st.column_config.TextColumn(width=125),

            "Status": st.column_config.TextColumn(width=175),

            "Customer": st.column_config.TextColumn(width=100),

            "Why review": st.column_config.TextColumn(width="large"),

            "Current overage": st.column_config.NumberColumn(format="dollar", width=150,

                help="Current overage only. Full forecast exposure is in the account details."),

            "Pacing Index": st.column_config.NumberColumn(format="%.2f", width=115,

                help="Pacing Index = Utilization % ÷ Contract Elapsed %. 1.00 is on pace; above 1.00 is ahead of pace; below 1.00 is behind pace."),

        }, on_select=select_queue_account, selection_mode="single-row", key=queue_key)

    st.caption(f"{len(filtered)} of {len(accounts)} accounts shown · P1 urgent / P2 high / P3 medium / P4 routine")

    if filtered.empty:

        st.info("No accounts match. Turn off Exceptions only or clear filters to broaden the queue.")

    with st.expander("Full Queue Metrics"):

        st.dataframe(styled, hide_index=True, width="stretch", column_config=config)

    

    audit_payload = build_audit_workbook_payload(
        contracts, usage, filtered, rules, as_of, source_name, source_hash)
    try:
        audit_payload_json = json.dumps(audit_payload, sort_keys=True, separators=(",", ":"))
        st.download_button("Download Queue & Formula Audit (XLSX)",
                           lambda: create_audit_workbook(audit_payload_json),
                           f"finance_action_queue_{as_of:%Y%m%d}.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as exc:
        logging.exception("Excel audit export failed")
        st.error(f"Excel audit export is unavailable: {exc}")



with details_tab:

    

    st.subheader("Customer Detail")

    if st.session_state.get("detail_customer") not in set(accounts.customer_id):

        st.session_state["detail_customer"] = accounts.iloc[0].customer_id

    selected = st.selectbox("Select customer", sorted(accounts.customer_id.tolist()), key="detail_customer",

        help="Search by customer ID. Queue row selection opens the same account here.")

    m = accounts.loc[accounts.customer_id.eq(selected)].iloc[0].to_dict()

    st.markdown(f"**{m['status']}** · {m['priority']} · {m['contract_lifecycle'].title()}" + (" · USAGE ACCELERATING" if m["usage_accelerating"] else ""))

    st.caption(f"Contract: {fmt(m['contract_start'], 'date')} to {fmt(m['contract_end'], 'date')} (end exclusive) · {int(m['term_months'])} months")

    # Put the decision ahead of supporting metrics.

    with st.container(border=True):

        st.markdown("**Why This Account Is Flagged**" if m["status"] != "ON TRACK" or m["usage_accelerating"] else "**Monitoring Assessment**")

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

    elapsed_contract_months = m["contract_elapsed_pct"] * m["term_months"]
    key_metrics = [
        ("Remaining Credit Available", f'{m["remaining_credits"]:,.1f}',
         "Total credit entitlement minus credits used. A negative value indicates consumption above entitlement."),

        ("Contract Months Elapsed", f'{elapsed_contract_months:.1f} / {int(m["term_months"])}',
         "Elapsed contract days expressed as the equivalent share of the contractual month term."),

        ("Utilization", fmt(m["utilization_pct"], "pct"),
         "Credits used divided by total credit entitlement."),

        ("Projected Utilization",
         "N/A" if m["contract_lifecycle"] == "NOT STARTED" else fmt(m["projected_utilization_pct"], "pct"),
         "Projected total contract usage divided by total credit entitlement."),

    ]

    metric_columns = st.columns(5)
    used_class = "over-entitlement-value" if m["total_credits_used"] > m["annual_entitlement_credits"] else ""
    with metric_columns[0]:
        st.markdown(
            '<div class="customer-kpi-card">'
            '<div class="customer-kpi-label" title="Cumulative credits consumed through the analysis date compared with contracted annual credit entitlement.">'
            'Credits Used / Total Credit Entitlement&nbsp; ⓘ</div>'
            f'<div class="customer-kpi-value"><span class="{used_class}">{m["total_credits_used"]:,.1f}</span>'
            f' / {m["annual_entitlement_credits"]:,.0f}</div></div>', unsafe_allow_html=True)

    for col, (label, value, help_text) in zip(metric_columns[1:], key_metrics):

        col.metric(label, value, help=help_text)

    with st.expander("Contract, Balance & Exhaustion Details"):

        fields = [

            ("Annual Contract Value", "annual_contract_value_usd", "usd",
             "Contract value from the source workbook for the annual term."),

            ("Entitlement", "annual_entitlement_credits", "number",
             "Total credits contracted for the annual term."),

            ("Remaining Credits", "remaining_credits", "number",
             "Entitlement minus credits used. A negative balance means usage is over entitlement."),

            ("Contract Elapsed", "contract_elapsed_pct", "pct",
             "Elapsed contract days divided by total contract days, measured through the analysis cutoff."),

            ("Pacing Index", "pacing_index", "ratio",
             "Utilization percentage divided by contract elapsed percentage. 1.00 is on pace; above 1.00 is ahead; below 1.00 is behind."),

            ("Days to Exhaustion", "days_to_exhaustion", "number",
             "Remaining credits divided by trailing 30-day average daily usage. Already exhausted accounts show zero."),

            ("Exhaustion Date", "estimated_exhaustion_date", "date",
             "First actual date entitlement was reached; otherwise the analysis date plus rounded-up days to exhaustion."),

            ("Current Overage", "estimated_overage_value", "usd",
             "Maximum of credits used minus entitlement and zero, multiplied by contract value per credit. This is an estimate, not an invoice."),

        ]

        detail_rows = "".join(
            '<tr><td><span class="calculation-metric" tabindex="0" data-help="' +
            escape(definition, quote=True) + '">' + escape(label) +
            '<span class="info-mark">ⓘ</span></span></td><td>' +
            escape(fmt(m[key], kind)) + '</td></tr>'
            for label, key, kind, definition in fields
        )
        st.markdown(
            '<table class="calculation-table"><thead><tr><th>Metric</th><th>Value</th></tr></thead><tbody>' +
            detail_rows + '</tbody></table>', unsafe_allow_html=True)

        st.caption("Negative balance means overconsumption. Exhausted accounts show the first actual exhaustion date. Future dates assume the recent rate continues, even beyond contract end.")

    

    left, right = st.columns([1.3, 1])

    with left:

        st.markdown("**Cumulative Usage vs. Contract Pacing**")

        chart_data = pacing_chart_data(m, usage).melt("date", var_name="Series", value_name="Credits")

        chart = alt.Chart(chart_data).mark_line(strokeWidth=2).encode(

            x=alt.X("date:T", title=None), y=alt.Y("Credits:Q", title="Cumulative credits"),

            color=alt.Color("Series:N", scale=alt.Scale(domain=["Actual cumulative usage", "Expected linear pacing"], range=["#8271DD", "#B6B3C0"]), legend=alt.Legend(orient="top", title=None)),

            tooltip=[alt.Tooltip("date:T", title="Boundary date"), "Series:N", alt.Tooltip("Credits:Q", format=",.1f")])

        st.altair_chart(chart.properties(height=310), width="stretch")

        st.caption("Actual usage stops at the analysis cutoff; expected pacing continues to the full contract entitlement.")

        if m["missing_usage_days"]:

            st.caption("Incomplete feed: this is observed cumulative usage only, not a complete consumption ledger.")

    with right:

        st.markdown("**Daily Usage Trend**")

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

    

    with st.expander("Recent Usage & Commercial Impact", expanded=False):

        detail_fields = {

            "trailing_30_day_credits": "Trailing 30-day credits", "prior_30_day_credits": "Prior 30-day credits",

            "trailing_30_day_average_daily_usage": "Trailing average daily usage", "prior_30_day_average_daily_usage": "Prior average daily usage",

            "usage_growth_30_day_pct": "30-day usage growth", "forecast_remaining_usage": "Forecast remaining credits",

            "projected_total_contract_usage": "Projected total credits", "projected_overage_value": "Total projected overage proxy",

            "projected_unused_contract_value": "Projected unused contract value", "implied_contract_value_per_credit": "Implied value per credit",

        }

        detail_definitions = {
            "trailing_30_day_credits": "Sum of daily credits in the 30 calendar days ending on the analysis date.",
            "prior_30_day_credits": "Sum of daily credits in the 30 calendar days immediately before the trailing period.",
            "trailing_30_day_average_daily_usage": "Trailing-period credits divided by eligible contract days in that period, up to 30 days. Unavailable when expected daily records are missing.",
            "prior_30_day_average_daily_usage": "Prior-period credits divided by eligible contract days in that period, up to 30 days. Unavailable when expected daily records are missing.",
            "usage_growth_30_day_pct": "Trailing 30-day credits divided by prior 30-day credits, minus 1. It is unavailable when the prior period is zero but recent usage is positive, or when daily records are incomplete.",
            "forecast_remaining_usage": "Trailing average daily usage multiplied by remaining contract days for an active contract. Forecasts are withheld when daily records are incomplete or the contract has not started.",
            "projected_total_contract_usage": "Credits used to date plus forecast remaining credits.",
            "projected_overage_value": "Maximum of projected total credits minus entitlement and zero, multiplied by implied value per credit. This is a commercial proxy, not an invoice amount.",
            "projected_unused_contract_value": "Maximum of entitlement minus projected total credits and zero, multiplied by implied value per credit. This is an adoption indicator, not a refund estimate.",
            "implied_contract_value_per_credit": "Annual contract value divided by annual credit entitlement.",
        }
        detail_rows = []
        for key, label in detail_fields.items():
            value = fmt(m[key], "pct" if key.endswith("pct") else "usd" if "value" in key else "number")
            detail_rows.append(
                '<tr><td><span class="calculation-metric" tabindex="0" '
                f'data-help="{escape(detail_definitions[key], quote=True)}">{escape(label)}</span></td>'
                f'<td>{escape(value)}</td></tr>')
        st.markdown(
            '<table class="calculation-table"><thead><tr><th>Metric</th><th>Value</th></tr></thead><tbody>'
            + "".join(detail_rows) + '</tbody></table>', unsafe_allow_html=True)

        st.caption(f"{m['growth_note']}. Total projected overage includes current overage; do not add the two. Unused contract value is an adoption/renewal indicator, not a refund or revenue-loss estimate.")

    

    

    with st.expander("Data & Calculation Audit"):

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

        st.markdown("**Selected Account: Full Calculated Record**")

        st.dataframe(pd.DataFrame([{"Field": k, "Value": str(v)} for k, v in m.items()]), hide_index=True, width="stretch")

        st.markdown("**Source Contract**")

        st.dataframe(contracts.loc[contracts.customer_id.eq(selected)], hide_index=True, width="stretch")

        st.markdown("**Source Daily Usage**")

        st.dataframe(usage.loc[usage.customer_id.eq(selected)], hide_index=True, width="stretch")

        st.download_button("Download monitoring assumptions (JSON)", json.dumps({"as_of": str(as_of.date()), "source_sha256": source_hash, "rules": asdict(rules)}, indent=2), "monitoring_assumptions.json", "application/json")



with dashboard_tab:
    st.markdown("### Monthly Report")
    st.caption(
        f"Create an editable PowerPoint for {as_of:%B %Y} with the portfolio outlook, "
        "top customers, open or resolved priority items, and the annual usage chart.")
    report_payload = build_monthly_report_payload(
        accounts, usage, as_of, source_name, source_hash,
        st.session_state.get("priority_resolutions", {}))
    report_payload_json = json.dumps(
        report_payload, sort_keys=True, separators=(",", ":"))
    try:
        st.download_button(
            "Generate Monthly Usage Report",
            lambda: create_monthly_report(report_payload_json),
            f"monthly_usage_report_{as_of:%Y_%m}.pptx",
            "application/vnd.openxmlformats-officedocument.presentationml.presentation",
            type="primary",
            key="generate_monthly_report",
        )
    except Exception as exc:
        logging.getLogger(__name__).exception("PowerPoint report generation failed")
        st.error(str(exc))
