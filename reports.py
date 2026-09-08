"""Deterministic PowerPoint report payload and local Artifact Tool bridge."""

from __future__ import annotations

import json
import os
import shutil
import subprocess
import tempfile
from pathlib import Path

import pandas as pd

from finance import credit_sum
from portfolio import portfolio_chart


def _finite(value):
    return None if pd.isna(value) else round(float(value), 6)


def _status_text(row):
    projected = row.get("projected_utilization_pct")
    if row["status"] == "OVER ENTITLEMENT":
        return f'{row["customer_id"]} is over entitlement at {row["utilization_pct"]:.1%} utilization.'
    if row["status"] == "EARLY EXHAUSTION RISK" and pd.notna(projected):
        return (f'{row["customer_id"]} has early exhaustion risk at '
                f'{row["utilization_pct"]:.1%} utilization and {projected:.1%} projected utilization.')
    if pd.notna(projected):
        return (f'{row["customer_id"]} is {row["status"].lower()} at '
                f'{row["utilization_pct"]:.1%} utilization and {projected:.1%} projected utilization.')
    return f'{row["customer_id"]} is {row["status"].lower()} at {row["utilization_pct"]:.1%} utilization.'


def build_monthly_report_payload(accounts, usage, as_of, source_name, source_hash,
                                 resolutions=None):
    """Build the report from calculated account rows and normalized usage only."""
    as_of = pd.Timestamp(as_of)
    year_start = pd.Timestamp(as_of.year, 1, 1)
    year_end = pd.Timestamp(as_of.year + 1, 1, 1)
    annual_accounts = accounts.loc[
        (accounts.contract_start < year_end) & (accounts.contract_end > year_start)
    ].copy()
    annual_chart = portfolio_chart(annual_accounts, usage, year_start, year_end, as_of)
    by_date = annual_chart.set_index("date")
    cutoff = min(as_of + pd.Timedelta(days=1), year_end)

    categories = []
    actual_contracts = []
    projected_contracts = []
    actual_usage = []
    projected_usage = []
    for month in range(1, 13):
        month_start = pd.Timestamp(as_of.year, month, 1)
        boundary = month_start + pd.DateOffset(months=1)
        month_end = boundary - pd.Timedelta(days=1)
        categories.append(month_start.strftime("%b %y"))

        if month_end <= as_of:
            contracted = credit_sum(annual_accounts.loc[
                annual_accounts.contract_start.le(month_end), "annual_entitlement_credits"])
            actual_contracts.append(contracted)
            projected_contracts.append(contracted if month == as_of.month else None)
        elif month == as_of.month:
            contracted = credit_sum(annual_accounts.loc[
                annual_accounts.contract_start.le(as_of), "annual_entitlement_credits"])
            actual_contracts.append(contracted)
            projected_contracts.append(contracted)
        else:
            actual_contracts.append(None)
            projected_contracts.append(_finite(by_date.loc[boundary, "Projected contracted credits"]))

        if boundary <= cutoff:
            consumed = _finite(by_date.loc[boundary, "Credits consumed"])
            actual_usage.append(consumed)
            projected_usage.append(consumed if month == as_of.month else None)
        elif month == as_of.month:
            consumed = credit_sum(usage.loc[
                (usage.date >= year_start) & (usage.date <= as_of), "credits_used"])
            actual_usage.append(consumed)
            projected_usage.append(consumed)
        else:
            actual_usage.append(None)
            projected_usage.append(_finite(by_date.loc[boundary, "Projected consumption"]))

    ytd_usage = credit_sum(usage.loc[
        usage.date.between(year_start, as_of), "credits_used"])
    year_end_consumption = projected_usage[-1] or actual_usage[-1]
    year_end_contracts = projected_contracts[-1] or actual_contracts[-1]
    entitlement = credit_sum(annual_accounts.annual_entitlement_credits)
    contract_value = credit_sum(annual_accounts.annual_contract_value_usd)
    total_used = credit_sum(annual_accounts.total_credits_used)
    utilization = total_used / entitlement if entitlement else None

    top_three_rows = (accounts.sort_values(
        ["annual_contract_value_usd", "customer_id"], ascending=[False, True])
        .head(3).to_dict("records"))
    top_three = [{
        "customer_id": row["customer_id"],
        "contract_value": _finite(row["annual_contract_value_usd"]),
        "status": row["status"],
        "utilization": _finite(row["utilization_pct"]),
        "projected_utilization": _finite(row["projected_utilization_pct"]),
        "tracking_sentence": _status_text(row),
    } for row in top_three_rows]

    resolutions = resolutions or {}
    urgent = accounts.loc[accounts.status.isin(
        ["OVER ENTITLEMENT", "EARLY EXHAUSTION RISK"])].copy()
    urgent["report_amount"] = urgent.apply(
        lambda row: row.estimated_overage_value if row.status == "OVER ENTITLEMENT"
        else row.projected_overage_value, axis=1)
    urgent = urgent.sort_values(["report_amount", "customer_id"], ascending=[False, True])
    issues = []
    for row in urgent.to_dict("records"):
        resolution_id = "|".join([source_hash, row["customer_id"], str(row["contract_start"])])
        resolved_at = resolutions.get(resolution_id)
        issues.append({
            "customer_id": row["customer_id"],
            "status": row["status"],
            "amount": _finite(row["report_amount"]),
            "owner": row["recommended_owner"],
            "resolution": (f'Resolved {pd.Timestamp(resolved_at):%b %d, %Y}'
                           if resolved_at else "Open"),
        })

    return {
        "title": "Monthly Usage Report",
        "month": as_of.strftime("%B %Y"),
        "as_of": as_of.strftime("%B %d, %Y"),
        "source_name": source_name,
        "source_hash": source_hash,
        "portfolio": {
            "ytd_usage": ytd_usage,
            "year_end_consumption": year_end_consumption,
            "year_end_contracts": year_end_contracts,
            "entitlement": entitlement,
            "contract_value": contract_value,
            "total_used": total_used,
            "utilization": utilization,
            "over_count": int(annual_accounts.status.eq("OVER ENTITLEMENT").sum()),
            "risk_count": int(annual_accounts.status.eq("EARLY EXHAUSTION RISK").sum()),
            "current_overage": credit_sum(annual_accounts.estimated_overage_value),
        },
        "top_three": top_three,
        "issues": issues,
        "chart": {
            "snapshot_month": as_of.month,
            "categories": categories,
            "actual_contracts": actual_contracts,
            "projected_contracts": projected_contracts,
            "actual_usage": actual_usage,
            "projected_usage": projected_usage,
        },
    }


def _runtime_paths():
    node = os.environ.get("PPTX_NODE") or shutil.which("node")
    module = os.environ.get("ARTIFACT_TOOL_MODULE")
    if not module:
        candidate = (Path.home() / ".cache" / "codex-runtimes" / "codex-primary-runtime" /
                     "dependencies" / "node" / "node_modules" / "@oai" / "artifact-tool" /
                     "dist" / "artifact_tool.mjs")
        if candidate.exists():
            module = str(candidate)
            bundled_node = (Path.home() / ".cache" / "codex-runtimes" / "codex-primary-runtime" /
                            "dependencies" / "node" / "bin" / "node.exe")
            if bundled_node.exists():
                node = str(bundled_node)
    if not node or not module or not Path(module).exists():
        raise RuntimeError(
            "PowerPoint export requires the local presentation runtime. "
            "Set PPTX_NODE and ARTIFACT_TOOL_MODULE, then restart Streamlit.")
    return node, module


def generate_monthly_report(payload):
    """Return an editable PPTX as bytes."""
    node, module = _runtime_paths()
    builder = Path(__file__).with_name("report_builder.mjs")
    logo = Path(__file__).parent / "assets" / "glean-logo.svg"
    with tempfile.TemporaryDirectory(prefix="usage-report-") as temp_dir:
        temp_path = Path(temp_dir)
        input_path = temp_path / "report.json"
        output_path = temp_path / "monthly-usage-report.pptx"
        input_path.write_text(json.dumps(payload, allow_nan=False), encoding="utf-8")
        env = os.environ.copy()
        env["ARTIFACT_TOOL_MODULE"] = module
        result = subprocess.run(
            [node, str(builder), str(input_path), str(output_path), str(logo)],
            capture_output=True, text=True, timeout=60, env=env,
        )
        if result.returncode:
            detail = (result.stderr or result.stdout).strip().splitlines()[-1:]
            raise RuntimeError("PowerPoint report generation failed" +
                               (f": {detail[0]}" if detail else "."))
        return output_path.read_bytes()
