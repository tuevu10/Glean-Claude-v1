"""Traceable Excel audit export backed by the local Artifact Tool runtime."""

from __future__ import annotations

from dataclasses import asdict
import json
import os
import shutil
import subprocess
import tempfile
from pathlib import Path

import pandas as pd


def _scalar(value):
    if pd.isna(value):
        return None
    if isinstance(value, pd.Timestamp):
        return value.strftime("%Y-%m-%d")
    if hasattr(value, "item"):
        return value.item()
    return value


def build_audit_workbook_payload(contracts, usage, filtered_accounts, rules, as_of,
                                 source_name, source_hash):
    """Serialize normalized source rows and workbook formula inputs only."""
    contract_fields = ["customer_id", "contract_start", "term_months",
                       "annual_entitlement_credits", "annual_contract_value_usd",
                       "source_excel_row"]
    usage_fields = ["date", "customer_id", "credits_used", "source_excel_row"]
    customer_ids = filtered_accounts["customer_id"].astype(str).tolist()
    return {
        "as_of": pd.Timestamp(as_of).strftime("%Y-%m-%d"),
        "source_name": source_name,
        "source_hash": source_hash,
        "rules": {key: _scalar(value) for key, value in asdict(rules).items()},
        "contracts": [
            {field: _scalar(row.get(field)) for field in contract_fields}
            for row in contracts.to_dict("records")
        ],
        "usage": [
            {field: _scalar(row.get(field)) for field in usage_fields}
            for row in usage.to_dict("records")
        ],
        "queue_customer_ids": customer_ids,
    }


def _runtime_paths():
    node = os.environ.get("XLSX_NODE") or shutil.which("node")
    module = os.environ.get("ARTIFACT_TOOL_MODULE")
    if not module:
        root = (Path.home() / ".cache" / "codex-runtimes" /
                "codex-primary-runtime" / "dependencies" / "node")
        candidate = root / "node_modules" / "@oai" / "artifact-tool" / "dist" / "artifact_tool.mjs"
        if candidate.exists():
            module = str(candidate)
            bundled_node = root / "bin" / "node.exe"
            if bundled_node.exists():
                node = str(bundled_node)
    if not node or not module or not Path(module).exists():
        raise RuntimeError(
            "Excel export requires the local spreadsheet runtime. "
            "Set XLSX_NODE and ARTIFACT_TOOL_MODULE, then restart Streamlit.")
    return node, module


def generate_audit_workbook(payload):
    """Return an editable XLSX with formulas as bytes."""
    node, module = _runtime_paths()
    builder = Path(__file__).with_name("excel_export_builder.mjs")
    with tempfile.TemporaryDirectory(prefix="finance-audit-") as temp_dir:
        temp_path = Path(temp_dir)
        input_path = temp_path / "audit.json"
        output_path = temp_path / "finance-action-queue.xlsx"
        preview_path = temp_path / "preview.png"
        input_path.write_text(json.dumps(payload, allow_nan=False), encoding="utf-8")
        env = os.environ.copy()
        env["ARTIFACT_TOOL_MODULE"] = module
        result = subprocess.run(
            [node, str(builder), str(input_path), str(output_path), str(preview_path)],
            capture_output=True, text=True, timeout=90, env=env,
        )
        if result.returncode and not output_path.exists():
            detail = (result.stderr or result.stdout).strip().splitlines()[-8:]
            raise RuntimeError("Excel audit export failed" +
                               (f": {' | '.join(detail)}" if detail else "."))
        if not output_path.exists():
            raise RuntimeError("Excel audit export did not produce an output file.")
        return output_path.read_bytes()
