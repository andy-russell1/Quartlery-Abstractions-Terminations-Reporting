from __future__ import annotations

from datetime import date
from pathlib import Path
import re

import pandas as pd
import streamlit as st

from billing_engine import build_diffs, build_new_master, build_prev_master, run_sanity_checks
from bill_writer import generate_bill_bytes
from io_utils import build_full_zip, build_summary_workbook_bytes, dataframe_to_excel_bytes, load_excel_with_fallback, parse_date_from_filename
from validators import build_exception_tables, summarize_exceptions, validate_required_positions


st.set_page_config(page_title="Billing Workbook Replacement", layout="wide")
st.title("Billing Workbook Replacement")
st.caption("In-memory processing only. No uploaded files or intermediate data are written to disk.")


def _is_numeric_id(series: pd.Series) -> pd.Series:
    return (
        series.astype(str)
        .str.strip()
        .str.replace(r"\.0+$", "", regex=True)
        .str.match(r"^\d+$", na=False)
    )


def _backfill_prev_ids_by_property(prev_master: pd.DataFrame, new_master: pd.DataFrame) -> tuple[pd.DataFrame, int]:
    prev = prev_master.copy()
    prev["system_id"] = prev["system_id"].astype(str).str.strip().str.replace(r"\.0+$", "", regex=True)
    prev["property_name"] = prev["property_name"].astype(str).str.strip()

    new = new_master.copy()
    new["system_id"] = new["system_id"].astype(str).str.strip().str.replace(r"\.0+$", "", regex=True)
    new["property_name"] = new["property_name"].astype(str).str.strip()

    valid_new = new.loc[_is_numeric_id(new["system_id"]) & new["property_name"].ne(""), ["property_name", "system_id"]]
    prop_counts = valid_new["property_name"].str.lower().value_counts()
    unique_props = set(prop_counts[prop_counts == 1].index)
    name_to_id = (
        valid_new.assign(_k=valid_new["property_name"].str.lower())
        .loc[lambda d: d["_k"].isin(unique_props)]
        .set_index("_k")["system_id"]
        .to_dict()
    )

    invalid_prev = ~_is_numeric_id(prev["system_id"])
    mapped = prev.loc[invalid_prev, "property_name"].str.lower().map(name_to_id)
    can_fill = mapped.notna()
    if can_fill.any():
        fill_idx = prev.loc[invalid_prev].index[can_fill]
        prev.loc[fill_idx, "system_id"] = mapped.loc[can_fill].values

    return prev, int(can_fill.sum())


if st.button("Clear session"):
    st.session_state.clear()
    st.rerun()

with st.form("parameters_form"):
    st.subheader("A) Parameters")
    c1, c2, c3 = st.columns(3)
    with c1:
        client_name = st.text_input("Client name")
        currency = st.text_input("Currency", value="GBP")
        exchange_rate = st.number_input("Exchange rate", value=1.0, format="%.6f")
        billing_period_start = st.date_input("Billing period start date", value=date.today())
        billing_period_end = st.date_input("Billing period end date", value=date.today())
    with c2:
        billing_frequency = st.text_input("Billing frequency", value="Quarterly in arrears")
        client_po_number = st.text_input("Client PO number")
        payment_terms = st.text_input("Payment terms", value="30 days")
        instruction_number = st.text_input("Instruction number")
        previous_bill_date = st.date_input("Previous bill date", value=date.today())
    with c3:
        next_bill_date = st.date_input("Next bill date", value=date.today())
        scope = st.text_input("Scope")
        person_to_be_billed = st.text_input("Person to be billed")
        unit_fee = st.number_input("Cost per abstraction/termination", min_value=0.0, value=0.0, format="%.2f")
        vat_applied = st.checkbox("VAT applied?", value=True)

    st.subheader("B) Upload files")
    pay_upload = st.file_uploader("Payable (HLSELIST) snapshot (Excel)", type=["xlsx", "xlsm", "xls"])
    rec_upload = st.file_uploader("Receivable (LEASELIST) snapshot (Excel)", type=["xlsx", "xlsm", "xls"])
    free_upload = st.file_uploader("Freehold (PROPLIST) snapshot (Excel)", type=["xlsx", "xlsm", "xls"])
    prev_upload = st.file_uploader("Prev Period Master (Excel)", type=["xlsx", "xlsm", "xls"])

    run_clicked = st.form_submit_button("C) Validate + Run")

if run_clicked:
    warnings: list[str] = []
    errors: list[str] = []

    if any(f is None for f in (pay_upload, rec_upload, free_upload, prev_upload)):
        errors.append("Please upload all 4 required files.")
    if billing_period_start >= billing_period_end:
        errors.append(
            f"Billing period start date must be before end date "
            f"({billing_period_start.strftime('%Y-%m-%d')} < {billing_period_end.strftime('%Y-%m-%d')})."
        )

    if errors:
        for err in errors:
            st.error(err)
        st.stop()

    params = {
        "client_name": client_name,
        "currency": currency,
        "exchange_rate": float(exchange_rate),
        "billing_period_start": billing_period_start,
        "billing_period_end": billing_period_end,
        "billing_frequency": billing_frequency,
        "client_po_number": client_po_number,
        "payment_terms": payment_terms,
        "unit_fee": float(unit_fee),
        "instruction_number": instruction_number,
        "previous_bill_date": previous_bill_date,
        "next_bill_date": next_bill_date,
        "scope": scope,
        "vat_applied": "Yes" if vat_applied else "No",
        "person_to_be_billed": person_to_be_billed,
    }

    try:
        pay_df, w, _ = load_excel_with_fallback(pay_upload, preferred_sheet="HLSELIST")
        warnings.extend(w)
        rec_df, w, _ = load_excel_with_fallback(rec_upload, preferred_sheet="LEASELIST")
        warnings.extend(w)
        free_df, w, _ = load_excel_with_fallback(free_upload, preferred_sheet="PROPLIST")
        warnings.extend(w)
        prev_df, w, _ = load_excel_with_fallback(prev_upload, preferred_sheet=None)
        warnings.extend(w)
    except ValueError as exc:
        st.error(str(exc))
        st.stop()

    warnings.extend(validate_required_positions(pay_df, [1, 8], "Payable"))
    warnings.extend(validate_required_positions(rec_df, [1, 2], "Receivable"))
    warnings.extend(validate_required_positions(free_df, [1, 2, 5], "Freehold"))
    warnings.extend(validate_required_positions(prev_df, [2], "Prev Master"))
    warnings.extend(run_sanity_checks(pay_df, rec_df, free_df, prev_df))

    new_master = build_new_master(pay_df, rec_df, free_df)
    prev_master = build_prev_master(prev_df)

    prev_master, mapped_count = _backfill_prev_ids_by_property(prev_master, new_master)
    if mapped_count:
        warnings.append(
            f"Backfilled {mapped_count} previous-period rows where system_id was non-numeric using property-name matches."
        )

    invalid_prev_remaining = int((~_is_numeric_id(prev_master["system_id"]) & prev_master["system_id"].astype(str).str.strip().ne("")).sum())
    if invalid_prev_remaining:
        warnings.append(
            f"Previous master still has {invalid_prev_remaining} non-numeric system_id values; these may inflate terminated counts."
        )

    result = build_diffs(new_master, prev_master)
    if "total_contracts" not in result.counts:
        result.counts["total_contracts"] = int(
            result.new_master.loc[
                result.new_master["system_id"].astype(str).str.strip().ne(""),
                "system_id",
            ]
            .astype(str)
            .str.replace(r"\.0+$", "", regex=True)
            .nunique()
        )

    exception_tables = build_exception_tables(result.new_master, result.prev_master)
    warnings.extend(summarize_exceptions(exception_tables))

    today_str = date.today().strftime("%Y%m%d")
    new_master_filename = f"{today_str} Master.xlsx"
    summary_filename = "Summary.xlsx"
    bill_filename = "Bill.xlsx"
    safe_client = re.sub(r'[\\/:*?"<>|]+', "", (client_name or "").strip())
    safe_client = re.sub(r"\s+", " ", safe_client).strip() or "Client"
    full_pack_filename = f"{today_str} {safe_client} billing_doc.zip"

    new_master_bytes = dataframe_to_excel_bytes(
        {
            "New Master": result.new_master_out,
            "Prev Master": result.prev_master_out,
        }
    )
    summary_bytes = build_summary_workbook_bytes(
        result.counts,
        result.terminations_2col,
        result.abstractions_2col,
    )

    template_path = Path(__file__).resolve().parent / "templates" / "bill_template.xlsx"
    if not template_path.exists():
        st.error(f"Missing template: {template_path}")
        st.stop()

    bill_bytes = generate_bill_bytes(
        template_path=template_path,
        params=params,
        abstractions_2col=result.abstractions_2col,
        terminations_2col=result.terminations_2col,
        contract_count=result.counts.get("total_contracts", 0),
    )

    zip_bytes = build_full_zip(
        {
            new_master_filename: new_master_bytes,
            summary_filename: summary_bytes,
            bill_filename: bill_bytes,
        }
    )

    st.session_state["run_outputs"] = {
        "warnings": warnings,
        "result": result,
        "exceptions": exception_tables,
        "new_master_bytes": new_master_bytes,
        "summary_bytes": summary_bytes,
        "bill_bytes": bill_bytes,
        "zip_bytes": zip_bytes,
        "new_master_filename": new_master_filename,
        "summary_filename": summary_filename,
        "bill_filename": bill_filename,
        "full_pack_filename": full_pack_filename,
    }

if "run_outputs" in st.session_state:
    out = st.session_state["run_outputs"]
    result = out["result"]
    exceptions = out["exceptions"]

    st.subheader("D) Results")

    parsed_prev = parse_date_from_filename(prev_upload.name if prev_upload else None)
    compare_start = parsed_prev or previous_bill_date
    st.info(f"Comparing period: {compare_start.strftime('%Y-%m-%d')} -> {date.today().strftime('%Y-%m-%d')}")

    tabs = st.tabs(["Executive Summary", "Abstracted table", "Terminated table", "Exceptions / validation warnings"])

    with tabs[0]:
        total_contracts = int(
            result.counts.get(
                "total_contracts",
                result.new_master.loc[
                    result.new_master["system_id"].astype(str).str.strip().ne(""),
                    "system_id",
                ]
                .astype(str)
                .str.replace(r"\.0+$", "", regex=True)
                .nunique(),
            )
        )
        abs_count = int(result.counts.get("abs_count", len(result.abstractions_2col)))
        term_count = int(result.counts.get("term_count", len(result.terminations_2col)))

        new_rows = int(result.new_master["system_id"].astype(str).str.strip().ne("").sum())
        prev_rows = int(result.prev_master["system_id"].astype(str).str.strip().ne("").sum())
        row_net = new_rows - prev_rows
        churn_net = abs_count - term_count

        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Total contracts", total_contracts)
        c2.metric("Abstracted", abs_count)
        c3.metric("Terminated", term_count)
        c4.metric("New Master Records", new_rows)
        c5.metric("Prev Master Records", prev_rows)

        st.caption(f"Net row change: {row_net} | Abstracted - Terminated: {churn_net}")

        st.write("Abstracted rows")
        st.dataframe(result.abstractions_2col, use_container_width=True)
        st.write("Terminated rows")
        st.dataframe(result.terminations_2col, use_container_width=True)

    with tabs[1]:
        abstracted_view = result.new_master_out.loc[
            result.new_master_out["abstracted_flag"] == "Abstracted",
            ["abstracted_flag", "system_id", "property_name"],
        ]
        st.dataframe(abstracted_view, use_container_width=True)

    with tabs[2]:
        terminated_view = result.prev_master_out.loc[
            result.prev_master_out["terminated_flag"] == "Terminated",
            ["terminated_flag", "system_id", "property_name"],
        ]
        st.dataframe(terminated_view, use_container_width=True)

    with tabs[3]:
        if out["warnings"]:
            for w in out["warnings"]:
                st.warning(w)
        else:
            st.success("No validation warnings.")

        st.write("Missing system_id in new master")
        st.dataframe(exceptions.missing_system_id_new, use_container_width=True)
        st.write("Missing system_id in prev master")
        st.dataframe(exceptions.missing_system_id_prev, use_container_width=True)
        st.write("Duplicate system_id in new master")
        st.dataframe(exceptions.duplicate_system_id_new, use_container_width=True)
        st.write("Duplicate system_id in prev master")
        st.dataframe(exceptions.duplicate_system_id_prev, use_container_width=True)
        st.write("system_id present in multiple sources")
        st.dataframe(exceptions.multi_source_duplicates, use_container_width=True)

    st.subheader("E) Downloads")
    d1, d2, d3, d4 = st.columns(4)
    with d1:
        st.download_button(
            "Download New Period Master",
            data=out["new_master_bytes"],
            file_name=out["new_master_filename"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    with d2:
        st.download_button(
            "Download Summary.xlsx",
            data=out["summary_bytes"],
            file_name=out["summary_filename"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    with d3:
        st.download_button(
            "Download Bill.xlsx",
            data=out["bill_bytes"],
            file_name=out["bill_filename"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    with d4:
        st.download_button(
            "Download Full Pack (.zip)",
            data=out["zip_bytes"],
            file_name=out.get("full_pack_filename", "billing_doc.zip"),
            mime="application/zip",
        )
else:
    st.info("Complete parameters and uploads, then click 'Validate + Run'.")

