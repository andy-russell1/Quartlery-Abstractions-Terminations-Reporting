from __future__ import annotations

from datetime import date, datetime
from io import BytesIO
from importlib.util import find_spec
import xml.etree.ElementTree as ET
import re
import zipfile
from typing import Any
from pathlib import Path

import pandas as pd
from streamlit.runtime.uploaded_file_manager import UploadedFile


def parse_date_from_filename(filename: str | None) -> date | None:
    if not filename:
        return None

    patterns = [r"(20\d{2}[01]\d[0-3]\d)", r"(20\d{2}-[01]\d-[0-3]\d)"]
    for pattern in patterns:
        m = re.search(pattern, filename)
        if not m:
            continue
        raw = m.group(1)
        try:
            if "-" in raw:
                return datetime.strptime(raw, "%Y-%m-%d").date()
            return datetime.strptime(raw, "%Y%m%d").date()
        except ValueError:
            continue
    return None


def load_excel_with_fallback(upload: UploadedFile, preferred_sheet: str | None = None) -> tuple[pd.DataFrame, list[str], str]:
    warnings: list[str] = []
    payload = upload.getvalue()
    payload_stripped = payload.lstrip()

    # Some upstream exports are Excel 2003 XML content with a .xls extension.
    if payload_stripped.startswith(b"<?xml") or payload_stripped.startswith(b"\xef\xbb\xbf<?xml"):
        return _load_excel_2003_xml(upload.name, payload, preferred_sheet)

    suffix = Path(upload.name or "").suffix.lower()
    if suffix in {".xlsx", ".xlsm"}:
        candidate_engines = ["openpyxl"]
    elif suffix == ".xls":
        candidate_engines = ["xlrd", "openpyxl"]
    else:
        # Unknown/missing extension: try modern engine first, then legacy.
        candidate_engines = ["openpyxl", "xlrd"]

    excel = None
    errors_by_engine: list[tuple[str, str]] = []
    for engine in candidate_engines:
        if engine == "xlrd" and find_spec("xlrd") is None:
            errors_by_engine.append((engine, "xlrd is not installed"))
            continue
        try:
            excel = pd.ExcelFile(BytesIO(payload), engine=engine)
            break
        except Exception as exc:  # pragma: no cover - defensive fallback
            errors_by_engine.append((engine, str(exc)))
            continue

    if excel is None:
        details = "; ".join(f"{eng}: {err}" for eng, err in errors_by_engine)
        msg = f"{upload.name}: unable to open as Excel. Ensure file is .xlsx/.xlsm (or .xls with xlrd installed)."
        if details:
            msg = f"{msg} Details: {details}"
        raise ValueError(msg)

    if preferred_sheet and preferred_sheet in excel.sheet_names:
        chosen = preferred_sheet
    else:
        chosen = excel.sheet_names[0]
        if preferred_sheet:
            warnings.append(
                f"{upload.name}: expected sheet '{preferred_sheet}' not found; used '{chosen}' instead."
            )

    df = pd.read_excel(BytesIO(payload), sheet_name=chosen, engine=excel.engine)
    return df, warnings, chosen


def _local_name(tag: str) -> str:
    if "}" in tag:
        return tag.split("}", 1)[1]
    return tag


def _load_excel_2003_xml(
    filename: str | None,
    payload: bytes,
    preferred_sheet: str | None = None,
) -> tuple[pd.DataFrame, list[str], str]:
    warnings: list[str] = []

    root = ET.fromstring(payload)

    worksheets: list[ET.Element] = []
    for elem in root.iter():
        if _local_name(elem.tag) == "Worksheet":
            worksheets.append(elem)

    if not worksheets:
        raise ValueError(f"{filename}: XML file does not contain SpreadsheetML worksheets.")

    def worksheet_name(ws: ET.Element) -> str:
        for k, v in ws.attrib.items():
            if k.endswith("}Name") or k == "Name":
                return str(v)
        return "Sheet1"

    sheet_names = [worksheet_name(ws) for ws in worksheets]
    chosen_idx = 0
    if preferred_sheet and preferred_sheet in sheet_names:
        chosen_idx = sheet_names.index(preferred_sheet)
    elif preferred_sheet:
        warnings.append(
            f"{filename}: expected sheet '{preferred_sheet}' not found; used '{sheet_names[0]}' instead."
        )

    ws = worksheets[chosen_idx]

    table = None
    for child in ws.iter():
        if _local_name(child.tag) == "Table":
            table = child
            break
    if table is None:
        return pd.DataFrame(), warnings, sheet_names[chosen_idx]

    rows_out: list[list[Any]] = []
    max_cols = 0

    for row_elem in table:
        if _local_name(row_elem.tag) != "Row":
            continue

        row_vals: list[Any] = []
        col_pos = 1

        for cell_elem in row_elem:
            if _local_name(cell_elem.tag) != "Cell":
                continue

            idx_attr = None
            for k, v in cell_elem.attrib.items():
                if k.endswith("}Index") or k == "Index":
                    idx_attr = v
                    break
            if idx_attr is not None:
                try:
                    target = int(idx_attr)
                    while col_pos < target:
                        row_vals.append(None)
                        col_pos += 1
                except ValueError:
                    pass

            cell_val: Any = None
            for sub in cell_elem:
                if _local_name(sub.tag) == "Data":
                    cell_val = sub.text
                    break
            row_vals.append(cell_val)
            col_pos += 1

        if len(row_vals) > max_cols:
            max_cols = len(row_vals)
        rows_out.append(row_vals)

    for r in rows_out:
        if len(r) < max_cols:
            r.extend([None] * (max_cols - len(r)))

    if not rows_out:
        return pd.DataFrame(), warnings, sheet_names[chosen_idx]

    header = [("" if h is None else str(h)) for h in rows_out[0]]
    data_rows = rows_out[1:] if len(rows_out) > 1 else []
    df = pd.DataFrame(data_rows, columns=header)
    return df, warnings, sheet_names[chosen_idx]


def dataframe_to_excel_bytes(sheets: dict[str, pd.DataFrame]) -> bytes:
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        for sheet_name, df in sheets.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    out.seek(0)
    return out.getvalue()


def build_summary_workbook_bytes(
    counts: dict[str, int],
    terminations: pd.DataFrame,
    abstractions: pd.DataFrame,
) -> bytes:
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        summary_rows = [
            ("Total contracts", counts.get("total_contracts", 0)),
            ("Terminated count", counts.get("term_count", 0)),
            ("Abstracted count", counts.get("abs_count", 0)),
        ]
        summary_df = pd.DataFrame(summary_rows, columns=["Metric", "Value"])
        summary_df.to_excel(writer, sheet_name="Summary", index=False, startrow=0)
        ws = writer.sheets["Summary"]

        term_label_row0 = len(summary_df) + 2
        ws.cell(row=term_label_row0 + 1, column=1, value="Terminations")
        terminations.to_excel(
            writer,
            sheet_name="Summary",
            index=False,
            startrow=term_label_row0 + 1,
        )

        abs_label_row0 = term_label_row0 + len(terminations) + 4
        ws.cell(row=abs_label_row0 + 1, column=1, value="Abstractions")
        abstractions.to_excel(
            writer,
            sheet_name="Summary",
            index=False,
            startrow=abs_label_row0 + 1,
        )

    out.seek(0)
    return out.getvalue()


def build_full_zip(file_bytes: dict[str, bytes]) -> bytes:
    out = BytesIO()
    with zipfile.ZipFile(out, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        for filename, payload in file_bytes.items():
            zf.writestr(filename, payload)
    out.seek(0)
    return out.getvalue()


def ensure_date(value: Any) -> date | None:
    if value is None:
        return None
    if isinstance(value, date):
        return value
    if isinstance(value, datetime):
        return value.date()
    return None
