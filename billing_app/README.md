# Billing Workbook Replacement (Streamlit)

This app replaces an Excel/VBA billing workbook with a Streamlit workflow.

## Run

```bash
cd billing_app
pip install -r requirements.txt
streamlit run app.py
```

## What it does

1. Accepts 4 uploads:
   - Payable (`HLSELIST`)
   - Receivable (`LEASELIST`)
   - Freehold (`PROPLIST`)
   - Prev Period Master
2. Takes manual billing parameters.
3. Builds:
   - New Period Master (`YYYYMMDD Master.xlsx`)
   - `Summary.xlsx`
   - `Bill.xlsx` from `templates/bill_template.xlsx` (single sheet output named `LA ongoing fees`)
4. Shows results tabs:
   - Executive Summary
   - Abstracted table
   - Terminated table
   - Exceptions / validation warnings
5. Provides individual downloads and a full in-memory zip pack.

## Security and data persistence

- Uploads are read directly from Streamlit `UploadedFile` bytes.
- Processing is performed in memory with `BytesIO`.
- No upload data or intermediate DataFrames are written to disk.
- `st.cache_data` is not used.
- A **Clear session** button clears `st.session_state` and reruns the app.

## Template mapping notes

`bill_writer.py` contains a `PARAM_CELL_MAP` dictionary with editable cell references for parameter placement.
Some addresses are placeholders and are marked with `TODO` comments.

## Limitations

- This implementation is position-based like the legacy macro and assumes expected columns by index.
- If expected sheet names are missing, the app falls back to the first sheet and logs warnings.
- The Bill layout depends on the provided template having a `Bill` sheet and section block at `B16:H17`.
- Legacy `.xls` XML Spreadsheet exports are supported in-memory (no conversion step required).
