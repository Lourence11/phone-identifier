import io
import re
from typing import Optional, List, Dict, Any
import streamlit as st
import polars as pl
from pydantic import BaseModel, Field, ValidationError
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Page Config
st.set_page_config(page_title="PH Phone Validator & Cleaner", layout="wide")

# =====================================================================
# 1. PYDANTIC GUARDS (File & Row Validation)
# =====================================================================
class RawRowRecord(BaseModel):
    row_num: int
    raw_phone: str = Field(..., min_length=1)

    @classmethod
    def clean_cell(cls, val: Any) -> str:
        if val is None:
            return ""
        # Handle floats converted from numbers (e.g. 9171234567.0)
        s = str(val).strip()
        if s.endswith(".0"):
            s = s[:-2]
        return s


# =====================================================================
# 2. DETERMINISTIC PH NUMBER VALIDATION ENGINE
# =====================================================================
VALID_PROVINCIAL_AREAS = {
    "32", "33", "34", "35", "36", "38",
    "42", "43", "44", "45", "46", "47", "48", "49",
    "52", "53", "54", "55", "56",
    "62", "63", "64", "65",
    "72", "74", "75", "77", "78",
    "82", "83", "84", "85", "86", "87", "88"
}

def analyze_ph_phone(raw_val: str) -> Dict[str, Any]:
    if not raw_val or raw_val.strip() == "":
        return {
            "is_valid": False,
            "line_type": "Blank / Empty",
            "e164_format": "",
            "local_format": "",
            "audit_remark": "Missing Number"
        }

    # Clean non-digit characters except leading plus
    cleaned = re.sub(r'[\s\-\(\)\.]', '', str(raw_val).strip())

    # Strip national prefixes (+63, 63, 0)
    if cleaned.startswith('+63'):
        cleaned = cleaned[3:]
    elif cleaned.startswith('63'):
        cleaned = cleaned[2:]
    elif cleaned.startswith('0'):
        cleaned = cleaned[1:]

    # A. Check Mobile (10 digits starting with 9 or 8)
    if len(cleaned) == 10 and re.match(r'^(9\d{9}|8[1-9]\d{8})$', cleaned):
        return {
            "is_valid": True,
            "line_type": "Mobile",
            "e164_format": f"+63{cleaned}",
            "local_format": f"0{cleaned}",
            "audit_remark": "Valid Mobile"
        }

    # B. Metro Manila / GMA Landline (Area 02 + 8 local digits with NTC PTE prefix)
    # Valid PTEs: 3 (Bayantel), 5 (Eastern), 6 (ABS-CBN), 7 (Globe), 8 (PLDT)
    if len(cleaned) == 9 and cleaned.startswith('2'):
        pte = cleaned[1]
        if pte in {'3', '5', '6', '7', '8'}:
            return {
                "is_valid": True,
                "line_type": "Landline (GMA 02)",
                "e164_format": f"+63{cleaned}",
                "local_format": f"02-{cleaned[1:]}",
                "audit_remark": "Valid GMA 8-Digit"
            }
        return {
            "is_valid": False,
            "line_type": "Invalid Landline",
            "e164_format": cleaned,
            "local_format": cleaned,
            "audit_remark": f"Area 02 invalid operator prefix ({pte})"
        }

    # C. Provincial Landlines (2-digit area code + 7 local digits)
    if len(cleaned) == 9 and cleaned[:2] in VALID_PROVINCIAL_AREAS:
        area = cleaned[:2]
        return {
            "is_valid": True,
            "line_type": "Landline (Provincial)",
            "e164_format": f"+63{cleaned}",
            "local_format": f"0{area}-{cleaned[2:]}",
            "audit_remark": "Valid Provincial Landline"
        }

    # Flag legacy 7-digit Metro Manila numbers
    if len(cleaned) == 8 and cleaned.startswith('2'):
        return {
            "is_valid": False,
            "line_type": "Outdated Format",
            "e164_format": cleaned,
            "local_format": cleaned,
            "audit_remark": "Old 7-digit Manila landline (needs 8-digit migration)"
        }

    return {
        "is_valid": False,
        "line_type": "Invalid / Unknown",
        "e164_format": cleaned,
        "local_format": cleaned,
        "audit_remark": "Invalid length or prefix"
    }


# =====================================================================
# 3. OPENPYXL WORKBOOK BUILDER (Dynamic Formula Injection)
# =====================================================================
def build_excel_export(df: pl.DataFrame, phone_col: str) -> io.BytesIO:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Validated Numbers"

    orig_cols = [c for c in df.columns if not c.startswith("_meta_")]
    new_cols = [
        "Identified Type",
        "Normalized (E.164)",
        "Local Format",
        "Validation Status",
        "Dial Status (Formula)"
    ]
    all_headers = orig_cols + new_cols
    ws.append(all_headers)

    # Styling definitions
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    data_font = Font(name="Segoe UI", size=10)
    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9")
    )

    type_col_idx = len(orig_cols) + 1
    type_col_letter = get_column_letter(type_col_idx)

    # Populate rows
    for row_num, row in enumerate(df.iter_rows(named=True), start=2):
        row_data = [row[col] for col in orig_cols]

        line_type = row["_meta_type"]
        e164 = row["_meta_e164"]
        local_fmt = row["_meta_local"]
        audit_note = row["_meta_remark"]

        # Native Dynamic Excel Formula injected into each row
        # Reads the type cell; outputs DIALABLE vs REVIEW
        dynamic_formula = f'=IF(OR({type_col_letter}{row_num}="Mobile", ISNUMBER(SEARCH("Landline", {type_col_letter}{row_num}))), "DIALABLE", "REVIEW / INVALID")'

        full_row = row_data + [line_type, e164, local_fmt, audit_note, dynamic_formula]
        ws.append(full_row)

    # Apply styling & borders
    for col_idx in range(1, len(all_headers) + 1):
        h_cell = ws.cell(row=1, column=col_idx)
        h_cell.fill = header_fill
        h_cell.font = header_font
        h_cell.alignment = Alignment(horizontal="center", vertical="center")

    for r in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=len(all_headers)):
        for cell in r:
            cell.font = data_font
            cell.border = thin_border

    # Dynamic column widths
    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 14)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


# =====================================================================
# 4. STREAMLIT INTERFACE
# =====================================================================
st.title("Philippine Phone & Landline Identifier")
st.caption("Upload your spreadsheet, pick the target number column, and export a scrubbed file with dynamic Excel validation formulas.")

uploaded_file = st.file_uploader("Upload Excel or CSV file", type=["xlsx", "xls", "csv"])

if uploaded_file:
    # 1. Read into Polars
    try:
        if uploaded_file.name.endswith(".csv"):
            df = pl.read_csv(uploaded_file.getvalue(), infer_schema_length=10000)
        else:
            df = pl.read_excel(uploaded_file.getvalue())
    except Exception as e:
        st.error(f"Error loading file: {e}")
        st.stop()

    st.success(f"File loaded successfully: **{df.shape[0]} rows**, **{df.shape[1]} columns**")

    # 2. Select the target phone column
    all_columns = df.columns
    # Auto-detect sensible default
    default_idx = 0
    for idx, col in enumerate(all_columns):
        if any(term in col.lower() for term in ["phone", "tel", "landline", "mobile", "contact"]):
            default_idx = idx
            break

    selected_col = st.selectbox(
        "Select the column containing Phone / Landline / Mobile numbers:",
        options=all_columns,
        index=default_idx
    )

    if st.button("Identify and Clean Numbers", type="primary"):
        with st.spinner("Validating with Pydantic & Polars..."):
            # 3. Pydantic Guard Check
            raw_values = df[selected_col].to_list()
            guard_errors = []
            cleaned_strings = []

            for i, val in enumerate(raw_values, start=2):
                cleaned_str = RawRowRecord.clean_cell(val)
                try:
                    record = RawRowRecord(row_num=i, raw_phone=cleaned_str)
                    cleaned_strings.append(record.raw_phone)
                except ValidationError:
                    cleaned_strings.append("")
                    guard_errors.append(f"Row {i}: Missing or empty contact number.")

            if guard_errors and len(guard_errors) == len(raw_values):
                st.error("Every row in the selected column is empty or invalid. Please check your column selection.")
                st.stop()

            # 4. Polars Vectorized Mapping
            meta_results = [analyze_ph_phone(p) for p in cleaned_strings]

            meta_df = pl.DataFrame({
                "_meta_type": [m["line_type"] for m in meta_results],
                "_meta_e164": [m["e164_format"] for m in meta_results],
                "_meta_local": [m["local_format"] for m in meta_results],
                "_meta_remark": [m["audit_remark"] for m in meta_results],
            })

            processed_df = df.hstack(meta_df)

            # 5. Display Breakdown & Metrics
            st.divider()
            types_counts = meta_df["_meta_type"].value_counts().to_dicts()

            col1, col2, col3, col4 = st.columns(4)
            mobiles = sum(c["count"] for c in types_counts if c["_meta_type"] == "Mobile")
            landlines = sum(c["count"] for c in types_counts if "Landline" in c["_meta_type"])
            invalids = sum(c["count"] for c in types_counts if "Invalid" in c["_meta_type"] or "Unknown" in c["_meta_type"])
            blanks = sum(c["count"] for c in types_counts if c["_meta_type"] == "Blank / Empty")

            col1.metric("Mobile Numbers", mobiles)
            col2.metric("Landlines", landlines)
            col3.metric("Invalid / Outdated", invalids)
            col4.metric("Blank / Empty", blanks)

            # 6. Preview Result Table
            st.subheader("Data Preview")
            preview_cols = [selected_col, "_meta_type", "_meta_e164", "_meta_local", "_meta_remark"]
            st.dataframe(processed_df.select(preview_cols).head(10), use_container_width=True)

            # 7. Generate openpyxl Workbook with Dynamic Formulas
            excel_buffer = build_excel_export(processed_df, selected_col)

            st.download_button(
                label="📥 Download Cleaned Excel File (With Live Formulas)",
                data=excel_buffer,
                file_name=f"Cleaned_{uploaded_file.name.rsplit('.', 1)[0]}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
