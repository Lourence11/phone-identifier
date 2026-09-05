import io
import re
from typing import Optional, List, Dict, Any
import streamlit as st
import polars as pl
from pydantic import BaseModel, Field, ValidationError
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="PH Multi-Channel Phone Identifier", layout="wide")

# =====================================================================
# 1. PYDANTIC GUARDS (Input Cell Scrubbing)
# =====================================================================
class NumberCell(BaseModel):
    raw_val: str

    @classmethod
    def sanitize(cls, val: Any) -> str:
        if val is None:
            return ""
        s = str(val).strip()
        if s.endswith(".0"):  # Fix Excel auto-float conversions
            s = s[:-2]
        return s


# =====================================================================
# 2. VALIDATION ENGINES (Mobile, Metro 02, Provincial)
# =====================================================================
VALID_PROVINCIAL_AREAS = {
    "32", "33", "34", "35", "36", "38",
    "42", "43", "44", "45", "46", "47", "48", "49",
    "52", "53", "54", "55", "56",
    "62", "63", "64", "65",
    "72", "74", "75", "77", "78",
    "82", "83", "84", "85", "86", "87", "88"
}

def clean_to_digits(raw_val: str) -> str:
    cleaned = re.sub(r'[\s\-\(\)\.]', '', str(raw_val).strip())
    if cleaned.startswith('+63'):
        cleaned = cleaned[3:]
    elif cleaned.startswith('63'):
        cleaned = cleaned[2:]
    elif cleaned.startswith('0'):
        cleaned = cleaned[1:]
    return cleaned

def validate_number(raw_val: str, expected_type: str) -> Dict[str, Any]:
    cleaned = clean_to_digits(raw_val)
    if not cleaned:
        return {"status": "BLANK", "type": "Empty", "e164": "", "local": ""}

    # 1. Mobile Check (10 digits starting with 9 or 8)
    is_mobile = len(cleaned) == 10 and re.match(r'^(9\d{9}|8[1-9]\d{8})$', cleaned)

    # 2. GMA / Area 02 Check (9 digits starting with 2 + valid PTE digit)
    is_gma = (
        len(cleaned) == 9 
        and cleaned.startswith('2') 
        and cleaned[1] in {'3', '5', '6', '7', '8'}
    )

    # 3. Provincial Check (9 digits starting with registered area code)
    is_provincial = len(cleaned) == 9 and cleaned[:2] in VALID_PROVINCIAL_AREAS

    # Check for legacy 7-digit Metro Manila numbers
    is_outdated_gma = len(cleaned) == 8 and cleaned.startswith('2')

    if is_mobile:
        return {
            "status": "VALID",
            "type": "Mobile",
            "e164": f"+63{cleaned}",
            "local": f"0{cleaned}"
        }
    elif is_gma:
        return {
            "status": "VALID",
            "type": "Landline (GMA 02)",
            "e164": f"+63{cleaned}",
            "local": f"02-{cleaned[1:5]}-{cleaned[5:]}"
        }
    elif is_provincial:
        area = cleaned[:2]
        return {
            "status": "VALID",
            "type": f"Landline (Area 0{area})",
            "e164": f"+63{cleaned}",
            "local": f"0{area}-{cleaned[2:5]}-{cleaned[5:]}"
        }
    elif is_outdated_gma:
        return {
            "status": "INVALID",
            "type": "Old 7-digit Manila",
            "e164": cleaned,
            "local": cleaned
        }
    else:
        return {
            "status": "INVALID",
            "type": "Invalid Format",
            "e164": cleaned,
            "local": cleaned
        }


# =====================================================================
# 3. EXCEL EXPORTER (Dynamic Native Formulas)
# =====================================================================
def export_multi_column_excel(df: pl.DataFrame, col_map: Dict[str, str]) -> io.BytesIO:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Validated Numbers"

    base_cols = [c for c in df.columns if not c.startswith("_meta_")]
    headers = list(base_cols)

    # Build target output columns
    for label in ["Mobile", "Telephone", "Landline"]:
        if col_map.get(label):
            headers.extend([
                f"{label} Status",
                f"{label} E.164",
                f"{label} Local"
            ])
    
    headers.append("Master Reachability (Formula)")
    ws.append(headers)

    # Styles
    navy_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
    data_font = Font(name="Segoe UI", size=9)
    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9")
    )

    # Row iteration
    for row_idx, row in enumerate(df.iter_rows(named=True), start=2):
        row_vals = [row[c] for c in base_cols]
        status_col_letters = []

        curr_col_idx = len(base_cols) + 1
        for label in ["Mobile", "Telephone", "Landline"]:
            if col_map.get(label):
                status_col_letters.append(get_column_letter(curr_col_idx))
                row_vals.extend([
                    row[f"_meta_{label}_status"],
                    row[f"_meta_{label}_e164"],
                    row[f"_meta_{label}_local"]
                ])
                curr_col_idx += 3

        # Dynamic formula: checks if any of the three columns have a "VALID" status
        or_conditions = ",".join([f'{col}{row_idx}="VALID"' for col in status_col_letters])
        master_formula = f'=IF(OR({or_conditions}), "REACHABLE", "NO VALID NUMBERS")'
        row_vals.append(master_formula)

        ws.append(row_vals)

    # Format cells
    for col_idx in range(1, len(headers) + 1):
        c = ws.cell(row=1, column=col_idx)
        c.fill = navy_fill
        c.font = header_font
        c.alignment = Alignment(horizontal="center", vertical="center")

    for r in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=len(headers)):
        for c in r:
            c.font = data_font
            c.border = thin_border

    # Auto-fit column widths
    for col in ws.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 14)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


# =====================================================================
# 4. STREAMLIT UI
# =====================================================================
st.title("Philippine Phone, Telephone & Landline Identifier")
st.caption("Validates Mobile (09XX/08XX), Metro Manila Area 02 (8-digit NTC), and Provincial Area Codes simultaneously.")

uploaded_file = st.file_uploader("Upload Excel or CSV file", type=["xlsx", "xls", "csv"])

if uploaded_file:
    # 1. Ingest file into Polars
    if uploaded_file.name.endswith(".csv"):
        df = pl.read_csv(uploaded_file.getvalue(), infer_schema_length=10000)
    else:
        df = pl.read_excel(uploaded_file.getvalue())

    st.success(f"File loaded: **{df.shape[0]} rows**, **{df.shape[1]} columns**")

    # 2. Match column headers
    cols = ["(None / Skip)"] + df.columns

    def auto_match(pattern: str) -> int:
        for i, c in enumerate(df.columns):
            if re.search(pattern, c, re.IGNORECASE):
                return i + 1
        return 0

    st.subheader("Select Columns to Validate")
    col1, col2, col3 = st.columns(3)

    with col1:
        sel_mobile = st.selectbox(
            "📱 Mobile Phone Column",
            options=cols,
            index=auto_match(r"mobile|cell|cel")
        )
    with col2:
        sel_tele = st.selectbox(
            "☎️ Telephone / Area 02 Column",
            options=cols,
            index=auto_match(r"telephone|tele|phone")
        )
    with col3:
        sel_landline = st.selectbox(
            "🏢 Provincial Landline Column",
            options=cols,
            index=auto_match(r"landline|provincial")
        )

    col_map = {
        "Mobile": sel_mobile if sel_mobile != "(None / Skip)" else None,
        "Telephone": sel_tele if sel_tele != "(None / Skip)" else None,
        "Landline": sel_landline if sel_landline != "(None / Skip)" else None,
    }

    if not any(col_map.values()):
        st.warning("Please assign at least one column to process.")
        st.stop()

    if st.button("Identify and Clean All Numbers", type="primary"):
        with st.spinner("Processing with Polars and Pydantic..."):
            meta_dict = {}

            for label, col_name in col_map.items():
                if not col_name:
                    continue

                raw_cells = [NumberCell.sanitize(v) for v in df[col_name].to_list()]
                parsed = [validate_number(val, label) for val in raw_cells]

                meta_dict[f"_meta_{label}_status"] = [p["status"] for p in parsed]
                meta_dict[f"_meta_{label}_type"] = [p["type"] for p in parsed]
                meta_dict[f"_meta_{label}_e164"] = [p["e164"] for p in parsed]
                meta_dict[f"_meta_{label}_local"] = [p["local"] for p in parsed]

            meta_df = pl.DataFrame(meta_dict)
            processed_df = df.hstack(meta_df)

            # Metrics
            st.divider()
            m_cols = st.columns(len([k for k, v in col_map.items() if v]))
            idx = 0
            for label, col_name in col_map.items():
                if col_name:
                    valid_count = sum(1 for s in meta_dict[f"_meta_{label}_status"] if s == "VALID")
                    total_count = len(df)
                    m_cols[idx].metric(f"{label} Valid", f"{valid_count} / {total_count}")
                    idx += 1

            # Preview
            st.subheader("Processed Data Preview")
            display_cols = [c for c in processed_df.columns if not c.startswith("_meta_")] + [
                f"_meta_{l}_local" for l in ["Mobile", "Telephone", "Landline"] if col_map.get(l)
            ]
            st.dataframe(processed_df.select(display_cols), use_container_width=True)

            # Export with Dynamic Formulas
            excel_bytes = export_multi_column_excel(processed_df, col_map)
            st.download_button(
                label="📥 Download Cleaned Excel (With Live =IF() Formulas)",
                data=excel_bytes,
                file_name=f"Cleaned_{uploaded_file.name.rsplit('.', 1)[0]}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
