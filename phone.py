import io
import re
from typing import Optional, List, Dict, Any
import streamlit as st
import polars as pl
from pydantic import BaseModel, Field, ValidationError
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="PH Phone & Landline Identifier", layout="wide")

# =====================================================================
# 1. PYDANTIC GUARDS (Input Data Cleansing)
# =====================================================================
class NumberCell(BaseModel):
    raw_val: str

    @classmethod
    def sanitize(cls, val: Any) -> str:
        if val is None:
            return ""
        s = str(val).strip()
        # Clean Excel auto-float conversion (e.g., 9171234567.0 -> 9171234567)
        if s.endswith(".0"):
            s = s[:-2]
        return s


# =====================================================================
# 2. TELECOM VALIDATION ENGINE (NTC Compliant + Anti-Dummy Rules)
# =====================================================================
VALID_PROVINCIAL_AREAS = {
    "32", "33", "34", "35", "36", "38",
    "42", "43", "44", "45", "46", "47", "48", "49",
    "52", "53", "54", "55", "56",
    "62", "63", "64", "65",
    "72", "74", "75", "77", "78",
    "82", "83", "84", "85", "86", "87", "88"
}

# Major mobile prefixes (Globe, Smart, Dito, TNT, TM)
VALID_MOBILE_PREFIXES_2DIGIT = {"90", "91", "92", "93", "94", "95", "96", "97", "98", "99", "89", "81"}

SEQUENTIAL_PATTERNS = {"1234567", "2345678", "3456789", "7654321", "9876543", "8765432"}

def clean_to_digits(raw_val: str) -> str:
    """Strips formatting, spaces, +63, 63, or trunk prefix 0."""
    cleaned = re.sub(r'[\s\-\(\)\.]', '', str(raw_val).strip())
    if cleaned.startswith('+63'):
        cleaned = cleaned[3:]
    elif cleaned.startswith('63'):
        cleaned = cleaned[2:]
    elif cleaned.startswith('0'):
        cleaned = cleaned[1:]
    return cleaned

def is_dummy_or_restricted(local_digits: str, min_repeat_len: int = 5) -> tuple[bool, str]:
    """
    Blocks repeated digits (1111111, 0000000), sequence patterns,
    and NTC reserved prefixes (local subscriber numbers cannot begin with 0 or 1).
    """
    # 1. Total repeated digits (e.g. 111-1111)
    if len(set(local_digits)) == 1:
        return True, "Repeated Dummy Number"

    # 2. Runs of repeating digits
    for char in set(local_digits):
        if char * min_repeat_len in local_digits:
            return True, "Suspicious Repetitive Digits"

    # 3. Sequential runs
    if local_digits[:7] in SEQUENTIAL_PATTERNS:
        return True, "Sequential Dummy Pattern"

    # 4. NTC Rule: Fixed line subscriber numbers cannot begin with 0 or 1 (reserved for emergency / hotlines)
    if local_digits.startswith(('0', '1')):
        return True, "Reserved Prefix (Starts with 0 or 1)"

    return False, ""

def validate_number(raw_val: str, expected_type: str) -> Dict[str, Any]:
    cleaned = clean_to_digits(raw_val)
    if not cleaned:
        return {
            "status": "BLANK",
            "type": "Empty",
            "e164": "",
            "local": "",
            "remark": "Empty Cell"
        }

    if not cleaned.isdigit():
        return {
            "status": "INVALID",
            "type": "Malformed",
            "e164": cleaned,
            "local": cleaned,
            "remark": "Contains non-numeric characters"
        }

    # -------------------------------------------------------------
    # 1. MOBILE (10 Digits)
    # -------------------------------------------------------------
    if len(cleaned) == 10 and cleaned[:2] in VALID_MOBILE_PREFIXES_2DIGIT:
        subscriber_part = cleaned[3:]
        is_dummy, reason = is_dummy_or_restricted(subscriber_part, min_repeat_len=6)
        if is_dummy:
            return {
                "status": "INVALID",
                "type": "Mobile (Dummy)",
                "e164": f"+63{cleaned}",
                "local": f"0{cleaned}",
                "remark": f"Fake/Dummy Mobile: {reason}"
            }
        return {
            "status": "VALID",
            "type": "Mobile",
            "e164": f"+63{cleaned}",
            "local": f"0{cleaned}",
            "remark": "Valid Philippine Mobile"
        }

    # -------------------------------------------------------------
    # 2. METRO MANILA / GMA LANDLINE (Area Code 2 + 8-Digit Local)
    # -------------------------------------------------------------
    if len(cleaned) == 9 and cleaned.startswith('2'):
        pte_digit = cleaned[1]   # 3=Bayan, 5=Eastern, 6=ABS-CBN, 7=Globe, 8=PLDT
        subscriber_local = cleaned[2:]

        if pte_digit not in {'3', '5', '6', '7', '8'}:
            return {
                "status": "INVALID",
                "type": "Landline (Area 02)",
                "e164": cleaned,
                "local": cleaned,
                "remark": f"Invalid Area 02 Operator Prefix ({pte_digit})"
            }

        is_dummy, reason = is_dummy_or_restricted(subscriber_local)
        if is_dummy:
            return {
                "status": "INVALID",
                "type": "Landline (Area 02)",
                "e164": f"+63{cleaned}",
                "local": f"02-{cleaned[1:5]}-{cleaned[5:]}",
                "remark": f"Invalid Subscriber: {reason}"
            }

        return {
            "status": "VALID",
            "type": "Landline (GMA 02)",
            "e164": f"+63{cleaned}",
            "local": f"02-{cleaned[1:5]}-{cleaned[5:]}",
            "remark": "Valid Metro Manila / GMA 8-Digit Landline"
        }

    # -------------------------------------------------------------
    # 3. PROVINCIAL LANDLINE (2-Digit Area Code + 7-Digit Local)
    # -------------------------------------------------------------
    if len(cleaned) == 9 and cleaned[:2] in VALID_PROVINCIAL_AREAS:
        area = cleaned[:2]
        subscriber_local = cleaned[2:]

        is_dummy, reason = is_dummy_or_restricted(subscriber_local)
        if is_dummy:
            return {
                "status": "INVALID",
                "type": f"Landline (Area 0{area})",
                "e164": f"+63{cleaned}",
                "local": f"0{area}-{subscriber_local[:3]}-{subscriber_local[3:]}",
                "remark": f"Invalid Provincial Subscriber: {reason}"
            }

        return {
            "status": "VALID",
            "type": f"Landline (Area 0{area})",
            "e164": f"+63{cleaned}",
            "local": f"0{area}-{subscriber_local[:3]}-{subscriber_local[3:]}",
            "remark": "Valid Provincial Landline"
        }

    # -------------------------------------------------------------
    # 4. EXPLICIT ERROR CATCHERS
    # -------------------------------------------------------------
    if len(cleaned) == 8 and cleaned.startswith('2'):
        return {
            "status": "INVALID",
            "type": "Legacy Landline",
            "e164": cleaned,
            "local": cleaned,
            "remark": "Unmigrated 7-digit Manila landline (needs 8-digit operator prefix)"
        }

    return {
        "status": "INVALID",
        "type": "Unrecognized",
        "e164": cleaned,
        "local": cleaned,
        "remark": f"Invalid length ({len(cleaned)} digits) or unregistered prefix"
    }


# =====================================================================
# 3. EXCEL EXPORTER (Side-by-Side Context Layout + Live Formulas)
# =====================================================================
def export_multi_column_excel(df: pl.DataFrame, col_map: Dict[str, str]) -> io.BytesIO:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Scrubbed Numbers"

    # Separate non-phone columns (e.g., ID, Account Name, Address)
    phone_cols_set = {v for v in col_map.values() if v}
    non_phone_cols = [c for c in df.columns if not c.startswith("_meta_") and c not in phone_cols_set]

    # Structure headers: Master Status -> Non-Phone Data -> Grouped Channels
    headers = ["Master Reachability (Formula)"] + non_phone_cols
    col_style_map = {}

    theme_colors = {
        "Master": "1F4E78",     # Dark Navy
        "Base": "2F5597",       # Soft Blue
        "Mobile": "1B365D",     # Deep Navy
        "Telephone": "3B444B",  # Slate Gray
        "Landline": "164E63"    # Deep Teal
    }

    status_col_letters = []
    current_idx = len(headers) + 1

    # Place each original column immediately followed by its validation columns
    for label in ["Mobile", "Telephone", "Landline"]:
        orig_col = col_map.get(label)
        if orig_col:
            # 1. Original Input Column
            headers.append(orig_col)
            col_style_map[current_idx] = theme_colors[label]
            current_idx += 1

            # 2. Status Column
            headers.append(f"{label} Status")
            col_style_map[current_idx] = theme_colors[label]
            status_col_letters.append(get_column_letter(current_idx))
            current_idx += 1

            # 3. Formatted Local & E.164 + Remark
            headers.extend([f"{label} Local", f"{label} E.164", f"{label} Remark"])
            col_style_map[current_idx] = theme_colors[label]
            col_style_map[current_idx + 1] = theme_colors[label]
            col_style_map[current_idx + 2] = theme_colors[label]
            current_idx += 3

    ws.append(headers)

    data_font = Font(name="Segoe UI", size=9)
    thin_border = Border(
        left=Side(style="thin", color="E0E0E0"),
        right=Side(style="thin", color="E0E0E0"),
        top=Side(style="thin", color="E0E0E0"),
        bottom=Side(style="thin", color="E0E0E0")
    )

    # Populate rows
    for row_idx, row in enumerate(df.iter_rows(named=True), start=2):
        # Native Dynamic Formula: check if ANY of the status columns are "VALID"
        or_tests = ",".join([f'{col}{row_idx}="VALID"' for col in status_col_letters])
        master_formula = f'=IF(OR({or_tests}), "DIALABLE", "ALL INVALID / UNREACHABLE")'

        row_vals = [master_formula] + [row[c] for c in non_phone_cols]

        for label in ["Mobile", "Telephone", "Landline"]:
            orig_col = col_map.get(label)
            if orig_col:
                row_vals.extend([
                    row[orig_col],
                    row[f"_meta_{label}_status"],
                    row[f"_meta_{label}_local"],
                    row[f"_meta_{label}_e164"],
                    row[f"_meta_{label}_remark"]
                ])

        ws.append(row_vals)

    # Format header styles
    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_idx)
        fill_color = col_style_map.get(col_idx, theme_colors["Master"])
        cell.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")
        cell.font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Format data cells
    for r in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=len(headers)):
        for cell in r:
            cell.font = data_font
            cell.border = thin_border
            if cell.column == 1 or "Status" in headers[cell.column - 1]:
                cell.alignment = Alignment(horizontal="center")

    # Column width auto-fit
    for col in ws.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(min(max_len + 3, 35), 14)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


# =====================================================================
# 4. STREAMLIT UI
# =====================================================================
st.title("Philippine Contact Number Scrubbing Engine")
st.caption("Upload your spreadsheet to validate Mobile (09XX/08XX), Metro Manila Area 02 (8-digit NTC), and Provincial Landlines side-by-side.")

uploaded_file = st.file_uploader("Upload Excel or CSV file", type=["xlsx", "xls", "csv"])

if uploaded_file:
    try:
        if uploaded_file.name.endswith(".csv"):
            df = pl.read_csv(uploaded_file.getvalue(), infer_schema_length=10000)
        else:
            df = pl.read_excel(uploaded_file.getvalue())
    except Exception as e:
        st.error(f"Error reading file: {e}")
        st.stop()

    st.success(f"File loaded: **{df.shape[0]} records**, **{df.shape[1]} columns**")

    cols = ["(None / Skip)"] + df.columns

    def auto_match(pattern: str) -> int:
        for i, c in enumerate(df.columns):
            if re.search(pattern, c, re.IGNORECASE):
                return i + 1
        return 0

    st.subheader("Select Columns to Verify")
    c1, c2, c3 = st.columns(3)

    with c1:
        sel_mobile = st.selectbox(
            "📱 Mobile Phone Column",
            options=cols,
            index=auto_match(r"mobile|cell|cel")
        )
    with c2:
        sel_tele = st.selectbox(
            "☎️ Telephone / Area 02 Column",
            options=cols,
            index=auto_match(r"telephone|tele")
        )
    with c3:
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
        st.warning("Please assign at least one column to verify.")
        st.stop()

    if st.button("Run Full Validation & Clean", type="primary"):
        with st.spinner("Processing records with Pydantic and Polars..."):
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
                meta_dict[f"_meta_{label}_remark"] = [p["remark"] for p in parsed]

            meta_df = pl.DataFrame(meta_dict)
            processed_df = df.hstack(meta_df)

            # Summary Metrics
            st.divider()
            active_labels = [k for k, v in col_map.items() if v]
            m_cols = st.columns(len(active_labels))
            for i, label in enumerate(active_labels):
                valid_num = sum(1 for s in meta_dict[f"_meta_{label}_status"] if s == "VALID")
                m_cols[i].metric(f"{label} Valid", f"{valid_num} / {len(df)}")

            # Grouped Results Preview
            st.subheader("Validation Results (Side-by-Side View)")
            phone_cols_set = {v for v in col_map.values() if v}
            display_cols = [c for c in df.columns if not c.startswith("_meta_") and c not in phone_cols_set]

            for label in ["Mobile", "Telephone", "Landline"]:
                orig_col = col_map.get(label)
                if orig_col:
                    display_cols.extend([
                        orig_col,
                        f"_meta_{label}_status",
                        f"_meta_{label}_local",
                        f"_meta_{label}_remark"
                    ])

            st.dataframe(processed_df.select(display_cols), use_container_width=True)

            # Excel Download with Live Formulas
            excel_bytes = export_multi_column_excel(processed_df, col_map)
            st.download_button(
                label="📥 Download Cleaned Excel File",
                data=excel_bytes,
                file_name=f"Verified_{uploaded_file.name.rsplit('.', 1)[0]}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
