import io
import re
from typing import Optional, List, Dict, Any
import httpx
import polars as pl
from pydantic import BaseModel, Field, field_validator, ValidationError
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# =====================================================================
# 1. PYDANTIC GUARDS (Input Layer Validation)
# =====================================================================
class ContactRecord(BaseModel):
    account_id: str = Field(..., min_length=1, description="Unique account/customer reference")
    contact_name: str = Field(..., min_length=1)
    raw_phone: str = Field(..., min_length=7, description="Phone string before cleaning")

    @field_validator("raw_phone", mode="before")
    @classmethod
    def strip_whitespace(cls, v: Any) -> str:
        if v is None:
            raise ValueError("Phone field cannot be empty or null")
        s = str(v).strip()
        if not s:
            raise ValueError("Phone field cannot be an empty string")
        return s


def validate_raw_upload(data_rows: List[Dict[str, Any]]) -> List[ContactRecord]:
    """Validates raw dictionary inputs before downstream processing."""
    valid_records = []
    errors = []
    for idx, row in enumerate(data_rows):
        try:
            valid_records.append(ContactRecord(**row))
        except ValidationError as e:
            errors.append({"row_index": idx + 2, "details": e.errors()})
    
    if errors:
        raise ValueError(f"Input validation failed on {len(errors)} records: {errors[:3]}")
    return valid_records


# =====================================================================
# 2. DETERMINISTIC FORMAT VALIDATOR & LIVE HLR REACHABILITY
# =====================================================================
VALID_PROVINCIAL_AREAS = {
    "32", "33", "34", "35", "36", "38",
    "42", "43", "44", "45", "46", "47", "48", "49",
    "52", "53", "54", "55", "56",
    "62", "63", "64", "65",
    "72", "74", "75", "77", "78",
    "82", "83", "84", "85", "86", "87", "88"
}

def analyze_ph_syntax(raw_val: str) -> Dict[str, Any]:
    cleaned = re.sub(r'[\s\-\(\)\.]', '', str(raw_val).strip())
    
    # Strip national prefixes
    if cleaned.startswith('+63'):
        cleaned = cleaned[3:]
    elif cleaned.startswith('63'):
        cleaned = cleaned[2:]
    elif cleaned.startswith('0'):
        cleaned = cleaned[1:]

    # Mobile: 10 digits starting with 9 or 8
    if len(cleaned) == 10 and re.match(r'^(9\d{9}|8[1-9]\d{8})$', cleaned):
        return {
            "is_valid_format": True,
            "line_type": "Mobile",
            "clean_number": f"+63{cleaned}",
            "error_reason": None
        }

    # Metro Manila / GMA Landline (Area 2 + 8 local digits with NTC PTE prefix)
    if len(cleaned) == 9 and cleaned.startswith('2'):
        pte_digit = cleaned[1]
        if pte_digit in {'3', '5', '6', '7', '8'}:
            return {
                "is_valid_format": True,
                "line_type": "GMA Landline",
                "clean_number": f"+63{cleaned}",
                "error_reason": None
            }
        return {
            "is_valid_format": False,
            "line_type": "Invalid Landline",
            "clean_number": cleaned,
            "error_reason": "Invalid NTC PTE prefix for Area 02"
        }

    # Provincial Landlines (2-digit area + 7-digit local)
    if len(cleaned) == 9 and cleaned[:2] in VALID_PROVINCIAL_AREAS:
        return {
            "is_valid_format": True,
            "line_type": "Provincial Landline",
            "clean_number": f"+63{cleaned}",
            "error_reason": None
        }

    return {
        "is_valid_format": False,
        "line_type": "Unknown / Bad Format",
        "clean_number": cleaned,
        "error_reason": "Length mismatch or unallocated code"
    }


def verify_hlr_live_status(e164_number: str, api_key: Optional[str] = None) -> str:
    """
    Queries live network registry (HLR).
    Returns: 'ACTIVE', 'DISCONNECTED / ABSENT', or 'NOT_SUPPORTED_FOR_LANDLINE'.
    """
    if not e164_number.startswith("+639") and not e164_number.startswith("+638"):
        return "N/A (Landline)"

    if not api_key:
        # Mocking real-world HLR behavior without live third-party credits
        return "SIMULATED_ACTIVE" if not e164_number.endswith("0000") else "DISCONNECTED"

    try:
        # Example using a standard REST HLR endpoint
        url = "https://api.hlr-lookups.com/v2/lookup"
        headers = {"Authorization": f"Bearer {api_key}"}
        resp = httpx.post(url, json={"msisdn": e164_number}, headers=headers, timeout=4.0)
        
        if resp.status_code == 200:
            status = resp.json().get("status", "").upper()
            return "ACTIVE" if status in ["ACTIVE", "OK", "LIVE"] else "DISCONNECTED"
    except Exception:
        return "NETWORK_TIMEOUT"
    
    return "UNKNOWN"


# =====================================================================
# 3. POLARS BATCH TRANSFORMATION ENGINE
# =====================================================================
def process_contacts_polars(records: List[ContactRecord]) -> pl.DataFrame:
    raw_dicts = [r.model_dump() for r in records]
    df = pl.DataFrame(raw_dicts)

    # 1. Structural parse
    parsed_metadata = [analyze_ph_syntax(p) for p in df["raw_phone"].to_list()]
    parsed_df = pl.DataFrame(parsed_metadata)
    
    df = df.hstack(parsed_df)

    # 2. Check live availability
    reachability = [
        verify_hlr_live_status(row["clean_number"]) if row["is_valid_format"] else "INVALID_FORMAT"
        for row in df.iter_rows(named=True)
    ]
    
    df = df.with_columns(pl.Series("live_status", reachability))
    return df


# =====================================================================
# 4. OPENPYXL WORKBOOK BUILDER (Dynamic Formulas & Layout)
# =====================================================================
def export_to_dynamic_excel(df: pl.DataFrame, output_path: str):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Contact Scrubbing"

    headers = [
        "Account ID", "Customer Name", "Raw Input", 
        "Normalized (E.164)", "Line Type", "Live HLR Status", 
        "Contactable? (Formula)", "Audit Flag (Formula)"
    ]
    ws.append(headers)

    # Data population
    for row_idx, row in enumerate(df.iter_rows(named=True), start=2):
        acc = row["account_id"]
        name = row["contact_name"]
        raw = row["raw_phone"]
        clean = row["clean_number"]
        line_type = row["line_type"]
        hlr = row["live_status"]

        # Native Dynamic Excel formulas (E = Line Type, F = Live HLR Status, C = Raw Input)
        # Column G: Flag as Contactable if valid and not marked disconnected
        formula_contactable = (
            f'=IF(AND(ISNUMBER(SEARCH("Valid", E{row_idx})), F{row_idx}<>"DISCONNECTED"), "YES", '
            f'IF(F{row_idx}="SIMULATED_ACTIVE", "YES", "NO"))'
        )
        
        # Column H: Flag records requiring manual call/dial intervention
        formula_audit = f'=IF(G{row_idx}="NO", "Flagged: Unreachable/Bad", "Ready")'

        ws.append([acc, name, raw, clean, line_type, hlr, formula_contactable, formula_audit])

    # Formatting and Styling
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    data_font = Font(name="Segoe UI", size=10)
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    for col_num in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=len(headers)):
        for cell in row:
            cell.font = data_font
            cell.border = thin_border
            if cell.column in [7, 8]:
                cell.alignment = Alignment(horizontal="center")

    # Dynamic column autosizing
    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 15)

    wb.save(output_path)
