# excel_handler.py
import os
from openpyxl import Workbook, load_workbook
from datetime import date, datetime
from utils import log_error
from shutil import copy2

# Define the standard columns. This is the single source of truth.
DEFAULT_COLS = ["S.No.", "Name", "Email", "Phone", "OriginalFile", "DateApplied", "ResumePath"]
# Column indices for quick lookup (0-based)
EMAIL_COL_IDX = 2
DATE_COL_IDX = 5


def ensure_excel(path: str, columns=None):
    """Creates an Excel file with standard columns if it doesn't exist."""
    if os.path.exists(path):
        return
    wb = Workbook()
    ws = wb.active
    ws.title = "Applicants"
    
    # Use standard columns regardless of input
    ws.append(DEFAULT_COLS)
    
    # Adjust column widths for readability
    try:
        ws.column_dimensions['A'].width = 8  # S.No.
        ws.column_dimensions['B'].width = 30 # Name
        ws.column_dimensions['C'].width = 40 # Email
        ws.column_dimensions['D'].width = 30 # Phone
        ws.column_dimensions['E'].width = 35 # OriginalFile
        ws.column_dimensions['F'].width = 15 # DateApplied
        ws.column_dimensions['G'].width = 50 # ResumePath
    except Exception:
        pass # Not critical
        
    wb.save(path)


def append_row(path: str, data: dict):
    """
    Append a data row to Excel; create file/header if needed.
    data: dict with keys matching columns (Name, Email, Phone, OriginalFile, ResumePath)
    """
    try:
        if os.path.exists(path):
            wb = load_workbook(path)
        else:
            # This case should be handled by ensure_excel, but as a fallback:
            wb = Workbook()

        ws = wb.active
        serial_num = 1
        
        # set header if empty
        if ws.max_row == 1 and ws.cell(1, 1).value is None:
            ws.append(DEFAULT_COLS)
        else:
            # S.No. is the current max_row (since header is 1, first data is 2)
            serial_num = ws.max_row 

        today = date.today().isoformat()
        
        # Build row based on DEFAULT_COLS order
        row = [
            serial_num,
            data.get("Name", ""),
            data.get("Email", ""),
            data.get("Phone", ""),
            data.get("OriginalFile", ""), # New field
            today,
            data.get("ResumePath", "")
        ]
        
        ws.append(row)
        wb.save(path)
    except PermissionError as pe:
        log_error(f"Excel append failed (PermissionError): {pe}. File might be open.")
        raise # Re-raise so UI can handle it
    except Exception as e:
        log_error(f"Excel append failed: {e}\n{repr(e)}")
        raise # Re-raise so UI can handle it


def read_all_rows(path: str):
    """Return list of value-tuples from Excel (skips header)."""
    if not os.path.exists(path):
        return []
    
    rows = []
    try:
        wb = load_workbook(path, data_only=True)
        ws = wb.active
        for r in ws.iter_rows(min_row=2, values_only=True):
            if any(r): # Only add rows that are not entirely empty
                rows.append(r)
    except Exception as e:
        log_error(f"Failed to read Excel file {path}: {e}")
        # Try to return empty list instead of crashing
        return []
    return rows


def email_duplicate_within_days(path: str, email_str: str, days: int) -> bool:
    """
    Return True if any email in email_str exists in Excel and DateApplied is within `days`.
    Conservative behavior: if date parsing fails, treat as duplicate.
    """
    if not os.path.exists(path) or not email_str:
        return False
        
    # Get set of emails to check
    emails_to_check = set(e.strip().lower() for e in email_str.split(',') if e.strip())
    if not emails_to_check:
        return False
        
    try:
        wb = load_workbook(path, data_only=True)
        ws = wb.active

        # Read all rows into memory for faster checking
        # This assumes the excel file is not excessively large (e.g., > 50k rows)
        # For very large files, iter_rows is better but slower.
        all_rows = list(ws.iter_rows(min_row=2, values_only=True))
        if not all_rows:
            return False

        # Create a set of all emails already in the sheet for quick lookup
        existing_emails_map = {} # Map email -> applied_date
        
        for r in all_rows:
            row_email_data = r[EMAIL_COL_IDX] if len(r) > EMAIL_COL_IDX else None
            row_date = r[DATE_COL_IDX] if len(r) > DATE_COL_IDX else None

            if not row_email_data:
                continue
                
            # Handle multiple emails in the cell
            row_emails = [e.strip().lower() for e in str(row_email_data).split(',') if e.strip()]
            for remail in row_emails:
                if remail not in existing_emails_map:
                     existing_emails_map[remail] = row_date # Store first date found

        # Now check for duplicates
        for check_email in emails_to_check:
            if check_email in existing_emails_map:
                row_date = existing_emails_map[check_email]
                # If there's no date stored, treat as duplicate (safer)
                if not row_date:
                    return True

                # Normalize various date types (datetime, date, iso string)
                try:
                    if isinstance(row_date, str):
                        applied_dt = datetime.fromisoformat(row_date)
                    elif isinstance(row_date, datetime):
                        applied_dt = row_date
                    elif isinstance(row_date, date):
                        # row_date is date-like (datetime.date)
                        applied_dt = datetime.combine(row_date, datetime.min.time())
                    else:
                        return True # Unknown date format, treat as duplicate

                    delta_days = (datetime.now() - applied_dt).days
                    if delta_days <= int(days):
                        return True # Found a duplicate within the window
                except Exception:
                    # If parsing date fails, conservatively treat as duplicate
                    return True
                    
    except Exception as e:
        log_error(f"email_duplicate_within_days error: {e}\n{repr(e)}")
        # Be conservative on error
        return True # Treat as duplicate if check fails

    return False


def save_to_excel(data: dict, excel_path: str = "resumes_data.xlsx", columns=None):
    """
    Compatibility wrapper.
    This simply calls append_row(...) from this module.
    """
    append_row(excel_path, data)        # set header if empty
        if ws.max_row == 1 and ws.cell(1, 1).value is None:
            if not columns:
                columns = ["Name", "Email", "Phone", "Skills", "Experience", "DateApplied", "ResumePath"]
            ws.append(columns)

        today = date.today().isoformat()
        row = [
            data.get("Name", ""),
            data.get("Email", ""),
            data.get("Phone", ""),
            data.get("Skills", ""),
            data.get("Experience", ""),
            today,
            data.get("ResumePath", "")
        ]
        ws.append(row)
        wb.save(path)
    except Exception as e:
        # log the error and re-raise so calling code can handle it
        try:
            log_error("Excel append failed: " + str(e) + "\n" + repr(e))
        except Exception:
            pass
        raise


def read_all_rows(path: str):
    """Return list of value-tuples from Excel (skips header)."""
    if not os.path.exists(path):
        return []
    wb = load_workbook(path, data_only=True)
    ws = wb.active
    rows = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        rows.append(r)
    return rows


def email_duplicate_within_days(path: str, email: str, days: int) -> bool:
    """
    Return True if same email exists in Excel and DateApplied is within `days`.
    Conservative behavior: if date parsing fails, treat as duplicate.
    """
    if not os.path.exists(path) or not email:
        return False
    try:
        wb = load_workbook(path, data_only=True)
        ws = wb.active
        from datetime import datetime

        email_norm = str(email).strip().lower()
        for r in ws.iter_rows(min_row=2, values_only=True):
            # Expecting row like: Name, Email, Phone, Skills, Experience, DateApplied, ResumePath
            row_email = r[1] if len(r) > 1 else None
            row_date = r[5] if len(r) > 5 else None

            if not row_email:
                continue
            if str(row_email).strip().lower() != email_norm:
                continue

            # If there's no date stored, treat as duplicate (safer)
            if not row_date:
                return True

            # Normalize various date types (datetime, date, iso string)
            try:
                if isinstance(row_date, str):
                    applied_dt = datetime.fromisoformat(row_date)
                else:
                    # openpyxl may return a datetime.date or datetime.datetime
                    if isinstance(row_date, datetime):
                        applied_dt = row_date
                    else:
                        # row_date is date-like (datetime.date)
                        applied_dt = datetime.combine(row_date, datetime.min.time())
                delta_days = (datetime.now() - applied_dt).days
                if delta_days <= int(days):
                    return True
            except Exception:
                # If parsing date fails, conservatively treat as duplicate
                return True
    except Exception as e:
        try:
            log_error("email_duplicate_within_days error: " + str(e) + "\n" + repr(e))
        except Exception:
            pass
        return False

    return False


def save_to_excel(data: dict, excel_path: str = "resumes_data.xlsx", columns=None):
    """
    Compatibility wrapper so other modules can call save_to_excel(...).
    This simply calls append_row(...) from this module.
    """
    append_row(excel_path, data, columns)

def get_headers(path: str):
    """
    Return a list of header strings from the first row of the active sheet.
    If file does not exist, returns empty list.
    """
    if not os.path.exists(path):
        return []
    wb = load_workbook(path, data_only=True)
    ws = wb.active
    headers = []
    for cell in ws[1]:
        headers.append(cell.value if cell.value is not None else "")
    return headers

def update_headers(path: str, new_headers):
    """
    Replace header row of the Excel file with new_headers (list of str).
    Backup created: <file>.bak
    Data rows are padded/truncated to match new header length.
    """
    if not os.path.exists(path):
        raise FileNotFoundError(f"Excel file not found: {path}")

    # create backup
    bak_path = path + ".bak"
    try:
        copy2(path, bak_path)
    except Exception:
        # if backup fails, set bak_path = None but continue
        bak_path = None

    wb = load_workbook(path)
    ws = wb.active

    # collect existing rows (values_only)
    rows = []
    for r in ws.iter_rows(values_only=True):
        rows.append(list(r) if r is not None else [])

    data_rows = rows[1:] if len(rows) > 1 else []

    # build new workbook
    new_wb = Workbook()
    new_ws = new_wb.active
    new_ws.title = ws.title if ws.title else "Applicants"

    # write header
    new_ws.append(list(new_headers))

    new_len = len(new_headers)
    for r in data_rows:
        r = r if r is not None else []
        # ensure exact length
        row_vals = r[:new_len] + [""] * max(0, new_len - len(r))
        new_ws.append(row_vals)

    new_wb.save(path)
    return True

