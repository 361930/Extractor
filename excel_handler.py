# excel_handler.py
import os
from openpyxl import Workbook, load_workbook
from openpyxl.utils.exceptions import InvalidFileException
# --- NEW: Imports for Excel Data Validation ---
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter
# --- END NEW ---
from shutil import copy2
from utils import log_error
import zipfile # Added for error handling

# Define the standard columns
DEFAULT_COLS = ["S.No.", "Name", "Email", "Phone", "Status"]

# --- NEW: Reusable function to add dropdowns ---
def _apply_dropdown_validation(ws):
    """Applies data validation dropdowns to the 'Status' column of a worksheet."""
    try:
        # Define the dropdown options (must match app.py)
        STATUS_OPTIONS = [
            "", "Accepted", "Rejected", "On Hold", 
            "Interview Scheduled", "Pending Review"
        ]
        
        # Create the data validation rule
        dv = DataValidation(type="list", formula1=f'"{",".join(STATUS_OPTIONS)}"', allow_blank=True)
        dv.error = "Your entry is not in the list."
        dv.errorTitle = "Invalid Entry"
        dv.prompt = "Please select from the list."
        dv.promptTitle = "Select Status"
        
        # Add it to the worksheet
        ws.add_data_validation(dv)
        
        # Find the column letter for "Status"
        status_col_index = -1
        for i, cell in enumerate(ws[1]): # Iterate over header row
            if cell.value == "Status":
                status_col_index = i + 1
                break
        
        if status_col_index == -1:
            print("Could not find 'Status' column to apply dropdowns.")
            return

        status_col_letter = get_column_letter(status_col_index)
        
        # Apply the rule to the entire "Status" column (skipping header)
        dv.add(f'{status_col_letter}2:{status_col_letter}1048576')
        print(f"Applied status dropdowns to worksheet.")
        
    except Exception as e:
        log_error(f"Failed to add data validation to Excel: {e}")
# --- END NEW ---


def validate_or_create_excel(path: str):
    """
    Ensures the Excel file at 'path' exists and has the correct headers.
    If it exists with *incorrect* headers, it backs up the old file
    and creates a new, correct one.
    """
    if not os.path.exists(path):
        # File doesn't exist, create a new one
        _create_new_excel(path)
        return

    try:
        wb = load_workbook(path)
        ws = wb.active
        
        # Read headers from the first row
        existing_headers = [cell.value for cell in ws[1]]
        
        if existing_headers == DEFAULT_COLS:
            # --- FIX: Apply validation to existing file ---
            print("Excel file is valid. Applying/checking dropdowns...")
            _apply_dropdown_validation(ws)
            wb.save(path) # Save any changes
            # --- END FIX ---
            return
        else:
            # Headers are wrong! Backup and create new.
            log_error(f"Excel headers mismatch in {path}. Backing up and creating new file.")
            
            bak_path = path + ".bak"
            i = 1
            while os.path.exists(bak_path):
                bak_path = f"{path}.bak{i}"
                i += 1
                
            os.rename(path, bak_path)
            log_error(f"Backed up old file to {bak_path}")
            
            _create_new_excel(path)
            
    except (InvalidFileException, KeyError, zipfile.BadZipFile):
        # File is corrupted or not an Excel file
        log_error(f"Corrupted Excel file {path}. Moving it and creating new.")
        if os.path.exists(path):
            os.remove(path)
        _create_new_excel(path)
    except Exception as e:
        log_error(f"Error validating Excel file {path}: {e}. Creating new.")
        if os.path.exists(path):
            os.remove(path) # Remove to avoid loops
        _create_new_excel(path)

def _create_new_excel(path: str):
    """Internal helper to create a new Excel file with headers."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Applicants"
    ws.append(DEFAULT_COLS)
    
    # --- NEW: Call the reusable function ---
    _apply_dropdown_validation(ws)
    # --- END NEW ---
    
    wb.save(path)

def get_next_serial_number(path: str) -> int:
    """Gets the next serial number by reading the last S.No. in the file."""
    if not os.path.exists(path):
        return 1
    try:
        wb = load_workbook(path)
        ws = wb.active
        for row in range(ws.max_row, 1, -1):
            cell_value = ws.cell(row=row, column=1).value
            if cell_value is not None:
                try:
                    return int(cell_value) + 1
                except ValueError:
                    continue 
        return 1
    except (InvalidFileException, KeyError, zipfile.BadZipFile):
        log_error(f"Corrupted Excel file {path} while getting S.No. Creating new.")
        os.remove(path)
        _create_new_excel(path)
        return 1
    except Exception as e:
        log_error(f"Error reading serial number from {path}: {e}")
        return 1

# --- FIX: Bug `TypeError: append_row() missing...` ---
# This version matches the call in the corrected app.py
def append_row(path: str, data: dict) -> int:
    """
    Append a data row to Excel. Returns the new serial number.
    data: dict with keys "Name", "Email", "Phone"
    """
    try:
        wb = load_workbook(path)
        ws = wb.active
        
        serial_num = get_next_serial_number(path)
        
        # --- FIX: Handle email being a list or string ---
        email_val = data.get("Email", "")
        if isinstance(email_val, list):
            email_val = ", ".join(email_val)
        
        row = [
            serial_num,
            data.get("Name", ""),
            email_val,
            data.get("Phone", ""),
            "", # Status - defaults to empty
        ]
        ws.append(row)
        wb.save(path)
        return serial_num
    except Exception as e:
        log_error(f"Excel append failed: {e}")
        raise

def read_all_rows(path: str):
    """Return list of value-tuples from Excel (skips header)."""
    if not os.path.exists(path):
        return []
    try:
        wb = load_workbook(path, data_only=True)
        ws = wb.active
        rows = []
        for r in ws.iter_rows(min_row=2, values_only=True):
            row_data = list(r)
            if len(row_data) < len(DEFAULT_COLS):
                row_data.extend([""] * (len(DEFAULT_COLS) - len(row_data)))
            elif len(row_data) > len(DEFAULT_COLS):
                row_data = row_data[:len(DEFAULT_COLS)]
            rows.append(tuple(row_data))
        return rows
    except Exception as e:
        log_error(f"Failed to read Excel {path}: {e}")
        return []

def email_duplicate_within_days(path: str, email: str, days: int) -> bool:
    """
    Checks for email duplicates.
    NOTE: This is not currently used in the app, but kept for future use.
    """
    return False

def update_status(path: str, serial_num: int, new_status: str) -> bool:
    """Finds a row by its S.No. and updates its Status."""
    if not os.path.exists(path):
        return False
    try:
        wb = load_workbook(path)
        ws = wb.active
        
        status_col = -1
        for c_idx, cell in enumerate(ws[1], 1):
            if cell.value == "Status":
                status_col = c_idx
                break
        
        if status_col == -1:
            log_error("Could not find 'Status' column in Excel.")
            return False

        row_to_update = -1
        for r_idx, row in enumerate(ws.iter_rows(min_row=2), 2):
            if row[0].value == serial_num:
                row_to_update = r_idx
                break
        
        if row_to_update != -1:
            ws.cell(row=row_to_update, column=status_col, value=new_status)
            wb.save(path)
            return True
        else:
            log_error(f"Could not find S.No. {serial_num} to update status.")
            return False
    except Exception as e:
        log_error(f"Failed to update status for S.No. {serial_num}: {e}")
        return False

def export_by_status(excel_path: str, destination_folder: str):
    """
    Reads the master Excel file and creates a new Excel file
    for EACH unique status found (e.g., Accepted_Candidates.xlsx,
    Rejected_Candidates.xlsx, On_Hold_Candidates.xlsx).
    """
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"Master file not found: {excel_path}")

    master_wb = load_workbook(excel_path, data_only=True)
    master_ws = master_wb.active
    
    headers = [cell.value for cell in master_ws[1]]
    
    try:
        status_col_index = headers.index("Status")
    except ValueError:
        raise ValueError("Master file is missing the 'Status' column.")

    # --- UPDATED: Dynamic Status Handling ---
    
    # 1. Find all unique statuses and create workbooks for them
    status_workbooks = {} 
    
    for row in master_ws.iter_rows(min_row=2, values_only=True):
        if not row or len(row) <= status_col_index:
            continue
            
        status = row[status_col_index]
        
        if not status:
            continue
            
        if status not in status_workbooks:
            wb = Workbook()
            ws = wb.active
            ws.title = status[:30]
            ws.append(headers)
            status_workbooks[status] = {"wb": wb, "ws": ws}
            
        status_workbooks[status]["ws"].append(row)

    # 2. Save all the new workbooks
    created_file_paths = []
    
    if not status_workbooks:
        return []

    for status, data in status_workbooks.items():
        safe_filename = str(status).replace(" ", "_").replace("/", "-")
        
        file_path = os.path.join(destination_folder, f"{safe_filename}_Candidates.xlsx")
        data["wb"].save(file_path)
        created_file_paths.append(file_path)
    
    return created_file_paths
