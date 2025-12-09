# app.py — GUI Application
import os
import sys
import threading
import traceback
import subprocess
from pathlib import Path
from shutil import copy2
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import Set
# import time # No longer needed

# --- 'watchdog' imports have been REMOVED ---

# App modules
from utils import load_config, save_config, log_error, ensure_dirs, check_ollama_connection
from parser import parse_resume
from excel_handler import (
    validate_or_create_excel,
    read_all_rows, append_row, 
    update_status, export_by_status, DEFAULT_COLS
)

# ----------------- Configuration / Workspace -----------------
HOME = Path.home()
WORKSPACE = HOME / "Desktop" / "ResumeParserWorkspace"
RESUMES_DIR = WORKSPACE / "resumes"
EXCEL_DIR = WORKSPACE / "excel"
ensure_dirs() 
WORKSPACE.mkdir(parents=True, exist_ok=True)
RESUMES_DIR.mkdir(parents=True, exist_ok=True)
EXCEL_DIR.mkdir(parents=True, exist_ok=True)

CONFIG = load_config()

# --- Validate Active Excel Path ---
active_excel_path_str = CONFIG.get("active_excel", str(EXCEL_DIR / "resumes_data.xlsx"))
ACTIVE_EXCEL = Path(active_excel_path_str)

if not ACTIVE_EXCEL.exists() and ACTIVE_EXCEL.parent != EXCEL_DIR:
    ACTIVE_EXCEL = EXCEL_DIR / "resumes_data.xlsx"
    CONFIG["active_excel"] = str(ACTIVE_EXCEL)
    save_config(CONFIG)

DUPLICATE_CHECK_ENABLED = CONFIG.get("duplicate_check_enabled", True)

# --- NEW: Status Options ---
STATUS_OPTIONS = [
    "", 
    "Accepted", 
    "Rejected", 
    "On Hold", 
    "Interview Scheduled", 
    "Pending Review"
]

# ----------------- Load Ollama model -----------------
try:
    # Check connection and model
    success, model_or_error = check_ollama_connection()
    if not success:
        raise OSError(model_or_error)
    MODEL_DISPLAY_NAME = f"(Model: {model_or_error} via Ollama)"
    # nlp=None is just a placeholder, parser.py doesn't use it
    nlp = None
except Exception as e:
    nlp = None
    MODEL_DISPLAY_NAME = "(Model: NOT LOADED)"
    log_error(f"Ollama connection error: {e}\n{traceback.format_exc()}")
    
    # We must show the error on startup
    # Create a dummy root to show the message
    dummy_root = tk.Tk()
    dummy_root.withdraw()
    messagebox.showerror("Fatal Model Error",
                         "Could not connect to the Ollama server.\n\n"
                         f"Error: {e}\n\n"
                         "Please ensure:\n"
                         "1. The Ollama application is running.\n"
                         "2. You have run: `ollama pull llama3:8b`")
    dummy_root.destroy()
    # We let the app continue, but parsing will fail.

# ----------------- Global Threading Control -----------------
PARSING_THREAD = None
STOP_EVENT = None
# --- REMOVED: File Watcher globals ---


# ----------------- Main window -----------------
root = tk.Tk()
root.title(f"Resume Parser {MODEL_DISPLAY_NAME}")
root.geometry("1200x700")

# --- Styling ---
style = ttk.Style(root)
style.theme_use('clam') 
style.configure("Treeview.Heading", font=('Arial', 10, 'bold'))
style.configure("Details.TFrame", background="#f0f0f0")
style.configure("Details.TLabel", background="#f0f0f0", font=('Arial', 10))
style.configure("Details.TButton", font=('Arial', 10, 'bold'))
style.configure("Title.TLabel", background="#f0f0f0", font=('Arial', 12, 'bold'))
style.configure("Stop.TButton", font=('Arial', 10, 'bold'), foreground="red", background="white")


# ----------------- PanedWindow Layout -----------------
paned_window = ttk.PanedWindow(root, orient=tk.VERTICAL)
paned_window.pack(fill=tk.BOTH, expand=True)

# --- Top Pane: Controls and Treeview ---
top_pane = ttk.Frame(paned_window, height=450)
paned_window.add(top_pane, weight=4) 

# Top controls
frame_top_controls = ttk.Frame(top_pane)
frame_top_controls.pack(fill="x", padx=10, pady=8)

# Left side controls
frame_left_controls = ttk.Frame(frame_top_controls)
frame_left_controls.pack(side="left")

btn_single = ttk.Button(frame_left_controls, text="Upload Resume", command=lambda: on_upload("file"))
btn_single.pack(side="left", padx=6)

btn_multi = ttk.Button(frame_left_controls, text="Upload Multiple Resumes", command=lambda: on_upload("files"))
btn_multi.pack(side="left", padx=6)

btn_folder = ttk.Button(frame_left_controls, text="Upload Folder", command=lambda: on_upload("folder"))
btn_folder.pack(side="left", padx=6)

btn_stop = ttk.Button(
    frame_left_controls, 
    text="Stop Parsing", 
    style="Stop.TButton",
    command=lambda: on_stop_parsing()
)

var_dup = tk.BooleanVar(value=DUPLICATE_CHECK_ENABLED)
chk_dup = ttk.Checkbutton(
    frame_left_controls, 
    text="Check for Duplicates", 
    variable=var_dup,
    command=lambda: save_dup_config()
)
chk_dup.pack(side="left", padx=12, pady=5)


# Right side controls
frame_right_controls = ttk.Frame(frame_top_controls)
frame_right_controls.pack(side="right")

btn_export_sorted = ttk.Button(frame_right_controls, text="Export Sorted Files", command=lambda: export_sorted_files())
btn_export_sorted.pack(side="right", padx=6)

btn_open_excel = ttk.Button(frame_right_controls, text="Open Active Excel",
                            command=lambda: open_in_explorer(ACTIVE_EXCEL))
btn_open_excel.pack(side="right", padx=6)

# Excel selection frame
frame_excel_bar = ttk.Frame(top_pane)
frame_excel_bar.pack(fill="x", padx=10, pady=(0, 5))

btn_change_excel = ttk.Button(frame_excel_bar, text="Set Output Excel", width=16, command=lambda: select_active_excel())
btn_change_excel.pack(side="left", padx=6)

lbl_active = ttk.Label(frame_excel_bar, text=f"Active Excel: {ACTIVE_EXCEL.name}", anchor="w")
lbl_active.pack(side="left", fill="x", expand=True, padx=6)

# Search
frame_search = ttk.Frame(top_pane)
frame_search.pack(fill="x", padx=10, pady=6)
ttk.Label(frame_search, text="Search:").pack(side="left", padx=(6,6))
search_var = tk.StringVar()
entry_search = ttk.Entry(frame_search, textvariable=search_var, width=50)
entry_search.pack(side="left", fill="x", expand=True, padx=(0,6))

# Tree (results)
# --- UPDATED: Removed 'OriginalFile' column ---
tree_cols = ["S.No.", "Name", "Email", "Phone", "Status"]
tree = ttk.Treeview(
    top_pane, 
    columns=tree_cols, 
    show="headings",
    displaycolumns=tree_cols # Show all columns
)

tree.heading("S.No.", text="S.No.", anchor="w")
tree.column("S.No.", width=60, anchor="w")

tree.heading("Name", text="Name", anchor="w")
tree.column("Name", width=200, anchor="w")

tree.heading("Email", text="Email", anchor="w")
tree.column("Email", width=250, anchor="w")

tree.heading("Phone", text="Phone", anchor="w")
tree.column("Phone", width=150, anchor="w")

tree.heading("Status", text="Status", anchor="w")
tree.column("Status", width=150, anchor="w") # Made wider for new statuses

tree.pack(fill="both", expand=True, padx=10, pady=(0, 10))

# --- Bottom Pane: Details Panel ---
bottom_pane = ttk.Frame(paned_window, height=250, style="Details.TFrame")
paned_window.add(bottom_pane, weight=1)

ttk.Label(bottom_pane, text="Candidate Details", style="Title.TLabel").pack(anchor="w", padx=12, pady=(10, 6))

details_frame = ttk.Frame(bottom_pane, style="Details.TFrame")
details_frame.pack(fill="x", expand=True, padx=12)

# Grid layout for details
details_frame.columnconfigure(1, weight=1)

# --- Variables for Details Panel ---
detail_sno_var = tk.StringVar(value="-")
detail_name_var = tk.StringVar(value="-")
detail_email_var = tk.StringVar(value="-")
detail_phone_var = tk.StringVar(value="-")
detail_status_var = tk.StringVar()
# REMOVED: detail_original_file_var

# Row 0: S.No.
ttk.Label(details_frame, text="S.No:", style="Details.TLabel", font=('Arial', 10, 'bold')).grid(row=0, column=0, sticky="w", padx=5, pady=2)
ttk.Label(details_frame, textvariable=detail_sno_var, style="Details.TLabel").grid(row=0, column=1, sticky="w", padx=5, pady=2)

# Row 1: Name
ttk.Label(details_frame, text="Name:", style="Details.TLabel", font=('Arial', 10, 'bold')).grid(row=1, column=0, sticky="w", padx=5, pady=2)
ttk.Label(details_frame, textvariable=detail_name_var, style="Details.TLabel").grid(row=1, column=1, sticky="w", padx=5, pady=2)

# Row 2: Email
ttk.Label(details_frame, text="Email:", style="Details.TLabel", font=('Arial', 10, 'bold')).grid(row=2, column=0, sticky="w", padx=5, pady=2)
ttk.Label(details_frame, textvariable=detail_email_var, style="Details.TLabel").grid(row=2, column=1, sticky="w", padx=5, pady=2)

# Row 3: Phone
ttk.Label(details_frame, text="Phone:", style="Details.TLabel", font=('Arial', 10, 'bold')).grid(row=3, column=0, sticky="w", padx=5, pady=2)
ttk.Label(details_frame, textvariable=detail_phone_var, style="Details.TLabel").grid(row=3, column=1, sticky="w", padx=5, pady=2)

# Row 4: Status Dropdown
ttk.Label(details_frame, text="Status:", style="Details.TLabel", font=('Arial', 10, 'bold')).grid(row=4, column=0, sticky="w", padx=5, pady=8)
status_combobox = ttk.Combobox(
    details_frame,
    textvariable=detail_status_var,
    values=STATUS_OPTIONS, # --- UPDATED: Use new status list ---
    state="readonly",
    width=25 # --- NEW: Made wider ---
)
status_combobox.grid(row=4, column=1, sticky="w", padx=5, pady=8)

# Row 5: Save Button
btn_save_status = ttk.Button(
    details_frame, 
    text="Save Status", 
    style="Details.TButton",
    command=lambda: save_candidate_status()
)
btn_save_status.grid(row=5, column=1, sticky="w", padx=5, pady=10)

# REMOVED: btn_view_resume

# Store the currently selected real S.No. (IID)
current_selected_sno = None

# --- Status Bar ---
frame_status_bar = ttk.Frame(root)
frame_status_bar.pack(fill="x", side=tk.BOTTOM, padx=0, pady=0)
progress = ttk.Progressbar(frame_status_bar, orient="horizontal", mode="determinate")
progress.pack(fill="x", padx=12, pady=6)
lbl_status = ttk.Label(frame_status_bar, text="Ready", anchor="w")
lbl_status.pack(fill="x", padx=12, pady=(0, 6))


# ----------------- GUI Helper Functions -----------------
def ui_call(fn, *a, **kw):
    """Helper to run a function on the main GUI thread."""
    # Check if root window still exists
    if root.winfo_exists():
        root.after(0, lambda: fn(*a, **kw))

def add_status(text: str):
    """Append a status line to the status bar (thread-safe)."""
    ui_call(lbl_status.config, text=text)

def open_in_explorer(path: Path):
    try:
        if not path.exists():
            messagebox.showwarning("Not found", f"{path} does not exist.")
            return
        if os.name == "nt":
            os.startfile(path)
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(path)])
        else:
            subprocess.Popen(["xdg-open", str(path)])
    except Exception as e:
        messagebox.showerror("Open Error", f"Could not open {path}: {e}")

def set_active_excel(path: Path):
    """Updates the global ACTIVE_EXCEL, saves config, and updates UI."""
    global ACTIVE_EXCEL
    ACTIVE_EXCEL = path
    CONFIG["active_excel"] = str(ACTIVE_EXCEL)
    save_config(CONFIG)
    lbl_active.config(text=f"Active Excel: {ACTIVE_EXCEL.name}")
    validate_or_create_excel(str(ACTIVE_EXCEL)) 
    refresh_tree_from_excel()
    add_status(f"Active Excel set to {ACTIVE_EXCEL.name}")

def select_active_excel():
    """Lets user create or select an Excel file."""
    file = filedialog.asksaveasfilename(
        initialdir=str(EXCEL_DIR),
        title="Select or Create Excel File",
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")]
    )
    if not file:
        return
    set_active_excel(Path(file))

def save_dup_config():
    """Saves the duplicate check setting to config.json."""
    CONFIG["duplicate_check_enabled"] = var_dup.get()
    save_config(CONFIG)
    add_status(f"Duplicate check {'enabled' if var_dup.get() else 'disabled'}.")

# ----------------- GUI Behavior Functions -----------------
def refresh_tree_from_excel():
    """Load rows from Excel into the tree."""
    # --- REMOVED: global ITEMS_BEING_ADDED ---
    
    # Clear all items from the tree
    for r in tree.get_children():
        tree.delete(r)
    
    rows = read_all_rows(str(ACTIVE_EXCEL))
    if not rows:
        return
        
    try:
        headers = DEFAULT_COLS
        sno_idx = headers.index("S.No.")
        name_idx = headers.index("Name")
        email_idx = headers.index("Email")
        phone_idx = headers.index("Phone")
        status_idx = headers.index("Status")
        # REMOVED: original_file_idx
    except ValueError as e:
        messagebox.showerror("Excel Error", f"Your Excel file is missing a required column: {e}")
        return

    # Add data to tree
    for i, r in enumerate(rows):
        real_sno = r[sno_idx]
        if real_sno is None: # Skip empty S.No. rows
            continue
            
        # --- Simplified: iid can be a string or int, but we'll cast to int ---
        real_sno_int = int(real_sno)
        
        # --- FIX: Prevent `TclError: Item X already exists` ---
        # This check is still good, though the watcher was the main cause
        if tree.exists(real_sno_int):
            # Just update the values, in case they changed
            display_sno = tree.item(real_sno_int, "values")[0]
            tree.item(real_sno_int, values=(
                display_sno,
                r[name_idx],
                r[email_idx],
                r[phone_idx],
                r[status_idx]
            ))
            continue
            
        # If it's a new item, add it
        display_sno = i + 1
        
        tree.insert(
            "", 
            "end", 
            iid=real_sno_int, 
            values=(
                display_sno,
                r[name_idx],
                r[email_idx],
                r[phone_idx],
                r[status_idx]
            )
        )
    
    # --- REMOVED: ITEMS_BEING_ADDED = set() ---

def on_tree_select(event):
    """Fired when a user clicks a row. Populates the details panel."""
    global current_selected_sno
    
    selected_item = tree.focus() 
    if not selected_item:
        return
    
    try:
        current_selected_sno = int(selected_item) 
    except ValueError:
        current_selected_sno = None
        return 
    
    values = tree.item(selected_item, "values")
    
    # values = (display_sno, name, email, phone, status)
    detail_sno_var.set(str(current_selected_sno)) 
    detail_name_var.set(values[1])
    detail_email_var.set(values[2])
    detail_phone_var.set(values[3])
    detail_status_var.set(values[4])
    # REMOVED: detail_original_file_var

tree.bind("<<TreeviewSelect>>", on_tree_select)

# REMOVED: open_original_file()

def save_candidate_status():
    """Saves the status from the combobox to the Excel file."""
    if current_selected_sno is None:
        messagebox.showwarning("No Candidate", "Please select a candidate from the list.")
        return
        
    new_status = detail_status_var.get()
    
    def save_in_thread():
        try:
            success = update_status(str(ACTIVE_EXCEL), current_selected_sno, new_status)
            if success:
                ui_call(add_status, f"Status updated for S.No. {current_selected_sno}")
                # --- UPDATE: Manually update the tree row after saving ---
                # (display_sno, name, email, phone, status)
                current_values = list(tree.item(current_selected_sno, "values"))
                current_values[4] = new_status # Update status
                ui_call(tree.item, current_selected_sno, values=tuple(current_values))
            else:
                ui_call(add_status, f"Error updating status for S.No. {current_selected_sno}")
        except Exception as e:
            ui_call(messagebox.showerror, "Error", f"Failed to save status: {e}")

    threading.Thread(target=save_in_thread, daemon=True).start()


def on_search_change(*_):
    """Filters the tree view based on the search query."""
    q = search_var.get().strip().lower()
    for item_id in tree.get_children(): 
        values = tree.item(item_id, "values")
        # values = (display_sno, name, email, phone, status)
        # Search in name, email, phone, status (cols 1, 2, 3, 4)
        combined = " ".join([str(v).lower() for v in values[1:5]])
        
        if q == "" or q in combined:
            tree.reattach(item_id, "", "end")
        else:
            tree.detach(item_id)

entry_search.bind("<KeyRelease>", on_search_change)

# ----------------- Processing Worker -----------------
def process_files_sequential(file_paths, check_duplicates, stop_event):
    """Background worker to parse files, save to Excel, update UI."""
    # --- REMOVED: global ITEMS_BEING_ADDED ---
    
    ui_call(progress.config, mode="determinate", maximum=len(file_paths), value=0)
    add_status(f"Starting to process {len(file_paths)} files...")
    
    success_count = 0
    fail_count = 0
    skip_count = 0
    stopped = False
    
    existing_emails = set()
    if check_duplicates:
        try:
            add_status("Scanning for existing emails...")
            rows = read_all_rows(str(ACTIVE_EXCEL))
            email_idx = DEFAULT_COLS.index("Email")
            for r in rows:
                if r[email_idx]:
                    # --- FIX: Handle emails being a list (from previous bug) ---
                    emails_in_cell_raw = r[email_idx]
                    if isinstance(emails_in_cell_raw, list):
                        emails_in_cell_raw = ", ".join(emails_in_cell_raw)
                    
                    emails_in_cell = emails_in_cell_raw.lower().split(',')
                    for e in emails_in_cell:
                        if e.strip():
                            existing_emails.add(e.strip())
            add_status(f"Found {len(existing_emails)} unique emails in database.")
        except Exception as e:
            log_error(f"Error reading emails for duplicate check: {e}")
            ui_call(messagebox.showwarning, "Duplicate Check Error", 
                    f"Could not read existing emails: {e}\nDuplicate check may be incomplete.")
    
    for i, fp in enumerate(file_paths):
        if stop_event.is_set():
            stopped = True
            add_status(f"Parsing stopped by user at file {i+1}.")
            break 
            
        fname = Path(fp).name
        try:
            add_status(f"Processing ({i+1}/{len(file_paths)}): {fname}...")
            # We pass the dummy `nlp` object, it isn't used by the new parser
            data = parse_resume(fp, nlp) 
            if not data:
                # --- UPDATE: More helpful error ---
                add_status(f"Failed: {fname}. See logs/errors.log for details.")
                fail_count += 1
                continue

            # --- Duplicate Check Logic (Handles list and str) ---
            if check_duplicates:
                parsed_emails_raw = data.get("Email", "")
                parsed_emails = []
                
                # --- FIX: Handle 'list' object has no attribute 'lower' ---
                if isinstance(parsed_emails_raw, list):
                    parsed_emails = [e.lower().strip() for e in parsed_emails_raw if e.strip()]
                elif isinstance(parsed_emails_raw, str):
                    parsed_emails = [e.lower().strip() for e in parsed_emails_raw.split(',') if e.strip()]

                found_duplicate = False
                for e in parsed_emails:
                    if e in existing_emails:
                        found_duplicate = True
                        break
                
                if found_duplicate:
                    add_status(f"Skipped: {fname} (Email already exists)")
                    skip_count += 1
                    ui_call(progress.step, 1) 
                    continue
            
            # --- FIX: `append_row() missing 1...` ---
            # The function signature was changed in excel_handler.py
            new_sno = append_row(str(ACTIVE_EXCEL), data)
            
            # --- FIX: Re-added the UI update logic ---
            # Now that the watcher is gone, we add the row directly.
            # This fixes the "not processing" bug.
            display_sno = len(tree.get_children()) + 1
            ui_call(
                tree.insert,
                "",
                "end",
                iid=new_sno, # Use real S.No. as the unique ID
                values=(
                    display_sno,
                    data.get("Name", ""),
                    data.get("Email", ""),
                    data.get("Phone", ""),
                    "" # Default empty status
                )
            )
            # --- END FIX ---
            
            success_count += 1
            
            if check_duplicates:
                for e in parsed_emails:
                    existing_emails.add(e)

        except Exception as e:
            log_error(f"Processing error for {fp}: {e}\n" + traceback.format_exc())
            add_status(f"Error processing {fname}: {e}")
            fail_count += 1
        finally:
            ui_call(progress.step, 1)

    ui_call(progress.config, value=0)
    if stopped:
        summary = f"Process stopped. Added: {success_count}, Failed: {fail_count}, Skipped: {skip_count}."
    else:
        summary = f"Batch complete: {success_count} added, {fail_count} failed, {skip_count} skipped (duplicates)."
    
    add_status(summary)
    if not stopped: 
        ui_call(messagebox.showinfo, "Batch Complete", 
                            f"Processed {len(file_paths)} files.\n\n"
                            f"Added: {success_count}\n"
                            f"Failed: {fail_count}\n"
                            f"Skipped (Duplicates): {skip_count}")
    
    ui_call(parsing_complete)


def safe_copy_to_workspace(src_path: str):
    """
    Copies a file to the RESUMES_DIR, avoiding name collisions.
    Returns the destination Path.
    """
    src = Path(src_path)
    dst = RESUMES_DIR / src.name
    
    base = dst.stem
    ext = dst.suffix
    i = 1
    while dst.exists():
        dst = RESUMES_DIR / f"{base}_{i}{ext}"
        i += 1
    
    copy2(src, dst)
    return dst

def on_upload(upload_type: str):
    """Handles all upload button clicks."""
    global PARSING_THREAD, STOP_EVENT
    
    if PARSING_THREAD and PARSING_THREAD.is_alive():
        messagebox.showwarning("Busy", "A parsing job is already running.")
        return
        
    file_paths = []
    if upload_type == "file":
        f = filedialog.askopenfilename(filetypes=[("Resumes", "*.pdf;*.docx"), ("All files", "*.*")])
        if f: file_paths = [f]
    elif upload_type == "files":
        f = filedialog.askopenfilenames(filetypes=[("Resumes", "*.pdf;*.docx"), ("All files", "*.*")])
        if f: file_paths = list(f)
    elif upload_type == "folder":
        folder = filedialog.askdirectory()
        if folder:
            file_paths = [
                str(p) for p in Path(folder).rglob("*") 
                if p.suffix.lower() in [".pdf", ".docx"]
            ]
    
    if not file_paths:
        add_status("No files selected.")
        return

    copied_files_to_process = []
    try:
        add_status(f"Copying {len(file_paths)} files to workspace...")
        for fp in file_paths:
            dst = safe_copy_to_workspace(fp)
            copied_files_to_process.append(str(dst))
    except Exception as e:
        messagebox.showerror("File Copy Error", f"Failed to copy files to workspace: {e}")
        return

    check_dups = var_dup.get()
    STOP_EVENT = threading.Event()
    
    btn_single.config(state="disabled")
    btn_multi.config(state="disabled")
    btn_folder.config(state="disabled")
    btn_stop.pack(side="left", padx=(12, 6)) 
    
    PARSING_THREAD = threading.Thread(
        target=process_files_sequential,
        args=(copied_files_to_process, check_dups, STOP_EVENT),
        daemon=True
    )
    PARSING_THREAD.start()

def on_stop_parsing():
    """Sets the stop event for the parsing thread."""
    global STOP_EVENT
    if STOP_EVENT:
        add_status("Stopping... please wait for the current file to finish.")
        STOP_EVENT.set()
        btn_stop.config(state="disabled") 

def parsing_complete():
    """Called by the worker thread on the UI thread to re-enable UI."""
    global PARSING_THREAD, STOP_EVENT
    PARSING_THREAD = None
    STOP_EVENT = None
    
    btn_single.config(state="normal")
    btn_multi.config(state="normal")
    btn_folder.config(state="normal")
    
    btn_stop.pack_forget() 
    btn_stop.config(state="normal")

def export_sorted_files():
    """Asks for a destination and exports files by status."""
    if PARSING_THREAD and PARSING_THREAD.is_alive():
        messagebox.showwarning("Busy", "Please wait for the parsing job to finish before exporting.")
        return
        
    destination_folder = filedialog.askdirectory(title="Select Destination Folder for Exports")
    if not destination_folder:
        return

    add_status("Exporting sorted files...")

    def export_in_thread():
        try:
            # --- UPDATED: To handle dynamic statuses ---
            created_files = export_by_status(str(ACTIVE_EXCEL), destination_folder)
            
            if not created_files:
                ui_call(messagebox.showinfo, "Export Complete", "No candidates with a status were found to export.")
                ui_call(add_status, "Export complete. No files to create.")
                return

            # Create a summary message of all files created
            files_list_str = "\n".join(f"- {Path(f).name}" for f in created_files)
            
            ui_call(add_status, "Export complete!")
            ui_call(messagebox.showinfo, "Export Complete",
                    f"Files saved to {destination_folder}:\n"
                    f"{files_list_str}")
                    
        except Exception as e:
            log_error(f"Export failed: {e}\n{traceback.format_exc()}")
            ui_call(messagebox.showerror, "Export Error", f"Failed to export files: {e}")

    threading.Thread(target=export_in_thread, daemon=True).start()

# ----------------- REMOVED: File Watcher section -----------------
# (ExcelEventHandler class, start_file_watcher function)

def on_app_close():
    """Handle cleanup on app exit."""
    # --- SIMPLIFIED: No watcher to stop ---
    print("Closing application.")
    root.destroy()

# ----------------- Final initialization -----------------
# Set initial button states
btn_save_status.config(command=save_candidate_status)

# Ensure the active excel exists and load data
try:
    set_active_excel(ACTIVE_EXCEL) 
except Exception as e:
    log_error(f"Fatal error on startup trying to access Excel: {e}")
    messagebox.showerror("Fatal Error", f"Could not create or read active Excel file:\n{ACTIVE_EXCEL}\n\nError: {e}\n\nThe application may not function correctly.")

# --- REMOVED: start_file_watcher(ACTIVE_EXCEL) ---

# --- NEW: Handle app close gracefully ---
root.protocol("WM_DELETE_WINDOW", on_app_close)

add_status(f"Ready. Workspace: {WORKSPACE}")

# Start GUI
root.mainloop()
