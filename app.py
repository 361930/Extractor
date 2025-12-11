# app.py — GUI Application
import os
import sys
import threading
import traceback
import subprocess
import re  # Added for experience parsing
from pathlib import Path
from shutil import copy2
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import Set

# App modules
from utils import load_config, save_config, log_error, ensure_dirs, check_ollama_connection
from parser import parse_resume
from excel_handler import (
    validate_or_create_excel,
    read_all_rows, append_row, 
    update_status, export_by_status, DEFAULT_COLS
)
from db_handler import CandidateDB  # --- NEW IMPORT ---

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

# --- Initialize Database ---
DB = CandidateDB()

# --- Validate Active Excel Path ---
active_excel_path_str = CONFIG.get("active_excel", str(EXCEL_DIR / "resumes_data.xlsx"))
ACTIVE_EXCEL = Path(active_excel_path_str)

if not ACTIVE_EXCEL.exists() and ACTIVE_EXCEL.parent != EXCEL_DIR:
    ACTIVE_EXCEL = EXCEL_DIR / "resumes_data.xlsx"
    CONFIG["active_excel"] = str(ACTIVE_EXCEL)
    save_config(CONFIG)

DUPLICATE_CHECK_ENABLED = CONFIG.get("duplicate_check_enabled", True)

# --- Status Options ---
STATUS_OPTIONS = [
    "New Applicant", 
    "Re-Applicant (Updated)", 
    "Duplicate (Ignored)",
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
    nlp = None
except Exception as e:
    nlp = None
    MODEL_DISPLAY_NAME = "(Model: NOT LOADED)"
    log_error(f"Ollama connection error: {e}\n{traceback.format_exc()}")
    
    # We must show the error on startup
    dummy_root = tk.Tk()
    dummy_root.withdraw()
    messagebox.showerror("Fatal Model Error",
                         "Could not connect to the Ollama server.\n\n"
                         f"Error: {e}\n\n"
                         "Please ensure:\n"
                         "1. The Ollama application is running.\n"
                         "2. You have run: `ollama pull llama3:8b`")
    dummy_root.destroy()

# ----------------- Global Threading Control -----------------
PARSING_THREAD = None
STOP_EVENT = None

# ----------------- Main window -----------------
root = tk.Tk()
root.title(f"Resume Parser Pro {MODEL_DISPLAY_NAME}")
root.geometry("1250x750")

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

btn_multi = ttk.Button(frame_left_controls, text="Upload Multiple", command=lambda: on_upload("files"))
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
    text="Master Record Check (DB)", 
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
# --- UPDATED: Added 'Experience' column ---
tree_cols = ["S.No.", "Name", "Email", "Phone", "Experience", "Status"]
tree = ttk.Treeview(
    top_pane, 
    columns=tree_cols, 
    show="headings",
    displaycolumns=tree_cols 
)

tree.heading("S.No.", text="S.No.", anchor="w")
tree.column("S.No.", width=50, anchor="w")

tree.heading("Name", text="Name", anchor="w")
tree.column("Name", width=180, anchor="w")

tree.heading("Email", text="Email", anchor="w")
tree.column("Email", width=220, anchor="w")

tree.heading("Phone", text="Phone", anchor="w")
tree.column("Phone", width=120, anchor="w")

tree.heading("Experience", text="Exp (Yrs)", anchor="w") # New Column
tree.column("Experience", width=80, anchor="w")

tree.heading("Status", text="Status", anchor="w")
tree.column("Status", width=180, anchor="w")

tree.pack(fill="both", expand=True, padx=10, pady=(0, 10))

# --- Bottom Pane: Details Panel ---
bottom_pane = ttk.Frame(paned_window, height=250, style="Details.TFrame")
paned_window.add(bottom_pane, weight=1)

ttk.Label(bottom_pane, text="Candidate Details", style="Title.TLabel").pack(anchor="w", padx=12, pady=(10, 6))

details_frame = ttk.Frame(bottom_pane, style="Details.TFrame")
details_frame.pack(fill="x", expand=True, padx=12)

details_frame.columnconfigure(1, weight=1)

# --- Variables for Details Panel ---
detail_vars = {
    "S.No": tk.StringVar(value="-"),
    "Name": tk.StringVar(value="-"),
    "Email": tk.StringVar(value="-"),
    "Phone": tk.StringVar(value="-"),
    "Experience": tk.StringVar(value="-"), # New Variable
    "Status": tk.StringVar()
}

# Layout Details
row_idx = 0
for label_text, var in detail_vars.items():
    if label_text == "Status": continue
    ttk.Label(details_frame, text=f"{label_text}:", style="Details.TLabel", font=('Arial', 10, 'bold')).grid(row=row_idx, column=0, sticky="w", padx=5, pady=2)
    ttk.Label(details_frame, textvariable=var, style="Details.TLabel").grid(row=row_idx, column=1, sticky="w", padx=5, pady=2)
    row_idx += 1

# Status Dropdown
ttk.Label(details_frame, text="Status:", style="Details.TLabel", font=('Arial', 10, 'bold')).grid(row=row_idx, column=0, sticky="w", padx=5, pady=8)
status_combobox = ttk.Combobox(
    details_frame,
    textvariable=detail_vars["Status"],
    values=STATUS_OPTIONS, 
    state="readonly",
    width=25
)
status_combobox.grid(row=row_idx, column=1, sticky="w", padx=5, pady=8)
row_idx += 1

# Save Button
btn_save_status = ttk.Button(
    details_frame, 
    text="Save Status", 
    style="Details.TButton",
    command=lambda: save_candidate_status()
)
btn_save_status.grid(row=row_idx, column=1, sticky="w", padx=5, pady=10)

# Store the currently selected real S.No.
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
    if root.winfo_exists():
        root.after(0, lambda: fn(*a, **kw))

def add_status(text: str):
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
    global ACTIVE_EXCEL
    ACTIVE_EXCEL = path
    CONFIG["active_excel"] = str(ACTIVE_EXCEL)
    save_config(CONFIG)
    lbl_active.config(text=f"Active Excel: {ACTIVE_EXCEL.name}")
    validate_or_create_excel(str(ACTIVE_EXCEL)) 
    refresh_tree_from_excel()
    add_status(f"Active Excel set to {ACTIVE_EXCEL.name}")

def select_active_excel():
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
    CONFIG["duplicate_check_enabled"] = var_dup.get()
    save_config(CONFIG)
    add_status(f"Master Record Check {'enabled' if var_dup.get() else 'disabled'}.")

# ----------------- GUI Behavior Functions -----------------
def refresh_tree_from_excel():
    """Load rows from Excel into the tree."""
    for r in tree.get_children():
        tree.delete(r)
    
    rows = read_all_rows(str(ACTIVE_EXCEL))
    if not rows:
        return
        
    try:
        # Assuming DEFAULT_COLS are [S.No., Name, Email, Phone, Experience, Status]
        # We access by index to be safe
        pass 
    except ValueError as e:
        messagebox.showerror("Excel Error", f"Your Excel file is missing a required column: {e}")
        return

    # Add data to tree
    for i, r in enumerate(rows):
        # r structure: (SNo, Name, Email, Phone, Exp, Status)
        real_sno = r[0]
        if real_sno is None: continue
            
        real_sno_int = int(real_sno)
        
        # Ensure we have enough columns (handle old excel files without Experience)
        # Pad with empty string if Exp or Status missing
        r_list = list(r)
        while len(r_list) < 6:
            r_list.append("")
        
        values = (r_list[0], r_list[1], r_list[2], r_list[3], r_list[4], r_list[5])

        if tree.exists(real_sno_int):
            tree.item(real_sno_int, values=values)
            continue
            
        tree.insert("", "end", iid=real_sno_int, values=values)

def on_tree_select(event):
    """Fired when a user clicks a row. Populates the details panel."""
    global current_selected_sno
    
    selected_item = tree.focus() 
    if not selected_item: return
    
    try:
        current_selected_sno = int(selected_item) 
    except ValueError:
        current_selected_sno = None
        return 
    
    values = tree.item(selected_item, "values")
    # values = (sno, name, email, phone, experience, status)
    
    if len(values) >= 6:
        detail_vars["S.No"].set(values[0])
        detail_vars["Name"].set(values[1])
        detail_vars["Email"].set(values[2])
        detail_vars["Phone"].set(values[3])
        detail_vars["Experience"].set(values[4])
        detail_vars["Status"].set(values[5])

tree.bind("<<TreeviewSelect>>", on_tree_select)

def save_candidate_status():
    """Saves the status from the combobox to the Excel file."""
    if current_selected_sno is None:
        messagebox.showwarning("No Candidate", "Please select a candidate from the list.")
        return
        
    new_status = detail_vars["Status"].get()
    
    def save_in_thread():
        try:
            success = update_status(str(ACTIVE_EXCEL), current_selected_sno, new_status)
            if success:
                ui_call(add_status, f"Status updated for S.No. {current_selected_sno}")
                # Update tree manually
                current_values = list(tree.item(current_selected_sno, "values"))
                current_values[5] = new_status # Status is index 5
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
        # Search in Name, Email, Phone, Experience, Status
        combined = " ".join([str(v).lower() for v in values[1:]])
        
        if q == "" or q in combined:
            tree.reattach(item_id, "", "end")
        else:
            tree.detach(item_id)

entry_search.bind("<KeyRelease>", on_search_change)

# ----------------- Processing Worker (Updated with DB Logic) -----------------
def process_files_sequential(raw_paths, check_db, stop_event):
    """Background worker to parse files, check DB, save to Excel."""
    
    # --- 1. Copy Phase (Inside Thread) ---
    ui_call(progress.config, mode="indeterminate")
    add_status(f"Preparing {len(raw_paths)} files...")
    print(f"--- Starting Batch: {len(raw_paths)} files ---")
    
    processed_paths = []
    
    # Copy files
    for i, fp in enumerate(raw_paths):
        if stop_event.is_set(): break
        try:
            if i % 10 == 0:
                add_status(f"Copying file {i+1}/{len(raw_paths)}...")
            dst = safe_copy_to_workspace(fp)
            processed_paths.append(str(dst))
        except Exception as e:
            print(f"[Error] Failed to copy {fp}: {e}")
            log_error(f"Copy failed for {fp}: {e}")

    if stop_event.is_set():
        ui_call(progress.config, value=0)
        add_status("Stopped during file preparation.")
        ui_call(parsing_complete)
        return

    # --- 2. Parsing Phase ---
    ui_call(progress.config, mode="determinate", maximum=len(processed_paths), value=0)
    add_status(f"Starting analysis of {len(processed_paths)} resumes...")
    
    success_count = 0
    fail_count = 0
    update_count = 0
    dupe_count = 0
    stopped = False
    
    for i, fp in enumerate(processed_paths):
        if stop_event.is_set():
            stopped = True
            break 
            
        fname = Path(fp).name
        
        # --- VERBOSE CONSOLE OUTPUT ---
        print(f"[{i+1}/{len(processed_paths)}] Analyzing: {fname}")
        sys.stdout.flush() # Force CMD update
        
        try:
            add_status(f"Parsing ({i+1}/{len(processed_paths)}): {fname}...")
            
            # 1. PARSE
            data = parse_resume(fp) 
            
            if not data:
                print(f"  -> FAILED to extract data from {fname}")
                add_status(f"Failed: {fname}. See logs.")
                fail_count += 1
                continue

            print(f"  -> Extracted: {data.get('Name', 'Unknown')} | {data.get('Email', 'No Email')}")

            # 2. EXTRACT & NORMALIZE
            email = data.get("Email", "").strip().lower()
            if isinstance(email, list): email = email[0] if email else ""
            
            # Parse Experience (clean up "5 years" -> 5.0)
            new_exp_str = str(data.get("Experience", "0")).lower().replace("years", "").strip()
            try:
                # Find first float/int in string
                matches = re.findall(r"[\d\.]+", new_exp_str)
                new_exp_val = float(matches[0]) if matches else 0.0
            except:
                new_exp_val = 0.0

            status_to_save = "New Applicant"

            # 3. DB CHECK (MASTER RECORD)
            if check_db and email:
                existing_candidate = DB.get_candidate(email)
                
                if existing_candidate:
                    # Compare Experience
                    old_exp_str = str(existing_candidate["experience"])
                    try:
                        matches = re.findall(r"[\d\.]+", old_exp_str)
                        old_exp_val = float(matches[0]) if matches else 0.0
                    except:
                        old_exp_val = 0.0
                    
                    # Logic: If new exp > old exp + 0.5 (tolerance), it's an update
                    if new_exp_val > (old_exp_val + 0.5):
                        status_to_save = "Re-Applicant (Updated)"
                        print(f"  -> Existing candidate found. New exp ({new_exp_val}) > Old ({old_exp_val}). Updating.")
                        update_count += 1
                        
                        # Update DB
                        data['email'] = email
                        data['experience'] = str(new_exp_val)
                        data['name'] = data.get('Name')
                        data['phone'] = data.get('Phone')
                        data['resume_path'] = data.get('ResumePath')
                        DB.upsert_candidate(data, is_update=True)
                    else:
                        status_to_save = "Duplicate (Ignored)"
                        print(f"  -> Duplicate found (Exp {new_exp_val} vs {old_exp_val}). Marking as duplicate.")
                        dupe_count += 1
                else:
                    # Insert New into DB
                    data['email'] = email
                    data['experience'] = str(new_exp_val)
                    data['name'] = data.get('Name')
                    data['phone'] = data.get('Phone')
                    data['resume_path'] = data.get('ResumePath')
                    DB.upsert_candidate(data, is_update=False)
                    success_count += 1
            else:
                success_count += 1
            
            # 4. APPEND TO EXCEL
            # We add all records to Excel so user can see duplicates if they want.
            # Use status to filter them out later if needed.
            new_sno = append_row(str(ACTIVE_EXCEL), data, status=status_to_save)
            print(f"  -> Saved to Excel (S.No {new_sno})")
            
            # 5. UPDATE UI
            display_sno = len(tree.get_children()) + 1
            ui_call(
                tree.insert,
                "",
                "end",
                iid=new_sno, 
                values=(
                    display_sno,
                    data.get("Name", ""),
                    data.get("Email", ""),
                    data.get("Phone", ""),
                    data.get("Experience", "0"),
                    status_to_save
                )
            )

        except Exception as e:
            err_msg = f"Processing error for {fp}: {e}"
            print(f"  -> EXCEPTION: {err_msg}")
            log_error(err_msg + "\n" + traceback.format_exc())
            fail_count += 1
        finally:
            ui_call(progress.step, 1)

    ui_call(progress.config, value=0)
    
    if stopped:
        summary = "Parsing stopped by user."
    else:
        summary = (f"Batch Complete.\n"
                   f"New: {success_count}\n"
                   f"Updated (Re-applicants): {update_count}\n"
                   f"Duplicates (Ignored): {dupe_count}\n"
                   f"Failed: {fail_count}")
    
    add_status("Batch processing finished.")
    ui_call(messagebox.showinfo, "Result", summary)
    ui_call(parsing_complete)


def safe_copy_to_workspace(src_path: str):
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

    # --- CHANGED: Don't copy here. Pass raw paths to thread. ---
    
    check_db = var_dup.get()
    STOP_EVENT = threading.Event()
    
    btn_single.config(state="disabled")
    btn_multi.config(state="disabled")
    btn_folder.config(state="disabled")
    btn_stop.pack(side="left", padx=(12, 6)) 
    
    PARSING_THREAD = threading.Thread(
        target=process_files_sequential,
        args=(file_paths, check_db, STOP_EVENT), # Pass RAW paths
        daemon=True
    )
    PARSING_THREAD.start()

def on_stop_parsing():
    global STOP_EVENT
    if STOP_EVENT:
        add_status("Stopping... please wait for the current file.")
        STOP_EVENT.set()
        btn_stop.config(state="disabled") 

def parsing_complete():
    global PARSING_THREAD, STOP_EVENT
    PARSING_THREAD = None
    STOP_EVENT = None
    
    btn_single.config(state="normal")
    btn_multi.config(state="normal")
    btn_folder.config(state="normal")
    
    btn_stop.pack_forget() 
    btn_stop.config(state="normal")

def export_sorted_files():
    if PARSING_THREAD and PARSING_THREAD.is_alive():
        messagebox.showwarning("Busy", "Please wait for parsing to finish.")
        return
        
    destination_folder = filedialog.askdirectory(title="Select Destination for Exports")
    if not destination_folder:
        return

    add_status("Exporting sorted files...")

    def export_in_thread():
        try:
            created_files = export_by_status(str(ACTIVE_EXCEL), destination_folder)
            
            if not created_files:
                ui_call(messagebox.showinfo, "Export Complete", "No candidates with status found.")
                ui_call(add_status, "Export complete. No files created.")
                return

            files_list_str = "\n".join(f"- {Path(f).name}" for f in created_files)
            
            ui_call(add_status, "Export complete!")
            ui_call(messagebox.showinfo, "Export Complete",
                    f"Files saved to {destination_folder}:\n"
                    f"{files_list_str}")
                    
        except Exception as e:
            log_error(f"Export failed: {e}\n{traceback.format_exc()}")
            ui_call(messagebox.showerror, "Export Error", f"Failed to export files: {e}")

    threading.Thread(target=export_in_thread, daemon=True).start()

def on_app_close():
    root.destroy()

# ----------------- Final initialization -----------------
# Ensure the active excel exists and load data
try:
    set_active_excel(ACTIVE_EXCEL) 
except Exception as e:
    log_error(f"Fatal startup error: {e}")
    messagebox.showerror("Fatal Error", f"Could not access Excel:\n{e}")

root.protocol("WM_DELETE_WINDOW", on_app_close)
add_status(f"Ready. Workspace: {WORKSPACE}")

# Start GUI
root.mainloop()
