# app.py — GUI Application
import os
import sys
import threading
import traceback
import subprocess
import re
import datetime
from pathlib import Path
from shutil import copy2
import tkinter as tk
from tkinter import filedialog, messagebox
import tkinter.ttk as ttk # Standard ttk for stability

# --- Modern UI Library ---
import ttkbootstrap as tb
from ttkbootstrap.constants import *
from ttkbootstrap.dialogs import Messagebox
try:
    from ttkbootstrap.widgets import ToastNotification
except ImportError:
    from ttkbootstrap.toast import ToastNotification

# App modules
from utils import load_config, save_config, log_error, ensure_dirs, check_ollama_connection
from parser import parse_resume
from excel_handler import (
    validate_or_create_excel,
    read_all_rows, append_row, 
    update_status, export_by_status, DEFAULT_COLS
)
from db_handler import CandidateDB

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

# --- Status Dropdown Options ---
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
    success, model_or_error = check_ollama_connection()
    if not success:
        raise OSError(model_or_error)
    MODEL_DISPLAY_NAME = f"Model: {model_or_error}"
    nlp = None
except Exception as e:
    nlp = None
    MODEL_DISPLAY_NAME = "Model: NOT LOADED"
    log_error(f"Ollama connection error: {e}")

# ----------------- Globals -----------------
PARSING_THREAD = None
STOP_EVENT = None
CONFLICTS = []  # Store duplicate candidates for review
THEME_NAME = "cosmo" # Default light theme

# Cache for Treeview Data (Search Optimization)
# Format: List of tuples/lists matching tree columns
TREE_DATA_CACHE = [] 

# ----------------- Duplicate Resolver Window -----------------
class DuplicateResolver(tb.Toplevel):
    """
    A specific window to handle conflicts found during parsing.
    Blocks the main window until resolved.
    """
    def __init__(self, parent, conflicts, db, excel_path, tree_ref):
        super().__init__(master=parent, title="Resolve Duplicates", size=(1100, 750))
        
        self.conflicts = conflicts
        self.db = db
        self.excel_path = str(excel_path)
        self.tree = tree_ref
        self.resolved_indices = set()
        self.current_index = 0
        
        # Handle Window Close (X button) safely
        self.protocol("WM_DELETE_WINDOW", self.on_close_window)
        
        # Make this window modal
        self.transient(parent)
        self.grab_set()
        
        self.create_widgets()
        self.load_conflict(0)
        self.place_window_center()
        self.focus_set()

    def create_widgets(self):
        main_container = tb.Frame(self, padding=20)
        main_container.pack(fill=BOTH, expand=True)

        # Info Banner
        info_frame = tb.Frame(main_container, bootstyle="info", padding=10)
        info_frame.pack(fill=X, pady=(0, 10))
        # Fixed: Removed inverse=True, used bootstyle="inverse-info"
        tb.Label(info_frame, text="ℹ Duplicates detected in Master Database. Please review below.", bootstyle="inverse-info", font=('Segoe UI', 10)).pack(side=LEFT)

        # Layout: Paned Window
        paned = ttk.PanedWindow(main_container, orient=HORIZONTAL)
        paned.pack(fill=BOTH, expand=True)
        
        # --- Left Panel: List ---
        frame_list = tb.Frame(paned, padding=(0, 0, 10, 0))
        paned.add(frame_list, weight=1)
        
        tb.Label(frame_list, text=f"Conflicts ({len(self.conflicts)})", font=('Helvetica', 12, 'bold'), bootstyle="primary").pack(anchor="w", pady=(0, 10))
        
        list_scroll = tb.Scrollbar(frame_list)
        list_scroll.pack(side=RIGHT, fill=Y)
        
        self.listbox = tk.Listbox(frame_list, font=('Segoe UI', 10), yscrollcommand=list_scroll.set, activestyle="none", selectbackground="#007bff", selectforeground="white", borderwidth=1, relief="solid")
        self.listbox.pack(fill=BOTH, expand=True)
        self.listbox.bind('<<ListboxSelect>>', self.on_select)
        list_scroll.config(command=self.listbox.yview)
        
        for c in self.conflicts:
            name = c['new'].get('Name', 'Unknown')
            email = c['new'].get('Email', 'No Email')
            if not email and c.get('inferred_from_name'):
                email = "[Name Match]"
            self.listbox.insert(tk.END, f"{name}")
            
        # --- Right Panel: Comparison ---
        frame_comp = tb.Frame(paned, padding=(10, 0, 0, 0))
        paned.add(frame_comp, weight=3)
        
        header_frame = tb.Frame(frame_comp)
        header_frame.pack(fill=X, pady=(0, 20))
        tb.Label(header_frame, text="Comparison", font=('Helvetica', 16, 'bold')).pack(side=LEFT)
        self.lbl_status = tb.Label(header_frame, text="Action Required", bootstyle="warning", font=('Helvetica', 10, 'bold'))
        self.lbl_status.pack(side=RIGHT)

        comp_container = tb.Frame(frame_comp)
        comp_container.pack(fill=BOTH, expand=True)
        comp_container.columnconfigure(0, weight=1)
        comp_container.columnconfigure(1, weight=4)
        comp_container.columnconfigure(2, weight=4)
        
        tb.Label(comp_container, text="FIELD", font=('Helvetica', 10, 'bold'), bootstyle="secondary").grid(row=0, column=0, sticky="w", pady=10)
        
        old_header = ttk.LabelFrame(comp_container, text=" MASTER DB (Existing) ", padding=10)
        old_header.grid(row=0, column=1, sticky="ew", padx=10)
        tb.Label(old_header, text="Data currently in database", font=('Helvetica', 8), bootstyle="secondary").pack(anchor="w")

        new_header = ttk.LabelFrame(comp_container, text=" RESUME (New) ", padding=10)
        new_header.grid(row=0, column=2, sticky="ew", padx=10)
        tb.Label(new_header, text="Data parsed from file", font=('Helvetica', 8), bootstyle="secondary").pack(anchor="w")

        self.labels = {}
        fields = ["Name", "Email", "Phone", "Experience", "Last Applied"]
        
        for i, f in enumerate(fields):
            row = i + 1
            tb.Label(comp_container, text=f.upper(), font=('Helvetica', 9, 'bold'), bootstyle="secondary").grid(row=row, column=0, sticky="w", pady=10)
            
            self.labels[f"old_{f}"] = tb.Label(comp_container, text="-", font=('Segoe UI', 10), bootstyle="danger")
            self.labels[f"old_{f}"].grid(row=row, column=1, sticky="w", padx=15, pady=5)
            
            self.labels[f"new_{f}"] = tb.Label(comp_container, text="-", font=('Segoe UI', 10, 'bold'), bootstyle="success")
            self.labels[f"new_{f}"].grid(row=row, column=2, sticky="w", padx=15, pady=5)
            
            if i < len(fields) - 1:
                tb.Separator(comp_container, bootstyle="secondary-alpha").grid(row=row, column=0, columnspan=3, sticky="sew")

        # Buttons
        btn_frame = tb.Frame(frame_comp, padding=(0, 20, 0, 0))
        btn_frame.pack(fill=X, side=BOTTOM)
        
        self.btn_keep = tb.Button(btn_frame, text="Keep Master (Mark as Duplicate)", bootstyle="danger-outline", command=self.keep_old)
        self.btn_keep.pack(side=LEFT, expand=True, fill=X, padx=(0, 10))
        
        self.btn_update = tb.Button(btn_frame, text="Update Master (Use New Data)", bootstyle="success", command=self.replace_new)
        self.btn_update.pack(side=RIGHT, expand=True, fill=X, padx=(10, 0))

    def on_select(self, event):
        sel = self.listbox.curselection()
        if not sel: return
        self.load_conflict(sel[0])

    def load_conflict(self, index):
        self.current_index = index
        c = self.conflicts[index]
        old = c['old']
        new = c['new']
        
        self.labels["old_Name"].config(text=old.get('name', ''))
        self.labels["old_Email"].config(text=old.get('email', ''))
        self.labels["old_Phone"].config(text=old.get('phone', ''))
        self.labels["old_Experience"].config(text=f"{old.get('experience', '')} yrs")
        self.labels["old_Last Applied"].config(text=str(old.get('last_applied_date', '')))
        
        self.labels["new_Name"].config(text=new.get('Name', ''))
        self.labels["new_Email"].config(text=new.get('Email', ''))
        self.labels["new_Phone"].config(text=new.get('Phone', ''))
        self.labels["new_Experience"].config(text=f"{new.get('Experience', '')} yrs")
        self.labels["new_Last Applied"].config(text="Now")
        
        if index in self.resolved_indices:
            self.btn_keep.state(['disabled'])
            self.btn_update.state(['disabled'])
            self.lbl_status.config(text="RESOLVED", bootstyle="secondary")
        else:
            self.btn_keep.state(['!disabled'])
            self.btn_update.state(['!disabled'])
            self.lbl_status.config(text="ACTION REQUIRED", bootstyle="warning")

    def _finalize_action(self, status_msg, excel_status, update_db=False):
        c = self.conflicts[self.current_index]
        data = c['new']
        
        try:
            if update_db:
                # Clean exp
                try:
                    exp_str = str(data.get("Experience", "0")).lower().replace("years", "").strip()
                    matches = re.findall(r"[\d\.]+", exp_str)
                    exp_val = matches[0] if matches else "0"
                except: exp_val = "0"

                target_email = data.get('Email')
                if not target_email and c.get('old', {}).get('email'):
                    target_email = c['old']['email']
                    data['Email'] = target_email

                db_data = {
                    'email': target_email,
                    'name': data.get('Name'),
                    'phone': data.get('Phone'),
                    'experience': str(exp_val),
                    'resume_path': data.get('ResumePath')
                }
                self.db.upsert_candidate(db_data, is_update=True)
                
            new_sno = append_row(self.excel_path, data, status=excel_status)
            
            # --- UPDATE MAIN UI CACHE & TREE ---
            row_data = (
                new_sno, data.get("Name"), data.get("Email"), data.get("Phone"),
                data.get("Experience"), excel_status
            )
            # Add to global cache
            TREE_DATA_CACHE.append(row_data)
            
            # Check if it passes current search filter
            q = search_var.get().strip().lower()
            row_str = " ".join([str(x).lower() for x in row_data[1:]])
            
            if q == "" or q in row_str:
                self.tree.insert("", "end", iid=new_sno, values=row_data)
            # -----------------------------------
            
            self.listbox.itemconfig(self.current_index, {'bg': '#f0f0f0', 'fg': '#aaa'})
            self.resolved_indices.add(self.current_index)
            self.lbl_status.config(text=status_msg, bootstyle="success")
            self.btn_keep.state(['disabled'])
            self.btn_update.state(['disabled'])
            
            if self.current_index < len(self.conflicts) - 1:
                next_idx = self.current_index + 1
                self.listbox.selection_clear(0, tk.END)
                self.listbox.selection_set(next_idx)
                self.listbox.activate(next_idx)
                self.after(300, lambda: self.load_conflict(next_idx))
            elif len(self.resolved_indices) == len(self.conflicts):
                Messagebox.show_info("All duplicates resolved.", "Complete", parent=self)
                self.destroy()
                
        except Exception as e:
            Messagebox.show_error(f"Error: {e}", "Action Failed", parent=self)
            log_error(f"Resolution error: {e}")

    def keep_old(self):
        self._finalize_action("IGNORED", "Duplicate (Ignored)", update_db=False)

    def replace_new(self):
        self._finalize_action("UPDATED", "Re-Applicant (Updated)", update_db=True)

    def on_close_window(self):
        """Handle case where user closes window without resolving all."""
        unresolved = len(self.conflicts) - len(self.resolved_indices)
        if unresolved > 0:
            if Messagebox.show_question(f"{unresolved} items are unresolved. Close anyway?", "Confirm", buttons=['No:secondary', 'Yes:danger'], parent=self) == 'Yes':
                self.destroy()
        else:
            self.destroy()


# ----------------- Main GUI Application -----------------

root = tb.Window(themename=THEME_NAME)
root.title("Resume Parser Pro")
root.geometry("1280x850")
root.place_window_center()

# UI Vars
detail_sno_var = tk.StringVar(value="-")
detail_name_var = tk.StringVar(value="-")
detail_email_var = tk.StringVar(value="-")
detail_phone_var = tk.StringVar(value="-")
detail_exp_var = tk.StringVar(value="-")
detail_status_var = tk.StringVar()
search_var = tk.StringVar()

# --- UI Construction ---

header_frame = tb.Frame(root, padding=(20, 10))
header_frame.pack(fill=X)

title_lbl = tb.Label(header_frame, text="Resume Parser Pro", font=('Helvetica', 18, 'bold'), bootstyle="primary")
title_lbl.pack(side=LEFT)

subtitle_lbl = tb.Label(header_frame, text=f" | {MODEL_DISPLAY_NAME}", font=('Helvetica', 10), bootstyle="secondary")
subtitle_lbl.pack(side=LEFT, padx=10, pady=(6,0))

def toggle_theme():
    global THEME_NAME
    current = root.style.theme.name
    new_theme = "superhero" if current == "cosmo" else "cosmo"
    root.style.theme_use(new_theme)
    THEME_NAME = new_theme
    
btn_theme = tb.Checkbutton(header_frame, text="Dark Mode", bootstyle="round-toggle", command=toggle_theme)
btn_theme.pack(side=RIGHT)

paned_window = ttk.PanedWindow(root, orient=VERTICAL)
paned_window.pack(fill=BOTH, expand=True, padx=10, pady=10)

top_pane = tb.Frame(paned_window)
paned_window.add(top_pane, weight=4) 

controls_frame = tb.Frame(top_pane, padding=(10, 10))
controls_frame.pack(fill=X)

btn_single = tb.Button(controls_frame, text="Upload Resume", bootstyle="primary", command=lambda: on_upload("file"))
btn_single.pack(side=LEFT, padx=5)

btn_folder = tb.Button(controls_frame, text="Upload Folder", bootstyle="primary-outline", command=lambda: on_upload("folder"))
btn_folder.pack(side=LEFT, padx=5)

var_dup = tk.BooleanVar(value=DUPLICATE_CHECK_ENABLED)
chk_dup = tb.Checkbutton(controls_frame, text="Master DB Check", variable=var_dup, bootstyle="success-round-toggle", command=lambda: save_dup_config())
chk_dup.pack(side=LEFT, padx=20)

btn_stop = tb.Button(controls_frame, text="STOP", bootstyle="danger", command=lambda: on_stop_parsing())

btn_export = tb.Button(controls_frame, text="Export Sorted", bootstyle="success-outline", command=lambda: export_sorted_files())
btn_export.pack(side=RIGHT, padx=5)

btn_open_excel = tb.Button(controls_frame, text="Open Excel", bootstyle="info-outline", command=lambda: open_in_explorer(ACTIVE_EXCEL))
btn_open_excel.pack(side=RIGHT, padx=5)

sec_bar = tb.Frame(top_pane, padding=(10, 5))
sec_bar.pack(fill=X)

lbl_active = tb.Label(sec_bar, text=f"Active: {ACTIVE_EXCEL.name}", font=('Segoe UI', 9), bootstyle="secondary")
lbl_active.pack(side=LEFT, padx=5)
# Renamed button for clarity
tb.Button(sec_bar, text="Create / Open Database", bootstyle="link", command=lambda: select_active_excel()).pack(side=LEFT)

tb.Label(sec_bar, text="Search:", bootstyle="secondary").pack(side=RIGHT, padx=(10, 5))
entry_search = tb.Entry(sec_bar, textvariable=search_var, width=30)
entry_search.pack(side=RIGHT)

tree_frame = tb.Frame(top_pane, padding=10)
tree_frame.pack(fill=BOTH, expand=True)

tree_cols = ["S.No.", "Name", "Email", "Phone", "Experience", "Status"]
tree = tb.Treeview(tree_frame, columns=tree_cols, show="headings", bootstyle="primary")

cols_cfg = {
    "S.No.": 60, "Name": 200, "Email": 250, 
    "Phone": 120, "Experience": 80, "Status": 150
}
for c, w in cols_cfg.items():
    tree.heading(c, text=c)
    tree.column(c, width=w)

tree_scroll = tb.Scrollbar(tree_frame, orient=VERTICAL, command=tree.yview)
tree.configure(yscrollcommand=tree_scroll.set)
tree_scroll.pack(side=RIGHT, fill=Y)
tree.pack(fill=BOTH, expand=True)

bottom_pane = tb.Frame(paned_window, padding=10)
paned_window.add(bottom_pane, weight=1)

details_card = ttk.LabelFrame(bottom_pane, text=" Candidate Details ", padding=15)
details_card.pack(fill=BOTH, expand=True)

details_card.columnconfigure(1, weight=1)
details_card.columnconfigure(3, weight=1)
details_card.columnconfigure(5, weight=1)

tb.Label(details_card, text="Name:", font=('Segoe UI', 9, 'bold')).grid(row=0, column=0, sticky="e", padx=10, pady=5)
tb.Label(details_card, textvariable=detail_name_var, bootstyle="primary").grid(row=0, column=1, sticky="w")

tb.Label(details_card, text="Email:", font=('Segoe UI', 9, 'bold')).grid(row=0, column=2, sticky="e", padx=10, pady=5)
tb.Label(details_card, textvariable=detail_email_var).grid(row=0, column=3, sticky="w")

tb.Label(details_card, text="Phone:", font=('Segoe UI', 9, 'bold')).grid(row=0, column=4, sticky="e", padx=10, pady=5)
tb.Label(details_card, textvariable=detail_phone_var).grid(row=0, column=5, sticky="w")

tb.Label(details_card, text="Experience:", font=('Segoe UI', 9, 'bold')).grid(row=1, column=0, sticky="e", padx=10, pady=5)
tb.Label(details_card, textvariable=detail_exp_var).grid(row=1, column=1, sticky="w")

tb.Label(details_card, text="Status:", font=('Segoe UI', 9, 'bold')).grid(row=1, column=2, sticky="e", padx=10, pady=5)
combo_status = tb.Combobox(details_card, textvariable=detail_status_var, values=STATUS_OPTIONS, state="readonly", width=25)
combo_status.grid(row=1, column=3, sticky="w")

btn_save = tb.Button(details_card, text="Update Status", bootstyle="warning-outline", command=lambda: save_candidate_status())
btn_save.grid(row=1, column=4, padx=10, sticky="w")

status_bar = tb.Frame(root, bootstyle="light", padding=(10, 2))
status_bar.pack(fill=X, side=BOTTOM)

lbl_status = tb.Label(status_bar, text="Ready", font=('Segoe UI', 9), bootstyle="secondary")
lbl_status.pack(side=LEFT)

progress = tb.Progressbar(status_bar, orient=HORIZONTAL, mode='determinate', bootstyle="success-striped", length=200)
progress.pack(side=RIGHT)

current_selected_sno = None


# ----------------- Functions -----------------
def ui_call(fn, *a, **kw):
    if root.winfo_exists(): root.after(0, lambda: fn(*a, **kw))

def add_status(text): ui_call(lbl_status.config, text=text)

def open_in_explorer(path):
    if os.path.exists(path):
        os.startfile(path) if os.name == 'nt' else subprocess.Popen(['open', str(path)])
    else:
        Messagebox.show_error(f"Cannot open {path}", "File Not Found")

def set_active_excel(path):
    global ACTIVE_EXCEL
    ACTIVE_EXCEL = path
    CONFIG["active_excel"] = str(ACTIVE_EXCEL)
    save_config(CONFIG)
    lbl_active.config(text=f"Active: {ACTIVE_EXCEL.name}")
    validate_or_create_excel(str(ACTIVE_EXCEL))
    refresh_tree_from_excel()
    ToastNotification(title="Database Loaded", message=f"Ready to save to: {ACTIVE_EXCEL.name}", duration=3000).show_toast()

def select_active_excel():
    f = filedialog.asksaveasfilename(
        initialdir=str(EXCEL_DIR), 
        title="Create or Select Excel Database", 
        defaultextension=".xlsx", 
        filetypes=[("Excel Files", "*.xlsx")]
    )
    if f: set_active_excel(Path(f))

def save_dup_config():
    CONFIG["duplicate_check_enabled"] = var_dup.get()
    save_config(CONFIG)
    add_status(f"Master Record Check: {'Enabled' if var_dup.get() else 'Disabled'}")

def refresh_tree_from_excel():
    """Reloads Excel data into Global Cache and repopulates Tree."""
    global TREE_DATA_CACHE
    TREE_DATA_CACHE = [] # Clear cache
    
    # Clear visual tree
    for r in tree.get_children(): tree.delete(r)
    
    rows = read_all_rows(str(ACTIVE_EXCEL))
    for r in rows:
        r_list = list(r)
        while len(r_list) < 6: r_list.append("")
        
        if r_list[0]: # Ensure Valid S.No
            try:
                # Add to Cache
                TREE_DATA_CACHE.append(r_list[:6])
                # Add to Tree (initially showing all)
                tree.insert("", "end", iid=int(r_list[0]), values=r_list[:6])
            except: pass

def on_tree_select(event):
    global current_selected_sno
    sel = tree.focus()
    if not sel: return
    try: current_selected_sno = int(sel)
    except: return
    v = tree.item(sel, "values")
    if len(v) >= 6:
        detail_name_var.set(v[1])
        detail_email_var.set(v[2])
        detail_phone_var.set(v[3])
        detail_exp_var.set(v[4])
        detail_status_var.set(v[5])

tree.bind("<<TreeviewSelect>>", on_tree_select)

def save_candidate_status():
    if not current_selected_sno:
        Messagebox.show_warning("Please select a candidate first.", "Selection Required")
        return
    s = detail_status_var.get()
    if update_status(str(ACTIVE_EXCEL), current_selected_sno, s):
        v = list(tree.item(current_selected_sno, "values"))
        v[5] = s
        tree.item(current_selected_sno, values=tuple(v))
        
        # Update Cache too!
        for i, row in enumerate(TREE_DATA_CACHE):
            if int(row[0]) == current_selected_sno:
                new_row = list(row)
                new_row[5] = s
                TREE_DATA_CACHE[i] = tuple(new_row)
                break
                
        add_status(f"Status saved for #{current_selected_sno}")
    else:
        Messagebox.show_error("Could not save status to Excel.", "Error")

def on_search_change(*_):
    """Refilters tree based on cache."""
    q = search_var.get().strip().lower()
    
    # 1. Clear visible tree
    for k in tree.get_children():
        tree.detach(k) # Detach is faster than delete for reloading? 
                       # Actually, delete is safer to clear view.
        tree.delete(k)

    # 2. Re-insert matches from cache
    for row in TREE_DATA_CACHE:
        row_text = " ".join([str(x).lower() for x in row[1:]]) # search all fields except S.No
        if q == "" or q in row_text:
            try:
                tree.insert("", "end", iid=int(row[0]), values=row)
            except tk.TclError: pass 

entry_search.bind("<KeyRelease>", on_search_change)

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

# ----------------- Processing Logic -----------------
def process_files_sequential(raw_paths, check_db, stop_event):
    global CONFLICTS
    CONFLICTS = [] 
    
    ui_call(progress.config, mode="indeterminate")
    add_status(f"Preparing {len(raw_paths)} files...")
    sys.stdout.flush()
    
    processed_paths = []
    stopped_during_copy = False
    
    for i, fp in enumerate(raw_paths):
        if stop_event.is_set(): 
            stopped_during_copy = True
            break
        try:
            if i % 10 == 0: add_status(f"Copying file {i+1}...")
            dst = safe_copy_to_workspace(fp)
            processed_paths.append(str(dst))
        except Exception as e:
            print(f"[Error] Copy failed for {fp}: {e}")
            log_error(f"Copy error: {e}")

    if stopped_during_copy:
        ui_call(parsing_complete, 0, 0, 0, 0, True)
        return

    ui_call(progress.config, mode="determinate", maximum=len(processed_paths), value=0)
    add_status(f"Parsing {len(processed_paths)} resumes...")
    
    cnt_new, cnt_fail, cnt_dup = 0, 0, 0
    stopped_during_parse = False
    
    for i, fp in enumerate(processed_paths):
        if stop_event.is_set(): 
            stopped_during_parse = True
            break
        
        fname = Path(fp).name
        print(f"[{i+1}/{len(processed_paths)}] Parsing: {fname}")
        sys.stdout.flush()
        
        try:
            add_status(f"Processing ({i+1}): {fname}")
            data = parse_resume(fp)
            if not data:
                print(f"  -> Failed to parse {fname}")
                cnt_fail += 1
                ui_call(progress.step, 1)
                continue

            email = data.get("Email", "").strip().lower()
            if isinstance(email, list): email = email[0] if email else ""
            
            try:
                exp_str = str(data.get("Experience", "0")).lower().replace("years", "").strip()
                matches = re.findall(r"[\d\.]+", exp_str)
                exp_val = matches[0] if matches else "0"
            except: exp_val = "0"
            data['Experience'] = exp_val 

            is_conflict = False
            existing_record = None
            
            if check_db:
                # 1. Check by Email
                if email:
                    existing_record = DB.get_candidate(email)
                
                # 2. Check by Name if not found by email or email missing
                if not existing_record and data.get("Name"):
                    name_matches = DB.get_candidates_by_name(data.get("Name"))
                    if name_matches:
                        existing_record = name_matches[0]
                        data['inferred_from_name'] = True
                        print(f"  -> Match by Name found: {existing_record['email']}")

                if existing_record:
                    is_conflict = True
                    print(f"  -> DUPLICATE FOUND. Staging.")
                    CONFLICTS.append({"new": data, "old": existing_record, "file": fp})
                    cnt_dup += 1
            
            if not is_conflict:
                print(f"  -> New Candidate. Saving.")
                save_email = email if email else f"no_email_{os.urandom(4).hex()}"
                
                db_data = {
                    'email': save_email, 'name': data.get('Name'), 'phone': data.get('Phone'),
                    'experience': str(exp_val), 'resume_path': data.get('ResumePath')
                }
                if check_db: DB.upsert_candidate(db_data, is_update=False)
                new_sno = append_row(str(ACTIVE_EXCEL), data, status="New Applicant")
                
                # UPDATE UI & CACHE
                row_data = (
                    new_sno, data.get("Name"), data.get("Email"), data.get("Phone"),
                    data.get("Experience"), "New Applicant"
                )
                TREE_DATA_CACHE.append(row_data)
                
                # Update visible tree if matches search
                q_curr = search_var.get().strip().lower()
                row_str = " ".join([str(x).lower() for x in row_data[1:]])
                if q_curr == "" or q_curr in row_str:
                    ui_call(tree.insert, "", "end", iid=new_sno, values=row_data)
                
                cnt_new += 1
                
        except Exception as e:
            print(f"  -> ERROR processing {fname}: {e}")
            cnt_fail += 1
        finally:
            ui_call(progress.step, 1)

    ui_call(progress.config, value=0)
    ui_call(parsing_complete, cnt_new, cnt_dup, cnt_fail, 0, stopped_during_parse)

def parsing_complete(new, dup, fail, up, stopped):
    global PARSING_THREAD, STOP_EVENT
    PARSING_THREAD = None
    STOP_EVENT = None
    
    btn_single.config(state="normal")
    btn_folder.config(state="normal")
    btn_stop.pack_forget()
    
    add_status("Batch Processing Complete.")
    
    summary = f"New Candidates: {new}\nParsing Failures: {fail}\nDuplicates: {dup}"
    
    if stopped:
        summary = "Processing stopped by user.\n\n" + summary
    
    if dup > 0:
        try:
            Messagebox.show_info(f"{summary}\n\nDuplicates found. Opening Resolver.", "Results")
            DuplicateResolver(root, CONFLICTS, DB, ACTIVE_EXCEL, tree)
        except Exception as e:
            messagebox.showerror("UI Error", f"Could not open resolver: {e}")
            log_error(f"Resolver crash: {e}")
    else:
        Messagebox.show_info(summary, "Results")

def on_upload(mode):
    global PARSING_THREAD, STOP_EVENT
    if PARSING_THREAD: 
        Messagebox.show_warning("Parsing is already in progress.", "Busy")
        return
    
    paths = []
    if mode == "file":
        f = filedialog.askopenfilename(filetypes=[("Resumes", "*.pdf;*.docx")])
        if f: paths=[f]
    elif mode == "folder":
        d = filedialog.askdirectory()
        if d: paths=[str(p) for p in Path(d).rglob("*") if p.suffix.lower() in ['.pdf', '.docx']]
    
    if not paths: return
    
    STOP_EVENT = threading.Event()
    btn_single.config(state="disabled")
    btn_folder.config(state="disabled")
    btn_stop.pack(side=LEFT, padx=10)
    
    PARSING_THREAD = threading.Thread(
        target=process_files_sequential,
        args=(paths, var_dup.get(), STOP_EVENT),
        daemon=True
    )
    PARSING_THREAD.start()

def on_stop_parsing():
    if STOP_EVENT: 
        STOP_EVENT.set()
        add_status("Stopping... please wait.")

def export_sorted_files():
    d = filedialog.askdirectory()
    if d:
        files = export_by_status(str(ACTIVE_EXCEL), d)
        if files: Messagebox.show_info(f"Exported {len(files)} files.", "Success")
        else: Messagebox.show_info("No files were created.", "Info")

# --- Startup ---
try: set_active_excel(ACTIVE_EXCEL)
except: pass

def on_close():
    if PARSING_THREAD:
        if Messagebox.show_question("Parsing in progress. Quit?", "Exit", buttons=['No:secondary', 'Yes:danger']) == 'Yes':
            root.destroy()
    else:
        root.destroy()

root.protocol("WM_DELETE_WINDOW", on_close)
root.mainloop()
