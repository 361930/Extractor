# app.py — Cleaned, consolidated GUI for ResumeParserApp
import os
import sys
import threading
import traceback
import subprocess
import csv
from shutil import copy2
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, font

# App modules (existing files you already have)
from utils import load_config, save_config, load_nlp_prefer_transformer, log_error
from parser import parse_resume
from excel_handler import append_row, ensure_excel, email_duplicate_within_days, read_all_rows, DEFAULT_COLS

# ----------------- Configuration / Workspace -----------------
HOME = Path.home()
WORKSPACE = HOME / "Desktop" / "ResumeParserWorkspace"
RESUMES_DIR = WORKSPACE / "resumes"
EXCEL_DIR = WORKSPACE / "excel"
LOG_DIR = Path("logs")
LOG_DIR.mkdir(exist_ok=True)
WORKSPACE.mkdir(parents=True, exist_ok=True)
RESUMES_DIR.mkdir(parents=True, exist_ok=True)
EXCEL_DIR.mkdir(parents=True, exist_ok=True)

CONFIG = load_config()
ACTIVE_EXCEL = Path(CONFIG.get("active_excel", str(EXCEL_DIR / "resumes_data.xlsx")))
DUPLICATE_ENABLED = CONFIG.get("duplicate_check_enabled", True)
DUPLICATE_DAYS = CONFIG.get("duplicate_days", 30)

# Define the columns visible in the UI
UI_COLS = ["S.No.", "Name", "Email", "Phone", "OriginalFile"]

# ----------------- Load spaCy model (local-first) -----------------
nlp = None # Default
model_name = "N/A"
try:
    nlp, model_name = load_nlp_prefer_transformer(("en_core_web_trf", "en_core_web_sm"))
    CONFIG["last_used_model"] = model_name
    save_config(CONFIG)
except Exception as e:
    log_error("NLP model load error: " + str(e) + "\n" + traceback.format_exc())
    # show friendly popup but allow user to continue
    # Need a temporary root to show message box before mainloop
    _temp_root = tk.Tk()
    _temp_root.withdraw()
    messagebox.showwarning("Model Warning",
                         "Could not load a local spaCy model (e.g., 'en_core_web_sm').\n\n"
                         "Name extraction will be less accurate.\n"
                         "You can place the model folder next to the EXE to fix this.")
    _temp_root.destroy()
    nlp = None # Ensure nlp is None so parser can handle it

# ----------------- Main window -----------------
root = tk.Tk()
root.title(f"Resume Parser (Model: {model_name})")
root.geometry("1100x750")

# --- Style ---
style = ttk.Style(root)
try:
    # 'clam', 'alt', 'default', 'classic'
    style.theme_use("clam") 
except tk.TclError:
    pass # Use default

# Configure fonts
try:
    # Use a common, clean font
    font_family = "Segoe UI"
    # Check if Segoe UI is available, otherwise fall back
    if font_family not in font.families():
        font_family = "Verdana" # Another common clean font
    if font_family not in font.families():
        font_family = "Arial" # Safe fallback

    default_font = font.nametofont("TkDefaultFont")
    default_font.configure(family=font_family, size=10)
    text_font = font.nametofont("TkTextFont")
    text_font.configure(family=font_family, size=10)
    root.option_add("*Font", default_font)
except Exception as e:
    print(f"Font config error: {e}") # Non-fatal

# Configure widget styles
bg_color = "#f4f4f4" # Light grey background
text_color = "#1f1f1f"
button_bg = "#e1e1e1"
button_active_bg = "#c8c8c8"
select_bg = "#0078d4" # Blue selection
select_fg = "#ffffff" # White text on selection

root.configure(bg=bg_color)

style.configure("TFrame", background=bg_color)
style.configure("TLabel", background=bg_color, foreground=text_color, padding=4)
style.configure("TButton", padding=(10, 5), relief="flat",
                background=button_bg, foreground=text_color)
style.map("TButton",
    background=[('active', button_active_bg), ('!active', button_bg)],
    relief=[('pressed', 'sunken'), ('!pressed', 'flat')])

style.configure("TEntry", padding=6, relief="solid", borderwidth=1,
                fieldbackground="#ffffff", foreground=text_color)
style.configure("Treeview.Heading", font=(font_family, 10, "bold"), padding=8,
                background=button_bg, relief="flat")
style.map("Treeview.Heading",
    background=[('active', button_active_bg), ('!active', button_bg)])
style.configure("Treeview", rowheight=28, font=(font_family, 10),
                background="#ffffff", fieldbackground="#ffffff", foreground=text_color)
style.map("Treeview",
    background=[('selected', select_bg)],
    foreground=[('selected', select_fg)])

style.configure("Vertical.TScrollbar", relief="flat", troughcolor=bg_color)
style.configure("Horizontal.TScrollbar", relief="flat", troughcolor=bg_color)
style.configure("TProgressbar", troughcolor=bg_color, background=select_bg, relief="solid", borderwidth=1)
style.configure("TCheckbutton", background=bg_color, padding=4)
style.configure("TSpinbox", padding=4, relief="solid", borderwidth=1)

status_box_font = (font_family, 9)
# --- End Style ---


def ui_call(fn, *a, **kw):
    """Helper to run a function on the UI thread."""
    root.after(0, lambda: fn(*a, **kw))

# ----------------- Helpers -----------------
def add_status(text: str):
    """Append a status line to status_box (thread-safe)."""
    def _append():
        status_box.configure(state="normal")
        status_box.insert("end", text + "\n")
        status_box.see("end")
        status_box.configure(state="disabled")
    ui_call(_append)

def open_in_explorer(path: Path):
    try:
        if os.name == "nt": # Windows
            # Use os.startfile on the *directory* to select the file
            if path.is_file():
                subprocess.Popen(f'explorer /select,"{path}"')
            else:
                os.startfile(path)
        elif sys.platform == "darwin": # macOS
            subprocess.Popen(["open", "-R", path] if path.is_file() else ["open", path])
        else: # Linux
            subprocess.Popen(["xdg-open", path.parent if path.is_file() else path])
    except Exception as e:
        add_status(f"Error opening {path}: {e}")
        messagebox.showerror("Open Error", f"Could not open {path}: {e}")

def set_active_excel(path: Path):
    global ACTIVE_EXCEL
    ACTIVE_EXCEL = Path(path)
    CONFIG["active_excel"] = str(ACTIVE_EXCEL)
    save_config(CONFIG)
    lbl_active.config(text=f"Active Excel: {ACTIVE_EXCEL.name}")
    add_status(f"Active Excel set to: {ACTIVE_EXCEL}")

# ----------------- GUI Layout -----------------
# Top controls
frame_top = ttk.Frame(root)
frame_top.pack(fill="x", padx=10, pady=(10, 5))

lbl_active = ttk.Label(frame_top, text=f"Active Excel: {ACTIVE_EXCEL.name}", anchor="w")
lbl_active.pack(side="left", padx=6, fill="x", expand=True)

# Right-side buttons (pack in reverse order)
btn_open_excel = ttk.Button(frame_top, text="Open Excel", width=14,
                            command=lambda: open_in_explorer(ACTIVE_EXCEL) if ACTIVE_EXCEL.exists() else messagebox.showwarning("Not found", "Active Excel not found"))
btn_open_excel.pack(side="right", padx=(4, 6))

btn_export = ttk.Button(frame_top, text="Export Visible → CSV", width=20)
btn_export.pack(side="right", padx=4)

btn_change_excel = ttk.Button(frame_top, text="Create / Select Excel", width=20)
btn_change_excel.pack(side="right", padx=4)

btn_folder = ttk.Button(frame_top, text="Upload Folder", width=16)
btn_folder.pack(side="right", padx=4)

btn_multi = ttk.Button(frame_top, text="Upload Files", width=14)
btn_multi.pack(side="right", padx=4)

# Search
frame_search = ttk.Frame(root)
frame_search.pack(fill="x", padx=10, pady=6)
ttk.Label(frame_search, text="Search:").pack(side="left", padx=(6, 6))
search_var = tk.StringVar()
entry_search = ttk.Entry(frame_search, textvariable=search_var)
entry_search.pack(side="left", fill="x", expand=True, padx=(0, 6))

# Tree (results)
frame_tree = ttk.Frame(root)
frame_tree.pack(fill="both", expand=True, padx=10, pady=(4, 8))

tree = ttk.Treeview(frame_tree, columns=UI_COLS, show="headings")
for c in UI_COLS:
    tree.heading(c, text=c)
# Set column widths
tree.column("S.No.", width=60, stretch=tk.NO, anchor="center")
tree.column("Name", width=220)
tree.column("Email", width=280)
tree.column("Phone", width=200)
tree.column("OriginalFile", width=250)

# Scrollbars
yscroll = ttk.Scrollbar(frame_tree, orient="vertical", command=tree.yview)
xscroll = ttk.Scrollbar(root, orient="horizontal", command=tree.xview)
tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)

yscroll.pack(side="right", fill="y")
tree.pack(side="left", fill="both", expand=True)
xscroll.pack(fill="x", padx=10, pady=(0, 4))


# Progress & status area
progress = ttk.Progressbar(root, orient="horizontal", mode="determinate")
progress.pack(fill="x", padx=12, pady=6)

frame_status = ttk.Frame(root)
frame_status.pack(fill="x", padx=12, pady=(0,0))
status_box = tk.Text(frame_status, height=5, state="disabled", wrap="word",
                     relief="solid", borderwidth=1, font=status_box_font,
                     background="#ffffff", foreground=text_color)
status_box.pack(fill="both", expand=True)

# Duplicate controls (bottom-left)
frame_dup = ttk.Frame(root)
frame_dup.pack(fill="x", padx=12, pady=(4, 10))
var_dup = tk.BooleanVar(value=DUPLICATE_ENABLED)
chk_dup = ttk.Checkbutton(frame_dup, text="Enable duplicate-check by Email", variable=var_dup)
chk_dup.pack(side="left", padx=(6, 8))
ttk.Label(frame_dup, text="Window (days):").pack(side="left", padx=(6, 4))
var_days = tk.IntVar(value=DUPLICATE_DAYS)
spin_days = ttk.Spinbox(frame_dup, from_=1, to=365, textvariable=var_days, width=6)
spin_days.pack(side="left")

# ----------------- GUI behavior functions -----------------
def refresh_tree_from_excel(path: Path):
    """Load rows from Excel into the tree."""
    for r in tree.get_children():
        tree.delete(r)
    if not path.exists():
        add_status(f"Excel file not found at: {path}")
        return
        
    try:
        rows = read_all_rows(str(path))
        for r in rows:
            # expected order: S.No, Name, Email, Phone, OriginalFile, ...
            if not r or r[0] is None: continue # Skip empty rows
            
            # Get values for UI_COLS
            s_no = r[0]
            name = r[1] if len(r) > 1 else ""
            email = r[2] if len(r) > 2 else ""
            phone = r[3] if len(r) > 3 else ""
            orig_file = r[4] if len(r) > 4 else ""
            
            tree.insert("", "end", values=(s_no, name, email, phone, orig_file))
        
        # After refresh, trigger search filter again
        on_search_change()
        
    except Exception as e:
        log_error(f"Failed to refresh tree from excel: {e}\n{traceback.format_exc()}")
        add_status(f"Error: Could not read Excel file. Is it open or corrupted?")


def on_search_change(*_):
    q = search_var.get().strip().lower()
    for item in tree.get_children():
        values = tree.item(item, "values")
        try:
            # Search Name, Email, Phone, OriginalFile (indices 1, 2, 3, 4)
            combined = f"{values[1]} {values[2]} {values[3]} {values[4]}".lower()
            if q == "" or q in combined:
                # This logic doesn't work well if item is already attached
                # Instead, we should store all items and filter
                # For now, just show/hide
                # A better way is to detach all and re-attach matches
                pass # Re-visit this. Simple detach/reattach is complex with search.
        except Exception:
            pass # Ignore errors on filtering
            
    # Simple (but flickery) filter: detach all, re-attach matches
    all_items = tree.get_children()
    for item in all_items:
        tree.detach(item)
    
    for item in all_items:
        values = tree.item(item, "values")
        try:
            combined = f"{values[1]} {values[2]} {values[3]} {values[4]}".lower()
            if q == "" or q in combined:
                tree.reattach(item, "", "end")
        except Exception:
            if q == "": # Re-attach if search is empty
                tree.reattach(item, "", "end")


entry_search.bind("<KeyRelease>", on_search_change)

# ----------------- Excel creation / selection -----------------
def create_new_excel_dialog():
    # Columns are fixed now, so just ask for a name.
    file = filedialog.asksaveasfilename(initialdir=str(EXCEL_DIR),
                                        title="Create New Excel File",
                                        defaultextension=".xlsx",
                                        filetypes=[("Excel files","*.xlsx")])
    if not file:
        return
    try:
        ensure_excel(file) # This uses the new DEFAULT_COLS
        set_active_excel(Path(file))
        refresh_tree_from_excel(ACTIVE_EXCEL)
        add_status(f"Created and selected Excel: {file}")
    except Exception as e:
        log_error(f"Failed to create new excel: {e}\n{traceback.format_exc()}")
        messagebox.showerror("Error", f"Failed to create file: {e}")

def select_existing_excel():
    file = filedialog.askopenfilename(initialdir=str(EXCEL_DIR),
                                      title="Select Existing Excel File",
                                      filetypes=[("Excel files","*.xlsx")])
    if not file:
        return
    
    # Simple validation: Check if header matches
    try:
        wb = load_workbook(file, read_only=True)
        ws = wb.active
        headers = [cell.value for cell in ws[1]]
        wb.close()
        
        # Check if *at least* the first 5 columns match our UI_COLS
        if headers[:5] != UI_COLS:
            if not messagebox.askyesno("Header Mismatch",
                                      "The selected Excel file doesn't seem to have the standard columns (S.No., Name, Email, ...).\n\n"
                                      "Do you want to use it anyway? (Not recommended)"):
                return
    except Exception as e:
        log_error(f"Error checking excel header: {e}")
        if not messagebox.askyesno("Read Error",
                                  f"Could not read the file's headers. Error: {e}\n\n"
                                  "Do you want to select it anyway?"):
            return

    set_active_excel(Path(file))
    refresh_tree_from_excel(ACTIVE_EXCEL)
    add_status(f"Selected active Excel: {file}")

def on_change_excel():
    resp = messagebox.askquestion("Choose action", "Create a new Excel file? (Yes to create new, No to select existing)")
    if resp == "yes":
        create_new_excel_dialog()
    elif resp == "no":
        select_existing_excel()

btn_change_excel.config(command=on_change_excel)

# ----------------- UI Button State Toggle -----------------
def set_ui_processing(processing: bool):
    """Disable/enable buttons during processing."""
    state = "disabled" if processing else "normal"
    btn_multi.config(state=state)
    btn_folder.config(state=state)
    btn_change_excel.config(state=state)
    btn_open_excel.config(state=state)
    btn_export.config(state=state)
    entry_search.config(state=state)

# ----------------- Processing worker -----------------
def process_files_sequential(file_paths):
    """Background worker to parse files sequentially, save to Excel, update UI via ui_call."""
    ui_call(set_ui_processing, True)

    total = len(file_paths)
    processed = 0
    ui_call(progress.config, maximum=total, value=0)

    results = []
    for fp in file_paths:
        fname = Path(fp).name
        try:
            add_status(f"Processing {fname} ...")
            data = parse_resume(fp, nlp)
            if not data:
                add_status(f"{fname} — Skipped: no extractable text (scanned PDF or empty?)")
                results.append((fname, "Failed", "No text"))
                processed += 1
                ui_call(progress.step, 1)
                continue

            # Add original filename to data dict for excel handler
            data["OriginalFile"] = fname

            # duplicate check
            if var_dup.get() and email_duplicate_within_days(str(ACTIVE_EXCEL), data.get("Email", ""), int(var_days.get())):
                add_status(f"{fname} — Skipped: duplicate email found")
                results.append((fname, "Skipped", "Duplicate"))
            else:
                # ensure_excel(str(ACTIVE_EXCEL), DEFAULT_COLS) # Ensure it exists
                append_row(str(ACTIVE_EXCEL), data)
                # Don't update tree here. Refresh at the end.
                add_status(f"{fname} — Added to Excel")
                results.append((fname, "Added", "Saved"))

        except PermissionError:
            log_error(f"Processing error for {fp}: PermissionError\n" + traceback.format_exc())
            add_status(f"{fname} — ERROR: Permission denied. Is Excel open?")
            results.append((fname, "Error", "PermissionError"))
            ui_call(messagebox.showerror, "File Locked", f"Could not write to Excel file: {ACTIVE_EXCEL.name}.\n\nPlease close the file and try again.")
            break # Stop processing batch if excel is locked
        except Exception as e:
            log_error(f"Processing error for {fp}: {e}\n" + traceback.format_exc())
            add_status(f"{fname} — ERROR: {e}")
            results.append((fname, "Error", str(e)))
        finally:
            processed += 1
            ui_call(progress.step, 1)
            
    add_status("Batch complete. Refreshing table view...")
    ui_call(refresh_tree_from_excel, ACTIVE_EXCEL)

    # re-enable UI
    ui_call(set_ui_processing, False)
    ui_call(progress.config, value=0)

    # summary
    added = sum(1 for r in results if r[1] == "Added")
    skipped = sum(1 for r in results if r[1] == "Skipped")
    failed = sum(1 for r in results if r[1] in ("Failed","Error"))
    ui_call(messagebox.showinfo, "Batch Complete", f"Processed {processed}/{total} files.\n\nAdded: {added}\nSkipped (Duplicate): {skipped}\nFailed (No Text/Error): {failed}")

def safe_copy_to_workspace(src_path: str):
    """Copy src to RESUMES_DIR with collision-avoidance. Return destination Path or None."""
    try:
        src = Path(src_path)
    except Exception as e:
        add_status(f"Invalid source path: {src_path}. Error: {e}")
        return None
        
    # disallow selecting the active excel by mistake
    try:
        if ACTIVE_EXCEL.exists() and src.resolve() == ACTIVE_EXCEL.resolve():
            messagebox.showwarning("Wrong file", "Selected file is the active Excel. Please select a resume (PDF/DOCX).")
            return None
    except Exception:
        pass # resolution failed? continue cautiously

    dst = RESUMES_DIR / src.name
    base = dst.stem
    ext = dst.suffix
    i = 1
    # Avoid overwriting existing files
    while dst.exists():
        dst = RESUMES_DIR / f"{base}_{i}{ext}"
        i += 1

    try:
        copy2(src, dst)
    except Exception as e:
        add_status(f"Failed to copy {src.name}: {e}")
        log_error(f"Failed to copy {src.name} to {dst}: {e}\n{traceback.format_exc()}")
        return None # Return None on copy failure
    return dst

def on_upload_multiple():
    fps = filedialog.askopenfilenames(initialdir=str(RESUMES_DIR),
                                      title="Select Resume Files",
                                      filetypes=[("Resumes","*.pdf;*.docx"), ("All files","*.*")])
    if not fps:
        return
        
    add_status(f"Copying {len(fps)} files to workspace...")
    saved = []
    for f in fps:
        dst = safe_copy_to_workspace(f)
        if dst:
            saved.append(str(dst))
            
    if saved:
        add_status(f"Starting processing for {len(saved)} files...")
        threading.Thread(target=process_files_sequential, args=(saved,), daemon=True).start()
    else:
        add_status("No files were successfully copied.")

def on_upload_folder():
    dir_path = filedialog.askdirectory(initialdir=str(RESUMES_DIR),
                                      title="Select Folder Containing Resumes")
    if not dir_path:
        return
    
    add_status(f"Scanning folder: {dir                                                    data.get("Skills",""), data.get("Experience",""), "(today)"))
                add_status(f"{fname} — Added")
                results.append((fname, "Added", "Saved"))

        except Exception as e:
            log_error(f"Processing error for {fp}: {e}\n" + traceback.format_exc())
            add_status(f"{fname} — Error: {e}")
            results.append((fname, "Error", str(e)))
        finally:
            processed += 1
            ui_call(progress.step, 1)

    # re-enable UI
    ui_call(btn_single.config, state="normal")
    ui_call(btn_multi.config, state="normal")
    ui_call(btn_change_excel.config, state="normal")
    ui_call(btn_open_excel.config, state="normal")
    ui_call(btn_export.config, state="normal")

    # summary
    added = sum(1 for r in results if r[1] == "Added")
    skipped = sum(1 for r in results if r[1] == "Skipped")
    failed = sum(1 for r in results if r[1] in ("Failed","Error"))
    ui_call(messagebox.showinfo, "Batch Complete", f"Processed {total} files — {added} added, {skipped} skipped, {failed} failed.")

def safe_copy_to_workspace(src_path: str):
    """Copy src to RESUMES_DIR with collision-avoidance. Return destination Path or None."""
    src = Path(src_path)
    # disallow selecting the active excel by mistake
    try:
        if ACTIVE_EXCEL.exists() and src.resolve() == ACTIVE_EXCEL.resolve():
            messagebox.showwarning("Wrong file", "Selected file is the active Excel. Please select a resume (PDF/DOCX).")
            return None
    except Exception:
        # resolution failed? continue cautiously
        pass

    dst = RESUMES_DIR / src.name
    base = dst.stem
    ext = dst.suffix
    i = 1
    while dst.exists():
        dst = RESUMES_DIR / f"{base}_{i}{ext}"
        i += 1

    try:
        copy2(src, dst)
    except Exception as e:
        add_status(f"Failed to copy {src.name}: {e}")
        raise
    return dst

def on_upload_single():
    fp = filedialog.askopenfilename(initialdir=str(RESUMES_DIR), filetypes=[("Resumes","*.pdf;*.docx"), ("All files","*.*")])
    if not fp:
        return
    dst = safe_copy_to_workspace(fp)
    if not dst:
        return
    threading.Thread(target=process_files_sequential, args=([str(dst)],), daemon=True).start()

def on_upload_multiple():
    fps = filedialog.askopenfilenames(initialdir=str(RESUMES_DIR), filetypes=[("Resumes","*.pdf;*.docx"), ("All files","*.*")])
    if not fps:
        return
    saved = []
    for f in fps:
        dst = safe_copy_to_workspace(f)
        if dst:
            saved.append(str(dst))
    if saved:
        threading.Thread(target=process_files_sequential, args=(saved,), daemon=True).start()


btn_single.config(command=on_upload_single)
btn_multi.config(command=on_upload_multiple)

def edit_columns_dialog():
    """Open modal to edit/add/remove/reorder columns for ACTIVE_EXCEL."""
    if not ACTIVE_EXCEL or not ACTIVE_EXCEL.exists():
        messagebox.showwarning("No Excel", "No active Excel file found. Create or select an Excel file first.")
        return

    # load current headers
    headers = get_headers(str(ACTIVE_EXCEL))
    if not headers:
        headers = ["Name", "Email", "Phone", "Skills", "Experience", "DateApplied", "ResumePath"]

    dlg = tk.Toplevel(root)
    dlg.title("Edit Excel Columns")
    dlg.transient(root)
    dlg.grab_set()

    # center the dialog and ensure it's on top/focused (helps if dialog was off-screen)
    dlg.update_idletasks()
    w, h = 560, 460
    x = (root.winfo_screenwidth() - w) // 2
    y = (root.winfo_screenheight() - h) // 2
    dlg.geometry(f"{w}x{h}+{x}+{y}")
    dlg.lift()
    dlg.focus_force()


    lbl = ttk.Label(dlg, text=f"Editing: {ACTIVE_EXCEL.name}")
    lbl.pack(anchor="w", padx=10, pady=(8,4))

    frame = ttk.Frame(dlg)
    frame.pack(fill="both", expand=True, padx=10, pady=6)

    listbox = tk.Listbox(frame, selectmode="browse", height=12)
    listbox.pack(side="left", fill="both", expand=True, padx=(0,6))
    for h in headers:
        listbox.insert("end", h)

    # scrollbar
    scr = ttk.Scrollbar(frame, orient="vertical", command=listbox.yview)
    scr.pack(side="left", fill="y")
    listbox.config(yscrollcommand=scr.set)

    # right-side controls
    ctrl = ttk.Frame(frame)
    ctrl.pack(side="left", fill="y", padx=(6,0))

    entry_new = ttk.Entry(ctrl, width=24)
    entry_new.pack(pady=(2,6))

    def add_col():
        name = entry_new.get().strip()
        if not name:
            messagebox.showwarning("Empty", "Enter a column name.")
            return
        listbox.insert("end", name)
        entry_new.delete(0, "end")

    def remove_col():
        sel = listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        listbox.delete(idx)

    def move_up():
        sel = listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        if idx == 0:
            return
        val = listbox.get(idx)
        listbox.delete(idx)
        listbox.insert(idx-1, val)
        listbox.select_set(idx-1)

    def move_down():
        sel = listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        if idx == listbox.size()-1:
            return
        val = listbox.get(idx)
        listbox.delete(idx)
        listbox.insert(idx+1, val)
        listbox.select_set(idx+1)

    btn_add = ttk.Button(ctrl, text="Add", command=add_col, width=14)
    btn_add.pack(pady=4)
    btn_remove = ttk.Button(ctrl, text="Remove", command=remove_col, width=14)
    btn_remove.pack(pady=4)
    btn_up = ttk.Button(ctrl, text="Move Up", command=move_up, width=14)
    btn_up.pack(pady=4)
    btn_down = ttk.Button(ctrl, text="Move Down", command=move_down, width=14)
    btn_down.pack(pady=4)

    # Save / Cancel
    bottom = ttk.Frame(dlg)
    bottom.pack(fill="x", padx=10, pady=8)
    def do_save():
        new_headers = [listbox.get(i) for i in range(listbox.size())]
        if not new_headers:
            if not messagebox.askyesno("Empty headers", "You removed all headers. Continue?"):
                return
        # warn about truncation if new length shorter than existing
        old_headers = headers
        if len(new_headers) < len(old_headers):
            if not messagebox.askyesno("Truncate columns",
                                       "New header list is shorter than existing. This will truncate data in trailing columns. Proceed?"):
                return

        # Try to update headers; loop on PermissionError and allow Retry
        while True:
            try:
                add_status(f"Backing up and updating headers for {Path(ACTIVE_EXCEL).name}...")
                update_headers(str(ACTIVE_EXCEL), new_headers)
                refresh_tree_from_excel(ACTIVE_EXCEL)
                add_status("Headers updated.")
                dlg.destroy()
                break
            except PermissionError as pe:
                # File is likely open in Excel
                retry = messagebox.askretrycancel("File Locked",
                                                  f"Cannot write to {Path(ACTIVE_EXCEL).name} because it is open.\n\n"
                                                  f"Please close the file in Excel and click Retry.\n\nError: {pe}")
                if not retry:
                    add_status("Update cancelled (file locked).")
                    break
                # else loop and retry
            except Exception as e:
                messagebox.showerror("Error", f"Failed to update headers: {e}")
                add_status(f"Failed to update headers: {e}")
                break


    ttk.Button(bottom, text="Save", command=do_save).pack(side="right", padx=(4,0))
    ttk.Button(bottom, text="Cancel", command=dlg.destroy).pack(side="right")

    btn_edit_columns = ttk.Button(frame_top, text="Edit Columns", width=14, command=edit_columns_dialog)
    btn_edit_columns.pack(side="right", padx=6)

# ----------------- Export visible CSV -----------------
def export_visible_to_csv():
    save_path = filedialog.asksaveasfilename(initialdir=str(EXCEL_DIR), defaultextension=".csv", filetypes=[("CSV files","*.csv")])
    if not save_path:
        return
    try:
        with open(save_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(cols)
            for item in tree.get_children():
                writer.writerow(tree.item(item, "values"))
        messagebox.showinfo("Export", f"Exported {len(tree.get_children())} rows to {save_path}")
    except Exception as e:
        messagebox.showerror("Export Error", str(e))

btn_export.config(command=export_visible_to_csv)

# ----------------- Final initialization -----------------
# Ensure the active excel exists (create if not)
if not ACTIVE_EXCEL.exists():
    ensure_excel(str(ACTIVE_EXCEL))
set_active_excel(ACTIVE_EXCEL)
refresh_tree_from_excel(ACTIVE_EXCEL)
add_status("Ready — workspace: " + str(WORKSPACE))

# Start GUI
root.mainloop()

