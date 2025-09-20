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
from tkinter import ttk, filedialog, messagebox

# App modules (existing files you already have)
from utils import load_config, save_config, load_nlp_prefer_transformer, log_error
from parser import parse_resume
from excel_handler import append_row, ensure_excel, email_duplicate_within_days, read_all_rows, get_headers, update_headers

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

# ----------------- Load spaCy model (local-first) -----------------
try:
    nlp, model_name = load_nlp_prefer_transformer(("en_core_web_trf", "en_core_web_sm"))
    CONFIG["last_used_model"] = model_name
    save_config(CONFIG)
except Exception as e:
    log_error("NLP model load error: " + str(e) + "\n" + traceback.format_exc())
    # show friendly popup but allow user to continue (they can copy model later)
    tk.Tk().withdraw()
    messagebox.showerror("Model Error",
                         "Could not load a local spaCy model.\n\n"
                         "Please place an 'en_core_web_sm' or 'en_core_web_trf' folder next to the EXE or install the model.\n"
                         "The app will exit now.")
    sys.exit(1)

# ----------------- Main window -----------------
root = tk.Tk()
root.title("Resume Parser — Offline (Workspace on Desktop)")
root.geometry("1000x700")

def ui_call(fn, *a, **kw):
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
        if os.name == "nt":
            os.startfile(path)
        elif sys.platform == "darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])
    except Exception as e:
        messagebox.showerror("Open Error", f"Could not open {path}: {e}")

def set_active_excel(path: Path):
    global ACTIVE_EXCEL
    ACTIVE_EXCEL = Path(path)
    CONFIG["active_excel"] = str(ACTIVE_EXCEL)
    save_config(CONFIG)
    lbl_active.config(text=f"Active Excel: {ACTIVE_EXCEL}")

# ----------------- GUI Layout -----------------
# Top controls
frame_top = ttk.Frame(root)
frame_top.pack(fill="x", padx=10, pady=8)

lbl_active = ttk.Label(frame_top, text=f"Active Excel: {ACTIVE_EXCEL}")
lbl_active.pack(side="left", padx=6)

btn_open_excel = ttk.Button(frame_top, text="Open Active Excel", width=18,
                            command=lambda: open_in_explorer(ACTIVE_EXCEL) if ACTIVE_EXCEL.exists() else messagebox.showwarning("Not found", "Active Excel not found"))
btn_open_excel.pack(side="right", padx=6)

btn_export = ttk.Button(frame_top, text="Export Visible → CSV", width=18)
btn_export.pack(side="right", padx=6)

btn_change_excel = ttk.Button(frame_top, text="Create / Select Excel", width=20)
btn_change_excel.pack(side="right", padx=6)


btn_multi = ttk.Button(frame_top, text="Upload Multiple Resumes", width=22)
btn_multi.pack(side="right", padx=6)

btn_single = ttk.Button(frame_top, text="Upload Resume", width=16)
btn_single.pack(side="right", padx=6)

# Search
frame_search = ttk.Frame(root)
frame_search.pack(fill="x", padx=10, pady=6)
ttk.Label(frame_search, text="Search:").pack(side="left", padx=(0,6))
search_var = tk.StringVar()
entry_search = ttk.Entry(frame_search, textvariable=search_var, width=50)
entry_search.pack(side="left", padx=(0,6))

# Tree (results)
cols = ["Name", "Email", "Phone", "Skills", "Experience", "DateApplied"]
tree = ttk.Treeview(root, columns=cols, show="headings")
for c in cols:
    tree.heading(c, text=c)
    tree.column(c, width=150)
tree.pack(fill="both", expand=True, padx=10, pady=(4,8))

# Progress & status area
progress = ttk.Progressbar(root, orient="horizontal", mode="determinate")
progress.pack(fill="x", padx=12, pady=6)

frame_status = ttk.Frame(root)
frame_status.pack(fill="both", padx=12, pady=(0,12))
status_box = tk.Text(frame_status, height=6, state="disabled", wrap="word")
status_box.pack(fill="both", expand=True)

# Duplicate controls (bottom-left)
frame_dup = ttk.Frame(root)
frame_dup.pack(fill="x", padx=12, pady=(0,8))
var_dup = tk.BooleanVar(value=DUPLICATE_ENABLED)
chk_dup = ttk.Checkbutton(frame_dup, text="Enable duplicate-check by Email", variable=var_dup)
chk_dup.pack(side="left", padx=(0,8))
ttk.Label(frame_dup, text="Window (days):").pack(side="left", padx=(6,4))
var_days = tk.IntVar(value=DUPLICATE_DAYS)
spin_days = ttk.Spinbox(frame_dup, from_=1, to=365, textvariable=var_days, width=6)
spin_days.pack(side="left")

# ----------------- GUI behavior functions -----------------
def refresh_tree_from_excel(path: Path):
    """Load rows from Excel into the tree."""
    for r in tree.get_children():
        tree.delete(r)
    if not path.exists():
        return
    rows = read_all_rows(str(path))
    for r in rows:
        # expected order: Name, Email, Phone, Skills, Experience, DateApplied, ResumePath
        tree.insert("", "end", values=(r[0], r[1], r[2], r[3], r[4], r[5] if len(r) > 5 else ""))

def on_search_change(*_):
    q = search_var.get().strip().lower()
    for item in tree.get_children():
        values = tree.item(item, "values")
        combined = " ".join([str(v).lower() for v in values])
        if q == "" or q in combined:
            tree.reattach(item, "", "end")
        else:
            tree.detach(item)

entry_search.bind("<KeyRelease>", on_search_change)

# ----------------- Excel creation / selection -----------------
def create_new_excel_dialog():
    # Ask for filename and columns, then create in EXCEL_DIR
    cols_default = ",".join(cols + ["ResumePath"])
    dlg = tk.Toplevel(root)
    dlg.title("Create New Excel")
    dlg.geometry("480x220")
    ttk.Label(dlg, text="Enter comma-separated column names:").pack(pady=(10,6))
    txt = tk.Text(dlg, height=6)
    txt.pack(fill="both", padx=12, pady=6)
    txt.insert("1.0", cols_default)
    def do_create():
        cols_input = [c.strip() for c in txt.get("1.0","end").split(",") if c.strip()]
        file = filedialog.asksaveasfilename(initialdir=str(EXCEL_DIR),
                                            defaultextension=".xlsx", filetypes=[("Excel files","*.xlsx")])
        if not file:
            return
        ensure_excel(file, cols_input)
        set_active_excel(Path(file))
        refresh_tree_from_excel(ACTIVE_EXCEL)
        add_status(f"Created and selected Excel: {file}")
        dlg.destroy()
    ttk.Button(dlg, text="Create", command=do_create).pack(pady=8)

def select_existing_excel():
    file = filedialog.askopenfilename(initialdir=str(EXCEL_DIR), filetypes=[("Excel files","*.xlsx")])
    if not file:
        return
    set_active_excel(Path(file))
    refresh_tree_from_excel(ACTIVE_EXCEL)
    add_status(f"Selected active Excel: {file}")

def on_change_excel():
    resp = messagebox.askquestion("Choose action", "Create a new Excel file? (Yes to create new, No to select existing)")
    if resp == "yes":
        create_new_excel_dialog()
    else:
        select_existing_excel()

btn_change_excel.config(command=on_change_excel)

# ----------------- File selection helpers -----------------
def select_single_file():
    f = filedialog.askopenfilename(initialdir=str(RESUMES_DIR), filetypes=[("Resumes","*.pdf;*.docx"), ("All files","*.*")])
    return f

def select_multiple_files():
    f = filedialog.askopenfilenames(initialdir=str(RESUMES_DIR), filetypes=[("Resumes","*.pdf;*.docx"), ("All files","*.*")])
    return list(f)

# ----------------- Processing worker -----------------
def process_files_sequential(file_paths):
    """Background worker to parse files sequentially, save to Excel, update UI via ui_call."""
    ui_call(btn_single.config, state="disabled")
    ui_call(btn_multi.config, state="disabled")
    ui_call(btn_change_excel.config, state="disabled")
    ui_call(btn_open_excel.config, state="disabled")
    ui_call(btn_export.config, state="disabled")

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
                add_status(f"{fname} — Failed: no extractable text (scanned PDF?)")
                results.append((fname, "Failed", "No text"))
                processed += 1
                ui_call(progress.step, 1)
                continue

            # duplicate check
            if var_dup.get() and email_duplicate_within_days(str(ACTIVE_EXCEL), data.get("Email", ""), int(var_days.get())):
                add_status(f"{fname} — Skipped: duplicate within window")
                results.append((fname, "Skipped", "Duplicate"))
            else:
                ensure_excel(str(ACTIVE_EXCEL))
                append_row(str(ACTIVE_EXCEL), data)
                # update tree on UI thread
                ui_call(tree.insert, "", "end", values=(data.get("Name",""), data.get("Email",""), data.get("Phone",""),
                                                       data.get("Skills",""), data.get("Experience",""), "(today)"))
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
