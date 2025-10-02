# app.py — Modernized with CustomTkinter
import os
import sys
import threading
import traceback
import subprocess
import csv
from shutil import copy2
from pathlib import Path
import customtkinter as ctk
from tkinter import filedialog, messagebox, ttk

# App modules
from utils import load_config, save_config, load_nlp_prefer_transformer, log_error
from parser import parse_resume
from excel_handler import append_row, ensure_excel, email_duplicate_within_days, read_all_rows, get_headers, update_headers

# --- Configuration / Workspace ---
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

# --- Load spaCy model (local-first) ---
try:
    nlp, model_name = load_nlp_prefer_transformer(("en_core_web_trf", "en_core_web_sm"))
    CONFIG["last_used_model"] = model_name
    save_config(CONFIG)
except Exception as e:
    log_error("NLP model load error: " + str(e) + "\n" + traceback.format_exc())
    root = ctk.CTk()
    root.withdraw()
    messagebox.showerror("Model Error",
                         "Could not load a local spaCy model.\n\n"
                         "Please place 'en_core_web_sm' or 'en_core_web_trf' folder next to the app, or install the model via pip.\n"
                         "The app will now exit.")
    sys.exit(1)

# --- Main window ---
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

root = ctk.CTk()
root.title("Resume Parser")
root.geometry("1200x800")

def ui_call(fn, *a, **kw):
    root.after(0, lambda: fn(*a, **kw))

# --- Helpers ---
def add_status(text: str):
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
            subprocess.Popen(["open", str(path)])
        else:
            subprocess.Popen(["xdg-open", str(path)])
    except Exception as e:
        messagebox.showerror("Open Error", f"Could not open {path}: {e}")

def set_active_excel(path: Path):
    global ACTIVE_EXCEL
    ACTIVE_EXCEL = Path(path)
    CONFIG["active_excel"] = str(ACTIVE_EXCEL)
    save_config(CONFIG)
    lbl_active.configure(text=f"Active: {ACTIVE_EXCEL.name}")

# --- GUI Layout ---
root.grid_columnconfigure(0, weight=1)
root.grid_rowconfigure(2, weight=1)

# Top controls
frame_top = ctk.CTkFrame(root)
frame_top.grid(row=0, column=0, sticky="ew", padx=10, pady=10)

btn_single = ctk.CTkButton(frame_top, text="Upload Resume")
btn_single.pack(side="left", padx=6, pady=6)

btn_multi = ctk.CTkButton(frame_top, text="Upload Multiple")
btn_multi.pack(side="left", padx=6, pady=6)

btn_change_excel = ctk.CTkButton(frame_top, text="Create/Select Excel")
btn_change_excel.pack(side="left", padx=6, pady=6)

btn_edit_columns = ctk.CTkButton(frame_top, text="Edit Columns")
btn_edit_columns.pack(side="left", padx=6, pady=6)

btn_export = ctk.CTkButton(frame_top, text="Export to CSV")
btn_export.pack(side="left", padx=6, pady=6)

btn_open_excel = ctk.CTkButton(frame_top, text="Open Excel",
                            command=lambda: open_in_explorer(ACTIVE_EXCEL) if ACTIVE_EXCEL.exists() else messagebox.showwarning("Not found", "Active Excel not found"))
btn_open_excel.pack(side="left", padx=6, pady=6)

# Search and status
frame_search = ctk.CTkFrame(root)
frame_search.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 10))
frame_search.grid_columnconfigure(1, weight=1)

ctk.CTkLabel(frame_search, text="Search:").grid(row=0, column=0, padx=6, pady=6)
search_var = ctk.StringVar()
entry_search = ctk.CTkEntry(frame_search, textvariable=search_var, width=300)
entry_search.grid(row=0, column=1, sticky="ew", padx=6, pady=6)

lbl_active = ctk.CTkLabel(frame_search, text=f"Active: {ACTIVE_EXCEL.name}", anchor="e")
lbl_active.grid(row=0, column=2, padx=6, pady=6, sticky="e")
frame_search.grid_columnconfigure(2, weight=1)


# Tree (results)
style = ttk.Style()
bg_color = root._apply_appearance_mode(ctk.ThemeManager.theme["CTkFrame"]["fg_color"])
text_color = root._apply_appearance_mode(ctk.ThemeManager.theme["CTkLabel"]["text_color"])
header_bg = root._apply_appearance_mode(ctk.ThemeManager.theme["CTkButton"]["fg_color"])

style.theme_use("default")
style.configure("Treeview", background=bg_color, foreground=text_color, fieldbackground=bg_color, borderwidth=0)
style.map('Treeview', background=[('selected', ctk.ThemeManager.theme["CTkButton"]["fg_color"])])
style.configure("Treeview.Heading", background=header_bg, foreground=text_color, relief="flat", padding=(5, 5))
style.map("Treeview.Heading", background=[('active', ctk.ThemeManager.theme["CTkButton"]["hover_color"])])

tree_frame = ctk.CTkFrame(root)
tree_frame.grid(row=2, column=0, sticky="nsew", padx=10, pady=(0, 5))
tree_frame.grid_rowconfigure(0, weight=1)
tree_frame.grid_columnconfigure(0, weight=1)

cols = ["Name", "Email", "Phone", "Skills", "Experience", "DateApplied"]
tree = ttk.Treeview(tree_frame, columns=cols, show="headings")
for c in cols:
    tree.heading(c, text=c)
    tree.column(c, width=180, anchor='w')
tree.grid(row=0, column=0, sticky="nsew")

vsb = ctk.CTkScrollbar(tree_frame, command=tree.yview)
vsb.grid(row=0, column=1, sticky='ns')
tree.configure(yscrollcommand=vsb.set)
hsb = ctk.CTkScrollbar(tree_frame, orientation="horizontal", command=tree.xview)
hsb.grid(row=1, column=0, sticky='ew')
tree.configure(xscrollcommand=hsb.set)

# Progress, Status, and Controls
bottom_frame = ctk.CTkFrame(root)
bottom_frame.grid(row=3, column=0, sticky="ew", padx=10, pady=(5, 10))
bottom_frame.grid_columnconfigure(0, weight=1)

progress = ctk.CTkProgressBar(bottom_frame, orientation="horizontal", mode="determinate")
progress.set(0)
progress.grid(row=0, column=0, columnspan=2, sticky="ew", padx=10, pady=(5, 5))

status_box = ctk.CTkTextbox(bottom_frame, height=120, wrap="word")
status_box.grid(row=1, column=0, sticky="nsew", padx=(10, 5), pady=10)
status_box.configure(state="disabled")

# Duplicate controls
frame_dup = ctk.CTkFrame(bottom_frame)
frame_dup.grid(row=1, column=1, sticky="nsew", padx=(5, 10), pady=10)
var_dup = ctk.BooleanVar(value=DUPLICATE_ENABLED)
chk_dup = ctk.CTkCheckBox(frame_dup, text="Enable duplicate-check", variable=var_dup)
chk_dup.pack(anchor="w", padx=10, pady=(10, 5))
ctk.CTkLabel(frame_dup, text="Window (days):").pack(anchor="w", padx=10, pady=5)
var_days = ctk.IntVar(value=DUPLICATE_DAYS)
spin_days = ctk.CTkEntry(frame_dup, textvariable=var_days, width=120)
spin_days.pack(anchor="w", padx=10, pady=(0, 10))

# --- GUI behavior functions ---
def refresh_tree_from_excel(path: Path):
    for r in tree.get_children():
        tree.delete(r)
    if not path.exists():
        return
    rows = read_all_rows(str(path))
    for r in rows:
        tree.insert("", "end", values=r[:len(cols)])

def on_search_change(*_):
    q = search_var.get().strip().lower()
    for i in tree.get_children():
        tree.detach(i)
    if not q:
        for i in tree.get_children():
             tree.reattach(i, "", "end")
        return

    all_items = tree.get_children()
    for item in all_items:
        values = tree.item(item, "values")
        combined = " ".join([str(v).lower() for v in values])
        if q in combined:
            tree.reattach(item, "", "end")

entry_search.bind("<KeyRelease>", on_search_change)

def create_new_excel_dialog():
    cols_default = ",".join(cols + ["ResumePath"])
    dlg = ctk.CTkToplevel(root)
    dlg.title("Create New Excel")
    dlg.geometry("480x250")
    dlg.transient(root)
    dlg.grab_set()

    ctk.CTkLabel(dlg, text="Enter comma-separated column names:").pack(pady=(10, 6))
    txt = ctk.CTkTextbox(dlg, height=120)
    txt.pack(fill="both", padx=12, pady=6)
    txt.insert("1.0", cols_default)

    def do_create():
        cols_input = [c.strip() for c in txt.get("1.0", "end-1c").split(",") if c.strip()]
        file = filedialog.asksaveasfilename(initialdir=str(EXCEL_DIR),
                                            defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if not file: return
        ensure_excel(file, cols_input)
        set_active_excel(Path(file))
        refresh_tree_from_excel(ACTIVE_EXCEL)
        add_status(f"Created and selected: {Path(file).name}")
        dlg.destroy()
    ctk.CTkButton(dlg, text="Create", command=do_create).pack(pady=8)

def select_existing_excel():
    file = filedialog.askopenfilename(initialdir=str(EXCEL_DIR), filetypes=[("Excel files", "*.xlsx")])
    if not file: return
    set_active_excel(Path(file))
    refresh_tree_from_excel(ACTIVE_EXCEL)
    add_status(f"Selected: {Path(file).name}")

def on_change_excel():
    resp = messagebox.askquestion("Choose Action", "Create a new Excel file?", detail="Select 'Yes' to create a new file, or 'No' to select an existing one.")
    if resp == "yes":
        create_new_excel_dialog()
    elif resp == "no":
        select_existing_excel()

btn_change_excel.configure(command=on_change_excel)

def process_files_sequential(file_paths):
    ui_call(lambda: [w.configure(state="disabled") for w in [btn_single, btn_multi, btn_change_excel, btn_open_excel, btn_export, btn_edit_columns]])
    total = len(file_paths)
    ui_call(lambda: progress.set(0))
    results = []

    for i, fp in enumerate(file_paths):
        fname = Path(fp).name
        try:
            add_status(f"Processing {fname}...")
            data = parse_resume(fp, nlp)
            if not data:
                add_status(f"Failed: {fname} (no text, scanned PDF?)")
                results.append((fname, "Failed"))
                continue

            is_duplicate = var_dup.get() and email_duplicate_within_days(str(ACTIVE_EXCEL), data.get("Email", ""), int(var_days.get()))
            if is_duplicate:
                add_status(f"Skipped: {fname} (duplicate)")
                results.append((fname, "Skipped"))
            else:
                ensure_excel(str(ACTIVE_EXCEL))
                append_row(str(ACTIVE_EXCEL), data)
                ui_call(tree.insert, "", "end", values=[data.get(c, "") for c in cols])
                add_status(f"Added: {fname}")
                results.append((fname, "Added"))
        except Exception as e:
            log_error(f"Processing error for {fp}: {e}\n{traceback.format_exc()}")
            add_status(f"Error: {fname} ({e})")
            results.append((fname, "Error"))
        finally:
            ui_call(lambda p=(i + 1) / total: progress.set(p))

    ui_call(lambda: [w.configure(state="normal") for w in [btn_single, btn_multi, btn_change_excel, btn_open_excel, btn_export, btn_edit_columns]])
    added = sum(1 for r in results if r[1] == "Added")
    skipped = sum(1 for r in results if r[1] == "Skipped")
    failed = sum(1 for r in results if r[1] in ("Failed", "Error"))
    ui_call(messagebox.showinfo, "Batch Complete", f"Processed {total} files: {added} added, {skipped} skipped, {failed} failed.")

def safe_copy_to_workspace(src_path: str):
    src = Path(src_path)
    if ACTIVE_EXCEL.exists() and src.resolve() == ACTIVE_EXCEL.resolve():
        messagebox.showwarning("Wrong File", "Cannot process the active Excel file. Please select a resume.")
        return None
    dst = RESUMES_DIR / src.name
    i = 1
    while dst.exists():
        dst = RESUMES_DIR / f"{src.stem}_{i}{src.suffix}"
        i += 1
    copy2(src, dst)
    return dst

def on_upload(multiple=False):
    if multiple:
        fps = filedialog.askopenfilenames(initialdir=str(RESUMES_DIR), filetypes=[("Resumes", "*.pdf;*.docx"), ("All files", "*.*")])
    else:
        fp = filedialog.askopenfilename(initialdir=str(RESUMES_DIR), filetypes=[("Resumes", "*.pdf;*.docx"), ("All files", "*.*")])
        fps = [fp] if fp else []

    if not fps: return

    saved = [str(safe_copy_to_workspace(f)) for f in fps if f]
    if saved:
        threading.Thread(target=process_files_sequential, args=(saved,), daemon=True).start()

btn_single.configure(command=lambda: on_upload(False))
btn_multi.configure(command=lambda: on_upload(True))

def edit_columns_dialog():
    if not ACTIVE_EXCEL or not ACTIVE_EXCEL.exists():
        messagebox.showwarning("No Excel", "Create or select an Excel file first.")
        return

    headers = get_headers(str(ACTIVE_EXCEL)) or cols + ["ResumePath"]

    dlg = ctk.CTkToplevel(root)
    dlg.title("Edit Excel Columns")
    dlg.transient(root)
    dlg.grab_set()
    dlg.geometry("600x500")

    ctk.CTkLabel(dlg, text=f"Editing columns for: {ACTIVE_EXCEL.name}", font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=10, pady=8)

    frame = ctk.CTkFrame(dlg)
    frame.pack(fill="both", expand=True, padx=10, pady=6)

    list_frame = ctk.CTkScrollableFrame(frame, label_text="Columns")
    list_frame.pack(side="left", fill="both", expand=True, padx=(0, 6))

    col_entries = [ctk.CTkEntry(list_frame, placeholder_text=h) for h in headers]
    for entry in col_entries:
        entry.pack(fill="x", padx=5, pady=2)

    ctrl = ctk.CTkFrame(frame)
    ctrl.pack(side="left", fill="y", padx=(6, 0))

    def get_current_headers():
        return [entry.get() or entry.cget("placeholder_text") for entry in col_entries]

    def add_col():
        entry = ctk.CTkEntry(list_frame, placeholder_text="New Column")
        entry.pack(fill="x", padx=5, pady=2)
        col_entries.append(entry)

    def remove_col():
        if col_entries:
            col_entries.pop().destroy()

    ctk.CTkButton(ctrl, text="Add Column", command=add_col).pack(pady=4, padx=10, fill="x")
    ctk.CTkButton(ctrl, text="Remove Last", command=remove_col).pack(pady=4, padx=10, fill="x")

    def do_save():
        new_headers = get_current_headers()
        if not new_headers:
            if not messagebox.askyesno("Empty Headers", "Save with no columns?"): return

        if len(new_headers) < len(headers):
            if not messagebox.askyesno("Truncate Columns", "Fewer columns may cause data loss. Proceed?"): return

        try:
            add_status(f"Updating headers for {ACTIVE_EXCEL.name}...")
            update_headers(str(ACTIVE_EXCEL), new_headers)
            refresh_tree_from_excel(ACTIVE_EXCEL)
            add_status("Headers updated.")
            dlg.destroy()
        except PermissionError as pe:
            messagebox.askretrycancel("File Locked", f"Cannot write to {ACTIVE_EXCEL.name} as it is open.\nPlease close it and retry.\n\nError: {pe}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to update headers: {e}")
            add_status(f"Failed to update headers: {e}")

    bottom = ctk.CTkFrame(dlg)
    bottom.pack(fill="x", padx=10, pady=8)
    ctk.CTkButton(bottom, text="Save", command=do_save).pack(side="right")
    ctk.CTkButton(bottom, text="Cancel", command=dlg.destroy).pack(side="right", padx=10)

btn_edit_columns.configure(command=edit_columns_dialog)

def export_visible_to_csv():
    save_path = filedialog.asksaveasfilename(initialdir=str(WORKSPACE), defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
    if not save_path: return
    try:
        with open(save_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(cols)
            for item in tree.get_children():
                writer.writerow(tree.item(item, "values"))
        messagebox.showinfo("Export Complete", f"Exported {len(tree.get_children())} rows to {Path(save_path).name}")
    except Exception as e:
        messagebox.showerror("Export Error", str(e))

btn_export.configure(command=export_visible_to_csv)

# --- Final initialization ---
if not ACTIVE_EXCEL.exists():
    ensure_excel(str(ACTIVE_EXCEL))
set_active_excel(ACTIVE_EXCEL)
refresh_tree_from_excel(ACTIVE_EXCEL)
add_status(f"Ready. Workspace: {WORKSPACE}")

root.mainloop()