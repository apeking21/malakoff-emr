import os
import sys
import math
import getpass
import xlwings as xw
import tkinter as tk
from decimal import Decimal, InvalidOperation
from pathlib import Path
from datetime import datetime
from tkinter import filedialog, messagebox

# ---------------- Configuration ----------------
START_ROW = 6
INSERT_ROWS = True
APPLY_ROW5_FMT = True
CLEAR_STATIC_FILL = True

# Columns where SOURCE numbers should be written as strings with a trailing '%'
# so Excel parses them as percentages (e.g., 15 -> "15%").
# Use column letters or 1-based indices per sheet.
PERCENT_STRING_COLUMNS = {
    "LineItemsTaxes": ["E"],
    "LineItemsDiscounts": ["D"],
    "LineItemsCharges": ["D"],
}

DECIMAL_PRECISION_COLUMNS = {
    "Documents": {"AM": 5, "AN": 5, "AO": 5, "AP": 5, "AQ": 5, "AR": 5, "AS": 5, "AT": 5},
}

# BRUTE-FORCE left-pad with zeros to reach a required total length.
# Mapping: sheet_name -> { column_letter_or_1based_index : expected_length }
# Example:
# PAD_WITH_LEADING_ZEROS = {
#     "Documents": {"B": 2, "E": 5},
#     "LineItemsCharges": {2: 3},  # column 2 (B) -> total length 3
# }
PAD_WITH_LEADING_ZEROS = {
    "Documents": {"C": 2},
    "DocumentLineItems": {"D": 3},
    "LineItemsAddClassifications": {"D": 3},
    "LineItemsTaxes": {"D": 2},
    "DocumentTotalTax": {"C": 2}
}

# Extend named ranges after paste (sheet, name, column, start_row)
NAMED_RANGES_TO_EXTEND = [
    ("Documents", "InvoiceNumbersList", "B", 5),
    ("DocumentLineItems", "InvoiceItemsList", "C", 5),
]

# ---------------- Constants ----------------
XL_PASTE_FORMATS = -4122
XL_PASTE_VALIDATION = 6

# ---------------- Helpers (PyInstaller resources) ----------------
def resource_path(relative_path: str) -> Path:
    """Get absolute path to resource, works for dev and PyInstaller bundle."""
    if hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS) / relative_path
    return Path(relative_path)

# ---------------- Excel helpers ----------------
def last_row_col(sheet):
    ur = sheet.used_range
    return ur.last_cell.row, ur.last_cell.column

def _is_int_like(x, tol=1e-9):
    if isinstance(x, bool):
        return False
    if isinstance(x, int):
        return True
    if isinstance(x, float) and math.isfinite(x):
        return abs(x - round(x)) <= tol
    return False

def _digits_count_of_number(x):
    # assumes int-like
    try:
        n = int(round(float(x)))
    except Exception:
        try:
            n = int(Decimal(str(x)))
        except Exception:
            return 0
    return len(str(abs(n)))

def to_plain_string(v):
    """
    Convert any value to a non-scientific string.
    - Preserves strings as-is (leading zeros kept).
    - Numbers use Decimal to avoid scientific notation.
    - Integers -> no decimal point; non-integers trimmed of trailing zeros.
    """
    if v is None:
        return ""
    if isinstance(v, str):
        return v
    try:
        d = Decimal(str(v))
    except (InvalidOperation, ValueError, TypeError):
        return str(v)
    if d == d.to_integral_value():
        return str(d.quantize(Decimal("1")))
    s = format(d, "f")
    return s.rstrip("0").rstrip(".")

def pad_left_zeros(s, expected_len):
    """Left-pad with zeros to expected_len."""
    if s is None:
        return ""
    s = str(s)
    if expected_len and expected_len > 0 and len(s) < expected_len:
        return s.zfill(expected_len)
    return s

def _col_letter_to_index(c):
    # "A"->1, "Z"->26, "AA"->27 ...
    c = c.strip().upper()
    val = 0
    for ch in c:
        if 'A' <= ch <= 'Z':
            val = val * 26 + (ord(ch) - ord('A') + 1)
        else:
            raise ValueError(f"Invalid column letter: {c}")
    return val

def _columns_set_for_sheet(mapping, sheet_name, write_cols):
    """Return 0-based column indices from a {sheet: [letters|indices]} mapping."""
    cols = mapping.get(sheet_name, [])
    out = set()
    for item in cols:
        try:
            if isinstance(item, str):
                j1 = _col_letter_to_index(item)  # 1-based
            else:
                j1 = int(item)                   # 1-based
            if 1 <= j1 <= write_cols:
                out.add(j1 - 1)                  # 0-based
        except Exception:
            pass
    return out

def _percent_string_cols_for_sheet(sheet_name, write_cols):
    return _columns_set_for_sheet(PERCENT_STRING_COLUMNS, sheet_name, write_cols)

def _pad_columns_map_for_sheet(sheet_name, write_cols):
    """
    From PAD_WITH_LEADING_ZEROS[sheet] -> {col_spec: expected_len}
    Return dict { zero_based_col_index : expected_len }
    """
    mapping = PAD_WITH_LEADING_ZEROS.get(sheet_name, {})
    out = {}
    for col_spec, length in mapping.items():
        try:
            j1 = _col_letter_to_index(col_spec) if isinstance(col_spec, str) else int(col_spec)
            if 1 <= j1 <= write_cols:
                out[j1 - 1] = int(length)
        except Exception:
            pass
    return out

def _is_blank_matrix(matrix):
    """Return True if matrix has no real values (all None/empty strings)."""
    if matrix is None:
        return True
    if not isinstance(matrix, list):
        return matrix in (None, "")
    if matrix and not isinstance(matrix[0], list):
        row = matrix
        return all(v in (None, "") for v in row)
    for row in matrix:
        if not isinstance(row, list):
            row = [row]
        for v in row:
            if v not in (None, ""):
                return False
    return True

def extend_named_range_to_col(wb, sheet_name, range_name, col_letter, start_row):
    ws = wb.sheets[sheet_name]
    last_row_in_sheet = ws.cells.last_cell.row
    probe = ws.range(f"{col_letter}{last_row_in_sheet}").end('up').row
    last_row = max(probe, start_row)
    refers_to = f"='{sheet_name}'!${col_letter}${start_row}:${col_letter}${last_row}"

    # Workbook-scoped preferred
    try:
        nm = wb.names[range_name]
        nm.refers_to = refers_to
        return
    except KeyError:
        pass
    # Sheet-scoped fallback
    try:
        nm = ws.names[range_name]
        nm.refers_to = refers_to
        return
    except KeyError:
        wb.names.add(name=range_name, refers_to=refers_to)

# ---------------- Core copy ----------------
def copy_rows_from6(src_sheet, dst_sheet, sheet_name,
                    start_row=START_ROW, insert_rows=INSERT_ROWS,
                    apply_row5_fmt=APPLY_ROW5_FMT, clear_static_fill=CLEAR_STATIC_FILL):
    # 1) Bounds & read source data
    src_last_row, src_last_col = last_row_col(src_sheet)
    if src_last_row < start_row:
        return 0, 0

    src_data_block = src_sheet.range((start_row, 1), (src_last_row, src_last_col))
    data = src_data_block.value
    write_rows = src_last_row - start_row + 1
    write_cols = src_last_col

    # Normalize to 2D
    if write_rows == 1 and write_cols == 1:
        data = [[data]]
    elif write_rows == 1:
        data = [data]
    elif write_cols == 1:
        data = [[r] for r in data]

    # Nothing real? bail out
    if _is_blank_matrix(data):
        return 0, 0

    # Grab Excel displayed text for fallback (preserves literal text/leading zeros)
    displayed = [[None for _ in range(write_cols)] for _ in range(write_rows)]
    for i in range(write_rows):
        for j in range(write_cols):
            try:
                displayed[i][j] = src_sheet.range((start_row + i, 1 + j)).api.Text
            except Exception:
                displayed[i][j] = None

    # 2) Prepare destination area
    if insert_rows and write_rows > 0:
        dst_sheet.api.Rows(f"{start_row}:{start_row + write_rows - 1}").Insert()

    dest_block = dst_sheet.range((start_row, 1), (start_row + write_rows - 1, write_cols))

    # Copy row 5 validations + formats and enforce alignment
    if apply_row5_fmt and write_cols > 0:
        row5_src = dst_sheet.range((5, 1), (5, write_cols))
        row5_src.api.Copy()
        dest_block.api.PasteSpecial(Paste=XL_PASTE_VALIDATION)
        row5_src.api.Copy()
        dest_block.api.PasteSpecial(Paste=XL_PASTE_FORMATS)
        try:
            dest_block.api.HorizontalAlignment = row5_src.api.HorizontalAlignment
            dest_block.api.VerticalAlignment = row5_src.api.VerticalAlignment
            dest_block.api.WrapText = row5_src.api.WrapText
            dest_block.api.IndentLevel = row5_src.api.IndentLevel
        except Exception:
            pass
        if clear_static_fill:
            dest_block.api.Interior.Pattern = 0  # xlPatternNone

    # --- Special column categories ---
    percent_string_cols = _percent_string_cols_for_sheet(sheet_name, write_cols)
    pad_cols_map = _pad_columns_map_for_sheet(sheet_name, write_cols)  # {col_index_0based: expected_len}

    # New: precision columns (e.g., 5dp)
    precision_cols = {}
    prec_def = DECIMAL_PRECISION_COLUMNS.get(sheet_name, {})
    for col_spec, dp in prec_def.items():
        try:
            j1 = _col_letter_to_index(col_spec) if isinstance(col_spec, str) else int(col_spec)
            if 1 <= j1 <= write_cols:
                precision_cols[j1 - 1] = int(dp)
        except Exception:
            pass

    # Helper to analyze a column for numeric/text safety
    def _analyze_column(j):
        has_text_leading_zero = False
        has_very_long_int = False
        any_decimal = False
        all_numeric = True

        for i in range(write_rows):
            v = data[i][j]
            if v in (None, ""):
                continue
            disp = displayed[i][j]
            if isinstance(disp, str) and len(disp) > 1 and disp.lstrip().startswith("0") and disp.strip().isdigit():
                has_text_leading_zero = True
            if _is_int_like(v):
                if _digits_count_of_number(v) > 15:
                    has_very_long_int = True
            elif isinstance(v, (int, float)):
                any_decimal = True
            else:
                all_numeric = False

        return {
            "text_leading_zero": has_text_leading_zero,
            "very_long_int": has_very_long_int,
            "any_decimal": any_decimal,
            "all_numeric": all_numeric,
        }

    # --- Format destination columns ---
    for j in range(write_cols):
        if j in percent_string_cols:
            continue
        if j in pad_cols_map:
            # padded columns → text
            dest_block.columns[j].api.NumberFormat = "@"
            continue
        if j in precision_cols:
            # decimal precision override
            dp = precision_cols[j]
            fmt = f"0.{''.join(['0' for _ in range(dp)])}"
            dest_block.columns[j].api.NumberFormat = fmt
            continue

        info = _analyze_column(j)
        if info["text_leading_zero"] or info["very_long_int"] or not info["all_numeric"]:
            dest_block.columns[j].api.NumberFormat = "@"
        else:
            dest_block.columns[j].api.NumberFormat = "0.####################" if info["any_decimal"] else "0"

    # --- Transform and write ---
    for j in range(write_cols):
        if j in percent_string_cols:
            for i in range(write_rows):
                v = data[i][j]
                if v in (None, ""):
                    continue
                s = None
                if isinstance(v, (int, float)):
                    s = f"{v}%"
                elif isinstance(v, str):
                    v_str = v.strip()
                    if v_str.endswith("%"):
                        s = v_str
                    else:
                        try:
                            float(v_str)
                            s = f"{v_str}%"
                        except Exception:
                            s = v
                if s is not None:
                    data[i][j] = s
            continue

        if j in pad_cols_map:
            exp_len = pad_cols_map[j]
            for i in range(write_rows):
                raw = data[i][j]
                disp = displayed[i][j]
                s = (str(disp) if disp not in (None, "") else to_plain_string(raw))
                data[i][j] = pad_left_zeros(s, exp_len)
            continue

        if j in precision_cols:
            dp = precision_cols[j]
            for i in range(write_rows):
                v = data[i][j]
                if isinstance(v, (int, float)):
                    data[i][j] = round(v, dp)
                else:
                    try:
                        data[i][j] = round(float(v), dp)
                    except Exception:
                        pass
            continue

        info = _analyze_column(j)
        if info["text_leading_zero"] or info["very_long_int"] or not info["all_numeric"]:
            for i in range(write_rows):
                disp = displayed[i][j]
                s = (str(disp) if disp not in (None, "") else to_plain_string(data[i][j]))
                data[i][j] = s
        else:
            for i in range(write_rows):
                v = data[i][j]
                if isinstance(v, str):
                    v_str = v.strip()
                    try:
                        if "." in v_str:
                            data[i][j] = float(v_str)
                        else:
                            data[i][j] = int(v_str)
                    except Exception:
                        data[i][j] = v

    dest_block.value = data
    return write_rows, write_cols

def unblock_file_ntfs(pathlike):
    """Remove NTFS Mark-of-the-Web (Zone.Identifier) if present."""
    try:
        p = str(pathlike)
        zi = f"{p}:Zone.Identifier"
        if os.path.exists(zi):
            os.remove(zi)
    except Exception:
        pass

def run_process(template_path, source_path, output_dir):
    username = getpass.getuser()
    # Portal-friendly name: USER_HHMM_DDMMYYYY.xlsx
    ts = datetime.now().strftime("%H%M_%d%m%Y")
    final_name = f"{username}_{ts}.xlsx"
    final_output_path = Path(output_dir) / final_name

    template_path = Path(template_path).resolve()
    source_path = Path(source_path).resolve()

    app = xw.App(visible=False)
    app.display_alerts = False
    app.screen_updating = False
    app.calculation = 'manual'

    try:
        # Open TEMPLATE (no copying), modify in-memory, then SaveAs to output
        wb_template = app.books.open(str(template_path), update_links=False, read_only=False)
        wb_source   = app.books.open(str(source_path), update_links=False, read_only=True)

        src_names = {s.name for s in wb_source.sheets}
        dst_names = {s.name for s in wb_template.sheets}
        common = sorted(src_names & dst_names)

        for name in common:
            copy_rows_from6(
                src_sheet=wb_source.sheets[name],
                dst_sheet=wb_template.sheets[name],
                sheet_name=name,
                start_row=START_ROW,
                insert_rows=INSERT_ROWS,
                apply_row5_fmt=APPLY_ROW5_FMT,
                clear_static_fill=CLEAR_STATIC_FILL
            )

        # Extend named ranges after paste
        for sheet_name, range_name, col_letter, start_r in NAMED_RANGES_TO_EXTEND:
            try:
                extend_named_range_to_col(
                    wb=wb_template,
                    sheet_name=sheet_name,
                    range_name=range_name,
                    col_letter=col_letter,
                    start_row=start_r
                )
            except Exception:
                pass

        # Recalc & Save AS to output (template file remains unchanged on disk)
        app.calculation = 'automatic'
        try:
            wb_template.api.CalculateFullRebuild()
        except Exception:
            try:
                wb_template.api.CalculateFull()
            except Exception:
                pass

        # FileFormat=51 => .xlsx macro-free (adjust if your template is .xlsm)
        wb_template.api.SaveAs(Filename=str(final_output_path), FileFormat=51)
        unblock_file_ntfs(final_output_path)

    finally:
        try:
            wb_source.close(save=False)
        except Exception:
            pass
        try:
            wb_template.close(save=False)  # don't overwrite the template
        except Exception:
            pass
        try:
            app.quit()
            app.kill()
        except Exception:
            pass

    return str(final_output_path), ts

# ---------------- Tk UI ----------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Excel Merge Runner v3.3")
        # self.apply_window_icon("app.ico")
        self.geometry("620x250")           # initial size
        self.minsize(620, 250)             # min size
        self.resizable(True, False)        # horizontal resize only (path fields stretch)

        self.template_path = tk.StringVar()
        self.source_path = tk.StringVar()
        self.output_dir = tk.StringVar()

        self.template_ok = tk.BooleanVar(value=False)
        self.source_ok = tk.BooleanVar(value=False)
        self.output_ok = tk.BooleanVar(value=False)

        self._build_ui()

    def apply_window_icon(self, icon_filename: str = "app.ico"):
        """Try to set both window and taskbar icons, works in PyInstaller EXEs too."""
        try:
            icon_path = resource_path(icon_filename)
            if os.name == "nt":
                # Set for window title and taskbar
                self.iconbitmap(str(icon_path))
                try:
                    import ctypes
                    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("ExcelMergeRunner")
                except Exception:
                    pass
            else:
                self.iconbitmap(str(icon_path))
        except Exception as e:
            print(f"Icon not set: {e}")

    def _build_ui(self):
        pad = 10
        tk.Label(self, text="Select Template, Source, and Output Folder, then click Run").pack(pady=(12, 6))

        frm = tk.Frame(self)
        frm.pack(fill="x", expand=True, padx=pad, pady=(0, 8))  # expand so entries can stretch

        self._row_with_button(frm, row=0, btn_text="Select TEMPLATE workbook",
                              var_path=self.template_path, var_ok=self.template_ok, cmd=self.pick_template)
        self._row_with_button(frm, row=1, btn_text="Select SOURCE workbook",
                              var_path=self.source_path, var_ok=self.source_ok, cmd=self.pick_source)
        self._row_with_button(frm, row=2, btn_text="Select OUTPUT folder",
                              var_path=self.output_dir, var_ok=self.output_ok, cmd=self.pick_output)

        btn_frame = tk.Frame(self)
        btn_frame.pack(pady=(8, 6))
        self.run_btn = tk.Button(btn_frame, text="Run", width=18, state="disabled", command=self.on_run)
        self.run_btn.pack(side="left", padx=8)
        self.open_btn = tk.Button(btn_frame, text="Open folder", width=18, state="disabled", command=self.on_open_folder)
        self.open_btn.pack(side="left", padx=8)

        self.status_lbl = tk.Label(self, text="", fg="gray", wraplength=580, justify="center")
        self.status_lbl.pack(pady=(2, 6))

        self.after(200, self._update_run_state)

    def _row_with_button(self, parent, row, btn_text, var_path, var_ok, cmd):
        btn = tk.Button(parent, text=btn_text, width=28, command=cmd)
        btn.grid(row=row, column=0, padx=(0, 8), pady=6, sticky="w")

        lbl_tick = tk.Label(parent, text="✗", fg="red", width=2, anchor="w")
        lbl_tick.grid(row=row, column=1, sticky="w")
        var_ok.trace_add(
            "write",
            lambda *_: lbl_tick.config(text="✓" if var_ok.get() else "✗",
                                       fg=("green" if var_ok.get() else "red"))
        )

        ent = tk.Entry(parent, textvariable=var_path, state="readonly")
        ent.grid(row=row, column=2, padx=(8, 0), sticky="we")  # <-- stretch with window
        parent.grid_columnconfigure(2, weight=1)               # <-- column 2 grows

    def pick_template(self):
        path = filedialog.askopenfilename(
            title="Select TEMPLATE workbook",
            filetypes=(("Excel files", "*.xlsx;*.xlsm;*.xlsb"), ("All files", "*.*"))
        )
        if path:
            self.template_path.set(path)
            self.template_ok.set(True)
        self._update_run_state()

    def pick_source(self):
        path = filedialog.askopenfilename(
            title="Select SOURCE workbook (data)",
            filetypes=(("Excel files", "*.xlsx;*.xlsm;*.xlsb"), ("All files", "*.*"))
        )
        if path:
            self.source_path.set(path)
            self.source_ok.set(True)
        self._update_run_state()

    def pick_output(self):
        path = filedialog.askdirectory(title="Select OUTPUT folder")
        if path:
            self.output_dir.set(path)
            self.output_ok.set(True)
        self._update_run_state()

    def _update_run_state(self):
        ready = self.template_ok.get() and self.source_ok.get() and self.output_ok.get()
        self.run_btn.config(state=("normal" if ready else "disabled"))

    def on_run(self):
        self.run_btn.config(state="disabled")
        self.open_btn.config(state="disabled")
        self.status_lbl.config(text="Running... please wait.", fg="gray")
        self.update_idletasks()

        try:
            out_path, ts = run_process(
                template_path=self.template_path.get(),
                source_path=self.source_path.get(),
                output_dir=self.output_dir.get()
            )
            self._last_output_dir = self.output_dir.get()
            self.status_lbl.config(text=f"Done: {out_path}", fg="green")
            self.open_btn.config(state="normal")
        except Exception as e:
            self.status_lbl.config(text=f"Error: {e}", fg="red")
        finally:
            self._update_run_state()

    def on_open_folder(self):
        folder_path = getattr(self, "_last_output_dir", self.output_dir.get())
        if not folder_path:
            return
        try:
            if os.name == "nt":
                os.startfile(folder_path)  # type: ignore
            elif sys.platform == "darwin":
                import subprocess; subprocess.Popen(["open", folder_path])
            else:
                import subprocess; subprocess.Popen(["xdg-open", folder_path])
        except Exception as e:
            messagebox.showwarning("Open folder", f"Failed to open folder:\n{e}")

# ---------------- Entry point ----------------
if __name__ == "__main__":
    app = App()
    app.iconbitmap(resource_path("app.ico"))
    app.mainloop()