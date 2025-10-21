import os
import sys
import shutil
import getpass
import xlwings as xw
import tkinter as tk
import time  # <-- IMPORT ADDED for the pause
from pathlib import Path
from datetime import datetime
from tkinter import filedialog, messagebox

# --- Helper ---
def resource_path(relative_path: str) -> Path:
    """ Get absolute path to resource, works for dev and for PyInstaller bundle. """
    try:
        base_path = Path(sys._MEIPASS)
    except Exception:
        base_path = Path(".").resolve()
    return base_path / relative_path

# --- Core Process ---
def run_process(template_path, source_path, output_dir, runner_path):
    """
    Hybrid Workflow:
    1. Python copies the template to the final destination.
    2. Python launches a macro in Runner.xlsm, passing the file paths.
    3. VBA performs all Excel operations natively.
    """
    username = getpass.getuser()
    ts = datetime.now().strftime("%Y-%m-%d_%H%M")
    final_name = f"{username}_{ts}.xlsx"
    final_output_path = Path(output_dir) / final_name

    try:
        shutil.copy(template_path, final_output_path)
    except Exception as e:
        raise IOError(f"Failed to copy template file. Check permissions.\n{e}")

    source_path_abs = Path(source_path).resolve()
    output_path_abs = final_output_path.resolve()
    
    app = None
    try:
        app = xw.App(visible=False)
        app.display_alerts = False
        
        runner_wb = app.books.open(str(runner_path))
        
        time.sleep(2)

        macro_to_run = runner_wb.macro("VBA_Module.RunTheProcess")
        macro_to_run(str(source_path_abs), str(output_path_abs))
        
        runner_wb.close()

    finally:
        if app:
            try:
                app.quit()
            except Exception:
                app.kill()

    return str(final_output_path), ts

# ---------------- Tk UI (No changes below this line) ----------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Hybrid Excel Merge Runner v4.2")
        self.geometry("620x250")
        self.minsize(620, 250)
        self.resizable(True, False)
        self.template_path = tk.StringVar()
        self.source_path = tk.StringVar()
        self.output_dir = tk.StringVar()
        self.template_ok = tk.BooleanVar(value=False)
        self.source_ok = tk.BooleanVar(value=False)
        self.output_ok = tk.BooleanVar(value=False)
        self._build_ui()

    def _build_ui(self):
        pad = 10
        tk.Label(self, text="Select Template, Source, and Output Folder, then click Run").pack(pady=(12, 6))
        frm = tk.Frame(self)
        frm.pack(fill="x", expand=True, padx=pad, pady=(0, 8))
        self._row_with_button(frm, 0, "Select TEMPLATE workbook", self.template_path, self.template_ok, self.pick_template)
        self._row_with_button(frm, 1, "Select SOURCE workbook", self.source_path, self.source_ok, self.pick_source)
        self._row_with_button(frm, 2, "Select OUTPUT folder", self.output_dir, self.output_ok, self.pick_output)
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
        var_ok.trace_add("write", lambda *_: lbl_tick.config(text="✓" if var_ok.get() else "✗", fg=("green" if var_ok.get() else "red")))
        ent = tk.Entry(parent, textvariable=var_path, state="readonly")
        ent.grid(row=row, column=2, padx=(8, 0), sticky="ew")
        parent.grid_columnconfigure(2, weight=1)

    def pick_template(self):
        path = filedialog.askopenfilename(title="Select TEMPLATE workbook", filetypes=(("Excel files", "*.xlsx;*.xlsm;*.xlsb"), ("All files", "*.*")))
        if path: self.template_path.set(path); self.template_ok.set(True)
        self._update_run_state()

    def pick_source(self):
        path = filedialog.askopenfilename(title="Select SOURCE workbook (data)", filetypes=(("Excel files", "*.xlsx;*.xlsm;*.xlsb"), ("All files", "*.*")))
        if path: self.source_path.set(path); self.source_ok.set(True)
        self._update_run_state()

    def pick_output(self):
        path = filedialog.askdirectory(title="Select OUTPUT folder")
        if path: self.output_dir.set(path); self.output_ok.set(True)
        self._update_run_state()

    def _update_run_state(self):
        ready = self.template_ok.get() and self.source_ok.get() and self.output_ok.get()
        self.run_btn.config(state=("normal" if ready else "disabled"))

    def on_run(self):
        runner_path = resource_path("Runner.xlsm")
        if not runner_path.exists():
            messagebox.showerror("Error", f"Could not find the 'Runner.xlsm' file.\nPlease ensure it is in the same directory as this application.")
            return

        self.run_btn.config(state="disabled")
        self.open_btn.config(state="disabled")
        self.status_lbl.config(text="Running... Excel is processing in the background.", fg="gray")
        self.update_idletasks()
        try:
            out_path, _ = run_process(self.template_path.get(), self.source_path.get(), self.output_dir.get(), runner_path)
            self._last_output_dir = self.output_dir.get()
            self.status_lbl.config(text=f"Done: {out_path}", fg="green")
            self.open_btn.config(state="normal")
        except Exception as e:
            self.status_lbl.config(text=f"Error: {e}", fg="red")
            messagebox.showerror("Error", f"An unexpected error occurred:\n\n{e}")
        finally:
            self._update_run_state()

    def on_open_folder(self):
        folder_path = getattr(self, "_last_output_dir", self.output_dir.get())
        if not folder_path: return
        try:
            if os.name == "nt": os.startfile(folder_path)
            elif sys.platform == "darwin": import subprocess; subprocess.Popen(["open", folder_path])
            else: import subprocess; subprocess.Popen(["xdg-open", folder_path])
        except Exception as e:
            messagebox.showwarning("Open folder", f"Failed to open folder:\n{e}")

if __name__ == "__main__":
    app = App()
    try:
        app.iconbitmap(resource_path("app.ico"))
    except:
        pass
    app.mainloop()
