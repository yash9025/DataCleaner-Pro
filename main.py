import sys
import os
import argparse
import threading
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# Ensure src/ is importable
current_dir = Path(__file__).parent.resolve()
src_dir = current_dir / "src"
if str(src_dir) not in sys.path:
    sys.path.insert(0, str(src_dir))

from cleaner import DataCleaner
from reporter import ExcelReportGenerator

def run_pipeline(input_file, output_file=None, log_callback=None):
    """Executes the cleaning and reporting pipeline."""
    input_path = Path(input_file)
    if not output_file:
        output_dir = input_path.parent / "reports"
        output_dir.mkdir(parents=True, exist_ok=True)
        output_file = output_dir / f"{input_path.stem}_Cleaned_Report.xlsx"
    else:
        output_file = Path(output_file)
        output_file.parent.mkdir(parents=True, exist_ok=True)

    def log(msg):
        if log_callback:
            log_callback(msg)
        else:
            # Safe print for Windows consoles
            try:
                print(msg)
            except UnicodeEncodeError:
                print(msg.encode('ascii', 'replace').decode('ascii'))

    log(f"[LOAD] Ingesting raw dataset: {input_path.name}...")
    cleaner = DataCleaner(input_path)
    cleaner.load_data()
    log(f"   -> Raw records: {cleaner.audit_log['raw_rows']:,} rows across {cleaner.audit_log['raw_columns']} columns.")

    log("[CLEAN] Applying cleaning rules (casing, whitespace, dates, currency, dedup)...")
    cleaned_df, audit = cleaner.clean()
    log(f"   -> Duplicates eliminated: {audit['duplicates_removed']} rows.")
    log(f"   -> Missing values treated: {audit['missing_values_before']} -> {audit['missing_values_after']}.")
    log(f"   -> Retained records: {audit['final_rows']:,} rows in {audit['processing_time_sec']}s.")

    log("[REPORT] Designing multi-tab Executive Excel Report with openpyxl...")
    reporter = ExcelReportGenerator(cleaned_df, audit, output_path=output_file)
    final_path = reporter.save()
    log(f"[DONE] Automation pipeline completed successfully!")
    log(f"   -> Report generated at: {final_path.resolve()}")
    return final_path, audit

class ModernCleanerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("DataCleaner Pro - Executive Report Automation")
        self.root.geometry("680x560")
        self.root.minsize(620, 500)
        self.root.configure(bg="#F3F4F6")

        self.selected_file = None
        self.generated_report = None

        self._setup_styles()
        self._build_ui()

    def _setup_styles(self):
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Header.TLabel", background="#1F4E79", foreground="#FFFFFF", font=("Segoe UI", 14, "bold"))
        style.configure("SubHeader.TLabel", background="#1F4E79", foreground="#D0E1FD", font=("Segoe UI", 9))
        style.configure("Card.TFrame", background="#FFFFFF", relief="flat")
        style.configure("Primary.TButton", font=("Segoe UI", 10, "bold"), background="#1F4E79", foreground="#FFFFFF")
        style.map("Primary.TButton", background=[("active", "#15395B"), ("disabled", "#B0BEC5")])
        style.configure("Success.TButton", font=("Segoe UI", 10, "bold"), background="#2E7D32", foreground="#FFFFFF")
        style.map("Success.TButton", background=[("active", "#1B5E20"), ("disabled", "#B0BEC5")])

    def _build_ui(self):
        # 1. Header Banner
        header_frame = tk.Frame(self.root, bg="#1F4E79", height=85)
        header_frame.pack(fill="x", side="top")
        header_frame.pack_propagate(False)

        lbl_title = tk.Label(header_frame, text="DataCleaner Pro — Enterprise Automation", font=("Segoe UI", 15, "bold"), bg="#1F4E79", fg="#FFFFFF")
        lbl_title.pack(anchor="w", padx=20, pady=(12, 0))

        lbl_sub = tk.Label(header_frame, text="Clean messy CSV/Excel files into Executive Excel KPI Reports with 1 click", font=("Segoe UI", 9), bg="#1F4E79", fg="#D9E6F2")
        lbl_sub.pack(anchor="w", padx=20, pady=(2, 0))

        # 2. Main Container
        main_container = tk.Frame(self.root, bg="#F3F4F6", padx=20, pady=15)
        main_container.pack(fill="both", expand=True)

        # File Selection Card
        file_card = tk.LabelFrame(main_container, text=" Select Input Dataset ", font=("Segoe UI", 10, "bold"), bg="#FFFFFF", fg="#1F4E79", padx=15, pady=12)
        file_card.pack(fill="x", pady=(0, 12))

        self.file_entry = ttk.Entry(file_card, font=("Segoe UI", 9))
        self.file_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))

        btn_browse = ttk.Button(file_card, text="Browse File...", command=self.browse_file)
        btn_browse.pack(side="right")

        # Action Buttons
        btn_frame = tk.Frame(main_container, bg="#F3F4F6")
        btn_frame.pack(fill="x", pady=(0, 12))

        self.btn_run = tk.Button(
            btn_frame, text="Run Automation Pipeline", font=("Segoe UI", 11, "bold"),
            bg="#1F4E79", fg="#FFFFFF", activebackground="#15395B", activeforeground="#FFFFFF",
            relief="flat", cursor="hand2", command=self.start_processing_thread, padx=15, pady=6
        )
        self.btn_run.pack(side="left", fill="x", expand=True, padx=(0, 8))

        self.btn_open_excel = tk.Button(
            btn_frame, text="Open Cleaned Report", font=("Segoe UI", 11, "bold"),
            bg="#2E7D32", fg="#FFFFFF", activebackground="#1B5E20", activeforeground="#FFFFFF",
            relief="flat", cursor="hand2", command=self.open_excel_report, state="disabled", padx=15, pady=6
        )
        self.btn_open_excel.pack(side="right", fill="x", expand=True, padx=(8, 0))

        # Progress Bar
        self.progress = ttk.Progressbar(main_container, mode="indeterminate")
        self.progress.pack(fill="x", pady=(0, 10))

        # Activity Log Card
        log_card = tk.LabelFrame(main_container, text=" Pipeline Execution Log ", font=("Segoe UI", 10, "bold"), bg="#FFFFFF", fg="#1F4E79", padx=10, pady=10)
        log_card.pack(fill="both", expand=True)

        self.log_text = tk.Text(log_card, bg="#1E1E1E", fg="#D4D4D4", font=("Consolas", 9), relief="flat", wrap="word")
        self.log_text.pack(side="left", fill="both", expand=True)

        scrollbar = ttk.Scrollbar(log_card, command=self.log_text.yview)
        scrollbar.pack(side="right", fill="y")
        self.log_text.config(yscrollcommand=scrollbar.set)

        self.append_log("System Ready. Select a CSV or Excel file to begin cleaning.")

    def append_log(self, text):
        self.log_text.insert(tk.END, text + "\n")
        self.log_text.see(tk.END)

    def browse_file(self):
        default_dir = str(Path("data").resolve()) if Path("data").exists() else os.getcwd()
        f = filedialog.askopenfilename(
            initialdir=default_dir,
            title="Select Raw/Messy CSV or Excel File",
            filetypes=[("Excel & CSV Files", "*.xlsx *.xls *.csv"), ("CSV Files", "*.csv"), ("Excel Files", "*.xlsx *.xls")]
        )
        if f:
            self.selected_file = f
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, f)
            self.append_log(f"Selected: {Path(f).name}")

    def start_processing_thread(self):
        file_path = self.file_entry.get().strip()
        if not file_path or not Path(file_path).exists():
            messagebox.showerror("Error", "Please select a valid CSV or Excel file.")
            return

        self.btn_run.config(state="disabled")
        self.btn_open_excel.config(state="disabled")
        self.progress.start(10)
        
        thread = threading.Thread(target=self._process_worker, args=(file_path,), daemon=True)
        thread.start()

    def _process_worker(self, file_path):
        try:
            output_path, audit = run_pipeline(file_path, log_callback=self.append_log)
            self.generated_report = output_path
            self.root.after(0, self._on_success)
        except Exception as e:
            self.root.after(0, lambda: self._on_error(str(e)))

    def _on_success(self):
        self.progress.stop()
        self.btn_run.config(state="normal")
        self.btn_open_excel.config(state="normal")
        messagebox.showinfo("Success", f"Processing Complete!\n\nExecutive Report saved to:\n{self.generated_report}")

    def _on_error(self, err_msg):
        self.progress.stop()
        self.btn_run.config(state="normal")
        self.append_log(f"[ERROR] {err_msg}")
        messagebox.showerror("Error Occurred", f"Failed to clean file:\n{err_msg}")

    def open_excel_report(self):
        if self.generated_report and Path(self.generated_report).exists():
            try:
                os.startfile(str(self.generated_report))
            except Exception as e:
                messagebox.showerror("Error", f"Could not launch Excel: {e}")
        else:
            messagebox.showwarning("Warning", "No generated report found.")

def main():
    parser = argparse.ArgumentParser(description="DataCleaner Pro - Clean and summarize raw Excel/CSV data.")
    parser.add_argument("--file", "-f", type=str, help="Path to input messy file for CLI execution")
    parser.add_argument("--output", "-o", type=str, help="Path to output executive report")
    args = parser.parse_args()

    if args.file:
        run_pipeline(args.file, args.output)
    else:
        root = tk.Tk()
        app = ModernCleanerGUI(root)
        root.mainloop()

if __name__ == "__main__":
    main()
