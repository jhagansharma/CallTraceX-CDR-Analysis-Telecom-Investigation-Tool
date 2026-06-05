"""
CDR Forensic Analysis Tool v3.1 - Professional GUI
Law Enforcement Edition
Modern dark-themed interface with progress tracking
"""

import sys
import os
import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext, messagebox
import threading
import time

# Add core modules to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'core'))

# Check dependencies before importing
def _check_dependencies():
    """Check all required packages are installed"""
    missing = []
    for pkg in ['pandas', 'openpyxl', 'numpy']:
        try:
            __import__(pkg)
        except ImportError:
            missing.append(pkg)
    if missing:
        try:
            import tkinter.messagebox as mb
            root = tk.Tk()
            root.withdraw()
            mb.showerror(
                "Missing Dependencies",
                f"Required packages not installed:\n\n"
                f"{', '.join(missing)}\n\n"
                f"Please run: pip install {' '.join(missing)}\n\n"
                f"Or run install.bat again."
            )
            root.destroy()
        except Exception:
            print(f"ERROR: Missing packages: {', '.join(missing)}")
            print(f"Run: pip install {' '.join(missing)}")
        sys.exit(1)

_check_dependencies()

from parser import CDRParser
from analyzer import CDRAnalyzer
from reporter import ExcelReporter


# ============================================================
# COLORS / THEME
# ============================================================
BG = '#0f1117'
BG2 = '#161b22'
BG3 = '#1c2333'
CARD_BG = '#1a1f2e'
ACCENT = '#2f81f7'
ACCENT_HOVER = '#4493f8'
RED = '#f85149'
GREEN = '#3fb950'
YELLOW = '#d29922'
TEXT = '#e6edf3'
TEXT_DIM = '#7d8590'
TEXT_BRIGHT = '#ffffff'
BORDER = '#30363d'
LOG_BG = '#0d1117'
LOG_FG = '#58a6ff'


class CDRAnalysisGUI:
    """Professional CDR Analysis GUI"""

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("CDR FORENSIC ANALYSIS TOOL v3.1")
        self.root.geometry("1050x780")
        self.root.minsize(900, 650)
        self.root.configure(bg=BG)

        # Try to set icon (ignore if fails)
        try:
            self.root.iconname("CDR Tool")
        except Exception:
            pass

        self.csv_file = None
        self.output_file = None
        self.is_running = False
        self.start_time = None
        self.timer_id = None
        self.last_report_path = None

        self._create_styles()
        self._create_widgets()

    # ----------------------------------------------------------
    # STYLES
    # ----------------------------------------------------------
    def _create_styles(self):
        style = ttk.Style()
        style.theme_use('clam')

        style.configure('Card.TFrame', background=CARD_BG)
        style.configure('BG.TFrame', background=BG)

        style.configure('Title.TLabel', background=BG, foreground=TEXT_BRIGHT,
                        font=('Segoe UI', 20, 'bold'))
        style.configure('Sub.TLabel', background=BG, foreground=TEXT_DIM,
                        font=('Segoe UI', 10))
        style.configure('CardTitle.TLabel', background=CARD_BG, foreground=TEXT,
                        font=('Segoe UI', 10, 'bold'))
        style.configure('CardValue.TLabel', background=CARD_BG, foreground=GREEN,
                        font=('Consolas', 10))
        style.configure('Status.TLabel', background=BG, foreground=YELLOW,
                        font=('Segoe UI', 10, 'bold'))
        style.configure('Path.TLabel', background=CARD_BG, foreground=TEXT_DIM,
                        font=('Consolas', 9))

        style.configure('Accent.TButton', background=ACCENT, foreground=TEXT_BRIGHT,
                        font=('Segoe UI', 10, 'bold'), padding=(12, 6))
        style.map('Accent.TButton',
                  background=[('active', ACCENT_HOVER), ('disabled', '#21262d')])

        style.configure('Green.TButton', background='#238636', foreground=TEXT_BRIGHT,
                        font=('Segoe UI', 10, 'bold'), padding=(12, 6))
        style.map('Green.TButton',
                  background=[('active', '#2ea043'), ('disabled', '#21262d')])

        style.configure('Red.TButton', background='#da3633', foreground=TEXT_BRIGHT,
                        font=('Segoe UI', 12, 'bold'), padding=(16, 10))
        style.map('Red.TButton',
                  background=[('active', '#f85149'), ('disabled', '#21262d')])

        style.configure('Custom.Horizontal.TProgressbar',
                        troughcolor=BG3, background=ACCENT, thickness=6)

    # ----------------------------------------------------------
    # WIDGETS
    # ----------------------------------------------------------
    def _create_widgets(self):
        # Top bar
        top = ttk.Frame(self.root, style='BG.TFrame')
        top.pack(fill='x', padx=24, pady=(20, 0))

        ttk.Label(top, text="CDR FORENSIC ANALYSIS TOOL",
                  style='Title.TLabel').pack(side='left')

        self.status_label = ttk.Label(top, text="Ready", style='Status.TLabel')
        self.status_label.pack(side='right')

        ttk.Label(self.root, text="Made by jhagan  ·  v3.1" ,
                  style='Sub.TLabel').pack(anchor='w', padx=24, pady=(2, 12))

        # Separator
        sep = tk.Frame(self.root, bg=BORDER, height=1)
        sep.pack(fill='x', padx=24)

        # === FILE SELECTION CARD ===
        file_card = tk.Frame(self.root, bg=CARD_BG, highlightbackground=BORDER,
                             highlightthickness=1, padx=16, pady=12)
        file_card.pack(fill='x', padx=24, pady=(16, 0))

        # Input row
        input_row = tk.Frame(file_card, bg=CARD_BG)
        input_row.pack(fill='x', pady=(0, 8))

        tk.Label(input_row, text="INPUT CDR FILE", bg=CARD_BG, fg=TEXT_DIM,
                 font=('Segoe UI', 8, 'bold')).pack(anchor='w')

        input_inner = tk.Frame(input_row, bg=CARD_BG)
        input_inner.pack(fill='x', pady=(2, 0))

        self.input_entry = tk.Entry(input_inner, bg=BG3, fg=GREEN, insertbackground=TEXT,
                                    font=('Consolas', 10), relief='flat', bd=0,
                                    highlightbackground=BORDER, highlightthickness=1)
        self.input_entry.pack(side='left', fill='x', expand=True, ipady=6)
        self.input_entry.insert(0, ' Select a CDR CSV file...')
        self.input_entry.config(state='readonly')

        browse_btn = tk.Button(input_inner, text="  Browse CSV  ",
                               bg=ACCENT, fg=TEXT_BRIGHT, font=('Segoe UI', 9, 'bold'),
                               relief='flat', cursor='hand2', activebackground=ACCENT_HOVER,
                               command=self._browse_input)
        browse_btn.pack(side='right', padx=(8, 0))

        # Output row
        output_row = tk.Frame(file_card, bg=CARD_BG)
        output_row.pack(fill='x')

        tk.Label(output_row, text="OUTPUT REPORT", bg=CARD_BG, fg=TEXT_DIM,
                 font=('Segoe UI', 8, 'bold')).pack(anchor='w')

        output_inner = tk.Frame(output_row, bg=CARD_BG)
        output_inner.pack(fill='x', pady=(2, 0))

        self.output_entry = tk.Entry(output_inner, bg=BG3, fg=GREEN, insertbackground=TEXT,
                                     font=('Consolas', 10), relief='flat', bd=0,
                                     highlightbackground=BORDER, highlightthickness=1)
        self.output_entry.pack(side='left', fill='x', expand=True, ipady=6)
        self.output_entry.insert(0, ' Choose save location...')
        self.output_entry.config(state='readonly')

        save_btn = tk.Button(output_inner, text="  Save As  ",
                             bg='#238636', fg=TEXT_BRIGHT, font=('Segoe UI', 9, 'bold'),
                             relief='flat', cursor='hand2', activebackground='#2ea043',
                             command=self._browse_output)
        save_btn.pack(side='right', padx=(8, 0))

        # === STATS CARDS ===
        stats_frame = tk.Frame(self.root, bg=BG)
        stats_frame.pack(fill='x', padx=24, pady=(12, 0))

        self.stat_cards = {}
        stat_items = [
            ('records', 'RECORDS', '-'),
            ('target', 'TARGET', '-'),
            ('subscriber', 'SUBSCRIBER', '-'),
            ('sheets', 'SHEETS', '-'),
            ('circle', 'CIRCLE', '-'),
        ]
        for i, (key, label, default) in enumerate(stat_items):
            card = tk.Frame(stats_frame, bg=CARD_BG, highlightbackground=BORDER,
                            highlightthickness=1, padx=12, pady=8)
            card.pack(side='left', fill='both', expand=True, padx=(0 if i == 0 else 4, 0))

            tk.Label(card, text=label, bg=CARD_BG, fg=TEXT_DIM,
                     font=('Segoe UI', 7, 'bold')).pack(anchor='w')
            val_label = tk.Label(card, text=default, bg=CARD_BG, fg=GREEN,
                                 font=('Consolas', 12, 'bold'))
            val_label.pack(anchor='w')
            self.stat_cards[key] = val_label

        # === GENERATE BUTTON + PROGRESS ===
        btn_frame = tk.Frame(self.root, bg=BG)
        btn_frame.pack(fill='x', padx=24, pady=(12, 0))

        self.generate_btn = tk.Button(
            btn_frame, text="  GENERATE FORENSIC REPORT  ",
            bg='#da3633', fg=TEXT_BRIGHT, font=('Segoe UI', 13, 'bold'),
            relief='flat', cursor='hand2', activebackground='#f85149',
            command=self._start_analysis, state='disabled',
            disabledforeground='#484f58', padx=20, pady=10
        )
        self.generate_btn.pack(side='left', fill='x', expand=True)

        self.open_folder_btn = tk.Button(
            btn_frame, text="  📂 Open Report Folder  ",
            bg='#238636', fg=TEXT_BRIGHT, font=('Segoe UI', 10, 'bold'),
            relief='flat', cursor='hand2', activebackground='#2ea043',
            command=self._open_report_folder, state='disabled',
            disabledforeground='#484f58', padx=10, pady=10
        )
        self.open_folder_btn.pack(side='right', padx=(8, 0))

        # Progress bar area
        action_frame = tk.Frame(self.root, bg=BG)
        action_frame.pack(fill='x', padx=24, pady=(4, 0))

        # Progress bar
        self.progress = ttk.Progressbar(action_frame, mode='determinate',
                                         style='Custom.Horizontal.TProgressbar')
        self.progress.pack(fill='x', pady=(6, 0))

        self.progress_label = tk.Label(action_frame, text="", bg=BG, fg=TEXT_DIM,
                                       font=('Segoe UI', 8))
        self.progress_label.pack(side='left', anchor='w', pady=(2, 0))

        self.timer_label = tk.Label(action_frame, text="", bg=BG, fg=ACCENT,
                                    font=('Consolas', 9, 'bold'))
        self.timer_label.pack(side='right', anchor='e', pady=(2, 0))

        # === LOG AREA ===
        log_header = tk.Frame(self.root, bg=BG)
        log_header.pack(fill='x', padx=24, pady=(12, 0))

        tk.Label(log_header, text="ANALYSIS LOG", bg=BG, fg=TEXT_DIM,
                 font=('Segoe UI', 8, 'bold')).pack(side='left')

        self.clear_btn = tk.Button(log_header, text="Clear", bg=BG, fg=TEXT_DIM,
                                   font=('Segoe UI', 8), relief='flat', cursor='hand2',
                                   command=self._clear_log)
        self.clear_btn.pack(side='right')

        self.log_area = scrolledtext.ScrolledText(
            self.root, bg=LOG_BG, fg=LOG_FG, font=('Consolas', 9),
            relief='flat', bd=0, insertbackground=TEXT,
            highlightbackground=BORDER, highlightthickness=1,
            state='disabled', wrap='word'
        )
        self.log_area.pack(fill='both', expand=True, padx=24, pady=(4, 4))

        # Version footer bar
        footer = tk.Frame(self.root, bg=BG2, height=24)
        footer.pack(fill='x', side='bottom')
        tk.Label(footer, text="CDR Forensic Analysis Tool v3.1 | Law Enforcement Edition",
                 bg=BG2, fg=TEXT_DIM, font=('Segoe UI', 7)).pack(side='left', padx=12)
        tk.Label(footer, text="Developed for Indian Telecom CDR Analysis",
                 bg=BG2, fg=TEXT_DIM, font=('Segoe UI', 7)).pack(side='right', padx=12)

        # Tag colors for log
        self.log_area.tag_config('success', foreground=GREEN)
        self.log_area.tag_config('error', foreground=RED)
        self.log_area.tag_config('warn', foreground=YELLOW)
        self.log_area.tag_config('header', foreground=TEXT_BRIGHT, font=('Consolas', 9, 'bold'))
        self.log_area.tag_config('dim', foreground=TEXT_DIM)

        self._log("Welcome! Select a CDR CSV file and output location to begin.", 'dim')

    # ----------------------------------------------------------
    # FILE BROWSING
    # ----------------------------------------------------------
    def _browse_input(self):
        filename = filedialog.askopenfilename(
            title="Select CDR CSV File",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if filename:
            self.csv_file = filename
            self.input_entry.config(state='normal')
            self.input_entry.delete(0, tk.END)
            self.input_entry.insert(0, f' {filename}')
            self.input_entry.config(state='readonly')
            self._log(f"Input: {filename}", 'success')

            # Auto-suggest output name
            if not self.output_file:
                base = os.path.splitext(filename)[0]
                suggested = base + '_CDR_Report.xlsx'
                self.output_file = suggested
                self.output_entry.config(state='normal')
                self.output_entry.delete(0, tk.END)
                self.output_entry.insert(0, f' {suggested}')
                self.output_entry.config(state='readonly')
                self._log(f"Output auto-set: {suggested}", 'dim')

            self._check_ready()

    def _browse_output(self):
        initial = ''
        if self.csv_file:
            base = os.path.splitext(os.path.basename(self.csv_file))[0]
            initial = f"{base}_CDR_Report.xlsx"

        filename = filedialog.asksaveasfilename(
            title="Save Report As",
            defaultextension=".xlsx",
            initialfile=initial or "CDR_Analysis_Report.xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if filename:
            self.output_file = filename
            self.output_entry.config(state='normal')
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, f' {filename}')
            self.output_entry.config(state='readonly')
            self._log(f"Output: {filename}", 'success')
            self._check_ready()

    def _check_ready(self):
        if self.csv_file and self.output_file:
            self.generate_btn.config(state='normal')

    # ----------------------------------------------------------
    # LOG
    # ----------------------------------------------------------
    def _log(self, message, tag=None):
        self.log_area.config(state='normal')
        if tag:
            self.log_area.insert('end', message + '\n', tag)
        else:
            self.log_area.insert('end', message + '\n')
        self.log_area.see('end')
        self.log_area.config(state='disabled')
        self.root.update_idletasks()

    def _clear_log(self):
        self.log_area.config(state='normal')
        self.log_area.delete('1.0', 'end')
        self.log_area.config(state='disabled')

    def _set_status(self, text, color=YELLOW):
        self.status_label.config(text=text, foreground=color)

    def _set_progress(self, value, text=''):
        self.progress['value'] = value
        self.progress_label.config(text=text)
        self.root.update_idletasks()

    def _update_stat(self, key, value):
        if key in self.stat_cards:
            self.stat_cards[key].config(text=str(value))

    # ----------------------------------------------------------
    # ANALYSIS
    # ----------------------------------------------------------
    def _start_analysis(self):
        if self.is_running:
            return

        # Validate inputs before starting
        if not self.csv_file or not os.path.isfile(self.csv_file):
            messagebox.showerror("Input Error", "Please select a valid CDR CSV file.")
            return
        if os.path.getsize(self.csv_file) == 0:
            messagebox.showerror("Input Error", "Selected file is empty (0 bytes).")
            return
        if not self.output_file:
            messagebox.showerror("Output Error", "Please select an output location.")
            return

        # Check if output file is already open
        if os.path.exists(self.output_file):
            try:
                with open(self.output_file, 'a'):
                    pass
            except PermissionError:
                messagebox.showerror(
                    "File Locked",
                    f"Output file is open in another program (Excel?):\n\n"
                    f"{self.output_file}\n\n"
                    f"Please close the file and try again."
                )
                return

        # Check output directory exists
        out_dir = os.path.dirname(self.output_file)
        if out_dir and not os.path.isdir(out_dir):
            try:
                os.makedirs(out_dir, exist_ok=True)
            except Exception:
                messagebox.showerror("Output Error", f"Cannot create output directory:\n{out_dir}")
                return

        self.is_running = True
        self.start_time = time.time()
        self.generate_btn.config(state='disabled', text="  PROCESSING...  ")
        self._set_status("Analyzing...", YELLOW)
        self._set_progress(0, '')
        self._update_timer()

        thread = threading.Thread(target=self._perform_analysis, daemon=True)
        thread.start()

    def _perform_analysis(self):
        start_time = time.time()

        try:
            # ---- STEP 1: PARSE ----
            self._set_progress(5, 'Parsing CDR file...')
            self._log("\n" + "=" * 60, 'header')
            self._log("  STEP 1: PARSING CDR FILE", 'header')
            self._log("=" * 60, 'header')

            parser = CDRParser(self.csv_file)
            cleaned_df = parser.parse()

            metadata = parser.get_metadata_dict()
            metadata['_csv_path'] = self.csv_file  # For raw CDR sheet

            self._update_stat('records', len(cleaned_df))
            self._update_stat('target', parser.input_value)
            self._update_stat('subscriber', parser.subscriber_name[:20])
            self._update_stat('circle', parser.circle[:15])

            self._log(f"  Parsed {len(cleaned_df)} records", 'success')
            self._log(f"  Target: {parser.input_value}")
            self._log(f"  Subscriber: {parser.subscriber_name}")
            self._log(f"  Circle: {parser.circle}")
            self._set_progress(25, 'Parsing complete')

            # ---- STEP 2: ANALYZE ----
            self._set_progress(30, 'Running forensic analysis...')
            self._log("\n" + "=" * 60, 'header')
            self._log("  STEP 2: RUNNING FORENSIC ANALYSIS", 'header')
            self._log("=" * 60, 'header')

            analyzer = CDRAnalyzer(cleaned_df, metadata)
            results = analyzer.analyze_all()

            sheet_count = len([k for k, v in results.items()
                              if v is not None and len(v) > 0 and k != 'Smart Report'])
            self._update_stat('sheets', sheet_count + 2)  # +smart report +raw cdr
            self._log(f"  Generated {sheet_count} analysis sheets", 'success')
            self._set_progress(65, 'Analysis complete')

            # ---- STEP 3: GENERATE REPORT ----
            self._set_progress(70, 'Generating Excel report...')
            self._log("\n" + "=" * 60, 'header')
            self._log("  STEP 3: GENERATING EXCEL REPORT", 'header')
            self._log("=" * 60, 'header')

            reporter = ExcelReporter(parser.input_value)
            report_path = reporter.generate_report(results, self.output_file, cleaned_df, metadata)
            self._set_progress(95, 'Formatting...')
            self.last_report_path = report_path

            elapsed = time.time() - start_time

            # ---- SUCCESS ----
            self._set_progress(100, f'Complete in {elapsed:.1f}s')
            self._set_status("Complete!", GREEN)
            self.root.after(0, lambda: self.open_folder_btn.config(state='normal'))

            self._log("\n" + "=" * 60, 'success')
            self._log("  ANALYSIS COMPLETE!", 'success')
            self._log("=" * 60, 'success')
            self._log(f"\n  Records Analyzed : {len(cleaned_df)}")
            self._log(f"  Report Sheets    : {sheet_count + 2}")
            self._log(f"  Time Taken       : {elapsed:.1f} seconds")
            self._log(f"  Report Saved     : {report_path}", 'success')

            # Success popup
            self.root.after(0, lambda: messagebox.showinfo(
                "Analysis Complete",
                f"CDR Forensic Report Generated!\n\n"
                f"Target: {parser.input_value}\n"
                f"Subscriber: {parser.subscriber_name}\n"
                f"Records: {len(cleaned_df)}\n"
                f"Sheets: {sheet_count + 2}\n"
                f"Time: {elapsed:.1f}s\n\n"
                f"Saved to:\n{report_path}"
            ))

            # Ask to open file
            self.root.after(100, lambda: self._ask_open_file(report_path))

        except FileNotFoundError as e:
            self._set_progress(0, '')
            self._set_status("File Error!", RED)
            self._log(f"\n  FILE ERROR: {str(e)}", 'error')
            self.root.after(0, lambda: messagebox.showerror(
                "File Not Found", str(e)
            ))

        except PermissionError as e:
            self._set_progress(0, '')
            self._set_status("Permission Error!", RED)
            self._log(f"\n  PERMISSION ERROR: {str(e)}", 'error')
            self.root.after(0, lambda: messagebox.showerror(
                "Permission Error",
                f"File is locked or permission denied:\n\n{str(e)}\n\n"
                f"If the report file is open in Excel, please close it and try again."
            ))

        except ValueError as e:
            self._set_progress(0, '')
            self._set_status("Data Error!", RED)
            self._log(f"\n  DATA ERROR: {str(e)}", 'error')
            self.root.after(0, lambda: messagebox.showerror(
                "CDR Data Error",
                f"Problem with CDR file format:\n\n{str(e)}\n\n"
                f"Make sure you selected the correct CSV file from the telecom provider."
            ))

        except Exception as e:
            self._set_progress(0, '')
            self._set_status("Error!", RED)
            self._log(f"\n  ERROR: {str(e)}", 'error')

            import traceback
            self._log("\n  Full error details:", 'error')
            self._log(traceback.format_exc(), 'error')

            self.root.after(0, lambda: messagebox.showerror(
                "Analysis Error",
                f"An unexpected error occurred:\n\n{str(e)}\n\n"
                f"Please check the Analysis Log for details."
            ))

        finally:
            self.is_running = False
            if self.timer_id:
                self.root.after_cancel(self.timer_id)
                self.timer_id = None
            self.root.after(0, lambda: self.generate_btn.config(
                state='normal', text="  GENERATE FORENSIC REPORT  "
            ))

    def _ask_open_file(self, path):
        """Ask user if they want to open the file"""
        if messagebox.askyesno("Open Report?", "Do you want to open the generated report?"):
            try:
                if sys.platform == 'win32':
                    os.startfile(path)
                elif sys.platform == 'darwin':
                    os.system(f'open "{path}"')
                else:
                    os.system(f'xdg-open "{path}"')
            except Exception:
                pass

    def _open_report_folder(self):
        """Open the folder containing the last generated report"""
        if self.last_report_path and os.path.exists(self.last_report_path):
            folder = os.path.dirname(self.last_report_path)
            try:
                if sys.platform == 'win32':
                    os.startfile(folder)
                elif sys.platform == 'darwin':
                    os.system(f'open "{folder}"')
                else:
                    os.system(f'xdg-open "{folder}"')
            except Exception:
                pass

    def _update_timer(self):
        """Update elapsed time display during processing"""
        if self.is_running and self.start_time:
            elapsed = time.time() - self.start_time
            mins = int(elapsed // 60)
            secs = int(elapsed % 60)
            self.timer_label.config(text=f"⏱ {mins:02d}:{secs:02d}")
            self.timer_id = self.root.after(1000, self._update_timer)

    # ----------------------------------------------------------
    # RUN
    # ----------------------------------------------------------
    def run(self):
        # Center window
        self.root.update_idletasks()
        w = self.root.winfo_width()
        h = self.root.winfo_height()
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        x = (sw - w) // 2
        y = (sh - h) // 2
        self.root.geometry(f"{w}x{h}+{x}+{y}")

        self.root.mainloop()


def main():
    app = CDRAnalysisGUI()
    app.run()


if __name__ == "__main__":
    main()
