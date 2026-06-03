"""
product_similarity_app.py
=========================
GUI wrapper for the product_similarity.py analysis script.
Run with:
    & "C:\Program Files (x86)\Microsoft Visual Studio\Shared\Python39_64\python.exe" product_similarity_app.py

Requires product_similarity.py to be in the same directory.
Dependencies: pandas, rapidfuzz, xlsxwriter, openpyxl  (same as the script)
"""

import json
import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk

# ── Config persistence ─────────────────────────────────────────────────────────
# Saves last-used settings so they are restored on next launch.
CONFIG_FILE = Path(__file__).parent / ".product_similarity_config.json"

DEFAULT_CONFIG = {
    "input_file":  r"C:\Users\ejdiguilio\OneDrive - Vector Security\Desktop\ProductList.xlsx",
    "sheet":       "0",
    "no_header":   True,
    "threshold":   "95",
    "window":      "200",
    "scorer":      "WRatio",
    "output_dir":  "",   # empty = same folder as input
}

# Scorer options shown in the dropdown.
# Keys are display labels; values are rapidfuzz function names.
SCORERS = {
    "WRatio           — auto-picks best scorer (recommended)": "WRatio",
    "ratio            — character-level overlap, order-sensitive": "ratio",
    "partial_ratio    — best substring match (prefix/suffix variants)": "partial_ratio",
    "token_sort_ratio — sorts tokens before comparing (reordered words)": "token_sort_ratio",
    "token_set_ratio  — ignores duplicate tokens (most lenient)": "token_set_ratio",
}
# Reverse map: function name → display label
_SCORER_LABEL = {v: k for k, v in SCORERS.items()}


def load_config() -> dict:
    if CONFIG_FILE.exists():
        try:
            with open(CONFIG_FILE) as f:
                saved = json.load(f)
            return {**DEFAULT_CONFIG, **saved}
        except Exception:
            pass
    return dict(DEFAULT_CONFIG)


def save_config(cfg: dict) -> None:
    try:
        with open(CONFIG_FILE, "w") as f:
            json.dump(cfg, f, indent=2)
    except Exception:
        pass


# ── Log redirector ─────────────────────────────────────────────────────────────
class QueueWriter:
    """Redirects print() / stdout writes into a thread-safe Queue."""
    def __init__(self, q: queue.Queue):
        self._q = q

    def write(self, text: str):
        if text:
            self._q.put(text)

    def flush(self):
        pass


# ── Main application ───────────────────────────────────────────────────────────
class ProductSimilarityApp:

    TITLE   = "Product Similarity Finder"
    WIDTH   = 680
    HEIGHT  = 620
    BG      = "#f5f5f5"
    ACCENT  = "#4472C4"
    BTN_FG  = "white"
    FONT    = ("Segoe UI", 10)
    FONT_B  = ("Segoe UI", 10, "bold")
    MONO    = ("Consolas", 9)

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(self.TITLE)
        self.root.resizable(True, True)
        self.root.configure(bg=self.BG)
        self.root.geometry(f"{self.WIDTH}x{self.HEIGHT}")
        self.root.minsize(560, 500)

        self._cfg = load_config()
        self._log_queue: queue.Queue = queue.Queue()
        self._running = False
        self._last_output: str = ""   # path of last written output file

        self._build_ui()
        self._poll_log()

    # ── UI construction ────────────────────────────────────────────────────────

    def _build_ui(self):
        pad = {"padx": 12, "pady": 6}

        # ── Input section ──────────────────────────────────────────────────────
        input_frame = self._section(self.root, "Input File")

        tk.Label(input_frame, text="File:", font=self.FONT, bg=self.BG, anchor="w").grid(
            row=0, column=0, sticky="w", **pad)
        self._var_input = tk.StringVar(value=self._cfg["input_file"])
        tk.Entry(input_frame, textvariable=self._var_input, font=self.FONT, width=52).grid(
            row=0, column=1, sticky="ew", padx=(0, 6), pady=6)
        tk.Button(input_frame, text="Browse…", font=self.FONT,
                  command=self._browse_input).grid(row=0, column=2, pady=6, padx=(0, 12))

        tk.Label(input_frame, text="Sheet:", font=self.FONT, bg=self.BG, anchor="w").grid(
            row=1, column=0, sticky="w", **pad)
        self._var_sheet = tk.StringVar(value=self._cfg["sheet"])
        tk.Entry(input_frame, textvariable=self._var_sheet, font=self.FONT, width=12).grid(
            row=1, column=1, sticky="w", padx=(0, 6), pady=6)

        self._var_no_header = tk.BooleanVar(value=self._cfg["no_header"])
        tk.Checkbutton(input_frame, text="No header row (first row is data)",
                       variable=self._var_no_header, font=self.FONT,
                       bg=self.BG, activebackground=self.BG).grid(
            row=2, column=1, columnspan=2, sticky="w", padx=(0, 12), pady=(0, 6))

        input_frame.columnconfigure(1, weight=1)

        # ── Settings section ───────────────────────────────────────────────────
        settings_frame = self._section(self.root, "Settings")

        tk.Label(settings_frame, text="Similarity Threshold:", font=self.FONT,
                 bg=self.BG, anchor="w").grid(row=0, column=0, sticky="w", **pad)
        self._var_threshold = tk.StringVar(value=self._cfg["threshold"])
        thresh_entry = tk.Entry(settings_frame, textvariable=self._var_threshold,
                                font=self.FONT, width=6)
        thresh_entry.grid(row=0, column=1, sticky="w", pady=6)
        tk.Label(settings_frame, text="%  (0–100, default 95)",
                 font=self.FONT, bg=self.BG, fg="#666").grid(
            row=0, column=2, sticky="w", padx=6)

        tk.Label(settings_frame, text="Fuzzy Scorer:", font=self.FONT,
                 bg=self.BG, anchor="w").grid(row=1, column=0, sticky="w", **pad)
        saved_scorer_label = _SCORER_LABEL.get(self._cfg.get("scorer", "WRatio"),
                                               list(SCORERS.keys())[0])
        self._var_scorer = tk.StringVar(value=saved_scorer_label)
        scorer_cb = ttk.Combobox(
            settings_frame, textvariable=self._var_scorer,
            values=list(SCORERS.keys()), state="readonly",
            font=self.FONT, width=54)
        scorer_cb.grid(row=1, column=1, columnspan=3, sticky="ew",
                       padx=(0, 12), pady=6)

        tk.Label(settings_frame, text="Window Size:", font=self.FONT,
                 bg=self.BG, anchor="w").grid(row=2, column=0, sticky="w", **pad)
        self._var_window = tk.StringVar(value=self._cfg["window"])
        tk.Entry(settings_frame, textvariable=self._var_window,
                 font=self.FONT, width=6).grid(row=2, column=1, sticky="w", pady=6)
        tk.Label(settings_frame, text="neighbors per item for large lists (default 200)",
                 font=self.FONT, bg=self.BG, fg="#666").grid(
            row=2, column=2, sticky="w", padx=6)

        tk.Label(settings_frame, text="Output Folder:", font=self.FONT,
                 bg=self.BG, anchor="w").grid(row=3, column=0, sticky="w", **pad)
        self._var_output_dir = tk.StringVar(value=self._cfg["output_dir"])
        tk.Entry(settings_frame, textvariable=self._var_output_dir,
                 font=self.FONT, width=36).grid(
            row=3, column=1, columnspan=2, sticky="ew", padx=(0, 6), pady=6)
        tk.Button(settings_frame, text="Browse…", font=self.FONT,
                  command=self._browse_output).grid(row=3, column=3, pady=6, padx=(0, 12))
        tk.Label(settings_frame, text="(leave blank to save alongside input file)",
                 font=self.FONT, bg=self.BG, fg="#666").grid(
            row=4, column=1, columnspan=3, sticky="w", padx=(0, 12), pady=(0, 4))

        settings_frame.columnconfigure(1, weight=0)
        settings_frame.columnconfigure(2, weight=1)

        # ── Run button + progress ──────────────────────────────────────────────
        run_frame = tk.Frame(self.root, bg=self.BG)
        run_frame.pack(fill="x", padx=12, pady=4)

        self._btn_run = tk.Button(
            run_frame, text="▶  Run Analysis", font=("Segoe UI", 11, "bold"),
            bg=self.ACCENT, fg=self.BTN_FG, activebackground="#2a52a0",
            activeforeground="white", relief="flat", padx=20, pady=6,
            cursor="hand2", command=self._run)
        self._btn_run.pack(side="left")

        self._lbl_status = tk.Label(run_frame, text="", font=self.FONT,
                                    bg=self.BG, fg="#444")
        self._lbl_status.pack(side="left", padx=14)

        self._progress = ttk.Progressbar(run_frame, mode="indeterminate", length=160)

        # ── Log section ────────────────────────────────────────────────────────
        log_frame = self._section(self.root, "Log")
        log_frame.pack_configure(fill="both", expand=True)

        self._log = scrolledtext.ScrolledText(
            log_frame, font=self.MONO, bg="#1e1e1e", fg="#d4d4d4",
            insertbackground="white", relief="flat", wrap="word",
            state="disabled", height=10)
        self._log.pack(fill="both", expand=True, padx=8, pady=(0, 6))

        # Tag for highlighted lines (matches found, output path)
        self._log.tag_config("highlight", foreground="#4ec9b0")
        self._log.tag_config("error",     foreground="#f48771")

        # ── Bottom bar ─────────────────────────────────────────────────────────
        bottom = tk.Frame(self.root, bg=self.BG)
        bottom.pack(fill="x", padx=12, pady=(0, 10))

        tk.Button(bottom, text="Clear Log", font=self.FONT,
                  command=self._clear_log).pack(side="left")

        self._btn_open = tk.Button(
            bottom, text="Open Output File", font=self.FONT_B,
            bg="#107c10", fg=self.BTN_FG, activebackground="#0a5c0a",
            activeforeground="white", relief="flat", padx=12, pady=4,
            cursor="hand2", command=self._open_output, state="disabled")
        self._btn_open.pack(side="right")

    def _section(self, parent, title: str) -> tk.LabelFrame:
        """Styled label frame used for each section."""
        frame = tk.LabelFrame(parent, text=f"  {title}  ",
                              font=self.FONT_B, bg=self.BG,
                              fg=self.ACCENT, padx=2, pady=4,
                              relief="groove")
        frame.pack(fill="x", padx=12, pady=(8, 0))
        return frame

    # ── File dialogs ───────────────────────────────────────────────────────────

    def _browse_input(self):
        path = filedialog.askopenfilename(
            title="Select product list file",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
            initialfile=self._var_input.get() or DEFAULT_CONFIG["input_file"],
        )
        if path:
            self._var_input.set(path)

    def _browse_output(self):
        folder = filedialog.askdirectory(title="Select output folder")
        if folder:
            self._var_output_dir.set(folder)

    # ── Validation ─────────────────────────────────────────────────────────────

    def _validate(self) -> bool:
        input_file = self._var_input.get().strip()
        if not input_file:
            messagebox.showerror("Missing input", "Please select an input file.")
            return False
        if not Path(input_file).exists():
            messagebox.showerror("File not found", f"Cannot find:\n{input_file}")
            return False

        try:
            t = float(self._var_threshold.get())
            if not (0 < t <= 100):
                raise ValueError
        except ValueError:
            messagebox.showerror("Invalid threshold",
                                 "Threshold must be a number between 1 and 100.")
            return False

        try:
            w = int(self._var_window.get())
            if w < 1:
                raise ValueError
        except ValueError:
            messagebox.showerror("Invalid window size",
                                 "Window size must be a positive integer.")
            return False

        return True

    # ── Run analysis ───────────────────────────────────────────────────────────

    def _run(self):
        if self._running:
            return
        if not self._validate():
            return

        # Persist settings
        save_config({
            "input_file": self._var_input.get().strip(),
            "sheet":      self._var_sheet.get().strip(),
            "no_header":  self._var_no_header.get(),
            "threshold":  self._var_threshold.get().strip(),
            "window":     self._var_window.get().strip(),
            "scorer":     SCORERS.get(self._var_scorer.get(), "WRatio"),
            "output_dir": self._var_output_dir.get().strip(),
        })

        self._running = True
        self._last_output = ""
        self._btn_run.configure(state="disabled", text="Running…")
        self._btn_open.configure(state="disabled")
        self._progress.pack(side="right")
        self._progress.start(12)
        self._lbl_status.configure(text="Analysis in progress…", fg="#444")

        # Run in background thread so the UI stays responsive
        thread = threading.Thread(target=self._run_worker, daemon=True)
        thread.start()

    def _run_worker(self):
        """Background thread: builds CLI args and runs the analysis."""
        old_stdout = sys.stdout
        sys.stdout = QueueWriter(self._log_queue)

        try:
            # Build arg list matching product_similarity.py's argparse spec
            script = Path(__file__).parent / "product_similarity.py"
            input_file  = self._var_input.get().strip()
            sheet       = self._var_sheet.get().strip() or "0"
            threshold   = self._var_threshold.get().strip()
            window      = self._var_window.get().strip()
            output_dir  = self._var_output_dir.get().strip()
            no_header   = self._var_no_header.get()

            # Resolve the chosen scorer label → rapidfuzz function
            from rapidfuzz import fuzz as _fuzz
            scorer_name = SCORERS.get(self._var_scorer.get(), "WRatio")
            scorer_fn   = getattr(_fuzz, scorer_name, _fuzz.WRatio)
            print(f"  Scorer      : {scorer_name}")

            # Import and call directly (avoids subprocess overhead and keeps stdout capture)
            import importlib.util
            spec = importlib.util.spec_from_file_location("product_similarity", script)
            mod  = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(mod)

            # Build a fake args namespace matching parse_args() output
            class Args:
                pass

            args = Args()
            args.input      = input_file
            args.sheet      = int(sheet) if sheet.isdigit() else sheet
            args.col        = None
            args.no_header  = no_header
            args.threshold  = float(threshold)
            args.window     = int(window)
            args.output_dir = output_dir if output_dir else None
            args.log_level  = "INFO"

            # ── Replicate main() inline so we can capture the output path ──────
            from pathlib import Path as P

            input_path = P(args.input)
            print(f"Reading: {input_path}")

            header = None if args.no_header else 0
            sheet_val = int(args.sheet) if str(args.sheet).isdigit() else args.sheet
            df = mod.pd.read_excel(input_path, sheet_name=sheet_val,
                                   dtype=str, header=header)

            print(f"  Rows loaded : {len(df):,}")
            print(f"  Columns     : {list(df.columns)}")

            if args.col is not None:
                col = args.col
            elif args.no_header:
                col = 0   # no header means integer indices; always use first column
            else:
                col = mod.auto_detect_column(df)
            print(f"  Product col : {repr(col)}")

            raw_values = df.iloc[:, col] if isinstance(col, int) else df[col]
            raw_values = raw_values.dropna().astype(str)
            raw_values = raw_values[raw_values.str.strip() != ""]
            products = sorted(set(mod.normalize(v) for v in raw_values))
            n = len(products)
            print(f"  Unique products to compare : {n:,}")

            if n < 2:
                print("ERROR: Need at least 2 distinct product numbers.")
                self._log_queue.put("__DONE_ERROR__")
                return

            if n <= mod.LARGE_DATASET_CUTOFF:
                pair_count = n * (n - 1) // 2
                print(f"  Strategy    : full pairwise cdist ({pair_count:,} pairs)")
                matches = mod.find_similar_pairs_small(products, args.threshold, scorer_fn)
            else:
                effective = n * min(args.window, n - 1)
                print(f"  Strategy    : sorted sliding window  "
                      f"(window={args.window}, ~{effective:,} comparisons)")
                matches = mod.find_similar_pairs_large(
                    products, args.threshold, scorer_fn, args.window)

            print(f"  Matches found : {len(matches):,}")

            if not matches:
                print(f"No pairs met the {args.threshold}% threshold. No output written.")
                self._log_queue.put("__DONE_ERROR__")
                return

            results_df = (
                mod.pd.DataFrame(matches)
                .sort_values("Similarity Score", ascending=False)
                .reset_index(drop=True)
            )

            from datetime import datetime
            output_dir_path = P(args.output_dir) if args.output_dir else input_path.parent
            timestamp   = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_path = output_dir_path / f"ProductSimilarity_{timestamp}.xlsx"

            mod.write_output(results_df, output_path, args.threshold, n,
                             scorer_name=scorer_name)
            print(f"\nOutput written to:\n  {output_path}")

            # Signal success with the output path
            self._log_queue.put(f"__DONE_OK__{output_path}")

        except Exception as exc:
            import traceback
            print(f"\nERROR: {exc}")
            traceback.print_exc()
            self._log_queue.put("__DONE_ERROR__")
        finally:
            sys.stdout = old_stdout

    # ── Log polling ────────────────────────────────────────────────────────────

    def _poll_log(self):
        """Drain the log queue and write to the ScrolledText widget."""
        try:
            while True:
                msg = self._log_queue.get_nowait()

                if msg.startswith("__DONE_OK__"):
                    self._last_output = msg[len("__DONE_OK__"):]
                    self._on_done(success=True)
                elif msg == "__DONE_ERROR__":
                    self._on_done(success=False)
                else:
                    self._append_log(msg)
        except queue.Empty:
            pass
        finally:
            self.root.after(100, self._poll_log)

    def _append_log(self, text: str):
        self._log.configure(state="normal")

        # Highlight key result lines
        if "Matches found" in text or "Output written" in text:
            self._log.insert("end", text, "highlight")
        elif "ERROR" in text or "Traceback" in text:
            self._log.insert("end", text, "error")
        else:
            self._log.insert("end", text)

        self._log.configure(state="disabled")
        self._log.see("end")

    def _on_done(self, success: bool):
        self._running = False
        self._progress.stop()
        self._progress.pack_forget()
        self._btn_run.configure(state="normal", text="▶  Run Analysis")

        if success:
            self._lbl_status.configure(text="Done  ✓", fg="#107c10")
            self._btn_open.configure(state="normal")
        else:
            self._lbl_status.configure(text="Failed  ✗", fg="#c42b1c")

    def _clear_log(self):
        self._log.configure(state="normal")
        self._log.delete("1.0", "end")
        self._log.configure(state="disabled")
        self._lbl_status.configure(text="")

    def _open_output(self):
        if self._last_output and Path(self._last_output).exists():
            os.startfile(self._last_output)
        else:
            messagebox.showwarning("File not found",
                                   "Output file could not be located.")


# ── Entry point ────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    root = tk.Tk()
    app = ProductSimilarityApp(root)
    root.mainloop()
