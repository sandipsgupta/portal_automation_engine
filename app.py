"""
app.py — Tkinter GUI for Portal Automation Engine
Wraps main.py engine with:
  - Editable config (sheet path, headless, debug)
  - Live log window (stdout redirected to widget)
  - CSV run log saved to logs/ folder
  - Run/Stop controls
"""

import csv
import json
import os
import queue
import sys
import threading
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, font, messagebox, scrolledtext, ttk

# ── Ensure we can import the engine from the same folder ──────────────
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


# ──────────────────────────────────────────────────────────────────────
# CSV RUN LOGGER
# ──────────────────────────────────────────────────────────────────────
class RunLogger:
    """
    Writes a structured CSV log for each run.
    Columns: timestamp, invoice_no, payment_mode, amount, status, notes
    File: logs/run_YYYYMMDD_HHMMSS.csv
    """

    COLUMNS = ["timestamp", "invoice_no", "payment_mode", "amount", "status", "notes"]

    def __init__(self):
        self.log_dir  = Path("logs")
        self.log_dir.mkdir(exist_ok=True)
        timestamp     = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.log_path = self.log_dir / f"run_{timestamp}.csv"
        self._write_header()

    def _write_header(self):
        with open(self.log_path, "w", newline="", encoding="utf-8") as f:
            csv.DictWriter(f, fieldnames=self.COLUMNS).writeheader()

    def record(self, invoice_no="", payment_mode="", amount="",
               status="", notes=""):
        row = {
            "timestamp":    datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "invoice_no":   invoice_no,
            "payment_mode": payment_mode,
            "amount":       amount,
            "status":       status,
            "notes":        notes,
        }
        with open(self.log_path, "a", newline="", encoding="utf-8") as f:
            csv.DictWriter(f, fieldnames=self.COLUMNS).writerow(row)

    @property
    def path(self):
        return str(self.log_path)


# ──────────────────────────────────────────────────────────────────────
# STDOUT REDIRECTOR → tkinter queue
# ──────────────────────────────────────────────────────────────────────
class QueueWriter:
    """Redirects print() output to a thread-safe queue for the GUI."""

    def __init__(self, q: queue.Queue):
        self.queue = q

    def write(self, text):
        if text:
            self.queue.put(text)

    def flush(self):
        pass


# ──────────────────────────────────────────────────────────────────────
# MAIN GUI
# ──────────────────────────────────────────────────────────────────────
class PortalAutomationApp(tk.Tk):

    CONFIG_FILE = "config.json"

    # Colour palette — clean, professional, Windows-native feel
    BG          = "#F5F5F5"
    PANEL_BG    = "#FFFFFF"
    ACCENT      = "#5B2D8E"   # Mavic purple (matches portal)
    ACCENT_DARK = "#3D1A6B"
    TEXT        = "#1A1A1A"
    TEXT_MUTED  = "#6B6B6B"
    LOG_BG      = "#1E1E1E"   # dark terminal feel
    LOG_FG      = "#D4D4D4"
    LOG_OK      = "#4EC9B0"   # teal  — success
    LOG_WARN    = "#DCDCAA"   # yellow — warning/skip
    LOG_ERR     = "#F44747"   # red   — error
    LOG_INFO    = "#9CDCFE"   # blue  — info

    def __init__(self):
        super().__init__()

        self.title("Portal Automation Engine")
        self.geometry("820x680")
        self.minsize(700, 560)
        self.configure(bg=self.BG)
        self.resizable(True, True)

        # State
        self._run_thread  = None
        self._log_queue   = queue.Queue()
        self._running     = False
        self._run_logger  = None

        # Config vars
        self._sheet_path = tk.StringVar()
        self._headless   = tk.BooleanVar(value=False)
        self._debug      = tk.BooleanVar(value=False)

        self._load_config()
        self._build_ui()
        self._poll_log_queue()

    # ── CONFIG ─────────────────────────────────────────────────────────
    def _load_config(self):
        if Path(self.CONFIG_FILE).exists():
            with open(self.CONFIG_FILE) as f:
                cfg = json.load(f)
            self._sheet_path.set(cfg.get("sheet_path", r"D:\TATA_SHIPMENTS.xlsx"))
            self._headless.set(cfg.get("headless", False))
            self._debug.set(cfg.get("debug", False))
        else:
            self._sheet_path.set(r"D:\TATA_SHIPMENTS.xlsx")

    def _save_config(self):
        if Path(self.CONFIG_FILE).exists():
            with open(self.CONFIG_FILE) as f:
                cfg = json.load(f)
        else:
            cfg = {}
        cfg["sheet_path"] = self._sheet_path.get()
        cfg["headless"]   = self._headless.get()
        cfg["debug"]      = self._debug.get()
        with open(self.CONFIG_FILE, "w") as f:
            json.dump(cfg, f, indent=2)

    # ── BUILD UI ───────────────────────────────────────────────────────
    def _build_ui(self):
        # ── Header ────────────────────────────────────────────────────
        header = tk.Frame(self, bg=self.ACCENT, height=56)
        header.pack(fill="x")
        header.pack_propagate(False)

        tk.Label(
            header,
            text="⚡  Portal Automation Engine",
            bg=self.ACCENT, fg="white",
            font=("Segoe UI", 14, "bold"),
            pady=14,
        ).pack(side="left", padx=20)

        self._status_label = tk.Label(
            header, text="● Idle",
            bg=self.ACCENT, fg="#C8A8E9",
            font=("Segoe UI", 10),
        )
        self._status_label.pack(side="right", padx=20)

        # ── Config panel ──────────────────────────────────────────────
        panel = tk.Frame(self, bg=self.PANEL_BG, padx=20, pady=16)
        panel.pack(fill="x", padx=16, pady=(12, 0))

        tk.Label(
            panel, text="SETTINGS",
            bg=self.PANEL_BG, fg=self.TEXT_MUTED,
            font=("Segoe UI", 8, "bold"),
        ).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 8))

        # Sheet path
        tk.Label(
            panel, text="Sheet Path",
            bg=self.PANEL_BG, fg=self.TEXT,
            font=("Segoe UI", 10),
        ).grid(row=1, column=0, sticky="w", pady=4)

        sheet_entry = tk.Entry(
            panel,
            textvariable=self._sheet_path,
            font=("Segoe UI", 10),
            width=52,
            relief="solid",
            bd=1,
        )
        sheet_entry.grid(row=1, column=1, sticky="ew", padx=(8, 6), pady=4)

        tk.Button(
            panel, text="Browse",
            command=self._browse_sheet,
            bg=self.ACCENT, fg="white",
            font=("Segoe UI", 9),
            relief="flat",
            padx=10, pady=4,
            cursor="hand2",
            activebackground=self.ACCENT_DARK,
            activeforeground="white",
        ).grid(row=1, column=2, pady=4)

        # Toggles row
        toggle_frame = tk.Frame(panel, bg=self.PANEL_BG)
        toggle_frame.grid(row=2, column=0, columnspan=3, sticky="w", pady=(6, 0))

        tk.Checkbutton(
            toggle_frame,
            text="Run headless (hide browser)",
            variable=self._headless,
            bg=self.PANEL_BG, fg=self.TEXT,
            font=("Segoe UI", 10),
            activebackground=self.PANEL_BG,
            selectcolor=self.PANEL_BG,
        ).pack(side="left", padx=(0, 20))

        tk.Checkbutton(
            toggle_frame,
            text="Debug mode",
            variable=self._debug,
            bg=self.PANEL_BG, fg=self.TEXT,
            font=("Segoe UI", 10),
            activebackground=self.PANEL_BG,
            selectcolor=self.PANEL_BG,
        ).pack(side="left")

        panel.columnconfigure(1, weight=1)

        # ── Run / Stop buttons ────────────────────────────────────────
        btn_frame = tk.Frame(self, bg=self.BG)
        btn_frame.pack(fill="x", padx=16, pady=10)

        self._run_btn = tk.Button(
            btn_frame,
            text="▶  RUN",
            command=self._start_run,
            bg=self.ACCENT, fg="white",
            font=("Segoe UI", 11, "bold"),
            relief="flat",
            padx=28, pady=8,
            cursor="hand2",
            activebackground=self.ACCENT_DARK,
            activeforeground="white",
        )
        self._run_btn.pack(side="left")

        self._stop_btn = tk.Button(
            btn_frame,
            text="■  STOP",
            command=self._stop_run,
            bg="#C0392B", fg="white",
            font=("Segoe UI", 11, "bold"),
            relief="flat",
            padx=28, pady=8,
            cursor="hand2",
            state="disabled",
            activebackground="#922B21",
            activeforeground="white",
        )
        self._stop_btn.pack(side="left", padx=(8, 0))

        self._log_path_label = tk.Label(
            btn_frame,
            text="",
            bg=self.BG, fg=self.TEXT_MUTED,
            font=("Segoe UI", 9),
        )
        self._log_path_label.pack(side="right", padx=4)

        # ── Log window ────────────────────────────────────────────────
        log_frame = tk.Frame(self, bg=self.BG)
        log_frame.pack(fill="both", expand=True, padx=16, pady=(0, 12))

        tk.Label(
            log_frame,
            text="RUN LOG",
            bg=self.BG, fg=self.TEXT_MUTED,
            font=("Segoe UI", 8, "bold"),
        ).pack(anchor="w", pady=(0, 4))

        self._log_box = scrolledtext.ScrolledText(
            log_frame,
            bg=self.LOG_BG,
            fg=self.LOG_FG,
            font=("Cascadia Code", 9) if self._font_exists("Cascadia Code")
                 else ("Consolas", 9),
            relief="flat",
            wrap="word",
            state="disabled",
            padx=10, pady=10,
        )
        self._log_box.pack(fill="both", expand=True)

        # Colour tags for log lines
        self._log_box.tag_config("ok",   foreground=self.LOG_OK)
        self._log_box.tag_config("warn", foreground=self.LOG_WARN)
        self._log_box.tag_config("err",  foreground=self.LOG_ERR)
        self._log_box.tag_config("info", foreground=self.LOG_INFO)
        self._log_box.tag_config("dim",  foreground="#666666")

        # Clear log button
        tk.Button(
            log_frame,
            text="Clear log",
            command=self._clear_log,
            bg=self.BG, fg=self.TEXT_MUTED,
            font=("Segoe UI", 8),
            relief="flat",
            cursor="hand2",
        ).pack(anchor="e", pady=(4, 0))

    # ── HELPERS ────────────────────────────────────────────────────────
    def _font_exists(self, name):
        return name in font.families()

    def _browse_sheet(self):
        path = filedialog.askopenfilename(
            title="Select invoice sheet",
            filetypes=[
                ("CSV / Excel", "*.csv *.xlsx *.xls"),
                ("All files", "*.*"),
            ],
        )
        if path:
            self._sheet_path.set(path)

    def _set_status(self, text, colour="#C8A8E9"):
        self._status_label.config(text=text, fg=colour)

    def _clear_log(self):
        self._log_box.config(state="normal")
        self._log_box.delete("1.0", "end")
        self._log_box.config(state="disabled")

    # ── LOG WRITER ─────────────────────────────────────────────────────
    def _append_log(self, text: str):
        """Write a line to the log box with colour coding."""
        self._log_box.config(state="normal")

        # Pick tag based on content
        lower = text.lower()
        if any(k in lower for k in ("✅", "done", "complete", "saved", "clicked")):
            tag = "ok"
        elif any(k in lower for k in ("⚠️", "warning", "skipping", "skipped",
                                       "disabled", "fallback", "⏭")):
            tag = "warn"
        elif any(k in lower for k in ("❌", "error", "failed", "timeout")):
            tag = "err"
        elif any(k in lower for k in ("───", "===", "processing", "date from",
                                       "entries to", "navigating", "reading")):
            tag = "info"
        else:
            tag = None

        self._log_box.insert("end", text, tag)
        self._log_box.see("end")
        self._log_box.config(state="disabled")

    def _poll_log_queue(self):
        """Check queue every 50ms and flush to log box."""
        try:
            while True:
                text = self._log_queue.get_nowait()
                self._append_log(text)

                # Mirror to CSV run log with basic parsing
                if self._run_logger:
                    line = text.strip()
                    if "Processing:" in line:
                        # e.g. "Processing: IN123  [Cash  Rs.827]"
                        pass   # captured at fill_one_row level below
                    elif "✅ Done" in line or "✅ done" in line:
                        pass   # captured per-row
        except queue.Empty:
            pass
        finally:
            self.after(50, self._poll_log_queue)

    # ── RUN / STOP ─────────────────────────────────────────────────────
    def _start_run(self):
        if self._running:
            return

        sheet = self._sheet_path.get().strip()
        if not sheet or not Path(sheet).exists():
            messagebox.showerror(
                "Sheet not found",
                f"Cannot find sheet at:\n{sheet}\n\nPlease check the path."
            )
            return

        self._save_config()
        self._running = True
        self._run_btn.config(state="disabled")
        self._stop_btn.config(state="normal")
        self._set_status("● Running", "#90EE90")

        # New CSV log for this run
        self._run_logger = RunLogger()
        self._log_path_label.config(
            text=f"📄 {self._run_logger.path}"
        )
        self._run_logger.record(
            status="RUN STARTED",
            notes=f"Sheet: {sheet}"
        )

        self._append_log(
            f"\n{'─'*60}\n"
            f"  RUN STARTED  {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
            f"  Sheet: {sheet}\n"
            f"  Log:   {self._run_logger.path}\n"
            f"{'─'*60}\n"
        )

        # Run in background thread so GUI stays responsive
        self._run_thread = threading.Thread(
            target=self._run_engine, daemon=True
        )
        self._run_thread.start()

    def _run_engine(self):
        """Background thread — runs the full automation."""
        # Redirect stdout/stderr to the queue
        orig_stdout = sys.stdout
        orig_stderr = sys.stderr
        writer      = QueueWriter(self._log_queue)
        sys.stdout  = writer
        sys.stderr  = writer

        try:
            # Import engine here so errors surface in the log
            from main import (
                RetailerPortalEngine,
                load_config,
                load_rows_from_csv,
            )

            config = load_config()
            # Override with current GUI values
            config["sheet_path"] = self._sheet_path.get()
            config["headless"]   = self._headless.get()
            config["debug"]      = self._debug.get()

            engine = RetailerPortalEngine(config)

            engine.launch()
            engine.login()

            csv_path         = config["sheet_path"]
            rows, run_date   = load_rows_from_csv(csv_path)

            if not rows:
                print("⚠️  No processable rows found in sheet.")
            else:
                print(f"\nDate from sheet  : {run_date}")
                print(f"Entries to process: {len(rows)}")
                for r in rows:
                    print(f"  {r['invoice_no']}  {r['payment_mode']:8}  Rs.{r['amount']}")

                engine.navigate_to_settlement_page(run_date)

                for row in rows:
                    if not self._running:
                        print("\n⏹  Run stopped by user.")
                        self._run_logger.record(status="STOPPED", notes="User stopped run")
                        break
                    try:
                        engine.fill_one_row(
                            invoice_no   = row["invoice_no"],
                            payment_mode = row["payment_mode"],
                            cheque_no    = row.get("cheque_no", ""),
                            amount       = row["amount"],
                            date         = row["date"],
                        )
                        self._run_logger.record(
                            invoice_no   = row["invoice_no"],
                            payment_mode = row["payment_mode"],
                            amount       = row["amount"],
                            status       = "✅ saved",
                        )
                    except Exception as row_err:
                        print(f"  ❌ Row error: {row_err}")
                        self._run_logger.record(
                            invoice_no   = row["invoice_no"],
                            payment_mode = row["payment_mode"],
                            amount       = row["amount"],
                            status       = "❌ error",
                            notes        = str(row_err),
                        )

                if self._running:
                    print("\n✅ All entries processed.")
                    print("\n" + "="*55)
                    print("  ⏸  Review the table in the browser.")
                    print("  Final page Save will click automatically in 5 seconds...")
                    print("="*55)
                    # Give user 5s to look before final save
                    import time
                    for i in range(5, 0, -1):
                        if not self._running:
                            break
                        print(f"  Saving in {i}s...")
                        time.sleep(1)

                    if self._running:
                        engine.save_page()
                        self._run_logger.record(
                            status="✅ FINAL SAVE",
                            notes="Page-level save complete"
                        )

            engine.close()
            print("\n✅ Run complete.")
            self._run_logger.record(
                status="RUN COMPLETE",
                notes=f"Log saved: {self._run_logger.path}"
            )

        except Exception as e:
            import traceback
            print(f"\n❌ Fatal error: {e}")
            print(traceback.format_exc())
            if self._run_logger:
                self._run_logger.record(
                    status="❌ FATAL ERROR",
                    notes=str(e)
                )
        finally:
            sys.stdout = orig_stdout
            sys.stderr = orig_stderr
            self._running = False
            self.after(0, self._on_run_finished)

    def _stop_run(self):
        if self._running:
            self._running = False
            self._append_log("\n⏹  Stop requested — will halt after current row.\n")

    def _on_run_finished(self):
        self._running = False
        self._run_btn.config(state="normal")
        self._stop_btn.config(state="disabled")
        self._set_status("● Idle")
        if self._run_logger:
            self._append_log(
                f"\n📄 Run log saved: {self._run_logger.path}\n"
                f"   Send this file to your developer for review.\n"
            )


# ──────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app = PortalAutomationApp()
    app.mainloop()