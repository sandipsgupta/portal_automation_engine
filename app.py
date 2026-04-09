"""
app.py — Tkinter GUI for Portal Automation Engine
"""

import csv
import json
import os
import queue
import sys
import threading
import time
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, font, messagebox, scrolledtext

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from main import RetailerPortalEngine, load_config, load_rows_from_csv


# ──────────────────────────────────────────────────────────────────────
# PORTAL DEFINITIONS
# active=True  → selectable, engine wired up
# active=False → greyed out, coming soon
# otp=True     → shows OTP badge
# ──────────────────────────────────────────────────────────────────────
PORTALS = [
    {
        "id":     "tata_consumer",
        "label":  "Tata Consumer",
        "sub":    "Non-OTP  •  Active",
        "active": True,
        "otp":    False,
    },
    {
        "id":     "portal_2",
        "label":  "Portal 2",
        "sub":    "Non-OTP  •  Coming Soon",
        "active": False,
        "otp":    False,
    },
    {
        "id":     "portal_3",
        "label":  "Portal 3",
        "sub":    "OTP Login  •  Coming Soon",
        "active": False,
        "otp":    True,
    },
    {
        "id":     "portal_4",
        "label":  "Portal 4",
        "sub":    "OTP Login  •  Coming Soon",
        "active": False,
        "otp":    True,
    },
]


# ──────────────────────────────────────────────────────────────────────
# CSV RUN LOGGER
# ──────────────────────────────────────────────────────────────────────
class RunLogger:
    COLUMNS = ["timestamp", "invoice_no", "payment_mode", "amount", "status", "notes"]

    def __init__(self, portal_id="portal"):
        # When packaged as .exe logs save next to the .exe
        # When running as script logs save next to app.py
        import sys as _sys
        _base = Path(_sys.executable).parent if getattr(_sys, 'frozen', False) \
                else Path(__file__).parent
        self.log_dir = _base / "logs"
        self.log_dir.mkdir(exist_ok=True)
        ts            = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.log_path = self.log_dir / f"run_{portal_id}_{ts}.csv"
        self._write_header()

    def _write_header(self):
        with open(self.log_path, "w", newline="", encoding="utf-8") as f:
            csv.DictWriter(f, fieldnames=self.COLUMNS).writeheader()

    def record(self, invoice_no="", payment_mode="", amount="",
               status="", notes=""):
        with open(self.log_path, "a", newline="", encoding="utf-8") as f:
            csv.DictWriter(f, fieldnames=self.COLUMNS).writerow({
                "timestamp":    datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "invoice_no":   invoice_no,
                "payment_mode": payment_mode,
                "amount":       amount,
                "status":       status,
                "notes":        notes,
            })

    @property
    def path(self):
        return str(self.log_path)


# ──────────────────────────────────────────────────────────────────────
# STDOUT → queue
# ──────────────────────────────────────────────────────────────────────
class QueueWriter:
    def __init__(self, q):
        self.queue = q

    def write(self, text):
        if text:
            self.queue.put(text)

    def flush(self):
        pass


# ──────────────────────────────────────────────────────────────────────
# GUI helpers
# ──────────────────────────────────────────────────────────────────────
class ToggleSwitch(tk.Canvas):
    """A pill-shaped toggle switch bound to a tk.BooleanVar."""

    W, H, R = 44, 24, 11   # width, height, knob radius

    def __init__(self, parent, variable, on_color="#5B2D8E",
                 off_color="#CCCCCC", bg="#FFFFFF", **kw):
        super().__init__(parent, width=self.W, height=self.H,
                         bg=bg, highlightthickness=0, **kw)
        self._var      = variable
        self._on_color = on_color
        self._off_color = off_color
        self._draw()
        self._var.trace_add("write", lambda *_: self._draw())
        self.bind("<Button-1>", lambda _: self._var.set(not self._var.get()))

    def _draw(self):
        self.delete("all")
        on   = self._var.get()
        fill = self._on_color if on else self._off_color
        r, w, h = self.R, self.W, self.H
        # pill background
        self.create_oval(0, 0, h, h, fill=fill, outline="")
        self.create_oval(w - h, 0, w, h, fill=fill, outline="")
        self.create_rectangle(h // 2, 0, w - h // 2, h, fill=fill, outline="")
        # knob
        kx = w - h // 2 - 1 if on else h // 2 + 1
        ky = h // 2
        self.create_oval(kx - r, ky - r, kx + r, ky + r,
                         fill="white", outline="")


# ──────────────────────────────────────────────────────────────────────
class PortalAutomationApp(tk.Tk):

    CONFIG_FILE = "config.json"

    # Colours
    BG          = "#F0F0F0"
    PANEL_BG    = "#FFFFFF"
    ACCENT      = "#5B2D8E"
    ACCENT_DARK = "#3D1A6B"
    ACCENT_LITE = "#EDE7F6"
    TEXT        = "#1A1A1A"
    TEXT_MUTED  = "#757575"
    DISABLED_BG = "#FAFAFA"
    DISABLED_FG = "#BBBBBB"
    BORDER      = "#E0E0E0"
    LOG_BG      = "#1E1E1E"
    LOG_FG      = "#D4D4D4"
    LOG_OK      = "#4EC9B0"
    LOG_WARN    = "#DCDCAA"
    LOG_ERR     = "#F44747"
    LOG_INFO    = "#9CDCFE"

    def __init__(self):
        super().__init__()
        self.title("Portal Automation Engine")
        self.geometry("880x780")
        self.minsize(740, 640)
        self.configure(bg=self.BG)
        self.resizable(True, True)

        self._run_thread      = None
        self._log_queue       = queue.Queue()
        self._running         = False
        self._run_logger      = None
        self._log_text_buffer = []   # full text for TXT export

        # Config vars
        self._sheet_path      = tk.StringVar()
        self._headless        = tk.BooleanVar(value=False)
        self._debug           = tk.BooleanVar(value=False)
        self._selected_portal = tk.StringVar(value=PORTALS[0]["id"])

        self._load_config()
        self._build_ui()
        self._poll_log_queue()

    # ── CONFIG ─────────────────────────────────────────────────────────
    def _load_config(self):
        if Path(self.CONFIG_FILE).exists():
            with open(self.CONFIG_FILE) as f:
                cfg = json.load(f)
            self._sheet_path.set(cfg.get("sheet_path", ""))
            self._headless.set(cfg.get("headless", False))
            self._debug.set(cfg.get("debug", False))
            self._selected_portal.set(
                cfg.get("selected_portal", PORTALS[0]["id"])
            )
        else:
            self._sheet_path.set("")

    def _save_config(self):
        cfg = {}
        if Path(self.CONFIG_FILE).exists():
            with open(self.CONFIG_FILE) as f:
                cfg = json.load(f)
        cfg["sheet_path"]      = self._sheet_path.get()
        cfg["headless"]        = self._headless.get()
        cfg["debug"]           = self._debug.get()
        cfg["selected_portal"] = self._selected_portal.get()
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
        ).pack(side="left", padx=20, pady=14)

        self._status_label = tk.Label(
            header, text="● Idle",
            bg=self.ACCENT, fg="#C8A8E9",
            font=("Segoe UI", 10),
        )
        self._status_label.pack(side="right", padx=20)

        # ── Portal selector ───────────────────────────────────────────
        portal_outer = tk.Frame(self, bg=self.BG, padx=16, pady=12)
        portal_outer.pack(fill="x")

        tk.Label(
            portal_outer,
            text="SELECT PORTAL",
            bg=self.BG, fg=self.TEXT_MUTED,
            font=("Segoe UI", 8, "bold"),
        ).pack(anchor="w", pady=(0, 8))

        cards_frame = tk.Frame(portal_outer, bg=self.BG)
        cards_frame.pack(fill="x")

        for i, portal in enumerate(PORTALS):
            self._build_portal_card(cards_frame, portal, i)
        cards_frame.columnconfigure((0, 1, 2, 3), weight=1, uniform="card")

        # ── Settings panel ────────────────────────────────────────────
        panel = tk.Frame(self, bg=self.PANEL_BG, padx=20, pady=14)
        panel.pack(fill="x", padx=16)

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

        tk.Entry(
            panel,
            textvariable=self._sheet_path,
            font=("Segoe UI", 10),
            width=52, relief="solid", bd=1,
        ).grid(row=1, column=1, sticky="ew", padx=(8, 6), pady=4)

        tk.Button(
            panel, text="Browse",
            command=self._browse_sheet,
            bg=self.ACCENT, fg="white",
            font=("Segoe UI", 9),
            relief="flat", padx=10, pady=4,
            cursor="hand2",
            activebackground=self.ACCENT_DARK,
            activeforeground="white",
        ).grid(row=1, column=2, pady=4)

        # Toggles
        toggle_frame = tk.Frame(panel, bg=self.PANEL_BG)
        toggle_frame.grid(row=2, column=0, columnspan=3, sticky="w", pady=(6, 0))

        # Headless toggle
        ToggleSwitch(
            toggle_frame, variable=self._headless,
            on_color=self.ACCENT, bg=self.PANEL_BG,
        ).pack(side="left", padx=(0, 6))
        tk.Label(
            toggle_frame, text="Run headless (hide browser)",
            bg=self.PANEL_BG, fg=self.TEXT, font=("Segoe UI", 10),
        ).pack(side="left", padx=(0, 20))

        # Debug toggle
        ToggleSwitch(
            toggle_frame, variable=self._debug,
            on_color=self.ACCENT, bg=self.PANEL_BG,
        ).pack(side="left", padx=(0, 6))
        tk.Label(
            toggle_frame, text="Debug mode",
            bg=self.PANEL_BG, fg=self.TEXT, font=("Segoe UI", 10),
        ).pack(side="left")

        panel.columnconfigure(1, weight=1)

        # ── Action buttons ────────────────────────────────────────────
        btn_frame = tk.Frame(self, bg=self.BG)
        btn_frame.pack(fill="x", padx=16, pady=10)

        self._run_btn = tk.Button(
            btn_frame, text="▶  RUN",
            command=self._start_run,
            bg=self.ACCENT, fg="white",
            font=("Segoe UI", 11, "bold"),
            relief="flat", padx=28, pady=8,
            cursor="hand2",
            activebackground=self.ACCENT_DARK,
            activeforeground="white",
        )
        self._run_btn.pack(side="left")

        self._stop_btn = tk.Button(
            btn_frame, text="■  STOP",
            command=self._stop_run,
            bg="#999999", fg="white",
            font=("Segoe UI", 11, "bold"),
            relief="flat", padx=28, pady=8,
            cursor="arrow", state="disabled",
            disabledforeground="white",
            activebackground="#C0392B",
            activeforeground="white",
        )
        self._stop_btn.pack(side="left", padx=(8, 0))

        # Export Log button
        self._export_btn = tk.Button(
            btn_frame, text="⬇  Export Log",
            command=self._export_log_txt,
            bg=self.PANEL_BG, fg=self.TEXT_MUTED,
            font=("Segoe UI", 9),
            relief="solid", bd=1,
            padx=12, pady=8,
            cursor="hand2",
            activebackground=self.BG,
        )
        self._export_btn.pack(side="left", padx=(16, 0))

        # CSV log path label
        self._log_path_label = tk.Label(
            btn_frame, text="",
            bg=self.BG, fg=self.TEXT_MUTED,
            font=("Segoe UI", 9),
        )
        self._log_path_label.pack(side="right", padx=4)

        # ── Log window ────────────────────────────────────────────────
        log_frame = tk.Frame(self, bg=self.BG)
        log_frame.pack(fill="both", expand=True, padx=16, pady=(0, 12))

        log_header = tk.Frame(log_frame, bg=self.BG)
        log_header.pack(fill="x", pady=(0, 4))

        tk.Label(
            log_header, text="RUN LOG",
            bg=self.BG, fg=self.TEXT_MUTED,
            font=("Segoe UI", 8, "bold"),
        ).pack(side="left")

        tk.Button(
            log_header, text="Clear",
            command=self._clear_log,
            bg=self.BG, fg=self.TEXT_MUTED,
            font=("Segoe UI", 8),
            relief="flat", cursor="hand2",
        ).pack(side="right")

        self._log_box = scrolledtext.ScrolledText(
            log_frame,
            bg=self.LOG_BG, fg=self.LOG_FG,
            font=("Cascadia Code", 9) if "Cascadia Code" in font.families()
                 else ("Consolas", 9),
            relief="flat", wrap="word",
            state="disabled",
            padx=10, pady=10,
        )
        self._log_box.pack(fill="both", expand=True)

        # Log colour tags
        self._log_box.tag_config("ok",   foreground=self.LOG_OK)
        self._log_box.tag_config("warn", foreground=self.LOG_WARN)
        self._log_box.tag_config("err",  foreground=self.LOG_ERR)
        self._log_box.tag_config("info", foreground=self.LOG_INFO)

    # ── PORTAL CARDS ───────────────────────────────────────────────────
    def _build_portal_card(self, parent, portal, col):
        """
        Each portal gets a card with a radio button.
        Active portals → selectable, white bg, purple on select.
        Inactive       → greyed out, not selectable, 'Coming Soon' badge.
        """
        is_active = portal["active"]

        card_bg     = self.PANEL_BG if is_active else self.DISABLED_BG
        fg_main     = self.TEXT     if is_active else self.DISABLED_FG
        fg_sub      = self.TEXT_MUTED if is_active else "#CCCCCC"
        relief      = "solid"
        bd          = 1

        card = tk.Frame(
            parent,
            bg=card_bg,
            relief=relief, bd=bd,
            padx=12, pady=10,
        )
        card.grid(row=0, column=col, sticky="nsew", padx=(0, 8))

        # Radio button — disabled for inactive portals
        rb = tk.Radiobutton(
            card,
            variable=self._selected_portal,
            value=portal["id"],
            bg=card_bg,
            activebackground=card_bg,
            selectcolor=self.ACCENT_LITE,
            state="normal" if is_active else "disabled",
            cursor="hand2" if is_active else "arrow",
        )
        rb.pack(anchor="w")

        # Portal name
        tk.Label(
            card, text=portal["label"],
            bg=card_bg, fg=fg_main,
            font=("Segoe UI", 10, "bold"),
        ).pack(anchor="w", pady=(2, 0))

        # Sub label
        tk.Label(
            card, text=portal["sub"],
            bg=card_bg, fg=fg_sub,
            font=("Segoe UI", 8),
        ).pack(anchor="w")

        # OTP badge for OTP portals
        if portal.get("otp"):
            tk.Label(
                card, text="  OTP  ",
                bg="#F0E6FF", fg=self.ACCENT,
                font=("Segoe UI", 7, "bold"),
                relief="flat", padx=4,
            ).pack(anchor="w", pady=(4, 0))

    # ── HELPERS ────────────────────────────────────────────────────────
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
        self._log_text_buffer.clear()

    # ── EXPORT LOG AS TXT ──────────────────────────────────────────────
    def _export_log_txt(self):
        if not self._log_text_buffer:
            messagebox.showinfo("Export Log", "No log content to export yet.")
            return

        default_name = f"run_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
        path = filedialog.asksaveasfilename(
            title="Save Log as Text File",
            defaultextension=".txt",
            initialfile=default_name,
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
        )
        if path:
            with open(path, "w", encoding="utf-8") as f:
                f.write("".join(self._log_text_buffer))
            messagebox.showinfo(
                "Export Log",
                f"Log saved to:\n{path}\n\nSend this file for support."
            )

    # ── LOG WRITER ─────────────────────────────────────────────────────
    def _append_log(self, text: str):
        self._log_text_buffer.append(text)
        self._log_box.config(state="normal")

        lower = text.lower()
        if any(k in lower for k in ("✅", "done", "complete", "saved", "clicked")):
            tag = "ok"
        elif any(k in lower for k in ("⚠️", "warning", "skipping", "skipped",
                                       "disabled", "fallback", "⏭")):
            tag = "warn"
        elif any(k in lower for k in ("❌", "error", "failed", "timeout")):
            tag = "err"
        elif any(k in lower for k in ("───", "===", "processing", "date from",
                                       "entries to", "navigating", "reading",
                                       "run started")):
            tag = "info"
        else:
            tag = None

        self._log_box.insert("end", text, tag)
        self._log_box.see("end")
        self._log_box.config(state="disabled")

    def _poll_log_queue(self):
        try:
            while True:
                self._append_log(self._log_queue.get_nowait())
        except queue.Empty:
            pass
        finally:
            self.after(50, self._poll_log_queue)

    # ── RUN / STOP ─────────────────────────────────────────────────────
    def _start_run(self):
        if self._running:
            return

        sheet = self._sheet_path.get().strip()
        if sheet and not Path(sheet).exists():
            messagebox.showerror(
                "Sheet not found",
                f"Cannot find sheet at:\n{sheet}\n\nPlease check the path and try again."
            )
            return

        self._save_config()
        self._running = True
        self._run_btn.config(state="disabled")
        self._stop_btn.config(
            state="normal", bg="#C0392B", fg="white", cursor="hand2"
        )
        self._set_status("● Running", "#90EE90")

        portal_id        = self._selected_portal.get()
        self._run_logger = RunLogger(portal_id)
        self._log_path_label.config(text=f"📄 {self._run_logger.path}")
        self._run_logger.record(status="RUN STARTED", notes=f"Sheet: {sheet}")

        self._append_log(
            f"\n{'─'*60}\n"
            f"  RUN STARTED  {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
            f"  Portal : {portal_id}\n"
            f"  Sheet  : {sheet}\n"
            f"  CSV Log: {self._run_logger.path}\n"
            f"{'─'*60}\n"
        )

        self._run_thread = threading.Thread(target=self._run_engine, daemon=True)
        self._run_thread.start()

    def _run_engine(self):
        orig_stdout, orig_stderr = sys.stdout, sys.stderr
        writer = QueueWriter(self._log_queue)
        sys.stdout = sys.stderr = writer

        try:
            config               = load_config()
            config["sheet_path"] = self._sheet_path.get()
            config["headless"]   = self._headless.get()
            config["debug"]      = self._debug.get()

            engine = RetailerPortalEngine(config)
            try:
                engine.launch()
                engine.login()

                rows, from_date, to_date = load_rows_from_csv(config["sheet_path"])

                if not rows:
                    print("⚠️  No processable rows found in sheet.")
                else:
                    if from_date == to_date:
                        print(f"\nDate from sheet   : {from_date}")
                    else:
                        print(f"\nDate range        : {from_date}  →  {to_date}")
                    print(f"Entries to process: {len(rows)}")
                    for r in rows:
                        print(f"  {r['invoice_no']}  {r['payment_mode']:8}  Rs.{r['amount']}")

                    engine.navigate_to_settlement_page(from_date, to_date)

                    # Group flat payment list by invoice_no (preserving order)
                    invoices = {}
                    for row in rows:
                        invoices.setdefault(row["invoice_no"], []).append(row)

                    for invoice_no, payments in invoices.items():
                        if not self._running:
                            print("\n⏹  Run stopped by user.")
                            self._run_logger.record(status="STOPPED")
                            break
                        try:
                            engine.fill_invoice(invoice_no, payments)
                            for p in payments:
                                self._run_logger.record(
                                    invoice_no   = invoice_no,
                                    payment_mode = p["payment_mode"],
                                    amount       = p["amount"],
                                    status       = "✅ saved",
                                )
                        except Exception as inv_err:
                            print(f"  ❌ Invoice error: {inv_err}")
                            for p in payments:
                                self._run_logger.record(
                                    invoice_no   = invoice_no,
                                    payment_mode = p["payment_mode"],
                                    amount       = p["amount"],
                                    status       = "❌ error",
                                    notes        = str(inv_err),
                                )

                    if self._running:
                        print("\n✅ All entries processed.")
                        print("\n" + "="*55)
                        print("  ⏸  Review the table in the browser.")
                        print("  Final Save in 5 seconds...")
                        print("="*55)
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
            finally:
                engine.close()

            print("\n✅ Run complete.")
            self._run_logger.record(
                status="RUN COMPLETE",
                notes=f"CSV log: {self._run_logger.path}"
            )

        except Exception as e:
            import traceback
            print(f"\n❌ Fatal error: {e}")
            print(traceback.format_exc())
            if self._run_logger:
                self._run_logger.record(status="❌ FATAL ERROR", notes=str(e))
        finally:
            sys.stdout, sys.stderr = orig_stdout, orig_stderr
            self._running = False
            self.after(0, self._on_run_finished)

    def _stop_run(self):
        if self._running:
            self._running = False
            self._append_log("\n⏹  Stop requested — halting after current row.\n")

    def _on_run_finished(self):
        self._running = False
        self._run_btn.config(state="normal")
        self._stop_btn.config(
            state="disabled", bg="#999999", fg="white", cursor="arrow"
        )
        self._set_status("● Idle")
        if self._run_logger:
            self._append_log(
                f"\n📄 CSV log: {self._run_logger.path}\n"
                f"   Use '⬇ Export Log' to save the full text log.\n"
            )


# ──────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app = PortalAutomationApp()
    app.mainloop()