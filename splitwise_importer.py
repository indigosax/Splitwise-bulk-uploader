"""
Splitwise CSV Importer
A Windows GUI application to import expenses from CSV into Splitwise.
Author: Built for cnickerson@dfirent.com
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import csv
import json
import os
import threading
import time
import webbrowser
from datetime import datetime

try:
    import requests
except ImportError:
    import subprocess, sys
    subprocess.check_call([sys.executable, "-m", "pip", "install", "requests"])
    import requests

# ─────────────────────────────────────────────
# CONSTANTS
# ─────────────────────────────────────────────
API_BASE = "https://secure.splitwise.com/api/v3.0"
SETTINGS_FILE = os.path.join(os.path.expanduser("~"), ".splitwise_importer_settings.json")

TEAL       = "#01696F"
TEAL_DARK  = "#0C4E54"
BG         = "#F7F6F2"
SURFACE    = "#F9F8F5"
BORDER     = "#D4D1CA"
TEXT       = "#28251D"
TEXT_MUTED = "#7A7974"
SUCCESS    = "#437A22"
ERROR      = "#A12C7B"
WARNING    = "#964219"
WHITE      = "#FFFFFF"

FONT_H1    = ("Segoe UI", 16, "bold")
FONT_H2    = ("Segoe UI", 12, "bold")
FONT_BODY  = ("Segoe UI", 10)
FONT_SMALL = ("Segoe UI", 9)
FONT_MONO  = ("Consolas", 9)

# Standard CSV column names we recognize automatically
KNOWN_ALIASES = {
    "date":        ["date", "transaction date", "trans date", "posted date", "transaction_date"],
    "description": ["description", "desc", "memo", "note", "merchant", "name", "payee", "details"],
    "cost":        ["cost", "amount", "total", "price", "charge", "debit", "transaction amount"],
    "currency":    ["currency", "currency code", "ccy"],
    "category":    ["category", "cat", "type", "expense type"],
    "group_id":    ["group_id", "group id", "group", "splitwise group"],
    "notes":       ["notes", "note", "comment", "comments", "memo2"],
}


# ─────────────────────────────────────────────
# SETTINGS PERSISTENCE
# ─────────────────────────────────────────────
def load_settings():
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, "r") as f:
                return json.load(f)
        except Exception:
            pass
    return {}

def save_settings(data: dict):
    try:
        with open(SETTINGS_FILE, "w") as f:
            json.dump(data, f, indent=2)
    except Exception:
        pass


# ─────────────────────────────────────────────
# SPLITWISE API CLIENT
# ─────────────────────────────────────────────
class SplitwiseClient:
    def __init__(self, api_key: str):
        self.api_key = api_key.strip()
        self.session = requests.Session()
        self.session.headers.update({
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/x-www-form-urlencoded",
            "Accept": "application/json",
        })

    def get_current_user(self):
        r = self.session.get(f"{API_BASE}/get_current_user", timeout=10)
        r.raise_for_status()
        return r.json().get("user", {})

    def get_groups(self):
        r = self.session.get(f"{API_BASE}/get_groups", timeout=10)
        r.raise_for_status()
        groups = r.json().get("groups", [])
        return [{"id": g["id"], "name": g["name"]} for g in groups]

    def get_friends(self):
        r = self.session.get(f"{API_BASE}/get_friends", timeout=10)
        r.raise_for_status()
        friends = r.json().get("friends", [])
        result = []
        for f in friends:
            name = f"{f.get('first_name','')} {f.get('last_name','')}".strip()
            result.append({"id": f["id"], "name": name, "email": f.get("email","")})
        return result

    def create_expense(self, payload: dict):
        r = self.session.post(f"{API_BASE}/create_expense", data=payload, timeout=15)
        data = r.json()
        if r.status_code not in (200, 201):
            raise Exception(data.get("errors", {}).get("base", [str(r.status_code)])[0])
        errors = data.get("errors", {})
        if errors:
            base_errs = errors.get("base", [])
            if base_errs:
                raise Exception(base_errs[0])
        return data.get("expenses", [{}])[0]


# ─────────────────────────────────────────────
# TOOLTIP HELPER
# ─────────────────────────────────────────────
class ToolTip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tip = None
        widget.bind("<Enter>", self.show)
        widget.bind("<Leave>", self.hide)

    def show(self, _=None):
        x = self.widget.winfo_rootx() + 20
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 4
        self.tip = tk.Toplevel(self.widget)
        self.tip.wm_overrideredirect(True)
        self.tip.wm_geometry(f"+{x}+{y}")
        lbl = tk.Label(self.tip, text=self.text, font=FONT_SMALL,
                       bg="#28251D", fg=WHITE, padx=8, pady=4, wraplength=280,
                       justify="left")
        lbl.pack()

    def hide(self, _=None):
        if self.tip:
            self.tip.destroy()
            self.tip = None


# ─────────────────────────────────────────────
# MAIN APPLICATION
# ─────────────────────────────────────────────
class SplitwiseImporterApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Splitwise CSV Importer")
        self.configure(bg=BG)
        self.resizable(True, True)
        self.minsize(860, 640)

        self.settings = load_settings()
        self.client: SplitwiseClient | None = None
        self.groups: list = []
        self.friends: list = []
        self.csv_rows: list = []
        self.csv_headers: list = []
        self.column_vars: dict = {}   # field_name -> StringVar (selected CSV column)
        self.import_results: list = []

        self._build_ui()
        self._restore_settings()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ── UI CONSTRUCTION ──────────────────────
    def _build_ui(self):
        # ── Header bar ──
        header = tk.Frame(self, bg=TEAL)
        header.pack(fill="x")
        tk.Label(header, text="  Splitwise CSV Importer", font=("Segoe UI", 14, "bold"),
                 bg=TEAL, fg=WHITE, pady=10).pack(side="left")
        tk.Label(header, text="v1.0  |  dfirent.com",
                 font=FONT_SMALL, bg=TEAL, fg="#B2D8DA", padx=12).pack(side="right")

        # ── Notebook (tabs) ──
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("TNotebook", background=BG, borderwidth=0)
        style.configure("TNotebook.Tab", font=FONT_BODY, padding=[14, 6],
                        background=SURFACE, foreground=TEXT_MUTED)
        style.map("TNotebook.Tab",
                  background=[("selected", WHITE)],
                  foreground=[("selected", TEAL)])
        style.configure("TFrame", background=BG)

        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True, padx=0, pady=0)

        self.tab_auth    = ttk.Frame(self.nb)
        self.tab_import  = ttk.Frame(self.nb)
        self.tab_results = ttk.Frame(self.nb)
        self.tab_log     = ttk.Frame(self.nb)

        self.nb.add(self.tab_auth,    text="  🔑  Authentication  ")
        self.nb.add(self.tab_import,  text="  📂  Import CSV  ")
        self.nb.add(self.tab_results, text="  📊  Results  ")
        self.nb.add(self.tab_log,     text="  📋  Log  ")

        self._build_auth_tab()
        self._build_import_tab()
        self._build_results_tab()
        self._build_log_tab()

        # ── Status bar ──
        status_bar = tk.Frame(self, bg=SURFACE, bd=0, relief="flat")
        status_bar.pack(fill="x", side="bottom")
        tk.Frame(status_bar, bg=BORDER, height=1).pack(fill="x")
        self.status_var = tk.StringVar(value="Ready — enter your API key and connect.")
        self.status_label = tk.Label(status_bar, textvariable=self.status_var,
                                     font=FONT_SMALL, bg=SURFACE, fg=TEXT_MUTED,
                                     anchor="w", padx=12, pady=6)
        self.status_label.pack(fill="x")

    # ── AUTH TAB ─────────────────────────────
    def _build_auth_tab(self):
        pad = {"padx": 32, "pady": 10}
        f = self.tab_auth

        tk.Label(f, text="Connect to Splitwise", font=FONT_H1,
                 bg=BG, fg=TEXT).pack(anchor="w", **pad)
        tk.Label(f,
                 text="Paste your Splitwise OAuth2 access token or API key below.\n"
                      "You can generate one at: https://secure.splitwise.com/apps",
                 font=FONT_BODY, bg=BG, fg=TEXT_MUTED, justify="left").pack(anchor="w", padx=32)

        # API key frame
        key_frame = tk.Frame(f, bg=BG)
        key_frame.pack(fill="x", padx=32, pady=(8, 0))

        tk.Label(key_frame, text="API Key / Access Token:", font=FONT_BODY,
                 bg=BG, fg=TEXT).pack(anchor="w")

        entry_row = tk.Frame(key_frame, bg=BG)
        entry_row.pack(fill="x", pady=(4, 0))

        self.api_key_var = tk.StringVar()
        self.api_key_entry = tk.Entry(entry_row, textvariable=self.api_key_var,
                                      font=FONT_MONO, show="•", width=60,
                                      bg=WHITE, fg=TEXT, relief="flat",
                                      highlightthickness=1, highlightbackground=BORDER,
                                      highlightcolor=TEAL, insertbackground=TEXT)
        self.api_key_entry.pack(side="left", ipady=6, ipadx=4)

        self.show_key_btn = tk.Button(entry_row, text="Show", font=FONT_SMALL,
                                      bg=SURFACE, fg=TEXT_MUTED, relief="flat",
                                      activebackground=BORDER, cursor="hand2",
                                      command=self._toggle_key_visibility, padx=8)
        self.show_key_btn.pack(side="left", padx=(6, 0))

        ToolTip(self.api_key_entry, "Your Splitwise OAuth2 Bearer token or legacy API key.\n"
                                     "Never share this. It's stored locally in your user profile.")

        btn_row = tk.Frame(f, bg=BG)
        btn_row.pack(anchor="w", padx=32, pady=12)

        self.connect_btn = tk.Button(btn_row, text="Connect & Verify",
                                     font=FONT_BODY, bg=TEAL, fg=WHITE,
                                     activebackground=TEAL_DARK, activeforeground=WHITE,
                                     relief="flat", cursor="hand2", padx=16, pady=7,
                                     command=self._connect)
        self.connect_btn.pack(side="left")

        tk.Button(btn_row, text="Open Splitwise Dev Portal",
                  font=FONT_BODY, bg=SURFACE, fg=TEAL,
                  activebackground=BORDER, relief="flat", cursor="hand2",
                  padx=16, pady=7,
                  command=lambda: webbrowser.open("https://secure.splitwise.com/apps")
                  ).pack(side="left", padx=(10, 0))

        # Connection status card
        self.auth_status_frame = tk.Frame(f, bg=SURFACE, relief="flat",
                                          highlightthickness=1, highlightbackground=BORDER)
        self.auth_status_frame.pack(fill="x", padx=32, pady=16)

        self.auth_status_label = tk.Label(self.auth_status_frame,
                                          text="⚪  Not connected",
                                          font=FONT_BODY, bg=SURFACE, fg=TEXT_MUTED,
                                          anchor="w", padx=16, pady=12)
        self.auth_status_label.pack(fill="x")

        # Default split settings
        sep = tk.Frame(f, bg=BORDER, height=1)
        sep.pack(fill="x", padx=32, pady=(4, 12))

        tk.Label(f, text="Default Split Settings", font=FONT_H2, bg=BG, fg=TEXT).pack(anchor="w", padx=32)

        split_frame = tk.Frame(f, bg=BG)
        split_frame.pack(fill="x", padx=32, pady=(8, 0))

        # Default group
        tk.Label(split_frame, text="Default Group:", font=FONT_BODY, bg=BG, fg=TEXT,
                 width=18, anchor="w").grid(row=0, column=0, pady=6, sticky="w")
        self.default_group_var = tk.StringVar(value="No Group (0)")
        self.group_combo = ttk.Combobox(split_frame, textvariable=self.default_group_var,
                                        width=40, state="readonly", font=FONT_BODY)
        self.group_combo["values"] = ["No Group (0)"]
        self.group_combo.grid(row=0, column=1, padx=(8, 0), sticky="w")

        # Split equally
        tk.Label(split_frame, text="Split Equally:", font=FONT_BODY, bg=BG, fg=TEXT,
                 width=18, anchor="w").grid(row=1, column=0, pady=6, sticky="w")
        self.split_equally_var = tk.BooleanVar(value=True)
        chk = tk.Checkbutton(split_frame, variable=self.split_equally_var,
                              bg=BG, activebackground=BG, cursor="hand2")
        chk.grid(row=1, column=1, padx=(8, 0), sticky="w")
        ToolTip(chk, "When checked, the expense is split evenly among all group members.\n"
                     "Uncheck to assign as 100% paid by you, 0 owed by others.")

        # Currency
        tk.Label(split_frame, text="Default Currency:", font=FONT_BODY, bg=BG, fg=TEXT,
                 width=18, anchor="w").grid(row=2, column=0, pady=6, sticky="w")
        self.default_currency_var = tk.StringVar(value="USD")
        currency_entry = tk.Entry(split_frame, textvariable=self.default_currency_var,
                                  width=10, font=FONT_BODY, bg=WHITE, fg=TEXT, relief="flat",
                                  highlightthickness=1, highlightbackground=BORDER)
        currency_entry.grid(row=2, column=1, padx=(8, 0), sticky="w", ipady=4)
        ToolTip(currency_entry, "ISO 4217 currency code: USD, EUR, GBP, CAD, JPY, etc.")

    # ── IMPORT TAB ───────────────────────────
    def _build_import_tab(self):
        f = self.tab_import
        pad = {"padx": 32, "pady": 8}

        tk.Label(f, text="Import CSV File", font=FONT_H1, bg=BG, fg=TEXT).pack(anchor="w", **pad)
        tk.Label(f, text="Load a CSV file, map columns to Splitwise fields, preview, then import.",
                 font=FONT_BODY, bg=BG, fg=TEXT_MUTED).pack(anchor="w", padx=32, pady=(0, 8))

        # File picker row
        file_row = tk.Frame(f, bg=BG)
        file_row.pack(fill="x", padx=32, pady=(0, 4))

        self.file_path_var = tk.StringVar(value="No file selected")
        path_entry = tk.Entry(file_row, textvariable=self.file_path_var,
                              font=FONT_SMALL, state="readonly", width=62,
                              bg=SURFACE, fg=TEXT_MUTED, relief="flat",
                              highlightthickness=1, highlightbackground=BORDER,
                              readonlybackground=SURFACE)
        path_entry.pack(side="left", ipady=5, ipadx=4)

        tk.Button(file_row, text="Browse…", font=FONT_BODY,
                  bg=TEAL, fg=WHITE, activebackground=TEAL_DARK,
                  relief="flat", cursor="hand2", padx=12, pady=4,
                  command=self._browse_csv).pack(side="left", padx=(8, 0))

        # Column mapping
        self.mapping_frame = tk.LabelFrame(f, text=" Column Mapping ", font=FONT_SMALL,
                                           bg=BG, fg=TEXT_MUTED, padx=16, pady=10,
                                           relief="flat", bd=1,
                                           highlightthickness=1, highlightbackground=BORDER)
        self.mapping_frame.pack(fill="x", padx=32, pady=8)

        tk.Label(self.mapping_frame, text="Load a CSV file to configure column mapping.",
                 font=FONT_BODY, bg=BG, fg=TEXT_MUTED).pack(anchor="w")

        # Preview table
        preview_label_row = tk.Frame(f, bg=BG)
        preview_label_row.pack(fill="x", padx=32, pady=(4, 0))
        tk.Label(preview_label_row, text="Preview (first 10 rows)", font=FONT_H2,
                 bg=BG, fg=TEXT).pack(side="left")
        self.row_count_label = tk.Label(preview_label_row, text="",
                                        font=FONT_SMALL, bg=BG, fg=TEXT_MUTED)
        self.row_count_label.pack(side="left", padx=12)

        table_frame = tk.Frame(f, bg=BG)
        table_frame.pack(fill="both", expand=True, padx=32, pady=(4, 4))

        self.preview_tree = ttk.Treeview(table_frame, show="headings",
                                         selectmode="extended", height=8)
        vsb = ttk.Scrollbar(table_frame, orient="vertical",
                             command=self.preview_tree.yview)
        hsb = ttk.Scrollbar(table_frame, orient="horizontal",
                             command=self.preview_tree.xview)
        self.preview_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")
        self.preview_tree.pack(fill="both", expand=True)

        # Action buttons
        action_row = tk.Frame(f, bg=BG)
        action_row.pack(fill="x", padx=32, pady=(4, 12))

        self.dry_run_var = tk.BooleanVar(value=True)
        dry_chk = tk.Checkbutton(action_row, text="Dry Run (preview only — no data sent)",
                                  variable=self.dry_run_var, font=FONT_BODY,
                                  bg=BG, fg=TEXT, activebackground=BG, cursor="hand2",
                                  selectcolor=BG)
        dry_chk.pack(side="left")
        ToolTip(dry_chk, "Dry Run validates your CSV and mapping without posting any expenses.\n"
                          "Uncheck this to actually import into Splitwise.")

        self.import_btn = tk.Button(action_row, text="▶  Run Import",
                                    font=("Segoe UI", 10, "bold"), bg=TEAL, fg=WHITE,
                                    activebackground=TEAL_DARK, activeforeground=WHITE,
                                    relief="flat", cursor="hand2", padx=20, pady=7,
                                    state="disabled", command=self._run_import)
        self.import_btn.pack(side="right")

        self.progress_var = tk.DoubleVar(value=0)
        self.progress_bar = ttk.Progressbar(f, variable=self.progress_var,
                                            maximum=100, mode="determinate",
                                            length=400)
        self.progress_bar.pack(padx=32, pady=(0, 4), fill="x")

        style = ttk.Style()
        style.configure("green.Horizontal.TProgressbar",
                        troughcolor=SURFACE, background=SUCCESS)
        self.progress_bar.configure(style="green.Horizontal.TProgressbar")

    # ── RESULTS TAB ──────────────────────────
    def _build_results_tab(self):
        f = self.tab_results

        tk.Label(f, text="Import Results", font=FONT_H1,
                 bg=BG, fg=TEXT).pack(anchor="w", padx=32, pady=12)

        summary_frame = tk.Frame(f, bg=BG)
        summary_frame.pack(fill="x", padx=32, pady=(0, 8))

        for col, (label, color) in enumerate([
            ("Total Rows", TEXT),
            ("Succeeded", SUCCESS),
            ("Failed", ERROR),
            ("Skipped", WARNING),
        ]):
            card = tk.Frame(summary_frame, bg=SURFACE, relief="flat",
                            highlightthickness=1, highlightbackground=BORDER)
            card.grid(row=0, column=col, padx=(0, 12), ipadx=16, ipady=10, sticky="ew")
            summary_frame.columnconfigure(col, weight=1)

            var = tk.StringVar(value="—")
            setattr(self, f"stat_{label.lower().replace(' ','_')}_var", var)

            tk.Label(card, textvariable=var, font=("Segoe UI", 22, "bold"),
                     bg=SURFACE, fg=color).pack()
            tk.Label(card, text=label, font=FONT_SMALL, bg=SURFACE, fg=TEXT_MUTED).pack()

        # Results tree
        tree_frame = tk.Frame(f, bg=BG)
        tree_frame.pack(fill="both", expand=True, padx=32, pady=4)

        cols = ("row", "description", "cost", "date", "status", "message")
        self.results_tree = ttk.Treeview(tree_frame, columns=cols, show="headings",
                                          selectmode="browse", height=14)
        for col, (heading, width) in zip(cols, [
            ("#", 50), ("Description", 200), ("Cost", 100),
            ("Date", 100), ("Status", 90), ("Message", 320)
        ]):
            self.results_tree.heading(col, text=heading)
            self.results_tree.column(col, width=width, anchor="w")

        vsb2 = ttk.Scrollbar(tree_frame, orient="vertical",
                              command=self.results_tree.yview)
        self.results_tree.configure(yscrollcommand=vsb2.set)
        vsb2.pack(side="right", fill="y")
        self.results_tree.pack(fill="both", expand=True)

        # Tag colors
        self.results_tree.tag_configure("success", foreground=SUCCESS)
        self.results_tree.tag_configure("error",   foreground=ERROR)
        self.results_tree.tag_configure("dry_run", foreground=WARNING)

        tk.Button(f, text="Export Results to CSV",
                  font=FONT_BODY, bg=SURFACE, fg=TEAL,
                  activebackground=BORDER, relief="flat", cursor="hand2",
                  padx=14, pady=6, command=self._export_results
                  ).pack(anchor="e", padx=32, pady=8)

    # ── LOG TAB ──────────────────────────────
    def _build_log_tab(self):
        f = self.tab_log
        tk.Label(f, text="Activity Log", font=FONT_H1,
                 bg=BG, fg=TEXT).pack(anchor="w", padx=32, pady=12)

        self.log_text = scrolledtext.ScrolledText(
            f, font=FONT_MONO, bg="#1C1B19", fg="#CDCCCA",
            relief="flat", wrap="word", state="disabled",
            insertbackground=WHITE, padx=12, pady=8)
        self.log_text.pack(fill="both", expand=True, padx=32, pady=(0, 8))
        self.log_text.tag_configure("info",    foreground="#CDCCCA")
        self.log_text.tag_configure("success", foreground="#6DAA45")
        self.log_text.tag_configure("warning", foreground="#BB653B")
        self.log_text.tag_configure("error",   foreground="#D163A7")
        self.log_text.tag_configure("dim",     foreground="#5A5957")

        tk.Button(f, text="Clear Log", font=FONT_BODY,
                  bg=SURFACE, fg=TEXT_MUTED, activebackground=BORDER,
                  relief="flat", cursor="hand2", padx=12, pady=5,
                  command=self._clear_log).pack(anchor="e", padx=32, pady=(0, 12))

    # ── LOGIC ────────────────────────────────
    def _log(self, message: str, level: str = "info"):
        ts = datetime.now().strftime("%H:%M:%S")
        prefix = {"info": "  ", "success": "✓ ", "warning": "⚠ ", "error": "✗ "}.get(level, "  ")
        line = f"[{ts}] {prefix}{message}\n"
        self.log_text.configure(state="normal")
        self.log_text.insert("end", line, level)
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _clear_log(self):
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

    def _set_status(self, msg: str, color: str = TEXT_MUTED):
        self.status_var.set(msg)
        self.status_label.configure(fg=color)

    def _toggle_key_visibility(self):
        current = self.api_key_entry.cget("show")
        if current == "•":
            self.api_key_entry.configure(show="")
            self.show_key_btn.configure(text="Hide")
        else:
            self.api_key_entry.configure(show="•")
            self.show_key_btn.configure(text="Show")

    def _connect(self):
        api_key = self.api_key_var.get().strip()
        if not api_key:
            messagebox.showwarning("No API Key", "Please enter your Splitwise API key.")
            return

        self.connect_btn.configure(state="disabled", text="Connecting…")
        self._log("Attempting to connect to Splitwise API…")
        self._set_status("Connecting to Splitwise…")

        def _do():
            try:
                client = SplitwiseClient(api_key)
                user = client.get_current_user()
                groups = client.get_groups()
                friends = client.get_friends()

                self.client = client
                self.groups = groups
                self.friends = friends

                name = f"{user.get('first_name','')} {user.get('last_name','')}".strip()
                email = user.get("email", "")

                self.after(0, lambda: self._on_connect_success(name, email, groups, friends))
            except Exception as e:
                self.after(0, lambda: self._on_connect_fail(str(e)))

        threading.Thread(target=_do, daemon=True).start()

    def _on_connect_success(self, name, email, groups, friends):
        self.connect_btn.configure(state="normal", text="Connect & Verify")

        group_names = ["No Group (0)"] + [f"{g['name']} ({g['id']})" for g in groups]
        self.group_combo["values"] = group_names
        if len(group_names) > 1:
            self.default_group_var.set(group_names[1])

        msg = f"🟢  Connected as {name} ({email})  •  {len(groups)} groups  •  {len(friends)} friends"
        self.auth_status_label.configure(text=msg, fg=SUCCESS)
        self._set_status(f"Connected as {name}", SUCCESS)
        self._log(f"Connected as {name} ({email})", "success")
        self._log(f"Found {len(groups)} groups, {len(friends)} friends", "info")
        for g in groups:
            self._log(f"  Group: {g['name']} (ID {g['id']})", "dim")

        save_settings({"api_key": self.api_key_var.get()})
        if self.csv_rows:
            self.import_btn.configure(state="normal")

    def _on_connect_fail(self, error_msg):
        self.connect_btn.configure(state="normal", text="Connect & Verify")
        self.auth_status_label.configure(
            text=f"🔴  Connection failed: {error_msg}", fg=ERROR)
        self._set_status("Connection failed.", ERROR)
        self._log(f"Connection failed: {error_msg}", "error")
        messagebox.showerror("Connection Failed",
                             f"Could not connect to Splitwise.\n\n{error_msg}\n\n"
                             "Check your API key and internet connection.")

    def _browse_csv(self):
        path = filedialog.askopenfilename(
            title="Select CSV File",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
        if not path:
            return
        self._load_csv(path)

    def _load_csv(self, path: str):
        try:
            rows = []
            headers = []
            with open(path, newline="", encoding="utf-8-sig") as f:
                reader = csv.DictReader(f)
                headers = list(reader.fieldnames or [])
                for row in reader:
                    rows.append(dict(row))

            if not headers:
                messagebox.showerror("Invalid CSV", "The CSV file has no headers.")
                return

            self.csv_rows = rows
            self.csv_headers = headers
            self.file_path_var.set(path)
            self.row_count_label.configure(
                text=f"({len(rows)} rows)", fg=TEXT_MUTED)

            self._build_column_mapping(headers)
            self._populate_preview(rows[:10], headers)

            self._log(f"Loaded: {os.path.basename(path)} — {len(rows)} rows, {len(headers)} columns", "success")
            self._set_status(f"Loaded {len(rows)} rows from {os.path.basename(path)}", SUCCESS)

            if self.client:
                self.import_btn.configure(state="normal")

        except Exception as e:
            messagebox.showerror("Load Error", f"Could not read CSV:\n{e}")
            self._log(f"CSV load error: {e}", "error")

    def _build_column_mapping(self, headers: list):
        # Clear old widgets
        for widget in self.mapping_frame.winfo_children():
            widget.destroy()

        self.column_vars = {}
        NONE_OPTION = "(skip)"

        fields = [
            ("date",        "Date",        True,  "Transaction date — YYYY-MM-DD or MM/DD/YYYY"),
            ("description", "Description", True,  "Expense name/memo (required by Splitwise)"),
            ("cost",        "Amount",      True,  "Numeric amount, e.g. 45.00 (required)"),
            ("currency",    "Currency",    False, "ISO 4217 code — falls back to default if blank"),
            ("category",    "Category",    False, "Splitwise category name"),
            ("notes",       "Notes",       False, "Extra notes appended to the expense"),
        ]

        header_norm = {h.lower().strip(): h for h in headers}
        options = [NONE_OPTION] + headers

        title_row = tk.Frame(self.mapping_frame, bg=BG)
        title_row.pack(fill="x", pady=(0, 6))
        tk.Label(title_row, text="Splitwise Field", font=FONT_SMALL,
                 bg=BG, fg=TEXT_MUTED, width=18, anchor="w").pack(side="left")
        tk.Label(title_row, text="CSV Column to Use", font=FONT_SMALL,
                 bg=BG, fg=TEXT_MUTED).pack(side="left", padx=(8, 0))

        for field, label, required, tip in fields:
            row = tk.Frame(self.mapping_frame, bg=BG)
            row.pack(fill="x", pady=3)

            badge = " *" if required else "  "
            lbl_text = f"{label}{badge}"
            tk.Label(row, text=lbl_text, font=FONT_BODY, bg=BG, fg=TEXT,
                     width=18, anchor="w").pack(side="left")

            var = tk.StringVar()
            combo = ttk.Combobox(row, textvariable=var, values=options,
                                 state="readonly", width=32, font=FONT_BODY)
            combo.pack(side="left", padx=(8, 0))
            ToolTip(combo, tip)

            # Auto-detect column
            auto = NONE_OPTION
            for alias in KNOWN_ALIASES.get(field, []):
                if alias in header_norm:
                    auto = header_norm[alias]
                    break
            var.set(auto)
            self.column_vars[field] = var

        tk.Label(self.mapping_frame, text="  * Required fields",
                 font=FONT_SMALL, bg=BG, fg=TEXT_MUTED).pack(anchor="w", pady=(6, 0))

    def _populate_preview(self, rows: list, headers: list):
        tree = self.preview_tree
        tree.delete(*tree.get_children())
        tree["columns"] = headers
        for h in headers:
            tree.heading(h, text=h)
            tree.column(h, width=max(80, len(h) * 9), anchor="w")
        for row in rows:
            tree.insert("", "end", values=[row.get(h, "") for h in headers])

    def _run_import(self):
        if not self.csv_rows:
            messagebox.showwarning("No Data", "Please load a CSV file first.")
            return
        if not self.client and not self.dry_run_var.get():
            messagebox.showwarning("Not Connected",
                                   "Please connect to Splitwise before importing.\n"
                                   "Or enable Dry Run to validate without connecting.")
            return

        # Validate required fields are mapped
        for field in ("description", "cost"):
            if self.column_vars.get(field, tk.StringVar()).get() in ("(skip)", ""):
                messagebox.showerror("Mapping Error",
                                     f"'{field}' is required but not mapped to a column.")
                return

        # Get default group ID
        group_str = self.default_group_var.get()
        try:
            default_group_id = int(group_str.split("(")[-1].rstrip(")"))
        except Exception:
            default_group_id = 0

        dry_run   = self.dry_run_var.get()
        split_eq  = self.split_equally_var.get()
        currency  = self.default_currency_var.get().strip().upper() or "USD"

        self.import_btn.configure(state="disabled", text="Importing…")
        self.progress_var.set(0)
        self._clear_results()

        mode_label = "DRY RUN" if dry_run else "LIVE IMPORT"
        self._log(f"──── Starting {mode_label}: {len(self.csv_rows)} rows ────", "info")

        def _worker():
            results = []
            total = len(self.csv_rows)

            for idx, row in enumerate(self.csv_rows):
                result = self._process_row(row, idx + 1, default_group_id,
                                           split_eq, currency, dry_run)
                results.append(result)
                pct = ((idx + 1) / total) * 100
                self.after(0, lambda p=pct: self.progress_var.set(p))
                if not dry_run:
                    time.sleep(0.15)  # avoid rate-limiting

            self.after(0, lambda: self._finish_import(results, dry_run))

        threading.Thread(target=_worker, daemon=True).start()

    def _process_row(self, row: dict, row_num: int, default_group_id: int,
                     split_eq: bool, currency: str, dry_run: bool) -> dict:
        def get(field):
            col = self.column_vars.get(field, tk.StringVar()).get()
            if col in ("(skip)", ""):
                return ""
            return str(row.get(col, "")).strip()

        description = get("description") or f"Expense #{row_num}"
        cost_raw    = get("cost").replace("$", "").replace(",", "").strip()
        date_str    = get("date")
        currency_v  = get("currency") or currency
        notes       = get("notes")

        # Validate cost
        try:
            cost_f = float(cost_raw)
            if cost_f <= 0:
                return self._result(row_num, description, cost_raw, date_str,
                                    "skipped", "Amount is zero or negative")
            cost_str = f"{cost_f:.2f}"
        except (ValueError, TypeError):
            return self._result(row_num, description, cost_raw, date_str,
                                "error", f"Invalid amount: '{cost_raw}'")

        # Normalise date
        date_iso = ""
        if date_str:
            for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%m-%d-%Y", "%d/%m/%Y",
                        "%Y/%m/%d", "%d-%m-%Y", "%m/%d/%y"):
                try:
                    date_iso = datetime.strptime(date_str, fmt).strftime("%Y-%m-%dT%H:%M:%SZ")
                    break
                except ValueError:
                    continue

        if dry_run:
            msg = f"[DRY RUN] Would post: {description} — {currency_v} {cost_str}"
            self._log(msg, "warning")
            return self._result(row_num, description, cost_str, date_str, "dry_run", "OK (not sent)")

        # Build payload
        payload = {
            "cost": cost_str,
            "description": description,
            "currency_code": currency_v,
            "group_id": default_group_id,
            "split_equally": "true" if split_eq else "false",
        }
        if date_iso:
            payload["date"] = date_iso
        if notes:
            payload["details"] = notes

        try:
            self.client.create_expense(payload)
            self._log(f"  ✓ Row {row_num}: {description} — {currency_v} {cost_str}", "success")
            return self._result(row_num, description, cost_str, date_str, "success", "Imported")
        except Exception as e:
            self._log(f"  ✗ Row {row_num}: {description} — {e}", "error")
            return self._result(row_num, description, cost_str, date_str, "error", str(e))

    @staticmethod
    def _result(row_num, description, cost, date, status, message):
        return {"row": row_num, "description": description, "cost": cost,
                "date": date, "status": status, "message": message}

    def _finish_import(self, results: list, dry_run: bool):
        self.import_results = results
        self.import_btn.configure(state="normal", text="▶  Run Import")
        self.progress_var.set(100)

        counts = {"success": 0, "error": 0, "skipped": 0, "dry_run": 0}
        for r in results:
            counts[r["status"]] = counts.get(r["status"], 0) + 1

        total   = len(results)
        success = counts.get("success", 0) + counts.get("dry_run", 0)
        errors  = counts.get("error", 0)
        skipped = counts.get("skipped", 0)

        self.stat_total_rows_var.set(str(total))
        self.stat_succeeded_var.set(str(success))
        self.stat_failed_var.set(str(errors))
        self.stat_skipped_var.set(str(skipped))

        for r in results:
            tag = r["status"] if r["status"] in ("success", "error", "dry_run") else "info"
            self.results_tree.insert("", "end", tags=(tag,),
                                     values=(r["row"], r["description"],
                                             r["cost"], r["date"],
                                             r["status"].upper(), r["message"]))

        mode = "Dry run" if dry_run else "Import"
        summary = f"{mode} complete — {success} OK, {errors} errors, {skipped} skipped"
        color = SUCCESS if errors == 0 else ERROR
        self._set_status(summary, color)
        self._log(f"──── {summary} ────", "success" if errors == 0 else "warning")

        self.nb.select(self.tab_results)

    def _clear_results(self):
        self.results_tree.delete(*self.results_tree.get_children())
        for attr in ("total_rows", "succeeded", "failed", "skipped"):
            getattr(self, f"stat_{attr}_var").set("—")

    def _export_results(self):
        if not self.import_results:
            messagebox.showinfo("No Results", "Run an import first.")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV", "*.csv")],
            title="Export Results")
        if not path:
            return
        try:
            with open(path, "w", newline="", encoding="utf-8") as f:
                writer = csv.DictWriter(f, fieldnames=["row", "description", "cost",
                                                         "date", "status", "message"])
                writer.writeheader()
                writer.writerows(self.import_results)
            messagebox.showinfo("Exported", f"Results saved to:\n{path}")
            self._log(f"Results exported: {path}", "success")
        except Exception as e:
            messagebox.showerror("Export Error", str(e))

    def _restore_settings(self):
        key = self.settings.get("api_key", "")
        if key:
            self.api_key_var.set(key)
            self._log("Loaded saved API key from settings.", "dim")

    def _on_close(self):
        save_settings({"api_key": self.api_key_var.get()})
        self.destroy()


# ─────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────
if __name__ == "__main__":
    app = SplitwiseImporterApp()
    app.mainloop()
