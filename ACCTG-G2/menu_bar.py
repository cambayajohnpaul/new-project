"""Menu bar: File / Edit / View / Tools / Help, undo-redo, backups and the About dialog."""
import os
import shutil
import tkinter as tk
from datetime import datetime
from decimal import Decimal, InvalidOperation
from tkinter import filedialog, messagebox, ttk

from constants import DATA_FILE, HEADS, PRIMARY_COLOR, TAB_BG, fmt
from data import DataMixin
from exports import write_pdf, write_xlsx

APP_NAME = "Accounting App"
APP_VERSION = "1.0"
BACKUP_LIMIT = 5      # newest backups kept
HISTORY_LIMIT = 50    # undo steps kept
BACKUP_DIR = os.path.join(os.path.dirname(DATA_FILE), "backups")
BACKUP_PREFIX = "accounting_data-"


class MenuMixin(DataMixin):
    # ================= menu bar =================
    def build_menu(self, nb):
        self.nb = nb
        bar = tk.Menu(self.root)
        self.root.config(menu=bar)

        # File
        file_menu = tk.Menu(bar, tearoff=0)
        export_menu = tk.Menu(file_menu, tearoff=0)
        export_menu.add_command(label="Export CSV", command=self.export_worksheet_csv)
        export_menu.add_command(label="Export Excel", command=self.export_worksheet_excel)
        export_menu.add_command(label="Export PDF", command=self.export_worksheet_pdf)
        export_menu.add_separator()
        export_menu.add_command(label="Export Ledger CSV", command=self.export_ledger_csv)
        file_menu.add_cascade(label="Export", menu=export_menu)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.destroy)
        bar.add_cascade(label="File", menu=file_menu)

        # Edit
        edit_menu = tk.Menu(bar, tearoff=0)
        edit_menu.add_command(label="New Entry", accelerator="Ctrl+N", command=self.new_entry)
        edit_menu.add_separator()
        edit_menu.add_command(label="Undo", accelerator="Ctrl+Z", command=self.undo)
        self._undo_index = edit_menu.index("end")
        edit_menu.add_command(label="Redo", accelerator="Ctrl+Y", command=self.redo)
        self._redo_index = edit_menu.index("end")
        edit_menu.add_separator()
        edit_menu.add_command(label="Delete Entry", accelerator="Del", command=self.delete_selected_entry)
        edit_menu.add_separator()
        edit_menu.add_command(label="Preferences...", command=self.open_preferences)
        bar.add_cascade(label="Edit", menu=edit_menu)
        self.edit_menu = edit_menu

        # View
        view_menu = tk.Menu(bar, tearoff=0)
        view_menu.add_command(label="Refresh", accelerator="F5", command=self.refresh)
        bar.add_cascade(label="View", menu=view_menu)

        # Tools
        tools_menu = tk.Menu(bar, tearoff=0)
        tools_menu.add_command(label="Recalculate Worksheet", command=self.recalculate_worksheet)
        tools_menu.add_command(label="Validate Data", command=self.validate_data)
        tools_menu.add_command(label="Backup Manager...", command=self.backup_manager)
        bar.add_cascade(label="Tools", menu=tools_menu)

        # Help
        help_menu = tk.Menu(bar, tearoff=0)
        help_menu.add_command(label="Keyboard Shortcuts", command=self.show_shortcuts)
        help_menu.add_command(label="Documentation", command=self.show_documentation)
        help_menu.add_separator()
        help_menu.add_command(label="About", command=self.show_about)
        bar.add_cascade(label="Help", menu=help_menu)

        self.root.bind_all("<Control-n>", lambda _: self._shortcut(self.new_entry))
        self.root.bind_all("<Control-z>", lambda _: self._shortcut(self.undo, skip_text_fields=True))
        self.root.bind_all("<Control-y>", lambda _: self._shortcut(self.redo, skip_text_fields=True))
        self.root.bind_all("<Delete>", lambda _: self._shortcut(self.delete_selected_entry, skip_text_fields=True))
        self.root.bind_all("<F5>", lambda _: self._shortcut(self.refresh))

    def _shortcut(self, action, skip_text_fields=False):
        """Run a keyboard shortcut only from the main window (and not while typing, if asked)."""
        try:
            widget = self.root.focus_get()
        except KeyError:  # combobox pop-up lists can confuse focus_get
            widget = None
        if widget is not None and widget.winfo_toplevel() is not self.root:
            return
        if skip_text_fields and isinstance(widget, (tk.Entry, ttk.Entry, ttk.Combobox, tk.Text)):
            return
        action()

    # ================= File > Export =================
    def _worksheet_table(self):
        self.refresh_worksheet()
        headers = ["Account Name"] + HEADS
        rows = [([name] + list(cells), tag in ("total", "grand")) for name, cells, tag in self.ws_rows]
        return headers, rows

    def export_worksheet_excel(self):
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile="worksheet.xlsx",
                                            filetypes=[("Excel workbook", "*.xlsx")])
        if not path:
            return
        headers, rows = self._worksheet_table()
        try:
            write_xlsx(path, "Worksheet", headers, rows)
        except OSError as err:
            return messagebox.showerror("Export failed", str(err))
        messagebox.showinfo("Export", f"Worksheet exported to {os.path.basename(path)}")

    def export_worksheet_pdf(self):
        path = filedialog.asksaveasfilename(defaultextension=".pdf", initialfile="worksheet.pdf",
                                            filetypes=[("PDF file", "*.pdf")])
        if not path:
            return
        headers, rows = self._worksheet_table()
        text_rows = [([values[0]] + [fmt(v) for v in values[1:]], bold) for values, bold in rows]
        try:
            write_pdf(path, "Accounting Worksheet", headers, text_rows)
        except OSError as err:
            return messagebox.showerror("Export failed", str(err))
        messagebox.showinfo("Export", f"Worksheet exported to {os.path.basename(path)}")

    # ================= Edit =================
    def new_entry(self):
        self.nb.select(0)
        self.cb_dr.focus_set()

    def delete_selected_entry(self):
        tab = self.nb.index("current")
        if tab == 0:
            self.delete_journal_entry()
        elif tab == 1:
            self.delete_ledger_entry()
        else:
            messagebox.showinfo("Delete", "Open the Journal or Ledger tab and select an entry first.")

    def open_preferences(self):
        dlg = self._open_dialog("Preferences")
        body = ttk.Frame(dlg, padding=18)
        body.pack(fill="both", expand=True)
        ttk.Label(body, text="Table row height", font=("Segoe UI", 10, "bold"),
                  foreground=PRIMARY_COLOR).pack(anchor="w", pady=(0, 6))
        var = tk.IntVar(value=getattr(self, "_row_height", 28))

        def apply():
            self._row_height = var.get()
            ttk.Style().configure("Treeview", rowheight=self._row_height)

        for label, height in (("Compact", 22), ("Normal", 28), ("Comfortable", 34)):
            ttk.Radiobutton(body, text=label, value=height, variable=var, command=apply).pack(anchor="w", pady=1)
        ttk.Label(body, text="Applies until the app is closed.").pack(anchor="w", pady=(8, 0))
        ttk.Button(body, text="Close", style="Accent.TButton", command=dlg.destroy).pack(anchor="e", pady=(12, 0))
        self._show_dialog(dlg)

    # ---------- undo / redo ----------
    def init_history(self):
        self._undo, self._redo = [], []
        self._last_state = self._snapshot()
        self._update_edit_menu()

    def _snapshot(self):
        return [dict(e) for e in self.entries], dict(self.custom)

    def track_state(self):
        """Called after every refresh: remember the previous state whenever the data changed."""
        state = self._snapshot()
        if state != self._last_state:
            self._undo.append(self._last_state)
            del self._undo[:-HISTORY_LIMIT]
            self._redo.clear()
            self._last_state = state
        self._update_edit_menu()

    def _update_edit_menu(self):
        self.edit_menu.entryconfig(self._undo_index, state="normal" if self._undo else "disabled")
        self.edit_menu.entryconfig(self._redo_index, state="normal" if self._redo else "disabled")

    def _apply_state(self, state):
        """Switch the ledger to a saved state. Returns False (and changes nothing) if saving fails."""
        old_entries, old_custom = self.entries, self.custom
        self.entries, self.custom = [dict(e) for e in state[0]], dict(state[1])
        if not self.save():
            self.entries, self.custom = old_entries, old_custom
            return False
        self._last_state = self._snapshot()
        self.refresh()
        return True

    def undo(self):
        if not self._undo:
            return
        previous = self._last_state
        if self._apply_state(self._undo[-1]):
            self._undo.pop()
            self._redo.append(previous)
            self._update_edit_menu()

    def redo(self):
        if not self._redo:
            return
        previous = self._last_state
        if self._apply_state(self._redo[-1]):
            self._redo.pop()
            self._undo.append(previous)
            self._update_edit_menu()

    # ================= Tools =================
    def recalculate_worksheet(self):
        self.refresh()
        messagebox.showinfo("Recalculate Worksheet", "Worksheet recalculated.\n\n" + self.ws_status.get())

    def validate_data(self):
        names = set(self.all_account_names())
        problems = []
        for i, e in enumerate(self.entries, 1):
            try:
                amount = Decimal(e["amount"])
                if not amount.is_finite() or amount <= 0:
                    problems.append(f"Entry {i}: the amount must be greater than 0.")
            except (InvalidOperation, KeyError, TypeError):
                problems.append(f"Entry {i}: the amount is not a valid number.")
            for side in ("debit", "credit"):
                if e.get(side) not in names:
                    problems.append(f"Entry {i}: the {side} account '{e.get(side)}' does not exist.")
            if e.get("debit") == e.get("credit"):
                problems.append(f"Entry {i}: the debit and credit accounts are the same.")
        if not problems:
            _, ok = self.compute_worksheet()
            unbalanced = [name for name, balanced in ok.items() if not balanced]
            if unbalanced:
                problems.append("Worksheet columns out of balance: " + ", ".join(unbalanced))
        if not problems:
            return messagebox.showinfo(
                "Validate Data",
                f"No problems found.\n\nChecked {len(self.entries)} entries; all worksheet columns balance.")
        shown = "\n".join(problems[:12])
        extra = f"\n...and {len(problems) - 12} more." if len(problems) > 12 else ""
        messagebox.showwarning("Validate Data", f"{len(problems)} problem(s) found:\n\n{shown}{extra}")

    # ---------- backups ----------
    def save(self):
        """Back up the previous data file, then save as usual."""
        self._backup_current_file()
        return super().save()

    def _list_backups(self):
        """Backup file paths, newest first."""
        try:
            names = [n for n in os.listdir(BACKUP_DIR) if n.startswith(BACKUP_PREFIX) and n.endswith(".json")]
        except OSError:
            return []
        return [os.path.join(BACKUP_DIR, n) for n in sorted(names, reverse=True)]

    @staticmethod
    def _same_content(first, second):
        try:
            with open(first, "rb") as a, open(second, "rb") as b:
                return a.read() == b.read()
        except OSError:
            return False

    def _backup_current_file(self, force=False):
        """Copy the data file into the backups folder. Returns True if a backup was made."""
        if getattr(self, "_data_load_failed", False) or not os.path.isfile(DATA_FILE):
            return False
        try:
            os.makedirs(BACKUP_DIR, exist_ok=True)
            existing = self._list_backups()
            if not force and existing and self._same_content(DATA_FILE, existing[0]):
                return False
            stamp = datetime.now().strftime("%Y%m%d-%H%M%S-%f")
            shutil.copy2(DATA_FILE, os.path.join(BACKUP_DIR, f"{BACKUP_PREFIX}{stamp}.json"))
            for old in self._list_backups()[BACKUP_LIMIT:]:
                os.remove(old)
            return True
        except OSError:
            return False  # a failed backup must never block saving

    def backup_manager(self):
        dlg = self._open_dialog("Backup Manager")
        body = ttk.Frame(dlg, padding=16)
        body.pack(fill="both", expand=True)
        ttk.Label(body, justify="left",
                  text=f"A backup is made automatically before each save.\n"
                       f"The {BACKUP_LIMIT} newest are kept in:\n{BACKUP_DIR}").pack(anchor="w", pady=(0, 8))
        listbox = tk.Listbox(body, height=7, width=54, activestyle="none", exportselection=False,
                             font=("Segoe UI", 10))
        listbox.pack(fill="x")
        paths = []

        def reload():
            paths[:] = self._list_backups()
            listbox.delete(0, "end")
            for p in paths:
                when = datetime.fromtimestamp(os.path.getmtime(p)).strftime("%b %d, %Y  %I:%M:%S %p")
                listbox.insert("end", f"{when}    ({os.path.getsize(p) / 1024:.1f} KB)")
            if not paths:
                listbox.insert("end", "No backups yet.")

        def back_up_now():
            if not os.path.isfile(DATA_FILE):
                return messagebox.showinfo("Backup", "There is no data file to back up yet.", parent=dlg)
            if not self._backup_current_file(force=True):
                messagebox.showerror("Backup", "The backup could not be created.", parent=dlg)
            reload()

        def restore():
            selection = listbox.curselection()
            if not paths or not selection:
                return messagebox.showinfo("Restore", "Select a backup first.", parent=dlg)
            if not messagebox.askyesno(
                    "Restore backup",
                    "Replace the current data with the selected backup?\n"
                    "Your current data is backed up first.", parent=dlg):
                return
            try:
                with open(paths[selection[0]], "rb") as f:
                    content = f.read()  # read first: backing up below may rotate this file out
                self._backup_current_file()
                temp_path = DATA_FILE + ".restore.tmp"
                with open(temp_path, "wb") as f:
                    f.write(content)
                os.replace(temp_path, DATA_FILE)
            except OSError as err:
                return messagebox.showerror("Restore failed", str(err), parent=dlg)
            self.load()
            self.init_history()
            self.refresh()
            reload()
            messagebox.showinfo("Restore", "Backup restored.", parent=dlg)

        buttons = ttk.Frame(body)
        buttons.pack(fill="x", pady=(10, 0))
        ttk.Button(buttons, text="Back Up Now", command=back_up_now).pack(side="left")
        ttk.Button(buttons, text="Restore Selected", command=restore).pack(side="left", padx=8)
        ttk.Button(buttons, text="Close", style="Accent.TButton", command=dlg.destroy).pack(side="right")
        reload()
        self._show_dialog(dlg)

    # ================= Help =================
    def show_shortcuts(self):
        messagebox.showinfo(
            "Keyboard Shortcuts",
            "Ctrl+N    New entry (jumps to the Journal form)\n"
            "Ctrl+Z    Undo\n"
            "Ctrl+Y    Redo\n"
            "Del       Delete the selected entry (Journal or Ledger tab)\n"
            "F5        Refresh all tabs\n"
            "Enter     Add the entry while typing in the Journal form")

    def show_documentation(self):
        dlg = self._open_dialog("Documentation")
        body = ttk.Frame(dlg, padding=14)
        body.pack(fill="both", expand=True)
        text = tk.Text(body, width=74, height=22, wrap="word", font=("Segoe UI", 10), relief="flat",
                       background="white", padx=10, pady=8)
        text.insert("1.0", (
            "GETTING STARTED\n"
            "1. Journal: choose (or type) the debit and credit accounts, enter the amount and press "
            "Add Entry. Tick 'Adjusting entry' for adjustments. A new account name asks for its type.\n"
            "2. Ledger: pick an account, or 'All Accounts', to see its postings and running balance.\n"
            "3. Worksheet: the 12-column worksheet (Trial Balance, Adjustments, Adjusted Trial Balance, "
            "Income Statement, Owner's Equity, Balance Sheet). Tools > Recalculate Worksheet shows which "
            "column pairs balance.\n\n"
            "MENUS\n"
            "File > Export: save the worksheet as CSV, Excel (.xlsx) or PDF, or the ledger as CSV.\n"
            "Edit: new entry, undo / redo, delete the selected entry, preferences.\n"
            "View > Refresh: reload every tab.\n"
            "Tools: recalculate the worksheet, validate your data, manage backups.\n\n"
            f"DATA\nEverything is saved automatically to {os.path.basename(DATA_FILE)}. A backup of the previous "
            f"file is made before each save (the {BACKUP_LIMIT} newest are kept). Restore one from "
            "Tools > Backup Manager."))
        text.configure(state="disabled")
        text.pack(fill="both", expand=True)
        ttk.Button(body, text="Close", style="Accent.TButton", command=dlg.destroy).pack(anchor="e", pady=(10, 0))
        self._show_dialog(dlg)

    def show_about(self):
        dlg = self._open_dialog(f"About {APP_NAME}")
        body = ttk.Frame(dlg, padding=(26, 20))
        body.pack(fill="both", expand=True)
        heading = ("Segoe UI", 10, "bold")
        ttk.Label(body, text=APP_NAME, foreground=PRIMARY_COLOR, font=("Segoe UI", 18, "bold")).pack()
        ttk.Label(body, text="Journal  •  Ledger  •  Balance Sheet", foreground=PRIMARY_COLOR,
                  font=heading).pack(pady=(2, 10))
        ttk.Separator(body).pack(fill="x", pady=4)
        ttk.Label(body, text="Developers", foreground=PRIMARY_COLOR, font=heading).pack(anchor="w", pady=(6, 2))
        ttk.Label(body, text="BSIT 2nd Year — 1st Semester — Group 3").pack(anchor="w")
        ttk.Separator(body).pack(fill="x", pady=8)
        ttk.Label(body, text="Features", foreground=PRIMARY_COLOR, font=heading).pack(anchor="w", pady=(0, 2))
        features = (
            "• Zero-setup: runs with just Python (no pip install required)\n"
            "• Complete accounting cycle: Journal → Ledger → 12-Column Worksheet\n"
            f"• Automatic backup before every save (max {BACKUP_LIMIT} kept)\n"
            "• Data validation from the Tools menu\n"
            "• Undo / Redo for journal changes\n"
            "• Export to CSV, Excel (.xlsx), and PDF")
        ttk.Label(body, text=features, justify="left").pack(anchor="w")
        ttk.Separator(body).pack(fill="x", pady=8)
        ttk.Label(body, text=f"Version {APP_VERSION}  |  Data saved to {os.path.basename(DATA_FILE)}",
                  font=("Segoe UI", 8)).pack(anchor="w")
        ttk.Button(body, text="Close", style="Accent.TButton", command=dlg.destroy).pack(anchor="e", pady=(12, 0))
        self._show_dialog(dlg)

    # ================= dialog helpers =================
    def _open_dialog(self, title):
        dlg = tk.Toplevel(self.root)
        dlg.title(title)
        dlg.transient(self.root)
        dlg.resizable(False, False)
        dlg.configure(background=TAB_BG)
        return dlg

    def _show_dialog(self, dlg):
        dlg.update_idletasks()
        x = self.root.winfo_rootx() + (self.root.winfo_width() - dlg.winfo_reqwidth()) // 2
        y = self.root.winfo_rooty() + (self.root.winfo_height() - dlg.winfo_reqheight()) // 3
        dlg.geometry(f"+{max(x, 0)}+{max(y, 0)}")
        dlg.bind("<Escape>", lambda _: dlg.destroy())
        dlg.wait_visibility()
        dlg.grab_set()
        dlg.focus_set()