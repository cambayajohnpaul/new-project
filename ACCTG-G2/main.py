"""
Accounting App: Journal -> Ledger -> 12-Column Worksheet
Run with:  python main.py      (built-in modules only, nothing to install)
Data saves automatically to accounting_data.json beside this file.

The app is split into files so teammates can work on different parts:
    constants.py      shared constants and fmt()/balfmt()
    data.py           load/save data, account list
    style.py          look and feel, shared table builder
    journal_tab.py    Tab 1: Journal Entry
    ledger_tab.py     Tab 2: Ledger
    worksheet_tab.py  Tab 3: Worksheet / Balance Sheet
    menu_bar.py       File / Edit / View / Tools / Help menus, undo-redo, backups, About
    exports.py        Excel (.xlsx) and PDF writers used by the File > Export menu
"""
import tkinter as tk
from tkinter import ttk

from journal_tab import JournalMixin
from ledger_tab import LedgerMixin
from style import StyleMixin
from worksheet_tab import WorksheetMixin
from menu_bar import MenuMixin


class App(StyleMixin, JournalMixin, LedgerMixin, WorksheetMixin, MenuMixin):
    def __init__(self, root):
        self.root = root
        root.title("Accounting App")
        root.geometry("1280x740")
        root.minsize(1000, 600)
        self.entries, self.custom, self.ws_rows, self.ws_ledger_data = [], {}, [], []
        self.load()
        self.set_style()
        nb = ttk.Notebook(root)
        nb.pack(fill="both", expand=True, padx=8, pady=8)
        self.build_journal(nb)
        self.build_ledger(nb)
        self.build_worksheet(nb)
        self.build_menu(nb)
        self.init_history()
        self.refresh()

    # ---------- refresh everything ----------
    def refresh(self):
        names = self.all_account_names()
        self.cb_dr["values"] = names
        self.cb_cr["values"] = names

        # Update ledger dropdown
        ledger_names = ["All Accounts"] + names
        self.cb_ledger_acct["values"] = ledger_names

        # Keep current selection if it still exists
        current = self.v_ledger_acct.get()
        if current not in ledger_names:
            self.v_ledger_acct.set("All Accounts")

        self.refresh_journal()
        self.refresh_ledger()
        self.refresh_worksheet()
        self.track_state()


if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()