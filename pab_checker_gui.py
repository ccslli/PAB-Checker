#!/usr/bin/env python3
r"""
pab_checker_gui.py
==================
Window version of the PAB checker.

All the checks live in design_workbook_check.py (the terminal version); this
file only adds the window, so any fix to a check is made once, in that file,
and both versions get it.  Keep the two files in the same folder.

Usage (Windows, Excel installed, `pip install pywin32 xlwings`):

    python pab_checker_gui.py

Rename it to pab_checker_gui.pyw to start it without a console window.
"""

import os
import subprocess
import sys
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import design_workbook_check as core

APP_TITLE = "PAB Checker"
PLACEHOLDER = "Select a workbook..."

# colours: errors / warnings / passes / notes
RED, AMBER, GREEN, GREY = "#b3261e", "#8a5300", "#1e7b34", "#5f6368"


class CheckerApp:
    def __init__(self, root):
        self.root = root
        self.entries = []            # [(label, opener)] from core.list_open_workbooks()
        self.wb = None               # Excel workbook object last checked
        self.wb_label = None
        self.report = None
        self.checked_name = ""       # workbook name of the last check, for Copy results
        self.marks = {}              # tab name -> text mark, for jumping to a section

        root.title(APP_TITLE)
        root.geometry("1100x720")
        root.minsize(760, 460)
        self._fonts()
        self._build()
        root.bind("<F5>", lambda e: self.run_checks())
        root.after(100, self.refresh_workbooks)

    # ------------------------------------------------------------------ layout
    def _fonts(self):
        base = "Segoe UI" if sys.platform == "win32" else "TkDefaultFont"
        self.f_body = (base, 10)
        self.f_bold = (base, 10, "bold")
        self.f_head = (base, 12, "bold")
        self.f_summary = (base, 13, "bold")
        style = ttk.Style(self.root)
        if "vista" in style.theme_names():
            style.theme_use("vista")
        style.configure("Accent.TButton", font=self.f_bold, padding=(14, 4))
        style.configure("TButton", padding=(10, 4))
        style.configure("Treeview", rowheight=24, font=self.f_body)

    def _build(self):
        top = ttk.Frame(self.root, padding=(12, 10, 12, 4))
        top.pack(fill="x")
        ttk.Label(top, text="Workbook", font=self.f_bold).pack(side="left")
        self.combo = ttk.Combobox(top, state="readonly", font=self.f_body)
        self.combo.pack(side="left", fill="x", expand=True, padx=(8, 6))
        self.combo.set(PLACEHOLDER)
        ttk.Button(top, text="Refresh list", command=self.refresh_workbooks).pack(side="left", padx=3)
        ttk.Button(top, text="Open file...", command=self.open_file).pack(side="left", padx=3)

        bar = ttk.Frame(self.root, padding=(12, 4, 12, 8))
        bar.pack(fill="x")
        self.btn_run = ttk.Button(bar, text="Run checks  (F5)", style="Accent.TButton",
                                  command=self.run_checks)
        self.btn_run.pack(side="left")
        self.btn_qa = ttk.Button(bar, text="Run DOE QA script", command=self.run_qa)
        self.btn_qa.pack(side="left", padx=8)
        self.btn_copy = ttk.Button(bar, text="Copy results", command=self.copy_results, state="disabled")
        self.btn_copy.pack(side="right")
        self.status = ttk.Label(bar, text="", foreground=GREY, font=self.f_body)
        self.status.pack(side="left", padx=12)

        self.summary = tk.Label(self.root, text="Choose a workbook and press Run checks.",
                                font=self.f_summary, anchor="w", padx=14, pady=6, fg=GREY)
        self.summary.pack(fill="x")

        body = ttk.PanedWindow(self.root, orient="horizontal")
        body.pack(fill="both", expand=True, padx=12, pady=(0, 12))

        left = ttk.Frame(body)
        self.tree = ttk.Treeview(left, columns=("result",), show="tree headings", selectmode="browse")
        self.tree.heading("#0", text="Tab")
        self.tree.heading("result", text="Result")
        self.tree.column("#0", width=200, stretch=True)
        self.tree.column("result", width=110, anchor="w", stretch=False)
        self.tree.tag_configure("error", foreground=RED)
        self.tree.tag_configure("warn", foreground=AMBER)
        self.tree.tag_configure("ok", foreground=GREEN)
        self.tree.pack(fill="both", expand=True)
        self.tree.bind("<<TreeviewSelect>>", self._jump)
        body.add(left, weight=1)

        right = ttk.Frame(body)
        self.text = tk.Text(right, wrap="word", font=self.f_body, padx=14, pady=10,
                            relief="flat", borderwidth=1, highlightthickness=1,
                            highlightbackground="#d0d0d0", spacing1=2, spacing3=2, cursor="arrow")
        scroll = ttk.Scrollbar(right, orient="vertical", command=self.text.yview)
        self.text.configure(yscrollcommand=scroll.set, state="disabled")
        scroll.pack(side="right", fill="y")
        self.text.pack(side="left", fill="both", expand=True)
        body.add(right, weight=4)

        t = self.text
        t.tag_configure("head", font=self.f_head, spacing1=14, spacing3=4)
        t.tag_configure("ERROR", foreground=RED, lmargin1=8, lmargin2=26)
        t.tag_configure("WARN", foreground=AMBER, lmargin1=8, lmargin2=26)
        t.tag_configure("NOTE", foreground=GREY, lmargin1=8, lmargin2=26)
        t.tag_configure("OK", foreground=GREEN, lmargin1=8, lmargin2=26)

    # --------------------------------------------------------------- workbooks
    def refresh_workbooks(self, select=None):
        self._busy("Looking for open workbooks...")
        try:
            self.entries = core.list_open_workbooks()
        except Exception as exc:
            self.entries = []
            self._error("Could not read the open Excel workbooks.", exc)
        finally:
            self._idle()
        labels = [label for label, _ in self.entries]
        self.combo["values"] = labels
        keep = select or self.combo.get()
        if keep in labels:
            self.combo.set(keep)
        else:
            self.combo.set(PLACEHOLDER)
        self.status.configure(text="%d open workbook(s) found." % len(labels) if labels
                              else "No open workbooks - open one in Excel or use Open file...")

    def open_file(self):
        path = filedialog.askopenfilename(parent=self.root, title="Select an Excel workbook",
                                          filetypes=[("Excel workbooks", "*.xlsx *.xlsm")])
        if not path:
            return
        self._busy("Opening %s..." % os.path.basename(path))
        try:
            wb = core.open_workbook_path(path)
        except Exception as exc:
            self._idle()
            self._error("Could not open %s." % os.path.basename(path), exc)
            return
        self._idle()
        self.refresh_workbooks(select=str(wb.Name))
        self.wb, self.wb_label = wb, str(wb.Name)
        self.combo.set(self.wb_label)
        self.run_checks()

    def _selected_workbook(self):
        """Workbook object for the combobox choice; reuses the last one on a re-run."""
        label = self.combo.get()
        if label == PLACEHOLDER or not label:
            messagebox.showinfo(APP_TITLE, "Choose a workbook first.", parent=self.root)
            return None
        if self.wb is not None and label == self.wb_label:
            return self.wb
        opener = dict(self.entries).get(label)
        if opener is None:
            self.refresh_workbooks()
            opener = dict(self.entries).get(label)
        if opener is None:
            messagebox.showwarning(APP_TITLE, "%s is no longer open in Excel." % label, parent=self.root)
            return None
        wb = opener(confirm=lambda q: messagebox.askyesno(APP_TITLE, q, parent=self.root))
        if wb is None:
            return None
        self.wb, self.wb_label = wb, label
        return wb

    # ------------------------------------------------------------------ checks
    def run_checks(self):
        wb = self._selected_workbook()
        if wb is None:
            return
        try:
            name = str(wb.Name)
        except Exception:                       # workbook was closed since the last run
            self.wb = None
            self.refresh_workbooks()
            wb = self._selected_workbook()
            if wb is None:
                return
            name = str(wb.Name)
        self._busy("Reading %s..." % name)
        try:
            rep = core.check_workbook_report(wb)
        except Exception as exc:
            self._idle()
            self.wb = None
            self._error("The checks could not finish.", exc)
            return
        self._idle()
        self.report = rep
        self.checked_name = name
        self._show(rep, name)
        self.status.configure(text="Checked %s" % name)

    def _show(self, rep, name):
        n_err, n_warn = rep.count("ERROR"), rep.count("WARN")
        if n_err or n_warn:
            parts = []
            if n_err:
                parts.append("%d error%s" % (n_err, "" if n_err == 1 else "s"))
            if n_warn:
                parts.append("%d warning%s" % (n_warn, "" if n_warn == 1 else "s"))
            self.summary.configure(text="   |   ".join(parts), fg=RED if n_err else AMBER)
        else:
            self.summary.configure(text="No issues found", fg=GREEN)

        self.tree.delete(*self.tree.get_children())
        t = self.text
        t.configure(state="normal")
        t.delete("1.0", "end")
        self.marks = {}
        for tab, items in rep.sections(core.tab_order()):
            e = sum(1 for lv, _ in items if lv == "ERROR")
            w = sum(1 for lv, _ in items if lv == "WARN")
            if e:
                result, tag = "%d error%s" % (e, "" if e == 1 else "s"), "error"
            elif w:
                result, tag = "%d warning%s" % (w, "" if w == 1 else "s"), "warn"
            else:
                result, tag = "OK", "ok"
            self.tree.insert("", "end", iid=tab, text=tab, values=(result,), tags=(tag,))

            mark = "sec%d" % len(self.marks)
            t.mark_set(mark, "end-1c")
            t.mark_gravity(mark, "left")
            self.marks[tab] = mark
            t.insert("end", tab + "\n", "head")
            for lv, msg in items:
                t.insert("end", "-  " + msg + "\n", lv)
            if not e and not w:
                t.insert("end", "-  No issues found in %s.\n" % tab, "OK")
        t.configure(state="disabled")
        t.yview_moveto(0)
        self.btn_copy.configure(state="normal")

    def _jump(self, _event=None):
        sel = self.tree.selection()
        if sel and sel[0] in self.marks:
            self.text.yview(self.marks[sel[0]])

    def copy_results(self):
        if not self.report:
            return
        text = "Workbook: %s\n%s" % (self.checked_name, self.report.render(core.tab_order()))
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self.status.configure(text="Results copied to the clipboard.")

    # ------------------------------------------------------------ DOE QA script
    def run_qa(self):
        if self.report is not None and self.report.count("ERROR"):
            if not messagebox.askyesno(
                    APP_TITLE, "The last check found %d error(s).\n\nRun the DOE QA script anyway?"
                    % self.report.count("ERROR"), parent=self.root):
                return
        script = self._find_qa_script()
        if not script or not self._qa_folder_ready(os.path.dirname(script)):
            return
        python = sys.executable
        if os.path.basename(python).lower() == "pythonw.exe":      # QA script needs a console
            python = os.path.join(os.path.dirname(python), "python.exe")
        flags = getattr(subprocess, "CREATE_NEW_CONSOLE", 0)
        try:
            subprocess.Popen([python, script], cwd=os.path.dirname(script), creationflags=flags)
        except Exception as exc:
            self._error("Could not start the DOE QA script.", exc)
            return
        self.root.destroy()                     # the QA script carries on in its own window

    def _aborted(self):
        return not messagebox.askretrycancel(APP_TITLE, "DOE QA script aborted.\n\nTry again?",
                                             parent=self.root)

    def _find_qa_script(self):
        folder, name = core.load_qa_settings()
        if folder and not os.path.isdir(folder):
            messagebox.showinfo(APP_TITLE, "The saved DOE QA script folder is no longer accessible:\n\n%s\n\n"
                                "Please select it again." % folder, parent=self.root)
            folder = ""
        while True:
            if not folder:
                folder = filedialog.askdirectory(parent=self.root,
                                                 title="Select the folder with the DOE QA script")
                folder = os.path.normpath(folder) if folder else ""
            if folder:
                path = os.path.join(folder, name)
                if not os.path.isfile(path):
                    path = ""
                    if messagebox.askyesno(APP_TITLE, "%s was not found in:\n\n%s\n\nDo you want to locate "
                                           "the Python script in this folder?" % (name, folder),
                                           parent=self.root):
                        path = filedialog.askopenfilename(parent=self.root, initialdir=folder,
                                                          title="Select the DOE QA Python script",
                                                          filetypes=[("Python scripts", "*.py")])
                        path = os.path.normpath(path) if path and os.path.isfile(path) else ""
                if path:
                    core.save_qa_settings(os.path.dirname(path), os.path.basename(path))
                    return path
                folder = ""
            if self._aborted():
                return None

    def _qa_folder_ready(self, folder):
        while True:
            try:
                files = core.qa_xlsx_files(folder)
            except OSError as exc:
                files, problem = None, "Could not read the folder:\n%s\n\n%s" % (folder, exc)
            if files is not None and len(files) == 1:
                return True
            if files is not None:
                if not files:
                    problem = "No Excel file (.xlsx) was found in:\n%s" % folder
                else:
                    problem = "%d Excel files (.xlsx) were found in:\n%s\n\n%s" % (
                        len(files), folder, "\n".join("    " + f for f in files))
                problem += "\n\nThere must be exactly 1 .xlsx file in this folder for the DOE QA script."
            if not messagebox.askretrycancel(APP_TITLE, problem + "\n\nFix the folder, then press Retry "
                                             "to check again.", parent=self.root):
                messagebox.showinfo(APP_TITLE, "DOE QA script aborted.", parent=self.root)
                return False

    # ----------------------------------------------------------------- helpers
    def _busy(self, msg):
        self.status.configure(text=msg)
        self.root.configure(cursor="watch")
        for b in (self.btn_run, self.btn_qa):
            b.configure(state="disabled")
        self.root.update()

    def _idle(self):
        self.root.configure(cursor="")
        for b in (self.btn_run, self.btn_qa):
            b.configure(state="normal")

    def _error(self, msg, exc):
        traceback.print_exc()
        messagebox.showerror(APP_TITLE, "%s\n\n%s: %s" % (msg, type(exc).__name__, exc), parent=self.root)


def main():
    if core.pythoncom is None or core.xw is None:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(APP_TITLE, "pywin32 and xlwings are needed.\n\nRun:  pip install pywin32 xlwings")
        return 1
    root = tk.Tk()
    CheckerApp(root)
    root.mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())
