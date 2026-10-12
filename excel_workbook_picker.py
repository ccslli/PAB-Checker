"""
excel_workbook_picker.py
========================
Drop-in select_or_open_workbook() for scripts that work on an open Excel
workbook through xlwings / win32com.

Lists every workbook open in every running Excel instance - including files
in Protected View (the yellow 'Enable Editing' bar) - lets the user pick one,
or opens a new file with Excel's file dialog.  Returns the Excel Workbook COM
object (same as xlwings' book.api), or None if nothing was chosen.

Requires:  pip install pywin32 xlwings

Use it either by pasting these functions into a script, or by keeping this
file next to the script and importing it:

    from excel_workbook_picker import select_or_open_workbook
    wb = select_or_open_workbook()
    if wb is not None:
        ws = wb.Worksheets("Notes")
"""

import pythoncom
import xlwings as xw


def _excel_com():
    try:
        import win32com.client
        return win32com.client.GetActiveObject("Excel.Application")
    except Exception:
        return None


def _open_protected(pv):
    """Workbook behind a Protected View window: read it as-is if Excel allows,
    otherwise offer to click 'Enable Editing' for the user."""
    try:
        wb = pv.Workbook
        wb.Worksheets.Count                     # make sure it can be read
        return wb
    except Exception:
        pass
    try:
        answer = input("%s is in Protected View. Enable editing so it can be checked? "
                       "(y = yes, Enter = cancel): " % pv.SourceName).strip().lower()
    except EOFError:
        answer = ""
    if answer == "y":
        try:
            return pv.Edit()
        except Exception as exc:
            print("Could not enable editing: %s" % exc)
    return None


def select_or_open_workbook():
    pythoncom.CoInitialize()

    # (label shown to the user, function returning the Excel workbook object)
    workbook_choices = []
    seen = set()

    def add(label, key, getter):
        if key not in seen:
            seen.add(key)
            workbook_choices.append((label, getter))

    # Gather all open workbooks from all running Excel instances (xlwings).
    # xlwings cannot read an Excel instance whose window is showing a file in
    # Protected View, so that instance is read directly through Excel below.
    any_failed = False
    try:
        apps = list(xw.apps)
    except Exception:
        apps, any_failed = [], True
    for app in apps:
        try:
            for wb in app.books:
                add(wb.name, wb.fullname, lambda wb=wb: wb.api)
        except Exception:
            any_failed = True

    xl = _excel_com()
    if xl is not None:
        if any_failed:
            try:
                for i in range(1, xl.Workbooks.Count + 1):
                    wb = xl.Workbooks(i)
                    add(wb.Name, wb.FullName, lambda wb=wb: wb)
            except Exception:
                pass
        # files open in Protected View ('Enable Editing' bar)
        try:
            for i in range(1, xl.ProtectedViewWindows.Count + 1):
                pv = xl.ProtectedViewWindows(i)
                add("%s  (Protected View)" % pv.SourceName, "pv:" + pv.SourcePath + pv.SourceName,
                    lambda pv=pv: _open_protected(pv))
        except Exception:
            pass

    # If there are open workbooks, let user choose one
    if workbook_choices:
        print("Select from open workbooks:")
        for idx, (label, _) in enumerate(workbook_choices, start=1):
            print(f"{idx}: {label}")

        try:
            choice = int(input("Enter number or 0 to open a new file: ").strip())
            if 1 <= choice <= len(workbook_choices):
                return workbook_choices[choice - 1][1]()
        except Exception:
            pass

    # Fallback: use active Excel instance if one exists, otherwise create one
    app = None
    try:
        if xw.apps.count > 0:
            app = xw.apps.active
            if app is None:
                app = list(xw.apps)[0]
            app.api                             # fails for a Protected View window
    except Exception:
        app = None
    if app is None and xl is not None:          # open the file in the running Excel
        try:
            file_path = xl.GetOpenFilename(
                FileFilter="Excel Files (*.xlsx;*.xlsm), *.xlsx;*.xlsm",
                Title="Select an Excel workbook")
            if not file_path or file_path is False:
                return None
            return xl.Workbooks.Open(file_path)
        except Exception:
            pass
    if app is None:
        app = xw.App(visible=True, add_book=False)

    file_path = app.api.GetOpenFilename(
        FileFilter="Excel Files (*.xlsx;*.xlsm), *.xlsx;*.xlsm",
        Title="Select an Excel workbook"
    )

    if not file_path or file_path is False:
        return None

    return app.books.open(file_path).api


if __name__ == "__main__":
    book = select_or_open_workbook()
    print("Selected:", book.Name if book is not None else None)
