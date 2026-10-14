#!/usr/bin/env python3
r"""
pab_checker_gui.py
==================
Window version of the PAB checker (PySide6 / Qt).

All the checks live in design_workbook_check.py (the terminal version); this
file only adds the window, so any fix to a check is made once, in that file,
and both versions get it.  Keep the two files in the same folder, with the
branding folder (the Design Team icon) beside them.

Usage (Windows, Excel installed):

    pip install pyside6 pywin32 xlwings
    python pab_checker_gui.py

Rename it to pab_checker_gui.pyw to start it without a console window.

To build PAB Checker.exe, run build_exe.bat (or the PyInstaller command in it).
The branding folder is packed into the .exe, and the DOE QA script is run with
the Python installed on the PC.
"""

import html
import os
import subprocess
import sys
import traceback
from datetime import datetime

from PySide6.QtCore import QSettings, QTimer, Qt
from PySide6.QtGui import (QBrush, QColor, QCursor, QGuiApplication, QIcon, QKeySequence,
                           QPainter, QPalette, QPen, QPixmap, QRadialGradient, QShortcut)
from PySide6.QtWidgets import (QApplication, QCheckBox, QComboBox, QFileDialog, QHBoxLayout,
                               QHeaderView, QLabel, QMainWindow, QMessageBox, QPushButton,
                               QSizePolicy, QSplitter, QTextBrowser, QTreeWidget, QTreeWidgetItem,
                               QVBoxLayout, QWidget)

import design_workbook_check as core

APP_TITLE = "PAB Checker"
# Design Team icon (built by branding/make_dt_icon.py)
ICON_ICO = "dt_icon.ico"
ICON_PNGS = ("dt_icon_256.png", "dt_icon_1024.png")


def branding_dirs():
    """Where the branding folder can be: bundled inside a PyInstaller .exe
    (--add-data), next to the .exe, or next to this .py file."""
    dirs = []
    if getattr(sys, "frozen", False):
        dirs.append(os.path.join(getattr(sys, "_MEIPASS", ""), "branding"))
        dirs.append(os.path.join(os.path.dirname(os.path.abspath(sys.executable)), "branding"))
    dirs.append(os.path.join(os.path.dirname(os.path.abspath(__file__)), "branding"))
    return dirs
PLACEHOLDER = "Select a workbook..."

# ---- palettes: rack-panel greys, Cat6-blue accent, port-LED status colours ----
THEMES = {
    "light": dict(bg="#EEF1F4", surface="#FFFFFF", ink="#1C2733", muted="#5B6876", rule="#D3DAE1",
                  border="#C3CCD5", hover="#F3F6F9", select="#E3EDF8", pressed="#E3EAF2",
                  disabled="#9AA5B1", disabled_border="#DDE3E8", statusbar="#E4E8EC",
                  accent="#2563A6", highlight="#2563A6",
                  button="#2563A6", button_hover="#1E5290", button_border="#1E5290",
                  button_hover_border="#1E5290", button_disabled="#9DB6D3", button_disabled_text="#FFFFFF",
                  hover_bg="#FFFFFF", hover_border="#2563A6", hover_text="#2563A6", focus_border="#2563A6",
                  green="#2E8B3E", amber="#C77700", red="#C62828", note_bar="#C3CCD5"),
    # neutral near-black (no blue tint), matched to the VS Code editor background
    "dark": dict(bg="#121314", surface="#18191B", ink="#E4E5E7", muted="#9A9EA4", rule="#2A2C2F",
                 border="#3A3D42", hover="#1F2124", select="#1F3247", pressed="#26282C",
                 disabled="#5C6067", disabled_border="#2A2C2F", statusbar="#0E0F10",
                 accent="#6EA8E6", highlight="#3474BD",
                 button="#2A2A29", button_hover="#343433", button_border="#55595F",
                 button_hover_border="#55595F", button_disabled="#232322", button_disabled_text="#6A6E74",
                 hover_bg="#343433", hover_border="#55595F", hover_text="#E4E5E7", focus_border="#8A8F96",
                 green="#57C46F", amber="#F2A541", red="#F26B60", note_bar="#3A3D42"),
}
# current colours (set by apply_theme; read by the window when it draws)
SURFACE = INK = MUTED = GREEN = AMBER = RED = NOTE_BAR = ""

LABEL_FONT = "Bahnschrift, 'Segoe UI Semibold', 'Segoe UI', sans-serif"   # DIN-style, ships with Windows
BODY_FONT = "'Segoe UI', sans-serif"

STYLE = """
QMainWindow, QWidget#root {{ background: {bg}; }}
QWidget {{ font-family: "Segoe UI"; font-size: 10pt; color: {ink}; }}
QLabel#fieldLabel {{ color: {muted}; font-weight: 600; }}
QLabel#status {{ color: {muted}; }}
QComboBox {{
    background: {surface}; border: 1px solid {border}; border-radius: 4px;
    padding: 5px 10px; min-height: 22px;
}}
QComboBox:hover, QComboBox:focus {{ border-color: {focus_border}; }}
QComboBox QAbstractItemView {{
    background: {surface}; border: 1px solid {border};
    selection-background-color: {select}; selection-color: {ink};
}}
QPushButton {{
    background: {surface}; border: 1px solid {border}; border-radius: 4px;
    padding: 6px 16px; min-height: 22px;
}}
QPushButton:focus {{ border-color: {focus_border}; }}
QPushButton:hover {{ background: {hover_bg}; border-color: {hover_border}; color: {hover_text}; }}
QPushButton:pressed {{ background: {pressed}; }}
QPushButton:disabled {{ color: {disabled}; border-color: {disabled_border}; }}
QPushButton#primary {{
    background: {button}; color: white; border: 1px solid {button_border}; font-weight: 600;
}}
QPushButton#primary:hover {{ background: {button_hover}; border-color: {button_hover_border}; color: white; }}
QPushButton#primary:disabled {{
    background: {button_disabled}; border-color: {disabled_border}; color: {button_disabled_text};
}}
QTreeWidget, QTextBrowser {{
    background: {surface}; border: 1px solid {rule}; border-radius: 6px;
}}
QTreeWidget {{ padding: 4px 0; outline: 0; }}
QTreeWidget::item {{ padding: 6px 4px; border: none; }}
QTreeWidget::item:hover {{ background: {hover}; }}
QTreeWidget::item:selected {{ background: {select}; color: {ink}; }}
QHeaderView::section {{
    background: {surface}; border: none; border-bottom: 1px solid {rule};
    padding: 4px 8px 8px 8px; color: {muted}; font-weight: 600;
}}
QCheckBox {{ color: {muted}; spacing: 6px; }}
QSplitter::handle {{ background: transparent; }}
QStatusBar {{ background: {statusbar}; color: {muted}; border-top: 1px solid {rule}; }}
QStatusBar::item {{ border: none; }}
QStatusBar QCheckBox {{ padding: 2px 10px 2px 4px; }}
QToolTip {{ background: {surface}; color: {ink}; border: 1px solid {border}; }}
"""


def apply_theme(app, name):
    """Switch the whole application between the light and dark palettes."""
    t = THEMES[name]
    g = globals()
    for key, var in (("surface", "SURFACE"), ("ink", "INK"), ("muted", "MUTED"), ("green", "GREEN"),
                     ("amber", "AMBER"), ("red", "RED"), ("note_bar", "NOTE_BAR")):
        g[var] = t[key]
    # the Qt palette covers the parts the style sheet doesn't reach
    # (message boxes, scroll bars, menus)
    pal = QPalette()
    for role, key in ((QPalette.ColorRole.Window, "bg"), (QPalette.ColorRole.WindowText, "ink"),
                      (QPalette.ColorRole.Base, "surface"), (QPalette.ColorRole.AlternateBase, "hover"),
                      (QPalette.ColorRole.Text, "ink"), (QPalette.ColorRole.Button, "surface"),
                      (QPalette.ColorRole.ButtonText, "ink"), (QPalette.ColorRole.ToolTipBase, "surface"),
                      (QPalette.ColorRole.ToolTipText, "ink"), (QPalette.ColorRole.PlaceholderText, "muted"),
                      (QPalette.ColorRole.Highlight, "highlight"), (QPalette.ColorRole.Link, "accent")):
        pal.setColor(role, QColor(t[key]))
    pal.setColor(QPalette.ColorRole.HighlightedText, QColor("#FFFFFF"))
    for role in (QPalette.ColorRole.WindowText, QPalette.ColorRole.Text, QPalette.ColorRole.ButtonText):
        pal.setColor(QPalette.ColorGroup.Disabled, role, QColor(t["disabled"]))
    app.setPalette(pal)
    app.setStyleSheet(STYLE.format(**t))


def saved_dark_mode(settings):
    """The saved choice, or Windows' own light/dark setting the first time."""
    value = settings.value("dark_mode", None)
    if value is None:
        try:
            return QGuiApplication.styleHints().colorScheme() == Qt.ColorScheme.Dark
        except Exception:                       # Qt older than 6.5
            return False
    return str(value).lower() in ("true", "1")


def led_icon(color, size=14):
    """A small port-LED style status light."""
    pm = QPixmap(size, size)
    pm.fill(Qt.GlobalColor.transparent)
    p = QPainter(pm)
    p.setRenderHint(QPainter.RenderHint.Antialiasing)
    base = QColor(color)
    grad = QRadialGradient(size * 0.38, size * 0.35, size * 0.6)
    grad.setColorAt(0.0, base.lighter(165))
    grad.setColorAt(0.55, base)
    grad.setColorAt(1.0, base.darker(135))
    p.setBrush(QBrush(grad))
    p.setPen(QPen(base.darker(160), 1))
    p.drawEllipse(1, 1, size - 2, size - 2)
    p.end()
    return QIcon(pm)


def esc(text):
    """Plain message text -> HTML, keeping line breaks and indentation."""
    out = []
    for line in html.escape(text).split("\n"):
        stripped = line.lstrip(" ")
        out.append("&nbsp;" * (len(line) - len(stripped)) + stripped)
    return "<br>".join(out)


class CheckerWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.entries = []            # [(label, opener)] from core.list_open_workbooks()
        self.wb = None               # Excel workbook object last checked
        self.wb_label = None
        self.report = None
        self.checked_name = ""
        self.checked_at = None
        self.settings = QSettings("PAB Checker", "PAB Checker")
        self._make_leds()

        self.setWindowTitle(APP_TITLE)
        self.resize(1180, 760)
        self.setMinimumSize(820, 500)
        self._build()
        self.f5 = QShortcut(QKeySequence("F5"), self)
        self.f5.activated.connect(self.run_checks)
        QTimer.singleShot(100, self.refresh_workbooks)

    # ------------------------------------------------------------------ layout
    def _build(self):
        root = QWidget()
        root.setObjectName("root")
        outer = QVBoxLayout(root)
        outer.setContentsMargins(16, 14, 16, 12)
        outer.setSpacing(10)

        row1 = QHBoxLayout()
        row1.setSpacing(8)
        lbl = QLabel("Workbook")
        lbl.setObjectName("fieldLabel")
        self.combo = QComboBox()
        self.combo.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.combo.setMinimumContentsLength(30)
        self.combo.setSizeAdjustPolicy(QComboBox.SizeAdjustPolicy.AdjustToMinimumContentsLengthWithIcon)
        self.btn_refresh = QPushButton("Refresh list")
        self.btn_open = QPushButton("Open file...")
        self.btn_refresh.clicked.connect(lambda: self.refresh_workbooks())
        self.btn_open.clicked.connect(self.open_file)
        row1.addWidget(lbl)
        row1.addWidget(self.combo, 1)
        row1.addWidget(self.btn_refresh)
        row1.addWidget(self.btn_open)
        outer.addLayout(row1)

        row2 = QHBoxLayout()
        row2.setSpacing(8)
        self.btn_run = QPushButton("Run checks")
        self.btn_run.setObjectName("primary")
        self.btn_run.setToolTip("Read the workbook and run every check (F5)")
        self.btn_qa = QPushButton("Run DOE QA script")
        self.btn_copy = QPushButton("Copy results")
        self.btn_copy.setEnabled(False)
        self.btn_run.clicked.connect(self.run_checks)
        self.btn_qa.clicked.connect(self.run_qa)
        self.btn_copy.clicked.connect(self.copy_results)
        self.only_issues = QCheckBox("Only show tabs with issues")
        self.only_issues.toggled.connect(self._redraw)
        self.dark = QCheckBox("Dark mode")
        self.dark.setChecked(INK == THEMES["dark"]["ink"])
        self.dark.toggled.connect(self.set_dark_mode)
        row2.addWidget(self.btn_run)
        row2.addWidget(self.btn_qa)
        row2.addStretch(1)
        row2.addWidget(self.only_issues)
        row2.addSpacing(12)
        row2.addWidget(self.btn_copy)
        outer.addLayout(row2)

        self.summary = QLabel()
        self.summary.setTextFormat(Qt.TextFormat.RichText)
        self.summary.setContentsMargins(2, 6, 2, 4)
        outer.addWidget(self.summary)

        split = QSplitter(Qt.Orientation.Horizontal)
        split.setHandleWidth(10)
        self.tree = QTreeWidget()
        self.tree.setColumnCount(2)
        self.tree.setHeaderLabels(["Tab", "Result"])
        self.tree.setRootIsDecorated(False)
        self.tree.setIndentation(0)
        self.tree.header().setStretchLastSection(False)
        self.tree.header().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.tree.header().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.tree.itemSelectionChanged.connect(self._jump)
        self.view = QTextBrowser()
        self.view.setOpenLinks(False)
        self.view.document().setDocumentMargin(18)
        split.addWidget(self.tree)
        split.addWidget(self.view)
        split.setStretchFactor(0, 1)
        split.setStretchFactor(1, 4)
        split.setSizes([290, 870])
        outer.addWidget(split, 1)

        self.setCentralWidget(root)
        self.status = QLabel()
        self.status.setObjectName("status")
        self.statusBar().addWidget(self.status, 1)
        self.statusBar().addPermanentWidget(self.dark)          # bottom right corner
        self._empty_state()

    def _make_leds(self):
        self.leds = {k: led_icon(c) for k, c in (("error", RED), ("warn", AMBER), ("ok", GREEN))}

    def set_dark_mode(self, on):
        apply_theme(QApplication.instance(), "dark" if on else "light")
        self.settings.setValue("dark_mode", bool(on))
        self._make_leds()
        if self.report is not None:
            self._redraw()
        else:
            self._empty_state()

    def _empty_state(self):
        self.summary.setText('<span style="font-family:%s; font-size:15pt; color:%s">'
                             'Pick a workbook, then run the checks.</span>' % (LABEL_FONT, MUTED))
        self.view.setHtml('<p style="font-family:%s; color:%s">Results appear here, grouped by tab. '
                          'After fixing something in Excel, press <b>Run checks</b> (or F5) again '
                          'to re-read the same workbook.</p>' % (BODY_FONT, MUTED))

    # --------------------------------------------------------------- workbooks
    def refresh_workbooks(self, select=None):
        keep = select or self.combo.currentText()
        self._busy("Looking for open workbooks...")
        try:
            self.entries = core.list_open_workbooks()
        except Exception as exc:
            self.entries = []
            self._idle()
            self._error("Could not read the open Excel workbooks.", exc)
        self._idle()
        labels = [label for label, _ in self.entries]
        self.combo.blockSignals(True)
        self.combo.clear()
        self.combo.addItem(PLACEHOLDER)
        self.combo.addItems(labels)
        self.combo.setCurrentIndex(labels.index(keep) + 1 if keep in labels else 0)
        self.combo.blockSignals(False)
        self.status.setText("%d open workbook%s found." % (len(labels), "" if len(labels) == 1 else "s")
                            if labels else "No open workbooks. Open one in Excel, or use Open file...")

    def open_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "Select an Excel workbook", "",
                                              "Excel workbooks (*.xlsx *.xlsm)")
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
        name = str(wb.Name)
        self.refresh_workbooks(select=name)
        self.wb, self.wb_label = wb, name
        if self.combo.findText(name) < 0:
            self.combo.addItem(name)
        self.combo.setCurrentText(name)
        self.run_checks()

    def _ask(self, question):
        return QMessageBox.question(self, APP_TITLE, question) == QMessageBox.StandardButton.Yes

    def _selected_workbook(self):
        """Workbook object for the current choice; reuses the last one on a re-run."""
        label = self.combo.currentText()
        if not label or label == PLACEHOLDER:
            QMessageBox.information(self, APP_TITLE, "Choose a workbook first.")
            return None
        if self.wb is not None and label == self.wb_label:
            return self.wb
        opener = dict(self.entries).get(label)
        if opener is None:
            self.refresh_workbooks()
            opener = dict(self.entries).get(label)
        if opener is None:
            QMessageBox.warning(self, APP_TITLE, "%s is no longer open in Excel." % label)
            return None
        wb = opener(confirm=self._ask)
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
        self.report, self.checked_name, self.checked_at = rep, name, datetime.now()
        self.btn_copy.setEnabled(True)
        self._redraw()
        self.status.setText("Checked %s at %s" % (name, self.checked_at.strftime("%H:%M")))

    def _redraw(self, *_):
        rep = self.report
        if rep is None:
            return
        n_err, n_warn = rep.count("ERROR"), rep.count("WARN")

        def count(n, word, color):
            return ('<span style="font-family:%s; font-size:22pt; color:%s">%d</span>'
                    '<span style="font-size:11pt; color:%s">&nbsp;%s%s</span>'
                    % (LABEL_FONT, color, n, MUTED, word, "" if n == 1 else "s"))
        if n_err or n_warn:
            head = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;".join(
                [count(n_err, "error", RED if n_err else MUTED),
                 count(n_warn, "warning", AMBER if n_warn else MUTED)])
        else:
            head = ('<span style="font-family:%s; font-size:22pt; color:%s">No issues found</span>'
                    % (LABEL_FONT, GREEN))
        self.summary.setText('%s<span style="font-size:10pt; color:%s">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;'
                             'in %s</span>' % (head, MUTED, html.escape(self.checked_name)))

        self.tree.blockSignals(True)
        self.tree.clear()
        parts = ['<div style="font-family:%s; color:%s">' % (BODY_FONT, INK)]
        first = True
        for i, (tab, items) in enumerate(rep.sections(core.tab_order())):
            e = sum(1 for lv, _ in items if lv == "ERROR")
            w = sum(1 for lv, _ in items if lv == "WARN")
            if e:
                result, key, color = "%d error%s" % (e, "" if e == 1 else "s"), "error", RED
            elif w:
                result, key, color = "%d warning%s" % (w, "" if w == 1 else "s"), "warn", AMBER
            else:
                result, key, color = "OK", "ok", GREEN
            if self.only_issues.isChecked() and key == "ok":
                continue
            item = QTreeWidgetItem([tab, result])
            item.setIcon(0, self.leds[key])
            item.setForeground(1, QBrush(QColor(color)))
            item.setData(0, Qt.ItemDataRole.UserRole, "sec%d" % i)
            self.tree.addTopLevelItem(item)

            parts.append('<a name="sec%d"></a><p style="font-family:%s; font-size:14pt; color:%s; '
                         'margin-top:%dpx; margin-bottom:6px">%s</p>'
                         % (i, LABEL_FONT, INK, 0 if first else 20, html.escape(tab)))
            first = False
            for lv, msg in items:
                parts.append(self._row(lv, msg))
            if not e and not w:
                parts.append(self._row("OK", "No issues found in %s." % tab))
        if first:
            parts.append('<p style="color:%s">Every tab passed.</p>' % GREEN)
        parts.append("</div>")
        self.tree.blockSignals(False)
        self.view.setHtml("".join(parts))
        self.view.verticalScrollBar().setValue(0)

    @staticmethod
    def _row(level, msg):
        bar, label, text_color = {
            "ERROR": (RED, "Error", INK),
            "WARN": (AMBER, "Warning", INK),
            "NOTE": (NOTE_BAR, "", MUTED),
            "OK": (GREEN, "", GREEN),
        }[level]
        lead = ('<span style="color:%s; font-weight:600">%s</span>&nbsp;&nbsp;' % (bar, label)) if label else ""
        return ('<table width="100%%" cellspacing="0" cellpadding="0" style="margin-bottom:5px">'
                '<tr><td width="3" bgcolor="%s"></td><td width="10"></td>'
                '<td style="color:%s; padding-top:2px; padding-bottom:2px">%s%s</td></tr></table>'
                % (bar, text_color, lead, esc(msg)))

    def _jump(self):
        items = self.tree.selectedItems()
        if items:
            self.view.scrollToAnchor(items[0].data(0, Qt.ItemDataRole.UserRole))

    def copy_results(self):
        if self.report is None:
            return
        QGuiApplication.clipboard().setText(
            "Workbook: %s\n%s" % (self.checked_name, self.report.render(core.tab_order())))
        self.status.setText("Results copied to the clipboard.")

    # ------------------------------------------------------------ DOE QA script
    def run_qa(self):
        if self.report is not None and self.report.count("ERROR"):
            if not self._ask("The last check found %d error(s).\n\nRun the DOE QA script anyway?"
                             % self.report.count("ERROR")):
                return
        script = self._find_qa_script()
        if not script or not self._qa_folder_ready(os.path.dirname(script)):
            return
        python = core.python_for_scripts()      # an .exe can't run .py files itself
        if not python:
            QMessageBox.warning(self, APP_TITLE, "Python was not found on this PC, so the DOE QA script "
                                "can't be started from here.\n\nRun it from VS Code or a terminal instead.")
            return
        flags = getattr(subprocess, "CREATE_NEW_CONSOLE", 0)
        try:
            subprocess.Popen(python + [script], cwd=os.path.dirname(script), creationflags=flags)
        except Exception as exc:
            self._error("Could not start the DOE QA script.", exc)
            return
        self.close()                            # the QA script carries on in its own window

    def _choose(self, text, labels):
        """Question box with custom button labels -> the label clicked (last one = cancel)."""
        box = QMessageBox(self)
        box.setIcon(QMessageBox.Icon.Question)
        box.setWindowTitle(APP_TITLE)
        box.setText(text)
        buttons = {}
        for i, label in enumerate(labels):
            role = (QMessageBox.ButtonRole.RejectRole if i == len(labels) - 1
                    else QMessageBox.ButtonRole.AcceptRole if i == 0 else QMessageBox.ButtonRole.ActionRole)
            buttons[label] = box.addButton(label, role)
        box.setDefaultButton(buttons[labels[0]])
        box.exec()
        clicked = box.clickedButton()
        return next((label for label, b in buttons.items() if b is clicked), labels[-1])

    def _try_again(self):
        answer = QMessageBox.question(self, APP_TITLE, "DOE QA script aborted.\n\nTry again?",
                                      QMessageBox.StandardButton.Retry | QMessageBox.StandardButton.Cancel)
        return answer == QMessageBox.StandardButton.Retry

    def _find_qa_script(self):
        folder, name = core.load_qa_settings()
        if folder and not os.path.isdir(folder):
            QMessageBox.information(self, APP_TITLE, "The saved DOE QA script folder is no longer "
                                    "accessible:\n\n%s\n\nPlease select it again." % folder)
            folder = ""
        while True:
            if not folder:
                folder = QFileDialog.getExistingDirectory(self, "Select the folder with the DOE QA script")
                folder = os.path.normpath(folder) if folder else ""
            if folder:
                path = os.path.join(folder, name)
                if not os.path.isfile(path):
                    path = ""
                    if self._ask("%s was not found in:\n\n%s\n\nDo you want to locate the Python script "
                                 "in this folder?" % (name, folder)):
                        path, _ = QFileDialog.getOpenFileName(self, "Select the DOE QA Python script",
                                                              folder, "Python scripts (*.py)")
                        path = os.path.normpath(path) if path and os.path.isfile(path) else ""
                if path:
                    core.save_qa_settings(os.path.dirname(path), os.path.basename(path))
                    return path
                folder = ""
            if not self._try_again():
                return None

    def _qa_folder_ready(self, folder):
        while True:
            try:
                files = core.qa_xlsx_files(folder)
                problem = ""
            except OSError as exc:
                files, problem = None, "Could not read the folder:\n%s\n\n%s" % (folder, exc)
            if files is not None and len(files) == 1:
                found = files[0]
                if not self.checked_name or core.same_workbook(found, self.checked_name):
                    note = ("This matches the workbook you just checked." if self.checked_name
                            else "No workbook has been checked yet in this window.")
                    if self._ask("Excel file found in the DOE QA folder:\n\n%s\n\n%s\n\n"
                                 "Run the DOE QA script on this file?" % (found, note)):
                        return True
                    QMessageBox.information(self, APP_TITLE, "DOE QA script aborted.")
                    return False
                choice = self._choose(
                    "Excel file found in the DOE QA folder:\n\n%s\n\nThis is not the workbook you just "
                    "checked:\n\n%s\n\nDo you want to continue with this file?" % (found, self.checked_name),
                    ["Continue", "Check folder again", "Cancel"])
                if choice == "Continue":
                    return True
                if choice == "Check folder again":
                    continue
                QMessageBox.information(self, APP_TITLE, "DOE QA script aborted.")
                return False
            if files is not None:
                if not files:
                    problem = "No Excel file (.xlsx) was found in:\n%s" % folder
                else:
                    problem = "%d Excel files (.xlsx) were found in:\n%s\n\n%s" % (
                        len(files), folder, "\n".join("    " + f for f in files))
                problem += "\n\nThere must be exactly 1 .xlsx file in this folder for the DOE QA script."
            answer = QMessageBox.warning(self, APP_TITLE, problem + "\n\nFix the folder, then press Retry "
                                         "to check again.",
                                         QMessageBox.StandardButton.Retry | QMessageBox.StandardButton.Cancel)
            if answer != QMessageBox.StandardButton.Retry:
                QMessageBox.information(self, APP_TITLE, "DOE QA script aborted.")
                return False

    # ----------------------------------------------------------------- helpers
    def _busy(self, msg):
        self.status.setText(msg)
        QApplication.setOverrideCursor(QCursor(Qt.CursorShape.WaitCursor))
        for b in (self.btn_run, self.btn_qa, self.btn_refresh, self.btn_open):
            b.setEnabled(False)
        QApplication.processEvents()

    def _idle(self):
        while QApplication.overrideCursor() is not None:
            QApplication.restoreOverrideCursor()
        for b in (self.btn_run, self.btn_qa, self.btn_refresh, self.btn_open):
            b.setEnabled(True)

    def _error(self, msg, exc):
        traceback.print_exc()
        QMessageBox.critical(self, APP_TITLE, "%s\n\n%s: %s" % (msg, type(exc).__name__, exc))


def load_icon(app):
    """Set the Design Team icon; say where it looked if it can't be loaded.

    The .ico is tried first, but Qt reads .ico through a plugin that a PyInstaller
    build doesn't always include; then the icon comes out blank (Windows shows a
    generic one).  availableSizes() is empty in that case, so the PNGs are used."""
    for folder in branding_dirs():
        ico = os.path.join(folder, ICON_ICO)
        if os.path.isfile(ico):
            icon = QIcon(ico)
            if icon.availableSizes():
                app.setWindowIcon(icon)
                return True
        icon = QIcon()
        for name in ICON_PNGS:
            path = os.path.join(folder, name)
            if os.path.isfile(path):
                icon.addFile(path)
        if not icon.isNull():
            app.setWindowIcon(icon)
            return True
    print("Design Team icon not found. Looked for the 'branding' folder in:\n  %s"
          % "\n  ".join(branding_dirs()))
    return False


def main():
    if sys.platform == "win32":
        # Without this, Windows groups the window under python.exe in the taskbar and
        # shows Python's icon there instead of ours.
        try:
            import ctypes
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("DesignTeam.PABChecker")
        except Exception:
            pass
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    load_icon(app)
    apply_theme(app, "dark" if saved_dark_mode(QSettings("PAB Checker", "PAB Checker")) else "light")
    if core.pythoncom is None or core.xw is None:
        QMessageBox.critical(None, APP_TITLE, "pywin32 and xlwings are needed.\n\n"
                             "Run:  pip install pywin32 xlwings")
        return 1
    win = CheckerWindow()
    win.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
