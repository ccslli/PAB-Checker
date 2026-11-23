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

import base64
import html
import os
import subprocess
import sys
import time
import traceback
from datetime import datetime

from PySide6.QtCore import QRectF, QSettings, QThread, QTimer, Qt, Signal
from PySide6.QtGui import (QBrush, QColor, QCursor, QGuiApplication, QIcon, QKeySequence,
                           QPainter, QPalette, QPen, QPixmap, QRadialGradient, QShortcut)
from PySide6.QtWidgets import (QApplication, QButtonGroup, QCheckBox, QComboBox, QDialog, QFileDialog, QFrame,
                               QHBoxLayout, QHeaderView, QLabel, QLineEdit, QMainWindow, QMessageBox, QPushButton,
                               QRadioButton, QSizePolicy, QSplitter, QTextBrowser, QTreeWidget, QTreeWidgetItem,
                               QVBoxLayout, QWidget)

import design_workbook_check as core

try:                                    # the DT icon embedded as data (branding/dt_icon_data.py)
    from branding.dt_icon_data import PNGS as EMBEDDED_ICON
except ImportError:
    EMBEDDED_ICON = {}

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
SURFACE = INK = MUTED = GREEN = AMBER = RED = NOTE_BAR = ACCENT = ""

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
QPushButton#themeToggle {{
    background: {surface}; color: {ink}; border: 1px solid {border}; border-radius: 4px;
    padding: 1px 10px; min-height: 18px; margin: 3px 8px 3px 4px; font-size: 9pt;
}}
QPushButton#themeToggle:hover {{ border-color: {focus_border}; }}
QPushButton#hiddenSpot, QPushButton#hiddenSpot:hover, QPushButton#hiddenSpot:pressed,
QPushButton#hiddenSpot:focus {{ background: transparent; border: none; padding: 0; margin: 0; }}
QToolTip {{ background: {surface}; color: {ink}; border: 1px solid {border}; }}
QFrame#loaderCard {{ background: {surface}; border: 1px solid {border}; border-radius: 10px; }}
QLabel#loaderTitle {{ font-family: "Bahnschrift"; font-size: 14pt; color: {ink}; }}
QLabel#loaderFile, QLabel#loaderElapsed {{ color: {muted}; }}
QLabel#loaderStep {{ color: {ink}; }}
"""


def apply_theme(app, name):
    """Switch the whole application between the light and dark palettes."""
    t = THEMES[name]
    g = globals()
    for key, var in (("surface", "SURFACE"), ("ink", "INK"), ("muted", "MUTED"), ("green", "GREEN"),
                     ("amber", "AMBER"), ("red", "RED"), ("note_bar", "NOTE_BAR"),
                     ("accent", "ACCENT")):
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


class PortLights(QWidget):
    """Loading animation: a row of port lights blinking in sequence, like link
    activity on a switch.  The row fills whatever width it is given."""

    SLOT = 24                           # space per light

    def __init__(self, parent=None):
        super().__init__(parent)
        self.step = 0
        self.setFixedHeight(24)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.timer = QTimer(self)
        self.timer.timeout.connect(self._tick)
        self.timer.start(85)

    def _count(self):
        return max(6, self.width() // self.SLOT)

    def _tick(self):
        self.step += 1
        self.update()

    def stop(self):
        self.timer.stop()

    def paintEvent(self, _event):
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        p.setPen(Qt.PenStyle.NoPen)
        on, off = QColor(ACCENT), QColor(NOTE_BAR)
        n = self._count()
        span = self.width() - 12 - 8                # first light at the left edge, last at the right
        lit = self.step % n
        for i in range(n):
            level = max(0.0, 1.0 - ((lit - i) % n) / 3.0)                   # lit light + fading trail
            x = 4 + i * span / (n - 1)
            if level > 0.5:                                                  # soft glow round the lit one
                p.setBrush(QColor(on.red(), on.green(), on.blue(), int(60 * level)))
                p.drawEllipse(QRectF(x - 4, 2, 20, 20))
            c = QColor(int(off.red() + (on.red() - off.red()) * level),
                       int(off.green() + (on.green() - off.green()) * level),
                       int(off.blue() + (on.blue() - off.blue()) * level))
            p.setBrush(c)
            p.drawEllipse(QRectF(x, 6, 12, 12))
        p.end()


class LoaderDialog(QDialog):
    """Small box shown while a workbook is checked: file name, current step,
    port-light animation and a seconds counter.  Can't be dismissed by the user."""

    def __init__(self, parent, name):
        super().__init__(parent)
        self.setWindowFlags(Qt.WindowType.Dialog | Qt.WindowType.FramelessWindowHint)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setModal(True)
        self._finished = False
        card = QFrame(self)
        card.setObjectName("loaderCard")
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.addWidget(card)
        lay = QVBoxLayout(card)
        lay.setContentsMargins(26, 22, 26, 20)
        lay.setSpacing(8)
        title = QLabel("Checking workbook")
        title.setObjectName("loaderTitle")
        self.file = QLabel()
        self.file.setObjectName("loaderFile")
        self.file.setText(self.file.fontMetrics().elidedText(name, Qt.TextElideMode.ElideMiddle, 360))
        self.file.setToolTip(name)
        self.lights = PortLights()
        self.step = QLabel("Opening the workbook...")
        self.step.setObjectName("loaderStep")
        self.elapsed = QLabel("0 s")
        self.elapsed.setObjectName("loaderElapsed")
        row = QHBoxLayout()
        row.addWidget(self.step, 1)
        row.addWidget(self.elapsed)
        lay.addWidget(title)
        lay.addWidget(self.file)
        lay.addSpacing(8)
        lay.addWidget(self.lights)
        lay.addSpacing(4)
        lay.addLayout(row)
        self.setFixedWidth(420)
        self.t0 = time.monotonic()
        self.clock = QTimer(self)
        self.clock.timeout.connect(self._tick)
        self.clock.start(250)

    def _tick(self):
        self.elapsed.setText("%d s" % int(time.monotonic() - self.t0))

    def set_step(self, text):
        self.step.setText(text + "...")

    def show_centered(self):
        self.adjustSize()
        centre = self.parent().frameGeometry().center()
        self.move(centre.x() - self.width() // 2, centre.y() - self.height() // 2)
        self.show()
        QApplication.processEvents()

    def finish(self):
        self._finished = True
        self.lights.stop()
        self.clock.stop()
        self.done(0)

    def reject(self):                   # Esc does nothing while the check runs
        if self._finished:
            super().reject()

    def closeEvent(self, event):        # nor does Alt+F4
        if self._finished:
            event.accept()
        else:
            event.ignore()


class CheckWorker(QThread):
    """Runs the checks away from the window's thread, so the window (and the
    loading animation) keep moving.  Excel objects can't be shared between
    threads directly, so the workbook is handed over as a marshalled stream."""
    progress = Signal(str)
    done_ok = Signal(object)
    failed = Signal(str)

    def __init__(self, stream, fix_wireless, parent=None):
        super().__init__(parent)
        self.stream = stream
        self.fix_wireless = fix_wireless

    def run(self):
        core.pythoncom.CoInitialize()
        try:
            import win32com.client
            disp = core.pythoncom.CoGetInterfaceAndReleaseStream(self.stream, core.pythoncom.IID_IDispatch)
            wb = win32com.client.Dispatch(disp)
            self.done_ok.emit(core.check_workbook_report(wb, self.progress.emit, self.fix_wireless))
        except Exception:
            self.failed.emit(traceback.format_exc())
        finally:
            core.pythoncom.CoUninitialize()


class CheckerWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.entries = []            # [(label, opener)] from core.list_open_workbooks()
        self.wb = None               # Excel workbook object last checked
        self.wb_label = None
        self.report = None
        self.checked_name = ""
        self.checked_at = None
        self.worker = None           # background check in progress
        self.loader = None
        self.pending_name = ""
        self.settings = QSettings("PAB Checker", "PAB Checker")
        # restricted features: only for the Windows users in core.SPECIAL_USERS
        self.special = core.special_access()
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
        self.fix_wireless = QCheckBox("Fix network tags in WDD")
        self.fix_wireless.setToolTip("Before each check, convert Wireless_Design_Data column %s "
                                     "(row %d down) to numbers and format it \"0\". "
                                     "The result is shown at the bottom of the results."
                                     % (core.WDATA_FIX_COL, core.WDATA_FIX_FIRST_ROW))
        self.fix_wireless.setChecked(core.wireless_fix_enabled())
        self.fix_wireless.toggled.connect(core.set_wireless_fix_enabled)
        # bottom-right button: names the mode it switches TO ("Dark mode" while light)
        self.dark = QPushButton()
        self.dark.setObjectName("themeToggle")
        self.dark.setCursor(Qt.CursorShape.PointingHandCursor)
        self.dark.clicked.connect(lambda: self.set_dark_mode(INK != THEMES["dark"]["ink"]))
        self._label_theme_button()
        row2.addWidget(self.btn_run)
        row2.addWidget(self.btn_qa)
        row2.addSpacing(12)
        row2.addWidget(self.fix_wireless)
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
        if self.special:
            # restricted tools: an invisible button just left of the theme button
            # (core.SPECIAL_USERS only - for everyone else it doesn't exist)
            self.hidden_spot = QPushButton()
            self.hidden_spot.setObjectName("hiddenSpot")
            self.hidden_spot.setFixedSize(26, 22)
            self.hidden_spot.setFocusPolicy(Qt.FocusPolicy.NoFocus)
            self.hidden_spot.setFlat(True)
            self.hidden_spot.clicked.connect(self.open_special_tools)
            self.statusBar().addPermanentWidget(self.hidden_spot)
        self.statusBar().addPermanentWidget(self.dark)          # bottom right corner
        self._empty_state()

    def _make_leds(self):
        self.leds = {k: led_icon(c) for k, c in (("error", RED), ("warn", AMBER), ("ok", GREEN))}

    def _label_theme_button(self):
        dark = INK == THEMES["dark"]["ink"]
        self.dark.setText("Light mode" if dark else "Dark mode")
        self.dark.setToolTip("Switch to the %s theme" % ("light" if dark else "dark"))

    def set_dark_mode(self, on):
        apply_theme(QApplication.instance(), "dark" if on else "light")
        self.settings.setValue("dark_mode", bool(on))
        self._label_theme_button()
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
        if self.worker is not None:             # a check is already running
            return
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
        self.pending_name = name
        self._busy("Checking %s..." % name, cursor=False)
        self.loader = LoaderDialog(self, name)
        self.loader.show_centered()

        stream = None
        try:                                    # hand the workbook to the background thread
            com = getattr(wb, "_inner", wb)     # xlwings wraps the Excel object
            stream = core.pythoncom.CoMarshalInterThreadInterfaceInStream(
                core.pythoncom.IID_IDispatch, com._oleobj_)
        except Exception:
            stream = None
        if stream is None:                      # fall back to checking on this thread
            try:
                rep = core.check_workbook_report(wb, self._step_here, self.fix_wireless.isChecked())
            except Exception:
                self._check_failed(traceback.format_exc())
                return
            self._check_done(rep)
            return
        self.worker = CheckWorker(stream, self.fix_wireless.isChecked(), self)
        self.worker.progress.connect(self.loader.set_step)
        self.worker.done_ok.connect(self._check_done)
        self.worker.failed.connect(self._check_failed)
        self.worker.start()

    def _step_here(self, text):
        if self.loader:
            self.loader.set_step(text)
        QApplication.processEvents()

    def _end_check(self):
        if self.loader:
            self.loader.finish()
            self.loader = None
        if self.worker:
            self.worker.wait(5000)
            self.worker = None
        self._idle()

    def _check_done(self, rep):
        self._end_check()
        name = self.pending_name
        self.report, self.checked_name, self.checked_at = rep, name, datetime.now()
        self.btn_copy.setEnabled(True)
        self._redraw()
        self.status.setText("Checked %s at %s" % (name, self.checked_at.strftime("%H:%M")))

    def _check_failed(self, tb):
        self._end_check()
        self.wb = None
        print(tb)
        last = tb.strip().splitlines()[-1] if tb.strip() else "unknown error"
        QMessageBox.critical(self, APP_TITLE, "The checks could not finish.\n\n%s" % last)

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
        for j, (title, body) in enumerate(getattr(rep, "footer", [])):
            failed = any(line.startswith("Could not") for line in body)
            item = QTreeWidgetItem([title, "Failed" if failed else "Done"])
            item.setIcon(0, self.leds["warn" if failed else "ok"])
            item.setForeground(1, QBrush(QColor(AMBER if failed else MUTED)))
            item.setData(0, Qt.ItemDataRole.UserRole, "foot%d" % j)
            self.tree.addTopLevelItem(item)
            parts.append('<a name="foot%d"></a><p style="font-family:%s; font-size:14pt; color:%s; '
                         'margin-top:24px; margin-bottom:6px">%s</p>'
                         % (j, LABEL_FONT, INK, html.escape(title)))
            for line in body:
                parts.append(self._row("WARN" if line.startswith("Could not") else "NOTE", line))
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

    def open_special_tools(self):
        if not self.special:
            return
        dlg = getattr(self, "special_dialog", None)
        if dlg is None:
            dlg = self.special_dialog = ScanVlanDialog(self)
        dlg.refresh_notes()
        dlg.show()
        dlg.raise_()
        dlg.activateWindow()

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
    def _busy(self, msg, cursor=True):
        self.status.setText(msg)
        if cursor:
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


class ScanVlanDialog(QDialog):
    """Restricted tool (core.SPECIAL_USERS): QA Check - site scan vs Notes and the design file."""

    def __init__(self, win):
        super().__init__(win)
        import scan_vlan_check as svc
        self.svc, self.win = svc, win
        self.settings = win.settings
        self.setWindowTitle("QA Check")
        self.resize(900, 640)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(16, 14, 16, 12)
        lay.setSpacing(10)

        self.notes_label = QLabel()
        self.notes_label.setTextFormat(Qt.TextFormat.RichText)
        lay.addWidget(self.notes_label)

        fmt_row = QHBoxLayout()
        lab = QLabel("Scan format")
        lab.setObjectName("fieldLabel")
        fmt_row.addWidget(lab)
        self.fmt_group = QButtonGroup(self)
        self.fmt_buttons = {}
        saved = self.settings.value("scan_vlans/format", "new") or "new"
        tips = {"original": "The older scan files (VLAN in column E)",
                "new": "Network Scanner's Interfaces Report (VLAN in column F)"}
        for key, f in svc.FORMATS.items():
            rb = QRadioButton(f["title"])
            rb.setToolTip(tips.get(key, ""))
            rb.setChecked(key == saved)
            self.fmt_group.addButton(rb)
            self.fmt_buttons[key] = rb
            fmt_row.addWidget(rb)
        fmt_row.addStretch(1)
        lay.addLayout(fmt_row)

        file_row = QHBoxLayout()
        lab = QLabel("Scan file")
        lab.setObjectName("fieldLabel")
        self.path = QLineEdit(self.settings.value("scan_vlans/file", "") or "")
        self.path.setPlaceholderText("Browse... or paste the scan file's path")
        browse = QPushButton("Browse...")
        browse.clicked.connect(self.browse)
        self.btn_run = QPushButton("Check scan")
        self.btn_run.setObjectName("primary")
        self.btn_run.clicked.connect(self.run)
        file_row.addWidget(lab)
        file_row.addWidget(self.path, 1)
        file_row.addWidget(browse)
        file_row.addWidget(self.btn_run)
        lay.addLayout(file_row)

        design_row = QHBoxLayout()
        lab = QLabel("Design file")
        lab.setObjectName("fieldLabel")
        self.design = QLineEdit(self.settings.value("scan_vlans/design", "") or "")
        self.design.setPlaceholderText("Optional - its Drop List Draft is checked against the scan")
        browse2 = QPushButton("Browse...")
        browse2.clicked.connect(self.browse_design)
        design_row.addWidget(lab)
        design_row.addWidget(self.design, 1)
        design_row.addWidget(browse2)
        lay.addLayout(design_row)

        self.view = QTextBrowser()
        self.view.setOpenLinks(False)
        self.view.document().setDocumentMargin(14)
        lay.addWidget(self.view, 1)

        bottom = QHBoxLayout()
        self.btn_copy = QPushButton("Copy results")
        self.btn_copy.setEnabled(False)
        self.btn_copy.clicked.connect(lambda: QGuiApplication.clipboard().setText(self.text))
        close = QPushButton("Close")
        close.clicked.connect(self.close)
        bottom.addStretch(1)
        bottom.addWidget(self.btn_copy)
        bottom.addWidget(close)
        lay.addLayout(bottom)
        self.text = ""
        self.view.setHtml('<p style="color:%s">Pick the scan file and its format, then Check scan.</p>' % MUTED)

    def fmt(self):
        return next(k for k, b in self.fmt_buttons.items() if b.isChecked())

    def notes(self):
        rep = self.win.report
        parsed = getattr(rep, "parsed", None) if rep is not None else None
        return (parsed or {}).get("notes")

    def refresh_notes(self):
        if self.notes() is not None:
            self.notes_label.setText('<span style="color:%s">Notes from</span> <b>%s</b> '
                                     '<span style="color:%s">(checked %s)</span>'
                                     % (MUTED, esc(self.win.checked_name), MUTED,
                                        self.win.checked_at.strftime("%H:%M")))
        else:
            self.notes_label.setText('<span style="color:%s">Run checks on the workbook first - '
                                     'its Notes are what the scan is compared with.</span>' % AMBER)

    def browse(self):
        start = self.path.text().strip() or self.settings.value("scan_vlans/folder", "") or ""
        path, _ = QFileDialog.getOpenFileName(self, "Select the scan file", start,
                                              "Scan files (*.xlsx *.xlsm *.csv);;All files (*.*)")
        if path:
            self.path.setText(os.path.normpath(path))
            self.settings.setValue("scan_vlans/folder", os.path.dirname(path))

    def browse_design(self):
        start = self.design.text().strip() or self.settings.value("scan_vlans/design_folder", "") or ""
        path, _ = QFileDialog.getOpenFileName(self, "Select the design file", start,
                                              "Excel files (*.xlsm *.xlsx);;All files (*.*)")
        if path:
            self.design.setText(os.path.normpath(path))
            self.settings.setValue("scan_vlans/design_folder", os.path.dirname(path))

    def run(self):
        notes = self.notes()
        self.refresh_notes()
        if notes is None:
            QMessageBox.information(self, "QA Check", "Run checks on the workbook first.")
            return
        path = self.path.text().strip().strip('"')
        if not os.path.isfile(path):
            QMessageBox.warning(self, "QA Check", "Can't find the scan file:\n%s" % (path or "(none)"))
            return
        design = self.design.text().strip().strip('"')
        if design and not os.path.isfile(design):
            QMessageBox.warning(self, "QA Check", "Can't find the design file:\n%s" % design)
            return
        fmt = self.fmt()
        self.settings.setValue("scan_vlans/design", design)
        self.settings.setValue("scan_vlans/file", path)
        self.settings.setValue("scan_vlans/format", fmt)
        QApplication.setOverrideCursor(QCursor(Qt.CursorShape.WaitCursor))
        try:
            ports, info, every = self.svc.read_scan(path, fmt)
        except Exception as exc:
            QApplication.restoreOverrideCursor()
            QMessageBox.critical(self, "QA Check", "The scan file couldn't be read.\n\n%s" % exc)
            return
        dld, index = None, None
        if design:
            try:
                rows, source = self.svc.read_design(design)
                index = self.svc.design_index(rows)
                last = max((d["row"] for v in index.values() for d in v), default=0)
                dld = (design, "%s (%d rows read, last switch / port on row %d)" % (source, len(rows), last),
                       len(index), self.svc.check_design(ports, index))
            except Exception as exc:
                dld = (design, None, 0, [("ERROR", "The design file couldn't be read: %s" % exc, None)])
        groups, other = self.svc.check(ports, notes, every, index)
        QApplication.restoreOverrideCursor()
        self.show_results(path, fmt, info, groups, other, dld, index)

    def show_results(self, path, fmt, info, groups, other, dld=None, index=None):
        svc, row = self.svc, CheckerWindow._row
        html_parts, text = [], []

        def head(title):
            html_parts.append('<p style="font-family:%s; font-size:14pt; color:%s; margin-top:16px; '
                              'margin-bottom:6px">%s</p>' % (LABEL_FONT, INK, esc(title)))
            text.extend(["", title, "-" * len(title)])

        def line(level, msg):
            html_parts.append(row(level, msg))
            text.append(("[%s] " % level if level in ("ERROR", "WARN") else "") + msg)

        bad = sum(1 for g in groups if not g["ok"])
        head("%s  (%s)" % (os.path.basename(path), svc.FORMATS[fmt]["title"]))
        line("NOTE", "Compared with the Notes of %s." % self.win.checked_name)
        for i in info:
            line("NOTE", i)

        head("Special VLANs vs Notes")
        for g in groups:
            lab = "%s (%s)" % (svc.vlan_label(g["vlans"]), g["name"])
            total = len(g["notes_total"])
            tot = "%d %s row(s) in Notes in total" % (total, g["name"])
            if g["ok"]:
                line("OK", "%s: %d active in the scan = %d in Notes." % (lab, len(g["scan"]), len(g["notes"])))
            else:
                line("ERROR", "%s: %d active in the scan, but Notes (%s) lists %d existing port(s) - %s."
                     % (lab, len(g["scan"]), " + ".join(g["sections"]), len(g["notes"]), tot))
            for p in g["scan"]:
                if not g["scan_in_notes"][id(p)]:
                    line("NOTE", "    scan:  %s  (VLAN %s) - not in Notes" % (svc.port_text(p), p["vlan_raw"]))
            for e, how in g["notes_total"]:
                # only discrepancies: connected ports on the right VLAN and new ports aren't listed
                if how in ("matches scan", "no existing switch / port in Notes"):
                    continue
                line("NOTE", "    Notes: %s / %s - %s - %s" % (e["section"], e["component"],
                                                          e.get("old_raw") or e.get("raw") or "", how))

        head("Other active VLANs")
        if not other:
            line("OK", "No active ports on VLANs outside 2, 10-90, 21, 100 and trunk.")
        for v, ps in other.items():
            line("WARN", "VLAN %s: %d active port(s) - not one of the expected VLANs." % (v, len(ps)))
            for p in ps:
                line("NOTE", "    %s%s" % (svc.port_text(p), svc.design_note(p, index)))

        if dld:
            dpath, source, n, issues = dld
            head("Drop List Draft - %s" % os.path.basename(dpath))
            if source:
                line("NOTE", "Read from %s; %d existing switch / port(s) in columns A / B." % (source, n))
            if not issues:
                line("OK", "Every connected port is in Drop List Draft with the same VLAN and description, "
                           "and every non-AP port is patched (P) with a new switch / port.")
            for level, msg, _ in issues:
                line(level, msg)

        self.view.setHtml('<div style="font-family:\'Segoe UI\', sans-serif; font-size:10pt">%s</div>'
                          % "".join(html_parts))
        self.text = "\n".join(text).strip()
        self.btn_copy.setEnabled(True)
        self.setWindowTitle("QA Check - %s" % ("all special VLANs match Notes" if not bad
                                                      else "%d VLAN group(s) differ from Notes" % bad))


def embedded_icon():
    """The DT icon from the embedded PNGs - nothing on disk needed, so it also
    works inside a PyInstaller .exe however it was built."""
    icon = QIcon()
    for size, data in sorted(EMBEDDED_ICON.items()):
        pm = QPixmap()
        if pm.loadFromData(base64.b64decode("".join(data)), "PNG"):
            icon.addPixmap(pm)
    return icon


def load_icon(app):
    """Set the Design Team icon; say where it looked if it can't be loaded.

    The embedded copy is used first; the files in the branding folder are a fallback.

    The .ico is tried first, but Qt reads .ico through a plugin that a PyInstaller
    build doesn't always include; then the icon comes out blank (Windows shows a
    generic one).  availableSizes() is empty in that case, so the PNGs are used."""
    icon = embedded_icon()
    if not icon.isNull():
        app.setWindowIcon(icon)
        return True
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


def splash_text(text):
    """Update the start-up splash of a PyInstaller .exe (no-op when run as .py)."""
    try:
        import pyi_splash
        pyi_splash.update_text(text)
    except Exception:
        pass


def close_splash():
    try:
        import pyi_splash
        pyi_splash.close()
    except Exception:
        pass


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
        close_splash()
        QMessageBox.critical(None, APP_TITLE, "pywin32 and xlwings are needed.\n\n"
                             "Run:  pip install pywin32 xlwings")
        return 1
    splash_text("Opening the window...")
    win = CheckerWindow()
    win.show()
    close_splash()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
