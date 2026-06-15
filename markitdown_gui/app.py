from __future__ import annotations

import os
import sys
from multiprocessing import freeze_support
from pathlib import Path
from typing import Callable, Iterable

from PySide6.QtCore import (
    QEasingCurve,
    QObject,
    QPropertyAnimation,
    QThread,
    QTimer,
    Qt,
    QUrl,
    Signal,
)
from PySide6.QtGui import QColor, QDesktopServices, QFont, QIcon, QPainter, QPainterPath, QPen, QPixmap
from PySide6.QtWidgets import (
    QApplication,
    QAbstractItemView,
    QFrame,
    QFileDialog,
    QGraphicsOpacityEffect,
    QGridLayout,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QSizePolicy,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from .converter import ConversionResult, convert_files
from .excel_rule_engine import (
    TransformationTemplate,
    apply_template,
    infer_template,
    load_template,
    save_template,
)


SUPPORTED_EXTENSIONS = {
    ".csv",
    ".docx",
    ".epub",
    ".gif",
    ".htm",
    ".html",
    ".jpeg",
    ".jpg",
    ".json",
    ".m4a",
    ".md",
    ".mp3",
    ".pdf",
    ".png",
    ".pptx",
    ".rtf",
    ".txt",
    ".wav",
    ".webp",
    ".xls",
    ".xlsx",
    ".xml",
    ".zip",
}


APP_STYLE = """
QWidget {
    background: #f5f8ff;
    color: #1f2937;
    font-family: "Segoe UI", "Inter", Arial, sans-serif;
    font-size: 10.5pt;
}
QFrame#Shell {
    background: #ffffff;
    border: 1px solid #e6f0ff;
    border-radius: 18px;
}
QFrame#Hero {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
        stop:0 #eff6ff, stop:1 #ffffff);
    border: 1px solid #d6e8ff;
    border-radius: 18px;
}
QFrame#DropZone {
    background: #f8fbff;
    border: 2px dashed #91caff;
    border-radius: 16px;
}
QFrame#DropZone[dragging="true"] {
    background: #e6f4ff;
    border-color: #1677ff;
}
QFrame#MetricCard {
    background: #ffffff;
    border: 1px solid #d6e8ff;
    border-radius: 14px;
}
QFrame#FormatterCard {
    background: #ffffff;
    border: 1px solid #d6e8ff;
    border-radius: 14px;
}
QTabWidget::pane {
    border: none;
    background: transparent;
    top: -1px;
}
QTabBar::tab {
    background: transparent;
    border: none;
    border-bottom: 2px solid transparent;
    color: #64748b;
    font-weight: 600;
    padding: 10px 18px;
}
QTabBar::tab:hover {
    color: #1677ff;
}
QTabBar::tab:selected {
    color: #0958d9;
    border-bottom-color: #1677ff;
}
QLabel#Title {
    color: #0f172a;
    font-size: 26pt;
    font-weight: 700;
    font-family: Georgia, Merriweather, "Times New Roman", serif;
}
QLabel#Subtitle {
    color: #64748b;
    font-size: 11pt;
}
QLabel#MetricValue {
    color: #0958d9;
    font-size: 20pt;
    font-weight: 700;
}
QLabel#MetricLabel {
    color: #64748b;
    font-size: 9.5pt;
}
QLabel#DropTitle {
    color: #0958d9;
    font-size: 15pt;
    font-weight: 700;
}
QLabel#DropSubtitle {
    color: #64748b;
}
QPushButton {
    background: #ffffff;
    border: 1px solid #d9d9d9;
    border-radius: 8px;
    color: #1f2937;
    font-weight: 600;
    padding: 9px 14px;
}
QPushButton:hover {
    border-color: #4096ff;
    color: #0958d9;
}
QPushButton:disabled {
    background: #f5f5f5;
    border-color: #d9d9d9;
    color: #bfbfbf;
}
QPushButton#PrimaryButton {
    background: #1677ff;
    border-color: #1677ff;
    color: #ffffff;
}
QPushButton#PrimaryButton:hover {
    background: #0958d9;
    border-color: #0958d9;
}
QLineEdit {
    background: #ffffff;
    border: 1px solid #d9d9d9;
    border-radius: 8px;
    padding: 9px 11px;
}
QLineEdit:focus {
    border-color: #1677ff;
}
QTextEdit {
    background: #ffffff;
    border: 1px solid #d6e8ff;
    border-radius: 10px;
    padding: 10px;
    selection-background-color: #e6f4ff;
    selection-color: #0958d9;
}
QTextEdit:focus {
    border-color: #1677ff;
}
QTableWidget {
    background: #ffffff;
    border: 1px solid #e6f0ff;
    border-radius: 12px;
    gridline-color: #edf4ff;
    selection-background-color: #e6f4ff;
    selection-color: #0958d9;
}
QHeaderView::section {
    background: #f0f7ff;
    border: none;
    border-bottom: 1px solid #d6e8ff;
    color: #475569;
    font-weight: 700;
    padding: 9px;
}
QProgressBar {
    background: #e6f4ff;
    border: none;
    border-radius: 6px;
    height: 12px;
    text-align: center;
}
QProgressBar::chunk {
    background: #1677ff;
    border-radius: 6px;
}
"""


ICON_COLOR = "#1677ff"


def default_output_dir() -> Path:
    documents = Path.home() / "Documents"
    base_dir = documents if documents.is_dir() else Path.home()
    return base_dir / "File Optimizer for AI" / "Markdown"


def permission_error_message(path: Path) -> str:
    return (
        f"File Optimizer for AI cannot write to:\n{path}\n\n"
        "This folder is protected by Windows. Close the program, then "
        "right-click the File Optimizer for AI shortcut and select "
        "\"Run as administrator\".\n\n"
        "Alternatively, choose an output folder inside your Documents folder."
    )


def resource_path(relative_path: str) -> Path:
    base_path = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent.parent))
    return base_path / relative_path


def line_icon(name: str, color: str = ICON_COLOR) -> QIcon:
    pixmap = QPixmap(24, 24)
    pixmap.fill(Qt.transparent)
    painter = QPainter(pixmap)
    painter.setRenderHint(QPainter.Antialiasing)
    pen = QPen(QColor(color), 2.1, Qt.SolidLine, Qt.RoundCap, Qt.RoundJoin)
    painter.setPen(pen)

    path = QPainterPath()
    if name == "file":
        path.moveTo(7, 4)
        path.lineTo(14, 4)
        path.lineTo(19, 9)
        path.lineTo(19, 20)
        path.lineTo(7, 20)
        path.closeSubpath()
        painter.drawPath(path)
        painter.drawLine(14, 4, 14, 9)
        painter.drawLine(14, 9, 19, 9)
    elif name == "folder":
        path.moveTo(3, 7)
        path.lineTo(9, 7)
        path.lineTo(11, 9)
        path.lineTo(21, 9)
        path.lineTo(21, 19)
        path.lineTo(3, 19)
        path.closeSubpath()
        painter.drawPath(path)
    elif name == "remove":
        painter.drawLine(6, 6, 18, 18)
        painter.drawLine(18, 6, 6, 18)
    elif name == "reset":
        path.moveTo(7, 8)
        path.cubicTo(10, 4, 17, 5, 19, 10)
        path.cubicTo(21, 15, 17, 20, 12, 20)
        path.cubicTo(8, 20, 5, 17, 5, 13)
        painter.drawPath(path)
        painter.drawLine(7, 8, 7, 3)
        painter.drawLine(7, 8, 12, 8)
    elif name == "info":
        painter.drawEllipse(4, 4, 16, 16)
        painter.drawLine(12, 11, 12, 17)
        painter.drawPoint(12, 8)
    elif name == "check":
        painter.drawLine(5, 13, 10, 18)
        painter.drawLine(10, 18, 20, 7)
    elif name == "open":
        painter.drawLine(7, 17, 17, 7)
        painter.drawLine(11, 7, 17, 7)
        painter.drawLine(17, 7, 17, 13)
        painter.drawRect(5, 9, 10, 10)
    painter.end()
    return QIcon(pixmap)


class ConversionWorker(QObject):
    status_changed = Signal(str, str)
    finished = Signal(list, str)
    failed = Signal(str)

    def __init__(self, files: list[Path], output_dir: Path) -> None:
        super().__init__()
        self.files = files
        self.output_dir = output_dir

    def run(self) -> None:
        try:
            results = convert_files(
                self.files,
                self.output_dir,
                on_status=lambda path, status: self.status_changed.emit(
                    str(path), status
                ),
            )
            self.finished.emit(results, str(self.output_dir))
        except Exception as exc:
            self.failed.emit(str(exc))


class ExcelTaskWorker(QObject):
    status_changed = Signal(str)
    finished = Signal(object)
    failed = Signal(str)

    def __init__(self, task: Callable[..., object], *args: object) -> None:
        super().__init__()
        self.task = task
        self.args = args

    def run(self) -> None:
        try:
            result = self.task(
                *self.args,
                progress=lambda message: self.status_changed.emit(message),
            )
            self.finished.emit(result)
        except Exception as exc:
            self.failed.emit(str(exc))


class MetricCard(QFrame):
    def __init__(self, label: str, value: str = "0") -> None:
        super().__init__()
        self.setObjectName("MetricCard")
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 14, 16, 14)
        layout.setSpacing(4)

        self.value_label = QLabel(value)
        self.value_label.setObjectName("MetricValue")
        self.label_label = QLabel(label)
        self.label_label.setObjectName("MetricLabel")
        layout.addWidget(self.value_label)
        layout.addWidget(self.label_label)

    def set_value(self, value: str) -> None:
        self.value_label.setText(value)


class DropZone(QFrame):
    clicked = Signal()

    def __init__(self) -> None:
        super().__init__()
        self.setObjectName("DropZone")
        self.setProperty("dragging", False)
        self.setCursor(Qt.PointingHandCursor)
        self.setMinimumHeight(132)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 18, 18, 18)
        layout.setSpacing(5)
        layout.setAlignment(Qt.AlignCenter)

        title = QLabel("Drop files here or click Add files")
        title.setObjectName("DropTitle")
        title.setAlignment(Qt.AlignCenter)
        subtitle = QLabel("PDF | Excel | Word | PowerPoint | CSV | images | audio | ZIP")
        subtitle.setObjectName("DropSubtitle")
        subtitle.setAlignment(Qt.AlignCenter)

        layout.addWidget(title)
        layout.addWidget(subtitle)

    def mousePressEvent(self, event) -> None:
        if event.button() == Qt.LeftButton:
            self.clicked.emit()
        super().mousePressEvent(event)


class ExcelFormatterTab(QWidget):
    EXCEL_FILTER = "Excel workbooks (*.xlsx *.xlsm);;All files (*)"

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.template: TransformationTemplate | None = None
        self.worker_thread: QThread | None = None
        self.worker: ExcelTaskWorker | None = None
        self._build_ui()
        self._update_controls()

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 16, 0, 0)
        layout.setSpacing(16)

        hero = QFrame()
        hero.setObjectName("Hero")
        hero_layout = QVBoxLayout(hero)
        hero_layout.setContentsMargins(22, 18, 22, 18)
        hero_layout.setSpacing(6)
        title = QLabel("Example-Based Excel Formatter")
        title.setObjectName("Title")
        title.setFont(QFont("Georgia", 23, QFont.Bold))
        subtitle = QLabel(
            "Provide an example workbook and the output you expect. The app "
            "infers only transformations it can reproduce exactly, then saves "
            "and reuses those rules without an LLM."
        )
        subtitle.setObjectName("Subtitle")
        subtitle.setWordWrap(True)
        hero_layout.addWidget(title)
        hero_layout.addWidget(subtitle)
        layout.addWidget(hero)

        input_card = QFrame()
        input_card.setObjectName("FormatterCard")
        input_layout = QGridLayout(input_card)
        input_layout.setContentsMargins(18, 16, 18, 16)
        input_layout.setHorizontalSpacing(10)
        input_layout.setVerticalSpacing(10)

        self.example_input = QLineEdit()
        self.example_input.setPlaceholderText("Choose the original example workbook")
        self.example_browse = QPushButton("Browse")
        self.example_browse.setIcon(line_icon("folder"))
        self.example_browse.clicked.connect(
            lambda: self._choose_excel(self.example_input)
        )
        self.expected_output = QLineEdit()
        self.expected_output.setPlaceholderText("Choose the correctly formatted output")
        self.expected_browse = QPushButton("Browse")
        self.expected_browse.setIcon(line_icon("folder"))
        self.expected_browse.clicked.connect(
            lambda: self._choose_excel(self.expected_output)
        )
        self.analyze_button = QPushButton("Analyze Example Pair")
        self.analyze_button.setObjectName("PrimaryButton")
        self.analyze_button.setIcon(line_icon("check", "#ffffff"))
        self.analyze_button.clicked.connect(self._start_analysis)

        input_layout.addWidget(QLabel("Example input"), 0, 0)
        input_layout.addWidget(self.example_input, 0, 1)
        input_layout.addWidget(self.example_browse, 0, 2)
        input_layout.addWidget(QLabel("Expected output"), 1, 0)
        input_layout.addWidget(self.expected_output, 1, 1)
        input_layout.addWidget(self.expected_browse, 1, 2)
        input_layout.addWidget(self.analyze_button, 2, 2)
        input_layout.setColumnStretch(1, 1)
        layout.addWidget(input_card)

        actions_header = QHBoxLayout()
        actions_title = QLabel("Detected actions")
        actions_title.setStyleSheet("font-weight: 700; font-size: 11pt;")
        self.formatter_progress = QProgressBar()
        self.formatter_progress.setRange(0, 0)
        self.formatter_progress.setMaximumWidth(220)
        self.formatter_progress.setVisible(False)
        actions_header.addWidget(actions_title)
        actions_header.addStretch()
        actions_header.addWidget(self.formatter_progress)
        layout.addLayout(actions_header)

        self.actions_text = QTextEdit()
        self.actions_text.setReadOnly(True)
        self.actions_text.setPlaceholderText(
            "Analyze an example pair or load a saved template to see its actions."
        )
        self.actions_text.setMinimumHeight(190)
        layout.addWidget(self.actions_text, 1)

        controls = QHBoxLayout()
        controls.setSpacing(10)
        self.save_template_button = QPushButton("Save as Template")
        self.save_template_button.clicked.connect(self._save_current_template)
        self.load_template_button = QPushButton("Load Template")
        self.load_template_button.clicked.connect(self._load_template_file)
        self.apply_button = QPushButton("Apply to New File")
        self.apply_button.clicked.connect(self._apply_to_new_file)
        self.export_button = QPushButton("Export Output")
        self.export_button.clicked.connect(self._export_example_output)
        controls.addWidget(self.save_template_button)
        controls.addWidget(self.load_template_button)
        controls.addWidget(self.apply_button)
        controls.addWidget(self.export_button)
        controls.addStretch()
        layout.addLayout(controls)

        self.formatter_status = QLabel(
            "Choose an example input workbook and its expected output."
        )
        self.formatter_status.setObjectName("Subtitle")
        self.formatter_status.setWordWrap(True)
        layout.addWidget(self.formatter_status)

    def _choose_excel(self, target: QLineEdit) -> None:
        selected, _ = QFileDialog.getOpenFileName(
            self,
            "Choose an Excel workbook",
            str(Path.home()),
            self.EXCEL_FILTER,
        )
        if selected:
            target.setText(selected)
            self._update_controls()

    def _set_actions(self, actions: list[str]) -> None:
        self.actions_text.setPlainText("\n".join(actions))

    def _set_busy(self, busy: bool) -> None:
        for widget in (
            self.example_input,
            self.example_browse,
            self.expected_output,
            self.expected_browse,
            self.analyze_button,
            self.load_template_button,
        ):
            widget.setEnabled(not busy)
        self.formatter_progress.setVisible(busy)
        self._update_controls(busy)

    def _update_controls(self, busy: bool = False) -> None:
        has_template = self.template is not None
        self.save_template_button.setEnabled(has_template and not busy)
        self.apply_button.setEnabled(has_template and not busy)
        self.export_button.setEnabled(
            has_template and bool(self.example_input.text().strip()) and not busy
        )

    def _start_task(
        self,
        task: Callable[..., object],
        args: tuple[object, ...],
        on_success: Callable[[object], None],
    ) -> None:
        self._set_busy(True)
        self.worker_thread = QThread(self)
        self.worker = ExcelTaskWorker(task, *args)
        self.worker.moveToThread(self.worker_thread)
        self.worker_thread.started.connect(self.worker.run)
        self.worker.status_changed.connect(self.formatter_status.setText)
        self.worker.finished.connect(on_success)
        self.worker.failed.connect(self._task_failed)
        self.worker.finished.connect(self.worker_thread.quit)
        self.worker.failed.connect(self.worker_thread.quit)
        self.worker_thread.finished.connect(self.worker.deleteLater)
        self.worker_thread.finished.connect(self.worker_thread.deleteLater)
        self.worker_thread.finished.connect(self._clear_worker_refs)
        self.worker_thread.start()

    def _start_analysis(self) -> None:
        input_path = self.example_input.text().strip()
        expected_path = self.expected_output.text().strip()
        if not input_path or not expected_path:
            QMessageBox.warning(
                self,
                "Missing files",
                "Choose both an example input and an expected output workbook.",
            )
            return
        self.template = None
        self._set_actions(["Analyzing workbook structure and values..."])
        self.formatter_status.setText("Analyzing example pair...")
        self._start_task(
            infer_template,
            (input_path, expected_path),
            self._analysis_finished,
        )

    def _analysis_finished(self, result: object) -> None:
        self.template = result
        self._set_actions(self.template.actions)
        self._set_busy(False)
        self.formatter_status.setText(
            "Analysis complete. Review the detected actions and warnings."
        )

    def _save_current_template(self) -> None:
        if self.template is None:
            return
        selected, _ = QFileDialog.getSaveFileName(
            self,
            "Save transformation template",
            str(Path.home() / "excel-transformation-template.json"),
            "JSON template (*.json)",
        )
        if not selected:
            return
        try:
            saved = save_template(self.template, selected)
        except Exception as exc:
            QMessageBox.critical(self, "Could not save template", str(exc))
            return
        self.formatter_status.setText(f"Template saved: {saved}")

    def _load_template_file(self) -> None:
        selected, _ = QFileDialog.getOpenFileName(
            self,
            "Load transformation template",
            str(Path.home()),
            "JSON template (*.json);;All files (*)",
        )
        if not selected:
            return
        try:
            self.template = load_template(selected)
        except Exception as exc:
            QMessageBox.critical(self, "Could not load template", str(exc))
            return
        self._set_actions(self.template.actions)
        self._update_controls()
        self.formatter_status.setText(f"Template loaded: {selected}")

    def _apply_to_new_file(self) -> None:
        if self.template is None:
            return
        input_path, _ = QFileDialog.getOpenFileName(
            self,
            "Choose a new input workbook",
            str(Path.home()),
            self.EXCEL_FILTER,
        )
        if not input_path:
            return
        output_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save transformed workbook",
            str(Path(input_path).with_name(f"{Path(input_path).stem} - transformed.xlsx")),
            "Excel workbook (*.xlsx)",
        )
        if output_path:
            self._start_apply(input_path, output_path)

    def _export_example_output(self) -> None:
        if self.template is None:
            return
        input_path = self.example_input.text().strip()
        if not input_path:
            QMessageBox.warning(
                self,
                "Missing input",
                "Choose the example input workbook before exporting.",
            )
            return
        output_path, _ = QFileDialog.getSaveFileName(
            self,
            "Export reproduced output",
            str(Path(input_path).with_name(f"{Path(input_path).stem} - reproduced.xlsx")),
            "Excel workbook (*.xlsx)",
        )
        if output_path:
            self._start_apply(input_path, output_path)

    def _start_apply(self, input_path: str, output_path: str) -> None:
        self.formatter_status.setText("Applying transformation template...")
        self._start_task(
            apply_template,
            (self.template, input_path, output_path),
            self._apply_finished,
        )

    def _apply_finished(self, result: object) -> None:
        self._set_busy(False)
        self.formatter_status.setText(f"Output saved: {result}")
        if QMessageBox.question(
            self,
            "Transformation complete",
            f"Output saved to:\n{result}\n\nOpen the containing folder?",
        ) == QMessageBox.Yes:
            QDesktopServices.openUrl(
                QUrl.fromLocalFile(str(Path(result).resolve().parent))
            )

    def _task_failed(self, message: str) -> None:
        self._set_busy(False)
        self.formatter_status.setText("Operation failed.")
        QMessageBox.critical(self, "Could not complete operation", message)

    def _clear_worker_refs(self) -> None:
        self.worker = None
        self.worker_thread = None


class MarkItDownWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("File Optimizer for AI")
        self.resize(1120, 760)
        self.setMinimumSize(880, 620)
        self.setAcceptDrops(True)

        self.files: list[Path] = []
        self.rows: dict[Path, int] = {}
        self.output_dir = default_output_dir()
        self.worker_thread: QThread | None = None
        self.worker: ConversionWorker | None = None
        self.worker_count = os.cpu_count() or 1

        self._build_ui()
        self._update_metrics()
        QTimer.singleShot(0, self._animate_window_in)

    def _build_ui(self) -> None:
        root = QWidget()
        root_layout = QVBoxLayout(root)
        root_layout.setContentsMargins(22, 22, 22, 22)

        self.shell = QFrame()
        self.shell.setObjectName("Shell")
        shell_layout = QVBoxLayout(self.shell)
        shell_layout.setContentsMargins(24, 24, 24, 24)
        shell_layout.setSpacing(18)
        root_layout.addWidget(self.shell)
        self.setCentralWidget(root)

        self.tabs = QTabWidget()
        self.tabs.setDocumentMode(True)
        shell_layout.addWidget(self.tabs)

        converter_page = QWidget()
        converter_layout = QVBoxLayout(converter_page)
        converter_layout.setContentsMargins(0, 16, 0, 0)
        converter_layout.setSpacing(18)
        self.tabs.addTab(converter_page, "Markdown Optimizer")

        self.excel_formatter = ExcelFormatterTab()
        self.tabs.addTab(self.excel_formatter, "Excel Formatter")

        hero = QFrame()
        hero.setObjectName("Hero")
        hero_layout = QHBoxLayout(hero)
        hero_layout.setContentsMargins(22, 20, 22, 20)
        hero_layout.setSpacing(18)

        title_block = QVBoxLayout()
        title = QLabel("File Optimizer for AI")
        title.setObjectName("Title")
        title.setFont(QFont("Georgia", 26, QFont.Bold))
        subtitle = QLabel(
            "Put your PDF, Excel, Word, or slides. Get a clean Markdown file "
            "that uses fewer tokens and works better with ChatGPT, Claude, Gemini."
        )
        subtitle.setObjectName("Subtitle")
        subtitle.setWordWrap(True)
        title_block.addWidget(title)
        title_block.addWidget(subtitle)
        title_block.addStretch()
        hero_layout.addLayout(title_block, 1)

        self.files_card = MetricCard("Files queued")
        self.converted_card = MetricCard("Converted")
        self.tokens_card = MetricCard("Estimated token reduction")
        self.cores_card = MetricCard("CPU cores available", str(self.worker_count))
        hero_layout.addWidget(self.files_card)
        hero_layout.addWidget(self.converted_card)
        hero_layout.addWidget(self.tokens_card)
        hero_layout.addWidget(self.cores_card)
        converter_layout.addWidget(hero)

        self.drop_zone = DropZone()
        self.drop_zone.clicked.connect(self._add_files)
        converter_layout.addWidget(self.drop_zone)

        toolbar = QHBoxLayout()
        toolbar.setSpacing(10)
        self.add_files_button = QPushButton("Add files")
        self.add_files_button.setIcon(line_icon("file"))
        self.add_files_button.clicked.connect(self._add_files)
        self.add_folder_button = QPushButton("Add folder")
        self.add_folder_button.setIcon(line_icon("folder"))
        self.add_folder_button.clicked.connect(self._add_folder)
        self.remove_button = QPushButton("Remove selected")
        self.remove_button.setIcon(line_icon("remove"))
        self.remove_button.clicked.connect(self._remove_selected)
        self.clear_button = QPushButton("Clear all")
        self.clear_button.setIcon(line_icon("reset"))
        self.clear_button.clicked.connect(self._clear_files)
        toolbar.addWidget(self.add_files_button)
        toolbar.addWidget(self.add_folder_button)
        toolbar.addWidget(self.remove_button)
        toolbar.addWidget(self.clear_button)
        toolbar.addStretch()
        converter_layout.addLayout(toolbar)

        self.table = QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(["File", "Type", "Location", "Status"])
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.verticalHeader().setVisible(False)
        self.table.setAlternatingRowColors(False)
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.Stretch)
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        converter_layout.addWidget(self.table, 1)

        output_grid = QGridLayout()
        output_grid.setHorizontalSpacing(10)
        output_grid.setVerticalSpacing(8)
        output_label = QLabel("Output folder")
        self.output_input = QLineEdit(str(self.output_dir))
        self.output_input.textChanged.connect(self._set_output_dir)
        self.browse_button = QPushButton("Browse")
        self.browse_button.setIcon(line_icon("folder"))
        self.browse_button.clicked.connect(self._choose_output)
        self.open_output_button = QPushButton("Open folder")
        self.open_output_button.setIcon(line_icon("open"))
        self.open_output_button.clicked.connect(self._open_output_folder)
        output_grid.addWidget(output_label, 0, 0)
        output_grid.addWidget(self.output_input, 1, 0)
        output_grid.addWidget(self.browse_button, 1, 1)
        output_grid.addWidget(self.open_output_button, 1, 2)
        converter_layout.addLayout(output_grid)

        footer = QHBoxLayout()
        footer.setSpacing(12)
        self.info_button = QPushButton("Info")
        self.info_button.setIcon(line_icon("info"))
        self.info_button.clicked.connect(self._show_info)
        self.summary_label = QLabel("Add files to begin.")
        self.summary_label.setObjectName("Subtitle")
        self.progress = QProgressBar()
        self.progress.setRange(0, 1)
        self.progress.setValue(0)
        self.progress.setVisible(False)
        self.convert_button = QPushButton("Convert to Markdown")
        self.convert_button.setObjectName("PrimaryButton")
        self.convert_button.setIcon(line_icon("check", "#ffffff"))
        self.convert_button.clicked.connect(self._start_conversion)
        footer.addWidget(self.info_button)
        footer.addWidget(self.summary_label, 1)
        footer.addWidget(self.progress, 1)
        footer.addWidget(self.convert_button)
        converter_layout.addLayout(footer)

    def _show_info(self) -> None:
        box = QMessageBox(self)
        box.setWindowTitle("How File Optimizer for AI works")
        box.setIcon(QMessageBox.Information)
        box.setTextFormat(Qt.RichText)
        box.setText(
            "<b>How it works</b><br>"
            "This app uses Microsoft MarkItDown to convert supported files into "
            "Markdown. Markdown is usually cleaner and lighter than raw PDF, "
            "Excel, Word, or slide files, so it can use fewer estimated tokens "
            "when pasted into AI tools.<br><br>"
            "<b>Limitations</b><br>"
            "Token savings are estimates, complex layouts may not convert "
            "perfectly, scanned PDFs may need OCR, and audio conversion can "
            "require FFmpeg.<br><br>"
            "<span style='color:#64748b;'>Made with love by "
            "Tareq Rahman Pantho.</span>"
        )
        box.setStandardButtons(QMessageBox.Ok)
        box.exec()

    def _add_files(self) -> None:
        patterns = " ".join(f"*{extension}" for extension in sorted(SUPPORTED_EXTENSIONS))
        selected, _ = QFileDialog.getOpenFileNames(
            self,
            "Select files to convert",
            str(Path.home()),
            f"Supported files ({patterns});;All files (*)",
        )
        self._queue_files(Path(item) for item in selected)

    def _add_folder(self) -> None:
        selected = QFileDialog.getExistingDirectory(
            self,
            "Select a folder",
            str(Path.home()),
        )
        if not selected:
            return
        folder = Path(selected)
        files = sorted(
            item
            for item in folder.iterdir()
            if item.is_file() and item.suffix.lower() in SUPPORTED_EXTENSIONS
        )
        self._queue_files(files)
        if not files:
            QMessageBox.information(
                self,
                "No supported files",
                "The selected folder contains no supported files.",
            )

    def _queue_files(self, paths: Iterable[Path]) -> None:
        for path in paths:
            resolved = path.resolve()
            if not resolved.is_file() or resolved in self.rows:
                continue

            row = self.table.rowCount()
            self.table.insertRow(row)
            self.files.append(resolved)
            self.rows[resolved] = row
            self._set_row(row, resolved, "Ready")
        self._update_row_map()
        self._update_metrics()

    def _set_row(self, row: int, path: Path, status: str) -> None:
        values = [
            path.name,
            path.suffix.lstrip(".").upper() or "FILE",
            str(path.parent),
            status,
        ]
        for column, value in enumerate(values):
            item = QTableWidgetItem(value)
            if column in (1, 3):
                item.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(row, column, item)

    def _update_row_map(self) -> None:
        self.rows = {path: index for index, path in enumerate(self.files)}

    def _remove_selected(self) -> None:
        selected_rows = sorted(
            {index.row() for index in self.table.selectedIndexes()},
            reverse=True,
        )
        for row in selected_rows:
            self.table.removeRow(row)
            del self.files[row]
        self._update_row_map()
        self._update_metrics()

    def _clear_files(self) -> None:
        self.files.clear()
        self.rows.clear()
        self.table.setRowCount(0)
        self._update_metrics()

    def _set_output_dir(self, value: str) -> None:
        if value.strip():
            self.output_dir = Path(value.strip()).expanduser()

    def _choose_output(self) -> None:
        selected = QFileDialog.getExistingDirectory(
            self,
            "Choose output folder",
            str(self.output_dir),
        )
        if selected:
            self.output_dir = Path(selected)
            self.output_input.setText(str(self.output_dir))

    def _open_output_folder(self) -> None:
        try:
            self.output_dir.mkdir(parents=True, exist_ok=True)
        except PermissionError:
            self._show_permission_error(self.output_dir)
            return
        except OSError as exc:
            QMessageBox.critical(
                self,
                "Cannot open output folder",
                f"The output folder could not be created:\n{exc}",
            )
            return
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(self.output_dir.resolve())))

    def _animate_window_in(self) -> None:
        self.shell_effect = QGraphicsOpacityEffect(self.shell)
        self.shell.setGraphicsEffect(self.shell_effect)
        self.shell_animation = QPropertyAnimation(
            self.shell_effect, b"opacity", self
        )
        self.shell_animation.setDuration(320)
        self.shell_animation.setStartValue(0.0)
        self.shell_animation.setEndValue(1.0)
        self.shell_animation.setEasingCurve(QEasingCurve.OutCubic)
        self.shell_animation.start()

    def _show_permission_error(self, path: Path) -> None:
        QMessageBox.critical(
            self,
            "Administrator permission required",
            permission_error_message(path),
        )

    def _prepare_output_dir(self) -> bool:
        try:
            self.output_dir.mkdir(parents=True, exist_ok=True)
            probe = self.output_dir / ".file_optimizer_write_test"
            probe.touch(exist_ok=True)
            probe.unlink()
            return True
        except PermissionError:
            self._show_permission_error(self.output_dir)
        except OSError as exc:
            if getattr(exc, "winerror", None) == 5:
                self._show_permission_error(self.output_dir)
            else:
                QMessageBox.critical(
                    self,
                    "Output folder error",
                    f"The output folder is not available:\n{exc}",
                )
        return False

    def _update_metrics(
        self,
        converted: int = 0,
        token_savings_percent: float = 0.0,
        summary: str | None = None,
    ) -> None:
        queued = len(self.files)
        self.files_card.set_value(str(queued))
        self.converted_card.set_value(str(converted))
        self.tokens_card.set_value(f"{token_savings_percent:.1f}%")
        if summary is not None:
            self.summary_label.setText(summary)
        elif queued:
            self.summary_label.setText(f"{queued} file{'s' if queued != 1 else ''} ready.")
        else:
            self.summary_label.setText("Add files to begin.")

    def _set_controls_enabled(self, enabled: bool) -> None:
        for widget in (
            self.add_files_button,
            self.add_folder_button,
            self.remove_button,
            self.clear_button,
            self.output_input,
            self.browse_button,
            self.open_output_button,
            self.convert_button,
            self.info_button,
            self.drop_zone,
        ):
            widget.setEnabled(enabled)
        self.convert_button.setText(
            "Convert to Markdown" if enabled else "Converting..."
        )
        self.progress.setVisible(not enabled)
        if enabled:
            self.progress.setRange(0, 1)
            self.progress.setValue(0)
        else:
            self.progress.setRange(0, 0)

    def _start_conversion(self) -> None:
        if not self.files:
            QMessageBox.information(self, "No files", "Add at least one file.")
            return
        if not self.output_input.text().strip():
            QMessageBox.warning(self, "Output folder", "Choose an output folder.")
            return
        if not self._prepare_output_dir():
            return

        for path in self.files:
            self._set_status(path, "Waiting")
        self._set_controls_enabled(False)
        self._update_metrics(summary=f"Converting {len(self.files)} files...")

        self.worker_thread = QThread(self)
        self.worker = ConversionWorker(list(self.files), self.output_dir)
        self.worker.moveToThread(self.worker_thread)
        self.worker_thread.started.connect(self.worker.run)
        self.worker.status_changed.connect(self._set_status_from_signal)
        self.worker.finished.connect(self._finish_conversion)
        self.worker.failed.connect(self._conversion_failed)
        self.worker.finished.connect(self.worker_thread.quit)
        self.worker.failed.connect(self.worker_thread.quit)
        self.worker_thread.finished.connect(self.worker.deleteLater)
        self.worker_thread.finished.connect(self.worker_thread.deleteLater)
        self.worker_thread.finished.connect(self._clear_worker_refs)
        self.worker_thread.start()

    def _set_status_from_signal(self, path_text: str, status: str) -> None:
        self._set_status(Path(path_text), status)

    def _set_status(self, path: Path, status: str) -> None:
        row = self.rows.get(path)
        if row is None:
            return
        item = self.table.item(row, 3)
        if item is None:
            item = QTableWidgetItem()
            item.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(row, 3, item)
        item.setText(status)

    def _finish_conversion(self, results: list[ConversionResult], output_dir: str) -> None:
        self._set_controls_enabled(True)
        successful = sum(result.succeeded for result in results)
        failed = len(results) - successful
        source_tokens = sum(result.source_tokens_estimate for result in results)
        markdown_tokens = sum(result.markdown_tokens_estimate for result in results)
        token_savings_percent = (
            max(0, source_tokens - markdown_tokens) / source_tokens * 100
            if source_tokens
            else 0.0
        )
        self._update_metrics(
            converted=successful,
            token_savings_percent=token_savings_percent,
            summary=(
                f"Finished: {successful} converted, {failed} failed. "
                f"Estimated token reduction: {token_savings_percent:.1f}%."
            ),
        )

        if failed:
            details = "\n".join(
                f"{result.source.name}: {result.error}"
                for result in results
                if not result.succeeded
            )
            QMessageBox.warning(
                self,
                "Conversion complete",
                f"Converted {successful} of {len(results)} files.\n\n{details[:2000]}",
            )
            return

        if QMessageBox.question(
            self,
            "Conversion complete",
                (
                    f"Converted {successful} files successfully.\n\n"
                    f"Estimated token reduction: {token_savings_percent:.1f}%.\n\n"
                    "Open the output folder?"
                ),
        ) == QMessageBox.Yes:
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(Path(output_dir).resolve())))

    def _conversion_failed(self, message: str) -> None:
        self._set_controls_enabled(True)
        self._update_metrics(summary="Conversion could not start.")
        QMessageBox.critical(self, "Conversion error", message)

    def _clear_worker_refs(self) -> None:
        self.worker = None
        self.worker_thread = None

    def dragEnterEvent(self, event) -> None:
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            self.drop_zone.setProperty("dragging", True)
            self.drop_zone.style().unpolish(self.drop_zone)
            self.drop_zone.style().polish(self.drop_zone)

    def dragLeaveEvent(self, event) -> None:
        self.drop_zone.setProperty("dragging", False)
        self.drop_zone.style().unpolish(self.drop_zone)
        self.drop_zone.style().polish(self.drop_zone)
        super().dragLeaveEvent(event)

    def dropEvent(self, event) -> None:
        paths: list[Path] = []
        for url in event.mimeData().urls():
            path = Path(url.toLocalFile())
            if path.is_dir():
                paths.extend(
                    item
                    for item in sorted(path.iterdir())
                    if item.is_file() and item.suffix.lower() in SUPPORTED_EXTENSIONS
                )
            elif path.is_file():
                paths.append(path)
        self._queue_files(paths)
        self.drop_zone.setProperty("dragging", False)
        self.drop_zone.style().unpolish(self.drop_zone)
        self.drop_zone.style().polish(self.drop_zone)


def main() -> None:
    freeze_support()
    app = QApplication(sys.argv)
    app.setStyleSheet(APP_STYLE)
    icon_path = resource_path("assets/app.ico")
    if icon_path.exists():
        app.setWindowIcon(QIcon(str(icon_path)))
    window = MarkItDownWindow()
    window.show()
    sys.exit(app.exec())
