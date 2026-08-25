"""PySide6 GUI：拖拽上传 → 批量转换 → 状态列表 + 报告"""

from __future__ import annotations

import sys
from dataclasses import dataclass, field
from pathlib import Path

from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtGui import QCloseEvent, QDragEnterEvent
from PySide6.QtWidgets import (
    QCheckBox,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QSpinBox,
    QStatusBar,
    QVBoxLayout,
    QWidget,
)

from . import __version__
from .service import (
    FileResult,
    ConvertOptions,
    convert_file,
    collect_files,
    write_report,
)

STATUS_ICON = {"ok": "✅", "warning": "⚠️", "error": "❌", "pending": "⏳"}


class DropArea(QListWidget):
    """可拖拽入文件的列表控件"""

    files_dropped = Signal(list)  # list[Path]

    def dragEnterEvent(self, e: QDragEnterEvent) -> None:
        if e.mimeData().hasUrls():
            e.acceptProposedAction()

    def dragMoveEvent(self, e) -> None:
        if e.mimeData().hasUrls():
            e.acceptProposedAction()

    def dropEvent(self, e) -> None:
        if e.mimeData().hasUrls():
            paths = [Path(u.toLocalFile()) for u in e.mimeData().urls() if u.toLocalFile()]
            if paths:
                self.files_dropped.emit(paths)
            e.acceptProposedAction()


class ConvertWorker(QThread):
    """后台转换线程：逐文件处理，信号回报进度"""

    progress = Signal(int, int, str)     # (done, total, current_name)
    finished_all = Signal(object)        # list[FileResult]

    def __init__(self, files: list[Path], options: ConvertOptions, parent=None):
        super().__init__(parent)
        self.files = files
        self.options = options

    def run(self) -> None:
        results: list[FileResult] = []
        for i, f in enumerate(self.files, 1):
            results.append(convert_file(f, self.options))
            self.progress.emit(i, len(self.files), f.name)
        self.finished_all.emit(results)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"doc2md v{__version__} · 文件批量转 Markdown")
        self.resize(900, 580)
        self.setAcceptDrops(True)
        self._worker: ConvertWorker | None = None
        self._init_ui()

    # ---------- UI ----------
    def _init_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)

        bar = QHBoxLayout()
        self.btn_add = QPushButton("📂 添加文件/文件夹")
        self.btn_add.clicked.connect(self._add_dialog)
        self.btn_clear = QPushButton("清空")
        self.btn_clear.clicked.connect(self._clear)
        bar.addWidget(self.btn_add)
        bar.addWidget(self.btn_clear)
        bar.addStretch()
        bar.addWidget(QLabel("按标题切分(0=不切分):"))
        self.spin_level = QSpinBox()
        self.spin_level.setRange(0, 6)
        self.spin_level.setToolTip("大于0时，按该级别及以上标题把文档拆为多个md，利于RAG分块检索")
        bar.addWidget(self.spin_level)
        self.chk_flat = QCheckBox("全部输出到 第一个文件目录/output/")
        self.chk_flat.setChecked(True)
        bar.addWidget(self.chk_flat)
        self.chk_ocr = QCheckBox("扫描件自动 OCR")
        self.chk_ocr.setChecked(True)
        self.chk_ocr.setToolTip("对识别内容为空的图片和扫描版 PDF 自动运行本地 OCR")
        bar.addWidget(self.chk_ocr)
        layout.addLayout(bar)

        self.list = DropArea()
        self.list.files_dropped.connect(self.add_files)
        self.list.itemSelectionChanged.connect(self._remove_selected)
        layout.addWidget(self.list, 1)

        row = QHBoxLayout()
        self.progress = QProgressBar()
        self.progress.setRange(0, 1000)
        self.progress.setValue(0)
        self.progress.setFormat("就绪：将文件拖入窗口即可添加")
        row.addWidget(self.progress, 1)
        self.btn_start = QPushButton("▶ 开始转换")
        self.btn_start.clicked.connect(self._start)
        row.addWidget(self.btn_start)
        layout.addLayout(row)

        self.statusBar().showMessage(f"输出位置：各文件所在目录的 output/ 子目录 · 转换报告见 output/转换报告.csv")

    def closeEvent(self, e: QCloseEvent) -> None:
        if self._worker is not None and self._worker.isRunning():
            QMessageBox.warning(self, "doc2md", "正在转换，请等待完成或点击取消后再关闭。")
            e.ignore()
            return
        e.accept()

    # ---------- 操作 ----------
    def _add_dialog(self) -> None:
        paths, _ = QFileDialog.getOpenFileNames(
            self, "选择文件（可多选）", "",
            "支持的文件 (*.pdf *.docx *.doc *.pptx *.ppt *.xlsx *.xls *.csv *.html *.htm *.epub *.txt *.md *.png *.jpg *.jpeg *.wav *.mp3);;所有文件 (*)",
        )
        if paths:
            self.add_files([Path(p) for p in paths])

    def add_files(self, paths: list[Path]) -> None:
        files = collect_files(paths)
        existing = {self.list.item(i).data(Qt.UserRole) for i in range(self.list.count())}
        added = 0
        for f in files:
            if str(f) not in existing:
                item = QListWidgetItem(f"⏳ {f.name}")
                item.setToolTip(str(f))
                item.setData(Qt.UserRole, str(f))
                self.list.addItem(item)
                added += 1
        if added:
            self.statusBar().showMessage(f"已添加 {added} 个文件，待转换共 {self.list.count()} 个", 5000)

    def _remove_selected(self) -> None:
        if self._worker is not None and self._worker.isRunning():
            return
        for i in sorted((idx.row() for idx in self.list.selectedIndexes()), reverse=True):
            self.list.takeItem(i)

    def _clear(self) -> None:
        if self._worker is not None and self._worker.isRunning():
            return
        self.list.clear()
        self.progress.setValue(0)
        self.progress.setFormat("就绪：将文件拖入窗口即可添加")

    def _start(self) -> None:
        if self.list.count() == 0:
            QMessageBox.information(self, "提示", "请先拖入或添加要转换的文件")
            return
        files = [Path(self.list.item(i).data(Qt.UserRole)) for i in range(self.list.count())]
        out_dir: Path | None = None
        if self.chk_flat.isChecked():
            out_dir = files[0].parent / "output"
        options = ConvertOptions(
            split_at_heading_level=self.spin_level.value(),
            output_dir=out_dir,
            enable_ocr=self.chk_ocr.isChecked(),
        )
        self._set_running(True)
        for i in range(self.list.count()):
            it = self.list.item(i)
            it.setText(f"⏳ 转换中... {it.text().partition(' ')[2]}")
        self._worker = ConvertWorker(files, options, self)
        self._worker.progress.connect(self._on_progress)
        self._worker.finished_all.connect(self._on_all_done)
        self._worker.start()

    def _set_running(self, running: bool) -> None:
        self.btn_start.setEnabled(not running)
        self.btn_add.setEnabled(not running)
        self.btn_clear.setEnabled(not running)
        self.list.setEnabled(not running)
        self.spin_level.setEnabled(not running)

    def _on_progress(self, done: int, total: int, name: str) -> None:
        self.progress.setValue(int(done * 1000 / total))
        self.progress.setFormat(f"{done}/{total} · {name}")

    def _on_all_done(self, results: list[FileResult]) -> None:
        self._set_running(False)
        self.list.clear()
        for r in results:
            detail = f"{STATUS_ICON[r.status]} {r.src.name}  →  {r.message}  ({r.elapsed_ms}ms)"
            item = QListWidgetItem(detail)
            item.setToolTip(str(r.src))
            item.setData(Qt.UserRole, str(r.src))
            self.list.addItem(item)
        ok = sum(1 for r in results if r.status == "ok")
        warn = sum(1 for r in results if r.status == "warning")
        err = sum(1 for r in results if r.status == "error")
        self.progress.setFormat(f"完成 ✅{ok} ⚠️{warn} ❌{err}")
        # 报告写到第一个文件的 output 目录
        if results:
            report_dir = (self._worker.options.output_dir) or (results[0].src.parent / "output")
            try:
                report_dir.mkdir(parents=True, exist_ok=True)
                write_report(results, report_dir / "转换报告.csv")
                self.statusBar().showMessage(f"全部完成 · 报告: {report_dir / '转换报告.csv'}", 20000)
            except OSError as e:
                self.statusBar().showMessage(f"完成，但报告写入失败: {e}", 10000)


def run_gui() -> None:
    from PySide6.QtWidgets import QApplication

    app = QApplication(sys.argv)
    win = MainWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    run_gui()
