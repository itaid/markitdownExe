"""OCR 封装层：可插拔后端，默认 RapidOCR（纯 CPU + onnxruntime，体积最小、打包最稳）

选型依据（2026-08 实测，中文公文样例）：
- rapidocr_onnxruntime: 字迹准确率最高（含表格单元格），单页 ~3s，
  运行时体积 onnxruntime ~73MB，无重依赖 → 默认
- easyocr: 依赖 torch（~556MB），识别句块切碎、标点丢失 → 备选
- paddleocr 3.x: M 系芯片撞 PIR 编译器 bug 无法初始化 → 暂不接入

用法：
    from .ocr import OCR_AVAILABLE, ocr_image
    if OCR_AVAILABLE:
        results = ocr_image("scan.png")   # [(text, score), ...]

后端切换：环境变量 DOC2MD_OCR=rapid|paddle|easy

打包注意：EXE 用 PyInstaller 打包时必须收集 onnxruntime 的 .dll/.so
（spec 中 collect_all("onnxruntime")，已配置）。
"""

from __future__ import annotations

import os
import threading
from dataclasses import dataclass
from pathlib import Path

_LOCK = threading.Lock()
_backend = None  # 懒加载单例
_backend_name: str | None = None


@dataclass
class OCRLine:
    text: str
    score: float
    # 文本框四个角点 [[x, y] x4]，可用于排序/版式判断
    box: list[list[float]]


class OCREngine:
    """统一接口：各后端实现 ocr_image 即可接入"""

    name = "base"

    def ocr_image(self, image_path: str | Path) -> list[OCRLine]:
        raise NotImplementedError


class RapidOCREngine(OCREngine):
    name = "rapid"

    def __init__(self):
        from rapidocr_onnxruntime import RapidOCR

        self._impl = RapidOCR()

    def ocr_image(self, image_path: str | Path) -> list[OCRLine]:
        result, _elapse = self._impl(str(image_path))
        if not result:
            return []
        out = []
        for item in result:
            box, text, score = item
            out.append(OCRLine(text=text, score=float(score), box=[[float(x), float(y)] for x, y in box]))
        return sorted(out, key=lambda l: (l.box[0][1], l.box[0][0]))


class PaddleOCREngine(OCREngine):
    """备选后端：接口存在但不默认启用（macOS M 系当前版本有初始化 bug）"""

    name = "paddle"

    def __init__(self):
        from paddleocr import PaddleOCR

        self._impl = PaddleOCR(lang="ch")

    def ocr_image(self, image_path: str | Path) -> list[OCRLine]:
        result = self._impl.predict(str(image_path))
        out = []
        for res in result:
            for text, score in zip(res.get("rec_texts", []), res.get("rec_scores", [])):
                out.append(OCRLine(text=text, score=float(score), box=[[0, 0]] * 4))
        return out


class EasyOCREngine(OCREngine):
    """备选后端：准确率低、体积大（torch），仅在前两者都不可用时兜底"""

    name = "easy"

    def __init__(self):
        import easyocr

        self._impl = easyocr.Reader(["ch_sim", "en"], gpu=False, verbose=False)

    def ocr_image(self, image_path: str | Path) -> list[OCRLine]:
        result = self._impl.readtext(str(image_path))
        out = []
        for box, text, score in result:
            out.append(OCRLine(text=text, score=float(score), box=[[float(x), float(y)] for x, y in box]))
        return out


_ENGINES: dict[str, type[OCREngine]] = {
    "rapid": RapidOCREngine,
    "paddle": PaddleOCREngine,
    "easy": EasyOCREngine,
}


def get_engine() -> OCREngine:
    """按 DOC2MD_OCR 环境变量（默认 rapid）懒加载单例，加载失败抛异常"""
    global _backend, _backend_name
    with _LOCK:
        want = os.environ.get("DOC2MD_OCR", "rapid").lower()
        if _backend is not None and _backend_name == want:
            return _backend
        _backend = _ENGINES.get(want, RapidOCREngine)()
        _backend_name = want
        return _backend


def ocr_available() -> bool:
    """默认后端是否可用（用于 UI/CLI 显示能力提示，避免逐个探测开销）"""
    try:
        import rapidocr_onnxruntime  # noqa: F401

        return True
    except ImportError:
        return False


def ocr_image(image_path: str | Path) -> list[OCRLine]:
    """对单张图片做 OCR，按阅读顺序（先 y 后 x）返回文本行"""
    return get_engine().ocr_image(image_path)


def ocr_pdf(pdf_path: str | Path, ocr: bool = True) -> list[dict]:
    """把 PDF 每页渲染为图片并 OCR。返回 [{page, lines}, ...]

    需要 pypdfium2（渲染）+ OCR 后端。用于扫描版 PDF 兜底。
    """
    import pypdfium2 as pdfium

    pdf = pdfium.PdfDocument(str(pdf_path))
    pages = []
    try:
        for i, page in enumerate(pdf, 1):
            # 2x 缩放（144 dpi）在速度与识别率之间平衡
            bitmap = page.render(scale=2)
            pil_image = bitmap.to_pil()
            tmp = Path(pdf_path).with_suffix(f".ocr_page_{i}.png")
            pil_image.save(tmp)
            try:
                lines = ocr_image(tmp) if ocr else []
            finally:
                tmp.unlink(missing_ok=True)
            pages.append({"page": i, "lines": lines})
    finally:
        pdf.close()
    return pages


def lines_to_markdown(lines: list[OCRLine]) -> str:
    parts: dict[int, list[OCRLine]] = {}
    for l in lines:
        band = round(l.box[0][1] / 15)
        parts.setdefault(band, []).append(l)
    md_lines = []
    for band in sorted(parts):
        row = sorted(parts[band], key=lambda l: l.box[0][0])
        md_lines.append(" ".join(l.text for l in row))
    return "\n".join(md_lines)
