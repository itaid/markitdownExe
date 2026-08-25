"""转换核心：调用 markitdown + 后处理（文件名规范化 / 按章节分块 / 校验）

本模块无 GUI 依赖，可独立测试，也是在 Windows EXE 里执行的部分。
"""

from __future__ import annotations

import re as _re
import re
import time
from dataclasses import dataclass, field
from pathlib import Path

from markitdown import MarkItDown

# 转译后支持的扩展名 → 友好名称
SUPPORTED_EXTENSIONS = [
    ".pdf", ".docx", ".doc", ".pptx", ".ppt",
    ".xlsx", ".xls", ".csv", ".html", ".htm",
    ".epub", ".msg", ".xml", ".json", ".ipynb",
    ".txt", ".md", ".rst", ".zip",
    ".png", ".jpg", ".jpeg", ".webp", ".bmp", ".tiff",
    ".wav", ".mp3", ".m4a",
]

# 输出 md 内容低于该字数时提示"内容过少"
WARN_MIN_CHARS = 50

# 需人工确认的文件（扫描件/图片/音频）
NEEDS_REVIEW_EXT = {".png", ".jpg", ".jpeg", ".webp", ".bmp", ".tiff", ".wav", ".mp3", ".m4a"}
IMAGE_OR_SCAN_EXT = {".png", ".jpg", ".jpeg", ".webp", ".bmp", ".tiff"}

# 二进制容器格式的魔数校验：扩展名与内容不符时提前报错，避免被当纯文本转出一堆乱码
_MAGIC_CHECKS = [
    ((".docx", ".xlsx", ".pptx", ".epub"), b"PK\x03\x04", "OFFICE/EPUB 文件（ZIP 容器）"),
    ((".doc", ".xls", ".ppt"), b"\xd0\xcf\x11\xe0", "旧版 Office 文件（OLE2 容器）"),
    ((".pdf",), b"%PDF", "PDF"),
    ((".zip",), b"PK", "ZIP"),
]


def _magic_check(path):
    """扩展名要求特定容器格式时校验文件头；通过返回 None"""
    suffix = path.suffix.lower()
    try:
        with path.open("rb") as fh:
            head = fh.read(8)
    except OSError as e:
        return "无法读取文件: %s" % e
    if len(head) < 4:
        return "文件过小，可能已损坏"
    for exts, sig, label in _MAGIC_CHECKS:
        if suffix in exts and not head.startswith(sig):
            return ("文件内容不是有效的 %s（扩展名 %s 与内容不符，可能已损坏或被重命名）"
                    % (label, suffix))
    return None


@dataclass
class FileResult:
    """单个文件的转换结果"""
    src: Path
    status: str = "pending"        # pending / ok / warning / error
    dst: list[Path] = field(default_factory=list)
    message: str = ""
    char_count: int = 0
    elapsed_ms: int = 0
    ocr_used: bool = False


@dataclass
class ConvertOptions:
    """转换选项"""
    # 按标题层级切分：文档中出现该级别及以上标题时，切分为单独 md 文件
    # 0 = 不切分，输出单个 md
    split_at_heading_level: int = 0
    # 切分输出目录结构: {out_dir}/{stem}/01-章节名.md
    output_dir: Path | None = None   # None = 原目录下的 output/
    # 扫描件/图片 OCR 兜底（默认开启；无 OCR 依赖时自动降级为警告）
    enable_ocr: bool = True


def normalize_filename(stem: str, max_len: int = 80) -> str:
    """规范化文件名：去非法字符、压缩空格、去"最终版/副本/v1/（2)"类尾部噪音"""
    # Windows/通用非法字符
    s = re.sub(r'[\\/:*?"<>|\r\n\t]', "_", stem)
    # " - "/"--" 这类连接符替换为 _；压缩连续空格
    s = re.sub(r"\s*[-–—]\s*", "_", s)
    s = re.sub(r"[\s_]{2,}", "_", s).replace(" _", "_").strip("_ ")
    # 多轮移除尾部噪音。保留独立尾部的 "(2)"（编号有业务含义），
    # 仅当数字括号与 final/最终版/副本/copy/vN 等噪音词相邻时一并去除
    noise_word = re.compile(r"[_\-\s]?(?:final|最终版|副本|copy|v?\d+)$")
    paren_near_noise = re.compile(
        r"[_\-\s]*[（(]\d+[)）](?=[_\-\s]*(?:final|最终版|副本|copy|v?\d+)$)")
    for _ in range(5):
        new1 = paren_near_noise.sub("", s)
        new2 = noise_word.sub("", new1).rstrip(" ._\-")
        if not new2:      # 删光整个名字（如文件就叫 copy）则保留原名
            break
        if new2 == s:
            break
        s = new2
    return s[:max_len] or "unnamed"


def split_by_headings(md: str, level: int) -> list[tuple[str, str]]:
    """按 Markdown 标题切分。

    返回 [(section_title, section_md), ...]，其中第一个元素
    标题为 "__preamble__"（首个符合条件标题前的内容）。
    """
    if level <= 0:
        return [("", md)]

    # 匹配 <= level 的标题行，如 "## " 且 level>=2
    pattern = re.compile(rf"^(#{{{1},{level}}})\s+(.+)$", re.M)
    parts: list[tuple[str, str]] = []
    matches = list(pattern.finditer(md))
    if not matches:
        return [("", md)]

    head = md[: matches[0].start()].strip()
    if head:
        parts.append(("__preamble__", head))

    for i, m in enumerate(matches):
        title = m.group(2).strip()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(md)
        parts.append((title, f"{m.group(0)}\n\n{md[m.end():end]}".strip()))
    return parts


def _sanitize_section_title(title: str, max_len: int = 60) -> str:
    s = re.sub(r"[\\/:*?\"<>|\r\n\t]", "_", title).strip()
    s = re.sub(r"\s*[-–—]\s+", "_", s)
    s = re.sub(r"[\s_]{2,}", "_", s).strip("_ ")
    return s[:max_len] or "正文"


def _clean_md(text: str) -> str:
    """清理转换器残留：HTML 注释、多余空行、行尾空白"""
    text = _re.sub(r"<!--.*?-->", "", text, flags=_re.S)
    text = _re.sub(r"[ \t]+$", "", text, flags=_re.M)
    text = _re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def _has_meaningful_text(md_body: str) -> bool:
    """判断章节内容是否有实际文本（过滤纯 HTML 注释、分隔线等残留）"""
    import re as _re
    t = _re.sub(r"<!--.*?-->", "", md_body, flags=_re.S)   # 去注释
    t = _re.sub(r"^\s*[#*\-|>\s]+\s*$", "", t, flags=_re.M)  # 去引导符行
    return bool(t.strip())


def _run_ocr_for_file(src: Path) -> str | None:
    """对扫描件做 OCR，返回 Markdown 文本；无可用 OCR 引擎时返回 None"""
    from . import ocr as ocr_mod

    if not ocr_mod.ocr_available():
        try:
            ocr_mod.get_engine()  # 触发真实导入检测
        except Exception:  # noqa: BLE001
            return None
    if src.suffix.lower() == ".pdf":
        pages = ocr_mod.ocr_pdf(src)
        parts = [f"## 第 {p['page']} 页\n\n" + ocr_mod.lines_to_markdown(p["lines"]) for p in pages]
        return "\n\n".join(parts).strip() or None
    lines = ocr_mod.ocr_image(src)
    return ocr_mod.lines_to_markdown(lines).strip() or None


def convert_file(
    src: Path,
    options: ConvertOptions,
) -> FileResult:
    """转换单个文件。不抛异常，错误信息写入 FileResult.message。"""
    result = FileResult(src=src)
    start = time.time()

    err = _magic_check(src)
    if err:
        result.status = "error"
        result.message = err
        result.elapsed_ms = int((time.time() - start) * 1000)
        return result

    md = MarkItDown()

    try:
        res = md.convert_local(src)  # 安全：仅本地文件，不走网络
        text = res.text_content or ""
        result.char_count = len(text)
    except Exception as e:  # noqa: BLE001 - 需要把错误信息带给 GUI
        base_err = f"转换失败: {type(e).__name__}: {e}"
        # 可 OCR 类型 → 常规解析失败时先试 OCR（扫描件常见路径）
        ocr_text = None
        if options.enable_ocr and src.suffix.lower() in IMAGE_OR_SCAN_EXT:
            try:
                ocr_text = _run_ocr_for_file(src)
            except Exception as oe:  # noqa: BLE001
                base_err = f"{base_err}；OCR 兜底也失败: {type(oe).__name__}: {oe}"
        if ocr_text:
            text = "# OCR 识别内容\n\n" + ocr_text
            result.char_count = len(text)
            result.ocr_used = True
        else:
            result.status = "error"
            result.message = base_err
            result.elapsed_ms = int((time.time() - start) * 1000)
            return result

    # 扫描件/图片/扫描版 PDF → OCR 兜底（解析出的正文过短时触发）
    ocr_targets = IMAGE_OR_SCAN_EXT | {".pdf"}
    if options.enable_ocr and src.suffix.lower() in ocr_targets:
        content_len = len(re.sub(r"\s+", "", text))
        if content_len < 200:
            try:
                ocr_text = _run_ocr_for_file(src)
            except Exception as e:  # noqa: BLE001
                if not text.strip():
                    result.status = "error"
                    result.message = f"OCR 失败: {type(e).__name__}: {e}"
                    result.elapsed_ms = int((time.time() - start) * 1000)
                    return result
                ocr_text = None  # 有少量解析内容时降级为普通输出
            if ocr_text:
                prefix = _clean_md(text)
                text = (prefix + "\n\n" if prefix else "") + "# OCR 识别内容\n\n" + ocr_text
                result.char_count = len(text)
                result.ocr_used = True
            elif not text.strip():
                result.status = "warning"
                result.message = "文字层为空且 OCR 未识别到文字（可能为非文字内容）"
                result.elapsed_ms = int((time.time() - start) * 1000)
                return result

    # 空/短内容校验
    if not text.strip():
        result.status = "error"
        result.message = "转换结果为空（可能是扫描件、纯图片 PDF 或加密文件）"
        result.elapsed_ms = int((time.time() - start) * 1000)
        return result

    # 输出目录
    out_dir = options.output_dir or (src.parent / "output")
    out_dir.mkdir(parents=True, exist_ok=True)
    stem = normalize_filename(src.stem)

    if (options.split_at_heading_level > 0 and result.src.suffix.lower() not in NEEDS_REVIEW_EXT
            and not result.ocr_used):
        # 按章节切分为多个 md
        sec_dir = out_dir / stem
        sec_dir.mkdir(parents=True, exist_ok=True)
        sections = split_by_headings(text, options.split_at_heading_level)
        sections = [(t, b) for t, b in sections if _has_meaningful_text(b)]
        written: list[Path] = []
        for i, (title, body) in enumerate(sections, 1):
            body = _clean_md(body)
            if not body.strip():
                continue
            name = "前言" if title == "__preamble__" else _sanitize_section_title(title)
            fname = f"{i:02d}-{name}.md"
            written.append(sec_dir / fname)
            (sec_dir / fname).write_text(body.strip() + "\n", encoding="utf-8")
        result.dst = written
    else:
        dst = out_dir / f"{stem}.md"
        dst.write_text(_clean_md(text) + "\n", encoding="utf-8")
        result.dst = [dst]

    # 状态判定
    if result.char_count < WARN_MIN_CHARS:
        result.status = "warning"
        result.message = f"内容仅 {result.char_count} 字，请人工确认（可能为扫描件/图片）"
    elif src.suffix.lower() in NEEDS_REVIEW_EXT:
        if result.ocr_used:
            result.status = "ok"
            result.message = "OCR 识别已写入 md"
        else:
            result.status = "warning"
            result.message = "图片/音频仅提取到元数据（未启用/失败 OCR）"
    else:
        result.status = "ok"
        result.message = f"输出 {len(result.dst)} 个文件" if len(result.dst) > 1 else "转换成功"

    result.elapsed_ms = int((time.time() - start) * 1000)
    return result


def collect_files(paths: list[Path]) -> list[Path]:
    """把用户拖入的列表展开为待转换文件列表（目录 → 递归取支持的文件）"""
    files: list[Path] = []
    for p in paths:
        if p.is_dir():
            for f in sorted(p.rglob("*")):
                if f.is_file() and f.suffix.lower() in SUPPORTED_EXTENSIONS:
                    if "output" not in f.parts:  # 跳过上次输出，避免重复转换
                        files.append(f)
        elif p.is_file():
            files.append(p)
    # 去重保序
    seen: set[Path] = set()
    out = []
    for f in files:
        rp = f.resolve()
        if rp not in seen:
            seen.add(rp)
            out.append(f)
    return out


def write_report(results: list[FileResult], report_path: Path) -> None:
    """生成转换报告 CSV"""
    import csv

    with report_path.open("w", newline="", encoding="utf-8-sig") as fh:
        w = csv.writer(fh)
        w.writerow(["状态", "源文件", "输出文件", "字数", "OCR", "耗时(ms)", "备注"])
        for r in results:
            w.writerow([
                r.status,
                r.src.name,
                "; ".join(p.name for p in r.dst),
                r.char_count,
                "是" if r.ocr_used else "",
                r.elapsed_ms,
                r.message,
            ])
