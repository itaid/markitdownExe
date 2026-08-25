"""转换核心：调用 markitdown + 后处理（文件名规范化 / 按章节分块 / 校验）

本模块无 GUI 依赖，可独立测试，也是在 Windows EXE 里执行的部分。
"""

from __future__ import annotations

import base64
import io
import posixpath
import re as _re
import re
import time
import zipfile
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

# 内嵌图片处理（docx 的 mammoth 输出 base64 data URI；pptx 的图片被直接丢弃）
_DATA_URI_RE = re.compile(r"!\[[^\]]*\]\(data:image/([a-zA-Z0-9+-]+);base64,([A-Za-z0-9+/=]+)\)")
_EXT_BY_MIME = {"png": ".png", "jpeg": ".jpg", "jpg": ".jpg", "webp": ".webp", "gif": ".gif", "bmp": ".bmp"}
_MIN_IMAGE_BYTES = 4096        # 低于此视为图标/表情，丢弃
_MAX_OCR_IMAGES = 30           # 单文件 OCR 图片数量上限，超出的只落盘不识别
_SLIDE_MARKER = re.compile(r"<!-- Slide number: (\d+) -->")

IMG_DIR_HOLDER: dict = {}      # 当前文档的图片目录（_image_block 使用）

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
    img_count: int = 0


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


def _soft_ocr_image(data: bytes, ext: str) -> str | None:
    """对图片字节做 OCR；图太小/无引擎/异常 → None（不阻断主流程）"""
    import tempfile

    suffix = ext if ext.startswith(".") else "." + ext
    tf = tempfile.NamedTemporaryFile(suffix=suffix, delete=False)
    try:
        tf.write(data)
        tf.close()
    except OSError:
        tf.close()
        Path(tf.name).unlink(missing_ok=True)
        return None
    try:
        try:
            from PIL import Image

            with Image.open(tf.name) as im:
                if im.width < 150 or im.height < 60:  # 小到识别不了
                    return None
        except Exception:  # noqa: BLE001
            pass
        from . import ocr as ocr_mod

        lines = ocr_mod.ocr_image(tf.name)
        if not lines:
            return None
        return ocr_mod.lines_to_markdown(lines) or None
    except Exception:  # noqa: BLE001
        return None
    finally:
        Path(tf.name).unlink(missing_ok=True)


def _image_block(data: bytes, name: str, rel_dir: str, options: ConvertOptions, counter: list) -> str | None:
    counter[0] += 1
    img_dir = Path(IMG_DIR_HOLDER["dir"])
    img_dir.mkdir(parents=True, exist_ok=True)
    (img_dir / name).write_bytes(data)
    block = f"![图片{counter[0]}]({rel_dir}/{name})"
    ocr_md = None
    if options.enable_ocr and counter[0] <= _MAX_OCR_IMAGES:
        ocr_md = _soft_ocr_image(data, name)
    if ocr_md:
        quoted = "\n".join("> " + ln for ln in ocr_md.splitlines() if ln.strip())
        block += f"\n\n> 图片{counter[0]} OCR 内容\n{quoted}"
    return block


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

    # 输出目录
    out_dir = options.output_dir or (src.parent / "output")
    out_dir.mkdir(parents=True, exist_ok=True)
    stem = normalize_filename(src.stem)
    use_split = (options.split_at_heading_level > 0
                 and src.suffix.lower() not in NEEDS_REVIEW_EXT)

    # 内嵌图片抽取（docx/pptx）：消掉 base64 大段，落盘真实文件 + OCR 内联
    if src.suffix.lower() in (".docx", ".pptx", ".ppt") and text.strip():
        text, n_imgs = _process_embedded_images(
            text, src, out_dir / stem, out_dir, stem, use_split, options)
        result.img_count = n_imgs
        result.char_count = len(text)

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

    if use_split and not result.ocr_used:
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

    # 状态判定（图片已 OCR 内联时不计入短文本警告）
    if result.char_count < WARN_MIN_CHARS and result.img_count == 0:
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
        msg = f"输出 {len(result.dst)} 个文件" if len(result.dst) > 1 else "转换成功"
        if result.img_count:
            msg += f"，提取 {result.img_count} 张图片"
        result.message = msg

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
        w.writerow(["状态", "源文件", "输出文件", "字数", "OCR", "图片", "耗时(ms)", "备注"])
        for r in results:
            w.writerow([
                r.status,
                r.src.name,
                "; ".join(p.name for p in r.dst),
                r.char_count,
                "是" if r.ocr_used else "",
                r.img_count or "",
                r.elapsed_ms,
                r.message,
            ])
def _process_embedded_images(
    text: str, src: Path, base_dir: Path, out_dir: Path, stem: str,
    use_split: bool, options: ConvertOptions,
) -> tuple[str, int]:
    """抽取 docx/pptx 内嵌图片：落盘真实文件 + 位置引用 + OCR 内联。

    - docx: mammoth 输出 ![](data:xxx;base64,...)，就地替换（位置天然正确）
    - pptx: markitdown 丢弃图片，从 pptx(zip) 的 ppt/media 按 slide rels 映射插回对应页

    返回 (new_text, 抽取图片张数)。
    """
    if use_split:
        IMG_DIR_HOLDER["dir"] = str(base_dir / "img")
        rel_dir = "img"
    else:
        IMG_DIR_HOLDER["dir"] = str(out_dir / f"{stem}_img")
        rel_dir = f"{stem}_img"
    counter = [0]
    if src.suffix.lower() == ".docx":
        text = _docx_extract_images(text, src, rel_dir, options, counter)
    elif src.suffix.lower() in (".pptx", ".ppt"):
        text = _pptx_insert_slide_images(text, src, rel_dir, options, counter)

    # 清理任何残留的 data URI 占位/截断位（防御性）
    text = re.sub(r"^![^\n]*\(data:image/[a-zA-Z0-9+-]+;base64[^)]*\)\s*$", "", text, flags=re.M)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text, counter[0]


def _pptx_insert_slide_images(text, src, rel_dir, options, counter):
    """从 pptx 包内取每张幻灯片的图片，插到 markitdown 的 slide 标记之后"""
    try:
        raw = src.read_bytes()
    except OSError:
        return text
    try:
        z = zipfile.ZipFile(io.BytesIO(raw))
    except zipfile.BadZipFile:
        return text
    names = set(z.namelist())

    def _resolve(t):
        return posixpath.normpath(posixpath.join("ppt/slides", t)).lstrip("/")

    per_slide: dict[int, list[str]] = {}
    for nm in sorted(names):
        m = re.fullmatch(r"ppt/slides/slide(\d+)\.xml", nm)
        if not m:
            continue
        sn = int(m.group(1))
        rels_nm = f"ppt/slides/_rels/slide{sn}.xml.rels"
        if rels_nm not in names:
            continue
        rels_xml = z.read(rels_nm).decode("utf-8", "ignore")
        rid_map = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rels_xml))
        slide_xml = z.read(f"ppt/slides/slide{sn}.xml").decode("utf-8", "ignore")
        blocks: list[str] = []
        for rid in dict.fromkeys(re.findall(r'r:(?:embed|link)="(rId\d+)"', slide_xml)):
            target = rid_map.get(rid, "")
            if "media/" not in target:  # hyperlink 等非图片关系
                continue
            path = _resolve(target)
            if path not in names:
                continue
            data = z.read(path)
            if len(data) < _MIN_IMAGE_BYTES:
                continue
            ext = Path(path).suffix or ".png"
            name = f"s{sn}_{counter[0] + 1:02d}{ext}"
            blk = _image_block(data, name, rel_dir, options, counter)
            if blk:
                blocks.append(blk)
        if blocks:
            per_slide[sn] = blocks

    # 按 marker 切开（捕获组 => [前导, "1", 页1正文, "2", 页2正文, ...]），
    # 在每页第一个非空行（标题）之后插入该页图片
    parts = _SLIDE_MARKER.split(text)
    if not parts or (len(parts) % 2 != 1):
        return text
    out = [parts[0]]
    for i in range(1, len(parts) - 1, 2):
        sn, page_body = int(parts[i]), parts[i + 1]
        blocks = per_slide.get(sn)
        if not blocks:
            out.append(f"<!-- Slide number: {sn} -->\n" + page_body)
            continue
        lines = page_body.splitlines()
        first = 0
        while first < len(lines) and not lines[first].strip():
            first += 1
        insert_idx = (first + 1) if first < len(lines) else len(lines)
        lines = lines[:insert_idx] + "\n\n".join(blocks).splitlines() + lines[insert_idx:]
        out.append(f"<!-- Slide number: {sn} -->\n" + "\n".join(lines))
    return "\n".join(out)


def _docx_extract_images(text: str, src: Path, rel_dir: str,
                         options, counter: list, extra_only: bool = False) -> str:
    """docx 内嵌图片：直接读 docx.zip 的 word/media 抽取（替代 mammoth，修复截断/位置 bug）。"""
    try:
        raw = src.read_bytes()
    except OSError:
        return text
    try:
        z = zipfile.ZipFile(io.BytesIO(raw))
    except zipfile.BadZipFile:
        return text
    names = set(z.namelist())
    doc_xml, rels_xml = "word/document.xml", "word/_rels/document.xml.rels"
    if doc_xml not in names or rels_xml not in names:
        return text
    rid_map: dict[str, str] = {}
    for m in re.finditer(r'<Relationship\s+([^>]+)/?>',
                         z.read(rels_xml).decode("utf-8", "ignore")):
        attrs = dict(re.findall(r'(\w+)="([^"]*)"', m.group(1)))
        rid_map[attrs.get("Id", "")] = attrs.get("Target", "")
    xml = z.read(doc_xml).decode("utf-8", "ignore")
    # 第一次：按 w:p 顺序扫，收集大图
    ordered, para_pos = [], []
    _gpara = 0
    for pm in re.finditer(r'<w:p[ >](.*?)</w:p>', xml, re.S):
        _gpara += 1
        for rid in dict.fromkeys(re.findall(r'\br?:(?:embed|link)="(rId\d+)"', pm.group(1))):
            target = rid_map.get(rid, "")
            if not target or "media/" not in target:
                continue
            path = posixpath.normpath(posixpath.join("word", target)).lstrip("/")
            if path not in names:
                continue
            data = z.read(path)
            if len(data) < _MIN_IMAGE_BYTES:
                continue
            counter[0] += 1
            ext = Path(path).suffix or ".png"
            name = f"{counter[0]:02d}{ext}"
            (Path(IMG_DIR_HOLDER["dir"]) / name).parent.mkdir(parents=True, exist_ok=True)
            (Path(IMG_DIR_HOLDER["dir"]) / name).write_bytes(data)
            ordered.append(name)
            para_pos.append(_gpara)
    if not ordered:
        return text

    def _block(i, alt):
        name = ordered[i]
        blk = f"![{alt}]({rel_dir}/{name})"
        if options.enable_ocr and i + 1 <= _MAX_OCR_IMAGES:
            ochr = _soft_ocr_image((Path(IMG_DIR_HOLDER["dir"]) / name).read_bytes(), name)
            if ochr:
                q = "\n".join("> " + ln for ln in ochr.splitlines() if ln.strip())
                blk += f"\n\n> 图片{i + 1} OCR 内容\n{q}"
        return blk

    # extra_only 模式：只追加尚未在正文中出现的图片
    if extra_only:
        missing = [i for i, n in enumerate(ordered) if n not in text]
        if not missing:
            return text
        parts = [_block(i, f"图片{i+1}") for i in missing]
        return text.rstrip() + "\n\n" + "\n\n".join(parts)

    # 判定：mammoth 是否保留了完整图片引用（无 alt 图 => 截断占位）
    md_imgs = list(re.finditer(r"!\[[^\]]*\]\([^)]*base64[^)]*\)", text))
    if len(md_imgs) == len(ordered):
        # mammoth 有 alt 路径：in 有结构 ![](data:${base64<截断>})，按序替换
        out = text
        for i in reversed(range(len(ordered))):
            m = md_imgs[i]
            alt = re.match(r"!\[([^\]]*)]", m.group(0)).group(1).strip() or f"图片{i+1}"
            out = out[:m.start()] + _block(i, alt) + out[m.end():]
        return out

    # 无 alt 图路径：按文档 w:p 顺序估算段落到 text 上的位置
    # 分母：整个文档 w:p 数；分子：para_pos[i]
    n_wp = len(re.findall(r"<w:p[ =>]", xml))
    texts = text.split("\n\n")
    for i in reversed(range(len(ordered))):
        est = int(para_pos[i] * max(len(texts)-1, 1) / max(n_wp, 1))
        texts[est] = (texts[est].rstrip() + "\n\n") + _block(i, f"图片{i+1}") if texts[est].strip() else _block(i, f"图片{i+1}")
    return "\n".join(texts)
