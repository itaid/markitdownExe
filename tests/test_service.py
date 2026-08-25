"""转换核心单元测试（无 GUI 依赖，跨平台 CI 可跑）"""
from pathlib import Path

import pytest

from doc2md_tool.service import (
    ConvertOptions,
    collect_files,
    convert_file,
    normalize_filename,
    split_by_headings,
)


MD_DOC = """# 员工手册

## 总则
本文档规定公司员工的基本行为准则。

## 考勤制度
- 工作时间 9:00-18:00
- 打卡方式：钉钉

### 请假流程
主管审批 -> HR备案

## 保密条款
员工不得泄露商业机密。
"""


def test_normalize_filename():
    assert normalize_filename("新建文件夹 (2) 最终版") == "新建文件夹"
    assert normalize_filename("a/b:c*d?") == "a_b_c_d"
    assert normalize_filename("   copy  ") == "copy"
    assert normalize_filename("   ") == "unnamed"
    assert len(normalize_filename("x" * 200)) <= 80


def test_split_by_headings():
    sections = split_by_headings(MD_DOC, 2)
    # 首标题 “# 员工手册” 之前无内容 → 不产生 preamble；level 1 标题也作为切分点
    titles = [t for t, _ in sections]
    assert titles == ["员工手册", "总则", "考勤制度", "保密条款"]
    # 每个小节带上自己的标题行
    assert sections[2][1].startswith("## 考勤制度")
    assert "请假流程" in sections[2][1]
    # level=1 时不按 ## 切
    assert len(split_by_headings(MD_DOC, 1)) == 1  # 全部归入 “员工手册”
    # level=0 不切
    assert len(split_by_headings(MD_DOC, 0)) == 1
    # 前面有引言内容时产生 preamble
    with_preamble = "引言段落\\n\\n" + MD_DOC
    assert split_by_headings(with_preamble, 2)[0][0] == "__preamble__"


def test_convert_md_file(tmp_path: Path):
    src = tmp_path / "测试文档.md"
    src.write_text(MD_DOC, encoding="utf-8")
    out = tmp_path / "output"
    r = convert_file(src, ConvertOptions(output_dir=out))
    assert r.status == "ok", r.message
    assert r.dst == [out / "测试文档.md"]
    content = r.dst[0].read_text(encoding="utf-8")
    assert "考勤制度" in content


def test_convert_split(tmp_path: Path):
    src = tmp_path / "制度.md"
    src.write_text(MD_DOC, encoding="utf-8")
    out = tmp_path / "output"
    r = convert_file(src, ConvertOptions(split_at_heading_level=2, output_dir=out))
    assert r.status == "ok", r.message
    assert len(r.dst) == 4  # preamble + 3 个二级标题
    assert (out / "制度").is_dir()
    for p in r.dst:
        assert p.exists()
        assert p.read_text(encoding="utf-8").strip()


def test_convert_report(tmp_path: Path):
    import csv

    src = tmp_path / "a.md"
    src.write_text("hello world test content", encoding="utf-8")
    out = tmp_path / "output"
    r = convert_file(src, ConvertOptions(output_dir=out))
    from doc2md_tool.service import write_report

    write_report([r], out / "report.csv")
    with (out / "report.csv").open(encoding="utf-8-sig") as fh:
        rows = list(csv.reader(fh))
    assert rows[0] == ["状态", "源文件", "输出文件", "字数", "OCR", "耗时(ms)", "备注"]
    assert rows[1][1] == "a.md"


def test_collect_files_dir_ignores_output():
    import tempfile

    with tempfile.TemporaryDirectory() as d:
        root = Path(d)
        (root / "x.pdf").write_bytes(b"%PDF-fake")
        (root / "output").mkdir()
        (root / "output" / "old.md").write_text("skip me")
        (root / "notes.txt").write_text("hi")
        files = collect_files([root])
        names = {f.name for f in files}
        assert names == {"x.pdf", "notes.txt"}


def test_error_result_on_bad_file(tmp_path: Path):
    src = tmp_path / "broken.docx"
    src.write_bytes(b"this is not a real docx file" * 20)
    r = convert_file(src, ConvertOptions(output_dir=tmp_path / "output"))
    # 损坏文件应转为 error/warning 而不是抛异常
    assert r.status in ("error", "warning")
    assert r.message


def test_ocr_graceful_when_unavailable(tmp_path: Path, monkeypatch):
    """无 OCR 引擎时：转换不崩，图片降级为 warning"""
    import doc2md_tool.service as svc

    monkeypatch.setattr(svc, "_run_ocr_for_file", lambda src: None)
    # 造一个无文字层的图片：纯色 PNG
    from PIL import Image
    src = tmp_path / "blank.png"
    Image.new("RGB", (100, 100), "white").save(src)
    r = svc.convert_file(src, svc.ConvertOptions(output_dir=tmp_path / "out", enable_ocr=True))
    assert r.status in ("warning", "error")
    assert r.message
