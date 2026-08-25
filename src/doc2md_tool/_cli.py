"""CLI 入口：doc2md file1 [file2 ...] [-o out_dir] [--split N]"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from . import __version__
from .service import ConvertOptions, FileResult, convert_file, collect_files, write_report


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(prog="doc2md", description="批量将文件转换为 Markdown")
    ap.add_argument("paths", nargs="+", help="文件或文件夹路径（可多个）")
    ap.add_argument("-o", "--output-dir", default=None, help="输出目录（默认：各文件所在目录的 output/）")
    ap.add_argument(
        "--split", type=int, default=0, metavar="LEVEL",
        help="按标题级别切分文档为多个 md（0=不切分，1~6）",
    )
    ap.add_argument(
        "--no-ocr", action="store_true",
        help="禁用扫描件/图片 OCR（默认自动 OCR）",
    )
    ap.add_argument("--version", action="version", version=f"doc2md {__version__}")
    args = ap.parse_args(argv)

    out_dir = Path(args.output_dir) if args.output_dir else None
    if out_dir is not None:
        out_dir.mkdir(parents=True, exist_ok=True)
    options = ConvertOptions(split_at_heading_level=args.split, output_dir=out_dir,
                            enable_ocr=not args.no_ocr)

    files = collect_files([Path(p) for p in args.paths])
    results = []
    # 用户显式指定的不存在路径 → 直接报错（静默跳过会迷惑用户）
    for p in args.paths:
        if not Path(p).exists():
            results.append(FileResult(src=Path(p), status="error", message="路径不存在"))
            print(f"❌ {p} → -  (路径不存在)")
    if not files and not results:
        print("未找到可转换的文件", file=sys.stderr)
        return 1
    if not files:
        # 所有路径都无效：写报告后直接返回失败
        if out_dir is not None:
            write_report(results, out_dir / "转换报告.csv")
        print(f"共 {len(results)} 个路径均转换失败", file=sys.stderr)
        return 1
    for f in files:
        r = convert_file(f, options)
        results.append(r)
        icon = {"ok": "✅", "warning": "⚠️", "error": "❌"}[r.status]
        targets = " ".join(p.name for p in r.dst) if r.dst else "-"
        print(f"{icon} {f.name} → {targets}  ({r.message})")

    # CLI 模式：报告写到输出目录
    if out_dir is not None:
        write_report(results, out_dir / "转换报告.csv")
    elif len(files) == 1:
        d = files[0].parent / "output"
        d.mkdir(exist_ok=True)
        write_report(results, d / "转换报告.csv")

    err = sum(1 for r in results if r.status == "error")
    return 1 if err else 0


if __name__ == "__main__":
    sys.exit(main())
