# doc2md

拖拽文件批量转 Markdown 的桌面工具（Windows EXE / macOS / CLI 通用），基于微软 [markitdown](https://github.com/microsoft/markitdown) + 内置本地 OCR，针对 RAG（如 RAGFlow）流水线优化。

## 特性

- 🖱️ 拖拽文件/文件夹，批量转换（PDF / Word / PPT / Excel / HTML / EPUB / 音频 等 20+ 格式）
- 📝 **内置本地 OCR**：扫描件 / 纯图片 PDF / 图片在常规解析无内容时自动识别为文字；纯本地运行不上传，可在界面关闭
- 🔪 可选**按标题级别切分**：一份制度文档自动拆成 `01-前言.md`、`02-适用范围.md`... 每个分块自带标题，直接对齐 RAG 分块粒度
- 🧹 文件名自动规范化（去非法字符、去"最终版/副本"噪音、保留业务编号）
- 📋 生成 `转换报告.csv`：每文件状态 / 字数 / 是否 OCR / 耗时 / 错误原因，带 BOM 可直接 Excel 打开
- ⚠️ 坏文件防护：魔数校验（假 docx 直接报错不吐乱码）、加密/空文件标黄、路径不存在明确提示
- 💻 GUI 和 CLI 双入口

## 使用

### GUI（Windows 用户）

从 [Releases / Actions Artifacts](https://github.com/yourname/doc2md/actions) 下载 `doc2md_vX.Y_Z_win64.zip`，解压后双击 `doc2md.exe`：

1. 把文件（或整个文件夹）拖进窗口
2. 可选设置：按标题切分级别（制度类文档推荐 2）、扫描件自动 OCR（默认开）
3. 点「开始转换」
4. 结果在**文件所在目录的 `output/`**：单文件格式为 `文档名.md`，切分格式为 `文档名/01-章节.md`...
5. 同目录生成 `转换报告.csv`，⚠️/❌ 条目请人工确认后重新上传 RAG

### CLI

```bash
pip install "doc2md-tool[gui,ocr]"
doc2md ./资料/ -o ./output/ --split 2   # 按二级标题切分
doc2md ./扫描件.pdf                      # 自动 OCR 兜底
doc2md ./图片.png --no-ocr               # 关闭 OCR
```

## 给 RAG 用户

转换端切块优于在 RAG 引擎内部分块：每个输出 md = 一个语义完整的章节 = 一个理想 chunk，文件名即章节名，检索命中直接溯源。`--split 2`（按二级标题）是大多数制度/流程类文档的最佳起点；标题结构不清晰的文档保持 `--split 0`。

## 打包与发布（tag 自动版）

- 打 tag（如 `git tag v0.2.0 && git push --tags`）或 push `main` → GitHub Actions 自动：ubuntu 跑单测 → windows-latest 上 PyInstaller 打 EXE
- 版本号规则：触发 tag > 手动触发输入的 tag > `latest`，会同步写入程序关于信息
- 产物：`doc2md_vX.Y.Z_win64.zip`（Actions → Artifacts，保留 90 天）
- 图标：`pack/icon.ico`（源图 `pack/icon_512.png`）

## OCR 选型说明（三选一的实测结论）

| 引擎 | 中文准确率 | 运行时体积 | 结论 |
|---|---|---|---|
| **rapidocr_onnxruntime** | 最高（表格单元格均正确，含标点） | ~90MB（onnxruntime, 纯 CPU） | ✅ 默认 |
| easyocr | 中（句块切碎、标点大量丢失） | ~700MB（torch） | 备选后端已封装 |
| paddleocr 3.x | —（无法运行） | ~200MB | 暂不接入：M 系 macOS PIR 编译器 bug |

后端可切换：环境变量 `DOC2MD_OCR=paddle|easy`，统一接口与扩展位见 `src/doc2md_tool/ocr.py`（新增引擎只需实现 `OCREngine.ocr_image`）。

## Development

```bash
python3.12 -m venv .venv && source .venv/bin/activate
pip install -e ".[gui,ocr]" pytest
pytest
python -m doc2md_tool.gui   # 界面调试
```

核心逻辑（`service.py` / `ocr.py`）无 GUI 依赖，CI 在 ubuntu 上直接跑单测；EXE 打包仅 Windows CI 执行。

## License

MIT（图标素材除外，图标遵循来源网站许可）
