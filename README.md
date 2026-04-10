# PPTconvert

一个本地化的公考试卷整理工具，当前主链路是：

`PDF -> 结构化工程 -> Word / PPT / JSON`

项目已经不再把 GUI 设计成“Word 转 PPT 工具箱”。现在的 GUI 聚焦 PDF 试卷整理，适合把整套真题按科目抽取、预览、修补后再导出。

## 当前状态

- GUI：单一 PDF 向导，步骤为导入 PDF、识别设置、结果预览、导出结果
- 科目识别：支持政治理论、常识判断、言语理解与表达、数量关系、判断推理、资料分析
- 导出结果：支持题本 Word、授课 PPT、工程清单 JSON
- 编辑能力：支持在 GUI 中查看结构树、调整资料分析材料与题目
- 兼容能力：CLI 仍保留旧的 `Word -> PPT` 入口，便于兼容历史用法

## 推荐工作流

### GUI

```bash
python main.py
```

GUI 默认进入 PDF 工作流：

1. 选择 PDF
2. 勾选需要识别的科目
3. 预览结果并修补结构
4. 导出 Word / PPT / JSON

### CLI

```bash
# PDF -> Word
python main.py --pdf-input exam.pdf --docx-output exam_题本.docx

# PDF -> PPT
python main.py --pdf-input exam.pdf --ppt-output exam.pptx

# PDF -> Word + PPT，并筛选题号
python main.py --pdf-input exam.pdf --docx-output exam_题本.docx --ppt-output exam.pptx --question-range 66-85,111-120

# PDF -> 仅部分科目
python main.py --pdf-input exam.pdf --ppt-output verbal_only.pptx --subject verbal
python main.py --pdf-input exam.pdf --ppt-output mixed.pptx --subject politics,common_sense,verbal

# 兼容旧流程：Word -> PPT
python main.py -i exam.docx -o exam.pptx
```

## 能力边界

当前这套规则对以下场景做了较多修补：

- 六大模块识别
- PDF 内嵌图片与页面图片补提取
- 资料分析材料截图复用
- 文本表格材料按 PDF 区域裁图导出
- `D` 选项续行、行内下一题、题号独立成行等切题边界
- 双栏页面的读序修正

仍然需要警惕的场景：

- 扫描版 PDF 或 OCR 错字严重的 PDF
- 跨页材料
- 多列表格读序异常
- 极端版式下的题号、选项、材料混排

这类问题目前仍建议用真实试卷做回归，再按错例补规则。

## 安装

### 环境要求

- Python 3.10+
- Windows 为主

### 安装依赖

```bash
pip install -r requirements.txt
```

依赖包含：

- `pymupdf`
- `python-docx`
- `python-pptx`
- `Pillow`
- `ttkbootstrap`

## 项目结构

```text
PPTconvert/
├── main.py
├── core/         # PDF 抽取/解析、旧版 Word->PPT 核心
├── ingest/       # PDF -> 工程模型构建
├── domain/       # 统一工程模型、筛选与编辑操作
├── exporters/    # Word / PPT / JSON 导出
├── workflows/    # 端到端流程编排
├── gui/          # PDF 向导界面
├── tests/        # 回归测试
├── docs/         # 架构、进度、GitHub 协作说明
├── templates/    # PPT 模板
├── examples/     # 可分享的样例
└── outputs/      # 本地导出目录（默认忽略）
```

更详细的模块说明见 [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md)。

## 测试

全量回归：

```bash
python -m unittest discover -s tests -v
```

常用定向测试：

```bash
python -m unittest tests.test_pdf_exam_parse -v
python -m unittest tests.test_pdf_exam_extract -v
python -m unittest tests.test_exam_project -v
```

## 仓库约定

- 根目录下的 PDF、DOCX、PPT、资产目录和工程 JSON 视为本地输入/输出，默认不提交
- 可分享样例请放到 `examples/`
- 本地导出建议放到 `outputs/`

## 文档

- [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md)
- [docs/STATUS.md](docs/STATUS.md)
- [docs/GITHUB.md](docs/GITHUB.md)
- [CONTRIBUTING.md](CONTRIBUTING.md)

## License

MIT
