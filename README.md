# PPTconvert

`PPTconvert` 现在是一个面向公考试题整理与授课输出的桌面工具，不再只是早期的 `Word -> PPT` 小脚本。

当前主链路已经收敛成两条共享工程流：

- `PDF -> 抽取 / 解析 / 科目推断 -> ExamProject -> 共享预览 / 编辑 -> Word / PPT / JSON`
- `Word -> 解析 / 科目推断 -> ExamProject -> 共享预览 / 编辑 -> PPT / JSON`

## 当前能力

### PDF 试卷整理

- 支持政治理论、常识判断、言语理解与表达、数量关系、判断推理、资料分析
- 支持整份文档科目提示：`自动识别 / 政治理论 / 常识判断 / 言语理解与表达 / 数量关系 / 判断推理 / 资料分析`
- 支持无大标题、单科整卷、标题中途缺失时的启发式科目推断
- 支持资料分析材料的文本、图片、表格区域保留与裁图复用
- 支持导出题本 Word、授课 PPT、工程 JSON

### Word 生成 PPT

- 支持直接导入整理好的 `.docx`
- 先解析为统一的 `ExamProject`，再进入共享预览页
- 可在导出前继续修改题干、选项、选项布局、科目归类
- 保留旧版 `Word -> PPT` 输出能力，但已经接到新的共享预览/编辑链路

### 共享预览与人工修订

- 结构树预览篇题、材料、题目
- 题干实时编辑
- 选项文字、顺序、增删、图片替换 / 清除 / PDF 重裁
- 单题选项布局覆盖：`跟随全局 / 一行四项 / 两行两列 / 四行竖排`
- 资料分析材料原貌预览
- 题干图片预览
- 未保存修改保护
- 未知科目保留与整段改科目

## GUI 工作流

启动方式：

```powershell
python main.py
```

当前 GUI 有三个主标签页：

- `PDF 试卷整理`
- `Word 生成 PPT`
- `PPT 导出设置`

PDF 与 Word 都会汇入同一套预览/编辑区域，再共用导出设置。

## CLI 用法

### 启动 GUI

```powershell
python main.py
```

### 处理 PDF

```powershell
python main.py --pdf-input exam.pdf --docx-output exam_题本.docx
python main.py --pdf-input exam.pdf --ppt-output exam.pptx --subject data
python main.py --pdf-input exam.pdf --ppt-output exam.pptx --subject politics,common_sense,verbal
```

### 处理 Word

```powershell
python main.py -i input.docx -o output.pptx -t template.pptx
```

常用参数包括：

- `--pdf-input`
- `--docx-output`
- `--ppt-output`
- `--manifest-output`
- `--subject`
- `--question-range`
- `--template`

## 目录概览

- [main.py](C:/Users/17679/Desktop/PPTconvert/main.py)
  命令行入口；无参数时启动 GUI。
- [gui/app.py](C:/Users/17679/Desktop/PPTconvert/gui/app.py)
  主界面、共享预览编辑、两条导入流程。
- [workflows/project_flow.py](C:/Users/17679/Desktop/PPTconvert/workflows/project_flow.py)
  `PDF/Word -> ExamProject -> 导出` 编排入口。
- [core/pdf_exam_extract.py](C:/Users/17679/Desktop/PPTconvert/core/pdf_exam_extract.py)
  PDF 文本块、图片块、页面图片信息抽取。
- [core/pdf_exam_parse.py](C:/Users/17679/Desktop/PPTconvert/core/pdf_exam_parse.py)
  PDF 题目、材料、选项、科目切分。
- [core/word_parser.py](C:/Users/17679/Desktop/PPTconvert/core/word_parser.py)
  Word 题目解析与科目修正。
- [core/subject_inference.py](C:/Users/17679/Desktop/PPTconvert/core/subject_inference.py)
  无标题 / 单科 / 标题缺失时的启发式科目推断。
- [ingest/pdf/project_builder.py](C:/Users/17679/Desktop/PPTconvert/ingest/pdf/project_builder.py)
  PDF 解析结果转 `ExamProject`。
- [ingest/docx/project_builder.py](C:/Users/17679/Desktop/PPTconvert/ingest/docx/project_builder.py)
  Word 解析结果转 `ExamProject`。
- [domain/](C:/Users/17679/Desktop/PPTconvert/domain)
  统一工程模型、筛选器、编辑动作。
- [exporters/](C:/Users/17679/Desktop/PPTconvert/exporters)
  Word / PPT / JSON 导出与材料裁图。
- [tests/](C:/Users/17679/Desktop/PPTconvert/tests)
  解析、工程构建、编辑、导出回归。

## 文档

- [架构说明](C:/Users/17679/Desktop/PPTconvert/docs/ARCHITECTURE.md)
- [当前进度](C:/Users/17679/Desktop/PPTconvert/docs/STATUS.md)
- [本次整理记录](C:/Users/17679/Desktop/PPTconvert/docs/WORKLOG.md)
- [GitHub 协作说明](C:/Users/17679/Desktop/PPTconvert/docs/GITHUB.md)

## 测试

常用回归命令：

```powershell
python -m unittest discover -s tests -v
python -m py_compile .\main.py .\gui\app.py
```

## 当前边界

- 扫描版 PDF、OCR 错字、多栏复杂版式仍会影响切题与分科
- 没有大标题、且多科混排、且文本信号很弱的文档，目前仍以启发式推断为主
- 资料分析跨页、极端复杂表格、超弱结构化 Word 仍需继续补规则
- 共享预览已经能覆盖大部分人工修订，但还不是最终导出版面的完全替代
