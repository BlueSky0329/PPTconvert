# 架构说明

## 总览

项目当前围绕统一工程模型 `ExamProject` 组织，已经不是两套完全割裂的 `PDF` 与 `Word` 逻辑。

当前整体结构是：

`PDF / Word -> 解析与启发式分类 -> ExamProject -> 共享预览 / 编辑 -> Word / PPT / JSON`

设计上的核心目标是：

- 让 `PDF` 与 `Word` 最终汇入同一套预览和编辑能力
- 把“题目切分、科目归类、材料保留、人工修订”放在导出之前解决
- 将 `Word / PPT / JSON` 视为 `ExamProject` 的不同输出层，而不是彼此独立的流程

## 当前主流程

### 1. PDF 流

`PDF -> pdf_exam_extract -> pdf_exam_parse -> ingest.pdf.project_builder -> ExamProject`

职责拆分如下：

1. [core/pdf_exam_extract.py](C:/Users/17679/Desktop/PPTconvert/core/pdf_exam_extract.py)
   - 基于 PyMuPDF 抽取文本块、图片块、页面图片信息
   - 处理双栏页读序与页面图片补提取
2. [core/pdf_exam_parse.py](C:/Users/17679/Desktop/PPTconvert/core/pdf_exam_parse.py)
   - 识别篇题、材料、题号、选项
   - 处理题干/选项续行、材料拆组、标题缺失 fallback
3. [core/subject_inference.py](C:/Users/17679/Desktop/PPTconvert/core/subject_inference.py)
   - 在无标题、单科整卷、篇题缺失时做启发式科目推断
4. [ingest/pdf/project_builder.py](C:/Users/17679/Desktop/PPTconvert/ingest/pdf/project_builder.py)
   - 将解析结果映射为 `ExamProject`
   - 保存题干资源、材料图片、PDF 页面区域

### 2. Word 流

`Word -> word_parser -> ingest.docx.project_builder -> ExamProject`

职责拆分如下：

1. [core/word_parser.py](C:/Users/17679/Desktop/PPTconvert/core/word_parser.py)
   - 解析整理后的题本 Word
   - 提取题号、题干、选项、材料、题干图片
   - 在标题缺失时调用启发式科目推断
2. [core/subject_inference.py](C:/Users/17679/Desktop/PPTconvert/core/subject_inference.py)
   - 作为 Word 无标题和漂移修正的共用分类器
3. [ingest/docx/project_builder.py](C:/Users/17679/Desktop/PPTconvert/ingest/docx/project_builder.py)
   - 将 Word 解析结果转成 `ExamProject`
   - 让 Word 与 PDF 共享同一套后续预览与导出逻辑

### 3. 共享工程流

`ExamProject -> selectors / project_editor -> GUI preview -> exporters`

这部分是当前架构的中心：

- [domain/models.py](C:/Users/17679/Desktop/PPTconvert/domain/models.py)
  定义统一工程模型
- [domain/selectors.py](C:/Users/17679/Desktop/PPTconvert/domain/selectors.py)
  负责科目、题号范围筛选，并保留 `unknown` 兜底段
- [domain/project_editor.py](C:/Users/17679/Desktop/PPTconvert/domain/project_editor.py)
  负责 GUI 的人工修题操作
- [gui/app.py](C:/Users/17679/Desktop/PPTconvert/gui/app.py)
  负责两条导入流、共享预览、共享导出设置

## 核心模型

统一工程模型位于 [domain/models.py](C:/Users/17679/Desktop/PPTconvert/domain/models.py)。

关键对象包括：

- `ExamProject`
  整份工程，承载来源信息、选中科目、章节列表
- `Section`
  一个科目段落，可能是普通客观题，也可能是资料分析
- `MaterialSet`
  资料分析材料，包含正文、材料图片、页面区域和挂载题目
- `QuestionNode`
  统一题目对象，包含题干、题干资源、选项、答案、选项布局
- `OptionNode`
  统一选项对象，支持文本、图片、原 PDF 区域信息
- `AssetRef`
  图片等资源引用
- `PageRegion`
  从 PDF 记录下来的页码与裁切区域

统一模型的价值在于：

- PDF 和 Word 可以共享预览编辑界面
- Word / PPT / JSON 导出不用重复理解上游文档结构
- GUI 的人工修正不需要分别改两条链路

## GUI 结构

当前 GUI 位于 [gui/app.py](C:/Users/17679/Desktop/PPTconvert/gui/app.py)，主要包含三块：

- `PDF 试卷整理`
  负责导入 PDF、科目选择、预览生成、结果导出
- `Word 生成 PPT`
  负责导入 Word 并转入共享预览
- `PPT 导出设置`
  负责模板、字号、版式等导出参数

无论来源是 PDF 还是 Word，都会进入同一套预览编辑区：

- 左侧结构树：篇题 / 材料 / 题目
- 右侧题目编辑：题干、选项、选项布局、图片修订
- 右侧材料原貌：资料分析材料截图 / 原图
- 右侧结构详情：便于排查模型状态

## 导出层

统一导出位于 [exporters/](C:/Users/17679/Desktop/PPTconvert/exporters)：

- [docx_booklet.py](C:/Users/17679/Desktop/PPTconvert/exporters/docx_booklet.py)
  `ExamProject -> Word`
- [pptx_slides.py](C:/Users/17679/Desktop/PPTconvert/exporters/pptx_slides.py)
  `ExamProject -> PPT`
- [manifest_json.py](C:/Users/17679/Desktop/PPTconvert/exporters/manifest_json.py)
  `ExamProject <-> JSON`
- [material_crops.py](C:/Users/17679/Desktop/PPTconvert/exporters/material_crops.py)
  从原 PDF 按页面区域裁图，供 Word / PPT / GUI 共用

其中 PPT 输出层最终仍会复用 [core/ppt_generator.py](C:/Users/17679/Desktop/PPTconvert/core/ppt_generator.py)，也就是说旧版 Word 生成器与新工程流在输出层已经汇合。

## 编排层

[workflows/project_flow.py](C:/Users/17679/Desktop/PPTconvert/workflows/project_flow.py) 负责端到端编排：

- 构建 PDF 工程
- 构建 Word 工程
- 按科目与题号筛选
- 导出 Word / PPT / JSON

[main.py](C:/Users/17679/Desktop/PPTconvert/main.py) 则提供：

- 无参数启动 GUI
- `--pdf-input` 走 PDF 工程流
- `-i / --input` 走 Word 输入流

## 测试结构

[tests/](C:/Users/17679/Desktop/PPTconvert/tests) 按风险点分层：

- `test_pdf_exam_extract.py`
  页面读序、图片补提取
- `test_pdf_exam_parse.py`
  题目切分、材料切分、科目 fallback
- `test_exam_project.py`
  `ParsedExam -> ExamProject`
- `test_word_parser.py`
  Word 解析与无标题推断
- `test_word_project_builder.py`
  `WordQuestion -> ExamProject`
- `test_project_editor.py`
  GUI 编辑动作、答案同步
- `test_docx_booklet.py`
  Word 导出
- `test_ppt_generator.py`
  PPT 输出层
- `test_manifest_json.py`
  JSON round-trip
- `test_selectors.py`
  科目筛选与 `unknown` 保留

## 当前边界

当前架构已经能稳住大部分“单科 / 缺标题 / 可人工修题”的使用场景，但仍有边界：

- 多科混排且没有任何大标题时，只能依赖启发式推断
- 扫描版 PDF 与 OCR 错字仍会影响切题与分类
- 复杂跨页资料分析材料需要继续增强
- GUI 预览主要服务人工校对，还不是最终导出版面的完全等价渲染
