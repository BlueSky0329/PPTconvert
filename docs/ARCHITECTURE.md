# 架构说明

## 总览

项目已经从早期的 `Word -> PPT` 单链路，演进为以 `ExamProject` 为中心的统一工程流：

`PDF -> 抽取 -> 解析 -> 工程模型 -> 预览/编辑 -> Word / PPT / JSON`

其中：

- GUI 主要服务 PDF 工作流
- CLI 既支持 PDF 工作流，也保留旧版 `Word -> PPT` 兼容入口
- `ExamProject` 是 PDF 工作流的统一中间层

## 目录职责

### 根目录

- `main.py`
  - 命令行入口
  - 无参数启动 GUI
  - `--pdf-input` 走 PDF 工程流
  - `-i/--input` 走旧版 Word -> PPT 兼容流

### `gui/`

- `app.py`
  - PDF 向导界面
  - 负责导入、筛选、预览、导出和局部编辑
- `ui_constants.py`
  - 主题、文案和常量配置

### `workflows/`

- `project_flow.py`
  - 端到端编排入口
  - 负责：
    - 构建 PDF 工程
    - 按科目/题号筛选
    - 导出 Word / PPT / JSON

### `domain/`

- `models.py`
  - 统一工程模型定义
- `selectors.py`
  - 题号范围、科目筛选
- `project_editor.py`
  - GUI 里对工程的局部编辑操作

核心对象：

- `ExamProject`
- `Section`
- `MaterialSet`
- `QuestionNode`
- `OptionNode`
- `AssetRef`
- `PageRegion`

### `core/`

这里同时承载两类能力。

#### 1. PDF 抽取与解析

- `pdf_exam_extract.py`
  - 用 PyMuPDF 抽取文本块、图片块和页面图片信息
  - 负责页面读序、双栏识别、图片补提取
- `pdf_exam_parse.py`
  - 将抽取结果解析为六大模块的结构化中间模型
  - 处理题号、材料、选项、续行、切题边界
- `pdf_exam_models.py`
  - PDF 解析阶段的中间数据结构
- `exam_docx_writer.py`
  - 旧版 PDF -> Word 输出模块，当前仍保留
- `pdf_exam_pipeline.py`
  - 旧版 PDF 工作流入口，当前以兼容为主

#### 2. 旧版 Word -> PPT 链路

- `word_parser.py`
  - 解析整理好的 Word 题本
- `models.py`
  - 旧版 PPT 生成使用的数据结构
- `ppt_generator.py`
  - 生成 PPT
- `template_manager.py`
  - 加载和创建模板
- `template_style.py`
  - 提取模板占位样式

说明：

- 新的 PDF 主链路导出 PPT 时，最终仍会复用 `core.ppt_generator`
- 也就是说，旧版 PPT 生成器目前已经成为新旧两条流共享的输出层

### `ingest/`

- `pdf/layout.py`
  - 抽取 PDF 文本行和块级几何信息
- `pdf/project_builder.py`
  - 将 PDF 解析结果映射为 `ExamProject`
  - 负责：
    - 题干与选项落盘
    - 材料正文、材料图片、页面区域保存
    - 资源文件复制到工程资产目录

### `exporters/`

- `docx_booklet.py`
  - `ExamProject -> Word`
- `pptx_slides.py`
  - `ExamProject -> PPT`
- `manifest_json.py`
  - `ExamProject -> JSON`
- `material_crops.py`
  - 资料分析材料按 PDF 页面区域裁图
  - 被 Word / PPT / GUI 复用

### `tests/`

测试按风险点分层：

- `test_pdf_exam_extract.py`
  - 页面读序、图片补提取
- `test_pdf_exam_parse.py`
  - 题目切分、材料切分、选项边界
- `test_exam_project.py`
  - `ParsedExam -> ExamProject`
- `test_docx_booklet.py`
  - Word 导出
- `test_project_editor.py`
  - GUI 编辑动作
- `test_word_parser.py`
  - 旧版 Word -> PPT 兼容流
- `test_ppt_generator.py`
  - PPT 输出层

## 主流程

### PDF -> ExamProject

1. `core/pdf_exam_extract.py`
   - 提取文本与图片
   - 补充页面缺失图片
   - 对明显双栏页重排块顺序
2. `core/pdf_exam_parse.py`
   - 识别篇题
   - 切材料
   - 切题
   - 切选项
3. `ingest/pdf/project_builder.py`
   - 生成 `ExamProject`
   - 复制资产到工程目录
4. `domain/selectors.py`
   - 按科目、题号筛选

### ExamProject -> 导出

1. `exporters/docx_booklet.py`
   - 输出题本 Word
   - 资料分析材料优先插入 PDF 区域裁图
2. `exporters/pptx_slides.py`
   - 将 `ExamProject` 转成旧版 `Question`
   - 调用 `core.ppt_generator` 生成 PPT
3. `exporters/manifest_json.py`
   - 导出工程清单 JSON

### ExamProject -> GUI

1. GUI 读取 `ExamProject`
2. 左侧树展示科目、材料、题目
3. 右侧预览文本、选项和资料分析材料原貌
4. 用户可做局部修补后导出

## 关键设计

### 统一工程模型

PDF 解析结果不直接写 Word 或 PPT，而是先落到 `ExamProject`。这样：

- GUI 和 CLI 共享同一套数据
- 筛选逻辑只写一份
- 导出器可以独立演进

### 材料双表示

资料分析材料目前同时保存三份信息：

- `body_lines`
  - 文本版正文
- `body_assets`
  - 抽取到的原始图片
- `body_regions`
  - 对应 PDF 页面区域

好处是：

- Word 可以在文本回退和截图之间切换
- PPT 可以优先保留原始表格/图表外观
- GUI 可以展示“材料原貌”

### 输出层复用

新链路没有重写一套 PPT 生成器，而是先将 `ExamProject` 映射为旧版 `Question`，继续复用已经稳定的 `core.ppt_generator`。这降低了重构成本，但也意味着：

- `core/` 目录暂时同时承载新旧两代能力
- 后续若继续整理代码，可以考虑把旧版兼容层再明确拆分

## 当前已知边界

- OCR 质量差的扫描 PDF 仍不稳定
- 跨页材料仍主要靠启发式处理
- 极端多栏或复杂表格排版仍可能读序异常
- 工程 JSON 当前主要用于导出与检查，不是完整的可回载编辑格式

## 推荐后续整理方向

1. 为真实 PDF 建立回归样本集
2. 将 GUI 编辑能力和工程 JSON 回载打通
3. 进一步把旧版 Word -> PPT 兼容层与新 PDF 主链隔离
4. 给解析器增加调试视图或诊断日志导出
