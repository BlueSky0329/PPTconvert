# 重构蓝图

## 目标

把项目从“Word -> PPT 为主，PDF -> Word 为附加功能”改成“PDF 试卷整理为主，Word/PPT 都是导出结果”。

目标业务流程：

1. 导入试卷 PDF
2. 自动识别题号、科目、材料块、题目、选项、图片
3. 按科目 / 题号范围筛选
4. 预览并人工修正
5. 导出题本 Word
6. 从同一份结构化数据直接导出 PPT

关键原则：

- `PDF` 是主输入，不再把 `Word` 当主数据源
- `Word` 是导出物，不应再反向解析后生成 `PPT`
- `Word` 和 `PPT` 必须共享同一份标准化题目数据
- 资料分析必须保留原始 PDF 的页面坐标，以便裁切材料图

---

## 当前架构的主要问题

### 1. 主链路错位

当前入口仍然围绕 `Word -> PPT` 设计，`PDF -> Word` 是旁路功能。

这与真实业务不一致。真实业务是先拿试卷 PDF，再摘题，再出题本，再出 PPT。

### 2. 数据发生二次解析和信息损耗

当前 `PDF -> Word -> PPT` 需要先写出 Word，再重新解析 Word 才能生成 PPT。

这会带来几个问题：

- 解析规则重复维护
- 图片、题号、材料边界容易丢失
- 无法稳定复用 PDF 中的原始版面信息

### 3. 资料分析缺少几何信息

当前 PDF 解析结果主要保留文本和图片路径，没有把材料块的 `page + bbox` 保留下来。

这导致 PPT 里无法稳定地把“同一篇材料”裁成一张图，供五道题重复使用。

### 4. 模块职责还不够清晰

当前 `core` 中同时混有：

- 输入解析
- 导出
- GUI 直接依赖的生成逻辑
- 面向 Word 的模型
- 面向 PDF 的模型

模型是分裂的，缺少一个真正统一的领域模型。

---

## 目标架构

建议改成四层：

1. `domain`
2. `ingest`
3. `pipeline`
4. `exporters`

### 1. domain

定义唯一的数据标准，所有导入和导出都围绕它。

核心对象建议：

- `ExamProject`
- `PaperSource`
- `QuestionSelection`
- `Section`
- `Question`
- `MaterialSet`
- `Option`
- `AssetRef`
- `PageRegion`

其中：

- `Question` 表示普通单题
- `MaterialSet` 表示一份材料 + 多道关联题
- `PageRegion` 保存 PDF 页码和裁切区域
- `AssetRef` 保存题干图、选项图、材料图等资源引用

### 2. ingest

负责把不同来源转成统一领域模型。

建议包含：

- `ingest/pdf/`
- `ingest/docx/`
- `ingest/debug_markdown/`

其中：

- `pdf` 是主适配器
- `docx` 只做兼容输入，不再是核心主链路
- `debug_markdown` 用于输出中间检查结果，方便人工核对

### 3. pipeline

负责业务编排，而不是文件格式转换本身。

建议包含：

- `extract`: 从 PDF 抽页面元素
- `detect`: 识别题号、材料、科目、题型
- `select`: 按科目 / 题号筛选
- `normalize`: 形成标准化题目对象
- `review`: 提供人工修正入口
- `render_assets`: 渲染材料裁图和题干图

### 4. exporters

不同导出器只读取统一模型，不再自己做业务判断。

建议包含：

- `exporters/docx_booklet.py`
- `exporters/pptx_slides.py`
- `exporters/json_manifest.py`

其中：

- `docx_booklet` 负责题本
- `pptx_slides` 负责一页一题
- `json_manifest` 负责保存中间工程文件，便于复用和重导出

---

## 统一数据模型

建议把“题目”统一成以下两种节点：

### 普通题

- 原题号
- 所属科目
- 题干文本
- 题干图片
- 选项 A-D
- 选项文字 / 图片
- 答案（可选）
- 来源页码

### 资料分析组

- 组号 / 材料号
- 所属科目
- 材料文本
- 材料内图片
- 材料对应的 `PageRegion`
- 关联题目列表（通常 5 题）

资料分析子题：

- 原题号
- 题干文本
- 选项 A-D
- 题干附图（若有）

关键点：

- `材料图` 不应该从 Word 再取
- 应该直接从 PDF 的原始坐标裁图
- 每个资料分析组生成一个或多个材料图资源，供 5 道题重复复用

---

## 资料分析的正确处理方式

资料分析不要再把材料正文优先变成长文本后再进入 PPT。

建议改成双轨输出：

1. Word 题本：材料按文本 + 图片正常排版
2. PPT：材料优先使用 PDF 裁切后的图片

这样做的原因：

- 公考材料通常表格、图表、排版复杂
- 直接裁图比重排文本更稳定
- 五道题共用同一张材料图，业务上更符合授课场景

为此，PDF 抽取阶段必须保留：

- 页码
- 文本块 bbox
- 图片 bbox
- 题号和材料块的归属关系

后续通过“材料块所有 bbox 的并集”生成裁图区域。

---

## GUI 重做建议

GUI 不应再分成“Word -> PPT”和“PDF -> Word”两个几乎平级的功能页。

建议改成一个向导式流程：

1. `导入 PDF`
2. `识别设置`
3. `结果预览`
4. `导出 Word`
5. `生成 PPT`

### 第一步：导入 PDF

- 选择 PDF
- 自动识别试卷名称
- 自动建立工程目录

### 第二步：识别设置

- 处理范围：全部 / 按科目 / 按题号
- 科目：数量关系 / 资料分析 / 判断推理 / 言语理解等
- 题号范围：如 `66-85`、`111-125`
- 是否启用 OCR 兜底

### 第三步：结果预览

- 左侧题号树
- 右侧题目详情
- 可修正题号归属、材料分组、漏识别选项

### 第四步：导出 Word

- 选择题本模板
- 是否保留原题号
- 是否插入来源信息

### 第五步：生成 PPT

- 选择 PPT 模板
- 一页一题
- 资料分析自动复用材料图

---

## 与 MarkItDown 的借鉴关系

MarkItDown 值得借鉴的是架构思路，不是最终输出格式。

可借鉴点：

- 文件输入统一走转换器接口
- 转换过程可插拔
- 支持流式处理，不依赖大量临时文件
- 可生成中间检查结果，方便调试

不建议照搬的点：

- 本项目的标准中间格式不应是 Markdown
- 我们需要保留题号、材料分组、图片资产、页面坐标
- 这些信息比 Markdown 层级更关键

因此更合适的做法是：

- 参考 MarkItDown 的“转换器架构”
- 但本项目自己的中间标准应为 `ExamProject JSON`
- Markdown 只作为调试输出，不作为主数据源

---

## 推荐目录结构

```text
PPTconvert/
├── app/
│   ├── services/
│   ├── workflows/
│   └── viewmodels/
├── domain/
│   ├── models.py
│   ├── enums.py
│   └── selectors.py
├── ingest/
│   ├── pdf/
│   │   ├── extractor.py
│   │   ├── layout_models.py
│   │   ├── detector.py
│   │   ├── assembler.py
│   │   └── cropper.py
│   ├── docx/
│   │   └── importer.py
│   └── debug_markdown/
│       └── exporter.py
├── exporters/
│   ├── docx_booklet.py
│   ├── pptx_slides.py
│   └── manifest_json.py
├── workflows/
│   ├── import_pdf.py
│   ├── build_booklet.py
│   └── build_ppt.py
├── gui/
│   ├── wizard.py
│   ├── pages/
│   └── dialogs/
├── tests/
│   ├── domain/
│   ├── ingest/
│   ├── exporters/
│   └── workflows/
└── main.py
```

---

## 分阶段重构路线

### Phase 1：建立统一模型

先不动 GUI，大改数据层。

任务：

- 新建 `domain` 模型
- 让 `PDF -> 统一模型`
- 让 `Word -> 统一模型` 作为兼容适配器
- 新增 `json manifest` 导出

完成标准：

- 同一份 `ExamProject` 能生成 Word
- 不再需要先写 Word 再重新解析

### Phase 2：PPT 改为直接读取统一模型

任务：

- 重写 `PPTGenerator` 输入接口
- 普通题直接读 `Question`
- 资料分析题读 `MaterialSet + Question`
- 增加材料 PDF 裁图能力

完成标准：

- `PDF -> 统一模型 -> PPT`
- 资料分析 5 题共用同一材料图

### Phase 3：GUI 改成向导式

任务：

- 重做界面流程
- 增加选择范围和预览修正
- 统一导出入口

完成标准：

- 用户只需导入一次 PDF
- 后续 Word/PPT 都从同一工程导出

### Phase 4：规则增强与插件化

任务：

- 支持更多科目识别
- OCR 兜底
- 调试导出 Markdown
- 模板系统和导出器解耦

---

## 我建议的实际开工顺序

如果现在开始动手，不要先改 GUI，也不要先改模板。

正确顺序：

1. 建立统一领域模型
2. 重写 PDF 导入为统一模型
3. 让 Word 和 PPT 都读取统一模型
4. 最后再改 GUI

原因很简单：

- 你的核心问题不是界面
- 是主数据流设计错了
- 数据流不重做，界面改再漂亮都只是包旧逻辑

---

## 本轮重构的明确结论

这个项目应该从“文档转换工具”改成“试卷整理工作流工具”。

产品核心不再是：

- Word 转 PPT

而应该是：

- PDF 试卷导入
- 结构化摘题
- 按科目 / 题号筛选
- 输出题本 Word
- 输出授课 PPT

统一模型、统一工程、统一导出，才是这次重构的中心。
