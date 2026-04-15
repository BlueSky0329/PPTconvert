# 当前进度

更新时间：2026-04-15

## 当前定位

项目现在是一个面向公考试题整理与授课输出的桌面工具，不再只是早期的 `Word -> PPT` 小工具。

当前主链路已经变成两条共享工程流：

- `PDF -> 抽取 / 解析 / 科目推断 -> ExamProject -> 共享预览 / 编辑 -> Word / PPT / JSON`
- `Word -> 解析 / 科目推断 -> ExamProject -> 共享预览 / 编辑 -> PPT / JSON`

GUI 当前保留双入口：

- `PDF 试卷整理`
- `Word 生成 PPT`

并共用一套 `PPT 导出设置` 与预览编辑区。

## 本轮整理重点

### 1. 恢复并统一 Word 工作流

- GUI 中重新接回 `Word 生成 PPT`
- `Word` 不再直接跳过预览，而是先转成统一的 `ExamProject`
- `Word` 与 `PDF` 现在共用同一套结构树、题目编辑、材料预览、导出设置

### 2. 增强无标题 / 单科场景鲁棒性

- 新增整份文档科目提示：`自动识别 / 政治理论 / 常识判断 / 言语理解与表达 / 数量关系 / 判断推理 / 资料分析`
- 新增启发式科目推断器，覆盖 PDF 和 Word 两条链路
- 支持：
  - 单科整卷
  - 缺少部分篇题标题
  - 完全无大标题但仍有题号与选项结构
- PDF 文件名会先做 profile 判断：
  - `set_paper`
  - `single_subject_book`
  - `unknown`
- 套卷与单科题库已经分成两套无标题策略，不再用固定题号范围硬编码整卷结构
- 低置信度内容不再直接丢弃，而是落到 `unknown` / 待确认科目

### 3. 共享预览与人工修题继续完善

- 题干实时修改
- 选项文字编辑
- 选项顺序调整、插入、删除
- 选项图片查看、替换、清除、从原 PDF 区域重裁
- 单题选项布局覆盖
- 题干图片预览
- 资料分析材料原貌预览
- 工程 JSON 回载继续编辑
- 未保存修改保护

### 4. 导出链路补齐

- `ExamProject -> Word / PPT / JSON` 保持统一
- Word 导出已补齐图片选项的选项块输出
- 资料分析材料截图逻辑由 Word / PPT / GUI 共用

### 5. 本地 AI 质检、修复与学习框架

- 已接入本地 AI 质检层：
  - `review_confidence`
  - `review_issues`
  - `suggested_subject`
- 已接入本地 AI 修复层：
  - 单题清洗
  - 题号回正
  - 串题边界回拆
  - 题干/选项图片重挂
  - 资料分析材料回收
  - `policy` 模式下的 learned repair policy 排序与提示
  - 离线模仿学习轨迹策略管线：
    - `repair_trajectories.jsonl` 数据集构建脚本
    - `repair_trajectory_policy.pkl` 训练脚本
    - 运行时轨迹策略加载与规则兜底
    - 可选 `--gui-manifests` / `--gui-jsonl` 把工程 manifest 或导出的 GUI 日志并入 trajectory 数据集
  - GUI 已支持：
    - `balanced / policy` 修复模式切换
    - 单题修复策略 / 轨迹解释可视化
    - 人工题干 / 选项 / 布局 / 图片编辑日志写入工程 manifest
    - `build_gui_repair_log_dataset.py` 可把工程日志汇总成真实修题数据集
  - repair `action / action_family` 模型已达 `ready_for_runtime = true`
- 已接入扫描/OCR PDF 诊断与自动修补工具：
  - core 层可输出扫描/OCR 风险报告
  - CLI 已支持：
    - `--ocr-diagnose`
    - `--ocr-diagnose-json`
    - `--ocr-auto-repair`
  - GUI 已支持：
    - `扫描/OCR 诊断`
    - `OCR 自动修补`
- 已接入 PDF 文本噪声学习层：
  - 候选页眉 / 页脚 / 页码 / 广告语灰区补杀
  - 默认模型 `pdf_noise_text_classifier.pkl` 已达 `ready_for_runtime = true`
  - 运行时仍保持“规则优先，学习模型只补灰区”
- 已接入 PDF 图片噪声学习层：
  - 候选广告图 / 二维码 / 横幅图 / 整页背景图 / 淡水印灰区补杀
  - 默认模型 `pdf_noise_image_classifier.pt` 已达 `ready_for_runtime = true`
  - 在支持 CUDA 的环境里可直接用 GPU 重训
  - 运行时仍保持“规则优先，学习模型只补灰区”
  - 额外补了两层保守规则：
    - 大幅低方差背景图
    - 超浅整页水印图
  - 图片训练集构建已改成 staging 原子替换，避免中途失败把旧资产目录删空
- 已接入本地知识库：
  - [LOCAL_AI_KB.md](C:/Users/17679/Desktop/PPTconvert/docs/LOCAL_AI_KB.md)
  - [GONGKAO_TAXONOMY.md](C:/Users/17679/Desktop/PPTconvert/docs/GONGKAO_TAXONOMY.md)
  - [GONGKAO_CORPUS.md](C:/Users/17679/Desktop/PPTconvert/docs/GONGKAO_CORPUS.md)
- 已搭建监督学习训练框架：
  - 金标准 PDF 目录
  - 训练集构建脚本
  - 本地轻量分类模型训练脚本
  - 运行时 gated 融合逻辑

## 当前能力概览

### 解析覆盖

已支持：

- 政治理论
- 常识判断
- 言语理解与表达
- 数量关系
- 判断推理
- 资料分析
- 待确认科目（兜底容器）

### 已处理的高频边界

- `66. 题干` 同行切分
- `D` 选项续行
- `D` 选项后直接跟下一题
- 资料分析无“材料二/三/四”时的自动拆组
- 页面图片缺失补提取
- 双栏页误判
- 客观题篇首说明误并入首题
- 标题中途缺失导致的科目漂移
- 无标题整卷的单科推断
- `30 余颗卫星...` 这类数量单位正文误拆成假题号
- `图1 A. 图2 B. 图3 C. 图4 D.` 这类图片先于标签的图片选项
- 页眉广告、二维码、扫码语、页脚页码误入题目流
- 题干中的数字续行、碎片选项、纯图片选项回挂

## 当前学习语料与训练状态

- 金标准 PDF：`7` 份
- 训练样本：`6338` 题
- 当前样本分布：
  - `politics`: `479`
  - `common_sense`: `25`
  - `verbal`: `2071`
  - `quant`: `897`
  - `reasoning`: `1908`
  - `data`: `958`
- 已新增 `【02】最新行测5000题题库/政治理论题本.pdf`，作为 `single_subject_book -> politics` 金标准语料
- 对 `2025` 年前后口径的当前约定：
  - `2025` 年前不少行测语料里，政治理论和常识判断仍存在混编或交叉表述
  - 当前项目里，`习近平新时代中国特色社会主义思想 / 党的创新理论 / 政策文件` 更偏向 `politics`
  - `法律制度 / 自然科学 / 人文科技常识` 更偏向 `common_sense`
- 当前已训练出一版本地科目分类模型，并按 `single_subject_book` 优先重训
- 当前分组评估方式仍是 `grouped-by-pdf`，但会保留每个标签的最低训练支持量，避免唯一主来源整本被挪到测试侧
- 目前 grouped-by-pdf `macro_f1` 为 `0.7128`
- 当前模型状态：`ready_for_runtime = true`
- 含义：
  - 训练链路已经打通
  - 模型已经达到当前运行阈值
  - 默认模型会自动参与规则融合，但规则引擎仍是主判断源
- PDF 文本噪声模型：
  - 数据集：`13078` 条候选文本块
  - 指标：`macro_f1 = 0.8573`，`noise precision = 0.9867`，`noise recall = 0.7420`
  - 当前模型状态：`ready_for_runtime = true`
- PDF 图片噪声模型：
  - 数据集：`4444` 个候选图片块
  - 标签分布：
    - `noise`: `3302`
    - `content`: `1142`
  - 弱标注来源新增：
    - `rule_background_image`: `31`
  - 在 `RTX 5090 + CUDA` 环境下重训结果：
    - `macro_f1 = 0.9941`
    - `noise precision = 0.9981`
    - `noise recall = 0.9924`
  - 当前模型状态：`ready_for_runtime = true`
  - 含义：
    - 学习模型已经能参与候选广告图 / 二维码 / 横幅图 / 整页背景图的高精度补杀
    - 规则层已经能额外挡住纯背景图和淡水印式整页图
    - 仍然不会替代已有规则去“按图片大小直接乱删”

## 当前已知边界

- 扫描版 PDF / OCR 错字仍可能造成错切
- 多科混排且没有标题、并且文本信号很弱时，仍以启发式推断为主
- 资料分析跨页材料仍有继续增强空间
- 极端复杂表格和多栏版式仍可能出现读序异常
- 共享预览已经足够用于人工修题，但还不是最终导出版面的完全等价渲染
- 当前本地学习模型的数据分布仍明显偏科：
  - `politics` 已明显补强，但 `common_sense` 仍然样本很少
  - 更适合继续扩 `common_sense` 金标准和跨年份套卷语料，再继续提模
- PDF 噪声学习层已经覆盖文本块与图片块灰区：
  - 文本块主要处理页眉 / 页脚 / 页码 / 广告语
  - 图片块主要处理广告图 / 二维码 / 横幅图 / 大幅低方差背景图 / 淡水印整页图
  - 当前仍保持“规则先挡住明显噪声，学习模型只处理灰区候选”的保守接法
  - 如果后续要继续增强，最值当的是补更难的淡水印、半透明章印和整页底纹纹理
- 离线模仿学习轨迹管线已经打通，并且已经产出一版默认 `repair_trajectory_policy.pkl`
- 截至 `2026-04-15`，trajectory 数据集为 `20359` 条，其中 GUI 真实日志贡献暂时还是 `0`
- 也就是说：入口已经支持真实 GUI 轨迹，但当前仓库里还没有可直接并入的真实日志文件
- 当前 trajectory bundle 评估为：
  - `split_strategy = grouped_by_pdf_partial_labels`
  - `macro_f1 = 0.1303`
  - `ready_for_runtime = false`
- 结论是：这条线已经从“没有模型”推进到“有默认 bundle、可继续迭代”，但还需要真实 GUI 修题轨迹来跨过上线阈值
- 强化学习目前仍停留在设计阶段，尚未开始训练真实修复策略模型

## 当前测试情况

截至本次整理，全量回归已通过：

```text
python -m unittest discover -s tests -v
```

当前测试结果：`289` 通过，`1` 个按环境条件跳过

回归覆盖重点：

- PDF 抽取与读序
- PDF 解析与切题
- 科目推断与无标题 fallback
- Word 解析
- `ParsedExam / WordQuestion -> ExamProject`
- Word 导出
- 项目编辑与答案同步
- GUI 相关的数据层逻辑
- 本地 AI 质检与修复
- 本地学习模型加载与 gated 融合
- 扫描/OCR 诊断与自动修补
- 离线模仿学习轨迹策略
- GUI 修题日志 roundtrip 与数据集导出

## 接手建议

1. 用更多真实套卷和单科题库继续扩充金标准 PDF
2. 优先补 `common_sense` 金标准样本，并继续补跨年份 `politics/common_sense` 套卷边界语料
3. 基于已沉淀的 GUI 修题日志，继续构建“修复动作 / 修复轨迹”真实训练集，优先补 trajectory bundle
4. 继续增强扫描/OCR PDF 的诊断与自动修补，重点补淡水印、章印、底纹和真实扫描版误切样本
