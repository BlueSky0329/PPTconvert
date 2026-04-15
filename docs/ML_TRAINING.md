# 本地学习模型训练方案

## 目标

我们把本地 AI 拆成两层：

1. 监督学习分类器  
   直接学习“整道题属于哪个科目”，作为 `core/subject_inference.py` 的第二判断源。
2. 修复策略模型  
   学习“遇到串题、空选项、图片挂错时应该采取什么修复动作”。这一层先做离线模仿学习，后续再接离线强化学习。

当前仓库已经落地了第一层的训练与运行时接入：

- 金标准 PDF 目录：[gold_pdf_catalog.json](/C:/Users/17679/Desktop/PPTconvert/data/gold_pdf_catalog.json)
- 训练集构建脚本：[build_gold_subject_dataset.py](/C:/Users/17679/Desktop/PPTconvert/scripts/build_gold_subject_dataset.py)
- 训练脚本：[train_local_subject_model.py](/C:/Users/17679/Desktop/PPTconvert/scripts/train_local_subject_model.py)
- 运行时模型加载器：[learned_subject_model.py](/C:/Users/17679/Desktop/PPTconvert/core/learned_subject_model.py)

同时，第二层“修复动作学习”的数据准备也已经接好：

- 修复动作数据集构建脚本：[build_repair_action_dataset.py](/C:/Users/17679/Desktop/PPTconvert/scripts/build_repair_action_dataset.py)
- 默认输出：[repair_actions.jsonl](/C:/Users/17679/Desktop/PPTconvert/data/datasets/repair_actions.jsonl)
- 摘要输出：[repair_actions.summary.json](/C:/Users/17679/Desktop/PPTconvert/data/datasets/repair_actions.summary.json)
- 修复动作训练脚本：[train_repair_action_model.py](/C:/Users/17679/Desktop/PPTconvert/scripts/train_repair_action_model.py)
- 运行时策略加载器：[repair_action_model.py](/C:/Users/17679/Desktop/PPTconvert/core/repair_action_model.py)

另外，针对 PDF 页眉 / 页脚 / 页码 / 广告语灰区，当前仓库也已经补上了一个轻量文本噪声分类器：

- 数据集构建脚本：[build_pdf_noise_text_dataset.py](/C:/Users/17679/Desktop/PPTconvert/scripts/build_pdf_noise_text_dataset.py)
- 训练脚本：[train_pdf_noise_text_model.py](/C:/Users/17679/Desktop/PPTconvert/scripts/train_pdf_noise_text_model.py)
- 运行时加载器：[pdf_noise_model.py](/C:/Users/17679/Desktop/PPTconvert/core/pdf_noise_model.py)
- 抽取层接入点：[pdf_exam_extract.py](/C:/Users/17679/Desktop/PPTconvert/core/pdf_exam_extract.py)

同时，针对广告图 / 二维码 / 顶部横幅图 / 整页背景图 / 淡水印这类图片灰区，当前仓库也已经补上了一个基于 PyTorch 的轻量视觉分类器：

- 数据集构建脚本：[build_pdf_noise_image_dataset.py](/C:/Users/17679/Desktop/PPTconvert/scripts/build_pdf_noise_image_dataset.py)
- 训练脚本：[train_pdf_noise_image_model.py](/C:/Users/17679/Desktop/PPTconvert/scripts/train_pdf_noise_image_model.py)
- 运行时加载器：[pdf_noise_image_model.py](/C:/Users/17679/Desktop/PPTconvert/core/pdf_noise_image_model.py)
- 抽取层接入点：[pdf_exam_extract.py](/C:/Users/17679/Desktop/PPTconvert/core/pdf_exam_extract.py)

## 当前基线结果

截至 `2026-04-15`，本地科目分类模型已经按“千题库优先、套卷补缺”重训过一轮：

- 金标准 PDF：`7` 份
- 样本总数：`6338`
- 科目分布：
  - `politics`: `479`
  - `common_sense`: `25`
  - `verbal`: `2071`
  - `quant`: `897`
  - `reasoning`: `1908`
  - `data`: `958`
- 当前分组评估方式：按 `source_pdf` 做 grouped split，并保留每个标签的最低训练支持量
- 当前 baseline `macro_f1`：`0.7128`
- 当前运行时状态：`ready_for_runtime = true`
- 当前训练偏好：优先吸收 `single_subject_book`，并用套卷样本回填 `common_sense` 与跨科边界

这个结果说明两件事：

1. 训练链路已经打通，数据集、训练、序列化、运行时融合都可用。
2. 模型已经达到仓库当前运行阈值，可以作为规则引擎的第二判断源参与线上推断。

所以当前仓库采用的策略是：

- 继续保留规则引擎为主
- 训练模型作为加权融合信号，而不是直接替代规则判断
- 只有评估达标的模型才允许自动并入运行时推断

## PDF 噪声学习现状

截至 `2026-04-15`，PDF 文本噪声模型已经按“候选灰区样本”重训过一轮：

- 数据集：`13078` 条候选文本块
- 标签分布：
  - `noise`: `6211`
  - `content`: `6867`
- 当前分组评估方式：按 `source_pdf` 做 grouped split
- 当前 baseline `macro_f1`：`0.8573`
- `noise precision`：`0.9867`
- `noise recall`：`0.7420`
- 当前运行时状态：`ready_for_runtime = true`

这里的关键点不是“让模型替代规则”，而是：

- 规则先过滤明确广告语、页码、页眉页脚噪声
- 学习模型只看候选灰区
- 评估阈值优先看 `noise precision`，避免误删正文

所以这条模型的定位是“高精度补杀灰区噪声”，不是全文本去噪器。

## PDF 图片噪声学习现状

截至 `2026-04-15`，PDF 图片噪声模型已经在支持 CUDA 的本地环境里重训过一轮：

- 数据集：`4444` 个候选图片块
- 标签分布：
  - `noise`: `3302`
  - `content`: `1142`
- 新增弱标注来源：
  - `rule_background_image`: `31`
- 当前分组评估方式：按 `source_pdf` 做 grouped split
- 当前训练设备：`cuda`
- 当前 baseline `macro_f1`：`0.9941`
- `noise precision`：`0.9981`
- `noise recall`：`0.9924`
- 当前运行时状态：`ready_for_runtime = true`

这里的接法和文本噪声层保持一致：

- 规则先过滤明显的顶部横幅、小角标、二维码、logo
- 对整页低方差黑底 / 白底背景图，再额外走一层视觉统计规则
- 对超浅整页水印图，也会先走一层保守规则过滤
- 学习模型只看候选灰区图片
- 评估阈值优先压住误删风险，避免把题图或资料分析图表误当装饰图删掉

所以这条模型的定位是“对疑似广告图 / 二维码 / 横幅图 / 背景图做高精度补杀”，不是通用视觉理解器。

另外，图片噪声数据集构建脚本现在已经改成 staging 原子替换：

- 构建成功之前，不会覆盖旧的 `jsonl`、`summary` 和资产目录
- 这样即使缺少 `PyMuPDF` 或中途报错，也不会把上一版可训练数据破坏掉

## 为什么先做监督学习

你现在手里的 PDF 是“金标准分类语料”，这意味着最先该做的是监督学习，而不是直接上强化学习。

原因很简单：

- 监督学习天然适合“题目 -> 正确科目”这种有标准答案的问题。
- 现在最缺的是稳定分类，不是探索式策略。
- 强化学习更适合“多步修复动作链”，比如先判断串题，再回收图片，再补题号。

所以这条线应该分两步：

1. 先用金标准 PDF 把题目分类模型训起来
2. 再用人工修改日志和合成脏样本训练修复策略

## 当前语料组织

当前训练目录里我们明确区分了两种 PDF：

- `set_paper`：完整套卷，比如 `模拟卷十一.pdf`
- `single_subject_book`：单科题库，比如 `行测——判断推理（2000题）.pdf`

这个区分会直接影响训练样本生成：

- 套卷题目标签来自解析后的 section kind
- 单科题库的标签直接由目录清单指定

这样可以避免把“单科题库”错误地当成标准套卷去学习。

对于 `politics / common_sense`，当前还有一个明确的时点差异：

- `2025` 年前，不少行测资料里政治理论与常识判断仍会混编或共享表述
- `2025` 年后，政治理论更集中在 `习近平新时代中国特色社会主义思想 / 党的创新理论 / 政策文件`
- 当前仓库把 [政治理论题本.pdf](/C:/Users/17679/Desktop/PPTconvert/【02】最新行测5000题题库/政治理论题本.pdf) 作为 `single_subject_book -> politics`
- `法律制度 / 自然科学 / 人文科技常识` 仍优先归到 `common_sense`

## 训练步骤

### 1. 构建训练集

```powershell
python .\\scripts\\build_gold_subject_dataset.py
```

默认输出：

- [subject_gold.jsonl](/C:/Users/17679/Desktop/PPTconvert/data/datasets/subject_gold.jsonl)
- [subject_gold.summary.json](/C:/Users/17679/Desktop/PPTconvert/data/datasets/subject_gold.summary.json)

每条样本包含：

- `subject`
- `stem`
- `options`
- `material_header`
- `material_text`
- `image_count`
- `source_pdf`
- `feature_record`

### 2. 安装训练依赖

```powershell
pip install -r .\\requirements-ml.txt
```

### 3. 训练本地分类模型

```powershell
python .\\scripts\\train_local_subject_model.py
```

默认输出：

- `data/models/subject_classifier.pkl`
- `data/models/subject_classifier.metrics.json`

模型会被 [subject_inference.py](/C:/Users/17679/Desktop/PPTconvert/core/subject_inference.py) 检查。

- 默认情况下，只要 [subject_classifier.pkl](/C:/Users/17679/Desktop/PPTconvert/data/models/subject_classifier.pkl) 存在且 `ready_for_runtime = true`，它就会自动参与规则融合
- 如果指标还不够，会保留为候选模型，不会自动污染线上判断
- 如需显式关闭默认模型，可设置 `PPTCONVERT_ENABLE_PICKLED_SUBJECT_MODEL=0`
- 如需试用自定义模型路径，可设置 `PPTCONVERT_ENABLE_PICKLED_SUBJECT_MODEL=1` 并配合 `PPTCONVERT_SUBJECT_MODEL`
- 如需强制试用尚未达标的模型，可设置环境变量 `PPTCONVERT_FORCE_SUBJECT_MODEL=1`

## 当前模型结构

当前不是大模型微调，而是一个很适合这个项目的本地轻量分类器：

- 文本特征：题干、选项、材料头、材料正文
- 结构特征：图片数、数字密度、选项数量、空格填空结构、长文段长度
- 训练器：线性分类器

这样做的好处是：

- 本地可训练
- 小样本也能起效果
- 很容易和规则引擎融合
- 出错时可解释

## 修复动作数据准备

这一步不是直接训练，而是把“正确工程 -> 合成脏状态 -> 正确修复动作”固化成离线样本。

### 构建修复动作数据集

```powershell
python .\\scripts\\build_repair_action_dataset.py
```

默认输出：

- [repair_actions.jsonl](/C:/Users/17679/Desktop/PPTconvert/data/datasets/repair_actions.jsonl)
- [repair_actions.summary.json](/C:/Users/17679/Desktop/PPTconvert/data/datasets/repair_actions.summary.json)

当前脚本已经会合成这些动作样本：

- `renumber_current_question`
- `split_embedded_next_question`
- `move_spilled_option_back`
- `move_data_intro_back_to_material`
- `move_data_assets_to_material`
- `reassign_stem_image_to_options`

每条样本包含：

- `action`
- `action_family`
- `corruption`
- `source_pdf`
- `subject`
- `state_record`
- `target`

其中 `state_record` 已经带上：

- 当前题
- 前一题 / 后一题
- 材料正文
- 基础结构统计特征

这意味着下一步如果做模仿学习，我们已经有统一的状态表示，不需要重新定义训练输入。

### 训练修复动作模型

```powershell
python .\\scripts\\train_repair_action_model.py
```

默认输出：

- `data/models/repair_action_classifier.pkl`
- `data/models/repair_action_classifier.metrics.json`

这个模型当前的定位不是直接自动接管 GUI，而是先用来评估“结构修复动作是否可学”。输入是 `state_record`，标签是 `action`。

如果想先训练更稳定的父类动作模型，可以直接：

```powershell
python .\\scripts\\train_repair_action_model.py --label-field action_family
```

这会先学习三大类：

- `boundary_repair`
- `material_repair`
- `asset_repair`

截至 `2026-04-15`，这两个 repair 头已经在本地重训过一轮：

- `action` 头 `macro_f1`：`0.8461`
- `action_family` 头 `macro_f1`：`0.7928`
- 当前运行时状态：两个 bundle 都已经 `ready_for_runtime = true`

### 当前运行时接法

修复策略模型现在已经能被运行时读取，但仍然保持“规则优先、策略只做软参与”的接法。

- 默认情况下，`AIRepairService` 仍然先跑确定性的边界修复和本地规则补丁
- 只有在 `policy` 模式下，learned repair model 才会参与 flagged 题排序，并把“更像哪类修复动作”的提示追加到 patch 摘要
- 当前仓库里的默认 repair bundles 已经达到 `ready_for_runtime = true`
- 默认仍然不会在 `balanced` 模式下抢占规则修复；只有切到 `policy` 模式才会参与排序和提示
- 如果想显式开启 `policy` 模式，可以配合环境变量：

```powershell
$env:PPTCONVERT_AI_REPAIR_MODE='policy'
python .\\main.py --pdf-input exam.pdf --ai-repair
```

如果想指定自定义模型路径，可以再补：

```powershell
$env:PPTCONVERT_REPAIR_FAMILY_MODEL='C:\\path\\to\\repair_action_family_classifier.pkl'
$env:PPTCONVERT_REPAIR_MODEL='C:\\path\\to\\repair_action_classifier.pkl'
```

这套接法的关键原则是：

- learned model 只做排序、提示、加权
- 不让 learned model 直接绕过规则修复去“拍脑袋改题”
- 如需关闭默认 repair bundle，可设置 `PPTCONVERT_ENABLE_PICKLED_REPAIR_MODEL=0`
- 如需强制试用未达标的自定义 repair bundle，可设置 `PPTCONVERT_FORCE_REPAIR_MODEL=1`

### 离线模仿学习轨迹管线

截至 `2026-04-15`，多步修复轨迹这条线已经补齐了“数据集构建 -> 训练脚本 -> 运行时加载 -> GUI/CLI 展示”的完整管线：

- 数据集脚本：[build_repair_trajectory_dataset.py](/C:/Users/17679/Desktop/PPTconvert/scripts/build_repair_trajectory_dataset.py)
- 训练脚本：[train_repair_trajectory_model.py](/C:/Users/17679/Desktop/PPTconvert/scripts/train_repair_trajectory_model.py)
- 运行时加载器：[repair_trajectory_model.py](/C:/Users/17679/Desktop/PPTconvert/core/repair_trajectory_model.py)
- 运行时接线：[ai_repair.py](/C:/Users/17679/Desktop/PPTconvert/core/ai_repair.py)

当前真实状态是：

- 轨迹策略已经能参与 `policy` 模式下的排序提示和轨迹解释
- 没有 learned trajectory bundle 时，会自动退回规则轨迹兜底
- 数据集脚本现在支持 `--gui-manifests`，可以把工程 manifest 里的真实 GUI 修题日志并入 trajectory 数据集
- 数据集脚本现在也支持 `--gui-jsonl`，可以直接读取 [gui_repair_logs.jsonl](/C:/Users/17679/Desktop/PPTconvert/data/datasets/gui_repair_logs.jsonl) 这类导出日志文件
- 当前仓库里已经产出默认 `repair_trajectory_policy.pkl`
- 当前 trajectory 数据集为 `20359` 条，来源仍然全部是 synthetic；当前仓库里还没有可并入的真实 GUI manifest
- 默认训练脚本现在会在大样本数据上自动切到更快的 `SGDClassifier`；`--rebalance-train` 作为可选实验开关保留
- 当前这一版 trajectory bundle 的结果是：
  - `split_strategy = grouped_by_pdf_partial_labels`
  - `macro_f1 = 0.1303`
  - `ready_for_runtime = false`
- GUI 里的真实人工修题日志现在会写入工程 manifest，并可用 [build_gui_repair_log_dataset.py](/C:/Users/17679/Desktop/PPTconvert/scripts/build_gui_repair_log_dataset.py) 汇总成 JSONL
- 所以“离线模仿学习管线”已经完成，而且默认 bundle 也训出来了，但“trajectory 模型达标上线”还没有完成

## 强化学习该怎么接

强化学习不建议直接拿“题目分类”来做。更合适的对象是“修复动作策略”。

### 推荐的动作空间

- `merge_numeric_continuation`
- `move_spilled_option_back`
- `split_embedded_next_question`
- `reassign_stem_image_to_options`
- `move_data_intro_back_to_material`
- `renumber_current_question`

### 状态表示

状态不是原始 PDF，而是当前工程里的结构化题目节点：

- 当前题干
- 当前选项
- 前后题邻域
- 图片挂载位置
- 材料上下文
- 当前质检问题列表

### 奖励设计

建议先做离线奖励，不直接在线探索：

- 修复后题号连续：`+1`
- 修复后空选项消失：`+1`
- 修复后图片挂载正确：`+1`
- 修复后质检项减少：`+1`
- 引入新错误：`-2`
- 把原本正确题改坏：`-3`

### 数据从哪里来

第一阶段不要等人工日志，先用“合成脏样本”做离线模仿学习。

这条线现在已经实现成脚本和运行时骨架了，当前合成错误主要覆盖：

1. 从金标准项目出发
2. 人为注入常见错误：
   - 数字续行断裂
   - D 选项串到下一题
   - 题干吞掉下一题号
   - 选项图挂到题干
   - 材料图挂到首题
3. 把“脏状态 -> 正确动作 -> 修复后状态”记录成离线轨迹

第二阶段现在已经开始：GUI 里真实人工修改会写进工程日志，能够沉淀成高质量策略数据；下一步就是把这些真实 manifest 真正喂给 trajectory 训练。

## 推荐训练路线

### 第一阶段：监督学习

- 目标：提高 `politics/common_sense/verbal/quant/reasoning/data` 分类稳定性
- 数据：当前金标准 PDF
- 指标：按 `source_pdf` 分组的宏平均 F1
- 当前现实问题：
  - `common_sense` 样本仍然严重不足
  - `politics` 虽然已经补进单科题本，但跨年份边界语料仍不够
  - 所以下一步应该先补金标准语料，再继续提模

### 第二阶段：离线模仿学习

- 目标：预测修复动作
- 数据：合成脏样本 + 人工修题日志
- 指标：动作准确率、修复后质检下降量
- 当前状态：
  - 单步 `action / action_family` policy 已经 `ready_for_runtime = true`
  - 多步 trajectory policy 的数据集 / 训练 / 运行时管线已打通，默认 bundle 也已训练完成
- 真实 GUI 修题日志已经能导出成 JSONL，下一步最值当的是把它接进 trajectory 训练，而不是继续堆合成规则
  - 现在这条接线已经支持两种入口：原工程 manifest 或导出的 `gui_repair_logs.jsonl`

### 第三阶段：离线强化学习

- 目标：多步修复策略
- 数据：修复轨迹
- 奖励：题号连续、空选项减少、图片归位、质检问题减少
- 当前状态：
  - 还没有开始真实训练
  - 当前仍停留在奖励设计和数据沉淀阶段

## 参考方法

这条路线不是拍脑袋的，底层参考的是文档理解和策略学习里更适合本项目的方向：

- [LayoutLMv3](https://arxiv.org/abs/2204.08387)
- [Donut](https://arxiv.org/abs/2111.15664)
- [DAgger](https://arxiv.org/abs/1011.0686)
- [Offline Reinforcement Learning: Tutorial, Review, and Perspectives on Open Problems](https://arxiv.org/abs/2005.01643)

对这个项目来说，最稳的原则是：

- 分类先做监督学习
- 修复先做模仿学习
- 强化学习最后再接，而且只用于“结构修复动作”

这样我们不会把一个本来已经能用的工程，变成一个黑箱且不可控的系统。
