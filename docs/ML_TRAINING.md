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

## 当前基线结果

截至 `2026-04-14`，第一版本地监督学习基线已经实际跑通：

- 金标准 PDF：`6` 份
- 样本总数：`5947`
- 科目分布：
  - `politics`: `35`
  - `common_sense`: `25`
  - `verbal`: `2071`
  - `quant`: `897`
  - `reasoning`: `1908`
  - `data`: `1011`
- 当前分组评估方式：按 `source_pdf` 做 grouped split
- 当前 baseline `macro_f1`：`0.1369`
- 当前运行时状态：`ready_for_runtime = false`

这个结果说明两件事：

1. 训练链路已经打通，数据集、训练、序列化、运行时 gated 加载都可用。
2. 当前模型还不够强，不能直接替代规则引擎。

所以当前仓库采用的策略是：

- 继续保留规则引擎为主
- 训练模型只作为候选能力
- 只有评估达标的模型才允许自动并入运行时推断

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

- 如果评估指标达到运行阈值，会自动参与规则融合
- 如果指标还不够，会保留为候选模型，不会自动污染线上判断
- 如需强制试用，可设置环境变量 `PPTCONVERT_FORCE_SUBJECT_MODEL=1`

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

第一阶段不要等人工日志，先用“合成脏样本”做离线模仿学习：

1. 从金标准项目出发
2. 人为注入常见错误：
   - 数字续行断裂
   - D 选项串到下一题
   - 题干吞掉下一题号
   - 选项图挂到题干
   - 材料图挂到首题
3. 把“脏状态 -> 正确动作 -> 修复后状态”记录成离线轨迹

第二阶段再把 GUI 里真实人工修改记录下来，作为高质量策略数据。

## 推荐训练路线

### 第一阶段：监督学习

- 目标：提高 `politics/common_sense/verbal/quant/reasoning/data` 分类稳定性
- 数据：当前金标准 PDF
- 指标：按 `source_pdf` 分组的宏平均 F1
- 当前现实问题：
  - `politics / common_sense` 样本严重不足
  - 当前训练集主要来自 `verbal / reasoning / quant / data` 单科题库
  - 所以下一步应该先补金标准语料，再继续提模

### 第二阶段：离线模仿学习

- 目标：预测修复动作
- 数据：合成脏样本 + 人工修题日志
- 指标：动作准确率、修复后质检下降量

### 第三阶段：离线强化学习

- 目标：多步修复策略
- 数据：修复轨迹
- 奖励：题号连续、空选项减少、图片归位、质检问题减少

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
