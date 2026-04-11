# 本次整理记录

日期：2026-04-12

## 整理目标

本轮整理的重点不是单点修 bug，而是把仓库整理成“功能、文档、测试、提交状态”一致的可继续维护状态。

## 本次纳入的主要改动

### 1. 工作流层

- 恢复 `Word 生成 PPT` GUI 入口
- 将 `Word` 解析结果接入统一 `ExamProject`
- 让 `PDF` 与 `Word` 共用同一套预览、编辑、导出设置

### 2. 解析与鲁棒性

- 为 PDF 和 Word 增加整份文档科目提示
- 新增启发式科目推断器
- 支持单科整卷、无标题整卷、标题中途缺失
- 低置信度内容落到 `unknown`，避免静默丢题

### 3. 人工修题能力

- 题干实时编辑
- 选项编辑、移动、增删
- 选项图片替换、清除、PDF 重裁
- 单题选项布局覆盖
- 工程 JSON 回载继续编辑
- 未保存修改保护

### 4. 导出相关

- Word 导出补齐图片选项块输出
- 资料分析材料裁图逻辑在 GUI / Word / PPT 之间共享

## 文档整理

本次同步更新：

- [README.md](C:/Users/17679/Desktop/PPTconvert/README.md)
- [docs/STATUS.md](C:/Users/17679/Desktop/PPTconvert/docs/STATUS.md)
- [docs/ARCHITECTURE.md](C:/Users/17679/Desktop/PPTconvert/docs/ARCHITECTURE.md)
- [docs/GITHUB.md](C:/Users/17679/Desktop/PPTconvert/docs/GITHUB.md)

目标是让新接手者不用翻聊天记录，也能知道项目现在能做什么、主链路在哪、边界在哪。

## 清理约定

整理仓库时默认不提交：

- 根目录真实 PDF / DOCX / PPTX
- 根目录导出产物
- `*_assets/`
- `*_工程.json`
- 缓存目录与解释器缓存

## 验证基线

提交前基线命令：

```powershell
python -m unittest discover -s tests -v
python -m py_compile .\main.py .\gui\app.py
```

## 当前留存风险

- 多科无标题混排仍主要依赖启发式推断
- OCR 噪声与复杂版式依然会影响分段和切题
- GUI 预览仍然是校对工具，不是最终导出版面的完全仿真器
