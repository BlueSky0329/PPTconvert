# GitHub 协作与发布

本文档面向当前这个已经存在远程仓库的项目，不再讨论“如何第一次创建 GitHub 仓库”。

## 当前仓库

- 远程：`origin = https://github.com/BlueSky0329/PPTconvert.git`
- 默认分支：`main`

## 提交前检查

提交前至少做两件事：

```powershell
python -m unittest discover -s tests -v
python -m py_compile .\main.py .\gui\app.py
```

如果本次修改集中在 PDF 解析，也建议补跑：

```powershell
python -m unittest tests.test_pdf_exam_extract tests.test_pdf_exam_parse tests.test_exam_project -v
```

## 仓库边界

以下内容默认不提交：

- 根目录下的真实 PDF
- 根目录下导出的 DOCX / PPTX
- `*_assets/`
- `*_工程.json`
- `outputs/`
- 本地缓存目录

如果要共享样例，请放到 `examples/`，并确认不含隐私与版权风险。

## 日常提交流程

```powershell
git status
git add .
git commit -m "feat: describe your change"
git push origin main
```

更稳妥的多人协作流程：

```powershell
git checkout -b feat/your-topic
git add .
git commit -m "feat: describe your change"
git push -u origin feat/your-topic
```

## 提交信息建议

- `feat:` 新功能
- `fix:` 缺陷修复
- `refactor:` 重构
- `docs:` 文档整理
- `test:` 测试补充
- `chore:` 纯维护性改动

示例：

- `feat: add pdf project workflow and gui wizard`
- `fix: recover data material intro spilled into previous option`
- `docs: rewrite architecture and progress docs`

## 推送前自查

推送前确认：

1. `git status` 中没有误加入的本地试卷或导出文件
2. 测试已通过
3. README 与行为一致
4. 大改动有对应测试

## 发布建议

适合打标签的节点：

- GUI 工作流稳定可演示
- 六大模块识别可用
- 一轮真实试卷回归通过

示例：

```powershell
git tag v0.3.0
git push origin v0.3.0
```

## 不建议直接做的事

- 不要把真实试卷 PDF 直接提交到根目录
- 不要把导出的 Word/PPT 当源码版本管理
- 不要在未跑测试的情况下直接推送解析规则改动
