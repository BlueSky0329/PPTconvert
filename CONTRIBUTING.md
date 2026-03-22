# 参与贡献

感谢你有意改进本项目。

## 开发环境

- Python **3.10+**（推荐 3.12）
- 建议使用虚拟环境：

```bash
python -m venv .venv
# Windows
.venv\Scripts\activate
pip install -r requirements.txt
python main.py
```

## 提交前建议

- 本地运行 GUI 或命令行做一次完整转换流程自测。
- 避免将 **个人试卷、大体积 docx/pptx** 或 **打包产物**（`dist/`、`build/`、`.venv/`）提交进仓库；这些路径已在 `.gitignore` 中忽略。

## 代码风格

- 类型注解与清晰函数命名优先。
- 与现有模块职责划分保持一致：解析在 `core/`，界面在 `gui/`。

## 报告问题

提交 Issue 时请尽量说明：

- Python 版本、操作系统
- 输入 Word 的排版特点（是否表格、是否公式）
- 期望行为与实际行为
