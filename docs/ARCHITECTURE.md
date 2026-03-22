# 项目结构说明

```
PPTconvert/
├── main.py                 # 入口：无参数启动 GUI；带 -i 等参数走命令行
├── requirements.txt        # 运行时依赖
├── requirements-build.txt  # 仅打包 exe 时依赖（PyInstaller）
├── PPTconvert.spec         # PyInstaller 配置
├── build_exe.bat           # Windows 一键打包脚本
├── core/                   # 核心逻辑
│   ├── word_parser.py      # 解析 .docx 题目结构、表格、图片
│   ├── word_math.py        # Word OMML 公式 → 文本近似
│   ├── ppt_generator.py    # 生成 .pptx 幻灯片
│   ├── ppt_style.py        # 颜色等样式辅助
│   ├── template_manager.py # 模板幻灯片解析与占位
│   ├── template_style.py   # 从模板读取字体/对齐等
│   ├── image_extractor.py  # 题目配图提取与临时文件
│   └── models.py           # 题目等数据模型
└── gui/                    # 图形界面（ttkbootstrap）
    ├── app.py              # 主窗口与交互流程
    ├── ui_constants.py     # 标题、主题名等常量
    └── font_data.py        # 字体列表构建
```

## 数据流（简要）

1. **Word** → `WordParser` 读段落/表格 → `Question` 列表（含公式文本、图片引用）。
2. **配置** → `PPTConfig`（边距、选项布局、字体颜色等）；若使用模板则样式以模板解析结果为主。
3. **生成** → `PPTGenerator` 写入 `python-pptx` 幻灯片，必要时插入图片。

## 扩展时建议修改位置

- 解析规则 / 题号识别：`core/word_parser.py`
- 版式与占位：`core/ppt_generator.py`、`core/template_manager.py`
- 界面文案与主题：`gui/ui_constants.py`、`gui/app.py`
