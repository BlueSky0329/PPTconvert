import os
from typing import Optional

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN


# 默认幻灯片尺寸 (16:9)
DEFAULT_WIDTH = Inches(13.333)
DEFAULT_HEIGHT = Inches(7.5)


class TemplateManager:
    """PPT 模板管理器"""

    def __init__(self):
        self._template_path: Optional[str] = None
        self._prs: Optional[Presentation] = None

    def load_template(self, template_path: str) -> Presentation:
        """加载用户指定的 PPT 模板"""
        if not os.path.exists(template_path):
            raise FileNotFoundError(f"模板文件不存在: {template_path}")
        self._template_path = template_path
        self._prs = Presentation(template_path)
        return self._prs

    def create_default(self) -> Presentation:
        """创建默认空白演示文稿 (16:9)"""
        self._prs = Presentation()
        self._prs.slide_width = DEFAULT_WIDTH
        self._prs.slide_height = DEFAULT_HEIGHT
        self._template_path = None
        return self._prs

    def get_presentation(self) -> Presentation:
        if self._prs is None:
            return self.create_default()
        return self._prs

    def get_blank_layout(self) -> object:
        """获取空白版式"""
        prs = self.get_presentation()
        # 优先寻找名为 "Blank" 或空白的版式
        for layout in prs.slide_layouts:
            if layout.name.lower() in ('blank', '空白', '空白版式'):
                return layout
        # 退而求其次用最后一个版式（通常是空白）
        return prs.slide_layouts[-1]

    def get_layout_names(self) -> list[str]:
        """返回模板中所有可用版式名称"""
        prs = self.get_presentation()
        return [layout.name for layout in prs.slide_layouts]

    @property
    def is_custom_template(self) -> bool:
        return self._template_path is not None

    @property
    def slide_width(self) -> int:
        return self.get_presentation().slide_width

    @property
    def slide_height(self) -> int:
        return self.get_presentation().slide_height
