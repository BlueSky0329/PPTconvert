import os
import tempfile
import shutil
from typing import Optional

from docx.oxml.ns import qn


class ImageExtractor:
    """从 Word 文档中提取图片并保存到临时目录"""

    def __init__(self, temp_dir: Optional[str] = None):
        if temp_dir:
            self.temp_dir = temp_dir
            os.makedirs(temp_dir, exist_ok=True)
            self._owns_temp = False
        else:
            self.temp_dir = tempfile.mkdtemp(prefix="pptconvert_")
            self._owns_temp = True
        self._counter = 0

    def extract_from_paragraph(self, paragraph, question_number: int) -> list[str]:
        """从段落中提取所有内联图片，返回保存的文件路径列表"""
        image_paths = []
        for run in paragraph.runs:
            drawing_elements = run._element.findall(qn('w:drawing'))
            for drawing in drawing_elements:
                blip_elements = drawing.findall('.//' + qn('a:blip'))
                for blip in blip_elements:
                    r_embed = blip.get(qn('r:embed'))
                    if r_embed:
                        path = self._save_image(paragraph, r_embed, question_number)
                        if path:
                            image_paths.append(path)
        return image_paths

    def _save_image(self, paragraph, r_embed: str, question_number: int) -> Optional[str]:
        """根据关系 ID 提取图片二进制数据并保存"""
        try:
            part = paragraph.part
            related_part = part.related_parts.get(r_embed)
            if related_part is None:
                return None

            image_data = related_part.blob
            content_type = related_part.content_type

            ext = self._get_extension(content_type)
            self._counter += 1
            filename = f"q{question_number}_img{self._counter}{ext}"
            filepath = os.path.join(self.temp_dir, filename)

            with open(filepath, 'wb') as f:
                f.write(image_data)

            return filepath
        except Exception:
            return None

    @staticmethod
    def _get_extension(content_type: str) -> str:
        mapping = {
            'image/png': '.png',
            'image/jpeg': '.jpg',
            'image/gif': '.gif',
            'image/bmp': '.bmp',
            'image/tiff': '.tiff',
            'image/x-emf': '.emf',
            'image/x-wmf': '.wmf',
        }
        return mapping.get(content_type, '.png')

    def cleanup(self):
        """清理临时目录（在 PPT 生成完成后再调用，勿在析构时自动删除，否则图片路径会失效）"""
        if self._owns_temp and os.path.exists(self.temp_dir):
            shutil.rmtree(self.temp_dir, ignore_errors=True)
