import logging
import os
import shutil
import tempfile
import weakref
from typing import Optional

from docx.oxml.ns import qn

LOGGER = logging.getLogger(__name__)


class ImageExtractor:
    def __init__(self, temp_dir: Optional[str] = None):
        if temp_dir:
            self.temp_dir = temp_dir
            os.makedirs(temp_dir, exist_ok=True)
            self._owns_temp = False
        else:
            self.temp_dir = tempfile.mkdtemp(prefix="pptconvert_")
            self._owns_temp = True

        self._counter = 0
        self._closed = False
        self._finalizer = weakref.finalize(
            self,
            self._cleanup_temp_dir,
            self.temp_dir,
            self._owns_temp,
        )

    def extract_from_paragraph(self, paragraph, question_number: int) -> list[str]:
        image_paths = []
        for run in paragraph.runs:
            drawing_elements = run._element.findall(qn("w:drawing"))
            for drawing in drawing_elements:
                blip_elements = drawing.findall(".//" + qn("a:blip"))
                for blip in blip_elements:
                    r_embed = blip.get(qn("r:embed"))
                    if not r_embed:
                        continue
                    path = self._save_image(paragraph, r_embed, question_number)
                    if path:
                        image_paths.append(path)
        return image_paths

    def _save_image(self, paragraph, r_embed: str, question_number: int) -> Optional[str]:
        try:
            part = paragraph.part
            related_part = part.related_parts.get(r_embed)
            if related_part is None:
                LOGGER.warning(
                    "Skipping image for question %s: relationship %s was not found",
                    question_number,
                    r_embed,
                )
                return None

            image_data = related_part.blob
            content_type = related_part.content_type

            ext = self._get_extension(content_type)
            self._counter += 1
            filename = f"q{question_number}_img{self._counter}{ext}"
            filepath = os.path.join(self.temp_dir, filename)

            with open(filepath, "wb") as file_obj:
                file_obj.write(image_data)

            return filepath
        except Exception:
            LOGGER.warning(
                "Failed to extract image for question %s (relationship %s)",
                question_number,
                r_embed,
                exc_info=True,
            )
            return None

    @staticmethod
    def _get_extension(content_type: str) -> str:
        mapping = {
            "image/png": ".png",
            "image/jpeg": ".jpg",
            "image/gif": ".gif",
            "image/bmp": ".bmp",
            "image/tiff": ".tiff",
            "image/x-emf": ".emf",
            "image/x-wmf": ".wmf",
        }
        return mapping.get(content_type, ".png")

    @staticmethod
    def _cleanup_temp_dir(path: str, owns_temp: bool) -> None:
        if owns_temp and path and os.path.exists(path):
            shutil.rmtree(path, ignore_errors=True)

    def cleanup(self):
        if self._closed:
            return
        self._closed = True
        if self._finalizer.alive:
            self._finalizer()

    def __del__(self):
        try:
            self.cleanup()
        except Exception:
            pass
