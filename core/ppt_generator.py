import os
from typing import Optional

from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

from core.models import Question
from core.template_manager import TemplateManager
from core.ppt_style import align_from_string
from core.template_style import (
    ExtractedRunStyle,
    TemplateSlideStyle,
    delete_all_slides,
    extract_style_from_slide,
    merge_template_style_into_config,
    neutralize_option_colors_if_no_template_rgb,
)

EMU_PER_INCH = 914400


class PPTConfig:
    """PPT 生成配置（边距、间距、字体、颜色、排列等）"""

    def __init__(self):
        self.margin_left_in = 0.8
        self.margin_right_in = 0.8
        self.margin_top_in = 0.5

        self.stem_height_with_image_in = 1.5
        self.stem_height_no_image_in = 2.5
        self.stem_font_size = Pt(20)
        self.font_name = "微软雅黑"
        self.font_bold_stem = True
        self.stem_color = RGBColor(0x1A, 0x1A, 0x2E)
        self.stem_align = "left"

        self.image_max_width = Inches(5)
        self.image_max_height = Inches(2.5)
        self.image_h_align = "center"

        self.gap_after_stem_in = 0.2
        self.gap_after_image_in = 0.15
        self.gap_before_options_in = 0.2

        self.option_layout = "grid"
        self.one_row_height_in = 0.55
        self.one_row_gap_in = 0.06
        self.grid_layout = "ab_cd"
        self.grid_row_height_in = 0.9
        self.grid_col_gap_in = 0.15
        self.list_row_height_in = 0.7
        self.option_font_size = Pt(18)
        self.option_font_bold = False
        self.option_color = RGBColor(0x2D, 0x2D, 0x2D)
        self.option_letter_color = RGBColor(0x00, 0x6B, 0xBD)
        self.option_letter_bold = True
        self.option_align = "left"

        self.number_color = self.option_letter_color

    def stem_align_pp(self) -> int:
        return align_from_string(self.stem_align)

    def option_align_pp(self) -> int:
        return align_from_string(self.option_align)


def _scale_image(img_path: str, max_w_in: float, max_h_in: float):
    """读取图片尺寸并按比例缩放到限定范围，返回 (width_emu, height_emu)。"""
    px_w = px_h = None
    dpi = 96.0
    try:
        with Image.open(img_path) as img:
            px_w, px_h = img.size
    except Exception:
        pass

    if px_w and px_h and px_w > 0 and px_h > 0:
        img_w = px_w / dpi
        img_h = px_h / dpi
        scale = min(max_w_in / img_w, max_h_in / img_h, 1.0)
        return Inches(img_w * scale), Inches(img_h * scale)
    return Inches(min(max_w_in, 4.0)), Inches(min(max_h_in, 2.5))


def _align_pic_left(base_left: int, area_width: int, pic_width, h_align: str) -> int:
    pw = int(pic_width)
    if h_align == "right":
        return base_left + area_width - pw
    if h_align == "center":
        return max(base_left, base_left + (area_width - pw) // 2)
    return base_left


class PPTGenerator:
    """PPT 幻灯片生成器"""

    def __init__(self, template_manager: Optional[TemplateManager] = None,
                 config: Optional[PPTConfig] = None):
        self.tm = template_manager or TemplateManager()
        self.config = config or PPTConfig()
        self._prs: Optional[Presentation] = None
        self._tpl_style: Optional[TemplateSlideStyle] = None

    def generate(self, questions: list[Question], output_path: str,
                 template_path: Optional[str] = None,
                 progress_callback=None):
        self._tpl_style = None

        if template_path:
            self._prs = self.tm.load_template(template_path)
            if len(self._prs.slides) > 0:
                first = next(iter(self._prs.slides))
                self._tpl_style = extract_style_from_slide(first)
            self.config = PPTConfig()
            if self._tpl_style:
                merge_template_style_into_config(self.config, self._tpl_style)
                neutralize_option_colors_if_no_template_rgb(self.config, self._tpl_style)
            delete_all_slides(self._prs)
        else:
            self._prs = self.tm.create_default()

        total = len(questions)
        for idx, question in enumerate(questions):
            self._add_question_slide(question)
            if progress_callback:
                progress_callback(idx + 1, total)

        os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
        self._prs.save(output_path)
        return output_path

    # ── slide dispatch ──

    def _add_question_slide(self, question: Question):
        slide = self._prs.slides.add_slide(self.tm.get_blank_layout())
        ts = self._tpl_style
        if ts and ts.stem_rect and len(ts.option_rects) >= 4:
            self._layout_from_template(slide, question, ts)
        else:
            self._layout_default(slide, question)

    # ── template layout ──

    def _layout_from_template(self, slide, question: Question, ts: TemplateSlideStyle):
        cfg = self.config
        sr = ts.stem_rect

        self._add_stem_box(
            slide, question,
            sr.left, sr.top, sr.width, sr.height,
            style_override=ts.stem,
        )

        stem_bottom = sr.top + sr.height
        self._place_images_template(
            slide, question, ts, cfg, stem_bottom,
        )

        for i, opt in enumerate(question.options[:4]):
            if i >= len(ts.option_rects):
                break
            r = ts.option_rects[i]
            letter_ex, body_ex = self._resolve_option_style(ts, i)
            self._add_option_box(
                slide, opt.letter, opt.text,
                r.left, r.top, r.width, r.height,
                letter_tpl=letter_ex, body_tpl=body_ex,
            )

    def _place_images_template(self, slide, question, ts, cfg, stem_bottom):
        if not question.image_paths:
            return
        if ts.image_rect:
            ir = ts.image_rect
            top = ir.top
            for p in question.image_paths:
                if os.path.exists(p):
                    top = self._insert_image_in_rect(
                        slide, p, ir.left, ir.top, ir.width, ir.height, top,
                    )
        else:
            gap_after = int(Inches(cfg.gap_after_stem_in))
            gap_before = int(Inches(cfg.gap_before_options_in))
            opt_top = min(r.top for r in ts.option_rects)
            top = stem_bottom + gap_after
            band_h = max(1, opt_top - top - gap_before)
            for p in question.image_paths:
                if os.path.exists(p):
                    top = self._insert_image(
                        slide, p, ts.stem_rect.left, top,
                        ts.stem_rect.width, band_h,
                    )

    @staticmethod
    def _resolve_option_style(ts: TemplateSlideStyle, idx: int):
        letter_ex = body_ex = None
        if ts.option_box_styles and idx < len(ts.option_box_styles):
            ob = ts.option_box_styles[idx]
            letter_ex, body_ex = ob.letter, ob.body
        if letter_ex is None and body_ex is None and ts.option:
            letter_ex = body_ex = ts.option
        return letter_ex, body_ex

    # ── default layout ──

    def _layout_default(self, slide, question: Question):
        cfg = self.config
        slide_w = self._prs.slide_width
        ml = Inches(cfg.margin_left_in)
        mr = Inches(cfg.margin_right_in)
        mt = Inches(cfg.margin_top_in)
        cw = slide_w - ml - mr
        has_img = bool(question.image_paths)

        stem_h = Inches(
            cfg.stem_height_with_image_in if has_img else cfg.stem_height_no_image_in
        )
        self._add_stem_box(slide, question, ml, mt, cw, stem_h)

        top = mt + stem_h + Inches(cfg.gap_after_stem_in)
        if has_img:
            for p in question.image_paths:
                if os.path.exists(p):
                    top = self._insert_image(slide, p, ml, top, cw)

        opts_top = top + Inches(cfg.gap_before_options_in)
        ol = (cfg.option_layout or "grid").lower()
        if ol == "list":
            self._options_list(slide, question, ml, opts_top, cw)
        elif ol == "one_row":
            self._options_one_row(slide, question, ml, opts_top, cw)
        else:
            self._options_grid(slide, question, ml, opts_top, cw)

    # ── shared: stem ──

    def _add_stem_box(self, slide, question: Question,
                      left, top, width, height,
                      style_override: Optional[ExtractedRunStyle] = None):
        cfg = self.config
        box = slide.shapes.add_textbox(left, top, width, height)
        tf = box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = f"{question.number}. {question.stem}"

        ex = style_override
        p.font.size = Pt(ex.size_pt) if ex and ex.size_pt and ex.size_pt > 0 else cfg.stem_font_size
        p.font.bold = ex.bold if ex and ex.bold is not None else cfg.font_bold_stem
        p.font.name = ex.name if ex and ex.name else cfg.font_name
        p.font.color.rgb = ex.rgb if ex and ex.rgb else cfg.stem_color
        p.alignment = ex.alignment if ex and ex.alignment is not None else cfg.stem_align_pp()

    # ── shared: image insertion ──

    def _insert_image(self, slide, img_path: str,
                      left: int, top: int, area_w: int,
                      max_h_emu: int = 0) -> int:
        cfg = self.config
        max_w = cfg.image_max_width.inches
        max_h = cfg.image_max_height.inches
        if max_h_emu > 0:
            max_h = min(max_h, max_h_emu / EMU_PER_INCH)
        max_w = min(max_w, area_w / EMU_PER_INCH)

        w, h = _scale_image(img_path, max_w, max_h)
        pl = _align_pic_left(left, area_w, w, (cfg.image_h_align or "center").lower())
        pic = slide.shapes.add_picture(img_path, pl, top, width=w, height=h)
        return int(top + pic.height + Inches(cfg.gap_after_image_in))

    def _insert_image_in_rect(self, slide, img_path: str,
                              rect_l: int, rect_t: int,
                              rect_w: int, rect_h: int,
                              start_top: int) -> int:
        cfg = self.config
        bottom = rect_t + rect_h
        remaining = bottom - start_top
        if remaining <= 0:
            return start_top

        max_w = min(cfg.image_max_width.inches, rect_w / EMU_PER_INCH)
        max_h = min(cfg.image_max_height.inches, remaining / EMU_PER_INCH)
        w, h = _scale_image(img_path, max_w, max_h)
        pl = _align_pic_left(rect_l, rect_w, w, (cfg.image_h_align or "center").lower())
        pic = slide.shapes.add_picture(img_path, pl, start_top, width=w, height=h)
        nxt = int(start_top + pic.height + Inches(cfg.gap_after_image_in))
        return min(nxt, bottom)

    # ── shared: option box ──

    def _add_option_box(self, slide, letter: str, text: str,
                        left, top, width, height,
                        letter_tpl: Optional[ExtractedRunStyle] = None,
                        body_tpl: Optional[ExtractedRunStyle] = None):
        cfg = self.config
        box = slide.shapes.add_textbox(left, top, width, height)
        tf = box.text_frame
        tf.word_wrap = True
        tf.auto_size = None

        p = tf.paragraphs[0]
        rl = p.add_run()
        rl.text = f"{letter}. "
        self._style_run(rl, cfg.option_font_size, cfg.option_letter_bold,
                        cfg.option_letter_color, letter_tpl)

        rt = p.add_run()
        rt.text = text
        self._style_run(rt, cfg.option_font_size, cfg.option_font_bold,
                        cfg.option_color, body_tpl or letter_tpl)

        if body_tpl and body_tpl.alignment is not None:
            p.alignment = body_tpl.alignment
        elif letter_tpl and letter_tpl.alignment is not None:
            p.alignment = letter_tpl.alignment
        else:
            p.alignment = cfg.option_align_pp()

    def _style_run(self, run, default_size, default_bold, default_rgb,
                   tpl: Optional[ExtractedRunStyle]):
        cfg = self.config
        run.font.size = Pt(tpl.size_pt) if tpl and tpl.size_pt and tpl.size_pt > 0 else default_size
        run.font.bold = tpl.bold if tpl and tpl.bold is not None else default_bold
        run.font.name = tpl.name if tpl and tpl.name else cfg.font_name
        run.font.color.rgb = tpl.rgb if tpl and tpl.rgb else default_rgb

    # ── option layouts ──

    def _grid_positions(self, left, top, col_w, row_h):
        cfg = self.config
        g = Inches(cfg.grid_col_gap_in)
        c0, c1 = left, left + col_w + g
        r0, r1 = top, top + row_h
        if cfg.grid_layout == "ac_bd":
            return [(c0, r0), (c0, r1), (c1, r0), (c1, r1)]
        return [(c0, r0), (c1, r0), (c0, r1), (c1, r1)]

    def _options_grid(self, slide, q: Question, left, top, cw):
        cfg = self.config
        g = Inches(cfg.grid_col_gap_in)
        col_w = (cw - g) // 2
        row_h = Inches(cfg.grid_row_height_in)
        pos = self._grid_positions(left, top, col_w, row_h)
        w, h = col_w - Inches(0.3), row_h - Inches(0.1)
        for i, opt in enumerate(q.options[:4]):
            if i < len(pos):
                self._add_option_box(slide, opt.letter, opt.text, pos[i][0], pos[i][1], w, h)

    def _options_list(self, slide, q: Question, left, top, cw):
        rh = Inches(self.config.list_row_height_in)
        for i, opt in enumerate(q.options[:4]):
            self._add_option_box(slide, opt.letter, opt.text,
                                 left, top + i * rh, cw, rh - Inches(0.1))

    def _options_one_row(self, slide, q: Question, left, top, cw):
        cfg = self.config
        opts = q.options[:4]
        if not opts:
            return
        n = len(opts)
        gap = Inches(cfg.one_row_gap_in)
        cell_w = (cw - gap * (n - 1)) // n if n > 1 else cw
        row_h = Inches(cfg.one_row_height_in)
        pad = Inches(0.04)
        for i, opt in enumerate(opts):
            x = left + i * (cell_w + gap)
            self._add_option_box(slide, opt.letter, opt.text,
                                 x, top, cell_w - pad, row_h - pad)
