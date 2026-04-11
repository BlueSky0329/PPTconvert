import os
import shutil
import tempfile
import threading
from typing import Optional
import tkinter as tk
import tkinter.font as tkfont
from tkinter import colorchooser, filedialog, messagebox, simpledialog

import ttkbootstrap as ttk
from PIL import Image, ImageTk
from ttkbootstrap.constants import *

from core.word_parser import WordParser
from domain.models import ALL_SUBJECT_KINDS, SUBJECT_DISPLAY_NAMES
from domain.project_editor import (
    clear_option_image,
    insert_option_after,
    insert_material_after,
    merge_adjacent_materials,
    move_option,
    move_data_question,
    reclassify_objective_section,
    remove_option,
    remove_question,
    replace_option_image,
    rename_material,
    renumber_question,
    set_question_option_layout,
    update_option_text,
    update_question_stem,
)
from core.ppt_generator import PPTGenerator, PPTConfig
from core.ppt_style import parse_hex_color
from core.models import Question
from exporters.manifest_json import load_project_manifest_project
from exporters.material_crops import crop_material_regions, crop_page_regions
from exporters.pptx_slides import project_to_ppt_questions
from gui.font_data import build_font_values
from gui import ui_constants as U
from workflows.project_flow import build_word_project

_PAD = 14
_PDF_WIZARD_STEPS = (
    ("导入 PDF", "选择试卷文件，向导会按文件名预填默认输出路径。"),
    ("识别设置", "决定要进入工程的题目范围；下一步会生成或刷新结构化预览。"),
    ("结果预览", "校对题号、材料和题目归属；这一步的人工修正会直接用于导出。"),
    ("导出结果", "从同一份题目工程导出题本 Word、授课 PPT 和工程 JSON。"),
)
_PDF_SUBJECT_ORDER = tuple(ALL_SUBJECT_KINDS)
_PDF_QUESTION_LAYOUT_CHOICES = (
    ("", "跟随全局"),
    ("one_row", "一行四项"),
    ("grid", "两行两列"),
    ("list", "四行竖排"),
)
_PDF_QUESTION_LAYOUT_LABELS = {
    "": "跟随全局",
    "one_row": "一行四项",
    "grid": "两行两列",
    "list": "四行竖排",
}
_DOCUMENT_SUBJECT_CHOICES = (
    ("auto", "自动识别"),
    ("politics", "政治理论"),
    ("common_sense", "常识判断"),
    ("verbal", "言语理解与表达"),
    ("quant", "数量关系"),
    ("reasoning", "判断推理"),
    ("data", "资料分析"),
)
_DOCUMENT_SUBJECT_LABELS = {key: label for key, label in _DOCUMENT_SUBJECT_CHOICES}


class PPTConvertApp:
    """PDF 试卷整理图形界面。"""

    def __init__(self):
        self.root = ttk.Window(
            title=U.APP_TITLE,
            themename=U.THEME_NAME,
            size=(1320, 920),
            minsize=(1080, 760),
            resizable=(True, True),
        )
        self.root.update_idletasks()
        self._fit_window_to_screen()

        # Tk variables
        self.word_path = tk.StringVar()
        self.template_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.use_template = tk.BooleanVar(value=False)
        self.word_document_subject = tk.StringVar(value=_DOCUMENT_SUBJECT_LABELS["auto"])
        self.pdf_document_subject = tk.StringVar(value=_DOCUMENT_SUBJECT_LABELS["auto"])
        self.option_layout = tk.StringVar(value="grid")
        self.font_size_stem = tk.IntVar(value=20)
        self.font_size_option = tk.IntVar(value=18)

        self.margin_left = tk.DoubleVar(value=0.8)
        self.margin_right = tk.DoubleVar(value=0.8)
        self.margin_top = tk.DoubleVar(value=0.5)
        self.stem_h_img = tk.DoubleVar(value=1.5)
        self.stem_h_no = tk.DoubleVar(value=2.5)
        self.gap_stem = tk.DoubleVar(value=0.2)
        self.gap_img = tk.DoubleVar(value=0.15)
        self.gap_opts = tk.DoubleVar(value=0.2)
        self.stem_align = tk.StringVar(value="left")
        self.image_align = tk.StringVar(value="center")
        self.image_max_w = tk.DoubleVar(value=5.0)
        self.image_max_h = tk.DoubleVar(value=2.5)
        self.grid_layout = tk.StringVar(value="ab_cd")
        self.grid_row_h = tk.DoubleVar(value=0.9)
        self.grid_col_gap = tk.DoubleVar(value=0.15)
        self.list_row_h = tk.DoubleVar(value=0.7)
        self.option_align = tk.StringVar(value="left")
        self.font_name = tk.StringVar(value="微软雅黑")
        self.stem_bold = tk.BooleanVar(value=True)
        self.option_letter_bold = tk.BooleanVar(value=True)
        self.option_text_bold = tk.BooleanVar(value=False)
        self.color_stem = tk.StringVar(value="#1A1A2E")
        self.color_option = tk.StringVar(value="#2D2D2D")
        self.color_letter = tk.StringVar(value="#006BBD")
        self.one_row_h = tk.DoubleVar(value=0.55)
        self.one_row_gap = tk.DoubleVar(value=0.06)
        self.preview_has_image = tk.BooleanVar(value=True)

        self._font_list = build_font_values()
        self._preview_after: Optional[str] = None
        self.questions: list[Question] = []
        self.parser: WordParser | None = None
        self.pdf_project = None
        self._pdf_project_context: dict[str, str] = {}
        self._pdf_preview_payloads: dict[str, dict] = {}
        self._pdf_material_preview_dir: Optional[str] = None
        self._pdf_material_preview_cache: dict[str, tuple[str, list[str]]] = {}
        self._pdf_material_preview_paths: list[str] = []
        self._pdf_material_preview_source = ""
        self._pdf_material_preview_title = ""
        self._pdf_material_preview_index = 0
        self._pdf_material_preview_photo = None
        self._pdf_question_preview_photos: list[ImageTk.PhotoImage] = []
        self._pdf_stem_preview_paths: list[str] = []
        self._pdf_stem_preview_index = 0
        self._pdf_stem_preview_photo = None

        self.pdf_path = tk.StringVar()
        self.pdf_word_out = tk.StringVar()
        self.pdf_ppt_out = tk.StringVar()
        self.pdf_manifest_out = tk.StringVar()
        self.pdf_question_range = tk.StringVar()
        self._pdf_question_layout_var = tk.StringVar(value="")
        self._pdf_question_editor_message = tk.StringVar(value="选择一道题后，可在这里实时修改题干，并为该题单独切换选项布局。")
        self._pdf_subject_vars = {
            kind: tk.BooleanVar(value=True)
            for kind in _PDF_SUBJECT_ORDER
        }
        self._pdf_wizard_step = 0
        self._pdf_wizard_pending_step: Optional[int] = None
        self._pdf_step_frames: list[ttk.Frame] = []
        self._pdf_step_buttons: list[ttk.Button] = []
        self._pdf_question_editor_target = None
        self._pdf_editor_updating = False
        self._pdf_section_subject_var = tk.StringVar(value=_DOCUMENT_SUBJECT_LABELS["auto"])
        self._pdf_question_layout_buttons: list[ttk.Radiobutton] = []
        self._pdf_option_editors: dict[str, tk.Text] = {}
        self._pdf_option_image_labels: dict[str, ttk.Label] = {}
        self._pdf_option_view_buttons: dict[str, ttk.Button] = {}
        self._pdf_option_recrop_buttons: dict[str, ttk.Button] = {}
        self._pdf_option_clear_buttons: dict[str, ttk.Button] = {}
        self._pdf_option_replace_buttons: dict[str, ttk.Button] = {}
        self._pdf_option_move_up_buttons: dict[str, ttk.Button] = {}
        self._pdf_option_move_down_buttons: dict[str, ttk.Button] = {}
        self._pdf_option_insert_buttons: dict[str, ttk.Button] = {}
        self._pdf_option_remove_buttons: dict[str, ttk.Button] = {}
        self._pdf_project_dirty = False

        self._build_ui()
        self._bind_pdf_wizard_updates()
        self._bind_preview_updates()
        try:
            self._pdf_question_layout_var.trace_add("write", lambda *_: self._on_pdf_question_layout_change())
        except Exception:
            pass
        self._schedule_preview_refresh()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    # UI

    def _build_ui(self):
        outer = ttk.Frame(self.root)
        outer.pack(fill=BOTH, expand=YES)

        main = ttk.Frame(outer, padding=_PAD)
        main.pack(fill=BOTH, expand=YES)

        self._build_header(main)
        self._build_workspace_tabs(main)
        self._build_progress_footer(outer)

    def _build_workspace_tabs(self, parent):
        notebook = ttk.Notebook(parent, bootstyle="primary")
        notebook.pack(fill=BOTH, expand=YES)
        self._workspace_notebook = notebook

        pdf_tab, pdf_body = self._make_scrollable_tab(notebook)
        word_tab, word_body = self._make_scrollable_tab(notebook)
        settings_tab, settings_body = self._make_scrollable_tab(notebook)
        self._pdf_workspace_tab = pdf_tab
        self._word_workspace_tab = word_tab
        self._settings_workspace_tab = settings_tab
        self._ppt_settings_tab = settings_tab

        notebook.add(pdf_tab, text=" PDF 试卷整理 ")
        notebook.add(word_tab, text=" Word 生成 PPT ")
        notebook.add(settings_tab, text=" PPT 导出设置 ")

        self._build_pdf_tab(pdf_body)
        self._build_word_tab(word_body)
        self._build_ppt_settings_tab(settings_body)

    def _make_scrollable_tab(self, parent):
        host = ttk.Frame(parent)
        canvas = tk.Canvas(host, highlightthickness=0, bd=0)
        canvas.pack(side=LEFT, fill=BOTH, expand=YES)
        scrollbar = ttk.Scrollbar(host, orient=VERTICAL, command=canvas.yview)
        scrollbar.pack(side=RIGHT, fill=Y)
        canvas.configure(yscrollcommand=scrollbar.set)

        body = ttk.Frame(canvas)
        window_id = canvas.create_window((0, 0), window=body, anchor="nw")

        def _sync_scrollregion(_event=None):
            canvas.configure(scrollregion=canvas.bbox("all"))

        def _sync_width(event):
            canvas.itemconfigure(window_id, width=event.width)

        def _on_mousewheel(event):
            delta = event.delta
            if delta == 0:
                return
            step = -1 if delta > 0 else 1
            canvas.yview_scroll(step, "units")

        body.bind("<Configure>", _sync_scrollregion)
        canvas.bind("<Configure>", _sync_width)
        canvas.bind("<Enter>", lambda _event: canvas.bind_all("<MouseWheel>", _on_mousewheel))
        canvas.bind("<Leave>", lambda _event: canvas.unbind_all("<MouseWheel>"))
        return host, body

    def _fit_window_to_screen(self):
        try:
            self.root.state("zoomed")
            return
        except Exception:
            pass

        try:
            screen_w = max(1200, int(self.root.winfo_screenwidth() * 0.96))
            screen_h = max(760, int(self.root.winfo_screenheight() * 0.92))
            self.root.geometry(f"{screen_w}x{screen_h}+18+18")
        except Exception:
            self.root.place_window_center()

    # Header

    def _build_header(self, parent):
        hdr = ttk.Frame(parent, padding=(20, 16, 20, 12))
        hdr.pack(fill=X, pady=(0, 6))

        top_row = ttk.Frame(hdr)
        top_row.pack(fill=X)
        ttk.Label(top_row, text=U.APP_TITLE, font=("", 17, "bold")).pack(side=LEFT)
        ttk.Label(
            top_row, text=U.STEP_HINT,
            font=("", 9), bootstyle="secondary",
        ).pack(side=RIGHT)

        ttk.Label(
            hdr, text=U.HERO_SUB,
            wraplength=880, font=("", 10), bootstyle="secondary",
        ).pack(anchor=W, pady=(4, 0))

        ttk.Separator(hdr, orient=HORIZONTAL).pack(fill=X, pady=(10, 0))

    # 1. Files

    def _build_file_section(self, parent):
        frame = ttk.Labelframe(parent, text=" ① 选择文件 ", bootstyle="primary", padding=_PAD)
        frame.pack(fill=X, pady=(0, 10))

        self._file_row(frame, "Word 文件", self.word_path,
                       self._browse_word, "浏览...", pady=(0, 0))
        self._file_row(frame, "输出 PPT", self.output_path,
                       self._browse_output, "另存为...", pady=(6, 0))
        row = ttk.Frame(frame)
        row.pack(fill=X, pady=(8, 0))
        ttk.Label(row, text="整份科目", width=10).pack(side=LEFT)
        ttk.Combobox(
            row,
            textvariable=self.word_document_subject,
            values=[label for _key, label in _DOCUMENT_SUBJECT_CHOICES],
            state="readonly",
            width=18,
        ).pack(side=LEFT)
        ttk.Label(
            row,
            text="适合单科题库或没有大标题的资料；自动识别不稳时可直接固定整份 Word 的科目。",
            bootstyle="secondary",
        ).pack(side=LEFT, padx=(8, 0))

    def _file_row(self, parent, label, var, cmd, btn_text, pady=(0, 0)):
        row = ttk.Frame(parent)
        row.pack(fill=X, pady=pady)
        ttk.Label(row, text=label, width=10).pack(side=LEFT)
        ttk.Entry(row, textvariable=var).pack(side=LEFT, fill=X, expand=YES, padx=(0, 8))
        ttk.Button(row, text=btn_text, command=cmd,
                   bootstyle="outline", width=9).pack(side=RIGHT)

    # 2. Template

    def _build_template_section(self, parent):
        frame = ttk.Labelframe(parent, text=" PPT 模板（可选） ", bootstyle="info", padding=_PAD)
        frame.pack(fill=X, pady=(0, 10))

        row = ttk.Frame(frame)
        row.pack(fill=X)
        ttk.Checkbutton(
            row, text="使用 .pptx 模板", variable=self.use_template,
            command=self._toggle_template, bootstyle="round-toggle",
        ).pack(side=LEFT)
        self.template_entry = ttk.Entry(row, textvariable=self.template_path, state=DISABLED)
        self.template_entry.pack(side=LEFT, fill=X, expand=YES, padx=(12, 8))
        self.template_btn = ttk.Button(
            row, text="选择...", command=self._browse_template,
            state=DISABLED, bootstyle="info-outline", width=8,
        )
        self.template_btn.pack(side=RIGHT)

        self._tpl_hint = ttk.Label(
            frame,
            text=(
                "启用模板后，字体、颜色和对齐以模板第一页为准，下方版式设置不再参与生成。\n"
                "推荐在模板第一页放置 [题干]、[图片]、[选项A] 到 [选项D] 文本框；\n"
                "或者按顺序放置 1 个题干框和 4 个选项框。单个文本框里写 A. 到 D. 时会自动拆成 2x2。"
            ),
            bootstyle="secondary", font=("", 9), wraplength=860, justify=LEFT,
        )
        self._tpl_hint.pack(anchor=W, pady=(8, 0))
        self._tpl_hint.pack_forget()

    # 3. Config

    def _build_config_section(self, parent):
        self._config_frame = ttk.Labelframe(
            parent, text=" PPT 版式与样式（非模板模式） ", bootstyle="primary", padding=(2, 8),
        )
        self._config_frame.pack(fill=X, pady=(0, 10))

        self._tpl_overlay_label = ttk.Label(
            self._config_frame,
            text="已启用模板：版式与样式以模板第一页为准，此处设置不再参与生成。",
            font=("", 10), bootstyle="warning", padding=(16, 12),
        )

        nb = ttk.Notebook(self._config_frame, bootstyle="primary")
        nb.pack(fill=X, padx=8, pady=(4, 8))
        self._config_notebook = nb

        tab_layout = ttk.Frame(nb, padding=10)
        tab_options = ttk.Frame(nb, padding=10)
        tab_font = ttk.Frame(nb, padding=10)
        tab_preview = ttk.Frame(nb, padding=10)
        nb.add(tab_layout, text=" 布局 ")
        nb.add(tab_options, text=" 选项排列 ")
        nb.add(tab_font, text=" 字体与颜色 ")
        nb.add(tab_preview, text=" 示意图 ")

        self._build_tab_layout(tab_layout)
        self._build_tab_options(tab_options)
        self._build_tab_font(tab_font)
        self._build_tab_preview(tab_preview)

    # 4. Question table

    def _build_question_table(self, parent):
        frame = ttk.Labelframe(
            parent,
            text=" ④ 解析结果 ",
            bootstyle="primary", padding=(2, 8),
        )
        frame.pack(fill=BOTH, expand=YES, pady=(0, 6))
        inner = ttk.Frame(frame, padding=(10, 8))
        inner.pack(fill=BOTH, expand=YES)

        cols = ("num", "stem", "options", "image")
        self.tree = ttk.Treeview(inner, columns=cols, show="headings", height=7,
                                 bootstyle="info")
        self.tree.heading("num", text="#")
        self.tree.heading("stem", text="题干摘要")
        self.tree.heading("options", text="选项")
        self.tree.heading("image", text="配图")
        self.tree.column("num", width=40, anchor=CENTER)
        self.tree.column("stem", width=460)
        self.tree.column("options", width=50, anchor=CENTER)
        self.tree.column("image", width=56, anchor=CENTER)

        sb = ttk.Scrollbar(inner, orient=VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side=LEFT, fill=BOTH, expand=YES)
        sb.pack(side=RIGHT, fill=Y)

        ttk.Label(
            frame, text=U.TIP_SECONDARY, font=("", 9), bootstyle="secondary",
        ).pack(anchor=W, padx=10, pady=(2, 4))

    # PPT tab buttons + shared footer

    def _build_ppt_tab_buttons(self, parent):
        bar = ttk.Frame(parent, padding=(16, 8))
        bar.pack(side=BOTTOM, fill=X)

        row = ttk.Frame(bar)
        row.pack(fill=X)

        ttk.Button(
            row,
            text="PPT 设置",
            command=self._open_ppt_settings_tab,
            bootstyle="info-outline",
            width=10,
        ).pack(side=LEFT)

        ttk.Button(row, text="清空", command=self._clear_all,
                   bootstyle="secondary-outline", width=7).pack(side=RIGHT, padx=3)
        ttk.Button(row, text="生成 PPT", command=self._generate_ppt,
                   bootstyle="success-outline", width=10).pack(side=RIGHT, padx=3)
        ttk.Button(row, text="解析并预览", command=self._parse_preview,
                   bootstyle="info-outline", width=10).pack(side=RIGHT, padx=3)
        ttk.Button(row, text="一键生成", command=self._convert_all,
                   bootstyle="success", width=10).pack(side=RIGHT, padx=3)

    def _build_pdf_tab(self, parent):
        frame = ttk.Frame(parent, padding=_PAD)
        frame.pack(fill=BOTH, expand=YES)

        hdr = ttk.Frame(frame, padding=(12, 12, 12, 6))
        hdr.pack(fill=X)
        ttk.Label(hdr, text="PDF 试卷工作流", font=("", 15, "bold")).pack(anchor=W)
        ttk.Label(
            hdr, text=U.PDF_TAB_HINT, wraplength=860, font=("", 9), bootstyle="secondary",
        ).pack(anchor=W, pady=(6, 0))

        self._build_pdf_wizard(frame)

    def _build_word_tab(self, parent):
        frame = ttk.Frame(parent, padding=_PAD)
        frame.pack(fill=BOTH, expand=YES)

        hdr = ttk.Frame(frame, padding=(12, 12, 12, 6))
        hdr.pack(fill=X)
        ttk.Label(hdr, text="Word 生成 PPT", font=("", 15, "bold")).pack(anchor=W)
        ttk.Label(
            hdr,
            text=U.WORD_TAB_HINT,
            wraplength=860,
            font=("", 9),
            bootstyle="secondary",
        ).pack(anchor=W, pady=(6, 0))

        helper = ttk.Labelframe(frame, text=" 使用说明 ", bootstyle="info", padding=_PAD)
        helper.pack(fill=X, pady=(0, 8))
        ttk.Label(
            helper,
            text=(
                "直接选择你自己的 Word 题库即可解析。PPT 模板、字体、颜色和选项布局"
                "统一在“PPT 导出设置”标签页里调整。"
            ),
            wraplength=860,
            justify=LEFT,
            bootstyle="secondary",
        ).pack(side=LEFT, fill=X, expand=YES)
        ttk.Button(
            helper,
            text="打开 PPT 设置",
            command=self._open_ppt_settings_tab,
            bootstyle="info-outline",
            width=14,
        ).pack(side=RIGHT, padx=(10, 0))

        content = ttk.Frame(frame)
        content.pack(fill=BOTH, expand=YES)
        self._build_file_section(content)
        self._build_question_table(content)
        self._build_ppt_tab_buttons(content)

    def _build_ppt_settings_tab(self, parent):
        frame = ttk.Frame(parent, padding=_PAD)
        frame.pack(fill=BOTH, expand=YES)

        hdr = ttk.Frame(frame, padding=(12, 12, 12, 6))
        hdr.pack(fill=X)
        ttk.Label(hdr, text="PPT 导出设置", font=("", 15, "bold")).pack(anchor=W)
        ttk.Label(
            hdr,
            text=U.PPT_SETTINGS_HINT,
            wraplength=860,
            font=("", 9),
            bootstyle="secondary",
        ).pack(anchor=W, pady=(6, 0))

        self._build_template_section(frame)
        self._build_config_section(frame)

    def _build_pdf_wizard(self, parent):
        step_bar = ttk.Frame(parent, padding=(12, 6, 12, 6))
        step_bar.pack(fill=X, pady=(2, 6))
        self._pdf_step_buttons = []
        for index, (title, _hint) in enumerate(_PDF_WIZARD_STEPS):
            button = ttk.Button(
                step_bar,
                text=f"{index + 1}. {title}",
                command=lambda i=index: self._request_pdf_wizard_step(i),
                width=16,
                bootstyle="secondary-outline",
            )
            button.pack(side=LEFT, padx=(0, 8))
            self._pdf_step_buttons.append(button)

        intro = ttk.Frame(parent, padding=(12, 0, 12, 6))
        intro.pack(fill=X)
        self._pdf_step_title_label = ttk.Label(intro, font=("", 12, "bold"))
        self._pdf_step_title_label.pack(anchor=W)
        self._pdf_step_hint_label = ttk.Label(
            intro,
            wraplength=860,
            font=("", 9),
            bootstyle="secondary",
            justify=LEFT,
        )
        self._pdf_step_hint_label.pack(anchor=W, pady=(4, 0))

        self._pdf_step_host = ttk.Frame(parent)
        self._pdf_step_host.pack(fill=BOTH, expand=YES)

        self._pdf_step_frames = []
        for builder in (
            self._build_pdf_step_import,
            self._build_pdf_step_settings,
            self._build_pdf_step_preview,
            self._build_pdf_step_export,
        ):
            step_frame = ttk.Frame(self._pdf_step_host)
            builder(step_frame)
            self._pdf_step_frames.append(step_frame)

        nav = ttk.Frame(parent, padding=(12, 8, 12, 0))
        nav.pack(fill=X)
        self._pdf_nav_status_label = ttk.Label(nav, bootstyle="secondary")
        self._pdf_nav_status_label.pack(side=LEFT, fill=X, expand=YES)
        self._pdf_prev_btn = ttk.Button(
            nav,
            text="上一步",
            command=self._go_prev_pdf_step,
            bootstyle="secondary-outline",
            width=10,
        )
        self._pdf_prev_btn.pack(side=RIGHT, padx=(6, 0))
        self._pdf_next_btn = ttk.Button(
            nav,
            text="下一步",
            command=self._go_next_pdf_step,
            bootstyle="primary",
            width=16,
        )
        self._pdf_next_btn.pack(side=RIGHT)

        self._show_pdf_wizard_step(0)

    def _build_pdf_step_import(self, parent):
        box = ttk.Labelframe(parent, text=" 第一步：导入试卷 ", bootstyle="info", padding=_PAD)
        box.pack(fill=X, pady=(0, 8))

        self._file_row(box, "PDF 文件", self.pdf_path, self._browse_pdf, "浏览...", pady=(0, 0))
        manifest_row = ttk.Frame(box)
        manifest_row.pack(fill=X, pady=(8, 0))
        ttk.Label(
            manifest_row,
            text="已有工程",
            width=10,
        ).pack(side=LEFT)
        ttk.Label(
            manifest_row,
            text="如果之前已经导出过工程 JSON，可以直接载入继续修改。",
            bootstyle="secondary",
        ).pack(side=LEFT, fill=X, expand=YES)
        ttk.Button(
            manifest_row,
            text="载入工程 JSON",
            command=self._load_pdf_manifest_project,
            bootstyle="info-outline",
            width=14,
        ).pack(side=RIGHT)
        ttk.Label(
            box,
            text="导出文件路径会在最后一步统一确认；当前只需要先锁定试卷 PDF。",
            font=("", 9),
            bootstyle="secondary",
        ).pack(anchor=W, pady=(8, 0))

        self._pdf_import_summary = ttk.Label(
            box,
            wraplength=840,
            justify=LEFT,
            bootstyle="secondary",
        )
        self._pdf_import_summary.pack(anchor=W, pady=(10, 0))

        tips = ttk.Labelframe(parent, text=" 工作流说明 ", bootstyle="secondary", padding=_PAD)
        tips.pack(fill=X, pady=(0, 8))
        ttk.Label(
            tips,
            text=(
                "向导会按“导入 PDF -> 识别设置 -> 结果预览 -> 导出结果”推进。\n"
                "资料分析 PPT 会优先复用从 PDF 页面裁切出来的材料图，不再走 Word 二次解析。"
            ),
            wraplength=840,
            justify=LEFT,
        ).pack(anchor=W)

    def _build_pdf_step_settings(self, parent):
        box = ttk.Labelframe(parent, text=" 第二步：识别设置 ", bootstyle="info", padding=_PAD)
        box.pack(fill=X, pady=(0, 8))

        subject_box = ttk.Frame(box)
        subject_box.pack(fill=X, pady=(0, 0))
        ttk.Label(subject_box, text="处理科目", width=10).pack(side=LEFT, anchor=N)
        subject_panel = ttk.Frame(subject_box)
        subject_panel.pack(side=LEFT, fill=X, expand=YES)

        row1 = ttk.Frame(subject_panel)
        row1.pack(anchor=W)
        row2 = ttk.Frame(subject_panel)
        row2.pack(anchor=W, pady=(6, 0))
        action_row = ttk.Frame(subject_panel)
        action_row.pack(anchor=W, pady=(8, 0))

        for index, kind in enumerate(_PDF_SUBJECT_ORDER):
            host = row1 if index < 3 else row2
            ttk.Checkbutton(
                host,
                text=SUBJECT_DISPLAY_NAMES.get(kind, kind),
                variable=self._pdf_subject_vars[kind],
                bootstyle="round-toggle",
            ).pack(side=LEFT, padx=(0, 10))

        ttk.Button(
            action_row,
            text="全选",
            command=lambda: self._set_all_pdf_subjects(True),
            bootstyle="secondary-outline",
            width=8,
        ).pack(side=LEFT, padx=(0, 6))
        ttk.Button(
            action_row,
            text="清空",
            command=lambda: self._set_all_pdf_subjects(False),
            bootstyle="secondary-outline",
            width=8,
        ).pack(side=LEFT)

        doc_subject_row = ttk.Frame(box)
        doc_subject_row.pack(fill=X, pady=(10, 0))
        ttk.Label(doc_subject_row, text="整份科目", width=10).pack(side=LEFT)
        ttk.Combobox(
            doc_subject_row,
            textvariable=self.pdf_document_subject,
            values=[label for _key, label in _DOCUMENT_SUBJECT_CHOICES],
            state="readonly",
            width=18,
        ).pack(side=LEFT)
        ttk.Label(
            doc_subject_row,
            text="单科整卷或没有大标题时可固定整份 PDF 科目；自动识别会走启发式判断。",
            font=("", 9),
            bootstyle="secondary",
        ).pack(side=LEFT, padx=(8, 0))

        range_row = ttk.Frame(box)
        range_row.pack(fill=X, pady=(8, 0))
        ttk.Label(range_row, text="题号范围", width=10).pack(side=LEFT)
        ttk.Entry(range_row, textvariable=self.pdf_question_range).pack(
            side=LEFT, fill=X, expand=YES, padx=(0, 8)
        )
        ttk.Label(
            range_row,
            text="例如 66-85,111-115；留空表示导出当前科目全部题目",
            font=("", 9),
            bootstyle="secondary",
        ).pack(side=LEFT)

        self._pdf_settings_summary = ttk.Label(
            box,
            wraplength=840,
            justify=LEFT,
            bootstyle="secondary",
        )
        self._pdf_settings_summary.pack(anchor=W, pady=(10, 0))

        action_row = ttk.Frame(box)
        action_row.pack(fill=X, pady=(12, 0))
        ttk.Button(
            action_row,
            text="生成预览并进入下一步",
            command=self._start_pdf_preview_step,
            bootstyle="primary",
            width=18,
        ).pack(side=RIGHT)

    def _build_pdf_step_preview(self, parent):
        ttk.Label(
            parent,
            text="第三步会把当前筛选结果整理成题目工程。你可以直接修改题号、材料标题、题目归属，导出时会复用当前工程。",
            wraplength=860,
            justify=LEFT,
            font=("", 9),
            bootstyle="secondary",
        ).pack(anchor=W, pady=(0, 8), padx=12)
        self._pdf_preview_summary = ttk.Label(
            parent,
            wraplength=860,
            justify=LEFT,
            bootstyle="secondary",
        )
        self._pdf_preview_summary.pack(anchor=W, pady=(0, 8), padx=12)
        self._build_pdf_preview(parent)

    def _build_pdf_step_export(self, parent):
        summary = ttk.Labelframe(parent, text=" 第四步：导出结果 ", bootstyle="info", padding=_PAD)
        summary.pack(fill=X, pady=(0, 8))
        self._pdf_export_summary = ttk.Label(
            summary,
            wraplength=840,
            justify=LEFT,
            bootstyle="secondary",
        )
        self._pdf_export_summary.pack(anchor=W)

        box = ttk.Labelframe(parent, text=" 导出文件 ", bootstyle="primary", padding=_PAD)
        box.pack(fill=X, pady=(0, 8))

        self._file_row(box, "题本 Word", self.pdf_word_out, self._browse_pdf_word, "另存为...", pady=(0, 0))
        self._file_row(box, "授课 PPT", self.pdf_ppt_out, self._browse_pdf_ppt, "另存为...", pady=(6, 0))
        self._file_row(box, "工程 JSON", self.pdf_manifest_out, self._browse_pdf_manifest, "另存为...", pady=(6, 0))

        template_row = ttk.Frame(box)
        template_row.pack(fill=X, pady=(10, 0))
        ttk.Label(template_row, text="PPT 设置", width=10).pack(side=LEFT)
        ttk.Label(
            template_row,
            text="模板、字号、颜色和布局统一在“PPT 导出设置”标签页调整。",
            bootstyle="secondary",
        ).pack(side=LEFT, fill=X, expand=YES)
        ttk.Button(
            template_row,
            text="打开设置",
            command=self._open_ppt_settings_tab,
            bootstyle="info-outline",
            width=10,
        ).pack(side=RIGHT)

        action_row = ttk.Frame(parent, padding=(0, 4))
        action_row.pack(fill=X)
        ttk.Button(
            action_row,
            text="重新预览",
            command=self._preview_pdf_project,
            bootstyle="secondary-outline",
            width=12,
        ).pack(side=LEFT)
        ttk.Button(
            action_row,
            text="仅导出 Word",
            command=self._export_pdf_word,
            bootstyle="info-outline",
            width=12,
        ).pack(side=RIGHT, padx=(4, 0))
        ttk.Button(
            action_row,
            text="仅导出 PPT",
            command=self._export_pdf_ppt,
            bootstyle="success-outline",
            width=12,
        ).pack(side=RIGHT, padx=4)
        ttk.Button(
            action_row,
            text="导出 Word + PPT",
            command=self._export_pdf_bundle,
            bootstyle="success",
            width=16,
        ).pack(side=RIGHT, padx=4)

    def _build_progress_footer(self, parent):
        bar = ttk.Frame(parent, padding=(16, 10))
        bar.pack(side=BOTTOM, fill=X)

        self.progress = ttk.Progressbar(bar, mode="determinate", bootstyle="success-striped")
        self.progress.pack(fill=X, pady=(0, 8))

        row = ttk.Frame(bar)
        row.pack(fill=X)
        self.status_label = ttk.Label(row, text="就绪", anchor=W, bootstyle="secondary")
        self.status_label.pack(side=LEFT, fill=X, expand=YES)

    def _open_ppt_settings_tab(self):
        notebook = getattr(self, "_workspace_notebook", None)
        settings_tab = getattr(self, "_ppt_settings_tab", None)
        if notebook is not None and settings_tab is not None:
            notebook.select(settings_tab)

    def _open_pdf_preview_workspace(self):
        notebook = getattr(self, "_workspace_notebook", None)
        pdf_tab = getattr(self, "_pdf_workspace_tab", None)
        if notebook is not None and pdf_tab is not None:
            notebook.select(pdf_tab)
        self._show_pdf_wizard_step(2)

    def _format_question_ranges_for_gui(self, ranges) -> str:
        parts: list[str] = []
        for question_range in ranges or []:
            start = getattr(question_range, "start", None)
            end = getattr(question_range, "end", None)
            if start is None or end is None:
                continue
            if int(start) == int(end):
                parts.append(str(int(start)))
            else:
                parts.append(f"{int(start)}-{int(end)}")
        return ",".join(parts)

    def _document_subject_key(self, raw_value: str) -> str:
        normalized = (raw_value or "").strip()
        if normalized in {"unknown", "待确认科目", "未知科目"}:
            return "unknown"
        for key, label in _DOCUMENT_SUBJECT_CHOICES:
            if normalized == key or normalized == label:
                return key
        return "auto"

    def _document_subject_label(self, raw_value: str) -> str:
        key = self._document_subject_key(raw_value)
        if key == "unknown":
            return "待确认科目"
        return _DOCUMENT_SUBJECT_LABELS.get(key, _DOCUMENT_SUBJECT_LABELS["auto"])

    def _apply_project_subject_selection(self, subjects) -> None:
        normalized = [str(subject) for subject in (subjects or []) if str(subject)]
        if not normalized:
            self._set_all_pdf_subjects(True)
            return
        selected = set(normalized)
        for kind, var in self._pdf_subject_vars.items():
            var.set(kind in selected)

    def _default_pdf_base_path(self) -> str:
        for candidate in (
            self.pdf_path.get().strip(),
            self.word_path.get().strip(),
            self.pdf_manifest_out.get().strip(),
            self._pdf_project_context.get("docx_path", ""),
            self._pdf_project_context.get("manifest_path", ""),
        ):
            if candidate:
                return os.path.splitext(candidate)[0]
        return os.path.join(os.getcwd(), "工程")

    def _selected_pdf_subjects(self) -> list[str]:
        return [
            kind
            for kind in _PDF_SUBJECT_ORDER
            if self._pdf_subject_vars[kind].get()
        ]

    def _current_pdf_subject_spec(self) -> str:
        selected = self._selected_pdf_subjects()
        if not selected:
            return ""
        if tuple(selected) == _PDF_SUBJECT_ORDER:
            return "all"
        return ",".join(selected)

    def _selected_pdf_subject_labels(self) -> str:
        selected = self._selected_pdf_subjects()
        if not selected:
            return "未选择科目"
        return " / ".join(SUBJECT_DISPLAY_NAMES.get(kind, kind) for kind in selected)

    def _set_all_pdf_subjects(self, selected: bool):
        for var in self._pdf_subject_vars.values():
            var.set(selected)

    def _bind_pdf_wizard_updates(self):
        watched = [
            self.pdf_path,
            self.pdf_question_range,
            self.pdf_word_out,
            self.pdf_ppt_out,
            self.pdf_manifest_out,
            self.template_path,
            self.pdf_document_subject,
            self.word_document_subject,
        ]
        for var in watched:
            try:
                var.trace_add("write", lambda *_: self._refresh_pdf_wizard_ui())
            except Exception:
                pass
        for var in self._pdf_subject_vars.values():
            try:
                var.trace_add("write", lambda *_: self._refresh_pdf_wizard_ui())
            except Exception:
                pass

    def _pdf_project_matches_current_selection(self) -> bool:
        if self.pdf_project is None:
            return False
        source_kind = self._pdf_project_context.get("source_kind", "")
        if source_kind in {"manifest", "word"}:
            return True
        return (
            self._pdf_project_context.get("pdf_path") == self.pdf_path.get().strip()
            and self._pdf_project_context.get("subject_spec", "all") == self._current_pdf_subject_spec()
            and self._pdf_project_context.get("range_spec", "") == self.pdf_question_range.get().strip()
            and self._pdf_project_context.get("document_subject_hint", "auto")
            == self._document_subject_key(self.pdf_document_subject.get())
        )

    def _pdf_can_enter_step(self, index: int, *, show_message: bool) -> bool:
        if index <= 0:
            return True

        if self._pdf_project_matches_current_selection():
            return True

        pdf_file = self.pdf_path.get().strip()
        if not pdf_file:
            if show_message:
                messagebox.showwarning("提示", "请先在第一步选择 PDF 文件")
            return False
        if not os.path.exists(pdf_file):
            if show_message:
                messagebox.showerror("错误", f"文件不存在：{pdf_file}")
            return False
        if not self._selected_pdf_subjects():
            if show_message:
                messagebox.showwarning("提示", "请至少选择一个科目")
            return False

        if index >= 2 and not self._pdf_project_matches_current_selection():
            if show_message:
                messagebox.showinfo("提示", "请先在第二步生成当前设置对应的预览")
            return False
        return True

    def _request_pdf_wizard_step(self, index: int):
        if index == self._pdf_wizard_step:
            return
        if index < self._pdf_wizard_step:
            self._show_pdf_wizard_step(index)
            return
        if index == 1:
            if self._pdf_can_enter_step(index, show_message=True):
                self._show_pdf_wizard_step(index)
            return
        if index == 2:
            self._start_pdf_preview_step()
            return
        if self._pdf_can_enter_step(index, show_message=True):
            self._show_pdf_wizard_step(index)

    def _show_pdf_wizard_step(self, index: int):
        if not self._pdf_step_frames:
            return
        self._pdf_wizard_step = max(0, min(index, len(self._pdf_step_frames) - 1))
        for step_index, frame in enumerate(self._pdf_step_frames):
            frame.pack_forget()
            if step_index == self._pdf_wizard_step:
                frame.pack(fill=BOTH, expand=YES)
        self._refresh_pdf_wizard_ui()

    def _go_prev_pdf_step(self):
        if self._pdf_wizard_step > 0:
            self._show_pdf_wizard_step(self._pdf_wizard_step - 1)

    def _go_next_pdf_step(self):
        if self._pdf_wizard_step == 0:
            if self._pdf_can_enter_step(1, show_message=True):
                self._show_pdf_wizard_step(1)
        elif self._pdf_wizard_step == 1:
            self._start_pdf_preview_step()
        elif self._pdf_wizard_step == 2:
            if self._pdf_can_enter_step(3, show_message=True):
                self._show_pdf_wizard_step(3)

    def _start_pdf_preview_step(self):
        if not self._pdf_can_enter_step(1, show_message=True):
            return
        self._pdf_wizard_pending_step = 2
        self._preview_pdf_project()

    def _refresh_pdf_wizard_ui(self):
        if not getattr(self, "_pdf_step_title_label", None):
            return

        title, hint = _PDF_WIZARD_STEPS[self._pdf_wizard_step]
        self._pdf_step_title_label.configure(text=f"第 {self._pdf_wizard_step + 1} 步 · {title}")
        self._pdf_step_hint_label.configure(text=hint)

        for index, button in enumerate(getattr(self, "_pdf_step_buttons", [])):
            if index == self._pdf_wizard_step:
                style = "primary"
            elif index == 0 and (
                self.pdf_path.get().strip()
                or self._pdf_project_context.get("source_kind", "") == "word"
            ):
                style = "success-outline"
            elif index in (1, 2) and self._pdf_project_matches_current_selection():
                style = "success-outline"
            else:
                style = "secondary-outline"
            button.configure(bootstyle=style)

        self._pdf_prev_btn.configure(state=NORMAL if self._pdf_wizard_step > 0 else DISABLED)

        if self._pdf_wizard_step == 0:
            self._pdf_next_btn.configure(text="下一步：识别设置", state=NORMAL, bootstyle="primary")
        elif self._pdf_wizard_step == 1:
            self._pdf_next_btn.configure(text="生成预览", state=NORMAL, bootstyle="primary")
        elif self._pdf_wizard_step == 2:
            state = NORMAL if self._pdf_project_matches_current_selection() else DISABLED
            self._pdf_next_btn.configure(text="下一步：导出结果", state=state, bootstyle="success")
        else:
            self._pdf_next_btn.configure(text="已到最后一步", state=DISABLED, bootstyle="secondary")

        range_text = self.pdf_question_range.get().strip() or "当前科目全部题目"
        pdf_file = self.pdf_path.get().strip()
        manifest_file = self._pdf_project_context.get("manifest_path", "")
        source_kind = self._pdf_project_context.get("source_kind", "")
        docx_file = self._pdf_project_context.get("docx_path", "")
        if source_kind == "word" and docx_file:
            base_name = os.path.basename(docx_file)
        elif pdf_file:
            base_name = os.path.basename(pdf_file)
        elif manifest_file:
            base_name = os.path.basename(manifest_file)
        else:
            base_name = "未选择来源文件"
        subject_text = self._selected_pdf_subject_labels()
        effective_document_subject = (
            self._pdf_project_context.get("document_subject_hint", self.pdf_document_subject.get())
            if source_kind in {"word", "manifest"}
            else self.pdf_document_subject.get()
        )
        document_subject_text = self._document_subject_label(effective_document_subject)
        preview_ready = self._pdf_project_matches_current_selection()
        if preview_ready:
            preview_state = f"当前预览已就绪，共 {self.pdf_project.question_count} 道题。"
        elif self.pdf_project is not None:
            preview_state = "已有旧预览，但与当前设置不一致，请重新生成。"
        else:
            preview_state = "尚未生成预览。"

        import_summary = "未选择 PDF 文件。"
        if source_kind == "word" and docx_file:
            docx_base = os.path.splitext(docx_file)[0]
            import_summary = (
                f"当前 Word：{os.path.basename(docx_file)}\n"
                f"素材目录：{self._pdf_project_context.get('asset_dir', '-')}\n"
                f"整份科目：{self._document_subject_label(self._pdf_project_context.get('document_subject_hint', 'auto'))}\n"
                f"默认课件：{self.pdf_ppt_out.get().strip() or docx_base + '.pptx'}\n"
                f"默认工程：{self.pdf_manifest_out.get().strip() or docx_base + '_工程.json'}"
            )
        elif pdf_file:
            import_summary = (
                f"当前试卷：{base_name}\n"
                f"整份科目：{document_subject_text}\n"
                f"默认题本：{self.pdf_word_out.get().strip() or os.path.splitext(pdf_file)[0] + '_真题.docx'}\n"
                f"默认课件：{self.pdf_ppt_out.get().strip() or os.path.splitext(pdf_file)[0] + '_授课.pptx'}\n"
                f"默认工程：{self.pdf_manifest_out.get().strip() or os.path.splitext(pdf_file)[0] + '_工程.json'}"
            )
        elif source_kind == "manifest" and manifest_file:
            manifest_base = os.path.splitext(manifest_file)[0]
            import_summary = (
                f"当前工程：{os.path.basename(manifest_file)}\n"
                f"来源 PDF：{self.pdf_path.get().strip() or '未记录 / 不可用'}\n"
                f"默认题本：{self.pdf_word_out.get().strip() or manifest_base + '_真题.docx'}\n"
                f"默认课件：{self.pdf_ppt_out.get().strip() or manifest_base + '_授课.pptx'}"
            )
        if getattr(self, "_pdf_import_summary", None):
            self._pdf_import_summary.configure(text=import_summary)

        source_label = "当前试卷"
        if source_kind == "word" and docx_file:
            source_label = "当前 Word"
            project_subjects = [
                SUBJECT_DISPLAY_NAMES.get(section.kind, section.kind)
                for section in getattr(self.pdf_project, "sections", [])
                if section.kind in ALL_SUBJECT_KINDS
            ]
            subject_text = "、".join(project_subjects) if project_subjects else "未识别科目"
            range_text = "Word 全部题目"
        settings_summary = (
            f"{source_label}：{base_name}\n"
            f"整份科目：{document_subject_text}\n"
            f"处理科目：{subject_text}\n"
            f"题号范围：{range_text}\n"
            f"{preview_state}"
        )
        if getattr(self, "_pdf_settings_summary", None):
            self._pdf_settings_summary.configure(text=settings_summary)

        preview_summary = (
            f"{source_label}：{base_name}\n"
            f"整份科目：{document_subject_text}\n"
            f"筛选：{subject_text}；题号范围 {range_text}\n"
            f"{preview_state}"
        )
        if getattr(self, "_pdf_preview_summary", None):
            self._pdf_preview_summary.configure(text=preview_summary)

        if preview_ready:
            asset_dir = self._pdf_project_context.get("asset_dir", "-")
            export_summary = (
                f"当前工程共 {self.pdf_project.question_count} 道题，素材目录：{asset_dir}\n"
                "导出会直接复用当前预览中的人工修改。"
            )
        else:
            export_summary = "请先在上一步生成当前设置对应的预览工程。"
        if getattr(self, "_pdf_export_summary", None):
            self._pdf_export_summary.configure(text=export_summary)

        if getattr(self, "_pdf_nav_status_label", None):
            self._pdf_nav_status_label.configure(text=preview_state)

    def _build_pdf_preview(self, parent):
        frame = ttk.Labelframe(parent, text=" 结果预览 ", bootstyle="primary", padding=(10, 8))
        frame.pack(fill=BOTH, expand=YES, pady=(6, 0))

        split = ttk.Panedwindow(frame, orient=HORIZONTAL)
        split.pack(fill=BOTH, expand=YES)

        left = ttk.Frame(split)
        right = ttk.Frame(split)
        split.add(left, weight=2)
        split.add(right, weight=5)

        cols = ("kind", "source", "count")
        self.pdf_tree = ttk.Treeview(
            left,
            columns=cols,
            show="tree headings",
            height=12,
            bootstyle="info",
        )
        self.pdf_tree.heading("#0", text="节点")
        self.pdf_tree.heading("kind", text="类型")
        self.pdf_tree.heading("source", text="题号")
        self.pdf_tree.heading("count", text="数量")
        self.pdf_tree.column("#0", width=280)
        self.pdf_tree.column("kind", width=70, anchor=CENTER)
        self.pdf_tree.column("source", width=70, anchor=CENTER)
        self.pdf_tree.column("count", width=60, anchor=CENTER)
        self.pdf_tree.pack(side=LEFT, fill=BOTH, expand=YES)
        sb = ttk.Scrollbar(left, orient=VERTICAL, command=self.pdf_tree.yview)
        self.pdf_tree.configure(yscrollcommand=sb.set)
        sb.pack(side=RIGHT, fill=Y)
        self.pdf_tree.bind("<<TreeviewSelect>>", self._on_pdf_preview_select)

        action_box = ttk.Frame(right)
        action_box.pack(fill=X, pady=(0, 6))
        action_row = ttk.Frame(action_box)
        action_row.pack(fill=X)
        ttk.Button(
            action_row,
            text="重新生成预览",
            command=self._preview_pdf_project,
            bootstyle="secondary-outline",
            width=12,
        ).pack(side=LEFT, padx=(0, 4))
        ttk.Button(
            action_row,
            text="改题号",
            command=self._edit_selected_question_number,
            bootstyle="info-outline",
            width=8,
        ).pack(side=LEFT, padx=4)
        ttk.Button(
            action_row,
            text="删题",
            command=self._remove_selected_question,
            bootstyle="danger-outline",
            width=7,
        ).pack(side=LEFT, padx=4)
        ttk.Label(action_row, text="整段改科目", bootstyle="secondary").pack(side=LEFT, padx=(16, 4))
        ttk.Combobox(
            action_row,
            textvariable=self._pdf_section_subject_var,
            values=[label for key, label in _DOCUMENT_SUBJECT_CHOICES if key != "data"] + ["待确认科目"],
            state="readonly",
            width=14,
        ).pack(side=LEFT, padx=(0, 4))
        ttk.Button(
            action_row,
            text="应用",
            command=self._reclassify_selected_section,
            bootstyle="info-outline",
            width=7,
        ).pack(side=LEFT, padx=(0, 4))
        action_row2 = ttk.Frame(action_box)
        action_row2.pack(fill=X, pady=(6, 0))
        ttk.Button(
            action_row2,
            text="材料改名",
            command=self._rename_selected_material,
            bootstyle="warning-outline",
            width=10,
        ).pack(side=LEFT, padx=4)
        ttk.Button(
            action_row2,
            text="下方新建材料",
            command=self._insert_material_after_selection,
            bootstyle="warning-outline",
            width=12,
        ).pack(side=LEFT, padx=4)
        ttk.Button(
            action_row2,
            text="并入上一材料",
            command=lambda: self._merge_selected_material(-1),
            bootstyle="warning-outline",
            width=12,
        ).pack(side=LEFT, padx=4)
        ttk.Button(
            action_row2,
            text="并入下一材料",
            command=lambda: self._merge_selected_material(1),
            bootstyle="warning-outline",
            width=12,
        ).pack(side=LEFT, padx=4)
        ttk.Button(
            action_row2,
            text="移到上一材料",
            command=lambda: self._move_selected_question_between_materials(-1),
            bootstyle="secondary-outline",
            width=12,
        ).pack(side=LEFT, padx=4)
        ttk.Button(
            action_row2,
            text="移到下一材料",
            command=lambda: self._move_selected_question_between_materials(1),
            bootstyle="secondary-outline",
            width=12,
        ).pack(side=LEFT, padx=4)

        detail_tabs = ttk.Notebook(right, bootstyle="info")
        detail_tabs.pack(fill=BOTH, expand=YES)

        editor_tab = ttk.Frame(detail_tabs, padding=(0, 0, 0, 0))
        material_tab = ttk.Frame(detail_tabs, padding=(0, 0, 0, 0))
        detail_tab = ttk.Frame(detail_tabs, padding=(0, 0, 0, 0))
        detail_tabs.add(editor_tab, text=" 题目编辑 ")
        detail_tabs.add(material_tab, text=" 材料原貌 ")
        detail_tabs.add(detail_tab, text=" 结构详情 ")

        self._build_pdf_question_editor(editor_tab)
        self._build_pdf_material_preview_panel(material_tab)

        detail_host = ttk.Frame(detail_tab)
        detail_host.pack(fill=BOTH, expand=YES)
        self.pdf_detail = tk.Text(detail_host, wrap="word", height=16)
        self.pdf_detail.pack(side=LEFT, fill=BOTH, expand=YES)
        detail_scroll = ttk.Scrollbar(detail_host, orient=VERTICAL, command=self.pdf_detail.yview)
        detail_scroll.pack(side=RIGHT, fill=Y)
        self.pdf_detail.configure(yscrollcommand=detail_scroll.set)
        self.pdf_detail.configure(state="disabled")

    def _build_pdf_question_editor(self, parent):
        ttk.Label(
            parent,
            textvariable=self._pdf_question_editor_message,
            wraplength=440,
            justify=LEFT,
            bootstyle="secondary",
        ).pack(anchor=W, pady=(0, 8))

        layout_box = ttk.Labelframe(parent, text=" 单题选项布局 ", bootstyle="secondary", padding=(8, 6))
        layout_box.pack(fill=X, pady=(0, 8))
        ttk.Label(
            layout_box,
            text="保持“跟随全局”时，会继续使用第四步里的全局排版；单独切换后，只影响当前这道题。",
            wraplength=430,
            justify=LEFT,
            bootstyle="secondary",
        ).pack(anchor=W, pady=(0, 6))
        button_row = ttk.Frame(layout_box)
        button_row.pack(fill=X)
        for value, label in _PDF_QUESTION_LAYOUT_CHOICES:
            button = ttk.Radiobutton(
                button_row,
                text=label,
                value=value,
                variable=self._pdf_question_layout_var,
            )
            button.pack(side=LEFT, padx=(0, 8))
            self._pdf_question_layout_buttons.append(button)

        stem_box = ttk.Labelframe(parent, text=" 题干编辑 ", bootstyle="info", padding=(8, 6))
        stem_box.pack(fill=X, pady=(0, 8))
        ttk.Label(
            stem_box,
            text="这里的修改会直接进入当前工程，后续导出 Word / PPT 会复用当前版本。",
            wraplength=430,
            justify=LEFT,
            bootstyle="secondary",
        ).pack(anchor=W, pady=(0, 6))
        stem_editor_host = ttk.Frame(stem_box)
        stem_editor_host.pack(fill=X, expand=YES)
        self._pdf_question_stem_editor = tk.Text(stem_editor_host, wrap="word", height=7)
        self._pdf_question_stem_editor.pack(side=LEFT, fill=BOTH, expand=YES)
        stem_scroll = ttk.Scrollbar(stem_editor_host, orient=VERTICAL, command=self._pdf_question_stem_editor.yview)
        stem_scroll.pack(side=RIGHT, fill=Y)
        self._pdf_question_stem_editor.configure(yscrollcommand=stem_scroll.set)
        self._pdf_question_stem_editor.bind("<KeyRelease>", self._on_pdf_question_stem_change)
        self._pdf_question_stem_editor.bind("<FocusOut>", self._on_pdf_question_stem_change)

        stem_preview_box = ttk.Labelframe(parent, text=" 题干图片 ", bootstyle="secondary", padding=(8, 6))
        stem_preview_box.pack(fill=X, pady=(0, 8))
        stem_preview_nav = ttk.Frame(stem_preview_box)
        stem_preview_nav.pack(fill=X, pady=(0, 6))
        self._pdf_stem_preview_status = ttk.Label(
            stem_preview_nav,
            text="当前题目没有题干图片。",
        )
        self._pdf_stem_preview_status.pack(side=LEFT, fill=X, expand=YES)
        self._pdf_stem_preview_prev = ttk.Button(
            stem_preview_nav,
            text="上一张",
            command=lambda: self._step_pdf_stem_preview(-1),
            width=8,
            bootstyle="secondary-outline",
            state=DISABLED,
        )
        self._pdf_stem_preview_prev.pack(side=LEFT, padx=(4, 4))
        self._pdf_stem_preview_next = ttk.Button(
            stem_preview_nav,
            text="下一张",
            command=lambda: self._step_pdf_stem_preview(1),
            width=8,
            bootstyle="secondary-outline",
            state=DISABLED,
        )
        self._pdf_stem_preview_next.pack(side=LEFT, padx=(0, 4))
        self._pdf_stem_preview_open = ttk.Button(
            stem_preview_nav,
            text="打开原图",
            command=self._open_pdf_stem_preview_image,
            width=10,
            bootstyle="info-outline",
            state=DISABLED,
        )
        self._pdf_stem_preview_open.pack(side=LEFT)
        self._pdf_stem_preview_box = ttk.Label(
            stem_preview_box,
            text="当前题目没有题干图片。",
            anchor=CENTER,
            justify=CENTER,
        )
        self._pdf_stem_preview_box.pack(fill=X, expand=YES, ipady=40)
        self._pdf_stem_preview_box.bind(
            "<Configure>",
            lambda _event: self._render_pdf_stem_preview(),
        )

        option_box = ttk.Labelframe(parent, text=" 选项编辑 ", bootstyle="info", padding=(8, 6))
        option_box.pack(fill=X, pady=(0, 8))
        ttk.Label(
            option_box,
            text="可逐项修改 A/B/C/D 文本；如果选项里带图，也可以查看、替换或清除当前图片。",
            wraplength=430,
            justify=LEFT,
            bootstyle="secondary",
        ).pack(anchor=W, pady=(0, 6))
        self._pdf_option_editor_host = ttk.Frame(option_box)
        self._pdf_option_editor_host.pack(fill=X, expand=YES)

        preview_box = ttk.Labelframe(parent, text=" 实时排版预览 ", bootstyle="primary", padding=(8, 6))
        preview_box.pack(fill=BOTH, expand=YES)
        ttk.Label(
            preview_box,
            text="基于当前题干、选项和布局即时刷新，用来快速判断一行 / 两行 / 四行效果。",
            wraplength=430,
            justify=LEFT,
            bootstyle="secondary",
        ).pack(anchor=W, pady=(0, 6))
        self._pdf_question_preview_canvas = tk.Canvas(
            preview_box,
            height=300,
            bg="#edf2f7",
            highlightthickness=0,
        )
        self._pdf_question_preview_canvas.pack(fill=BOTH, expand=YES)
        self._pdf_question_preview_canvas.bind(
            "<Configure>",
            lambda _event: self._render_pdf_question_editor_preview(),
        )
        self._clear_pdf_question_editor()

    def _build_pdf_material_preview_panel(self, parent):
        ttk.Label(
            parent,
            text="资料分析材料会优先显示 PDF 区域原貌；如果区域截图不可用，会回退到材料图片。",
            wraplength=440,
            justify=LEFT,
            bootstyle="secondary",
        ).pack(anchor=W, pady=(0, 8))

        preview_box = ttk.Labelframe(parent, text=" 材料预览 ", bootstyle="secondary", padding=(8, 6))
        preview_box.pack(fill=BOTH, expand=YES)
        preview_nav = ttk.Frame(preview_box)
        preview_nav.pack(fill=X, pady=(0, 6))
        self._pdf_material_preview_status = ttk.Label(
            preview_nav,
            text="选择资料分析材料或题目后，可查看 PDF 区域原貌。",
        )
        self._pdf_material_preview_status.pack(side=LEFT, fill=X, expand=YES)
        self._pdf_material_preview_prev = ttk.Button(
            preview_nav,
            text="上一张",
            command=lambda: self._step_pdf_material_preview(-1),
            width=8,
            bootstyle="secondary-outline",
            state=DISABLED,
        )
        self._pdf_material_preview_prev.pack(side=LEFT, padx=(4, 4))
        self._pdf_material_preview_next = ttk.Button(
            preview_nav,
            text="下一张",
            command=lambda: self._step_pdf_material_preview(1),
            width=8,
            bootstyle="secondary-outline",
            state=DISABLED,
        )
        self._pdf_material_preview_next.pack(side=LEFT)
        self._pdf_material_preview_box = ttk.Label(
            preview_box,
            text="暂无材料原貌",
            anchor=CENTER,
            justify=CENTER,
        )
        self._pdf_material_preview_box.pack(fill=BOTH, expand=YES, ipady=80)
        self._pdf_material_preview_box.bind(
            "<Configure>",
            lambda _event: self._render_pdf_material_preview(),
        )

    # Tabs

    def _spin_row(self, parent, label, var, lo, hi, step, width=6):
        f = ttk.Frame(parent)
        f.pack(fill=X, pady=2)
        ttk.Label(f, text=label, width=22).pack(side=LEFT)
        ttk.Spinbox(f, from_=lo, to=hi, increment=step, textvariable=var,
                    width=width, format="%.2f").pack(side=LEFT)

    def _radio_row(self, parent, label, var, choices):
        f = ttk.Frame(parent)
        f.pack(fill=X, pady=3)
        ttk.Label(f, text=label, width=22).pack(side=LEFT)
        for val, txt in choices:
            ttk.Radiobutton(f, text=txt, variable=var, value=val).pack(side=LEFT, padx=4)

    # tab: layout

    def _build_tab_layout(self, p):
        ttk.Label(p, text="页面边距（英寸）", font=("", 9, "bold")).pack(anchor=W)
        self._spin_row(p, "左边距", self.margin_left, 0.1, 2.0, 0.05)
        self._spin_row(p, "右边距", self.margin_right, 0.1, 2.0, 0.05)
        self._spin_row(p, "上边距", self.margin_top, 0.1, 2.0, 0.05)

        ttk.Separator(p).pack(fill=X, pady=6)
        ttk.Label(p, text="题干区高度（英寸）", font=("", 9, "bold")).pack(anchor=W)
        self._spin_row(p, "有图时", self.stem_h_img, 0.5, 4.0, 0.1)
        self._spin_row(p, "无图时", self.stem_h_no, 0.5, 5.0, 0.1)

        ttk.Separator(p).pack(fill=X, pady=6)
        ttk.Label(p, text="间距（英寸）", font=("", 9, "bold")).pack(anchor=W)
        self._spin_row(p, "题干到下方", self.gap_stem, 0.0, 1.0, 0.05)
        self._spin_row(p, "图片间距", self.gap_img, 0.0, 1.0, 0.05)
        self._spin_row(p, "图片到选项区", self.gap_opts, 0.0, 1.5, 0.05)

        ttk.Separator(p).pack(fill=X, pady=6)
        self._radio_row(p, "题干对齐", self.stem_align,
                        [("left", "左"), ("center", "中"), ("right", "右")])
        self._radio_row(p, "图片位置", self.image_align,
                        [("left", "左"), ("center", "中"), ("right", "右")])
        self._spin_row(p, "图片最大宽（英寸）", self.image_max_w, 1.0, 12.0, 0.1)
        self._spin_row(p, "图片最大高（英寸）", self.image_max_h, 0.5, 6.0, 0.1)

    # tab: options

    def _build_tab_options(self, p):
        self._radio_row(p, "排列方式", self.option_layout,
                        [("grid", "2x2 网格"), ("list", "竖排"), ("one_row", "一行四项")])

        ttk.Separator(p).pack(fill=X, pady=6)
        ttk.Label(p, text="网格", font=("", 9, "bold")).pack(anchor=W)
        f = ttk.Frame(p)
        f.pack(fill=X, pady=3)
        ttk.Radiobutton(f, text="AB / CD", variable=self.grid_layout, value="ab_cd").pack(side=LEFT, padx=4)
        ttk.Radiobutton(f, text="AC / BD", variable=self.grid_layout, value="ac_bd").pack(side=LEFT, padx=4)
        self._spin_row(p, "网格行高", self.grid_row_h, 0.4, 2.0, 0.05)
        self._spin_row(p, "网格列间距", self.grid_col_gap, 0.0, 1.0, 0.05)

        ttk.Separator(p).pack(fill=X, pady=6)
        ttk.Label(p, text="竖排 / 一行", font=("", 9, "bold")).pack(anchor=W)
        self._spin_row(p, "列表行高", self.list_row_h, 0.4, 2.0, 0.05)
        self._spin_row(p, "一行四列行高", self.one_row_h, 0.35, 1.5, 0.05)
        self._spin_row(p, "一行四列间距", self.one_row_gap, 0.0, 0.5, 0.02)

        ttk.Separator(p).pack(fill=X, pady=6)
        self._radio_row(p, "选项文字对齐", self.option_align,
                        [("left", "左"), ("center", "中"), ("right", "右")])

    # tab: font & color

    def _build_tab_font(self, p):
        # font family
        r0 = ttk.Frame(p)
        r0.pack(fill=X, pady=4)
        ttk.Label(r0, text="字体", width=14).pack(side=LEFT)
        self._font_combo = ttk.Combobox(r0, textvariable=self.font_name,
                                        values=self._font_list, width=24, state="readonly")
        self._font_combo.pack(side=LEFT, padx=4)
        self._font_combo.bind("<<ComboboxSelected>>", lambda e: self._schedule_preview_refresh())

        # sizes
        r1 = ttk.Frame(p)
        r1.pack(fill=X, pady=4)
        ttk.Label(r1, text="题干字号", width=14).pack(side=LEFT)
        ttk.Spinbox(r1, from_=10, to=48, textvariable=self.font_size_stem, width=5).pack(side=LEFT, padx=4)
        ttk.Label(r1, text="选项字号").pack(side=LEFT, padx=(16, 0))
        ttk.Spinbox(r1, from_=8, to=40, textvariable=self.font_size_option, width=5).pack(side=LEFT, padx=4)

        # bold
        r2 = ttk.Frame(p)
        r2.pack(fill=X, pady=4)
        ttk.Checkbutton(r2, text="题干加粗", variable=self.stem_bold).pack(side=LEFT, padx=(0, 12))
        ttk.Checkbutton(r2, text="字母加粗", variable=self.option_letter_bold).pack(side=LEFT, padx=(0, 12))
        ttk.Checkbutton(r2, text="正文加粗", variable=self.option_text_bold).pack(side=LEFT)

        # live preview
        ttk.Separator(p).pack(fill=X, pady=8)
        pv = ttk.Labelframe(p, text=" 字体预览 ", bootstyle="info", padding=10)
        pv.pack(fill=X, pady=(0, 6))
        self._pv_stem = tk.Label(pv, text="1. 示例题干（2020·上海）", anchor=W, justify=LEFT, wraplength=500)
        self._pv_stem.pack(fill=X)
        self._pv_body = tk.Label(pv, text="选项正文：示例文本", anchor=W)
        self._pv_body.pack(fill=X, pady=(4, 0))
        self._pv_letter = tk.Label(pv, text="A.  B.  C.  D.", anchor=W)
        self._pv_letter.pack(fill=X, pady=(2, 0))

        # colors
        ttk.Separator(p).pack(fill=X, pady=8)
        ttk.Label(p, text="颜色（点击色块选色）", font=("", 9, "bold")).pack(anchor=W)
        self._swatch("题干", self.color_stem, p)
        self._swatch("选项正文", self.color_option, p)
        self._swatch("选项字母", self.color_letter, p)

        ttk.Button(p, text="恢复默认", command=self._reset_defaults,
                   bootstyle="secondary-outline").pack(anchor=E, pady=(10, 0))

    def _swatch(self, label, var, parent):
        row = ttk.Frame(parent)
        row.pack(fill=X, pady=4)
        ttk.Label(row, text=label, width=14).pack(side=LEFT)
        sw = tk.Frame(row, width=40, height=28, relief=tk.SOLID, bd=1)
        sw.pack(side=LEFT, padx=4)
        sw.pack_propagate(False)

        def pick():
            c = colorchooser.askcolor(color=var.get(), title=label)
            if c and c[1]:
                var.set(c[1].upper())

        sw.bind("<Button-1>", lambda e: pick())
        ttk.Button(row, text="选色", command=pick, bootstyle="info-outline", width=6).pack(side=LEFT, padx=4)

        def sync(*_):
            hx = var.get().strip()
            try:
                sw.configure(bg=hx[:7] if hx.startswith("#") and len(hx) >= 7 else "#ccc")
            except tk.TclError:
                sw.configure(bg="#ccc")

        var.trace_add("write", lambda *_: sync())
        sync()

    # tab: layout preview

    def _build_tab_preview(self, p):
        ctrl = ttk.Frame(p)
        ctrl.pack(fill=X, pady=(0, 4))
        ttk.Checkbutton(ctrl, text="显示图片占位", variable=self.preview_has_image,
                        bootstyle="round-toggle").pack(side=LEFT)
        ttk.Label(ctrl, text="（示意图，非真实比例）", font=("", 9),
                  bootstyle="secondary").pack(side=LEFT, padx=8)

        self._layout_canvas = tk.Canvas(p, width=560, height=315, bg="#f0f4f8",
                                        highlightthickness=0)
        self._layout_canvas.pack(pady=4)
        self._layout_canvas.bind("<Configure>", lambda e: self._schedule_preview_refresh())

    # Preview refresh

    def _bind_preview_updates(self):
        watched = [
            self.margin_left, self.margin_right, self.margin_top,
            self.stem_h_img, self.stem_h_no, self.gap_stem, self.gap_img, self.gap_opts,
            self.option_layout, self.grid_layout,
            self.grid_row_h, self.grid_col_gap, self.list_row_h,
            self.one_row_h, self.one_row_gap, self.preview_has_image,
            self.font_name, self.font_size_stem, self.font_size_option,
            self.color_stem, self.color_option, self.color_letter,
            self.stem_bold, self.option_letter_bold, self.option_text_bold,
        ]
        for v in watched:
            try:
                v.trace_add("write", lambda *_: self._schedule_preview_refresh())
            except Exception:
                pass

    def _schedule_preview_refresh(self, *_):
        if self._preview_after is not None:
            try:
                self.root.after_cancel(self._preview_after)
            except Exception:
                pass
        self._preview_after = self.root.after(80, self._do_refresh)

    def _do_refresh(self):
        self._preview_after = None
        self._refresh_layout_canvas()
        self._refresh_font_preview()
        self._refresh_pdf_question_editor_message()
        self._render_pdf_question_editor_preview()

    def _refresh_font_preview(self):
        if not getattr(self, "_pv_stem", None):
            return
        fn = self.font_name.get().strip() or "微软雅黑"
        sz_s = max(10, self.font_size_stem.get() - 4)
        sz_o = max(9, self.font_size_option.get() - 2)

        def font(sz, w):
            try:
                return tkfont.Font(family=fn, size=sz, weight=w)
            except tk.TclError:
                return tkfont.Font(size=sz, weight=w)

        cs = self.color_stem.get().strip()
        co = self.color_option.get().strip()
        cl = self.color_letter.get().strip()
        if not cs.startswith("#"): cs = "#1A1A2E"
        if not co.startswith("#"): co = "#2D2D2D"
        if not cl.startswith("#"): cl = "#006BBD"

        ws = "bold" if self.stem_bold.get() else "normal"
        wl = "bold" if self.option_letter_bold.get() else "normal"
        wb = "bold" if self.option_text_bold.get() else "normal"

        self._pv_stem.configure(font=font(sz_s, ws), fg=cs)
        self._pv_body.configure(font=font(sz_o, wb), fg=co)
        self._pv_letter.configure(font=font(sz_o, wl), fg=cl)

    def _refresh_layout_canvas(self):
        cvs = getattr(self, "_layout_canvas", None)
        if cvs is None:
            return
        cvs.delete("all")
        W = int(cvs.cget("width")) or 560
        H = int(cvs.cget("height")) or 315
        sx, sy = W / 13.333, H / 7.5

        ml = self.margin_left.get() * sx
        mt = self.margin_top.get() * sy
        mr = self.margin_right.get() * sx
        cw = W - ml - mr

        sh_in = self.stem_h_img.get() if self.preview_has_image.get() else self.stem_h_no.get()
        sh = sh_in * sy
        gap1 = self.gap_stem.get() * sy
        img_h = min(self.image_max_h.get(), 2.2) * sy * 0.35 if self.preview_has_image.get() else 0
        gap_img = self.gap_img.get() * sy if self.preview_has_image.get() else 0
        gap_opts = self.gap_opts.get() * sy

        y0 = mt
        y_stem_end = y0 + sh
        y_img_end = y_stem_end + gap1 + (img_h + gap_img if img_h else 0)
        oy = y_img_end + gap_opts
        oh = max(28, H - oy - 10)

        # slide bg
        cvs.create_rectangle(0, 0, W, H, fill="#eaeef3", outline="#b0b8c4", width=2)
        # stem
        cvs.create_rectangle(ml, y0, ml + cw, y_stem_end, fill="#d4e8ff", outline="#4a90d9", width=2)
        cvs.create_text(ml + 8, y0 + 8, anchor=tk.NW, text="题干", fill="#1a4d80", font=("", 10, "bold"))
        # image
        if self.preview_has_image.get() and img_h > 0:
            iy = y_stem_end + gap1
            cvs.create_rectangle(ml, iy, ml + cw, iy + img_h, fill="#fff8e6", outline="#d9a84a", width=2, dash=(4, 3))
            cvs.create_text(ml + cw / 2, iy + img_h / 2, text="图片", fill="#996600")
        # options
        ol = self.option_layout.get()
        ox, ow = ml, cw
        self._draw_options(cvs, ol, ox, oy, ow, oh, sx)
        cvs.create_text(W / 2, H - 5, text="示意图", fill="#999", font=("", 8))

    def _draw_options(self, cvs, layout, ox, oy, ow, oh, sx):
        fills = ["#e3f2fd", "#e8f5e9", "#fce4ec", "#fff3e0"]
        outlines = ["#1976d2", "#43a047", "#c62828", "#e65100"]
        labels = ["A", "B", "C", "D"]

        if layout == "list":
            rh = oh / 4
            for i in range(4):
                cvs.create_rectangle(ox, oy + i * rh, ox + ow, oy + (i + 1) * rh - 2,
                                     fill=fills[i], outline=outlines[i])
                cvs.create_text(ox + 8, oy + i * rh + 6, anchor=tk.NW,
                                text=f"选项 {labels[i]}", fill=outlines[i])
        elif layout == "one_row":
            g = self.one_row_gap.get() * sx
            cw4 = (ow - g * 3) / 4
            for i in range(4):
                x0 = ox + i * (cw4 + g)
                cvs.create_rectangle(x0, oy, x0 + cw4, oy + oh, fill=fills[i], outline=outlines[i])
                cvs.create_text(x0 + cw4 / 2, oy + oh / 2, text=labels[i],
                                fill=outlines[i], font=("", 11, "bold"))
        else:
            gh, gw = oh / 2, ow / 2
            gl = self.grid_layout.get()
            order = [("A", 0, 0), ("C", 1, 0), ("B", 0, 1), ("D", 1, 1)] if gl == "ac_bd" \
                else [("A", 0, 0), ("B", 1, 0), ("C", 0, 1), ("D", 1, 1)]
            for lab, gx, gy in order:
                ci = "ABCD".index(lab)
                x0, y0 = ox + gx * gw, oy + gy * gh
                cvs.create_rectangle(x0, y0, x0 + gw - 2, y0 + gh - 2,
                                     fill=fills[ci], outline=outlines[ci])
                cvs.create_text(x0 + 8, y0 + 6, anchor=tk.NW, text=lab,
                                fill=outlines[ci], font=("", 10, "bold"))

    # Template toggle

    def _toggle_template(self):
        on = self.use_template.get()
        st = NORMAL if on else DISABLED
        self.template_entry.configure(state=st)
        self.template_btn.configure(state=st)

        if on:
            self._tpl_hint.pack(anchor=W, pady=(8, 0))
            self._config_notebook.pack_forget()
            self._tpl_overlay_label.pack(fill=X, padx=8, pady=10)
        else:
            self._tpl_hint.pack_forget()
            self._tpl_overlay_label.pack_forget()
            self._config_notebook.pack(fill=X, padx=8, pady=(4, 8))

    # Defaults reset

    def _reset_defaults(self):
        pairs = [
            (self.margin_left, 0.8), (self.margin_right, 0.8), (self.margin_top, 0.5),
            (self.stem_h_img, 1.5), (self.stem_h_no, 2.5),
            (self.gap_stem, 0.2), (self.gap_img, 0.15), (self.gap_opts, 0.2),
            (self.stem_align, "left"), (self.image_align, "center"),
            (self.image_max_w, 5.0), (self.image_max_h, 2.5),
            (self.option_layout, "grid"), (self.grid_layout, "ab_cd"),
            (self.grid_row_h, 0.9), (self.grid_col_gap, 0.15), (self.list_row_h, 0.7),
            (self.one_row_h, 0.55), (self.one_row_gap, 0.06),
            (self.option_align, "left"), (self.font_name, "微软雅黑"),
            (self.font_size_stem, 20), (self.font_size_option, 18),
            (self.stem_bold, True), (self.option_letter_bold, True), (self.option_text_bold, False),
            (self.color_stem, "#1A1A2E"), (self.color_option, "#2D2D2D"), (self.color_letter, "#006BBD"),
        ]
        for var, val in pairs:
            var.set(val)
        self._schedule_preview_refresh()

    # File dialogs

    def _browse_word(self):
        path = filedialog.askopenfilename(title="选择 Word 文件",
                                          filetypes=[("Word", "*.docx"), ("All", "*.*")])
        if path:
            self.word_path.set(path)
            if not self.output_path.get():
                self.output_path.set(os.path.splitext(path)[0] + ".pptx")

    def _browse_output(self):
        path = filedialog.asksaveasfilename(title="保存 PPT",
                                            defaultextension=".pptx",
                                            filetypes=[("PowerPoint", "*.pptx")])
        if path:
            self.output_path.set(path)

    def _browse_template(self):
        path = filedialog.askopenfilename(title="选择 PPT 模板",
                                          filetypes=[("PowerPoint", "*.pptx"), ("All", "*.*")])
        if path:
            self.template_path.set(path)

    def _browse_pdf(self):
        path = filedialog.askopenfilename(
            title="选择 PDF 试卷",
            filetypes=[("PDF", "*.pdf"), ("All", "*.*")],
        )
        if path:
            self.pdf_path.set(path)
            if not self.pdf_word_out.get().strip():
                self.pdf_word_out.set(os.path.splitext(path)[0] + "_真题.docx")
            if not self.pdf_ppt_out.get().strip():
                self.pdf_ppt_out.set(os.path.splitext(path)[0] + "_授课.pptx")
            if not self.pdf_manifest_out.get().strip():
                self.pdf_manifest_out.set(os.path.splitext(path)[0] + "_工程.json")

    def _browse_pdf_word(self):
        path = filedialog.asksaveasfilename(
            title="保存真题 Word",
            defaultextension=".docx",
            filetypes=[("Word", "*.docx"), ("All", "*.*")],
        )
        if path:
            self.pdf_word_out.set(path)

    def _browse_pdf_ppt(self):
        path = filedialog.asksaveasfilename(
            title="保存授课 PPT",
            defaultextension=".pptx",
            filetypes=[("PowerPoint", "*.pptx"), ("All", "*.*")],
        )
        if path:
            self.pdf_ppt_out.set(path)

    def _browse_pdf_manifest(self):
        path = filedialog.asksaveasfilename(
            title="保存工程清单 JSON",
            defaultextension=".json",
            filetypes=[("JSON", "*.json"), ("All", "*.*")],
        )
        if path:
            self.pdf_manifest_out.set(path)

    def _load_pdf_manifest_project(self):
        if not self._confirm_discard_pdf_project_edits("载入新的工程 JSON"):
            return
        path = filedialog.askopenfilename(
            title="选择工程 JSON",
            filetypes=[("JSON", "*.json"), ("All", "*.*")],
        )
        if not path:
            return
        try:
            project = load_project_manifest_project(path)
        except Exception as exc:
            messagebox.showerror("载入失败", str(exc), parent=self.root)
            return

        pdf_path = project.source.pdf_path or ""
        asset_dir = project.source.asset_dir or ""
        self.pdf_project = project
        self._pdf_project_dirty = False
        self._pdf_project_context = {
            "pdf_path": pdf_path,
            "subject_spec": "all",
            "range_spec": "",
            "asset_dir": asset_dir,
            "source_kind": "manifest",
            "manifest_path": path,
            "document_subject_hint": "auto",
        }

        self.pdf_path.set(pdf_path)
        self.pdf_document_subject.set(_DOCUMENT_SUBJECT_LABELS["auto"])
        self.pdf_manifest_out.set(path)
        default_base = os.path.splitext(pdf_path or path)[0]
        self.pdf_word_out.set(default_base + "_真题.docx")
        self.pdf_ppt_out.set(default_base + "_授课.pptx")

        self._apply_project_subject_selection(project.selected_subjects)
        self.pdf_question_range.set(self._format_question_ranges_for_gui(project.selected_ranges))
        self._pdf_project_context["subject_spec"] = self._current_pdf_subject_spec()
        self._pdf_project_context["range_spec"] = self.pdf_question_range.get().strip()

        self._reset_pdf_material_preview_session()
        self._populate_pdf_preview(project)
        self._show_pdf_wizard_step(2)
        self._set_status(f"已载入工程：{os.path.basename(path)}")

    def _preview_pdf_project(self):
        if not self._confirm_discard_pdf_project_edits("重新生成预览"):
            return
        self._run_pdf_project(export_word=False, export_ppt=False)

    def _export_pdf_word(self):
        self._run_pdf_project(export_word=True, export_ppt=False)

    def _export_pdf_ppt(self):
        self._run_pdf_project(export_word=False, export_ppt=True)

    def _export_pdf_bundle(self):
        self._run_pdf_project(export_word=True, export_ppt=True)

    def _run_pdf_project(self, *, export_word: bool, export_ppt: bool):
        from workflows.project_flow import build_pdf_project, export_project_outputs

        pdf_file = self.pdf_path.get().strip()
        subject_spec = self._current_pdf_subject_spec()
        source_kind = self._pdf_project_context.get("source_kind", "")
        is_word_project = source_kind == "word" and self.pdf_project is not None
        document_subject_hint = self._document_subject_key(self.pdf_document_subject.get())
        if not subject_spec and not is_word_project:
            messagebox.showwarning("提示", "请至少选择一个科目")
            return
        range_spec = self.pdf_question_range.get().strip()
        use_cached_project = (
            self.pdf_project is not None
            and self._pdf_project_context.get("pdf_path") == pdf_file
            and self._pdf_project_context.get("subject_spec", "all") == subject_spec
            and self._pdf_project_context.get("range_spec", "") == range_spec
            and self._pdf_project_context.get("document_subject_hint", "auto") == document_subject_hint
        )
        if is_word_project:
            use_cached_project = True
        if not use_cached_project:
            if not pdf_file:
                messagebox.showwarning("提示", "请先选择 PDF 文件")
                return
            if not os.path.exists(pdf_file):
                messagebox.showerror("错误", f"文件不存在：{pdf_file}")
                return

        docx_output = self.pdf_word_out.get().strip() if export_word else None
        ppt_output = self.pdf_ppt_out.get().strip() if export_ppt else None
        manifest_output = self.pdf_manifest_out.get().strip() or None
        default_base = self._default_pdf_base_path()
        if export_word and not docx_output:
            docx_output = default_base + "_真题.docx"
            self.pdf_word_out.set(docx_output)
        if export_ppt and not ppt_output:
            ppt_output = default_base + "_授课.pptx"
            self.pdf_ppt_out.set(ppt_output)
        if not (export_word or export_ppt):
            manifest_output = None
        elif not manifest_output:
            manifest_output = default_base + "_工程.json"
            self.pdf_manifest_out.set(manifest_output)

        template = self.template_path.get().strip() or None
        if template and not os.path.exists(template):
            messagebox.showerror("错误", "PPT 模板文件不存在")
            return

        self._set_status("正在整理题目工程…")
        self.progress["value"] = 0
        self.progress["maximum"] = 100
        if not use_cached_project and not self._confirm_discard_pdf_project_edits("按当前设置重新整理题目工程"):
            return

        def work():
            try:
                if use_cached_project:
                    project = self.pdf_project
                    asset_dir = self._pdf_project_context.get("asset_dir", "")
                else:
                    project, asset_dir = build_pdf_project(
                        pdf_file,
                        mode=subject_spec,
                        question_range_spec=range_spec,
                        document_subject_hint=None if document_subject_hint == "auto" else document_subject_hint,
                    )
                ppt_config = self._make_ppt_config() if export_ppt else None
                outputs = export_project_outputs(
                    project,
                    asset_dir=asset_dir,
                    docx_output=docx_output,
                    ppt_output=ppt_output,
                    manifest_output=manifest_output,
                    template_path=template,
                    ppt_config=ppt_config,
                )
                self.root.after(0, lambda: self._on_pdf_project_done(project, outputs))
            except Exception as exc:
                self.root.after(0, lambda e=exc: self._on_pdf_project_error(str(e)))

        threading.Thread(target=work, daemon=True).start()

    def _on_pdf_project_done(self, project, outputs):
        self.pdf_project = project
        self._pdf_project_dirty = False
        source_kind = self._pdf_project_context.get("source_kind", "")
        manifest_path = self._pdf_project_context.get("manifest_path", "")
        self._pdf_project_context = {
            "pdf_path": self.pdf_path.get().strip() if source_kind != "word" else "",
            "docx_path": self.word_path.get().strip() if source_kind == "word" else self._pdf_project_context.get("docx_path", ""),
            "subject_spec": self._current_pdf_subject_spec() if source_kind != "word" else "all",
            "range_spec": self.pdf_question_range.get().strip() if source_kind != "word" else "",
            "asset_dir": outputs.asset_dir,
            "source_kind": source_kind or "pdf",
            "manifest_path": self.pdf_manifest_out.get().strip() or manifest_path,
            "document_subject_hint": (
                self._document_subject_key(self.word_document_subject.get())
                if source_kind == "word"
                else self._document_subject_key(self.pdf_document_subject.get())
            ),
        }
        self._reset_pdf_material_preview_session()
        self._populate_pdf_preview(project)
        preview_only = not (outputs.docx_path or outputs.pptx_path or outputs.manifest_path)
        if preview_only:
            self._show_pdf_wizard_step(self._pdf_wizard_pending_step or 2)
        else:
            self._show_pdf_wizard_step(3)
        self._pdf_wizard_pending_step = None
        self.progress["value"] = self.progress["maximum"]
        self._set_status(f"PDF 工作流已完成，共 {project.question_count} 道题")
        result_lines = [f"共整理题目：{project.question_count}", f"素材目录：{outputs.asset_dir}"]
        if outputs.docx_path:
            result_lines.append(f"题本 Word：{outputs.docx_path}")
        if outputs.pptx_path:
            result_lines.append(f"授课 PPT：{outputs.pptx_path}")
        if outputs.manifest_path:
            result_lines.append(f"工程 JSON：{outputs.manifest_path}")
        if outputs.docx_path or outputs.pptx_path or outputs.manifest_path:
            messagebox.showinfo(
                "完成",
                "\n".join(result_lines),
            )

    def _on_pdf_project_error(self, msg: str):
        self._pdf_wizard_pending_step = None
        self.progress["value"] = 0
        self._set_status("PDF 工作流失败")
        messagebox.showerror("处理失败", msg)

    # -----------------------------------------------------------------------------
    # Core operations
    # -----------------------------------------------------------------------------

    def _refresh_question_tree(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        for question in self.questions:
            stem_short = question.display_stem[:50] + ("..." if len(question.display_stem) > 50 else "")
            image_count = len(question.image_paths)
            self.tree.insert(
                "",
                END,
                values=(
                    question.number,
                    stem_short,
                    len(question.options),
                    f"{image_count} 张" if image_count else "-",
                ),
            )

    def _clear_pdf_preview(self):
        self.pdf_project = None
        self._pdf_project_dirty = False
        self._pdf_wizard_pending_step = None
        self._pdf_project_context = {}
        self._pdf_preview_payloads.clear()
        self._reset_pdf_material_preview_session()
        self._clear_pdf_question_editor()
        if getattr(self, "pdf_tree", None):
            for item in self.pdf_tree.get_children():
                self.pdf_tree.delete(item)
        self._set_pdf_detail("")
        self._refresh_pdf_wizard_ui()

    def _set_pdf_detail(self, text: str):
        if not getattr(self, "pdf_detail", None):
            return
        self.pdf_detail.configure(state="normal")
        self.pdf_detail.delete("1.0", tk.END)
        self.pdf_detail.insert("1.0", text or "暂无内容")
        self.pdf_detail.configure(state="disabled")

    def _selected_pdf_item_id(self) -> str:
        if not getattr(self, "pdf_tree", None):
            return ""
        selected = self.pdf_tree.selection()
        return selected[0] if selected else ""

    def _question_tree_label(self, question) -> str:
        compact = " ".join((question.stem or "").split())
        stem_short = compact[:38]
        if len(compact) > 38:
            stem_short += "..."
        return stem_short or "未命名题目"

    def _option_layout_label(self, layout: str | None) -> str:
        normalized = (layout or "").strip().lower()
        return _PDF_QUESTION_LAYOUT_LABELS.get(normalized, _PDF_QUESTION_LAYOUT_LABELS[""])

    def _effective_question_option_layout(self, question) -> str:
        normalized = (getattr(question, "option_layout", None) or "").strip().lower()
        if normalized in {"grid", "list", "one_row"}:
            return normalized
        global_layout = (self.option_layout.get() or "grid").strip().lower()
        return global_layout if global_layout in {"grid", "list", "one_row"} else "grid"

    def _find_pdf_question_option(self, question, letter: str):
        normalized = (letter or "").strip().upper()
        for option in getattr(question, "options", []):
            if (option.letter or "").strip().upper() == normalized:
                return option
        return None

    def _mark_pdf_project_dirty(self):
        if self.pdf_project is None:
            return
        self._pdf_project_dirty = True

    def _confirm_discard_pdf_project_edits(self, action_text: str) -> bool:
        if not self._pdf_project_dirty or self.pdf_project is None:
            return True
        return messagebox.askyesno(
            "确认继续",
            f"当前预览里有人工修改；{action_text}会丢失这些改动。\n\n确定继续吗？",
            parent=self.root,
        )

    def _rebuild_pdf_option_editors(self, question=None):
        host = getattr(self, "_pdf_option_editor_host", None)
        if host is None:
            return
        for child in host.winfo_children():
            child.destroy()

        self._pdf_option_editors.clear()
        self._pdf_option_image_labels.clear()
        self._pdf_option_view_buttons.clear()
        self._pdf_option_recrop_buttons.clear()
        self._pdf_option_clear_buttons.clear()
        self._pdf_option_replace_buttons.clear()
        self._pdf_option_move_up_buttons.clear()
        self._pdf_option_move_down_buttons.clear()
        self._pdf_option_insert_buttons.clear()
        self._pdf_option_remove_buttons.clear()

        options = list(getattr(question, "options", []) or [])
        if not options:
            ttk.Label(
                host,
                text="当前题目暂无可编辑选项。",
                bootstyle="secondary",
            ).pack(anchor=W)
            return

        for option in options:
            card = ttk.Frame(host)
            card.pack(fill=X, pady=(0, 8))

            row = ttk.Frame(card)
            row.pack(fill=X)
            ttk.Label(row, text=f"{option.letter}.", width=3).pack(side=LEFT, anchor=N, pady=(4, 0))

            editor_host = ttk.Frame(row)
            editor_host.pack(side=LEFT, fill=BOTH, expand=YES, padx=(0, 8))
            editor = tk.Text(editor_host, wrap="word", height=2)
            editor.pack(fill=X, expand=YES)
            editor.insert("1.0", option.text or "")
            editor.bind(
                "<KeyRelease>",
                lambda _event, letter=option.letter: self._on_pdf_option_text_change(letter),
            )
            editor.bind(
                "<FocusOut>",
                lambda _event, letter=option.letter: self._on_pdf_option_text_change(letter),
            )
            self._pdf_option_editors[option.letter] = editor

            control_col = ttk.Frame(row)
            control_col.pack(side=RIGHT, anchor=N)
            status = ttk.Label(control_col, text="", width=26, bootstyle="secondary")
            status.pack(anchor=E)

            button_row1 = ttk.Frame(control_col)
            button_row1.pack(anchor=E, pady=(4, 0))
            view_btn = ttk.Button(
                button_row1,
                text="查看图",
                command=lambda letter=option.letter: self._open_pdf_option_image(letter),
                bootstyle="secondary-outline",
                width=7,
            )
            view_btn.pack(side=LEFT, padx=(0, 4))
            replace_btn = ttk.Button(
                button_row1,
                text="换图",
                command=lambda letter=option.letter: self._replace_pdf_option_image(letter),
                bootstyle="info-outline",
                width=7,
            )
            replace_btn.pack(side=LEFT, padx=(0, 4))
            recrop_btn = ttk.Button(
                button_row1,
                text="PDF重裁",
                command=lambda letter=option.letter: self._recrop_pdf_option_image(letter),
                bootstyle="info-outline",
                width=8,
            )
            recrop_btn.pack(side=LEFT, padx=(0, 4))
            clear_btn = ttk.Button(
                button_row1,
                text="清图",
                command=lambda letter=option.letter: self._clear_pdf_option_image(letter),
                bootstyle="warning-outline",
                width=7,
            )
            clear_btn.pack(side=LEFT)

            button_row2 = ttk.Frame(control_col)
            button_row2.pack(anchor=E, pady=(4, 0))
            move_up_btn = ttk.Button(
                button_row2,
                text="上移",
                command=lambda letter=option.letter: self._move_pdf_option(letter, -1),
                bootstyle="secondary-outline",
                width=7,
            )
            move_up_btn.pack(side=LEFT, padx=(0, 4))
            move_down_btn = ttk.Button(
                button_row2,
                text="下移",
                command=lambda letter=option.letter: self._move_pdf_option(letter, 1),
                bootstyle="secondary-outline",
                width=7,
            )
            move_down_btn.pack(side=LEFT, padx=(0, 4))
            insert_btn = ttk.Button(
                button_row2,
                text="下方加项",
                command=lambda letter=option.letter: self._insert_pdf_option_after(letter),
                bootstyle="success-outline",
                width=9,
            )
            insert_btn.pack(side=LEFT, padx=(0, 4))
            remove_btn = ttk.Button(
                button_row2,
                text="删项",
                command=lambda letter=option.letter: self._remove_pdf_option(letter),
                bootstyle="danger-outline",
                width=7,
            )
            remove_btn.pack(side=LEFT)

            self._pdf_option_image_labels[option.letter] = status
            self._pdf_option_view_buttons[option.letter] = view_btn
            self._pdf_option_recrop_buttons[option.letter] = recrop_btn
            self._pdf_option_replace_buttons[option.letter] = replace_btn
            self._pdf_option_clear_buttons[option.letter] = clear_btn
            self._pdf_option_move_up_buttons[option.letter] = move_up_btn
            self._pdf_option_move_down_buttons[option.letter] = move_down_btn
            self._pdf_option_insert_buttons[option.letter] = insert_btn
            self._pdf_option_remove_buttons[option.letter] = remove_btn
            self._refresh_pdf_option_image_status(option.letter)
        self._refresh_pdf_option_action_buttons()

    def _refresh_pdf_question_editor_message(self, payload: dict | None = None):
        payload = payload or self._selected_pdf_payload()
        if payload.get("kind") != "question":
            self._pdf_question_editor_message.set(
                "选择一道题后，可在这里实时修改题干、选项内容，并为该题单独切换选项布局。"
            )
            return

        question = payload.get("question")
        if question is None:
            self._pdf_question_editor_message.set(
                "选择一道题后，可在这里实时修改题干、选项内容，并为该题单独切换选项布局。"
            )
            return

        section_kind = payload.get("section_kind") or "unknown"
        section_name = SUBJECT_DISPLAY_NAMES.get(section_kind, section_kind)
        material = payload.get("material")
        parts = [section_name, f"原题号 {question.source_number or '-'}"]
        if material is not None:
            parts.append(material.header or material.material_id)
        layout_source = "单题覆盖" if question.option_layout else "跟随全局"
        layout_label = self._option_layout_label(question.option_layout or self._effective_question_option_layout(question))
        message = " · ".join(parts)
        message += f"\n当前布局：{layout_source}（{layout_label}）"
        if question.stem_assets:
            message += f" · 题干图片 {len(question.stem_assets)} 张"
        image_option_count = sum(1 for option in question.options if option.image_path)
        if image_option_count:
            message += f" · 图片选项 {image_option_count} 个"
        self._pdf_question_editor_message.set(message)

    def _set_pdf_question_editor_state(self, enabled: bool):
        state = NORMAL if enabled else DISABLED
        if getattr(self, "_pdf_question_stem_editor", None):
            self._pdf_question_stem_editor.configure(state=state)
        for button in getattr(self, "_pdf_question_layout_buttons", []):
            button.configure(state=state)
        for editor in getattr(self, "_pdf_option_editors", {}).values():
            editor.configure(state=state)
        for button_map in (
            getattr(self, "_pdf_option_view_buttons", {}),
            getattr(self, "_pdf_option_recrop_buttons", {}),
            getattr(self, "_pdf_option_replace_buttons", {}),
            getattr(self, "_pdf_option_clear_buttons", {}),
            getattr(self, "_pdf_option_move_up_buttons", {}),
            getattr(self, "_pdf_option_move_down_buttons", {}),
            getattr(self, "_pdf_option_insert_buttons", {}),
            getattr(self, "_pdf_option_remove_buttons", {}),
        ):
            for button in button_map.values():
                button.configure(state=state)
        if enabled:
            for letter in getattr(self, "_pdf_option_image_labels", {}).keys():
                self._refresh_pdf_option_image_status(letter)
            self._refresh_pdf_option_action_buttons()

    def _set_pdf_stem_preview_message(self, message: str, status: str):
        self._pdf_stem_preview_photo = None
        self._pdf_stem_preview_paths = []
        self._pdf_stem_preview_index = 0
        if getattr(self, "_pdf_stem_preview_status", None):
            self._pdf_stem_preview_status.configure(text=status)
        if getattr(self, "_pdf_stem_preview_prev", None):
            self._pdf_stem_preview_prev.configure(state=DISABLED)
        if getattr(self, "_pdf_stem_preview_next", None):
            self._pdf_stem_preview_next.configure(state=DISABLED)
        if getattr(self, "_pdf_stem_preview_open", None):
            self._pdf_stem_preview_open.configure(state=DISABLED)
        if getattr(self, "_pdf_stem_preview_box", None):
            self._pdf_stem_preview_box.configure(image="", text=message)

    def _refresh_pdf_stem_preview_for_question(self, question):
        if question is None:
            self._set_pdf_stem_preview_message("当前题目没有题干图片。", "暂无题干图片")
            return
        paths = [
            asset.path
            for asset in getattr(question, "stem_assets", []) or []
            if asset.path and os.path.exists(asset.path)
        ]
        if not paths:
            self._set_pdf_stem_preview_message("当前题目没有题干图片。", "暂无题干图片")
            return
        self._pdf_stem_preview_paths = paths
        self._pdf_stem_preview_index = min(self._pdf_stem_preview_index, len(paths) - 1)
        self._render_pdf_stem_preview()

    def _step_pdf_stem_preview(self, delta: int):
        if not self._pdf_stem_preview_paths:
            return
        self._pdf_stem_preview_index = (
            self._pdf_stem_preview_index + delta
        ) % len(self._pdf_stem_preview_paths)
        self._render_pdf_stem_preview()

    def _open_pdf_stem_preview_image(self):
        if not self._pdf_stem_preview_paths:
            return
        index = min(self._pdf_stem_preview_index, len(self._pdf_stem_preview_paths) - 1)
        image_path = self._pdf_stem_preview_paths[index]
        if not image_path or not os.path.exists(image_path):
            messagebox.showinfo("提示", "当前题干图片不存在。", parent=self.root)
            return
        os.startfile(os.path.abspath(image_path))

    def _render_pdf_stem_preview(self):
        if not getattr(self, "_pdf_stem_preview_box", None):
            return
        if not self._pdf_stem_preview_paths:
            return
        index = min(self._pdf_stem_preview_index, len(self._pdf_stem_preview_paths) - 1)
        image_path = self._pdf_stem_preview_paths[index]
        if not image_path or not os.path.exists(image_path):
            self._set_pdf_stem_preview_message("当前题干图片不存在。", "题干图片不可用")
            return

        target_width = self._pdf_stem_preview_box.winfo_width()
        max_width = max(240, target_width - 20) if target_width > 40 else 420
        with Image.open(image_path) as source_image:
            image = source_image.copy()
        image.thumbnail((max_width, 180), Image.Resampling.LANCZOS)
        self._pdf_stem_preview_photo = ImageTk.PhotoImage(image)
        self._pdf_stem_preview_box.configure(image=self._pdf_stem_preview_photo, text="")
        if getattr(self, "_pdf_stem_preview_status", None):
            self._pdf_stem_preview_status.configure(
                text=f"题干图片 {index + 1}/{len(self._pdf_stem_preview_paths)} · {os.path.basename(image_path)}"
            )
        state = NORMAL if len(self._pdf_stem_preview_paths) > 1 else DISABLED
        if getattr(self, "_pdf_stem_preview_prev", None):
            self._pdf_stem_preview_prev.configure(state=state)
        if getattr(self, "_pdf_stem_preview_next", None):
            self._pdf_stem_preview_next.configure(state=state)
        if getattr(self, "_pdf_stem_preview_open", None):
            self._pdf_stem_preview_open.configure(state=NORMAL)

    def _clear_pdf_question_editor(self, message: str | None = None):
        self._pdf_question_editor_target = None
        self._pdf_editor_updating = True
        if getattr(self, "_pdf_question_stem_editor", None):
            self._pdf_question_stem_editor.configure(state=NORMAL)
            self._pdf_question_stem_editor.delete("1.0", tk.END)
            self._pdf_question_stem_editor.configure(state=DISABLED)
        self._set_pdf_stem_preview_message("当前题目没有题干图片。", "暂无题干图片")
        self._rebuild_pdf_option_editors(None)
        self._pdf_question_layout_var.set("")
        self._set_pdf_question_editor_state(False)
        self._pdf_editor_updating = False
        self._pdf_question_editor_message.set(
            message or "选择一道题后，可在这里实时修改题干、选项内容，并为该题单独切换选项布局。"
        )
        self._render_pdf_question_editor_preview()

    def _populate_pdf_question_editor(self, payload: dict):
        if payload.get("kind") != "question":
            self._clear_pdf_question_editor("当前节点不是题目；请选择左侧一道题后再修改题干或布局。")
            return
        question = payload.get("question")
        if question is None:
            self._clear_pdf_question_editor()
            return

        self._pdf_question_editor_target = question
        self._pdf_editor_updating = True
        self._set_pdf_question_editor_state(True)
        self._pdf_question_stem_editor.configure(state=NORMAL)
        self._pdf_question_stem_editor.delete("1.0", tk.END)
        self._pdf_question_stem_editor.insert("1.0", question.stem or "")
        self._refresh_pdf_stem_preview_for_question(question)
        self._rebuild_pdf_option_editors(question)
        self._pdf_question_layout_var.set((question.option_layout or "").strip().lower())
        self._pdf_editor_updating = False
        self._set_pdf_question_editor_state(True)
        self._refresh_pdf_question_editor_message(payload)
        self._render_pdf_question_editor_preview()

    def _sync_selected_question_preview_payload(self):
        payload = self._selected_pdf_payload()
        if payload.get("kind") != "question":
            return
        question = payload.get("question")
        section = payload.get("section")
        material = payload.get("material")
        if question is None or section is None:
            return

        payload["text"] = self._question_preview_text(section, material, question)
        item_id = self._selected_pdf_item_id()
        if item_id:
            self.pdf_tree.item(
                item_id,
                text=self._question_tree_label(question),
                values=("question", question.source_number or "-", len(question.options)),
            )
        self._set_pdf_detail(payload["text"])
        self._refresh_pdf_question_editor_message(payload)
        self._render_pdf_question_editor_preview()

    def _on_pdf_question_stem_change(self, _event=None):
        if self._pdf_editor_updating or self._pdf_question_editor_target is None:
            return
        content = self._pdf_question_stem_editor.get("1.0", tk.END).rstrip("\n")
        update_question_stem(self._pdf_question_editor_target, content)
        self._mark_pdf_project_dirty()
        self._sync_selected_question_preview_payload()

    def _on_pdf_question_layout_change(self):
        if self._pdf_editor_updating or self._pdf_question_editor_target is None:
            return
        set_question_option_layout(self._pdf_question_editor_target, self._pdf_question_layout_var.get())
        self._mark_pdf_project_dirty()
        self._sync_selected_question_preview_payload()

    def _refresh_pdf_option_image_status(self, letter: str):
        question = self._pdf_question_editor_target
        status = self._pdf_option_image_labels.get(letter)
        if question is None or status is None:
            return
        option = self._find_pdf_question_option(question, letter)
        image_path = getattr(option, "image_path", None) if option is not None else None
        page_region = getattr(option, "page_region", None) if option is not None else None
        exists = bool(image_path and os.path.exists(image_path))
        if exists:
            label_text = f"图片：{os.path.basename(image_path)}"
        elif image_path:
            label_text = f"图片路径失效：{os.path.basename(image_path)}"
        else:
            label_text = "图片：无"
        status.configure(text=label_text)

        view_btn = self._pdf_option_view_buttons.get(letter)
        recrop_btn = self._pdf_option_recrop_buttons.get(letter)
        clear_btn = self._pdf_option_clear_buttons.get(letter)
        if view_btn is not None:
            view_btn.configure(state=NORMAL if exists else DISABLED)
        if recrop_btn is not None:
            recrop_btn.configure(state=NORMAL if page_region is not None else DISABLED)
        if clear_btn is not None:
            clear_btn.configure(state=NORMAL if image_path else DISABLED)

    def _refresh_pdf_option_action_buttons(self):
        question = self._pdf_question_editor_target
        options = list(getattr(question, "options", []) or [])
        option_count = len(options)
        for index, option in enumerate(options):
            letter = option.letter
            move_up_btn = self._pdf_option_move_up_buttons.get(letter)
            move_down_btn = self._pdf_option_move_down_buttons.get(letter)
            insert_btn = self._pdf_option_insert_buttons.get(letter)
            remove_btn = self._pdf_option_remove_buttons.get(letter)
            if move_up_btn is not None:
                move_up_btn.configure(state=NORMAL if index > 0 else DISABLED)
            if move_down_btn is not None:
                move_down_btn.configure(state=NORMAL if index < option_count - 1 else DISABLED)
            if insert_btn is not None:
                insert_btn.configure(state=NORMAL if option_count < 26 else DISABLED)
            if remove_btn is not None:
                remove_btn.configure(state=NORMAL if option_count > 1 else DISABLED)

    def _on_pdf_option_text_change(self, letter: str):
        if self._pdf_editor_updating or self._pdf_question_editor_target is None:
            return
        editor = self._pdf_option_editors.get(letter)
        if editor is None:
            return
        content = editor.get("1.0", tk.END).rstrip("\n")
        if update_option_text(self._pdf_question_editor_target, letter, content):
            self._mark_pdf_project_dirty()
            self._sync_selected_question_preview_payload()

    def _refresh_pdf_option_editor_after_structure_change(self):
        question = self._pdf_question_editor_target
        if question is None:
            return
        self._pdf_editor_updating = True
        self._rebuild_pdf_option_editors(question)
        self._pdf_editor_updating = False
        self._set_pdf_question_editor_state(True)
        self._mark_pdf_project_dirty()
        self._sync_selected_question_preview_payload()

    def _move_pdf_option(self, letter: str, direction: int):
        question = self._pdf_question_editor_target
        if question is None:
            return
        if move_option(question, letter, direction):
            self._refresh_pdf_option_editor_after_structure_change()

    def _insert_pdf_option_after(self, letter: str):
        question = self._pdf_question_editor_target
        if question is None:
            return
        if insert_option_after(question, letter):
            self._refresh_pdf_option_editor_after_structure_change()

    def _remove_pdf_option(self, letter: str):
        question = self._pdf_question_editor_target
        if question is None:
            return
        if not messagebox.askyesno("确认删项", f"确定删除 {letter} 选项吗？", parent=self.root):
            return
        if remove_option(question, letter):
            self._refresh_pdf_option_editor_after_structure_change()

    def _stage_pdf_option_image(self, selected_path: str) -> str:
        source = os.path.abspath(selected_path)
        asset_dir = (self._pdf_project_context.get("asset_dir") or "").strip()
        if not asset_dir:
            return source
        os.makedirs(asset_dir, exist_ok=True)
        name = os.path.basename(source)
        stem, ext = os.path.splitext(name)
        candidate = os.path.join(asset_dir, name)
        suffix = 1
        while os.path.exists(candidate) and os.path.abspath(candidate) != source:
            candidate = os.path.join(asset_dir, f"{stem}_manual_{suffix}{ext}")
            suffix += 1
        if os.path.abspath(candidate) != source:
            shutil.copy2(source, candidate)
        return candidate

    def _stage_pdf_option_crop(self, selected_path: str, letter: str) -> str:
        source = os.path.abspath(selected_path)
        asset_dir = (self._pdf_project_context.get("asset_dir") or "").strip()
        if not asset_dir:
            return source
        os.makedirs(asset_dir, exist_ok=True)
        question = self._pdf_question_editor_target
        source_number = getattr(question, "source_number", "") if question is not None else ""
        question_label = source_number or "question"
        ext = os.path.splitext(source)[1] or ".png"
        base_name = f"q{question_label}_{letter}_crop"
        candidate = os.path.join(asset_dir, f"{base_name}{ext}")
        suffix = 1
        while os.path.exists(candidate) and os.path.abspath(candidate) != source:
            candidate = os.path.join(asset_dir, f"{base_name}_{suffix}{ext}")
            suffix += 1
        if os.path.abspath(candidate) != source:
            shutil.copy2(source, candidate)
        return candidate

    def _open_pdf_option_image(self, letter: str):
        question = self._pdf_question_editor_target
        if question is None:
            return
        option = self._find_pdf_question_option(question, letter)
        image_path = getattr(option, "image_path", None) if option is not None else None
        if not image_path or not os.path.exists(image_path):
            messagebox.showinfo("提示", "当前选项没有可打开的图片。")
            return
        os.startfile(os.path.abspath(image_path))

    def _replace_pdf_option_image(self, letter: str):
        question = self._pdf_question_editor_target
        if question is None:
            return
        selected_path = filedialog.askopenfilename(
            title=f"为 {letter} 选项选择图片",
            filetypes=[
                ("Image", "*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.webp"),
                ("All", "*.*"),
            ],
        )
        if not selected_path:
            return
        staged_path = self._stage_pdf_option_image(selected_path)
        if replace_option_image(question, letter, staged_path):
            self._mark_pdf_project_dirty()
            self._refresh_pdf_option_image_status(letter)
            self._sync_selected_question_preview_payload()

    def _recrop_pdf_option_image(self, letter: str):
        question = self._pdf_question_editor_target
        if question is None:
            return
        option = self._find_pdf_question_option(question, letter)
        region = getattr(option, "page_region", None) if option is not None else None
        pdf_path = self._pdf_project_context.get("pdf_path") or ""
        if not pdf_path and getattr(self, "pdf_project", None) is not None:
            pdf_path = getattr(self.pdf_project.source, "pdf_path", "") or ""
        if not pdf_path or not os.path.exists(pdf_path):
            messagebox.showinfo("提示", "当前工程没有可用的原始 PDF。", parent=self.root)
            return
        if region is None:
            messagebox.showinfo("提示", "当前选项没有保存原 PDF 区域，暂时无法重裁。", parent=self.root)
            return
        temp_dir = tempfile.mkdtemp(prefix="pptconvert_option_crop_")
        try:
            crop_paths = crop_page_regions(
                pdf_path,
                [region],
                temp_dir,
                prefix=f"option_{letter.lower()}",
                margin=10.0,
                dpi=180,
            )
            if not crop_paths:
                messagebox.showinfo("提示", "没有裁出图片，请检查该选项的区域信息。", parent=self.root)
                return
            staged_path = self._stage_pdf_option_crop(crop_paths[0], letter)
            if replace_option_image(question, letter, staged_path):
                self._mark_pdf_project_dirty()
                self._refresh_pdf_option_image_status(letter)
                self._sync_selected_question_preview_payload()
        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)

    def _clear_pdf_option_image(self, letter: str):
        question = self._pdf_question_editor_target
        if question is None:
            return
        if clear_option_image(question, letter):
            self._mark_pdf_project_dirty()
            self._refresh_pdf_option_image_status(letter)
            self._sync_selected_question_preview_payload()

    def _draw_pdf_preview_thumbnail(self, canvas, image_path: str, x0: float, y0: float, width: float, height: float) -> bool:
        if not image_path or not os.path.exists(image_path):
            return False
        if width < 24 or height < 24:
            return False
        try:
            with Image.open(image_path) as source_image:
                image = source_image.copy()
        except Exception:
            return False

        image.thumbnail(
            (max(16, int(width)), max(16, int(height))),
            Image.Resampling.LANCZOS,
        )
        photo = ImageTk.PhotoImage(image)
        self._pdf_question_preview_photos.append(photo)
        canvas.create_image(
            x0 + width / 2,
            y0 + height / 2,
            image=photo,
            anchor=CENTER,
        )
        return True

    def _draw_pdf_question_preview_option(self, canvas, option, x0, y0, width, height, fill, outline):
        canvas.create_rectangle(x0, y0, x0 + width, y0 + height, fill=fill, outline=outline, width=2)
        option_label = f"{option.letter}."
        body_text = (option.text or "").strip()
        text = f"{option_label} {body_text}".strip()
        image_exists = bool(option.image_path and os.path.exists(option.image_path))
        option_align = (self.option_align.get() or "left").lower()
        if option_align == "center":
            anchor = tk.N
            justify = CENTER
            x = x0 + width / 2
        elif option_align == "right":
            anchor = tk.NE
            justify = RIGHT
            x = x0 + width - 8
        else:
            anchor = tk.NW
            justify = LEFT
            x = x0 + 8

        if image_exists:
            content_width = max(24, width - 16)
            text_height = max(28, min(height * 0.38, height - 44))
            preview_text = text or option_label
            wrapped_text, font_spec = self._fit_preview_text(
                preview_text,
                width=content_width,
                height=max(24, text_height - 4),
                target_points=float(self.font_size_option.get()),
                scale_px_per_in=self._pdf_preview_scale_px_per_in,
                bold=self.option_text_bold.get(),
            )
            canvas.create_text(
                x,
                y0 + 8,
                anchor=anchor,
                width=content_width,
                text=wrapped_text,
                justify=justify,
                fill=self.color_option.get().strip() or "#2D2D2D",
                font=font_spec,
            )

            image_top = y0 + text_height
            image_height = max(24, height - (image_top - y0) - 8)
            image_left = x0 + 8
            drawn = self._draw_pdf_preview_thumbnail(
                canvas,
                option.image_path,
                image_left,
                image_top,
                width - 16,
                image_height,
            )
            if not drawn:
                canvas.create_rectangle(
                    image_left,
                    image_top,
                    image_left + width - 16,
                    image_top + image_height,
                    outline="#d9a84a",
                    dash=(4, 2),
                )
                canvas.create_text(
                    x0 + width / 2,
                    image_top + image_height / 2,
                    text="图片不可预览",
                    fill="#996600",
                    font=("", 9, "bold"),
                )
            return

        wrapped_text, font_spec = self._fit_preview_text(
            text or option_label,
            width=max(40, width - 16),
            height=max(28, height - 12),
            target_points=float(self.font_size_option.get()),
            scale_px_per_in=self._pdf_preview_scale_px_per_in,
            bold=self.option_text_bold.get(),
        )
        canvas.create_text(
            x,
            y0 + 8,
            anchor=anchor,
            width=max(24, width - 16),
            text=wrapped_text,
            justify=justify,
            fill=self.color_option.get().strip() or "#2D2D2D",
            font=font_spec,
        )

    def _preview_font(self, pixel_size: int, *, bold: bool):
        family = self.font_name.get().strip() or "微软雅黑"
        weight = "bold" if bold else "normal"
        try:
            return tkfont.Font(family=family, size=-max(6, int(pixel_size)), weight=weight)
        except tk.TclError:
            return tkfont.Font(size=-max(6, int(pixel_size)), weight=weight)

    def _wrap_preview_text(self, text: str, font: tkfont.Font, width: float) -> list[str]:
        content = (text or "").replace("\r\n", "\n").replace("\r", "\n")
        max_width = max(24, int(width))
        wrapped: list[str] = []
        for paragraph in content.split("\n") or [""]:
            if not paragraph:
                wrapped.append("")
                continue
            current = ""
            for ch in paragraph:
                candidate = current + ch
                if current and font.measure(candidate) > max_width:
                    wrapped.append(current.rstrip())
                    current = ch
                else:
                    current = candidate
            wrapped.append(current.rstrip())
        return wrapped or [""]

    def _fit_preview_text(
        self,
        text: str,
        *,
        width: float,
        height: float,
        target_points: float,
        scale_px_per_in: float,
        bold: bool,
    ) -> tuple[str, tuple[str, int, str]]:
        target_px = max(8, int(round(target_points / 72.0 * scale_px_per_in)))
        min_px = max(7, min(target_px, 9))
        chosen_lines = [text or ""]
        chosen_px = min_px
        for pixel_size in range(target_px, min_px - 1, -1):
            font = self._preview_font(pixel_size, bold=bold)
            lines = self._wrap_preview_text(text or "", font, width)
            needed_height = len(lines) * font.metrics("linespace") + 4
            if needed_height <= max(20, int(height)):
                chosen_lines = lines
                chosen_px = pixel_size
                break
            chosen_lines = lines
            chosen_px = pixel_size
        weight = "bold" if bold else "normal"
        return "\n".join(chosen_lines), (self.font_name.get().strip() or "微软雅黑", -chosen_px, weight)

    def _render_pdf_question_editor_preview(self):
        canvas = getattr(self, "_pdf_question_preview_canvas", None)
        if canvas is None:
            return

        self._pdf_question_preview_photos = []
        canvas.delete("all")
        width = max(460, canvas.winfo_width() or 520)
        height = max(300, canvas.winfo_height() or 320)
        canvas.create_rectangle(0, 0, width, height, fill="#edf2f7", outline="")

        question = self._pdf_question_editor_target
        if question is None:
            canvas.create_text(
                width / 2,
                height / 2,
                text="选择左侧一道题后，这里会显示当前题目的实时排版预览。",
                width=width - 40,
                justify=CENTER,
                fill="#6b7280",
                font=("", 11),
            )
            return

        outer_margin = 14
        slide_ratio = 13.333 / 7.5
        available_width = max(300, width - outer_margin * 2)
        available_height = max(200, height - outer_margin * 2)
        if available_width / available_height > slide_ratio:
            slide_height = available_height
            slide_width = slide_height * slide_ratio
        else:
            slide_width = available_width
            slide_height = slide_width / slide_ratio

        slide_left = (width - slide_width) / 2
        slide_top = (height - slide_height) / 2
        slide_right = slide_left + slide_width
        slide_bottom = slide_top + slide_height

        scale = slide_width / 13.333
        self._pdf_preview_scale_px_per_in = scale
        margin_left = self.margin_left.get() * scale
        margin_right = self.margin_right.get() * scale
        margin_top = self.margin_top.get() * scale
        content_left = slide_left + margin_left
        content_top = slide_top + margin_top
        content_width = max(120, slide_width - margin_left - margin_right)

        canvas.create_rectangle(
            slide_left,
            slide_top,
            slide_right,
            slide_bottom,
            fill="#ffffff",
            outline="#b8c2cc",
            width=2,
        )

        stem_text = f"{question.source_number or '-'}".strip(".")
        stem_prefix = f"{stem_text}. " if stem_text else ""
        stem_value = f"{stem_prefix}{question.stem or '未填写题干'}".strip()
        has_image = bool(question.stem_assets)
        stem_height = (
            self.stem_h_img.get() if has_image else self.stem_h_no.get()
        ) * scale
        cursor_y = content_top
        wrapped_stem, stem_font = self._fit_preview_text(
            stem_value,
            width=max(48, content_width - 16),
            height=max(40, stem_height - 12),
            target_points=float(self.font_size_stem.get()),
            scale_px_per_in=scale,
            bold=self.stem_bold.get(),
        )
        canvas.create_rectangle(
            content_left,
            cursor_y,
            content_left + content_width,
            cursor_y + stem_height,
            fill="#eef6ff",
            outline="#4a90d9",
            width=2,
        )
        stem_align = (self.stem_align.get() or "left").lower()
        if stem_align == "center":
            stem_anchor = tk.N
            stem_justify = CENTER
            stem_x = content_left + content_width / 2
        elif stem_align == "right":
            stem_anchor = tk.NE
            stem_justify = RIGHT
            stem_x = content_left + content_width - 8
        else:
            stem_anchor = tk.NW
            stem_justify = LEFT
            stem_x = content_left + 8
        canvas.create_text(
            stem_x,
            cursor_y + 8,
            anchor=stem_anchor,
            width=max(40, content_width - 16),
            text=wrapped_stem,
            justify=stem_justify,
            fill=self.color_stem.get().strip() or "#1A1A2E",
            font=stem_font,
        )

        cursor_y += stem_height + self.gap_stem.get() * scale
        if has_image:
            image_band_height = min(self.image_max_h.get() * scale, slide_height * 0.22)
            image_band_width = min(self.image_max_w.get() * scale, content_width)
            image_align = (self.image_align.get() or "center").lower()
            if image_align == "right":
                image_left = content_left + content_width - image_band_width
            elif image_align == "left":
                image_left = content_left
            else:
                image_left = content_left + (content_width - image_band_width) / 2
            canvas.create_rectangle(
                image_left,
                cursor_y,
                image_left + image_band_width,
                cursor_y + image_band_height,
                fill="#fff8e6",
                outline="#d9a84a",
                width=2,
                dash=(5, 3),
            )
            canvas.create_text(
                image_left + image_band_width / 2,
                cursor_y + image_band_height / 2,
                text=f"题干图片 × {len(question.stem_assets)}",
                fill="#996600",
                font=("", 10, "bold"),
            )
            cursor_y += image_band_height + self.gap_img.get() * scale

        options_top = cursor_y + self.gap_opts.get() * scale
        options_height = max(56, slide_bottom - options_top - 12)
        layout = self._effective_question_option_layout(question)
        options = list(question.options[:4])
        fills = ["#e3f2fd", "#e8f5e9", "#fce4ec", "#fff3e0"]
        outlines = ["#1976d2", "#43a047", "#c62828", "#e65100"]

        if not options:
            canvas.create_text(
                width / 2,
                options_top + options_height / 2,
                text="当前题目没有可预览的选项。",
                fill="#6b7280",
                font=("", 10),
            )
        elif layout == "one_row":
            gap = max(4, self.one_row_gap.get() * scale)
            cell_width = (content_width - gap * (len(options) - 1)) / max(1, len(options))
            row_height = min(options_height, max(40, self.one_row_h.get() * scale))
            row_top = options_top + max(0, (options_height - row_height) / 2)
            for index, option in enumerate(options):
                x0 = content_left + index * (cell_width + gap)
                self._draw_pdf_question_preview_option(
                    canvas,
                    option,
                    x0,
                    row_top,
                    cell_width,
                    row_height,
                    fills[index],
                    outlines[index],
                )
        elif layout == "list":
            gap = max(4, scale * 0.06)
            row_height = max(40, self.list_row_h.get() * scale)
            needed_height = row_height * len(options) + gap * max(0, len(options) - 1)
            shrink = min(1.0, options_height / needed_height) if needed_height else 1.0
            row_height *= shrink
            gap *= shrink
            for index, option in enumerate(options):
                self._draw_pdf_question_preview_option(
                    canvas,
                    option,
                    content_left,
                    options_top + index * (row_height + gap),
                    content_width,
                    row_height,
                    fills[index],
                    outlines[index],
                )
        else:
            gap = max(6, self.grid_col_gap.get() * scale)
            col_width = (content_width - gap) / 2
            row_height = max(42, self.grid_row_h.get() * scale)
            needed_height = row_height * 2 + gap
            shrink = min(1.0, options_height / needed_height) if needed_height else 1.0
            row_height *= shrink
            gap *= shrink
            for index, option in enumerate(options):
                row = index // 2
                col = index % 2
                self._draw_pdf_question_preview_option(
                    canvas,
                    option,
                    content_left + col * (col_width + gap),
                    options_top + row * (row_height + gap),
                    col_width,
                    row_height,
                    fills[index],
                    outlines[index],
                )

        layout_label = self._option_layout_label(question.option_layout or layout)
        layout_source = "单题覆盖" if question.option_layout else "跟随全局"
        canvas.create_text(
            width - 18,
            height - 16,
            anchor=tk.SE,
            text=f"{layout_source} · {layout_label}",
            fill="#6b7280",
            font=("", 9),
        )

    def _selected_pdf_payload(self) -> dict:
        if not getattr(self, "pdf_tree", None):
            return {}
        selected = self.pdf_tree.selection()
        if not selected:
            return {}
        return self._pdf_preview_payloads.get(selected[0], {})

    def _reset_pdf_material_preview_session(self):
        if self._pdf_material_preview_dir and os.path.isdir(self._pdf_material_preview_dir):
            shutil.rmtree(self._pdf_material_preview_dir, ignore_errors=True)
        self._pdf_material_preview_dir = None
        self._pdf_material_preview_cache.clear()
        self._pdf_material_preview_paths = []
        self._pdf_material_preview_source = ""
        self._pdf_material_preview_title = ""
        self._pdf_material_preview_index = 0
        self._pdf_material_preview_photo = None
        self._set_pdf_material_preview_message(
            "选择资料分析材料或题目后，可查看 PDF 区域原貌。",
            "暂无材料原貌",
        )

    def _ensure_pdf_material_preview_dir(self) -> str:
        if self._pdf_material_preview_dir and os.path.isdir(self._pdf_material_preview_dir):
            return self._pdf_material_preview_dir
        self._pdf_material_preview_dir = tempfile.mkdtemp(prefix="pptconvert_gui_material_preview_")
        return self._pdf_material_preview_dir

    def _set_pdf_material_preview_message(self, message: str, status: str):
        self._pdf_material_preview_photo = None
        if getattr(self, "_pdf_material_preview_status", None):
            self._pdf_material_preview_status.configure(text=status)
        if getattr(self, "_pdf_material_preview_prev", None):
            self._pdf_material_preview_prev.configure(state=DISABLED)
        if getattr(self, "_pdf_material_preview_next", None):
            self._pdf_material_preview_next.configure(state=DISABLED)
        if getattr(self, "_pdf_material_preview_box", None):
            self._pdf_material_preview_box.configure(image="", text=message)

    def _current_material_preview_target(self, payload: dict):
        if payload.get("kind") == "material":
            return payload.get("material")
        if payload.get("kind") == "question" and payload.get("section_kind") == "data":
            return payload.get("material")
        return None

    def _material_preview_entries(self, material) -> tuple[str, list[str]]:
        cache_key = str(material.material_id or id(material))
        cached = self._pdf_material_preview_cache.get(cache_key)
        if cached:
            return cached

        source = ""
        paths: list[str] = []
        pdf_path = self._pdf_project_context.get("pdf_path") or ""
        if pdf_path and material.body_regions:
            paths = crop_material_regions(
                pdf_path,
                material,
                self._ensure_pdf_material_preview_dir(),
                dpi=144,
            )
            if paths:
                source = "PDF 区域预览"
        if not paths and material.body_assets:
            paths = [
                asset.path
                for asset in material.body_assets
                if asset.path and os.path.exists(asset.path)
            ]
            if paths:
                source = "材料图片"

        result = (source, paths)
        self._pdf_material_preview_cache[cache_key] = result
        return result

    def _show_pdf_material_preview_for_payload(self, payload: dict):
        material = self._current_material_preview_target(payload)
        if material is None:
            self._pdf_material_preview_paths = []
            self._pdf_material_preview_source = ""
            self._pdf_material_preview_title = ""
            self._pdf_material_preview_index = 0
            self._pdf_material_preview_photo = None
            self._set_pdf_material_preview_message(
                "选择资料分析材料或题目后，可查看 PDF 区域原貌。",
                "暂无材料原貌",
            )
            return

        source, paths = self._material_preview_entries(material)
        if not paths:
            self._pdf_material_preview_paths = []
            self._pdf_material_preview_source = ""
            self._pdf_material_preview_title = ""
            self._pdf_material_preview_index = 0
            self._pdf_material_preview_photo = None
            self._set_pdf_material_preview_message(
                "当前材料没有可用的区域截图或图片素材，请先参考下方结构化文本。",
                f"{material.header or material.material_id}：暂无可视预览",
            )
            return

        self._pdf_material_preview_paths = paths
        self._pdf_material_preview_source = source
        self._pdf_material_preview_title = material.header or material.material_id
        self._pdf_material_preview_index = 0
        self._render_pdf_material_preview()

    def _step_pdf_material_preview(self, delta: int):
        if not self._pdf_material_preview_paths:
            return
        self._pdf_material_preview_index = (
            self._pdf_material_preview_index + delta
        ) % len(self._pdf_material_preview_paths)
        self._render_pdf_material_preview()

    def _render_pdf_material_preview(self, title: str = ""):
        if not getattr(self, "_pdf_material_preview_box", None):
            return
        if not self._pdf_material_preview_paths:
            return
        index = min(self._pdf_material_preview_index, len(self._pdf_material_preview_paths) - 1)
        image_path = self._pdf_material_preview_paths[index]
        if not os.path.exists(image_path):
            self._set_pdf_material_preview_message(
                "预览图片不存在，请重新生成预览。",
                "材料原貌不可用",
            )
            return

        target_width = self._pdf_material_preview_box.winfo_width()
        max_width = max(260, target_width - 20) if target_width > 40 else 420
        with Image.open(image_path) as source_image:
            image = source_image.copy()
        image.thumbnail((max_width, 240), Image.Resampling.LANCZOS)
        self._pdf_material_preview_photo = ImageTk.PhotoImage(image)
        self._pdf_material_preview_box.configure(image=self._pdf_material_preview_photo, text="")

        preview_title = title or self._pdf_material_preview_title or "材料原貌"
        self._pdf_material_preview_status.configure(
            text=(
                f"{preview_title} · {self._pdf_material_preview_source} "
                f"{index + 1}/{len(self._pdf_material_preview_paths)}"
            )
        )
        state = NORMAL if len(self._pdf_material_preview_paths) > 1 else DISABLED
        self._pdf_material_preview_prev.configure(state=state)
        self._pdf_material_preview_next.configure(state=state)

    def _refresh_pdf_preview_after_edit(self, detail_text: str | None = None):
        if self.pdf_project is None:
            return
        self._reset_pdf_material_preview_session()
        self._clear_pdf_question_editor()
        self._populate_pdf_preview(self.pdf_project)
        if detail_text:
            self._set_pdf_detail(detail_text)

    def _edit_selected_question_number(self):
        payload = self._selected_pdf_payload()
        if payload.get("kind") != "question":
            messagebox.showinfo("提示", "请先在左侧选择一道题目")
            return
        question = payload.get("question")
        if question is None:
            return
        new_number = simpledialog.askstring(
            "修改题号",
            "输入新的原题号：",
            initialvalue=question.source_number or "",
            parent=self.root,
        )
        if new_number is None:
            return
        renumber_question(question, new_number)
        self._mark_pdf_project_dirty()
        self._refresh_pdf_preview_after_edit("已更新题号。")

    def _rename_selected_material(self):
        payload = self._selected_pdf_payload()
        if payload.get("kind") != "material":
            messagebox.showinfo("提示", "请先选择一个材料节点")
            return
        material = payload.get("material")
        if material is None:
            return
        new_header = simpledialog.askstring(
            "修改材料标题",
            "输入新的材料标题：",
            initialvalue=material.header or "",
            parent=self.root,
        )
        if new_header is None:
            return
        rename_material(material, new_header)
        self._mark_pdf_project_dirty()
        self._refresh_pdf_preview_after_edit("已更新材料标题。")

    def _insert_material_after_selection(self):
        payload = self._selected_pdf_payload()
        material = payload.get("material")
        if material is None:
            messagebox.showinfo("提示", "请先选择一个材料节点，或选择资料分析中的一道题目")
            return
        if payload.get("section_kind") != "data":
            messagebox.showinfo("提示", "只有资料分析部分支持新增材料组")
            return
        new_header = simpledialog.askstring(
            "新建材料",
            "输入新材料标题：",
            initialvalue="新材料",
            parent=self.root,
        )
        if new_header is None:
            return
        if insert_material_after(self.pdf_project, material, new_header):
            self._mark_pdf_project_dirty()
            self._refresh_pdf_preview_after_edit("已在当前材料后方新建材料组。")

    def _merge_selected_material(self, direction: int):
        payload = self._selected_pdf_payload()
        material = payload.get("material")
        if material is None:
            messagebox.showinfo("提示", "请先选择一个材料节点，或选择资料分析中的一道题目")
            return
        if payload.get("section_kind") != "data":
            messagebox.showinfo("提示", "只有资料分析部分支持合并材料组")
            return
        direction_text = "下一材料" if direction > 0 else "上一材料"
        if not messagebox.askyesno("确认合并", f"确定与{direction_text}合并吗？"):
            return
        merged = merge_adjacent_materials(self.pdf_project, material, direction)
        if not merged:
            messagebox.showinfo("提示", "当前材料已经在边界位置，无法继续合并")
            return
        self._mark_pdf_project_dirty()
        self._refresh_pdf_preview_after_edit("已完成材料组合并。")

    def _remove_selected_question(self):
        payload = self._selected_pdf_payload()
        if payload.get("kind") != "question":
            messagebox.showinfo("提示", "请先选择一道题目")
            return
        question = payload.get("question")
        if question is None or self.pdf_project is None:
            return
        if not messagebox.askyesno("确认删除", "确定从当前工程中移除这道题吗？"):
            return
        if remove_question(self.pdf_project, question):
            self._mark_pdf_project_dirty()
            self._refresh_pdf_preview_after_edit("已移除所选题目。")

    def _move_selected_question_between_materials(self, direction: int):
        payload = self._selected_pdf_payload()
        if payload.get("kind") != "question":
            messagebox.showinfo("提示", "请先选择资料分析中的一道题目")
            return
        if payload.get("section_kind") != "data":
            messagebox.showinfo("提示", "只有资料分析题支持跨材料移动")
            return
        question = payload.get("question")
        if question is None or self.pdf_project is None:
            return
        moved = move_data_question(self.pdf_project, question, direction)
        if not moved:
            messagebox.showinfo("提示", "当前题目已经在边界材料中，无法继续移动")
            return
        self._mark_pdf_project_dirty()
        self._refresh_pdf_preview_after_edit("已调整题目所属材料。")

    def _material_preview_text(self, material) -> str:
        lines = [
            f"材料编号：{material.material_id}",
            f"标题：{material.header or '-'}",
            f"题目数：{len(material.questions)}",
            f"材料区域数：{len(material.body_regions)}",
            "",
            "材料正文：",
            material.body or "-",
        ]
        if material.body_regions:
            lines.append("")
            lines.append("页面区域：")
            for region in material.body_regions:
                lines.append(
                    f"第 {region.page_number} 页  ({region.x0:.1f}, {region.y0:.1f}) - ({region.x1:.1f}, {region.y1:.1f})"
                )
        return "\n".join(lines)

    def _question_preview_text(self, section, material, question) -> str:
        effective_layout = self._effective_question_option_layout(question)
        layout_text = self._option_layout_label(question.option_layout or effective_layout)
        layout_source = "单题覆盖" if question.option_layout else "跟随全局"
        lines = [
            f"科目：{section.kind}",
            f"原题号：{question.source_number or '-'}",
            f"选项数：{len(question.options)}",
            f"题干图片：{len(question.stem_assets)}",
            f"选项布局：{layout_source}（{layout_text}）",
        ]
        if material is not None:
            lines.append(f"所属材料：{material.header or material.material_id}")
        if question.page_numbers:
            lines.append("来源页码：" + ", ".join(str(page_no) for page_no in question.page_numbers))
        lines.extend(["", "题干：", question.stem or "-"])
        if question.options:
            lines.extend(["", "选项："])
            for option in question.options:
                suffix = " [图]" if option.image_path else ""
                lines.append(f"{option.letter}. {option.text}{suffix}")
        return "\n".join(lines)

    def _selected_preview_section(self):
        payload = self._selected_pdf_payload()
        return payload.get("section")

    def _sync_selected_section_subject(self, payload: dict | None = None) -> None:
        section = (payload or {}).get("section") if payload else self._selected_preview_section()
        if section is None:
            self._pdf_section_subject_var.set("待确认科目")
            return
        self._pdf_section_subject_var.set(self._document_subject_label(getattr(section, "kind", "unknown")))

    def _reclassify_selected_section(self):
        if self.pdf_project is None:
            return
        payload = self._selected_pdf_payload()
        section = payload.get("section")
        if section is None:
            messagebox.showinfo("提示", "请先在左侧选中一个篇题或该篇题下的题目。")
            return
        if section.kind == "data":
            messagebox.showinfo("提示", "资料分析暂不支持在这里整段改科目；建议回到导入设置直接固定整份文档科目。")
            return
        new_kind = self._document_subject_key(self._pdf_section_subject_var.get())
        if new_kind == "auto":
            new_kind = "unknown"
        if not reclassify_objective_section(section, new_kind):
            return
        self._mark_pdf_project_dirty()
        self._refresh_pdf_preview_after_edit(f"已将当前篇题调整为：{self._document_subject_label(new_kind)}")

    def _populate_pdf_preview(self, project):
        self._pdf_preview_payloads.clear()
        self._clear_pdf_question_editor()
        for item in self.pdf_tree.get_children():
            self.pdf_tree.delete(item)

        section_index = 0
        for section in project.sections:
            section_index += 1
            count = len(section.questions) if section.kind != "data" else sum(
                len(material.questions) for material in section.material_sets
            )
            section_id = self.pdf_tree.insert(
                "",
                END,
                text=section.title or f"Section {section_index}",
                values=(section.kind, "", count),
                open=True,
            )
            self._pdf_preview_payloads[section_id] = {
                "kind": "section",
                "section": section,
                "section_kind": section.kind,
                "text": f"篇题：{section.title}\n科目：{section.kind}\n题目数：{count}",
            }

            if section.kind == "data":
                for material in section.material_sets:
                    material_id = self.pdf_tree.insert(
                        section_id,
                        END,
                        text=material.header or material.material_id,
                        values=("material", "", len(material.questions)),
                        open=True,
                    )
                    self._pdf_preview_payloads[material_id] = {
                        "kind": "material",
                        "section": section,
                        "section_kind": section.kind,
                        "material": material,
                        "text": self._material_preview_text(material),
                    }
                    for question in material.questions:
                        qid = self.pdf_tree.insert(
                            material_id,
                            END,
                            text=self._question_tree_label(question),
                            values=("question", question.source_number or "-", len(question.options)),
                        )
                        self._pdf_preview_payloads[qid] = {
                            "kind": "question",
                            "section": section,
                            "section_kind": section.kind,
                            "material": material,
                            "question": question,
                            "text": self._question_preview_text(section, material, question),
                        }
            else:
                for question in section.questions:
                    qid = self.pdf_tree.insert(
                        section_id,
                        END,
                        text=self._question_tree_label(question),
                        values=("question", question.source_number or "-", len(question.options)),
                    )
                    self._pdf_preview_payloads[qid] = {
                        "kind": "question",
                        "section": section,
                        "section_kind": section.kind,
                        "material": None,
                        "question": question,
                        "text": self._question_preview_text(section, None, question),
                    }

        self._set_pdf_detail(f"已加载 {project.question_count} 道题。\n请在左侧查看篇题、材料和题目结构。")
        self._refresh_pdf_wizard_ui()

    def _on_pdf_preview_select(self, _event=None):
        if not getattr(self, "pdf_tree", None):
            return
        selected = self.pdf_tree.selection()
        if not selected:
            self._clear_pdf_question_editor()
            self._show_pdf_material_preview_for_payload({})
            self._sync_selected_section_subject({})
            return
        payload = self._pdf_preview_payloads.get(selected[0], {})
        self._set_pdf_detail(payload.get("text", "暂无内容"))
        self._sync_selected_section_subject(payload)
        self._populate_pdf_question_editor(payload)
        self._show_pdf_material_preview_for_payload(payload)

    def _close_parser(self):
        if self.parser:
            self.parser.cleanup()
            self.parser = None

    def _word_project_matches_current_file(self) -> bool:
        return (
            self.pdf_project is not None
            and self._pdf_project_context.get("source_kind") == "word"
            and self._pdf_project_context.get("docx_path") == self.word_path.get().strip()
            and self._pdf_project_context.get("document_subject_hint", "auto")
            == self._document_subject_key(self.word_document_subject.get())
        )

    def _load_word_project_into_preview(self, project, *, docx_path: str, asset_dir: str):
        docx_base = os.path.splitext(os.path.abspath(docx_path))[0]
        default_ppt = self.output_path.get().strip() or docx_base + ".pptx"
        default_manifest = self.pdf_manifest_out.get().strip() or docx_base + "_工程.json"
        self.output_path.set(default_ppt)
        self.pdf_ppt_out.set(default_ppt)
        self.pdf_manifest_out.set(default_manifest)
        self.pdf_project = project
        self._pdf_project_dirty = False
        self._pdf_project_context = {
            "pdf_path": "",
            "docx_path": docx_path,
            "asset_dir": asset_dir,
            "subject_spec": "all",
            "range_spec": "",
            "source_kind": "word",
            "manifest_path": default_manifest,
            "document_subject_hint": self._document_subject_key(self.word_document_subject.get()),
        }
        self._reset_pdf_material_preview_session()
        self._populate_pdf_preview(project)

    def _make_ppt_config(self) -> PPTConfig:
        from pptx.util import Inches, Pt

        values = {
            "margin_left_in": self.margin_left.get(),
            "margin_right_in": self.margin_right.get(),
            "margin_top_in": self.margin_top.get(),
            "stem_height_with_image_in": self.stem_h_img.get(),
            "stem_height_no_image_in": self.stem_h_no.get(),
            "gap_after_stem_in": self.gap_stem.get(),
            "gap_after_image_in": self.gap_img.get(),
            "gap_before_options_in": self.gap_opts.get(),
            "stem_align": self.stem_align.get(),
            "image_h_align": self.image_align.get(),
            "image_max_width": Inches(self.image_max_w.get()),
            "image_max_height": Inches(self.image_max_h.get()),
            "option_layout": self.option_layout.get(),
            "grid_layout": self.grid_layout.get(),
            "grid_row_height_in": self.grid_row_h.get(),
            "grid_col_gap_in": self.grid_col_gap.get(),
            "list_row_height_in": self.list_row_h.get(),
            "one_row_height_in": self.one_row_h.get(),
            "one_row_gap_in": self.one_row_gap.get(),
            "option_align": self.option_align.get(),
            "font_name": self.font_name.get().strip() or "微软雅黑",
            "stem_font_size": Pt(self.font_size_stem.get()),
            "option_font_size": Pt(self.font_size_option.get()),
            "font_bold_stem": self.stem_bold.get(),
            "option_letter_bold": self.option_letter_bold.get(),
            "option_font_bold": self.option_text_bold.get(),
        }
        for attr, var in [
            ("stem_color", self.color_stem),
            ("option_color", self.color_option),
            ("option_letter_color", self.color_letter),
        ]:
            rgb = parse_hex_color(var.get())
            if rgb:
                values[attr] = rgb
        return PPTConfig.from_mapping(values)

    def _parse_word_file(self) -> bool:
        word_file = self.word_path.get().strip()
        if not word_file:
            messagebox.showwarning("提示", "请先选择 Word 文件")
            return False
        if not os.path.exists(word_file):
            messagebox.showerror("错误", f"文件不存在：{word_file}")
            return False
        if not self._confirm_discard_pdf_project_edits("重新解析 Word 结构"):
            return False

        self._set_status("正在解析...")
        try:
            self._close_parser()
            project, _parsed_questions, asset_dir = build_word_project(
                word_file,
                document_subject_hint=(
                    None
                    if self._document_subject_key(self.word_document_subject.get()) == "auto"
                    else self._document_subject_key(self.word_document_subject.get())
                ),
            )
            self.questions = project_to_ppt_questions(project)
            self._refresh_question_tree()
            self._load_word_project_into_preview(project, docx_path=word_file, asset_dir=asset_dir)
            self._set_status(f"解析完成，共 {project.question_count} 道题")
            return True
        except Exception as exc:
            self._set_status("解析失败")
            messagebox.showerror("解析错误", str(exc))
            return False

    def _parse_preview(self):
        if not self._parse_word_file():
            return
        self._open_pdf_preview_workspace()
        if not self.questions:
            messagebox.showinfo("提示", "未解析到任何题目，请检查 Word 格式")

    def _convert_all(self):
        word_file = self.word_path.get().strip()
        if not word_file:
            messagebox.showwarning("提示", "请先选择 Word 文件")
            return
        if not self.output_path.get().strip():
            self.output_path.set(os.path.splitext(word_file)[0] + ".pptx")
        if not self._parse_word_file():
            return
        if not self.questions:
            messagebox.showinfo("提示", "未解析到任何题目")
            return
        self._generate_ppt()

    def _generate_ppt(self):
        output = self.output_path.get().strip()
        if not output:
            messagebox.showwarning("提示", "请设置输出路径")
            return

        project = self.pdf_project if self._word_project_matches_current_file() else None
        if project is None:
            if not self._parse_word_file():
                return
            project = self.pdf_project if self._word_project_matches_current_file() else None
        if project is None:
            messagebox.showwarning("提示", "请先解析 Word 文件")
            return

        template = None
        if self.use_template.get():
            template = self.template_path.get().strip()
            if not template or not os.path.exists(template):
                messagebox.showerror("错误", "请选择有效的模板文件")
                return

        render_questions = project_to_ppt_questions(project)
        self.questions = render_questions
        self._set_status("正在生成 PPT...")
        self.progress["value"] = 0
        self.progress["maximum"] = len(render_questions)

        def work():
            try:
                config = self._make_ppt_config()
                generator = PPTGenerator(config=config)
                generator.generate(
                    render_questions,
                    output,
                    template_path=template,
                    progress_callback=self._on_progress,
                )
                self.root.after(0, lambda: self._on_done(output))
            except Exception as exc:
                self.root.after(0, lambda e=exc: self._on_error(str(e)))

        threading.Thread(target=work, daemon=True).start()

    def _on_progress(self, current, total):
        self.root.after(0, lambda: self._set_progress(current, total))

    def _set_progress(self, current, total):
        self.progress["value"] = current
        self._set_status(f"生成中：{current}/{total}")

    def _on_done(self, path):
        self.progress["value"] = self.progress["maximum"]
        self._set_status(f"完成：{path}")
        if messagebox.askyesno("完成", f"PPT 已生成：\n{path}\n\n打开所在目录？"):
            os.startfile(os.path.dirname(os.path.abspath(path)))

    def _on_error(self, msg):
        self._set_status("生成失败")
        messagebox.showerror("生成错误", msg)

    def _clear_all(self):
        if not self._confirm_discard_pdf_project_edits("清空当前工作区"):
            return
        self.word_path.set("")
        self.output_path.set("")
        self.template_path.set("")
        self.pdf_path.set("")
        self.pdf_word_out.set("")
        self.pdf_ppt_out.set("")
        self.pdf_manifest_out.set("")
        self.pdf_question_range.set("")
        self.word_document_subject.set(_DOCUMENT_SUBJECT_LABELS["auto"])
        self.pdf_document_subject.set(_DOCUMENT_SUBJECT_LABELS["auto"])
        self._pdf_section_subject_var.set("待确认科目")
        self._set_all_pdf_subjects(True)
        self.use_template.set(False)
        self._toggle_template()
        self.questions.clear()
        for item in self.tree.get_children():
            self.tree.delete(item)
        self._clear_pdf_preview()
        self._show_pdf_wizard_step(0)
        self.progress["value"] = 0
        self._set_status("已清空")
        self._close_parser()

    def _set_status(self, text: str):
        self.status_label.configure(text=text)

    def _on_close(self):
        if not self._confirm_discard_pdf_project_edits("关闭窗口"):
            return
        self._close_parser()
        self._reset_pdf_material_preview_session()
        self.root.destroy()

    def run(self):
        self.root.mainloop()
