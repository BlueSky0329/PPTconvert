import copy
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
from core.ai_repair import (
    AIRepairService,
    apply_ai_question_patch,
    build_repair_strategy_summary,
    inspect_repair_strategy,
    repair_question_boundary,
    repair_project_questions,
)
from core.pdf_ocr_diagnostics import auto_repair_ocr_project, diagnose_project_ocr_risks
from core.repair_log import append_project_repair_log, append_question_repair_log, capture_question_state
from core.subject_inference import infer_subject_diagnostics
from domain.models import ALL_SUBJECT_KINDS, SUBJECT_DISPLAY_NAMES
from domain.project_editor import (
    apply_all_safe_subject_suggestions,
    apply_section_subject_suggestion,
    clear_question_ppt_layout,
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
    section_subject_suggestion,
    set_question_ppt_layout_block,
    set_question_option_layout,
    update_option_text,
    update_question_stem,
)
from core.ppt_layout import (
    build_effective_question_layout,
    is_option_layout_block,
    normalize_layout_rect,
    option_layout_block_key,
    option_layout_block_letter,
    scale_layout_rect,
)
from core.ppt_generator import PPTGenerator, PPTConfig
from core.ppt_style import parse_hex_color
from core.template_manager import TemplateManager
from core.template_style import (
    describe_template_style,
    extract_best_style_from_presentation,
    template_style_has_full_layout,
)
from core.project_quality import (
    annotate_project_quality,
    is_flagged_question,
    iter_flagged_question_rows,
    question_max_severity,
    question_review_summary,
    severity_rank,
)
from core.models import Option, Question
from exporters.manifest_json import load_project_manifest_project
from exporters.material_crops import crop_material_regions, crop_page_regions
from exporters.pptx_slides import iter_project_question_nodes, project_to_ppt_questions
from exporters.review_report import export_quality_report
from gui.font_data import build_font_values
from gui import ui_constants as U
from workflows.project_flow import build_word_project

_PAD = 14
_PDF_WIZARD_STEPS = (
    ("导入 PDF", "选择试卷文件，向导会按文件名预填默认输出路径。"),
    ("识别设置", "决定要进入工程的题目范围；下一步会生成或刷新结构化预览。"),
    ("结果预览", "校对题号、材料和题目归属；这一步的人工修正会直接用于导出。"),
    ("导出结果", "从当前题目工程导出 Word / JSON；需要做 PPT 时，再把 Word 交给 Word 工作流继续解析。"),
)
_PPT_SLIDE_WIDTH_IN = 13.333
_PPT_SLIDE_HEIGHT_IN = 7.5
_PPT_LAYOUT_FIELD_STEP_IN = 0.05
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
_AI_MODE_CHOICES = (
    ("balanced", "规则优先"),
    ("policy", "策略增强"),
)
_AI_MODE_LABELS = {key: label for key, label in _AI_MODE_CHOICES}
_UI_PRIMARY = "#103f91"
_UI_PRIMARY_DARK = "#0b2b63"
_UI_INFO = "#0f766e"
_UI_WARNING = "#c96a09"
_UI_DANGER = "#b9382f"
_UI_SURFACE = "#f4f8ff"
_UI_TEXT_LIGHT = "#f8fbff"
_UI_MUTED = "#6b7280"


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
        self._template_status_var = tk.StringVar(value="未启用模板。")
        self._word_flow_status_var = tk.StringVar(value="先选题本，再解析检查；正常情况下下一步就是生成 PPT。")
        self._word_results_hint_var = tk.StringVar(value="还没有解析结果。先在左侧选择题本。")
        self.word_document_subject = tk.StringVar(value=_DOCUMENT_SUBJECT_LABELS["auto"])
        self.pdf_document_subject = tk.StringVar(value=_DOCUMENT_SUBJECT_LABELS["auto"])
        self.ai_only_flagged = tk.BooleanVar(value=True)
        self.ai_batch_limit = tk.IntVar(value=12)
        _default_ai_mode = os.environ.get("PPTCONVERT_AI_REPAIR_MODE", "balanced").strip().lower() or "balanced"
        self.ai_mode = tk.StringVar(value=_AI_MODE_LABELS.get(_default_ai_mode, _AI_MODE_LABELS["balanced"]))
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
        self._template_manager = TemplateManager()
        self._template_style_preview = None
        self._template_mode = "off"
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
        self._pdf_slide_payload_ids: list[str] = []
        self._pdf_slide_payload_to_item_id: dict[str, str] = {}
        self._pdf_slide_item_to_payload_id: dict[str, str] = {}
        self._pdf_slide_payload_to_number: dict[str, int] = {}
        self._pdf_review_payload_ids: list[str] = []
        self._pdf_review_payload_to_item_id: dict[str, str] = {}
        self._pdf_review_item_to_payload_id: dict[str, str] = {}
        self._pdf_preview_syncing_selection = False
        self._pdf_slide_syncing_selection = False
        self._pdf_review_syncing_selection = False

        self.pdf_path = tk.StringVar()
        self.pdf_word_out = tk.StringVar()
        self.pdf_ppt_out = tk.StringVar()
        self.pdf_manifest_out = tk.StringVar()
        self.pdf_question_range = tk.StringVar()
        self._pdf_question_layout_var = tk.StringVar(value="")
        self._pdf_question_editor_message = tk.StringVar(value="选择一道题后，可在这里实时修改题干，并为该题单独切换选项布局。")
        self._pdf_layout_editor_status_var = tk.StringVar(value="选择一道题后，可像编辑 PPT 一样拖动题干区、图片区和选项区。")
        self._pdf_layout_x_var = tk.DoubleVar(value=0.0)
        self._pdf_layout_y_var = tk.DoubleVar(value=0.0)
        self._pdf_layout_w_var = tk.DoubleVar(value=0.0)
        self._pdf_layout_h_var = tk.DoubleVar(value=0.0)
        self._pdf_ai_suggestion_var = tk.StringVar(value="AI 建议会在这里显示，当前先聚焦待确认题。")
        self._pdf_ai_strategy_var = tk.StringVar(value="修复策略会在这里显示。")
        self._ai_status_var = tk.StringVar(value="本地 AI 修复会直接写回当前工程，不依赖外部接口。")
        self._pdf_ocr_status_var = tk.StringVar(value="扫描/OCR 诊断会在这里显示。")
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
        self._pdf_question_editor_baseline_stem = ""
        self._pdf_option_editor_baseline_texts: dict[str, str] = {}
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
        self._pdf_preview_rects: dict[str, tuple[float, float, float, float]] = {}
        self._pdf_preview_slide_bounds: tuple[float, float, float, float] = (0.0, 0.0, 0.0, 0.0)
        self._pdf_preview_selected_block: str | None = None
        self._pdf_preview_drag_state: dict[str, object] | None = None
        self._pdf_preview_action_buttons: list[ttk.Button] = []
        self._pdf_layout_value_inputs: list[ttk.Spinbox] = []
        self._pdf_layout_field_buttons: list[ttk.Button] = []
        self._pdf_layout_fields_updating = False
        self._pdf_slide_status_var = tk.StringVar(value="选择左侧 PPT 页后，可以逐页预览并实时编辑。")
        self._pdf_review_status_var = tk.StringVar(value="AI 质检会在预览生成后自动标出待确认题目。")
        self._pdf_workspace_source_var = tk.StringVar(value="未选择来源")
        self._pdf_workspace_scope_var = tk.StringVar(value="未设置筛选")
        self._pdf_workspace_state_var = tk.StringVar(value="尚未生成预览。")
        self._ai_repair_busy = False
        self._ocr_tool_busy = False

        self._build_ui()
        self._bind_pdf_wizard_updates()
        self._bind_preview_updates()
        try:
            self._pdf_question_layout_var.trace_add("write", lambda *_: self._on_pdf_question_layout_change())
        except Exception:
            pass
        for variable in (self.ai_only_flagged, self.ai_mode):
            try:
                variable.trace_add("write", lambda *_: self._refresh_pdf_ai_suggestion())
            except Exception:
                pass
        for variable in (self.word_path, self.output_path, self.word_document_subject):
            try:
                variable.trace_add("write", lambda *_: self._refresh_word_flow_ui())
            except Exception:
                pass
        self._refresh_word_flow_ui()
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
        self._pdf_workspace_tab = pdf_tab
        self._word_workspace_tab = word_tab
        self._settings_workspace_tab = None
        self._ppt_settings_tab = word_tab

        notebook.add(pdf_tab, text=" PDF 试卷整理 ")
        notebook.add(word_tab, text=" Word 生成 PPT ")

        self._build_pdf_tab(pdf_body)
        self._build_word_tab(word_body)

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

    def _build_color_banner(self, parent, title: str, subtitle: str = "", *, bg: str, chips: tuple[tuple[str, str], ...] = ()):
        card = tk.Frame(parent, bg=bg, bd=0, highlightthickness=0)
        card.pack(fill=X, pady=(0, 10))

        top = tk.Frame(card, bg=bg)
        top.pack(fill=X, padx=18, pady=(16, 0))
        tk.Label(
            top,
            text=title,
            bg=bg,
            fg=_UI_TEXT_LIGHT,
            font=("", 16, "bold"),
            anchor="w",
        ).pack(side=LEFT)
        if chips:
            chip_row = tk.Frame(top, bg=bg)
            chip_row.pack(side=RIGHT)
            for text, color in chips:
                tk.Label(
                    chip_row,
                    text=text,
                    bg=color,
                    fg=_UI_TEXT_LIGHT,
                    font=("", 9, "bold"),
                    padx=10,
                    pady=4,
                ).pack(side=LEFT, padx=(8, 0))
        if subtitle:
            tk.Label(
                card,
                text=subtitle,
                bg=bg,
                fg="#dbeafe",
                font=("", 10),
                wraplength=980,
                justify=LEFT,
                anchor="w",
            ).pack(fill=X, padx=18, pady=(8, 16))
        else:
            tk.Frame(card, bg=bg, height=10).pack(fill=X)
        return card

    def _build_metric_tile(self, parent, title: str, textvariable: tk.StringVar, *, bg: str):
        card = tk.Frame(parent, bg=bg, bd=0, highlightthickness=0)
        tk.Label(
            card,
            text=title,
            bg=bg,
            fg=_UI_TEXT_LIGHT,
            font=("", 9, "bold"),
            anchor="w",
        ).pack(fill=X, padx=12, pady=(10, 4))
        tk.Label(
            card,
            textvariable=textvariable,
            bg=bg,
            fg=_UI_TEXT_LIGHT,
            font=("", 10),
            justify=LEFT,
            anchor="w",
            wraplength=260,
        ).pack(fill=X, padx=12, pady=(0, 10))
        return card

    # Header

    def _build_header(self, parent):
        hdr = ttk.Frame(parent, padding=(4, 8, 4, 8))
        hdr.pack(fill=X, pady=(0, 6))
        self._build_color_banner(
            hdr,
            U.APP_TITLE,
            bg=_UI_PRIMARY,
            chips=(
                ("PDF 整理", _UI_WARNING),
                ("Word 转 PPT", _UI_INFO),
                ("实时所见即所得", "#2563eb"),
            ),
        )

    # 1. Files

    def _build_file_section(self, parent):
        frame = ttk.Labelframe(parent, text=" ① 题本与输出 ", bootstyle="primary", padding=_PAD)
        frame.pack(fill=X, pady=(0, 10))

        self._file_row(frame, "题本 Word", self.word_path,
                       self._browse_word, "浏览...", pady=(0, 0))
        self._file_row(frame, "导出 PPT", self.output_path,
                       self._browse_output, "另存为...", pady=(6, 0))
        row = ttk.Frame(frame)
        row.pack(fill=X, pady=(8, 0))
        ttk.Label(row, text="整本科目", width=10).pack(side=LEFT)
        ttk.Combobox(
            row,
            textvariable=self.word_document_subject,
            values=[label for _key, label in _DOCUMENT_SUBJECT_CHOICES],
            state="readonly",
            width=18,
        ).pack(side=LEFT)
        ttk.Label(
            frame,
            text="单科题库或没有大标题时再固定；多数情况下保持“自动识别”即可。",
            bootstyle="secondary",
        ).pack(anchor=W, padx=(82, 0), pady=(6, 0))

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
        self.template_entry.bind("<FocusOut>", self._on_template_entry_changed)
        self.template_entry.bind("<Return>", self._on_template_entry_changed)
        self.template_btn = ttk.Button(
            row, text="选择...", command=self._browse_template,
            state=DISABLED, bootstyle="info-outline", width=8,
        )
        self.template_btn.pack(side=RIGHT)

        self._tpl_status = ttk.Label(
            frame,
            textvariable=self._template_status_var,
            bootstyle="secondary", font=("", 9), wraplength=860, justify=LEFT,
        )
        self._tpl_status.pack(anchor=W, pady=(8, 0))

    # 3. Config

    def _build_config_section(self, parent):
        self._config_frame = ttk.Labelframe(
            parent, text=" 版式与样式 ", bootstyle="primary", padding=(2, 8),
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
            text=" ④ 当前题目 ",
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
            frame, textvariable=self._word_results_hint_var, font=("", 9), bootstyle="secondary",
        ).pack(anchor=W, padx=10, pady=(2, 4))

    # PPT tab buttons + shared footer

    def _build_word_action_panel(self, parent):
        panel = ttk.Labelframe(parent, text=" ③ 生成流程 ", bootstyle="success", padding=(12, 10))
        panel.pack(fill=X, pady=(0, 10))

        ttk.Label(
            panel,
            textvariable=self._word_flow_status_var,
            wraplength=420,
            justify=LEFT,
            bootstyle="secondary",
        ).pack(anchor=W, pady=(0, 8))

        primary_row = ttk.Frame(panel)
        primary_row.pack(fill=X)

        self._word_parse_btn = ttk.Button(
            primary_row,
            text="解析并检查",
            command=self._parse_preview,
            bootstyle="info-outline",
            width=12,
        )
        self._word_parse_btn.pack(side=LEFT)
        self._word_generate_btn = ttk.Button(
            primary_row,
            text="生成 PPT",
            command=self._generate_ppt,
            bootstyle="success",
            width=10,
        )
        self._word_generate_btn.pack(side=LEFT, padx=(8, 0))
        self._word_convert_btn = ttk.Button(
            primary_row,
            text="一键生成 PPT",
            command=self._convert_all,
            bootstyle="success-outline",
            width=14,
        )
        self._word_convert_btn.pack(side=LEFT, padx=(8, 0))

        ttk.Button(
            primary_row,
            text="清空",
            command=self._clear_all,
            bootstyle="secondary-outline",
            width=7,
        ).pack(side=RIGHT)

        secondary_row = ttk.Frame(panel)
        secondary_row.pack(fill=X, pady=(8, 0))
        ttk.Label(
            secondary_row,
            text="只有要逐题改文字、配图或版式时，再进入编辑工作台：",
            bootstyle="secondary",
        ).pack(side=LEFT)
        self._word_editor_btn = ttk.Button(
            secondary_row,
            text="进入编辑工作台",
            command=self._open_word_editor_workspace,
            bootstyle="secondary-outline",
            width=14,
        )
        self._word_editor_btn.pack(side=LEFT, padx=(8, 0))

    def _build_pdf_tab(self, parent):
        frame = ttk.Frame(parent, padding=_PAD)
        frame.pack(fill=BOTH, expand=YES)
        self._build_color_banner(
            frame,
            "PDF 试卷工作流",
            bg=_UI_PRIMARY_DARK,
            chips=(
                ("先整理结构", _UI_WARNING),
                ("再进入预览", _UI_INFO),
            ),
        )

        self._build_pdf_wizard(frame)

    def _build_word_tab(self, parent):
        frame = ttk.Frame(parent, padding=_PAD)
        frame.pack(fill=BOTH, expand=YES)
        self._build_color_banner(
            frame,
            "Word 生成 PPT",
            bg="#155e75",
            chips=(
                ("先解析", "#2563eb"),
                ("直接生成", _UI_WARNING),
                ("可选精修", "#475569"),
            ),
        )

        shell = ttk.Panedwindow(frame, orient=HORIZONTAL)
        shell.pack(fill=BOTH, expand=YES)

        left = ttk.Frame(shell, padding=(0, 0, 12, 0))
        right = ttk.Frame(shell)
        shell.add(left, weight=3)
        shell.add(right, weight=5)

        self._build_file_section(left)
        self._build_word_settings_workspace(left)
        self._build_word_action_panel(left)
        self._build_question_table(right)

    def _build_word_settings_workspace(self, parent):
        frame = ttk.Labelframe(parent, text=" ② 生成设置 ", bootstyle="warning", padding=(2, 8))
        frame.pack(fill=BOTH, expand=YES, pady=(0, 10))

        notebook = ttk.Notebook(frame, bootstyle="warning")
        notebook.pack(fill=BOTH, expand=YES, padx=8, pady=(4, 8))
        self._word_settings_notebook = notebook

        template_tab = ttk.Frame(notebook, padding=(0, 0, 0, 0))
        layout_tab = ttk.Frame(notebook, padding=(0, 0, 0, 0))
        ai_tab = ttk.Frame(notebook, padding=(0, 0, 0, 0))
        notebook.add(template_tab, text=" 模板 ")
        notebook.add(layout_tab, text=" 版式与样式 ")
        notebook.add(ai_tab, text=" AI 修复 ")
        self._word_settings_template_tab = template_tab
        self._word_settings_layout_tab = layout_tab
        self._word_settings_ai_tab = ai_tab

        self._build_template_section(template_tab)
        self._build_config_section(layout_tab)
        self._build_ai_settings_section(ai_tab)

    def _build_ai_settings_section(self, parent):
        frame = ttk.Labelframe(parent, text=" AI 修复 ", bootstyle="warning", padding=_PAD)
        frame.pack(fill=X, pady=(0, 10))

        ttk.Label(
            frame,
            text="规则优先，必要时再用策略增强；修复会直接写回当前工程。",
            wraplength=860,
            justify=LEFT,
            bootstyle="secondary",
        ).pack(anchor=W, pady=(0, 8))

        row1 = ttk.Frame(frame)
        row1.pack(fill=X, pady=(0, 6))
        ttk.Label(row1, text="批量上限", width=10).pack(side=LEFT)
        ttk.Spinbox(row1, from_=1, to=50, textvariable=self.ai_batch_limit, width=8).pack(side=LEFT)
        ttk.Label(row1, text="修复模式", width=10).pack(side=LEFT, padx=(12, 0))
        ttk.Combobox(
            row1,
            textvariable=self.ai_mode,
            values=[label for _key, label in _AI_MODE_CHOICES],
            state="readonly",
            width=12,
        ).pack(side=LEFT)
        ttk.Checkbutton(
            row1,
            text="批量时仅处理待确认题",
            variable=self.ai_only_flagged,
            bootstyle="round-toggle",
        ).pack(side=LEFT, padx=(12, 0))

        row2 = ttk.Frame(frame)
        row2.pack(fill=X, pady=(0, 6))
        ttk.Label(
            row2,
            textvariable=self._ai_status_var,
            wraplength=680,
            justify=LEFT,
            bootstyle="secondary",
        ).pack(side=LEFT, padx=(12, 0))

        row3 = ttk.Frame(frame)
        row3.pack(fill=X, pady=(0, 4))
        self._pdf_ocr_diagnose_btn = ttk.Button(
            row3,
            text="扫描/OCR 诊断",
            command=self._run_pdf_ocr_diagnostics,
            bootstyle="info-outline",
            width=14,
            state=DISABLED,
        )
        self._pdf_ocr_diagnose_btn.pack(side=LEFT)
        self._pdf_ocr_repair_btn = ttk.Button(
            row3,
            text="OCR 自动修补",
            command=self._auto_repair_pdf_ocr,
            bootstyle="warning-outline",
            width=14,
            state=DISABLED,
        )
        self._pdf_ocr_repair_btn.pack(side=LEFT, padx=(8, 0))
        ttk.Label(
            frame,
            textvariable=self._pdf_ocr_status_var,
            wraplength=860,
            justify=LEFT,
            bootstyle="secondary",
        ).pack(anchor=W, padx=(12, 0))

    def _build_pdf_wizard(self, parent):
        shell = ttk.Frame(parent)
        shell.pack(fill=BOTH, expand=YES)

        rail = ttk.Frame(shell, padding=(0, 0, 12, 0))
        rail.pack(side=LEFT, fill=Y)
        main = ttk.Frame(shell)
        main.pack(side=LEFT, fill=BOTH, expand=YES)

        step_card = ttk.Labelframe(rail, text=" 工作流 ", bootstyle="secondary", padding=(10, 10))
        step_card.pack(fill=X, pady=(2, 10))
        self._pdf_step_buttons = []
        for index, (title, _hint) in enumerate(_PDF_WIZARD_STEPS):
            button = ttk.Button(
                step_card,
                text=f"{index + 1}. {title}",
                command=lambda i=index: self._request_pdf_wizard_step(i),
                width=18,
                bootstyle="secondary-outline",
            )
            button.pack(fill=X, pady=(0, 8))
            self._pdf_step_buttons.append(button)

        workspace_card = ttk.Labelframe(rail, text=" 当前工程 ", bootstyle="info", padding=(10, 10))
        workspace_card.pack(fill=X, pady=(0, 10))
        source_tile = self._build_metric_tile(workspace_card, "来源", self._pdf_workspace_source_var, bg=_UI_PRIMARY)
        source_tile.pack(fill=X, pady=(0, 8))
        scope_tile = self._build_metric_tile(workspace_card, "范围", self._pdf_workspace_scope_var, bg=_UI_INFO)
        scope_tile.pack(fill=X, pady=(0, 8))
        state_tile = self._build_metric_tile(workspace_card, "状态", self._pdf_workspace_state_var, bg=_UI_WARNING)
        state_tile.pack(fill=X)

        quick_card = ttk.Labelframe(rail, text=" 快捷入口 ", bootstyle="secondary", padding=(10, 10))
        quick_card.pack(fill=X)
        ttk.Button(
            quick_card,
            text="重新生成预览",
            command=self._preview_pdf_project,
            bootstyle="secondary-outline",
        ).pack(fill=X, pady=(0, 6))
        ttk.Button(
            quick_card,
            text="打开导出设置",
            command=self._open_ppt_settings_tab,
            bootstyle="info-outline",
        ).pack(fill=X, pady=(0, 6))
        self._pdf_handoff_btn = ttk.Button(
            quick_card,
            text="去 Word 生成 PPT",
            command=self._open_word_workspace_tab,
            bootstyle="info-outline",
        )
        self._pdf_handoff_btn.pack(fill=X)

        intro = ttk.Frame(main, padding=(12, 0, 12, 6))
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

        self._pdf_step_host = ttk.Frame(main)
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

        nav = ttk.Frame(main, padding=(12, 8, 12, 0))
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
        self._file_row(box, "工程 JSON", self.pdf_manifest_out, self._browse_pdf_manifest, "另存为...", pady=(6, 0))

        handoff_row = ttk.Frame(box)
        handoff_row.pack(fill=X, pady=(10, 0))
        ttk.Label(handoff_row, text="下一步", width=10).pack(side=LEFT)
        self._pdf_export_handoff_label = ttk.Label(
            handoff_row,
            text="需要做 PPT 时，请把 Word 交给“Word 生成 PPT”；解析后会进入同一套共享预览。",
            bootstyle="secondary",
        )
        self._pdf_export_handoff_label.pack(side=LEFT, fill=X, expand=YES)
        self._pdf_export_handoff_btn = ttk.Button(
            handoff_row,
            text="去 Word 生成 PPT",
            command=self._open_word_workspace_tab,
            bootstyle="info-outline",
            width=14,
        )
        self._pdf_export_handoff_btn.pack(side=RIGHT)

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
            text="导出 Word / 工程",
            command=self._export_pdf_word,
            bootstyle="info-outline",
            width=14,
        ).pack(side=RIGHT, padx=(4, 0))
        ttk.Button(
            action_row,
            text="导出 Word 并进入 PPT",
            command=self._export_pdf_word_and_open_ppt_flow,
            bootstyle="success",
            width=18,
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
        self._open_word_workspace_tab()
        notebook = getattr(self, "_word_settings_notebook", None)
        layout_tab = getattr(self, "_word_settings_layout_tab", None)
        if notebook is not None and layout_tab is not None:
            notebook.select(layout_tab)

    def _open_word_workspace_tab(self):
        notebook = getattr(self, "_workspace_notebook", None)
        word_tab = getattr(self, "_word_workspace_tab", None)
        if notebook is not None and word_tab is not None:
            notebook.select(word_tab)

    def _open_word_editor_workspace(self):
        if not self._word_project_matches_current_file():
            if not self._start_word_preview_flow(open_editor=False):
                return
        self._open_pdf_preview_workspace()
        self._word_flow_status_var.set("当前已进入逐题编辑工作台；修改完成后回到 Word 工作流直接生成 PPT。")
        self._word_results_hint_var.set("你正在逐题精修；改完后回到 Word 工作流直接生成 PPT。")

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
            next_text = "下一步：保存工程 / 返回生成" if self._pdf_project_context.get("source_kind", "") == "word" else "下一步：导出 Word / 继续 PPT"
            self._pdf_next_btn.configure(text=next_text, state=state, bootstyle="success")
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
        selected_subject_count = sum(1 for kind in _PDF_SUBJECT_ORDER if self._pdf_subject_vars[kind].get())
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
                f"默认 PPT：{self.output_path.get().strip() or docx_base + '.pptx'}\n"
                f"默认工程：{self.pdf_manifest_out.get().strip() or docx_base + '_工程.json'}"
            )
        elif pdf_file:
            import_summary = (
                f"当前试卷：{base_name}\n"
                f"整份科目：{document_subject_text}\n"
                f"默认题本：{self.pdf_word_out.get().strip() or os.path.splitext(pdf_file)[0] + '_真题.docx'}\n"
                f"默认工程：{self.pdf_manifest_out.get().strip() or os.path.splitext(pdf_file)[0] + '_工程.json'}"
            )
        elif source_kind == "manifest" and manifest_file:
            manifest_base = os.path.splitext(manifest_file)[0]
            import_summary = (
                f"当前工程：{os.path.basename(manifest_file)}\n"
                f"来源 PDF：{self.pdf_path.get().strip() or '未记录 / 不可用'}\n"
                f"默认题本：{self.pdf_word_out.get().strip() or manifest_base + '_真题.docx'}\n"
                f"默认工程：{self.pdf_manifest_out.get().strip() or manifest_base + '_工程.json'}"
            )
        if getattr(self, "_pdf_import_summary", None):
            self._pdf_import_summary.configure(text=import_summary)

        source_label = "当前试卷"
        compact_subject_text = subject_text
        compact_range_text = "全部题目" if not self.pdf_question_range.get().strip() else range_text
        if source_kind == "word" and docx_file:
            source_label = "当前 Word"
            project_subjects = [
                SUBJECT_DISPLAY_NAMES.get(section.kind, section.kind)
                for section in getattr(self.pdf_project, "sections", [])
                if section.kind in ALL_SUBJECT_KINDS
            ]
            subject_text = "、".join(project_subjects) if project_subjects else "未识别科目"
            range_text = "Word 全部题目"
            compact_subject_text = subject_text
            compact_range_text = "全部题目"
        elif selected_subject_count == len(_PDF_SUBJECT_ORDER):
            compact_subject_text = "全部科目"
        elif selected_subject_count >= 3:
            compact_subject_text = f"{selected_subject_count} 个科目"
        settings_summary = (
            f"{source_label}：{base_name}\n"
            f"整份科目：{document_subject_text}\n"
            f"处理科目：{subject_text}\n"
            f"题号范围：{range_text}\n"
            f"{preview_state}"
        )
        if getattr(self, "_pdf_settings_summary", None):
            self._pdf_settings_summary.configure(text=settings_summary)

        if preview_ready:
            asset_dir = self._pdf_project_context.get("asset_dir", "-")
            if source_kind == "word":
                export_summary = (
                    f"当前工程共 {self.pdf_project.question_count} 道题，素材目录：{asset_dir}\n"
                    "这一步更适合保存工程 JSON；PPT 请回到 Word 工作流直接生成。"
                )
            else:
                export_summary = (
                    f"当前工程共 {self.pdf_project.question_count} 道题，素材目录：{asset_dir}\n"
                    "这里会导出 Word / JSON；如果要做 PPT，请把导出的 Word 交给 Word 工作流继续解析。"
                )
        else:
            export_summary = "请先在上一步生成当前设置对应的预览工程。"
        if getattr(self, "_pdf_export_summary", None):
            self._pdf_export_summary.configure(text=export_summary)

        handoff_text = "去 Word 生成 PPT"
        handoff_hint = "需要做 PPT 时，请把 Word 交给“Word 生成 PPT”；解析后会进入同一套共享预览。"
        if source_kind == "word":
            handoff_text = "返回 Word 生成 PPT"
            handoff_hint = "当前工程来自 Word；改完文字、结构或版式后，回到 Word 工作流直接生成 PPT。"
        if getattr(self, "_pdf_handoff_btn", None):
            self._pdf_handoff_btn.configure(text=handoff_text)
        if getattr(self, "_pdf_export_handoff_btn", None):
            self._pdf_export_handoff_btn.configure(text=handoff_text)
        if getattr(self, "_pdf_export_handoff_label", None):
            self._pdf_export_handoff_label.configure(text=handoff_hint)

        if getattr(self, "_pdf_nav_status_label", None):
            self._pdf_nav_status_label.configure(text=preview_state)
        if getattr(self, "_pdf_workspace_source_var", None):
            self._pdf_workspace_source_var.set(base_name)
        if getattr(self, "_pdf_workspace_scope_var", None):
            self._pdf_workspace_scope_var.set(f"{compact_subject_text} · {compact_range_text}")
        if getattr(self, "_pdf_workspace_state_var", None):
            self._pdf_workspace_state_var.set(
                f"已就绪 · {self.pdf_project.question_count} 题"
                if preview_ready and self.pdf_project is not None
                else preview_state
            )

    def _build_pdf_preview(self, parent):
        frame = ttk.Frame(parent)
        frame.pack(fill=BOTH, expand=YES, pady=(0, 0))

        self._pdf_preview_summary = None

        split = ttk.Panedwindow(frame, orient=HORIZONTAL)
        split.pack(fill=BOTH, expand=YES)

        left = ttk.Labelframe(split, text=" 导航与定位 ", bootstyle="secondary", padding=(8, 8))
        stage = ttk.Frame(split)
        inspector = ttk.Frame(split)
        split.add(left, weight=2)
        split.add(stage, weight=5)
        split.add(inspector, weight=4)

        nav_banner = tk.Frame(left, bg=_UI_PRIMARY_DARK)
        nav_banner.pack(fill=X, pady=(0, 4))
        tk.Label(
            nav_banner,
            text="导航区",
            bg=_UI_PRIMARY_DARK,
            fg=_UI_TEXT_LIGHT,
            font=("", 11, "bold"),
            anchor="w",
        ).pack(fill=X, padx=10, pady=6)

        left_tabs = ttk.Notebook(left, bootstyle="info")
        left_tabs.pack(fill=BOTH, expand=YES)
        self._pdf_preview_left_tabs = left_tabs

        structure_tab = ttk.Frame(left_tabs)
        review_tab = ttk.Frame(left_tabs)
        slide_tab = ttk.Frame(left_tabs)
        left_tabs.add(structure_tab, text=" 结构树 ")
        left_tabs.add(review_tab, text=" 待确认 ")
        left_tabs.add(slide_tab, text=" PPT 页 ")
        self._pdf_preview_structure_tab = structure_tab
        self._pdf_preview_review_tab = review_tab
        self._pdf_preview_slide_tab = slide_tab

        cols = ("kind", "source", "count")
        self.pdf_tree = ttk.Treeview(
            structure_tab,
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
        sb = ttk.Scrollbar(structure_tab, orient=VERTICAL, command=self.pdf_tree.yview)
        self.pdf_tree.configure(yscrollcommand=sb.set)
        sb.pack(side=RIGHT, fill=Y)
        self.pdf_tree.bind("<<TreeviewSelect>>", self._on_pdf_preview_select)

        review_cols = ("severity", "score", "source", "subject", "issue")
        self._pdf_review_tree = ttk.Treeview(
            review_tab,
            columns=review_cols,
            show="headings",
            height=12,
            bootstyle="warning",
        )
        self._pdf_review_tree.heading("severity", text="级别")
        self._pdf_review_tree.heading("score", text="置信度")
        self._pdf_review_tree.heading("source", text="题号")
        self._pdf_review_tree.heading("subject", text="科目")
        self._pdf_review_tree.heading("issue", text="待确认原因")
        self._pdf_review_tree.column("severity", width=64, anchor=CENTER)
        self._pdf_review_tree.column("score", width=74, anchor=CENTER)
        self._pdf_review_tree.column("source", width=64, anchor=CENTER)
        self._pdf_review_tree.column("subject", width=92, anchor=CENTER)
        self._pdf_review_tree.column("issue", width=280)
        self._pdf_review_tree.pack(side=LEFT, fill=BOTH, expand=YES)
        review_sb = ttk.Scrollbar(review_tab, orient=VERTICAL, command=self._pdf_review_tree.yview)
        self._pdf_review_tree.configure(yscrollcommand=review_sb.set)
        review_sb.pack(side=RIGHT, fill=Y)
        self._pdf_review_tree.bind("<<TreeviewSelect>>", self._on_pdf_review_select)

        slide_cols = ("page", "source", "subject", "stem")
        self._pdf_slide_tree = ttk.Treeview(
            slide_tab,
            columns=slide_cols,
            show="headings",
            height=12,
            bootstyle="info",
        )
        self._pdf_slide_tree.heading("page", text="页码")
        self._pdf_slide_tree.heading("source", text="题号")
        self._pdf_slide_tree.heading("subject", text="科目")
        self._pdf_slide_tree.heading("stem", text="题干摘要")
        self._pdf_slide_tree.column("page", width=64, anchor=CENTER)
        self._pdf_slide_tree.column("source", width=64, anchor=CENTER)
        self._pdf_slide_tree.column("subject", width=92, anchor=CENTER)
        self._pdf_slide_tree.column("stem", width=280)
        self._pdf_slide_tree.pack(side=LEFT, fill=BOTH, expand=YES)
        slide_sb = ttk.Scrollbar(slide_tab, orient=VERTICAL, command=self._pdf_slide_tree.yview)
        self._pdf_slide_tree.configure(yscrollcommand=slide_sb.set)
        slide_sb.pack(side=RIGHT, fill=Y)
        self._pdf_slide_tree.bind("<<TreeviewSelect>>", self._on_pdf_slide_select)

        stage_header = tk.Frame(stage, bg=_UI_PRIMARY)
        stage_header.pack(fill=X, pady=(0, 4))
        tk.Label(
            stage_header,
            text="幻灯片工作台",
            bg=_UI_PRIMARY,
            fg=_UI_TEXT_LIGHT,
            font=("", 12, "bold"),
            anchor="w",
        ).pack(side=LEFT, padx=12, pady=7)

        stage_toolbar = ttk.Frame(stage)
        stage_toolbar.pack(fill=X, pady=(0, 4))
        slide_row = ttk.Frame(stage_toolbar)
        slide_row.pack(fill=X, pady=(0, 4))
        self._pdf_slide_status_label = ttk.Label(
            slide_row,
            textvariable=self._pdf_slide_status_var,
            bootstyle="secondary",
        )
        self._pdf_slide_status_label.pack(side=LEFT, fill=X, expand=YES)
        self._pdf_slide_prev_btn = ttk.Button(
            slide_row,
            text="上一页",
            command=lambda: self._step_pdf_slide(-1),
            bootstyle="secondary-outline",
            width=8,
            state=DISABLED,
        )
        self._pdf_slide_prev_btn.pack(side=LEFT, padx=(6, 4))
        self._pdf_slide_next_btn = ttk.Button(
            slide_row,
            text="下一页",
            command=lambda: self._step_pdf_slide(1),
            bootstyle="secondary-outline",
            width=8,
            state=DISABLED,
        )
        self._pdf_slide_next_btn.pack(side=LEFT)
        ttk.Label(
            slide_row,
            textvariable=self._pdf_review_status_var,
            bootstyle="warning",
        ).pack(side=LEFT, padx=(12, 0))
        ttk.Button(
            slide_row,
            text="下一个待确认",
            command=self._jump_to_next_pdf_review_item,
            bootstyle="warning-outline",
            width=12,
        ).pack(side=LEFT, padx=(8, 0))

        stage_split = ttk.Panedwindow(stage, orient=VERTICAL)
        stage_split.pack(fill=BOTH, expand=YES)

        preview_card = ttk.Labelframe(stage_split, text=" 所见即所得预览 ", bootstyle="primary", padding=(10, 8))
        stage_split.add(preview_card, weight=5)
        self._build_pdf_live_preview_panel(preview_card)

        stage_lower = ttk.Labelframe(stage_split, text=" 原始材料与结构详情 ", bootstyle="secondary", padding=(10, 8))
        stage_split.add(stage_lower, weight=3)
        lower_tabs = ttk.Notebook(stage_lower, bootstyle="info")
        lower_tabs.pack(fill=BOTH, expand=YES)
        material_tab = ttk.Frame(lower_tabs, padding=(0, 0, 0, 0))
        detail_tab = ttk.Frame(lower_tabs, padding=(0, 0, 0, 0))
        lower_tabs.add(material_tab, text=" 材料原貌 ")
        lower_tabs.add(detail_tab, text=" 结构详情 ")
        self._build_pdf_material_preview_panel(material_tab)

        detail_host = ttk.Frame(detail_tab)
        detail_host.pack(fill=BOTH, expand=YES)
        self.pdf_detail = tk.Text(detail_host, wrap="word", height=16)
        self.pdf_detail.pack(side=LEFT, fill=BOTH, expand=YES)
        detail_scroll = ttk.Scrollbar(detail_host, orient=VERTICAL, command=self.pdf_detail.yview)
        detail_scroll.pack(side=RIGHT, fill=Y)
        self.pdf_detail.configure(yscrollcommand=detail_scroll.set)
        self.pdf_detail.configure(state="disabled")

        inspector_header = tk.Frame(inspector, bg=_UI_WARNING)
        inspector_header.pack(fill=X, pady=(0, 4))
        tk.Label(
            inspector_header,
            text="检查器",
            bg=_UI_WARNING,
            fg=_UI_TEXT_LIGHT,
            font=("", 12, "bold"),
            anchor="w",
        ).pack(fill=X, padx=12, pady=7)

        action_box = ttk.Labelframe(inspector, text=" 工程动作 ", bootstyle="secondary", padding=(8, 6))
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
            text="导出 AI 质检报告",
            command=self._export_pdf_review_report,
            bootstyle="warning-outline",
            width=14,
        ).pack(side=LEFT, padx=4)
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
        ttk.Button(
            action_row,
            text="批量应用 AI 安全建议",
            command=self._apply_all_safe_ai_suggestions,
            bootstyle="warning-outline",
            width=18,
        ).pack(side=LEFT, padx=(8, 0))
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
        self._build_pdf_question_editor(inspector)

    def _build_pdf_live_preview_panel(self, parent):
        self._pdf_question_preview_canvas = tk.Canvas(
            parent,
            height=360,
            bg="#edf2f7",
            highlightthickness=0,
        )
        self._pdf_question_preview_canvas.pack(fill=BOTH, expand=YES)
        self._pdf_question_preview_canvas.bind(
            "<Configure>",
            lambda _event: self._render_pdf_question_editor_preview(),
        )
        self._pdf_question_preview_canvas.bind("<ButtonPress-1>", self._on_pdf_question_preview_press)
        self._pdf_question_preview_canvas.bind("<B1-Motion>", self._on_pdf_question_preview_drag)
        self._pdf_question_preview_canvas.bind("<ButtonRelease-1>", self._on_pdf_question_preview_release)
        self._pdf_question_preview_canvas.bind("<Motion>", self._on_pdf_question_preview_motion)
        for sequence in (
            "<Left>",
            "<Right>",
            "<Up>",
            "<Down>",
            "<Shift-Left>",
            "<Shift-Right>",
            "<Shift-Up>",
            "<Shift-Down>",
            "<Control-Left>",
            "<Control-Right>",
            "<Control-Up>",
            "<Control-Down>",
        ):
            self._pdf_question_preview_canvas.bind(sequence, self._on_pdf_question_preview_nudge)

    def _build_pdf_question_ai_panel(self, parent):
        ai_box = ttk.Labelframe(parent, text=" AI 修复建议 ", bootstyle="warning", padding=(8, 6))
        ai_box.pack(fill=X, pady=(0, 8))
        ttk.Label(
            ai_box,
            textvariable=self._pdf_ai_suggestion_var,
            wraplength=430,
            justify=LEFT,
            bootstyle="secondary",
        ).pack(anchor=W, pady=(0, 6))
        ttk.Label(
            ai_box,
            textvariable=self._pdf_ai_strategy_var,
            wraplength=430,
            justify=LEFT,
            bootstyle="info",
        ).pack(anchor=W, pady=(0, 6))
        ai_action_row = ttk.Frame(ai_box)
        ai_action_row.pack(fill=X)
        self._pdf_repair_current_ai_btn = ttk.Button(
            ai_action_row,
            text="AI 修当前题",
            command=self._repair_selected_question_with_ai,
            bootstyle="warning",
            width=12,
            state=DISABLED,
        )
        self._pdf_repair_current_ai_btn.pack(side=LEFT, padx=(0, 6))
        self._pdf_repair_batch_ai_btn = ttk.Button(
            ai_action_row,
            text="AI 批量修复",
            command=self._repair_flagged_questions_with_ai,
            bootstyle="warning-outline",
            width=12,
            state=DISABLED,
        )
        self._pdf_repair_batch_ai_btn.pack(side=LEFT, padx=(0, 6))
        self._pdf_apply_ai_suggestion_btn = ttk.Button(
            ai_action_row,
            text="应用当前篇题建议",
            command=self._apply_selected_ai_subject_suggestion,
            bootstyle="warning-outline",
            width=16,
            state=DISABLED,
        )
        self._pdf_apply_ai_suggestion_btn.pack(side=LEFT, padx=(0, 6))
        ttk.Button(
            ai_action_row,
            text="跳到下一个待确认",
            command=self._jump_to_next_pdf_review_item,
            bootstyle="secondary-outline",
            width=14,
        ).pack(side=LEFT)

    def _build_pdf_question_content_panel(self, parent):
        stem_box = ttk.Labelframe(parent, text=" 题干编辑 ", bootstyle="info", padding=(8, 6))
        stem_box.pack(fill=X, pady=(0, 8))
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
        self._pdf_option_editor_host = ttk.Frame(option_box)
        self._pdf_option_editor_host.pack(fill=X, expand=YES)

    def _build_pdf_question_layout_panel(self, parent):
        layout_box = ttk.Labelframe(parent, text=" 单题选项布局 ", bootstyle="secondary", padding=(8, 6))
        layout_box.pack(fill=X, pady=(0, 8))
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

        quick_box = ttk.Labelframe(parent, text=" 全局 PPT 快设 ", bootstyle="info", padding=(8, 6))
        quick_box.pack(fill=X, pady=(0, 8))
        quick_row = ttk.Frame(quick_box)
        quick_row.pack(fill=X)
        ttk.Label(quick_row, text="全局选项布局", bootstyle="secondary").pack(side=LEFT)
        ttk.Combobox(
            quick_row,
            textvariable=self.option_layout,
            values=["grid", "list", "one_row"],
            state="readonly",
            width=12,
        ).pack(side=LEFT, padx=(6, 10))
        ttk.Label(quick_row, text="题干字号", bootstyle="secondary").pack(side=LEFT)
        ttk.Spinbox(quick_row, from_=10, to=48, textvariable=self.font_size_stem, width=6).pack(side=LEFT, padx=(6, 10))
        ttk.Label(quick_row, text="选项字号", bootstyle="secondary").pack(side=LEFT)
        ttk.Spinbox(quick_row, from_=8, to=40, textvariable=self.font_size_option, width=6).pack(side=LEFT, padx=(6, 10))
        ttk.Button(
            quick_row,
            text="打开导出设置",
            command=self._open_ppt_settings_tab,
            bootstyle="info-outline",
            width=12,
        ).pack(side=RIGHT)

        preview_box = ttk.Labelframe(parent, text=" 版式工具 ", bootstyle="primary", padding=(8, 6))
        preview_box.pack(fill=BOTH, expand=YES)
        preview_toolbar = ttk.Frame(preview_box)
        preview_toolbar.pack(fill=X, pady=(0, 6))
        ttk.Label(
            preview_toolbar,
            textvariable=self._pdf_layout_editor_status_var,
            bootstyle="secondary",
        ).pack(side=LEFT, fill=X, expand=YES)
        self._pdf_reset_question_layout_btn = ttk.Button(
            preview_toolbar,
            text="重置当前题版式",
            command=self._reset_pdf_question_ppt_layout,
            bootstyle="secondary-outline",
            width=14,
            state=DISABLED,
        )
        self._pdf_reset_question_layout_btn.pack(side=LEFT)
        preview_action_row = ttk.Frame(preview_box)
        preview_action_row.pack(fill=X, pady=(0, 6))
        ttk.Label(
            preview_action_row,
            text="快捷动作：",
            bootstyle="secondary",
        ).pack(side=LEFT, padx=(0, 6))
        for text, action in (
            ("贴左", "left"),
            ("居中", "hcenter"),
            ("贴右", "right"),
            ("贴顶", "top"),
            ("中线", "vcenter"),
            ("贴底", "bottom"),
            ("铺宽", "fill_width"),
            ("铺高", "fill_height"),
        ):
            button = ttk.Button(
                preview_action_row,
                text=text,
                command=lambda value=action: self._align_selected_pdf_preview_block(value),
                bootstyle="light-outline",
                width=7,
                state=DISABLED,
            )
            button.pack(side=LEFT, padx=(0, 4))
            self._pdf_preview_action_buttons.append(button)
        ttk.Label(
            preview_action_row,
            text="方向键微调，Shift 加速，Ctrl+方向键缩放",
            bootstyle="secondary",
        ).pack(side=LEFT, padx=(8, 0))
        preview_metric_row = ttk.Frame(preview_box)
        preview_metric_row.pack(fill=X, pady=(0, 6))
        ttk.Label(
            preview_metric_row,
            text="精确尺寸（英寸）：",
            bootstyle="secondary",
        ).pack(side=LEFT, padx=(0, 6))
        for label, variable, upper in (
            ("X", self._pdf_layout_x_var, _PPT_SLIDE_WIDTH_IN),
            ("Y", self._pdf_layout_y_var, _PPT_SLIDE_HEIGHT_IN),
            ("宽", self._pdf_layout_w_var, _PPT_SLIDE_WIDTH_IN),
            ("高", self._pdf_layout_h_var, _PPT_SLIDE_HEIGHT_IN),
        ):
            ttk.Label(preview_metric_row, text=label, bootstyle="secondary").pack(side=LEFT)
            field = ttk.Spinbox(
                preview_metric_row,
                from_=0,
                to=upper,
                increment=_PPT_LAYOUT_FIELD_STEP_IN,
                textvariable=variable,
                width=7,
                state=DISABLED,
            )
            field.pack(side=LEFT, padx=(2, 8))
            field.bind("<Return>", self._apply_pdf_preview_numeric_layout)
            field.bind("<FocusOut>", self._apply_pdf_preview_numeric_layout)
            self._pdf_layout_value_inputs.append(field)
        apply_btn = ttk.Button(
            preview_metric_row,
            text="应用",
            command=self._apply_pdf_preview_numeric_layout,
            bootstyle="info-outline",
            width=7,
            state=DISABLED,
        )
        apply_btn.pack(side=LEFT, padx=(0, 6))
        self._pdf_layout_field_buttons.append(apply_btn)
        ttk.Label(
            preview_metric_row,
            text="直接输入位置和尺寸，回车或失焦会自动生效。",
            bootstyle="secondary",
        ).pack(side=LEFT, padx=(4, 0))

    def _build_pdf_question_editor(self, parent):
        ttk.Label(
            parent,
            textvariable=self._pdf_question_editor_message,
            wraplength=440,
            justify=LEFT,
            bootstyle="secondary",
        ).pack(anchor=W, pady=(0, 8))

        notebook = ttk.Notebook(parent, bootstyle="info")
        notebook.pack(fill=BOTH, expand=YES)

        content_tab = ttk.Frame(notebook, padding=(0, 0, 0, 0))
        layout_tab = ttk.Frame(notebook, padding=(0, 0, 0, 0))
        ai_tab = ttk.Frame(notebook, padding=(0, 0, 0, 0))
        notebook.add(content_tab, text=" 内容 ")
        notebook.add(layout_tab, text=" 版式 ")
        notebook.add(ai_tab, text=" AI / 审核 ")

        self._build_pdf_question_content_panel(content_tab)
        self._build_pdf_question_layout_panel(layout_tab)
        self._build_pdf_question_ai_panel(ai_tab)
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

    def _set_template_status(self, text: str, bootstyle: str = "secondary"):
        self._template_status_var.set(text)
        if getattr(self, "_tpl_status", None):
            self._tpl_status.configure(bootstyle=bootstyle)

    def _show_template_config_notebook(self):
        if getattr(self, "_config_notebook", None) and self._config_notebook.winfo_manager() != "pack":
            self._config_notebook.pack(fill=X, padx=8, pady=(4, 8))

    def _refresh_template_mode_ui(self):
        on = self.use_template.get()
        if not on:
            self._template_mode = "off"
            self._tpl_overlay_label.pack_forget()
            self._show_template_config_notebook()
            self._set_template_status("未启用模板。", "secondary")
            return

        if self._template_mode == "full":
            self._config_notebook.pack_forget()
            self._tpl_overlay_label.configure(text="已识别完整模板版式：导出时将直接沿用模板布局。")
            self._tpl_overlay_label.pack(fill=X, padx=8, pady=10)
            return

        self._tpl_overlay_label.pack_forget()
        self._show_template_config_notebook()

    def _inspect_selected_template(self, *, show_error: bool) -> bool:
        if not self.use_template.get():
            self._template_style_preview = None
            self._template_mode = "off"
            self._refresh_template_mode_ui()
            return False

        template = self.template_path.get().strip()
        if not template:
            self._template_style_preview = None
            self._template_mode = "pending"
            self._set_template_status("请选择模板文件。", "secondary")
            self._refresh_template_mode_ui()
            return False
        if not os.path.exists(template):
            self._template_style_preview = None
            self._template_mode = "error"
            self._set_template_status("模板文件不存在。", "danger")
            self._refresh_template_mode_ui()
            if show_error:
                messagebox.showerror("错误", f"模板文件不存在：{template}")
            return False

        try:
            prs = self._template_manager.load_template(template)
            style = extract_best_style_from_presentation(prs)
        except Exception as exc:
            self._template_style_preview = None
            self._template_mode = "error"
            self._set_template_status(f"模板读取失败：{exc}", "danger")
            self._refresh_template_mode_ui()
            if show_error:
                messagebox.showerror("错误", f"读取模板失败：{exc}")
            return False

        self._template_style_preview = style
        slide_text = f"第 {style.source_slide_index + 1} 页"
        summary = describe_template_style(style)
        if template_style_has_full_layout(style):
            self._template_mode = "full"
            self._set_template_status(
                f"已识别 {slide_text}：{summary}。导出时将直接沿用模板版式。",
                "success",
            )
        else:
            self._template_mode = "partial"
            self._set_template_status(
                f"已读取 {slide_text}：{summary}。会复用模板样式，布局继续使用当前设置。",
                "warning",
            )
        self._refresh_template_mode_ui()
        return True

    def _on_template_entry_changed(self, _event=None):
        if self.use_template.get():
            self._inspect_selected_template(show_error=False)

    def _toggle_template(self):
        on = self.use_template.get()
        st = NORMAL if on else DISABLED
        self.template_entry.configure(state=st)
        self.template_btn.configure(state=st)
        if on:
            self._inspect_selected_template(show_error=False)
        else:
            self._template_style_preview = None
            self._template_mode = "off"
            self._refresh_template_mode_ui()

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
            self._open_word_workspace_tab()
            self._set_status(f"已选择 Word：{os.path.basename(path)}")

    def _browse_output(self):
        path = filedialog.asksaveasfilename(title="保存 PPT",
                                            defaultextension=".pptx",
                                            filetypes=[("PowerPoint", "*.pptx")])
        if path:
            self.output_path.set(path)
            self._set_status(f"输出位置已更新：{os.path.basename(path)}")

    def _browse_template(self):
        path = filedialog.askopenfilename(title="选择 PPT 模板",
                                          filetypes=[("PowerPoint", "*.pptx"), ("All", "*.*")])
        if path:
            self.template_path.set(path)
            self._inspect_selected_template(show_error=True)

    def _browse_pdf(self):
        path = filedialog.askopenfilename(
            title="选择 PDF 试卷",
            filetypes=[("PDF", "*.pdf"), ("All", "*.*")],
        )
        if path:
            self.pdf_path.set(path)
            if not self.pdf_word_out.get().strip():
                self.pdf_word_out.set(os.path.splitext(path)[0] + "_真题.docx")
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

    def _export_pdf_word_and_open_ppt_flow(self):
        self._run_pdf_project(export_word=True, export_ppt=False, open_word_preview=True)

    def _run_pdf_project(self, *, export_word: bool, export_ppt: bool, open_word_preview: bool = False):
        from workflows.project_flow import build_pdf_project, export_project_outputs

        pdf_file = self.pdf_path.get().strip()
        subject_spec = self._current_pdf_subject_spec()
        source_kind = self._pdf_project_context.get("source_kind", "")
        is_cached_source_project = source_kind in {"word", "manifest"} and self.pdf_project is not None
        document_subject_hint = self._document_subject_key(self.pdf_document_subject.get())
        if not subject_spec and not is_cached_source_project:
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
        if is_cached_source_project:
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
        if not export_ppt:
            template = None
        if export_ppt and template and not os.path.exists(template):
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
                self.root.after(
                    0,
                    lambda: self._on_pdf_project_done(
                        project,
                        outputs,
                        open_word_preview=open_word_preview,
                    ),
                )
            except Exception as exc:
                self.root.after(0, lambda e=exc: self._on_pdf_project_error(str(e)))

        threading.Thread(target=work, daemon=True).start()

    def _on_pdf_project_done(self, project, outputs, *, open_word_preview: bool = False):
        self.pdf_project = project
        self._pdf_project_dirty = False
        source_kind = self._pdf_project_context.get("source_kind", "")
        manifest_path = self._pdf_project_context.get("manifest_path", "")
        is_cached_source_project = source_kind in {"word", "manifest"}
        self._pdf_project_context = {
            "pdf_path": (
                self.pdf_path.get().strip()
                if not is_cached_source_project
                else self._pdf_project_context.get("pdf_path", "")
            ),
            "docx_path": (
                self.word_path.get().strip()
                if source_kind == "word"
                else self._pdf_project_context.get("docx_path", "")
            ),
            "subject_spec": (
                self._current_pdf_subject_spec()
                if not is_cached_source_project
                else self._pdf_project_context.get("subject_spec", "all")
            ),
            "range_spec": (
                self.pdf_question_range.get().strip()
                if not is_cached_source_project
                else self._pdf_project_context.get("range_spec", "")
            ),
            "asset_dir": outputs.asset_dir,
            "source_kind": source_kind or "pdf",
            "manifest_path": self.pdf_manifest_out.get().strip() or manifest_path,
            "document_subject_hint": (
                self._document_subject_key(self.word_document_subject.get())
                if source_kind == "word"
                else (
                    self._pdf_project_context.get("document_subject_hint", "auto")
                    if source_kind == "manifest"
                    else self._document_subject_key(self.pdf_document_subject.get())
                )
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
        if open_word_preview and outputs.docx_path:
            self._set_status("PDF 已整理完成，正在把导出的 Word 送入 PPT 工作流…")
            handed_off = self._load_docx_into_word_workflow(
                outputs.docx_path,
                document_subject_hint="auto",
                auto_preview=True,
                skip_confirm=True,
            )
            if handed_off:
                return
            self._set_status("Word 已导出，可在“Word 生成 PPT”中继续解析。")
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
        self._refresh_word_flow_ui()

    def _refresh_word_flow_ui(self):
        has_word = bool(self.word_path.get().strip())
        has_output = bool(self.output_path.get().strip())
        parsed_current = self._word_project_matches_current_file() and bool(self.questions)

        if getattr(self, "_word_parse_btn", None):
            self._word_parse_btn.configure(state=(NORMAL if has_word else DISABLED))
        if getattr(self, "_word_convert_btn", None):
            self._word_convert_btn.configure(state=(NORMAL if has_word else DISABLED))
        if getattr(self, "_word_generate_btn", None):
            self._word_generate_btn.configure(state=(NORMAL if parsed_current and has_output else DISABLED))
        if getattr(self, "_word_editor_btn", None):
            self._word_editor_btn.configure(state=(NORMAL if parsed_current else DISABLED))

        if parsed_current:
            count = len(self.questions)
            self._word_flow_status_var.set(
                f"已解析 {count} 道题。常规下一步直接生成 PPT；只有版式要逐题调整时再进入编辑工作台。"
            )
            self._word_results_hint_var.set(
                f"当前显示 {count} 道题；快速检查无误后可直接生成 PPT。"
            )
            return

        if has_word:
            if has_output:
                self._word_flow_status_var.set("先点“解析并检查”；正常情况下，检查通过后下一步就是“生成 PPT”。")
            else:
                self._word_flow_status_var.set("先点“解析并检查”；如果赶时间，也可以直接用“一键生成 PPT”。")
            self._word_results_hint_var.set("当前还没有解析结果；完成解析后，这里会显示本次识别到的题目。")
            return

        self._word_flow_status_var.set("先选择 Word 题本。常规流程：选择文件 → 解析并检查 → 生成 PPT。")
        self._word_results_hint_var.set("还没有解析结果。先在左侧选择题本。")

    def _clear_pdf_preview(self):
        self.pdf_project = None
        self._pdf_project_dirty = False
        self._pdf_wizard_pending_step = None
        self._pdf_project_context = {}
        self._pdf_preview_payloads.clear()
        self._pdf_slide_payload_ids = []
        self._pdf_slide_payload_to_item_id = {}
        self._pdf_slide_item_to_payload_id = {}
        self._pdf_slide_payload_to_number = {}
        self._pdf_review_payload_ids = []
        self._pdf_review_payload_to_item_id = {}
        self._pdf_review_item_to_payload_id = {}
        self._reset_pdf_material_preview_session()
        self._clear_pdf_question_editor()
        if getattr(self, "pdf_tree", None):
            for item in self.pdf_tree.get_children():
                self.pdf_tree.delete(item)
        if getattr(self, "_pdf_review_tree", None):
            for item in self._pdf_review_tree.get_children():
                self._pdf_review_tree.delete(item)
        if getattr(self, "_pdf_slide_tree", None):
            for item in self._pdf_slide_tree.get_children():
                self._pdf_slide_tree.delete(item)
        self._refresh_pdf_slide_status("")
        self._refresh_pdf_review_status("")
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

    def _selected_pdf_slide_payload_id(self) -> str:
        slide_tree = getattr(self, "_pdf_slide_tree", None)
        if slide_tree is None:
            return ""
        selected = slide_tree.selection()
        if not selected:
            return ""
        return self._pdf_slide_item_to_payload_id.get(selected[0], "")

    def _selected_pdf_review_payload_id(self) -> str:
        review_tree = getattr(self, "_pdf_review_tree", None)
        if review_tree is None:
            return ""
        selected = review_tree.selection()
        if not selected:
            return ""
        return self._pdf_review_item_to_payload_id.get(selected[0], "")

    def _focus_pdf_left_tab(self, tab_name: str | None):
        left_tabs = getattr(self, "_pdf_preview_left_tabs", None)
        if left_tabs is None or not tab_name:
            return
        target_tab = {
            "structure": getattr(self, "_pdf_preview_structure_tab", None),
            "review": getattr(self, "_pdf_preview_review_tab", None),
            "slide": getattr(self, "_pdf_preview_slide_tab", None),
        }.get(tab_name)
        if target_tab is not None:
            left_tabs.select(target_tab)

    def _release_pdf_preview_sync_selection(self):
        self._pdf_preview_syncing_selection = False

    def _release_pdf_review_sync_selection(self):
        self._pdf_review_syncing_selection = False

    def _release_pdf_slide_sync_selection(self):
        self._pdf_slide_syncing_selection = False

    def _subject_hint_percent(self, value: float | None) -> int:
        try:
            numeric = float(value or 0.0)
        except (TypeError, ValueError):
            numeric = 0.0
        numeric = max(0.0, min(numeric, 1.0))
        return int(round(numeric * 100))

    def _apply_pdf_preview_selection(self, item_id: str, *, focus_left_tab: str | None = None):
        if not item_id:
            self._clear_pdf_question_editor()
            self._show_pdf_material_preview_for_payload({})
            self._sync_selected_section_subject({})
            self._sync_pdf_slide_selection("")
            self._sync_pdf_review_selection("")
            self._focus_pdf_left_tab(focus_left_tab)
            return

        payload = self._pdf_preview_payloads.get(item_id, {})
        self._set_pdf_detail(payload.get("text", "暂无内容"))
        self._sync_selected_section_subject(payload)
        self._populate_pdf_question_editor(payload)
        self._show_pdf_material_preview_for_payload(payload)
        self._sync_pdf_slide_selection(item_id)
        self._sync_pdf_review_selection(item_id)
        self._focus_pdf_left_tab(focus_left_tab)

    def _select_pdf_preview_item(self, item_id: str, *, focus_left_tab: str | None = None):
        if not item_id or not getattr(self, "pdf_tree", None):
            return
        current_selected = self._selected_pdf_item_id()
        if current_selected != item_id:
            try:
                self._pdf_preview_syncing_selection = True
                self.pdf_tree.selection_set(item_id)
                self.pdf_tree.focus(item_id)
                self.pdf_tree.see(item_id)
            finally:
                self.root.after_idle(self._release_pdf_preview_sync_selection)
        self._apply_pdf_preview_selection(item_id, focus_left_tab=focus_left_tab)

    def _find_preview_item_for_question(self, question) -> str:
        if question is None:
            return ""
        for item_id, payload in self._pdf_preview_payloads.items():
            if payload.get("kind") == "question" and payload.get("question") is question:
                return item_id
        return ""

    def _refresh_pdf_slide_status(self, payload_id: str = ""):
        total = len(self._pdf_slide_payload_ids)
        if total <= 0:
            self._pdf_slide_status_var.set("当前工程暂无可导出的 PPT 页。")
            if getattr(self, "_pdf_slide_prev_btn", None):
                self._pdf_slide_prev_btn.configure(state=DISABLED)
            if getattr(self, "_pdf_slide_next_btn", None):
                self._pdf_slide_next_btn.configure(state=DISABLED)
            return

        slide_number = self._pdf_slide_payload_to_number.get(payload_id, 0)
        if slide_number:
            payload = self._pdf_preview_payloads.get(payload_id, {})
            question = payload.get("question")
            source_number = getattr(question, "source_number", "") or "-"
            subject_label = self._document_subject_label(payload.get("section_kind") or "unknown")
            self._pdf_slide_status_var.set(
                f"当前 PPT 第 {slide_number}/{total} 页 · 原题号 {source_number} · {subject_label}"
            )
        else:
            self._pdf_slide_status_var.set(f"当前工程共 {total} 页 PPT；选择左侧某一页可逐页预览并实时编辑。")

        if getattr(self, "_pdf_slide_prev_btn", None):
            self._pdf_slide_prev_btn.configure(state=NORMAL if slide_number > 1 else DISABLED)
        if getattr(self, "_pdf_slide_next_btn", None):
            self._pdf_slide_next_btn.configure(state=NORMAL if slide_number and slide_number < total else DISABLED)

    def _refresh_pdf_review_status(self, payload_id: str = ""):
        total = len(self._pdf_review_payload_ids)
        if total <= 0:
            self._pdf_review_status_var.set("AI 质检未发现明显异常。")
            return

        if payload_id:
            payload = self._pdf_preview_payloads.get(payload_id, {})
            question = payload.get("question")
            if question is not None:
                score = int(round((getattr(question, "review_confidence", 1.0) or 1.0) * 100))
                self._pdf_review_status_var.set(
                    f"待确认 {total} 题 · 当前原题号 {question.source_number or '-'} · 置信度 {score}%"
                )
                return
        self._pdf_review_status_var.set(f"AI 质检标出 {total} 道待确认题，建议优先处理。")

    def _sync_pdf_slide_selection(self, payload_id: str):
        slide_tree = getattr(self, "_pdf_slide_tree", None)
        if slide_tree is None:
            return
        slide_item_id = self._pdf_slide_payload_to_item_id.get(payload_id, "")
        current_selected = slide_tree.selection()
        current_item_id = current_selected[0] if current_selected else ""
        if current_item_id != slide_item_id:
            try:
                self._pdf_slide_syncing_selection = True
                if slide_item_id:
                    slide_tree.selection_set(slide_item_id)
                    slide_tree.focus(slide_item_id)
                    slide_tree.see(slide_item_id)
                else:
                    slide_tree.selection_remove(slide_tree.selection())
            finally:
                self.root.after_idle(self._release_pdf_slide_sync_selection)
        self._refresh_pdf_slide_status(payload_id)

    def _sync_pdf_review_selection(self, payload_id: str):
        review_tree = getattr(self, "_pdf_review_tree", None)
        if review_tree is None:
            return
        review_item_id = self._pdf_review_payload_to_item_id.get(payload_id, "")
        current_selected = review_tree.selection()
        current_item_id = current_selected[0] if current_selected else ""
        if current_item_id != review_item_id:
            try:
                self._pdf_review_syncing_selection = True
                if review_item_id:
                    review_tree.selection_set(review_item_id)
                    review_tree.focus(review_item_id)
                    review_tree.see(review_item_id)
                else:
                    review_tree.selection_remove(review_tree.selection())
            finally:
                self.root.after_idle(self._release_pdf_review_sync_selection)
        self._refresh_pdf_review_status(payload_id)

    def _jump_to_next_pdf_review_item(self):
        if not self._pdf_review_payload_ids:
            return
        current_payload = self._selected_pdf_item_id() or self._selected_pdf_review_payload_id()
        if current_payload not in self._pdf_review_payload_ids:
            target_payload = self._pdf_review_payload_ids[0]
        else:
            index = self._pdf_review_payload_ids.index(current_payload)
            target_payload = self._pdf_review_payload_ids[(index + 1) % len(self._pdf_review_payload_ids)]
        self._select_pdf_preview_item(target_payload, focus_left_tab="review")

    def _review_severity_label(self, question) -> str:
        severity = question_max_severity(question)
        return {
            "error": "高风险",
            "warning": "注意",
            "info": "提示",
        }.get(severity, "稳定")

    def _step_pdf_slide(self, delta: int):
        if not self._pdf_slide_payload_ids:
            return
        current_payload = self._selected_pdf_item_id()
        if current_payload not in self._pdf_slide_payload_to_number:
            current_payload = self._selected_pdf_slide_payload_id()
        if current_payload not in self._pdf_slide_payload_to_number:
            target_payload = self._pdf_slide_payload_ids[0]
        else:
            index = self._pdf_slide_payload_to_number[current_payload] - 1
            target_index = max(0, min(len(self._pdf_slide_payload_ids) - 1, index + delta))
            target_payload = self._pdf_slide_payload_ids[target_index]
        self._select_pdf_preview_item(target_payload, focus_left_tab="slide")

    def _question_tree_label(self, question) -> str:
        compact = " ".join((question.stem or "").split())
        stem_short = compact[:38]
        if len(compact) > 38:
            stem_short += "..."
        label = stem_short or "未命名题目"
        if is_flagged_question(question):
            return f"[待确认] {label}"
        return label

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

    def _reanalyze_pdf_project(self):
        if self.pdf_project is None:
            return None
        return annotate_project_quality(self.pdf_project)

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
                lambda event, letter=option.letter: self._on_pdf_option_text_change(letter, event),
            )
            editor.bind(
                "<FocusOut>",
                lambda event, letter=option.letter: self._on_pdf_option_text_change(letter, event),
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
        slide_number = payload.get("slide_number")
        if slide_number:
            parts.insert(0, f"PPT 第 {slide_number} 页")
        if material is not None:
            parts.append(material.header or material.material_id)
        layout_source = "单题覆盖" if question.option_layout else "跟随全局"
        layout_label = self._option_layout_label(question.option_layout or self._effective_question_option_layout(question))
        message = " · ".join(parts)
        message += f"\n当前布局：{layout_source}（{layout_label}）"
        if getattr(question, "inferred_subtype", ""):
            message += f" · 子题型 {question.inferred_subtype}"
        if question.stem_assets:
            message += f" · 题干图片 {len(question.stem_assets)} 张"
        image_option_count = sum(1 for option in question.options if option.image_path)
        if image_option_count:
            message += f" · 图片选项 {image_option_count} 个"
        if is_flagged_question(question):
            message += (
                f"\nAI 质检：置信度 {int(round((question.review_confidence or 1.0) * 100))}%"
                f" · {question_review_summary(question)}"
            )
            hint = self._display_subject_hint(payload)
            if hint is not None:
                label, score, _reason = hint
                message += f" · 更像 {label}（{score}%）"
        else:
            message += "\nAI 质检：结构稳定"
        self._pdf_question_editor_message.set(message)

    def _display_subject_hint(self, payload: dict) -> tuple[str, int, str] | None:
        question = payload.get("question")
        section = payload.get("section")
        material = payload.get("material")
        if question is None:
            return None

        suggested_kind = getattr(question, "suggested_subject", None)
        suggested_confidence = getattr(question, "suggested_subject_confidence", None)
        suggested_reason = (getattr(question, "suggested_subject_reason", "") or "").strip()
        if suggested_kind:
            label = self._document_subject_label(suggested_kind)
            score = self._subject_hint_percent(suggested_confidence or question.review_confidence or 0.0)
            return label, score, suggested_reason or f"当前这道题更像 {label}。"

        diagnostics = infer_subject_diagnostics(
            stem=getattr(question, "stem", "") or "",
            options=[option.text or "" for option in getattr(question, "options", [])],
            material_text=(getattr(material, "body", "") or "") if material is not None else "",
            image_count=len(getattr(question, "stem_assets", []) or [])
            + sum(1 for option in getattr(question, "options", []) if getattr(option, "image_path", None)),
            material_header=(getattr(material, "header", "") or "") if material is not None else "",
            allow_data=True,
        )
        inferred_kind = diagnostics.kind
        inferred_confidence = diagnostics.confidence
        current_kind = getattr(section, "kind", "unknown")
        has_subject_issue = any(
            issue.code in {"subject_mismatch", "subject_suggestion", "unknown_subject"}
            for issue in (getattr(question, "review_issues", None) or [])
        )
        min_confidence = 0.38 if has_subject_issue else 0.58
        if inferred_kind in {"unknown", current_kind} or inferred_confidence < min_confidence:
            return None
        label = self._document_subject_label(inferred_kind)
        if diagnostics.subtype:
            label = f"{label} / {diagnostics.subtype}"
        score = self._subject_hint_percent(inferred_confidence)
        reason_prefix = "本地低置信度判断" if inferred_confidence < 0.58 else "本地判断"
        signal_text = ""
        if diagnostics.matched_signals:
            signal_text = "命中线索：" + "、".join(diagnostics.matched_signals[:3]) + "。"
        return label, score, f"{signal_text}{reason_prefix}这道题更像 {label}，建议人工确认。"

    def _refresh_pdf_ai_suggestion(self, payload: dict | None = None):
        payload = payload or self._selected_pdf_payload()
        mode = self._current_ai_mode()
        if mode == "policy":
            self._ai_status_var.set("本地 AI 修复当前处于策略增强模式：会叠加 learned repair policy 的排序、轨迹与提示。")
        else:
            self._ai_status_var.set("本地 AI 修复当前处于规则优先模式：会优先做安全规则修复。")
        self._refresh_pdf_ai_strategy(payload)
        button = getattr(self, "_pdf_apply_ai_suggestion_btn", None)
        repair_current_btn = getattr(self, "_pdf_repair_current_ai_btn", None)
        repair_batch_btn = getattr(self, "_pdf_repair_batch_ai_btn", None)
        can_batch = (
            self.pdf_project is not None
            and not self._ai_repair_busy
            and not self._ocr_tool_busy
            and (
                bool(self._pdf_review_payload_ids)
                if self.ai_only_flagged.get()
                else getattr(self.pdf_project, "question_count", 0) > 0
            )
        )
        if payload.get("kind") != "question":
            self._pdf_ai_suggestion_var.set("选中一道题后，这里会显示 AI 修复建议。")
            if button is not None:
                button.configure(state=DISABLED)
            if repair_current_btn is not None:
                repair_current_btn.configure(state=DISABLED)
            if repair_batch_btn is not None:
                repair_batch_btn.configure(state=NORMAL if can_batch else DISABLED)
            return

        question = payload.get("question")
        section = payload.get("section")
        if question is None or section is None:
            self._pdf_ai_suggestion_var.set("选中一道题后，这里会显示 AI 修复建议。")
            if button is not None:
                button.configure(state=DISABLED)
            if repair_current_btn is not None:
                repair_current_btn.configure(state=DISABLED)
            if repair_batch_btn is not None:
                repair_batch_btn.configure(state=NORMAL if can_batch else DISABLED)
            return

        target_kind, section_reason = section_subject_suggestion(section)
        hint = self._display_subject_hint(payload)
        if target_kind:
            target_label = self._document_subject_label(target_kind)
            self._pdf_ai_suggestion_var.set(
                f"整段建议：这组题更像 {target_label}。\n{section_reason}"
            )
            if button is not None:
                button.configure(state=NORMAL)
        elif hint is not None:
            label, score, reason = hint
            self._pdf_ai_suggestion_var.set(
                f"单题建议：这道题更像 {label}（置信度 {score}%）。\n"
                f"{reason or '当前只有单题级建议，暂不自动改整段。'}"
            )
        elif getattr(question, "review_issues", None):
            subject_issue = next(
                (
                    issue
                    for issue in question.review_issues
                    if issue.code in {"subject_mismatch", "subject_suggestion", "unknown_subject"} and issue.detail
                ),
                None,
            )
            if subject_issue is not None:
                self._pdf_ai_suggestion_var.set(
                    f"单题提醒：{subject_issue.title}\n{subject_issue.detail}"
                )
            else:
                self._pdf_ai_suggestion_var.set(
                    "AI 目前只给出质检提醒，没有形成可安全应用的整段建议。\n"
                    f"当前主要问题：{question_review_summary(question)}"
                )
        else:
            self._pdf_ai_suggestion_var.set("当前题目结构稳定，暂时没有 AI 修复建议。")
        if button is not None:
            button.configure(state=NORMAL if target_kind and not self._ai_repair_busy and not self._ocr_tool_busy else DISABLED)
        if repair_current_btn is not None:
            repair_current_btn.configure(
                state=NORMAL if not self._ai_repair_busy and not self._ocr_tool_busy else DISABLED
            )
        if repair_batch_btn is not None:
            repair_batch_btn.configure(state=NORMAL if can_batch else DISABLED)

    def _set_pdf_question_editor_state(self, enabled: bool):
        state = NORMAL if enabled else DISABLED
        if getattr(self, "_pdf_question_stem_editor", None):
            self._pdf_question_stem_editor.configure(state=state)
        for button in getattr(self, "_pdf_question_layout_buttons", []):
            button.configure(state=state)
        if getattr(self, "_pdf_reset_question_layout_btn", None):
            self._pdf_reset_question_layout_btn.configure(state=state)
        for button in getattr(self, "_pdf_preview_action_buttons", []):
            button.configure(state=state)
        for field in getattr(self, "_pdf_layout_value_inputs", []):
            field.configure(state=state)
        for button in getattr(self, "_pdf_layout_field_buttons", []):
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
        if getattr(self, "_pdf_apply_ai_suggestion_btn", None) is not None and not enabled:
            self._pdf_apply_ai_suggestion_btn.configure(state=DISABLED)
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
        self._pdf_question_editor_baseline_stem = ""
        self._pdf_option_editor_baseline_texts = {}
        self._pdf_editor_updating = True
        if getattr(self, "_pdf_question_stem_editor", None):
            self._pdf_question_stem_editor.configure(state=NORMAL)
            self._pdf_question_stem_editor.delete("1.0", tk.END)
            self._pdf_question_stem_editor.configure(state=DISABLED)
        self._set_pdf_stem_preview_message("当前题目没有题干图片。", "暂无题干图片")
        self._rebuild_pdf_option_editors(None)
        self._pdf_question_layout_var.set("")
        self._pdf_preview_selected_block = None
        self._pdf_preview_drag_state = None
        self._pdf_layout_editor_status_var.set(
            "选择一道题后，可像编辑 PPT 一样拖动题干区、图片区和选项区。"
        )
        self._sync_pdf_layout_fields()
        self._set_pdf_question_editor_state(False)
        self._pdf_editor_updating = False
        self._pdf_question_editor_message.set(
            message or "选择一道题后，可在这里实时修改题干、选项内容，并为该题单独切换选项布局。"
        )
        self._refresh_pdf_ai_suggestion({})
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
        self._pdf_question_editor_baseline_stem = question.stem or ""
        self._pdf_option_editor_baseline_texts = {
            option.letter: option.text or ""
            for option in getattr(question, "options", []) or []
        }
        self._pdf_editor_updating = True
        self._set_pdf_question_editor_state(True)
        self._pdf_question_stem_editor.configure(state=NORMAL)
        self._pdf_question_stem_editor.delete("1.0", tk.END)
        self._pdf_question_stem_editor.insert("1.0", question.stem or "")
        self._refresh_pdf_stem_preview_for_question(question)
        self._rebuild_pdf_option_editors(question)
        self._pdf_question_layout_var.set((question.option_layout or "").strip().lower())
        self._pdf_preview_selected_block = "stem"
        self._pdf_preview_drag_state = None
        self._pdf_editor_updating = False
        self._set_pdf_question_editor_state(True)
        self._refresh_pdf_question_editor_message(payload)
        self._refresh_pdf_ai_suggestion(payload)
        self._refresh_pdf_layout_editor_status(question)
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

        payload["text"] = self._question_preview_text(
            section,
            material,
            question,
            slide_number=payload.get("slide_number"),
        )
        item_id = self._selected_pdf_item_id()
        if item_id:
            self.pdf_tree.item(
                item_id,
                text=self._question_tree_label(question),
                values=("question", question.source_number or "-", len(question.options)),
            )
        slide_item_id = self._pdf_slide_payload_to_item_id.get(item_id)
        if slide_item_id and getattr(self, "_pdf_slide_tree", None):
            self._pdf_slide_tree.item(
                slide_item_id,
                values=(
                    payload.get("slide_number") or "-",
                    question.source_number or "-",
                    self._document_subject_label(payload.get("section_kind") or "unknown"),
                    self._question_tree_label(question),
                ),
            )
        self._set_pdf_detail(payload["text"])
        self._refresh_pdf_question_editor_message(payload)
        self._refresh_pdf_ai_suggestion(payload)
        self._render_pdf_question_editor_preview()

    def _on_pdf_question_stem_change(self, event=None):
        if self._pdf_editor_updating or self._pdf_question_editor_target is None:
            return
        question = self._pdf_question_editor_target
        content = self._pdf_question_stem_editor.get("1.0", tk.END).rstrip("\n")
        previous_content = question.stem or ""
        focusout = self._is_focusout_event(event)
        if content != previous_content:
            update_question_stem(question, content)
            self._mark_pdf_project_dirty()
            self._sync_selected_question_preview_payload()
        if focusout and content != self._pdf_question_editor_baseline_stem:
            before_state = self._capture_pdf_question_state_override(
                question,
                stem=self._pdf_question_editor_baseline_stem,
            )
            self._log_pdf_question_event(
                action="update_question_stem",
                question=question,
                before_state=before_state,
                metadata={"field": "stem"},
            )
            self._pdf_question_editor_baseline_stem = content

    def _on_pdf_question_layout_change(self):
        if self._pdf_editor_updating or self._pdf_question_editor_target is None:
            return
        before_state = self._capture_pdf_question_state(self._pdf_question_editor_target)
        set_question_option_layout(self._pdf_question_editor_target, self._pdf_question_layout_var.get())
        self._mark_pdf_project_dirty()
        self._sync_selected_question_preview_payload()
        self._log_pdf_question_event(
            action="set_question_option_layout",
            question=self._pdf_question_editor_target,
            before_state=before_state,
            metadata={"layout": self._pdf_question_layout_var.get() or ""},
        )

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

    def _on_pdf_option_text_change(self, letter: str, event=None):
        if self._pdf_editor_updating or self._pdf_question_editor_target is None:
            return
        question = self._pdf_question_editor_target
        editor = self._pdf_option_editors.get(letter)
        if editor is None:
            return
        option = self._find_pdf_question_option(question, letter)
        if option is None:
            return
        content = editor.get("1.0", tk.END).rstrip("\n")
        previous_content = option.text or ""
        focusout = self._is_focusout_event(event)
        if content != previous_content and update_option_text(question, letter, content):
            self._mark_pdf_project_dirty()
            self._sync_selected_question_preview_payload()
        if focusout and content != self._pdf_option_editor_baseline_texts.get(letter, ""):
            before_state = self._capture_pdf_question_state_override(
                question,
                option_text_overrides={letter: self._pdf_option_editor_baseline_texts.get(letter, "")},
            )
            self._log_pdf_question_event(
                action="update_option_text",
                question=question,
                before_state=before_state,
                metadata={"option_letter": letter},
            )
            self._pdf_option_editor_baseline_texts[letter] = content

    def _refresh_pdf_option_editor_after_structure_change(self, *, action: str = "", before_state: dict | None = None, metadata: dict | None = None):
        question = self._pdf_question_editor_target
        if question is None:
            return
        self._pdf_editor_updating = True
        self._rebuild_pdf_option_editors(question)
        self._pdf_editor_updating = False
        self._set_pdf_question_editor_state(True)
        self._mark_pdf_project_dirty()
        self._sync_selected_question_preview_payload()
        self._pdf_option_editor_baseline_texts = {
            option.letter: option.text or ""
            for option in getattr(question, "options", []) or []
        }
        if action and before_state:
            self._log_pdf_question_event(
                action=action,
                question=question,
                before_state=before_state,
                metadata=metadata,
            )

    def _move_pdf_option(self, letter: str, direction: int):
        question = self._pdf_question_editor_target
        if question is None:
            return
        before_state = self._capture_pdf_question_state(question)
        if move_option(question, letter, direction):
            self._refresh_pdf_option_editor_after_structure_change(
                action="move_option",
                before_state=before_state,
                metadata={"option_letter": letter, "direction": direction},
            )

    def _insert_pdf_option_after(self, letter: str):
        question = self._pdf_question_editor_target
        if question is None:
            return
        before_state = self._capture_pdf_question_state(question)
        if insert_option_after(question, letter):
            self._refresh_pdf_option_editor_after_structure_change(
                action="insert_option_after",
                before_state=before_state,
                metadata={"option_letter": letter},
            )

    def _remove_pdf_option(self, letter: str):
        question = self._pdf_question_editor_target
        if question is None:
            return
        if not messagebox.askyesno("确认删项", f"确定删除 {letter} 选项吗？", parent=self.root):
            return
        before_state = self._capture_pdf_question_state(question)
        if remove_option(question, letter):
            self._refresh_pdf_option_editor_after_structure_change(
                action="remove_option",
                before_state=before_state,
                metadata={"option_letter": letter},
            )

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
        before_state = self._capture_pdf_question_state(question)
        if replace_option_image(question, letter, staged_path):
            self._mark_pdf_project_dirty()
            self._refresh_pdf_option_image_status(letter)
            self._sync_selected_question_preview_payload()
            self._log_pdf_question_event(
                action="replace_option_image",
                question=question,
                before_state=before_state,
                metadata={"option_letter": letter},
            )

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
            before_state = self._capture_pdf_question_state(question)
            if replace_option_image(question, letter, staged_path):
                self._mark_pdf_project_dirty()
                self._refresh_pdf_option_image_status(letter)
                self._sync_selected_question_preview_payload()
                self._log_pdf_question_event(
                    action="recrop_option_image",
                    question=question,
                    before_state=before_state,
                    metadata={"option_letter": letter},
                )
        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)

    def _clear_pdf_option_image(self, letter: str):
        question = self._pdf_question_editor_target
        if question is None:
            return
        before_state = self._capture_pdf_question_state(question)
        if clear_option_image(question, letter):
            self._mark_pdf_project_dirty()
            self._refresh_pdf_option_image_status(letter)
            self._sync_selected_question_preview_payload()
            self._log_pdf_question_event(
                action="clear_option_image",
                question=question,
                before_state=before_state,
                metadata={"option_letter": letter},
            )

    def _draw_pdf_preview_thumbnail(self, canvas, image_path: str, x0: float, y0: float, width: float, height: float) -> bool:
        if not image_path or not os.path.exists(image_path):
            return False
        if width < 12 or height < 12:
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

    def _draw_pdf_preview_image_gallery(
        self,
        canvas,
        image_paths: list[str],
        x0: float,
        y0: float,
        width: float,
        height: float,
    ) -> int:
        valid_paths = [path for path in image_paths if path and os.path.exists(path)]
        if not valid_paths or width < 24 or height < 24:
            return 0
        count = len(valid_paths)
        if count == 1:
            return 1 if self._draw_pdf_preview_thumbnail(canvas, valid_paths[0], x0, y0, width, height) else 0
        gap = 6.0
        best_cols = 1
        best_score = -1.0
        max_cols = min(count, 4)
        for candidate_cols in range(1, max_cols + 1):
            candidate_rows = max(1, (count + candidate_cols - 1) // candidate_cols)
            candidate_cell_width = (width - gap * (candidate_cols - 1)) / candidate_cols
            candidate_cell_height = (height - gap * (candidate_rows - 1)) / candidate_rows
            candidate_score = min(candidate_cell_width, candidate_cell_height)
            if candidate_score > best_score:
                best_score = candidate_score
                best_cols = candidate_cols
        cols = best_cols
        rows = max(1, (count + cols - 1) // cols)
        cell_width = max(18.0, (width - gap * (cols - 1)) / cols)
        cell_height = max(18.0, (height - gap * (rows - 1)) / rows)
        drawn = 0
        for index, path in enumerate(valid_paths):
            row = index // cols
            col = index % cols
            cell_x0 = x0 + col * (cell_width + gap)
            cell_y0 = y0 + row * (cell_height + gap)
            if self._draw_pdf_preview_thumbnail(
                canvas,
                path,
                cell_x0,
                cell_y0,
                cell_width,
                cell_height,
            ):
                drawn += 1
        return drawn

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

    def _pdf_preview_block_label(self, block: str) -> str:
        if is_option_layout_block(block):
            letter = option_layout_block_letter(block)
            return f"{letter} 选项"
        return {
            "stem": "题干区",
            "image": "图片区",
            "options": "选项区",
        }.get((block or "").strip().lower(), "版式区")

    def _refresh_pdf_layout_editor_status(self, question=None):
        target = question or self._pdf_question_editor_target
        if target is None:
            self._pdf_layout_editor_status_var.set(
                "选择一道题后，可像编辑 PPT 一样拖动区块，或直接输入 X/Y/宽/高。"
            )
            return
        block = self._pdf_preview_selected_block or "stem"
        block_label = self._pdf_preview_block_label(block)
        override_count = len(getattr(target, "ppt_layout", {}) or {})
        if override_count:
            self._pdf_layout_editor_status_var.set(
                f"当前选中：{block_label}。已启用 {override_count} 个单题版式覆盖，可拖动、缩放，也可直接输入 X/Y/宽/高。"
            )
        else:
            self._pdf_layout_editor_status_var.set(
                f"当前选中：{block_label}。当前仍跟随全局版式，开始拖动或输入尺寸后会自动生成单题布局。"
            )

    def _pdf_preview_rect_to_inches(
        self,
        rect: tuple[float, float, float, float] | None,
    ) -> tuple[float, float, float, float]:
        if not rect:
            return (0.0, 0.0, 0.0, 0.0)
        slide_left, slide_top, slide_right, slide_bottom = self._pdf_preview_slide_bounds
        slide_width = max(1.0, slide_right - slide_left)
        slide_height = max(1.0, slide_bottom - slide_top)
        x0, y0, x1, y1 = rect
        return (
            max(0.0, (x0 - slide_left) * _PPT_SLIDE_WIDTH_IN / slide_width),
            max(0.0, (y0 - slide_top) * _PPT_SLIDE_HEIGHT_IN / slide_height),
            max(0.0, (x1 - x0) * _PPT_SLIDE_WIDTH_IN / slide_width),
            max(0.0, (y1 - y0) * _PPT_SLIDE_HEIGHT_IN / slide_height),
        )

    def _pdf_preview_inches_to_rect(
        self,
        x: float,
        y: float,
        width: float,
        height: float,
    ) -> tuple[float, float, float, float]:
        slide_left, slide_top, slide_right, slide_bottom = self._pdf_preview_slide_bounds
        slide_width = max(1.0, slide_right - slide_left)
        slide_height = max(1.0, slide_bottom - slide_top)
        px_x0 = slide_left + max(0.0, x) * slide_width / _PPT_SLIDE_WIDTH_IN
        px_y0 = slide_top + max(0.0, y) * slide_height / _PPT_SLIDE_HEIGHT_IN
        px_width = max(0.0, width) * slide_width / _PPT_SLIDE_WIDTH_IN
        px_height = max(0.0, height) * slide_height / _PPT_SLIDE_HEIGHT_IN
        return (px_x0, px_y0, px_x0 + px_width, px_y0 + px_height)

    def _constrain_pdf_preview_rect(
        self,
        block: str | None,
        rect: tuple[float, float, float, float],
    ) -> tuple[float, float, float, float]:
        bound_x0, bound_y0, bound_x1, bound_y1 = self._pdf_preview_parent_bounds(block)
        min_width = 52.0
        min_height = 34.0
        max_width = max(min_width, bound_x1 - bound_x0)
        max_height = max(min_height, bound_y1 - bound_y0)
        x0, y0, x1, y1 = rect
        width = min(max_width, max(min_width, x1 - x0))
        height = min(max_height, max(min_height, y1 - y0))
        x0 = min(bound_x1 - width, max(bound_x0, x0))
        y0 = min(bound_y1 - height, max(bound_y0, y0))
        return (x0, y0, x0 + width, y0 + height)

    def _sync_pdf_layout_fields(self, question=None):
        block = self._pdf_preview_selected_block or "stem"
        rect = self._pdf_preview_rects.get(block)
        x, y, width, height = self._pdf_preview_rect_to_inches(rect)
        self._pdf_layout_fields_updating = True
        try:
            self._pdf_layout_x_var.set(round(x, 2))
            self._pdf_layout_y_var.set(round(y, 2))
            self._pdf_layout_w_var.set(round(width, 2))
            self._pdf_layout_h_var.set(round(height, 2))
        finally:
            self._pdf_layout_fields_updating = False

    def _build_pdf_preview_render_question(self, question) -> Question:
        section, material = self._find_pdf_question_context(question)
        material_image_paths: list[str] = []
        question_image_paths = [
            asset.path
            for asset in getattr(question, "stem_assets", []) or []
            if getattr(asset, "path", None) and os.path.exists(asset.path)
        ]
        material_header = None
        material_text = None
        if section is not None and getattr(section, "kind", "") == "data" and material is not None:
            preview_source, preview_paths = self._material_preview_entries(material)
            material_image_paths = [path for path in preview_paths if path and os.path.exists(path)]
            if preview_source != "PDF 区域预览":
                material_header = (getattr(material, "header", "") or "").strip() or None
                material_text = (getattr(material, "body", "") or "").strip() or None
        image_paths = [*material_image_paths, *question_image_paths]
        return Question(
            number=getattr(question, "numeric_source_number", None) or 1,
            stem=getattr(question, "stem", "") or "",
            options=[
                Option(
                    letter=option.letter,
                    text=option.text,
                    image_path=option.image_path,
                )
                for option in getattr(question, "options", []) or []
            ],
            image_paths=image_paths,
            question_image_paths=question_image_paths,
            material_image_paths=material_image_paths,
            source_question_number=(getattr(question, "source_number", "") or "").strip() or None,
            material_header=material_header,
            material_text=material_text,
            option_layout=getattr(question, "option_layout", None),
            ppt_layout=copy.deepcopy(getattr(question, "ppt_layout", {}) or {}),
            section_kind=getattr(section, "kind", None) if section is not None else None,
            section_title=getattr(section, "title", None) if section is not None else None,
        )

    def _effective_pdf_question_layout(self, question) -> dict[str, dict[str, float]]:
        render_question = self._build_pdf_preview_render_question(question)
        return build_effective_question_layout(render_question, self._make_ppt_config())

    def _pdf_preview_hit_test(self, x: float, y: float) -> tuple[str | None, str | None]:
        option_blocks = sorted(
            block
            for block in self._pdf_preview_rects.keys()
            if is_option_layout_block(block)
        )
        for block in [*option_blocks, "options", "image", "stem"]:
            rect = self._pdf_preview_rects.get(block)
            if not rect:
                continue
            x0, y0, x1, y1 = rect
            if not (x0 <= x <= x1 and y0 <= y <= y1):
                continue
            handle_size = 10
            if x >= x1 - handle_size and y >= y1 - handle_size:
                return block, "resize"
            return block, "move"
        return None, None

    def _move_option_block_overrides_with_options_region(
        self,
        question,
        *,
        old_rect: tuple[float, float, float, float],
        new_rect: tuple[float, float, float, float],
    ):
        overrides = dict(getattr(question, "ppt_layout", {}) or {})
        option_override_blocks = [
            block for block in overrides.keys() if is_option_layout_block(block)
        ]
        if not option_override_blocks:
            return
        old_x0, old_y0, old_x1, old_y1 = old_rect
        new_x0, new_y0, new_x1, new_y1 = new_rect
        old_w = max(1e-6, old_x1 - old_x0)
        old_h = max(1e-6, old_y1 - old_y0)
        new_w = max(1e-6, new_x1 - new_x0)
        new_h = max(1e-6, new_y1 - new_y0)
        effective_layout = self._effective_pdf_question_layout(question)
        for block in option_override_blocks:
            child_rect = effective_layout.get(block)
            if not child_rect:
                continue
            child_x0, child_y0, child_x1, child_y1 = scale_layout_rect(
                child_rect,
                max(1.0, self._pdf_preview_slide_bounds[2] - self._pdf_preview_slide_bounds[0]),
                max(1.0, self._pdf_preview_slide_bounds[3] - self._pdf_preview_slide_bounds[1]),
                offset_x=self._pdf_preview_slide_bounds[0],
                offset_y=self._pdf_preview_slide_bounds[1],
            )
            rel_x = (child_x0 - old_x0) / old_w
            rel_y = (child_y0 - old_y0) / old_h
            rel_w = (child_x1 - child_x0) / old_w
            rel_h = (child_y1 - child_y0) / old_h
            moved_rect = (
                new_x0 + rel_x * new_w,
                new_y0 + rel_y * new_h,
                new_x0 + (rel_x + rel_w) * new_w,
                new_y0 + (rel_y + rel_h) * new_h,
            )
            self._apply_pdf_preview_layout_rect(question, block, moved_rect)

    def _pdf_preview_parent_bounds(self, block: str | None) -> tuple[float, float, float, float]:
        if is_option_layout_block(block):
            return self._pdf_preview_rects.get("options") or self._pdf_preview_slide_bounds
        return self._pdf_preview_slide_bounds

    def _commit_pdf_preview_block_rect(
        self,
        question,
        block: str,
        rect: tuple[float, float, float, float],
        *,
        action: str,
        metadata: dict | None = None,
    ) -> bool:
        current_rect = self._pdf_preview_rects.get(block)
        if current_rect is None:
            return False
        before_state = self._capture_pdf_question_state(question)
        if block == "options":
            self._move_option_block_overrides_with_options_region(
                question,
                old_rect=current_rect,
                new_rect=rect,
            )
        self._apply_pdf_preview_layout_rect(question, block, rect)
        after_layout = (getattr(question, "ppt_layout", {}) or {}).get(block)
        before_layout = (before_state.get("ppt_layout") or {}).get(block)
        if before_layout == after_layout:
            return False
        self._mark_pdf_project_dirty()
        self._sync_selected_question_preview_payload()
        payload = dict(metadata or {})
        payload.setdefault("block", block)
        payload.setdefault("rect", after_layout or {})
        self._log_pdf_question_event(
            action=action,
            question=question,
            before_state=before_state,
            metadata=payload,
        )
        self._refresh_pdf_layout_editor_status(question)
        self._render_pdf_question_editor_preview()
        return True

    def _align_selected_pdf_preview_block(self, action: str):
        question = self._pdf_question_editor_target
        block = self._pdf_preview_selected_block
        if question is None or not block:
            return
        rect = self._pdf_preview_rects.get(block)
        if not rect:
            return
        x0, y0, x1, y1 = rect
        width = x1 - x0
        height = y1 - y0
        bound_x0, bound_y0, bound_x1, bound_y1 = self._pdf_preview_parent_bounds(block)
        if action == "left":
            new_rect = (bound_x0, y0, bound_x0 + width, y1)
        elif action == "hcenter":
            new_x0 = bound_x0 + max(0.0, (bound_x1 - bound_x0 - width) / 2.0)
            new_rect = (new_x0, y0, new_x0 + width, y1)
        elif action == "right":
            new_x1 = bound_x1
            new_rect = (new_x1 - width, y0, new_x1, y1)
        elif action == "top":
            new_rect = (x0, bound_y0, x1, bound_y0 + height)
        elif action == "vcenter":
            new_y0 = bound_y0 + max(0.0, (bound_y1 - bound_y0 - height) / 2.0)
            new_rect = (x0, new_y0, x1, new_y0 + height)
        elif action == "bottom":
            new_y1 = bound_y1
            new_rect = (x0, new_y1 - height, x1, new_y1)
        elif action == "fill_width":
            new_rect = (bound_x0, y0, bound_x1, y1)
        elif action == "fill_height":
            new_rect = (x0, bound_y0, x1, bound_y1)
        else:
            return
        new_rect = self._constrain_pdf_preview_rect(block, new_rect)
        self._commit_pdf_preview_block_rect(
            question,
            block,
            new_rect,
            action="align_question_ppt_layout",
            metadata={"align": action},
        )

    def _apply_pdf_preview_numeric_layout(self, event=None):
        if self._pdf_layout_fields_updating:
            return "break" if event is not None else None
        question = self._pdf_question_editor_target
        block = self._pdf_preview_selected_block
        if question is None or not block:
            return "break" if event is not None else None
        try:
            x = float(self._pdf_layout_x_var.get())
            y = float(self._pdf_layout_y_var.get())
            width = float(self._pdf_layout_w_var.get())
            height = float(self._pdf_layout_h_var.get())
        except (tk.TclError, ValueError):
            self._sync_pdf_layout_fields(question)
            return "break" if event is not None else None
        rect = self._pdf_preview_inches_to_rect(x, y, width, height)
        rect = self._constrain_pdf_preview_rect(block, rect)
        self._commit_pdf_preview_block_rect(
            question,
            block,
            rect,
            action="set_question_ppt_layout_fields",
            metadata={
                "units": "in",
                "x": round(x, 2),
                "y": round(y, 2),
                "width": round(width, 2),
                "height": round(height, 2),
            },
        )
        return "break" if event is not None else None

    def _on_pdf_question_preview_nudge(self, event):
        question = self._pdf_question_editor_target
        block = self._pdf_preview_selected_block
        if question is None or not block:
            return "break"
        rect = self._pdf_preview_rects.get(block)
        if not rect:
            return "break"
        step = 8.0 if (event.state & 0x0001) else 2.0
        resize = bool(event.state & 0x0004)
        x0, y0, x1, y1 = rect
        bound_x0, bound_y0, bound_x1, bound_y1 = self._pdf_preview_parent_bounds(block)
        delta_x = 0.0
        delta_y = 0.0
        keysym = str(getattr(event, "keysym", "") or "")
        if keysym == "Left":
            delta_x = -step
        elif keysym == "Right":
            delta_x = step
        elif keysym == "Up":
            delta_y = -step
        elif keysym == "Down":
            delta_y = step
        else:
            return "break"
        if resize:
            min_width = 52.0
            min_height = 34.0
            new_x1 = min(bound_x1, max(x0 + min_width, x1 + delta_x))
            new_y1 = min(bound_y1, max(y0 + min_height, y1 + delta_y))
            new_rect = (x0, y0, new_x1, new_y1)
        else:
            width = x1 - x0
            height = y1 - y0
            new_x0 = min(bound_x1 - width, max(bound_x0, x0 + delta_x))
            new_y0 = min(bound_y1 - height, max(bound_y0, y0 + delta_y))
            new_rect = (new_x0, new_y0, new_x0 + width, new_y0 + height)
        new_rect = self._constrain_pdf_preview_rect(block, new_rect)
        self._commit_pdf_preview_block_rect(
            question,
            block,
            new_rect,
            action="nudge_question_ppt_layout",
            metadata={
                "keysym": keysym,
                "resize": resize,
                "step": step,
            },
        )
        return "break"

    def _update_pdf_preview_cursor(self, event=None):
        canvas = getattr(self, "_pdf_question_preview_canvas", None)
        if canvas is None:
            return
        if self._pdf_preview_drag_state:
            mode = str(self._pdf_preview_drag_state.get("mode") or "move")
            canvas.configure(cursor="size_nw_se" if mode == "resize" else "fleur")
            return
        if event is None:
            canvas.configure(cursor="")
            return
        block, mode = self._pdf_preview_hit_test(event.x, event.y)
        if not block:
            canvas.configure(cursor="")
        elif mode == "resize":
            canvas.configure(cursor="size_nw_se")
        else:
            canvas.configure(cursor="fleur")

    def _apply_pdf_preview_layout_rect(
        self,
        question,
        block: str,
        rect: tuple[float, float, float, float],
    ):
        slide_left, slide_top, slide_right, slide_bottom = self._pdf_preview_slide_bounds
        slide_width = max(1.0, slide_right - slide_left)
        slide_height = max(1.0, slide_bottom - slide_top)
        normalized = normalize_layout_rect(
            rect[0],
            rect[1],
            rect[2],
            rect[3],
            origin_x=slide_left,
            origin_y=slide_top,
            width=slide_width,
            height=slide_height,
        )
        set_question_ppt_layout_block(question, block, normalized)
        self._refresh_pdf_layout_editor_status(question)

    def _reset_pdf_question_ppt_layout(self):
        question = self._pdf_question_editor_target
        if question is None or not getattr(question, "ppt_layout", None):
            return
        before_state = self._capture_pdf_question_state(question)
        clear_question_ppt_layout(question)
        self._pdf_preview_selected_block = "stem"
        self._mark_pdf_project_dirty()
        self._sync_selected_question_preview_payload()
        self._refresh_pdf_layout_editor_status(question)
        self._log_pdf_question_event(
            action="reset_question_ppt_layout",
            question=question,
            before_state=before_state,
        )

    def _on_pdf_question_preview_press(self, event):
        question = self._pdf_question_editor_target
        if question is None:
            return
        canvas = getattr(self, "_pdf_question_preview_canvas", None)
        if canvas is not None:
            canvas.focus_set()
        block, mode = self._pdf_preview_hit_test(event.x, event.y)
        if not block:
            self._pdf_preview_selected_block = None
            self._refresh_pdf_layout_editor_status(question)
            self._render_pdf_question_editor_preview()
            self._update_pdf_preview_cursor(event)
            return
        rect = self._pdf_preview_rects.get(block)
        if not rect:
            return
        self._pdf_preview_selected_block = block
        self._pdf_preview_drag_state = {
            "block": block,
            "mode": mode,
            "start_x": float(event.x),
            "start_y": float(event.y),
            "origin_rect": rect,
            "before_state": self._capture_pdf_question_state(question),
        }
        self._refresh_pdf_layout_editor_status(question)
        self._render_pdf_question_editor_preview()
        self._update_pdf_preview_cursor(event)

    def _on_pdf_question_preview_drag(self, event):
        question = self._pdf_question_editor_target
        drag_state = self._pdf_preview_drag_state
        if question is None or not drag_state:
            return
        block = str(drag_state.get("block") or "")
        mode = str(drag_state.get("mode") or "move")
        x0, y0, x1, y1 = drag_state.get("origin_rect") or (0.0, 0.0, 0.0, 0.0)
        start_x = float(drag_state.get("start_x") or 0.0)
        start_y = float(drag_state.get("start_y") or 0.0)
        dx = float(event.x) - start_x
        dy = float(event.y) - start_y
        slide_left, slide_top, slide_right, slide_bottom = self._pdf_preview_slide_bounds
        if is_option_layout_block(block):
            option_bounds = self._pdf_preview_rects.get("options")
            if option_bounds:
                slide_left, slide_top, slide_right, slide_bottom = option_bounds
        min_width = 52.0
        min_height = 34.0
        if mode == "resize":
            new_x0, new_y0 = x0, y0
            new_x1 = min(slide_right, max(x0 + min_width, x1 + dx))
            new_y1 = min(slide_bottom, max(y0 + min_height, y1 + dy))
        else:
            rect_width = x1 - x0
            rect_height = y1 - y0
            new_x0 = min(slide_right - rect_width, max(slide_left, x0 + dx))
            new_y0 = min(slide_bottom - rect_height, max(slide_top, y0 + dy))
            new_x1 = new_x0 + rect_width
            new_y1 = new_y0 + rect_height
        new_x0, new_y0, new_x1, new_y1 = self._constrain_pdf_preview_rect(
            block,
            (new_x0, new_y0, new_x1, new_y1),
        )
        if block == "options":
            self._move_option_block_overrides_with_options_region(
                question,
                old_rect=(x0, y0, x1, y1),
                new_rect=(new_x0, new_y0, new_x1, new_y1),
            )
        self._apply_pdf_preview_layout_rect(question, block, (new_x0, new_y0, new_x1, new_y1))
        self._render_pdf_question_editor_preview()
        self._update_pdf_preview_cursor(event)

    def _on_pdf_question_preview_release(self, event):
        question = self._pdf_question_editor_target
        drag_state = self._pdf_preview_drag_state
        self._pdf_preview_drag_state = None
        if question is None or not drag_state:
            self._update_pdf_preview_cursor(event)
            return
        block = str(drag_state.get("block") or "")
        before_state = drag_state.get("before_state") or {}
        before_layout = (before_state.get("ppt_layout") or {}).get(block)
        after_layout = (getattr(question, "ppt_layout", {}) or {}).get(block)
        if before_layout != after_layout:
            self._mark_pdf_project_dirty()
            self._sync_selected_question_preview_payload()
            self._log_pdf_question_event(
                action="set_question_ppt_layout",
                question=question,
                before_state=before_state,
                metadata={
                    "block": block,
                    "mode": drag_state.get("mode") or "move",
                    "rect": after_layout or {},
                },
            )
        self._refresh_pdf_layout_editor_status(question)
        self._render_pdf_question_editor_preview()
        self._update_pdf_preview_cursor(event)

    def _on_pdf_question_preview_motion(self, event):
        self._update_pdf_preview_cursor(event)

    def _render_pdf_question_editor_preview(self):
        canvas = getattr(self, "_pdf_question_preview_canvas", None)
        if canvas is None:
            return

        self._pdf_question_preview_photos = []
        self._pdf_preview_rects = {}
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
            self._sync_pdf_layout_fields(None)
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
        self._pdf_preview_slide_bounds = (slide_left, slide_top, slide_right, slide_bottom)

        scale = slide_width / 13.333
        self._pdf_preview_scale_px_per_in = scale

        canvas.create_rectangle(
            slide_left,
            slide_top,
            slide_right,
            slide_bottom,
            fill="#ffffff",
            outline="#b8c2cc",
            width=2,
        )

        render_question = self._build_pdf_preview_render_question(question)
        layout = build_effective_question_layout(render_question, self._make_ppt_config())
        selected_block = self._pdf_preview_selected_block or "stem"
        if selected_block not in layout:
            selected_block = next(iter(layout.keys()), "stem")
            self._pdf_preview_selected_block = selected_block
        override_blocks = set((getattr(question, "ppt_layout", {}) or {}).keys())
        stem_value = PPTGenerator._stem_text_for_question(render_question).strip() or "未填写题干"
        stem_rect = layout.get("stem")
        if stem_rect:
            stem_x0, stem_y0, stem_x1, stem_y1 = scale_layout_rect(
                stem_rect,
                slide_width,
                slide_height,
                offset_x=slide_left,
                offset_y=slide_top,
            )
            self._pdf_preview_rects["stem"] = (stem_x0, stem_y0, stem_x1, stem_y1)
            stem_outline = "#4a90d9" if selected_block == "stem" else "#7ba7da"
            canvas.create_rectangle(
                stem_x0,
                stem_y0,
                stem_x1,
                stem_y1,
                fill="#eef6ff",
                outline=stem_outline,
                width=3 if selected_block == "stem" else 2,
            )
            canvas.create_text(
                stem_x0 + 8,
                stem_y0 + 6,
                anchor=tk.NW,
                text=f"题干区{' · 单题' if 'stem' in override_blocks else ''}",
                fill=stem_outline,
                font=("", 9, "bold"),
            )
            wrapped_stem, stem_font = self._fit_preview_text(
                stem_value,
                width=max(48, (stem_x1 - stem_x0) - 16),
                height=max(40, (stem_y1 - stem_y0) - 24),
                target_points=float(self.font_size_stem.get()),
                scale_px_per_in=scale,
                bold=self.stem_bold.get(),
            )
            stem_align = (self.stem_align.get() or "left").lower()
            if stem_align == "center":
                stem_anchor = tk.N
                stem_justify = CENTER
                stem_x = stem_x0 + (stem_x1 - stem_x0) / 2
            elif stem_align == "right":
                stem_anchor = tk.NE
                stem_justify = RIGHT
                stem_x = stem_x1 - 8
            else:
                stem_anchor = tk.NW
                stem_justify = LEFT
                stem_x = stem_x0 + 8
            canvas.create_text(
                stem_x,
                stem_y0 + 24,
                anchor=stem_anchor,
                width=max(40, (stem_x1 - stem_x0) - 16),
                text=wrapped_stem,
                justify=stem_justify,
                fill=self.color_stem.get().strip() or "#1A1A2E",
                font=stem_font,
            )

        preview_images = list(render_question.image_paths)
        image_rect = layout.get("image")
        if image_rect:
            image_x0, image_y0, image_x1, image_y1 = scale_layout_rect(
                image_rect,
                slide_width,
                slide_height,
                offset_x=slide_left,
                offset_y=slide_top,
            )
            self._pdf_preview_rects["image"] = (image_x0, image_y0, image_x1, image_y1)
            image_outline = "#d9a84a" if selected_block == "image" else "#e0b96a"
            canvas.create_rectangle(
                image_x0,
                image_y0,
                image_x1,
                image_y1,
                fill="#fff8e6",
                outline=image_outline,
                width=3 if selected_block == "image" else 2,
                dash=(5, 3),
            )
            canvas.create_text(
                image_x0 + 8,
                image_y0 + 6,
                anchor=tk.NW,
                text=f"图片区{' · 单题' if 'image' in override_blocks else ''}",
                fill="#996600",
                font=("", 9, "bold"),
            )
            if preview_images:
                drawn = self._draw_pdf_preview_image_gallery(
                    canvas,
                    preview_images,
                    image_x0 + 8,
                    image_y0 + 24,
                    max(24, (image_x1 - image_x0) - 16),
                    max(24, (image_y1 - image_y0) - 34),
                )
                if drawn and len(preview_images) > 1:
                    canvas.create_text(
                        image_x1 - 8,
                        image_y0 + 8,
                        anchor=tk.NE,
                        text=f"{len(preview_images)} 张",
                        fill="#996600",
                        font=("", 9, "bold"),
                    )
                elif not drawn:
                    canvas.create_text(
                        (image_x0 + image_x1) / 2,
                        (image_y0 + image_y1) / 2,
                        text=f"图片素材 × {len(preview_images)}",
                        fill="#996600",
                        font=("", 10, "bold"),
                    )
            else:
                canvas.create_text(
                    (image_x0 + image_x1) / 2,
                    (image_y0 + image_y1) / 2,
                    text="当前题目没有图片素材",
                    fill="#996600",
                    font=("", 10, "bold"),
                )

        options_rect = layout.get("options")
        layout_kind = self._effective_question_option_layout(question)
        options = list(render_question.options[:4])
        fills = ["#e3f2fd", "#e8f5e9", "#fce4ec", "#fff3e0"]
        outlines = ["#1976d2", "#43a047", "#c62828", "#e65100"]
        if options_rect:
            options_x0, options_y0, options_x1, options_y1 = scale_layout_rect(
                options_rect,
                slide_width,
                slide_height,
                offset_x=slide_left,
                offset_y=slide_top,
            )
            self._pdf_preview_rects["options"] = (options_x0, options_y0, options_x1, options_y1)
            option_outline = "#6d7f94" if selected_block == "options" else "#aab6c4"
            canvas.create_rectangle(
                options_x0,
                options_y0,
                options_x1,
                options_y1,
                fill="#f7fafc",
                outline=option_outline,
                width=3 if selected_block == "options" else 2,
            )
            canvas.create_text(
                options_x0 + 8,
                options_y0 + 6,
                anchor=tk.NW,
                text=f"选项区{' · 单题' if 'options' in override_blocks else ''}",
                fill=option_outline,
                font=("", 9, "bold"),
            )
            if not options:
                canvas.create_text(
                    (options_x0 + options_x1) / 2,
                    (options_y0 + options_y1) / 2,
                    text="当前题目没有可预览的选项。",
                    fill="#6b7280",
                    font=("", 10),
                )
            else:
                for index, option in enumerate(options):
                    block = option_layout_block_key(option.letter)
                    option_rect = layout.get(block)
                    if not option_rect:
                        continue
                    item_x0, item_y0, item_x1, item_y1 = scale_layout_rect(
                        option_rect,
                        slide_width,
                        slide_height,
                        offset_x=slide_left,
                        offset_y=slide_top,
                    )
                    self._pdf_preview_rects[block] = (item_x0, item_y0, item_x1, item_y1)
                    outline = outlines[index]
                    fill = fills[index]
                    if selected_block == block:
                        outline = "#111827"
                    self._draw_pdf_question_preview_option(
                        canvas,
                        option,
                        item_x0,
                        item_y0,
                        max(28, item_x1 - item_x0),
                        max(28, item_y1 - item_y0),
                        fill,
                        outline,
                    )
                    canvas.create_text(
                        item_x0 + 8,
                        item_y0 + 6,
                        anchor=tk.NW,
                        text=f"{option.letter}{' · 单题' if block in override_blocks else ''}",
                        fill=outline,
                        font=("", 8, "bold"),
                    )

        for block, rect in self._pdf_preview_rects.items():
            handle_fill = "#2563eb" if block == selected_block else "#94a3b8"
            x0, y0, x1, y1 = rect
            canvas.create_rectangle(
                x1 - 8,
                y1 - 8,
                x1 + 2,
                y1 + 2,
                fill=handle_fill,
                outline="#ffffff",
                width=1,
            )

        layout_label = self._option_layout_label(question.option_layout or layout_kind)
        layout_source = "单题覆盖" if question.option_layout else "跟随全局"
        canvas.create_text(
            width - 18,
            height - 16,
            anchor=tk.SE,
            text=f"{layout_source} · {layout_label} · 拖动块内移动，拖右下角缩放",
            fill="#6b7280",
            font=("", 9),
        )
        self._sync_pdf_layout_fields(question)

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
        selected_question = self._pdf_question_editor_target
        self._reset_pdf_material_preview_session()
        self._clear_pdf_question_editor()
        self._populate_pdf_preview(self.pdf_project)
        if selected_question is not None:
            item_id = self._find_preview_item_for_question(selected_question)
            if item_id:
                self._select_pdf_preview_item(item_id)
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
        before_state = self._capture_pdf_question_state(question)
        renumber_question(question, new_number)
        self._mark_pdf_project_dirty()
        self._log_pdf_question_event(
            action="renumber_question",
            question=question,
            before_state=before_state,
            metadata={"new_number": new_number},
        )
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

    def _question_preview_text(self, section, material, question, *, slide_number: int | None = None) -> str:
        effective_layout = self._effective_question_option_layout(question)
        layout_text = self._option_layout_label(question.option_layout or effective_layout)
        layout_source = "单题覆盖" if question.option_layout else "跟随全局"
        confidence_text = f"{int(round((getattr(question, 'review_confidence', 1.0) or 1.0) * 100))}%"
        lines = [
            f"科目：{section.kind}",
            f"原题号：{question.source_number or '-'}",
            f"选项数：{len(question.options)}",
            f"题干图片：{len(question.stem_assets)}",
            f"选项布局：{layout_source}（{layout_text}）",
            f"AI 质检置信度：{confidence_text}",
        ]
        if getattr(question, "inferred_subtype", ""):
            subtype_conf = getattr(question, "inferred_subtype_confidence", None)
            subtype_line = f"推断子题型：{question.inferred_subtype}"
            if subtype_conf is not None:
                subtype_line += f"（{int(round(max(0.0, min(subtype_conf, 1.0)) * 100))}%）"
            lines.append(subtype_line)
        if getattr(question, "inferred_signals", None):
            lines.append("命中线索：" + "、".join(question.inferred_signals[:4]))
        if slide_number:
            lines.insert(0, f"PPT 页码：第 {slide_number} 页")
        if material is not None:
            lines.append(f"所属材料：{material.header or material.material_id}")
        if question.page_numbers:
            lines.append("来源页码：" + ", ".join(str(page_no) for page_no in question.page_numbers))
        if getattr(question, "review_issues", None):
            lines.extend(["", "待确认项："])
            for issue in question.review_issues:
                detail = f"：{issue.detail}" if issue.detail else ""
                lines.append(f"[{issue.severity}] {issue.title}{detail}")
        if getattr(question, "suggested_subject", None):
            suggestion_label = self._document_subject_label(question.suggested_subject)
            if getattr(question, "inferred_subtype", ""):
                suggestion_label += f" / {question.inferred_subtype}"
            lines.extend(
                [
                    "",
                    "AI 建议：",
                    f"更像 {suggestion_label}"
                    + (
                        f"（置信度 {int(round(max(0.0, min(question.suggested_subject_confidence, 1.0)) * 100))}%）"
                        if question.suggested_subject_confidence is not None
                        else ""
                    ),
                ]
            )
            if question.suggested_subject_reason:
                lines.append(question.suggested_subject_reason)
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
        if not reclassify_objective_section(section, new_kind, project=self.pdf_project):
            return
        self._mark_pdf_project_dirty()
        self._log_pdf_project_event(
            action="reclassify_section_subject",
            metadata={
                "section_title": getattr(section, "title", ""),
                "new_kind": new_kind,
            },
        )
        self._refresh_pdf_preview_after_edit(f"已将当前篇题调整为：{self._document_subject_label(new_kind)}")

    def _apply_selected_ai_subject_suggestion(self):
        if self.pdf_project is None:
            return
        payload = self._selected_pdf_payload()
        section = payload.get("section")
        if section is None:
            messagebox.showinfo("提示", "请先选中一道题目，再应用 AI 建议。", parent=self.root)
            return
        target_kind, reason = section_subject_suggestion(section)
        if not target_kind:
            messagebox.showinfo("提示", reason or "当前没有可安全应用的 AI 建议。", parent=self.root)
            return
        changed = apply_section_subject_suggestion(section, project=self.pdf_project)
        if not changed:
            messagebox.showinfo("提示", "AI 建议与当前篇题一致，未发生变化。", parent=self.root)
            return
        self._mark_pdf_project_dirty()
        self._log_pdf_project_event(
            action="apply_ai_section_subject_suggestion",
            source="gui_ai",
            metadata={
                "section_title": getattr(section, "title", ""),
                "target_kind": target_kind,
                "reason": reason,
            },
        )
        self._refresh_pdf_preview_after_edit(
            f"已按 AI 建议将当前篇题调整为：{self._document_subject_label(target_kind)}"
        )

    def _apply_all_safe_ai_suggestions(self):
        if self.pdf_project is None:
            return
        applied = apply_all_safe_subject_suggestions(self.pdf_project)
        if applied <= 0:
            messagebox.showinfo("提示", "当前没有可安全批量应用的 AI 科目建议。", parent=self.root)
            return
        self._mark_pdf_project_dirty()
        self._log_pdf_project_event(
            action="apply_all_safe_ai_subject_suggestions",
            source="gui_ai",
            metadata={"applied_sections": applied},
        )
        self._refresh_pdf_preview_after_edit(f"已批量应用 {applied} 条 AI 安全建议。")

    def _ai_batch_limit_value(self) -> int:
        try:
            value = int(self.ai_batch_limit.get())
        except Exception:
            value = 12
        return max(1, min(50, value))

    def _current_ai_mode(self) -> str:
        label = (self.ai_mode.get() or "").strip()
        for key, display in _AI_MODE_CHOICES:
            if label == display:
                return key
        return "balanced"

    def _build_ai_service(self) -> AIRepairService:
        mode = self._current_ai_mode()
        service = AIRepairService(mode=mode)
        mode_label = _AI_MODE_LABELS.get(mode, mode)
        if mode == "policy":
            self._ai_status_var.set(f"本地 AI 修复器已启用：当前是 {mode_label}，会叠加 learned repair policy 的排序与提示。")
        else:
            self._ai_status_var.set(f"本地 AI 修复器已启用：当前是 {mode_label}，会优先做安全规则修复。")
        return service

    def _set_ai_repair_busy(self, busy: bool, status_text: str | None = None):
        self._ai_repair_busy = busy
        if status_text:
            self._set_status(status_text)
        self._refresh_pdf_ai_suggestion()
        self._refresh_ocr_tool_buttons()

    def _set_ocr_tool_busy(self, busy: bool, status_text: str | None = None):
        self._ocr_tool_busy = busy
        if status_text:
            self._set_status(status_text)
        self._refresh_ocr_tool_buttons()
        self._refresh_pdf_ai_suggestion()

    def _refresh_ocr_tool_buttons(self):
        can_use = self.pdf_project is not None and not self._ai_repair_busy and not self._ocr_tool_busy
        for button in (
            getattr(self, "_pdf_ocr_diagnose_btn", None),
            getattr(self, "_pdf_ocr_repair_btn", None),
        ):
            if button is not None:
                button.configure(state=NORMAL if can_use else DISABLED)

    def _ai_neighbor_questions(self, section, material, question):
        if section is None or question is None:
            return None, None
        if getattr(section, "kind", "") == "data":
            rows = list(getattr(material, "questions", []) or [])
        else:
            rows = list(getattr(section, "questions", []) or [])
        for index, item in enumerate(rows):
            if item is question:
                previous_question = rows[index - 1] if index > 0 else None
                next_question = rows[index + 1] if index + 1 < len(rows) else None
                return previous_question, next_question
        return None, None

    def _find_pdf_question_context(self, question):
        if self.pdf_project is None or question is None:
            return None, None
        for section, material, item in self.pdf_project.iter_questions():
            if item is question:
                return section, material
        return None, None

    def _capture_pdf_question_state(self, question, *, section=None, material=None):
        if self.pdf_project is None or question is None:
            return {}
        if section is None:
            section, material = self._find_pdf_question_context(question)
        if section is None:
            return {}
        previous_question, next_question = self._ai_neighbor_questions(section, material, question)
        return capture_question_state(
            section=section,
            material=material,
            question=question,
            previous_question=previous_question,
            next_question=next_question,
        )

    def _capture_pdf_question_state_override(
        self,
        question,
        *,
        section=None,
        material=None,
        stem: str | None = None,
        option_text_overrides: dict[str, str] | None = None,
    ):
        if self.pdf_project is None or question is None:
            return {}
        if section is None:
            section, material = self._find_pdf_question_context(question)
        if section is None:
            return {}
        question_copy = copy.deepcopy(question)
        if stem is not None:
            question_copy.stem = stem
        for letter, value in (option_text_overrides or {}).items():
            option = self._find_pdf_question_option(question_copy, letter)
            if option is not None:
                option.text = value
        previous_question, next_question = self._ai_neighbor_questions(section, material, question)
        return capture_question_state(
            section=section,
            material=material,
            question=question_copy,
            previous_question=previous_question,
            next_question=next_question,
        )

    def _log_pdf_question_event(
        self,
        *,
        action: str,
        question,
        section=None,
        material=None,
        before_state: dict | None = None,
        metadata: dict | None = None,
        source: str = "gui_manual",
    ):
        if self.pdf_project is None or question is None:
            return
        if section is None:
            section, material = self._find_pdf_question_context(question)
        if section is None:
            return
        previous_question, next_question = self._ai_neighbor_questions(section, material, question)
        append_question_repair_log(
            self.pdf_project,
            source=source,
            action=action,
            section=section,
            material=material,
            question=question,
            previous_question=previous_question,
            next_question=next_question,
            before_state=before_state,
            metadata=metadata,
        )

    def _log_pdf_project_event(self, *, action: str, metadata: dict | None = None, source: str = "gui_manual"):
        if self.pdf_project is None:
            return
        append_project_repair_log(
            self.pdf_project,
            source=source,
            action=action,
            metadata=metadata,
        )

    def _is_focusout_event(self, event) -> bool:
        return str(getattr(event, "type", "")) == "FocusOut"

    def _refresh_pdf_ai_strategy(self, payload: dict | None = None):
        payload = payload or self._selected_pdf_payload()
        if payload.get("kind") != "question":
            self._pdf_ai_strategy_var.set("修复策略会在这里显示。")
            return
        question = payload.get("question")
        section = payload.get("section")
        material = payload.get("material")
        if question is None or section is None:
            self._pdf_ai_strategy_var.set("修复策略会在这里显示。")
            return
        previous_question, next_question = self._ai_neighbor_questions(section, material, question)
        snapshot = inspect_repair_strategy(
            section=section,
            material=material,
            question=question,
            previous_question=previous_question,
            next_question=next_question,
            mode=self._current_ai_mode(),
        )
        self._pdf_ai_strategy_var.set(build_repair_strategy_summary(snapshot))

    def _run_pdf_ocr_diagnostics(self):
        if self.pdf_project is None:
            messagebox.showinfo("提示", "请先生成或载入工程。", parent=self.root)
            return
        self._set_ocr_tool_busy(True, "正在分析扫描/OCR 风险…")

        def work():
            try:
                report = diagnose_project_ocr_risks(
                    self.pdf_project,
                    pdf_path=self.pdf_path.get() or getattr(getattr(self.pdf_project, "source", None), "pdf_path", None),
                )
                self.root.after(0, lambda r=report: self._on_pdf_ocr_diagnostics_done(r))
            except Exception as exc:
                self.root.after(0, lambda e=exc: self._on_pdf_ocr_tool_error(str(e)))

        threading.Thread(target=work, daemon=True).start()

    def _on_pdf_ocr_diagnostics_done(self, report):
        self._set_ocr_tool_busy(False)
        summary = report.summary_text()
        if getattr(report, "suggestions", None):
            summary += "\n" + report.suggestions[0]
        self._pdf_ocr_status_var.set(summary)
        self._set_status("扫描/OCR 诊断完成")

    def _auto_repair_pdf_ocr(self):
        if self.pdf_project is None:
            messagebox.showinfo("提示", "请先生成或载入工程。", parent=self.root)
            return
        service = self._build_ai_service()
        self._set_ocr_tool_busy(True, "正在执行 OCR 自动修补…")
        self.progress["value"] = 0
        self.progress["maximum"] = self._ai_batch_limit_value()

        def work():
            try:
                summary = auto_repair_ocr_project(
                    self.pdf_project,
                    pdf_path=self.pdf_path.get() or getattr(getattr(self.pdf_project, "source", None), "pdf_path", None),
                    service=service,
                    only_flagged=bool(self.ai_only_flagged.get()),
                    limit=self._ai_batch_limit_value(),
                )
                self.root.after(0, lambda s=summary: self._on_pdf_ocr_auto_repair_done(s))
            except Exception as exc:
                self.root.after(0, lambda e=exc: self._on_pdf_ocr_tool_error(str(e)))

        threading.Thread(target=work, daemon=True).start()

    def _on_pdf_ocr_auto_repair_done(self, summary):
        self.progress["value"] = self.progress["maximum"]
        self._set_ocr_tool_busy(False)
        self._pdf_ocr_status_var.set(summary.summary_text())
        changed = (
            summary.section_subject_changes > 0
            or summary.batch_summary.changed_questions > 0
            or summary.batch_summary.structural_repairs > 0
        )
        if changed:
            self._mark_pdf_project_dirty()
            self._log_pdf_project_event(
                action="ocr_auto_repair",
                source="gui_ocr",
                metadata=summary.to_dict(),
            )
            self._refresh_pdf_preview_after_edit("OCR 自动修补已完成：" + summary.summary_text())
            return
        self._set_status("OCR 自动修补完成，但没有形成可写回的变化。")
        self._refresh_pdf_ai_suggestion()

    def _on_pdf_ocr_tool_error(self, message: str):
        self._set_ocr_tool_busy(False)
        self.progress["value"] = 0
        self._set_status("扫描/OCR 工具失败")
        messagebox.showerror("扫描/OCR 工具失败", message, parent=self.root)

    def _repair_selected_question_with_ai(self):
        if self.pdf_project is None:
            messagebox.showinfo("提示", "请先生成或载入工程。", parent=self.root)
            return
        payload = self._selected_pdf_payload()
        if payload.get("kind") != "question":
            messagebox.showinfo("提示", "请先在左侧选择一道题目。", parent=self.root)
            return
        question = payload.get("question")
        section = payload.get("section")
        material = payload.get("material")
        if question is None or section is None:
            return
        service = self._build_ai_service()
        previous_question, next_question = self._ai_neighbor_questions(section, material, question)
        original_number = question.source_number or "-"
        self._set_ai_repair_busy(True, f"AI 正在修复原题号 {original_number}…")
        self.progress["value"] = 0
        self.progress["maximum"] = 1

        def work():
            try:
                before_state = self._capture_pdf_question_state(question, section=section, material=material)
                boundary_changes, boundary_reason = repair_question_boundary(
                    self.pdf_project,
                    section,
                    material,
                    question,
                )
                previous_question, next_question = self._ai_neighbor_questions(section, material, question)
                result = service.repair_question(
                    section=section,
                    material=material,
                    question=question,
                    previous_question=previous_question,
                    next_question=next_question,
                )
                changes, subject_changed = apply_ai_question_patch(
                    question,
                    result.patch,
                    section=section,
                    material=material,
                    project=self.pdf_project,
                )
                annotate_project_quality(self.pdf_project)
                self.root.after(
                    0,
                    lambda r=result, c=changes, s=subject_changed, n=original_number, bc=boundary_changes, br=boundary_reason, bs=before_state: self._on_ai_single_repair_done(
                        n,
                        r,
                        c,
                        s,
                        bc,
                        br,
                        bs,
                    ),
                )
            except Exception as exc:
                self.root.after(0, lambda e=exc: self._on_ai_repair_error(str(e)))

        threading.Thread(target=work, daemon=True).start()

    def _on_ai_single_repair_done(
        self,
        original_number: str,
        result,
        changes: int,
        subject_changed: bool,
        boundary_changes: int = 0,
        boundary_reason: str = "",
        before_state: dict | None = None,
    ):
        self.progress["value"] = self.progress["maximum"]
        self._set_ai_repair_busy(False)
        if not getattr(result.patch, "should_apply", False):
            if boundary_changes > 0:
                self._mark_pdf_project_dirty()
                self._refresh_pdf_preview_after_edit(
                    f"本地 AI 已完成边界修复：{boundary_reason or '已把串到当前题开头的选项拆回上一题。'}"
                )
                return
            summary = (getattr(result.patch, "summary", "") or "AI 判断当前题暂不适合自动写回。").strip()
            self._set_status(summary)
            self._refresh_pdf_ai_suggestion()
            return

        total_changes = changes + boundary_changes
        if total_changes <= 0:
            summary = (getattr(result.patch, "summary", "") or "AI 完成检查，但没有形成需要写回的字段变化。").strip()
            self._set_status(summary)
            self._refresh_pdf_ai_suggestion()
            return

        if subject_changed:
            self._mark_pdf_project_dirty()
        else:
            self._mark_pdf_project_dirty()
        summary = (getattr(result.patch, "summary", "") or "AI 已写回当前题的结构修复。").strip()
        if boundary_changes > 0:
            summary = (boundary_reason or "已处理跨题边界串题") + "；" + summary
        question = self._pdf_question_editor_target
        if question is not None:
            self._log_pdf_question_event(
                action="ai_single_repair",
                question=question,
                before_state=before_state,
                metadata={
                    "summary": summary,
                    "field_changes": total_changes,
                    "subject_changed": bool(subject_changed),
                    "boundary_changes": int(boundary_changes),
                    "mode": self._current_ai_mode(),
                },
                source="gui_ai",
            )
        self._refresh_pdf_preview_after_edit(f"AI 已修复原题号 {original_number}：{summary}")

    def _repair_flagged_questions_with_ai(self):
        if self.pdf_project is None:
            messagebox.showinfo("提示", "请先生成或载入工程。", parent=self.root)
            return
        service = self._build_ai_service()
        only_flagged = bool(self.ai_only_flagged.get())
        if only_flagged and not self._pdf_review_payload_ids:
            messagebox.showinfo("提示", "当前没有待确认题，暂时不需要批量 AI 修复。", parent=self.root)
            return
        limit = self._ai_batch_limit_value()
        scope_text = "待确认题" if only_flagged else "当前工程题目"
        self._set_ai_repair_busy(True, f"AI 正在批量修复{scope_text}…")
        self.progress["value"] = 0
        self.progress["maximum"] = limit

        def work():
            try:
                summary = repair_project_questions(
                    self.pdf_project,
                    service=service,
                    only_flagged=only_flagged,
                    limit=limit,
                )
                self.root.after(0, lambda s=summary: self._on_ai_batch_repair_done(s))
            except Exception as exc:
                self.root.after(0, lambda e=exc: self._on_ai_repair_error(str(e)))

        threading.Thread(target=work, daemon=True).start()

    def _on_ai_batch_repair_done(self, summary):
        self.progress["value"] = self.progress["maximum"]
        self._set_ai_repair_busy(False)
        if summary.changed_questions > 0:
            self._mark_pdf_project_dirty()
            self._log_pdf_project_event(
                action="ai_batch_repair",
                source="gui_ai",
                metadata={
                    "attempted_questions": summary.attempted_questions,
                    "changed_questions": summary.changed_questions,
                    "total_field_changes": summary.total_field_changes,
                    "subject_changes": summary.subject_changes,
                    "structural_repairs": summary.structural_repairs,
                    "mode": self._current_ai_mode(),
                },
            )
            message = (
                f"AI 批量处理 {summary.attempted_questions} 题，"
                f"写回 {summary.changed_questions} 题，"
                f"共 {summary.total_field_changes} 处字段"
            )
            if summary.subject_changes:
                message += f"，其中科目调整 {summary.subject_changes} 处"
            if summary.errors:
                message += f"，另有 {len(summary.errors)} 题未成功"
            self._refresh_pdf_preview_after_edit(message + "。")
            return

        if summary.errors:
            preview = "\n".join(summary.errors[:3])
            messagebox.showwarning(
                "AI 修复",
                f"本次没有成功写回字段，失败 {len(summary.errors)} 题。\n\n{preview}",
                parent=self.root,
            )
            self._set_status("AI 批量修复未能写回字段。")
            self._refresh_pdf_ai_suggestion()
            return

        self._set_status("AI 批量检查完成，但没有形成需要写回的修复。")
        self._refresh_pdf_ai_suggestion()

    def _on_ai_repair_error(self, message: str):
        self._set_ai_repair_busy(False)
        self.progress["value"] = 0
        self._set_status("AI 修复失败")
        messagebox.showerror("AI 修复失败", message, parent=self.root)

    def _export_pdf_review_report(self):
        if self.pdf_project is None:
            messagebox.showinfo("提示", "请先生成预览工程。", parent=self.root)
            return
        default_path = self._default_pdf_base_path() + "_AI质检报告.json"
        path = filedialog.asksaveasfilename(
            title="导出 AI 质检报告",
            defaultextension=".json",
            initialfile=os.path.basename(default_path),
            initialdir=os.path.dirname(default_path) or None,
            filetypes=[("JSON", "*.json"), ("All", "*.*")],
        )
        if not path:
            return
        try:
            export_quality_report(self.pdf_project, path)
        except Exception as exc:
            messagebox.showerror("导出失败", str(exc), parent=self.root)
            return
        self._set_status(f"已导出 AI 质检报告：{path}")
        messagebox.showinfo("完成", f"AI 质检报告已导出：\n{path}", parent=self.root)

    def _populate_pdf_preview(self, project):
        self.pdf_project = project
        quality = self._reanalyze_pdf_project()
        self._pdf_preview_payloads.clear()
        self._clear_pdf_question_editor()
        for item in self.pdf_tree.get_children():
            self.pdf_tree.delete(item)
        if getattr(self, "_pdf_review_tree", None):
            for item in self._pdf_review_tree.get_children():
                self._pdf_review_tree.delete(item)
        if getattr(self, "_pdf_slide_tree", None):
            for item in self._pdf_slide_tree.get_children():
                self._pdf_slide_tree.delete(item)
        self._pdf_review_payload_ids = []
        self._pdf_review_payload_to_item_id = {}
        self._pdf_review_item_to_payload_id = {}
        self._pdf_slide_payload_ids = []
        self._pdf_slide_payload_to_item_id = {}
        self._pdf_slide_item_to_payload_id = {}
        self._pdf_slide_payload_to_number = {}

        section_index = 0
        question_payload_ids: dict[int, str] = {}
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
                            "text": "",
                        }
                        question_payload_ids[id(question)] = qid
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
                        "text": "",
                    }
                    question_payload_ids[id(question)] = qid

        flagged_rows = list(iter_flagged_question_rows(project))
        flagged_payload_order = {id(question): index for index, (_section, _material, question) in enumerate(flagged_rows)}

        for slide_number, (section, material, question) in enumerate(iter_project_question_nodes(project), start=1):
            payload_id = question_payload_ids.get(id(question))
            if not payload_id:
                continue
            payload = self._pdf_preview_payloads[payload_id]
            payload["slide_number"] = slide_number
            payload["text"] = self._question_preview_text(
                section,
                material,
                question,
                slide_number=slide_number,
            )
            if getattr(self, "_pdf_slide_tree", None):
                slide_item_id = self._pdf_slide_tree.insert(
                    "",
                    END,
                    values=(
                        slide_number,
                        question.source_number or "-",
                        self._document_subject_label(section.kind),
                        self._question_tree_label(question),
                    ),
                )
                self._pdf_slide_payload_ids.append(payload_id)
                self._pdf_slide_payload_to_item_id[payload_id] = slide_item_id
                self._pdf_slide_item_to_payload_id[slide_item_id] = payload_id
                self._pdf_slide_payload_to_number[payload_id] = slide_number
        if getattr(self, "_pdf_review_tree", None):
            sorted_flagged_payloads = sorted(
                (
                    payload_id
                    for payload_id in self._pdf_preview_payloads
                    if self._pdf_preview_payloads[payload_id].get("kind") == "question"
                    and is_flagged_question(self._pdf_preview_payloads[payload_id].get("question"))
                ),
                key=lambda payload_id: (
                    -severity_rank(question_max_severity(self._pdf_preview_payloads[payload_id]["question"])),
                    float(getattr(self._pdf_preview_payloads[payload_id]["question"], "review_confidence", 1.0) or 1.0),
                    flagged_payload_order.get(id(self._pdf_preview_payloads[payload_id]["question"]), 10**6),
                ),
            )
            for payload_id in sorted_flagged_payloads:
                payload = self._pdf_preview_payloads[payload_id]
                question = payload["question"]
                section = payload["section"]
                review_item_id = self._pdf_review_tree.insert(
                    "",
                    END,
                    values=(
                        self._review_severity_label(question),
                        f"{int(round((question.review_confidence or 1.0) * 100))}%",
                        question.source_number or "-",
                        self._document_subject_label(section.kind),
                        question_review_summary(question),
                    ),
                )
                self._pdf_review_payload_ids.append(payload_id)
                self._pdf_review_payload_to_item_id[payload_id] = review_item_id
                self._pdf_review_item_to_payload_id[review_item_id] = payload_id

        review_count = len(self._pdf_review_payload_ids)
        detail_text = f"已加载 {project.question_count} 道题。"
        if review_count:
            detail_text += (
                f"\nAI 质检标出 {review_count} 道待确认题"
                + (
                    f"，其中高风险 {quality.severe_questions} 题"
                    if quality is not None
                    else ""
                )
                + "，可切到左侧“待确认”页签优先处理。"
            )
        else:
            detail_text += "\nAI 质检暂未发现明显异常。"
        detail_text += "\n请在左侧查看篇题、材料和题目结构。"
        self._set_pdf_detail(detail_text)
        self._refresh_pdf_slide_status("")
        self._refresh_pdf_review_status("")
        self._refresh_ocr_tool_buttons()
        if self._pdf_review_payload_ids:
            self._select_pdf_preview_item(self._pdf_review_payload_ids[0], focus_left_tab="review")
        elif self._pdf_slide_payload_ids:
            self._select_pdf_preview_item(self._pdf_slide_payload_ids[0], focus_left_tab="slide")
        self._refresh_pdf_wizard_ui()

    def _on_pdf_preview_select(self, _event=None):
        if not getattr(self, "pdf_tree", None):
            return
        if self._pdf_preview_syncing_selection:
            return
        selected = self.pdf_tree.selection()
        if not selected:
            self._apply_pdf_preview_selection("", focus_left_tab="structure")
            return
        self._apply_pdf_preview_selection(selected[0], focus_left_tab="structure")

    def _on_pdf_slide_select(self, _event=None):
        if self._pdf_slide_syncing_selection:
            return
        payload_id = self._selected_pdf_slide_payload_id()
        if not payload_id:
            self._refresh_pdf_slide_status("")
            return
        self._select_pdf_preview_item(payload_id, focus_left_tab="slide")

    def _on_pdf_review_select(self, _event=None):
        if self._pdf_review_syncing_selection:
            return
        payload_id = self._selected_pdf_review_payload_id()
        if not payload_id:
            self._refresh_pdf_review_status("")
            return
        self._select_pdf_preview_item(payload_id, focus_left_tab="review")

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

    def _load_docx_into_word_workflow(
        self,
        docx_path: str,
        *,
        document_subject_hint: str = "auto",
        auto_preview: bool = False,
        skip_confirm: bool = False,
    ) -> bool:
        resolved = os.path.abspath(docx_path)
        self.word_path.set(resolved)
        self.output_path.set(os.path.splitext(resolved)[0] + ".pptx")
        self.word_document_subject.set(self._document_subject_label(document_subject_hint))
        self._open_word_workspace_tab()
        if not auto_preview:
            self._set_status(f"已准备好 Word 工作流：{os.path.basename(resolved)}")
            self._refresh_word_flow_ui()
            return True
        return self._start_word_preview_flow(skip_confirm=skip_confirm)

    def _load_word_project_into_preview(self, project, *, docx_path: str, asset_dir: str):
        docx_base = os.path.splitext(os.path.abspath(docx_path))[0]
        default_ppt = self.output_path.get().strip() or docx_base + ".pptx"
        default_manifest = self.pdf_manifest_out.get().strip() or docx_base + "_工程.json"
        self.output_path.set(default_ppt)
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
        if getattr(self, "_pdf_preview_left_tabs", None):
            self._pdf_preview_left_tabs.select(self._pdf_preview_slide_tab)
        self._refresh_word_flow_ui()

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

    def _parse_word_file(self, *, skip_confirm: bool = False) -> bool:
        word_file = self.word_path.get().strip()
        if not word_file:
            messagebox.showwarning("提示", "请先选择 Word 文件")
            return False
        if not os.path.exists(word_file):
            messagebox.showerror("错误", f"文件不存在：{word_file}")
            return False
        if not skip_confirm and not self._confirm_discard_pdf_project_edits("重新解析 Word 结构"):
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
            self._refresh_word_flow_ui()
            return True
        except Exception as exc:
            self._set_status("解析失败")
            self._word_flow_status_var.set("解析失败，请检查 Word 格式或文件内容。")
            messagebox.showerror("解析错误", str(exc))
            return False

    def _start_word_preview_flow(self, *, skip_confirm: bool = False, open_editor: bool = False) -> bool:
        self._open_word_workspace_tab()
        if not self._parse_word_file(skip_confirm=skip_confirm):
            return False
        if open_editor:
            self._open_pdf_preview_workspace()
            self._word_flow_status_var.set("当前已进入逐题编辑工作台；修改完成后回到 Word 工作流直接生成 PPT。")
            self._word_results_hint_var.set("你正在逐题精修；改完后回到 Word 工作流直接生成 PPT。")
        if not self.questions:
            messagebox.showinfo("提示", "未解析到任何题目，请检查 Word 格式")
        return True

    def _parse_preview(self):
        self._start_word_preview_flow()

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
        self._refresh_question_tree()
        self._clear_pdf_preview()
        self._show_pdf_wizard_step(0)
        self.progress["value"] = 0
        self._set_status("已清空")
        self._close_parser()
        self._refresh_word_flow_ui()

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
