import unittest
import tempfile
from types import SimpleNamespace

from core.project_quality import annotate_project_quality
from PIL import Image

from core.ppt_generator import PPTGenerator
from domain.models import AssetRef, ExamProject, MaterialSet, OptionNode, QuestionNode, Section
from gui.app import PPTConvertApp


class GuiPreviewBehaviorTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = PPTConvertApp()
        cls.app.root.withdraw()

    @classmethod
    def tearDownClass(cls):
        try:
            cls.app.root.update_idletasks()
            cls.app.root.update()
        except Exception:
            pass
        cls.app.root.destroy()

    def _current_left_tab(self) -> str:
        selected = self.app._pdf_preview_left_tabs.select()
        if selected == str(self.app._pdf_preview_structure_tab):
            return "structure"
        if selected == str(self.app._pdf_preview_review_tab):
            return "review"
        if selected == str(self.app._pdf_preview_slide_tab):
            return "slide"
        return selected

    def _build_flagged_project(self) -> ExamProject:
        project = ExamProject(
            title="示例工程",
            sections=[
                Section(
                    kind="common_sense",
                    title="常识判断",
                    questions=[
                        QuestionNode(
                            source_number="1",
                            stem="根据上述定义，下列符合定义的是",
                            options=[
                                OptionNode("A", "甲"),
                                OptionNode("B", "乙"),
                                OptionNode("C", "丙"),
                                OptionNode("D", "丁"),
                            ],
                        ),
                        QuestionNode(
                            source_number="2",
                            stem="下列关于宪法的说法正确的是",
                            options=[
                                OptionNode("A", "甲"),
                                OptionNode("B", "乙"),
                                OptionNode("C", "丙"),
                                OptionNode("D", "丁"),
                            ],
                        ),
                    ],
                )
            ],
        )
        annotate_project_quality(project)
        return project

    def _make_temp_image(self, directory: str, name: str, size=(180, 120), color=(80, 140, 220)) -> str:
        path = f"{directory}/{name}".replace("/", "\\")
        Image.new("RGB", size, color).save(path)
        return path

    def _build_data_project(self, temp_dir: str) -> tuple[ExamProject, QuestionNode]:
        material_image_1 = self._make_temp_image(temp_dir, "material_1.png", size=(220, 120), color=(220, 180, 90))
        material_image_2 = self._make_temp_image(temp_dir, "material_2.png", size=(200, 140), color=(120, 200, 140))
        stem_image = self._make_temp_image(temp_dir, "stem.png", size=(160, 110), color=(180, 120, 220))
        question = QuestionNode(
            source_number="101",
            stem="2024年全国地表水中未入渗补给地下水的资源量比入渗补给地下水的多：",
            options=[
                OptionNode("A", "不到1倍"),
                OptionNode("B", "1倍以上"),
                OptionNode("C", "2倍以上"),
                OptionNode("D", "3倍以上"),
            ],
            stem_assets=[AssetRef(kind="image", path=stem_image)],
        )
        material = MaterialSet(
            material_id="m1",
            header="材料一",
            body="2024年，全国平均年降水量为717.7毫米。\n地表水资源量为31655.4亿立方米。",
            body_assets=[
                AssetRef(kind="image", path=material_image_1),
                AssetRef(kind="image", path=material_image_2),
            ],
            questions=[question],
        )
        project = ExamProject(
            title="资料分析示例",
            sections=[
                Section(
                    kind="data",
                    title="资料分析",
                    material_sets=[material],
                )
            ],
        )
        annotate_project_quality(project)
        return project, question

    def _select_first_question(self, project: ExamProject):
        self.app._populate_pdf_preview(project)
        self.app.root.update_idletasks()
        self.app.root.update()

        section_item = self.app.pdf_tree.get_children()[0]
        question_item = next(
            child
            for child in self.app.pdf_tree.get_children(section_item)
            if self.app._pdf_preview_payloads.get(child, {}).get("kind") == "question"
        )

        self.app.pdf_tree.selection_set(question_item)
        self.app.pdf_tree.focus(question_item)
        self.app._on_pdf_preview_select()
        self.app.root.update_idletasks()
        self.app.root.update()
        return project.sections[0].questions[0]

    def _select_first_data_question(self, project: ExamProject):
        self.app._populate_pdf_preview(project)
        self.app.root.update_idletasks()
        self.app.root.update()

        section_item = self.app.pdf_tree.get_children()[0]
        material_item = next(
            child
            for child in self.app.pdf_tree.get_children(section_item)
            if self.app._pdf_preview_payloads.get(child, {}).get("kind") == "material"
        )
        question_item = next(
            child
            for child in self.app.pdf_tree.get_children(material_item)
            if self.app._pdf_preview_payloads.get(child, {}).get("kind") == "question"
        )

        self.app.pdf_tree.selection_set(question_item)
        self.app.pdf_tree.focus(question_item)
        self.app._on_pdf_preview_select()
        self.app.root.update_idletasks()
        self.app.root.update()
        return project.sections[0].material_sets[0].questions[0]

    def test_structure_question_selection_keeps_structure_tab_active(self):
        project = self._build_flagged_project()
        self.app._populate_pdf_preview(project)
        self.app.root.update_idletasks()
        self.app.root.update()

        section_item = self.app.pdf_tree.get_children()[0]
        question_item = next(
            child
            for child in self.app.pdf_tree.get_children(section_item)
            if self.app._pdf_preview_payloads.get(child, {}).get("kind") == "question"
        )

        self.app.pdf_tree.selection_set(question_item)
        self.app.pdf_tree.focus(question_item)
        self.app._on_pdf_preview_select()
        self.app.root.update_idletasks()
        self.app.root.update()

        self.assertEqual(self._current_left_tab(), "structure")
        self.assertEqual(self.app._selected_pdf_item_id(), question_item)

    def test_structure_question_selection_syncs_slide_selection(self):
        project = self._build_flagged_project()
        self.app._populate_pdf_preview(project)
        self.app.root.update_idletasks()
        self.app.root.update()

        section_item = self.app.pdf_tree.get_children()[0]
        question_item = next(
            child
            for child in self.app.pdf_tree.get_children(section_item)
            if self.app._pdf_preview_payloads.get(child, {}).get("kind") == "question"
        )

        self.app.pdf_tree.selection_set(question_item)
        self.app.pdf_tree.focus(question_item)
        self.app._on_pdf_preview_select()
        self.app.root.update_idletasks()
        self.app.root.update()

        self.assertEqual(self.app._selected_pdf_slide_payload_id(), question_item)

    def test_ai_suggestion_mentions_explicit_subject_name(self):
        project = self._build_flagged_project()
        self.app._populate_pdf_preview(project)
        self.app.root.update_idletasks()
        self.app.root.update()

        review_item = self.app._pdf_review_tree.get_children()[0]
        self.app._pdf_review_tree.selection_set(review_item)
        self.app._pdf_review_tree.focus(review_item)
        self.app._on_pdf_review_select()
        self.app.root.update_idletasks()
        self.app.root.update()

        suggestion = self.app._pdf_ai_suggestion_var.get()
        self.assertIn("判断推理", suggestion)

    def test_ai_strategy_box_shows_rule_mode_summary(self):
        project = self._build_flagged_project()
        self.app._populate_pdf_preview(project)
        self.app.root.update_idletasks()
        self.app.root.update()

        review_item = self.app._pdf_review_tree.get_children()[0]
        self.app._pdf_review_tree.selection_set(review_item)
        self.app._pdf_review_tree.focus(review_item)
        self.app._on_pdf_review_select()
        self.app.root.update_idletasks()
        self.app.root.update()

        strategy = self.app._pdf_ai_strategy_var.get()
        self.assertIn("当前模式：规则优先", strategy)
        self.assertIn("轨迹", strategy)

    def test_ocr_buttons_are_enabled_after_project_load(self):
        project = self._build_flagged_project()
        self.app._populate_pdf_preview(project)
        self.app.root.update_idletasks()
        self.app.root.update()

        self.assertEqual(str(self.app._pdf_ocr_diagnose_btn.cget("state")), "normal")
        self.assertEqual(str(self.app._pdf_ocr_repair_btn.cget("state")), "normal")

    def test_data_preview_renders_export_stem_text(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            project, question = self._build_data_project(temp_dir)
            self._select_first_data_question(project)
            self.app._render_pdf_question_editor_preview()

            render_question = self.app._build_pdf_preview_render_question(question)
            expected_text = PPTGenerator._stem_text_for_question(render_question)
            canvas_texts = [
                str(self.app._pdf_question_preview_canvas.itemcget(item, "text") or "")
                for item in self.app._pdf_question_preview_canvas.find_all()
                if self.app._pdf_question_preview_canvas.type(item) == "text"
            ]

            self.assertTrue(any("材料一" in text for text in canvas_texts))
            self.assertTrue(any("2024年，全国平均年降水量为717.7毫米。" in text for text in canvas_texts))
            self.assertTrue(any("101." in text and "未入渗补给地下水" in text for text in canvas_texts))
            self.assertIn("材料一", expected_text)

    def test_data_preview_renders_all_image_assets(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            project, _question = self._build_data_project(temp_dir)
            self._select_first_data_question(project)
            self.app._render_pdf_question_editor_preview()

            self.assertGreaterEqual(len(self.app._pdf_question_preview_photos), 3)

    def test_stem_focusout_records_original_text_in_repair_log(self):
        project = self._build_flagged_project()
        question = self._select_first_question(project)

        self.app._pdf_question_stem_editor.delete("1.0", "end")
        self.app._pdf_question_stem_editor.insert("1.0", "新的题干")
        self.app._on_pdf_question_stem_change()
        self.app._on_pdf_question_stem_change(SimpleNamespace(type="FocusOut"))

        self.assertEqual(question.stem, "新的题干")
        self.assertEqual(len(project.repair_log), 1)
        entry = project.repair_log[0]
        self.assertEqual(entry.action, "update_question_stem")
        self.assertIn("题干：根据上述定义，下列符合定义的是", entry.before_state["state_record"]["text"])
        self.assertIn("题干：新的题干", entry.after_state["state_record"]["text"])

    def test_option_focusout_records_original_option_text_in_repair_log(self):
        project = self._build_flagged_project()
        question = self._select_first_question(project)

        option_editor = self.app._pdf_option_editors["A"]
        option_editor.delete("1.0", "end")
        option_editor.insert("1.0", "新的甲")
        self.app._on_pdf_option_text_change("A")
        self.app._on_pdf_option_text_change("A", SimpleNamespace(type="FocusOut"))

        self.assertEqual(question.options[0].text, "新的甲")
        self.assertEqual(len(project.repair_log), 1)
        entry = project.repair_log[0]
        self.assertEqual(entry.action, "update_option_text")
        self.assertIn("A. 甲", entry.before_state["state_record"]["text"])
        self.assertIn("A. 新的甲", entry.after_state["state_record"]["text"])

    def test_dragging_preview_block_updates_question_ppt_layout(self):
        project = self._build_flagged_project()
        question = self._select_first_question(project)
        canvas = self.app._pdf_question_preview_canvas
        canvas.configure(width=720, height=360)
        self.app.root.update_idletasks()
        self.app.root.update()
        self.app._render_pdf_question_editor_preview()

        stem_rect = self.app._pdf_preview_rects["stem"]
        start_event = SimpleNamespace(x=int(stem_rect[0] + 24), y=int(stem_rect[1] + 28))
        drag_event = SimpleNamespace(x=int(stem_rect[0] + 56), y=int(stem_rect[1] + 46))

        self.app._on_pdf_question_preview_press(start_event)
        self.app._on_pdf_question_preview_drag(drag_event)
        self.app._on_pdf_question_preview_release(drag_event)

        self.assertIn("stem", question.ppt_layout)
        self.assertGreater(question.ppt_layout["stem"]["x"], 0.05)
        self.assertEqual(project.repair_log[-1].action, "set_question_ppt_layout")

    def test_dragging_option_block_updates_question_ppt_layout(self):
        project = self._build_flagged_project()
        question = self._select_first_question(project)
        canvas = self.app._pdf_question_preview_canvas
        canvas.configure(width=720, height=360)
        self.app.root.update_idletasks()
        self.app.root.update()
        self.app._render_pdf_question_editor_preview()

        option_rect = self.app._pdf_preview_rects["option_a"]
        start_event = SimpleNamespace(x=int(option_rect[0] + 18), y=int(option_rect[1] + 18))
        drag_event = SimpleNamespace(x=int(option_rect[0] + 42), y=int(option_rect[1] + 32))

        self.app._on_pdf_question_preview_press(start_event)
        self.app._on_pdf_question_preview_drag(drag_event)
        self.app._on_pdf_question_preview_release(drag_event)

        self.assertIn("option_a", question.ppt_layout)
        self.assertEqual(project.repair_log[-1].action, "set_question_ppt_layout")

    def test_moving_options_region_keeps_option_override_relative(self):
        project = self._build_flagged_project()
        question = self._select_first_question(project)
        canvas = self.app._pdf_question_preview_canvas
        canvas.configure(width=720, height=360)
        self.app.root.update_idletasks()
        self.app.root.update()
        self.app._render_pdf_question_editor_preview()

        option_rect = self.app._pdf_preview_rects["option_a"]
        press_option = SimpleNamespace(x=int(option_rect[0] + 16), y=int(option_rect[1] + 16))
        drag_option = SimpleNamespace(x=int(option_rect[0] + 32), y=int(option_rect[1] + 24))
        self.app._on_pdf_question_preview_press(press_option)
        self.app._on_pdf_question_preview_drag(drag_option)
        self.app._on_pdf_question_preview_release(drag_option)

        option_override_before = dict(question.ppt_layout["option_a"])
        region_rect = self.app._pdf_preview_rects["options"]
        press_region = SimpleNamespace(x=int(region_rect[0] + 6), y=int(region_rect[1] + 10))
        drag_region = SimpleNamespace(x=int(region_rect[0] + 36), y=int(region_rect[1] + 34))
        self.app._on_pdf_question_preview_press(press_region)
        self.app._on_pdf_question_preview_drag(drag_region)
        self.app._on_pdf_question_preview_release(drag_region)

        option_override_after = question.ppt_layout["option_a"]
        self.assertNotEqual(option_override_before["x"], option_override_after["x"])
        self.assertNotEqual(option_override_before["y"], option_override_after["y"])

    def test_align_action_updates_selected_block_layout(self):
        project = self._build_flagged_project()
        question = self._select_first_question(project)
        self.app._pdf_preview_selected_block = "stem"
        self.app._render_pdf_question_editor_preview()

        self.app._align_selected_pdf_preview_block("right")

        self.assertIn("stem", question.ppt_layout)
        self.assertEqual(project.repair_log[-1].action, "align_question_ppt_layout")

    def test_keyboard_nudge_updates_selected_block_layout(self):
        project = self._build_flagged_project()
        question = self._select_first_question(project)
        self.app._pdf_preview_selected_block = "stem"
        self.app._render_pdf_question_editor_preview()

        self.app._on_pdf_question_preview_nudge(SimpleNamespace(keysym="Right", state=0))

        self.assertIn("stem", question.ppt_layout)
        self.assertEqual(project.repair_log[-1].action, "nudge_question_ppt_layout")

    def test_numeric_layout_fields_apply_selected_block_layout(self):
        project = self._build_flagged_project()
        question = self._select_first_question(project)
        self.app._pdf_preview_selected_block = "stem"
        self.app._render_pdf_question_editor_preview()

        self.app._pdf_layout_x_var.set(1.2)
        self.app._pdf_layout_y_var.set(0.9)
        self.app._pdf_layout_w_var.set(5.4)
        self.app._pdf_layout_h_var.set(1.6)
        self.app._apply_pdf_preview_numeric_layout()

        self.assertIn("stem", question.ppt_layout)
        self.assertAlmostEqual(question.ppt_layout["stem"]["x"], 1.2 / 13.333, places=3)
        self.assertEqual(project.repair_log[-1].action, "set_question_ppt_layout_fields")

    def test_fill_width_action_updates_selected_block_layout(self):
        project = self._build_flagged_project()
        question = self._select_first_question(project)
        self.app._pdf_preview_selected_block = "stem"
        self.app._render_pdf_question_editor_preview()

        self.app._align_selected_pdf_preview_block("fill_width")

        self.assertIn("stem", question.ppt_layout)
        self.assertAlmostEqual(question.ppt_layout["stem"]["x"], 0.0, places=3)
        self.assertGreater(question.ppt_layout["stem"]["w"], 0.95)
        self.assertEqual(project.repair_log[-1].action, "align_question_ppt_layout")


if __name__ == "__main__":
    unittest.main()
