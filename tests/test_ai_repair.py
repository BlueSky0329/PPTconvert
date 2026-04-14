import unittest

from core.ai_repair import (
    AIRepairService,
    apply_ai_question_patch,
    repair_project_questions,
    repair_question_boundary,
)
from core.project_quality import annotate_project_quality, is_flagged_question
from domain.models import AssetRef, ExamProject, MaterialSet, OptionNode, QuestionNode, Section


class AIRepairTest(unittest.TestCase):
    def test_local_ai_service_repairs_combined_options_and_section(self):
        project = ExamProject(
            title="示例",
            sections=[
                Section(
                    kind="unknown",
                    title="题目列表",
                    questions=[
                        QuestionNode(
                            source_number="1.",
                            stem="下列说法正确的是",
                            options=[
                                OptionNode(letter="A", text="A. 甲 B. 乙"),
                                OptionNode(letter="B", text="C. 丙 D. 丁"),
                            ],
                        )
                    ],
                )
            ],
        )
        annotate_project_quality(project)
        section = project.sections[0]
        question = section.questions[0]
        service = AIRepairService()

        result = service.repair_question(section=section, material=None, question=question)
        changes, subject_changed = apply_ai_question_patch(question, result.patch, section=section, project=project)

        self.assertTrue(result.patch.should_apply)
        self.assertEqual(changes, 7)
        self.assertTrue(subject_changed)
        self.assertEqual(section.kind, "common_sense")
        self.assertEqual(question.source_number, "1")
        self.assertEqual(question.stem, "下列说法正确的是")
        self.assertEqual(question.option_layout, "one_row")
        self.assertEqual(len(question.options), 4)
        self.assertEqual(question.options[3].letter, "D")
        self.assertEqual(question.options[3].text, "丁")

    def test_repair_project_questions_only_repairs_flagged_rows(self):
        project = ExamProject(
            title="示例",
            sections=[
                Section(
                    kind="unknown",
                    title="待确认",
                    questions=[
                        QuestionNode(
                            source_number="1",
                            stem="下列关于法律常识的说法正确的是",
                            options=[
                                OptionNode(letter="A", text="甲"),
                                OptionNode(letter="B", text="乙"),
                                OptionNode(letter="C", text="丙"),
                                OptionNode(letter="D", text="丁"),
                            ],
                        )
                    ],
                ),
                Section(
                    kind="quant",
                    title="数量关系",
                    questions=[
                        QuestionNode(
                            source_number="66",
                            stem="甲乙两地相距 100 千米，求速度。",
                            options=[
                                OptionNode(letter="A", text="20"),
                                OptionNode(letter="B", text="25"),
                                OptionNode(letter="C", text="30"),
                                OptionNode(letter="D", text="40"),
                            ],
                        )
                    ],
                ),
            ],
        )
        annotate_project_quality(project)
        self.assertTrue(is_flagged_question(project.sections[0].questions[0]))
        self.assertFalse(is_flagged_question(project.sections[1].questions[0]))
        service = AIRepairService()

        summary = repair_project_questions(project, service=service, only_flagged=True, limit=5)

        self.assertEqual(summary.attempted_questions, 1)
        self.assertEqual(summary.changed_questions, 1)
        self.assertEqual(summary.subject_changes, 1)
        self.assertEqual(len(summary.errors), 0)
        self.assertEqual(project.sections[0].kind, "common_sense")
        self.assertFalse(is_flagged_question(project.sections[0].questions[0]))

    def test_boundary_repair_moves_spilled_d_option_back_to_previous_question(self):
        project = ExamProject(
            title="串题示例",
            sections=[
                Section(
                    kind="reasoning",
                    title="判断推理",
                    questions=[
                        QuestionNode(
                            source_number="71",
                            stem="根据上述定义，下列符合定义的是",
                            options=[
                                OptionNode(letter="A", text="甲"),
                                OptionNode(letter="B", text="乙"),
                                OptionNode(letter="C", text="丙"),
                            ],
                        ),
                        QuestionNode(
                            source_number="72",
                            stem="D. 丁 72. 如果甲成立，那么乙成立。由此可以推出",
                            options=[
                                OptionNode(letter="A", text="甲"),
                                OptionNode(letter="B", text="乙"),
                                OptionNode(letter="C", text="丙"),
                                OptionNode(letter="D", text="丁"),
                            ],
                        ),
                    ],
                )
            ],
        )
        annotate_project_quality(project)
        section = project.sections[0]
        previous_question = section.questions[0]
        current_question = section.questions[1]

        changes, reason = repair_question_boundary(project, section, None, current_question)

        self.assertEqual(changes, 3)
        self.assertIn("D", reason)
        self.assertEqual(len(previous_question.options), 4)
        self.assertEqual(previous_question.options[3].letter, "D")
        self.assertEqual(previous_question.options[3].text, "丁")
        self.assertEqual(current_question.stem, "如果甲成立，那么乙成立。由此可以推出")

    def test_boundary_repair_moves_multiple_spilled_options_back_to_previous_question(self):
        project = ExamProject(
            title="多选项串题示例",
            sections=[
                Section(
                    kind="reasoning",
                    title="判断推理",
                    questions=[
                        QuestionNode(
                            source_number="71",
                            stem="根据上述定义，下列符合定义的是",
                            options=[
                                OptionNode(letter="A", text="甲"),
                                OptionNode(letter="B", text="乙"),
                            ],
                        ),
                        QuestionNode(
                            source_number="72",
                            stem="C. 丙 D. 丁 72. 如果甲成立，那么乙成立。由此可以推出",
                            options=[
                                OptionNode(letter="A", text="甲"),
                                OptionNode(letter="B", text="乙"),
                                OptionNode(letter="C", text="丙"),
                                OptionNode(letter="D", text="丁"),
                            ],
                        ),
                    ],
                )
            ],
        )
        annotate_project_quality(project)
        section = project.sections[0]
        previous_question = section.questions[0]
        current_question = section.questions[1]

        changes, reason = repair_question_boundary(project, section, None, current_question)

        self.assertEqual(changes, 5)
        self.assertIn("C/D", reason)
        self.assertEqual([option.letter for option in previous_question.options], ["A", "B", "C", "D"])
        self.assertEqual(previous_question.options[2].text, "丙")
        self.assertEqual(previous_question.options[3].text, "丁")
        self.assertEqual(current_question.stem, "如果甲成立，那么乙成立。由此可以推出")

    def test_boundary_repair_recovers_question_number_from_stem_prefix(self):
        project = ExamProject(
            title="跳号示例",
            sections=[
                Section(
                    kind="quant",
                    title="数量关系",
                    questions=[
                        QuestionNode(
                            source_number="71",
                            stem="甲乙两地相距 100 千米",
                            options=[
                                OptionNode(letter="A", text="10"),
                                OptionNode(letter="B", text="20"),
                                OptionNode(letter="C", text="30"),
                                OptionNode(letter="D", text="40"),
                            ],
                        ),
                        QuestionNode(
                            source_number="73",
                            stem="72. 某工程队 5 天完成任务的 1/2，问全部完成需要几天？",
                            options=[
                                OptionNode(letter="A", text="6"),
                                OptionNode(letter="B", text="8"),
                                OptionNode(letter="C", text="10"),
                                OptionNode(letter="D", text="12"),
                            ],
                        ),
                    ],
                )
            ],
        )
        annotate_project_quality(project)
        section = project.sections[0]
        question = section.questions[1]

        changes, reason = repair_question_boundary(project, section, None, question)

        self.assertEqual(changes, 2)
        self.assertIn("题号", reason)
        self.assertEqual(question.source_number, "72")
        self.assertEqual(question.stem, "某工程队 5 天完成任务的 1/2，问全部完成需要几天？")

    def test_boundary_repair_splits_embedded_next_question_from_current_stem(self):
        project = ExamProject(
            title="嵌入下一题示例",
            sections=[
                Section(
                    kind="reasoning",
                    title="判断推理",
                    questions=[
                        QuestionNode(
                            source_number="72",
                            stem="根据上述定义，下列符合定义的是 73. 如果甲成立，那么乙成立。由此可以推出",
                            options=[
                                OptionNode(letter="A", text="甲"),
                                OptionNode(letter="B", text="乙"),
                                OptionNode(letter="C", text="丙"),
                                OptionNode(letter="D", text="丁"),
                            ],
                        ),
                        QuestionNode(
                            source_number="73",
                            stem="",
                            options=[
                                OptionNode(letter="A", text="甲"),
                                OptionNode(letter="B", text="乙"),
                                OptionNode(letter="C", text="丙"),
                                OptionNode(letter="D", text="丁"),
                            ],
                        ),
                    ],
                )
            ],
        )
        annotate_project_quality(project)
        section = project.sections[0]
        current_question = section.questions[0]
        next_question = section.questions[1]

        changes, reason = repair_question_boundary(project, section, None, current_question)

        self.assertEqual(changes, 2)
        self.assertIn("下一题", reason)
        self.assertEqual(current_question.stem, "根据上述定义，下列符合定义的是")
        self.assertEqual(next_question.stem, "如果甲成立，那么乙成立。由此可以推出")

    def test_boundary_repair_splits_embedded_next_question_from_current_option(self):
        project = ExamProject(
            title="选项串入下一题示例",
            sections=[
                Section(
                    kind="reasoning",
                    title="判断推理",
                    questions=[
                        QuestionNode(
                            source_number="72",
                            stem="根据上述定义，下列符合定义的是",
                            options=[
                                OptionNode(letter="A", text="甲"),
                                OptionNode(letter="B", text="乙"),
                                OptionNode(letter="C", text="丙"),
                                OptionNode(letter="D", text="丁 73. 如果甲成立，那么乙成立。由此可以推出"),
                            ],
                        ),
                        QuestionNode(
                            source_number="73",
                            stem="",
                            options=[
                                OptionNode(letter="A", text="甲"),
                                OptionNode(letter="B", text="乙"),
                                OptionNode(letter="C", text="丙"),
                                OptionNode(letter="D", text="丁"),
                            ],
                        ),
                    ],
                )
            ],
        )
        annotate_project_quality(project)
        section = project.sections[0]
        current_question = section.questions[0]
        next_question = section.questions[1]

        changes, reason = repair_question_boundary(project, section, None, current_question)

        self.assertEqual(changes, 2)
        self.assertIn("D 选项", reason)
        self.assertEqual(current_question.options[3].text, "丁")
        self.assertEqual(next_question.stem, "如果甲成立，那么乙成立。由此可以推出")

    def test_local_ai_service_formats_definition_judgment_stem(self):
        project = ExamProject(
            title="定义判断示例",
            sections=[
                Section(
                    kind="reasoning",
                    title="判断推理",
                    questions=[
                        QuestionNode(
                            source_number="71",
                            stem="行政指导是指行政机关在其职责范围内，采用非强制方式引导行政相对人作出或者不作出某种行为的活动 下列属于行政指导的是",
                            options=[
                                OptionNode(letter="A", text="甲"),
                                OptionNode(letter="B", text="乙"),
                                OptionNode(letter="C", text="丙"),
                                OptionNode(letter="D", text="丁"),
                            ],
                        )
                    ],
                )
            ],
        )
        annotate_project_quality(project)
        question = project.sections[0].questions[0]
        service = AIRepairService()

        result = service.repair_question(section=project.sections[0], material=None, question=question)

        self.assertTrue(result.patch.should_apply)
        self.assertEqual(result.patch.stem, "行政指导是指行政机关在其职责范围内，采用非强制方式引导行政相对人作出或者不作出某种行为的活动\n下列属于行政指导的是")
        self.assertIn("定义判断", result.patch.summary)

    def test_local_ai_service_formats_sentence_expression_stem(self):
        project = ExamProject(
            title="语句表达示例",
            sections=[
                Section(
                    kind="verbal",
                    title="言语理解与表达",
                    questions=[
                        QuestionNode(
                            source_number="36",
                            stem="①要提高服务质量 ②群众满意度才能提升 ③基层治理要更精细 ④机制建设也要同步推进 将以上4个句子重新排列，语序正确的是",
                            options=[
                                OptionNode(letter="A", text="④①③②"),
                                OptionNode(letter="B", text="③①④②"),
                                OptionNode(letter="C", text="①③④②"),
                                OptionNode(letter="D", text="②④①③"),
                            ],
                        )
                    ],
                )
            ],
        )
        annotate_project_quality(project)
        question = project.sections[0].questions[0]
        service = AIRepairService()

        result = service.repair_question(section=project.sections[0], material=None, question=question)

        self.assertTrue(result.patch.should_apply)
        self.assertIsNotNone(result.patch.stem)
        self.assertGreaterEqual(result.patch.stem.count("\n"), 4)
        self.assertIn("语句表达", result.patch.summary)

    def test_local_ai_service_formats_reading_comprehension_stem(self):
        project = ExamProject(
            title="片段阅读示例",
            sections=[
                Section(
                    kind="verbal",
                    title="言语理解与表达",
                    questions=[
                        QuestionNode(
                            source_number="42",
                            stem="近年来，城市更新逐渐从大拆大建转向精细治理，更强调公共空间品质、社区参与和历史肌理保护。这段文字意在说明",
                            options=[
                                OptionNode(letter="A", text="城市更新只需要控制成本"),
                                OptionNode(letter="B", text="城市更新更强调精细治理"),
                                OptionNode(letter="C", text="历史街区都应原样保留"),
                                OptionNode(letter="D", text="社区参与并不重要"),
                            ],
                        )
                    ],
                )
            ],
        )
        annotate_project_quality(project)
        question = project.sections[0].questions[0]
        service = AIRepairService()

        result = service.repair_question(section=project.sections[0], material=None, question=question)

        self.assertTrue(result.patch.should_apply)
        self.assertEqual(
            result.patch.stem,
            "近年来，城市更新逐渐从大拆大建转向精细治理，更强调公共空间品质、社区参与和历史肌理保护。\n这段文字意在说明",
        )

    def test_local_ai_service_prefers_grid_for_analogy_reasoning(self):
        project = ExamProject(
            title="类比推理示例",
            sections=[
                Section(
                    kind="reasoning",
                    title="判断推理",
                    questions=[
                        QuestionNode(
                            source_number="81",
                            stem="下列词项关系中，最贴近的一项是",
                            options=[
                                OptionNode(letter="A", text="工人之于工厂相当于教师之于学校"),
                                OptionNode(letter="B", text="医生之于医院相当于护士之于药房"),
                                OptionNode(letter="C", text="树木之于森林相当于水滴之于海洋"),
                                OptionNode(letter="D", text="铁路之于列车相当于公路之于飞机"),
                            ],
                        )
                    ],
                )
            ],
        )
        annotate_project_quality(project)
        question = project.sections[0].questions[0]
        service = AIRepairService()

        result = service.repair_question(section=project.sections[0], material=None, question=question)

        self.assertTrue(result.patch.should_apply)
        self.assertEqual(result.patch.option_layout, "grid")
        self.assertIn("类比推理", result.patch.summary)

    def test_local_ai_service_prefers_one_row_for_logic_fill_blank(self):
        project = ExamProject(
            title="逻辑填空示例",
            sections=[
                Section(
                    kind="verbal",
                    title="言语理解与表达",
                    questions=[
                        QuestionNode(
                            source_number="28",
                            stem="依次填入画横线部分最恰当的一项是",
                            options=[
                                OptionNode(letter="A", text="A. 精细 "),
                                OptionNode(letter="B", text="B. 精致 "),
                                OptionNode(letter="C", text="C. 精准 "),
                                OptionNode(letter="D", text="D. 精巧 "),
                            ],
                            option_layout="list",
                        )
                    ],
                )
            ],
        )
        annotate_project_quality(project)
        question = project.sections[0].questions[0]
        service = AIRepairService()

        result = service.repair_question(section=project.sections[0], material=None, question=question)

        self.assertTrue(result.patch.should_apply)
        self.assertEqual(result.patch.option_layout, "one_row")
        self.assertIn("逻辑填空", result.patch.summary)
        self.assertEqual([patch.text for patch in result.patch.options], ["精细", "精致", "精准", "精巧"])

    def test_local_ai_service_reassigns_graphic_stem_assets_to_options(self):
        project = ExamProject(
            title="图形推理示例",
            sections=[
                Section(
                    kind="reasoning",
                    title="判断推理",
                    questions=[
                        QuestionNode(
                            source_number="71",
                            stem="问号处应填入",
                            stem_assets=[
                                AssetRef(kind="image", path="a.png"),
                                AssetRef(kind="image", path="b.png"),
                                AssetRef(kind="image", path="c.png"),
                                AssetRef(kind="image", path="d.png"),
                            ],
                            options=[
                                OptionNode(letter="A", text="A"),
                                OptionNode(letter="B", text="B"),
                                OptionNode(letter="C", text="C"),
                                OptionNode(letter="D", text="D"),
                            ],
                        )
                    ],
                )
            ],
        )
        annotate_project_quality(project)
        question = project.sections[0].questions[0]
        service = AIRepairService()

        result = service.repair_question(section=project.sections[0], material=None, question=question)
        changes, _subject_changed = apply_ai_question_patch(question, result.patch, section=project.sections[0], project=project)

        self.assertTrue(result.patch.reassign_stem_assets_to_options)
        self.assertGreaterEqual(changes, 4)
        self.assertEqual(question.stem_assets, [])
        self.assertEqual([option.image_path for option in question.options], ["a.png", "b.png", "c.png", "d.png"])

    def test_local_ai_service_moves_embedded_data_intro_back_to_material(self):
        material = MaterialSet(
            material_id="m1",
            header="材料一",
            body="",
            questions=[
                QuestionNode(
                    source_number="101",
                    stem="2023年全市规模以上工业增加值同比增长8.4%，其中制造业增加值占比进一步提升。根据上述资料，下列说法正确的是",
                    options=[
                        OptionNode(letter="A", text="甲"),
                        OptionNode(letter="B", text="乙"),
                        OptionNode(letter="C", text="丙"),
                        OptionNode(letter="D", text="丁"),
                    ],
                )
            ],
        )
        project = ExamProject(
            title="资料分析示例",
            sections=[Section(kind="data", title="资料分析", material_sets=[material])],
        )
        annotate_project_quality(project)
        question = material.questions[0]
        service = AIRepairService()

        result = service.repair_question(section=project.sections[0], material=material, question=question)
        changes, _subject_changed = apply_ai_question_patch(
            question,
            result.patch,
            section=project.sections[0],
            material=material,
            project=project,
        )

        self.assertTrue(result.patch.should_apply)
        self.assertIn("材料说明", result.patch.summary)
        self.assertGreaterEqual(changes, 2)
        self.assertEqual(material.body, "2023年全市规模以上工业增加值同比增长8.4%，其中制造业增加值占比进一步提升。")
        self.assertEqual(question.stem, "根据上述资料，下列说法正确的是")

    def test_local_ai_service_moves_data_stem_assets_back_to_material(self):
        material = MaterialSet(
            material_id="m1",
            header="材料一",
            body="2024年全市工业增加值继续增长，相关图表如下。",
            questions=[
                QuestionNode(
                    source_number="101",
                    stem="根据上述资料，下列说法正确的是",
                    stem_assets=[
                        AssetRef(kind="image", path="chart.png", source_page=4),
                        AssetRef(kind="image", path="table.png", source_page=4),
                    ],
                    options=[
                        OptionNode(letter="A", text="甲"),
                        OptionNode(letter="B", text="乙"),
                        OptionNode(letter="C", text="丙"),
                        OptionNode(letter="D", text="丁"),
                    ],
                )
            ],
        )
        project = ExamProject(
            title="资料分析示例",
            sections=[Section(kind="data", title="资料分析", material_sets=[material])],
        )
        annotate_project_quality(project)
        question = material.questions[0]
        service = AIRepairService()

        result = service.repair_question(section=project.sections[0], material=material, question=question)
        changes, _subject_changed = apply_ai_question_patch(
            question,
            result.patch,
            section=project.sections[0],
            material=material,
            project=project,
        )

        self.assertTrue(result.patch.move_stem_assets_to_material)
        self.assertGreaterEqual(changes, 2)
        self.assertEqual(question.stem_assets, [])
        self.assertEqual([asset.path for asset in material.body_assets], ["chart.png", "table.png"])

    def test_boundary_repair_reconstructs_missing_question_from_previous_stem_tail(self):
        project = ExamProject(
            title="题干尾部漏题示例",
            sections=[
                Section(
                    kind="quant",
                    title="数量关系",
                    questions=[
                        QuestionNode(
                            source_number="71",
                            stem=(
                                "甲乙两地相距 100 千米。"
                                "72. 某工程队 5 天完成任务的 1/2，问全部完成需要几天？ "
                                "A. 6天 B. 8天 C. 10天 D. 12天"
                            ),
                            options=[
                                OptionNode(letter="A", text="10"),
                                OptionNode(letter="B", text="20"),
                                OptionNode(letter="C", text="30"),
                                OptionNode(letter="D", text="40"),
                            ],
                        ),
                        QuestionNode(
                            source_number="73",
                            stem="某商品原价 100 元，现打八折出售，现价是多少？",
                            options=[
                                OptionNode(letter="A", text="60"),
                                OptionNode(letter="B", text="70"),
                                OptionNode(letter="C", text="80"),
                                OptionNode(letter="D", text="90"),
                            ],
                        ),
                    ],
                )
            ],
        )
        annotate_project_quality(project)
        section = project.sections[0]
        current_question = section.questions[1]

        changes, reason = repair_question_boundary(project, section, None, current_question)

        self.assertEqual(changes, 2)
        self.assertIn("上一题题干尾部", reason)
        self.assertEqual([question.source_number for question in section.questions], ["71", "72", "73"])
        self.assertEqual(section.questions[0].stem, "甲乙两地相距 100 千米。")
        inserted_question = section.questions[1]
        self.assertEqual(inserted_question.stem, "某工程队 5 天完成任务的 1/2，问全部完成需要几天？")
        self.assertEqual([option.text for option in inserted_question.options], ["6天", "8天", "10天", "12天"])

    def test_boundary_repair_reconstructs_missing_question_block(self):
        project = ExamProject(
            title="漏题块示例",
            sections=[
                Section(
                    kind="quant",
                    title="数量关系",
                    questions=[
                        QuestionNode(
                            source_number="71",
                            stem="甲乙两地相距 100 千米",
                            options=[
                                OptionNode(letter="A", text="10"),
                                OptionNode(letter="B", text="20"),
                                OptionNode(letter="C", text="30"),
                                OptionNode(letter="D", text="40"),
                            ],
                        ),
                        QuestionNode(
                            source_number="73",
                            stem=(
                                "72. 某工程队 5 天完成任务的 1/2，问全部完成需要几天？ "
                                "A. 6天 B. 8天 C. 10天 D. 12天 "
                                "73. 某商品原价 100 元，现打八折出售，现价是多少？"
                            ),
                            options=[
                                OptionNode(letter="A", text="60"),
                                OptionNode(letter="B", text="70"),
                                OptionNode(letter="C", text="80"),
                                OptionNode(letter="D", text="90"),
                            ],
                        ),
                    ],
                )
            ],
        )
        annotate_project_quality(project)
        section = project.sections[0]
        current_question = section.questions[1]

        changes, reason = repair_question_boundary(project, section, None, current_question)

        self.assertEqual(changes, 2)
        self.assertIn("漏掉的第 72 题", reason)
        self.assertEqual([question.source_number for question in section.questions], ["71", "72", "73"])
        inserted_question = section.questions[1]
        self.assertEqual(inserted_question.stem, "某工程队 5 天完成任务的 1/2，问全部完成需要几天？")
        self.assertEqual([option.text for option in inserted_question.options], ["6天", "8天", "10天", "12天"])
        self.assertEqual(section.questions[2].stem, "某商品原价 100 元，现打八折出售，现价是多少？")

    def test_boundary_repair_reconstructs_missing_question_from_previous_option_tail(self):
        project = ExamProject(
            title="选项尾部漏题示例",
            sections=[
                Section(
                    kind="quant",
                    title="数量关系",
                    questions=[
                        QuestionNode(
                            source_number="71",
                            stem="甲乙两地相距 100 千米",
                            options=[
                                OptionNode(letter="A", text="10"),
                                OptionNode(letter="B", text="20"),
                                OptionNode(letter="C", text="30"),
                                OptionNode(
                                    letter="D",
                                    text="40 72. 某工程队 5 天完成任务的 1/2，问全部完成需要几天？ A. 6天 B. 8天 C. 10天 D. 12天",
                                ),
                            ],
                        ),
                        QuestionNode(
                            source_number="73",
                            stem="某商品原价 100 元，现打八折出售，现价是多少？",
                            options=[
                                OptionNode(letter="A", text="60"),
                                OptionNode(letter="B", text="70"),
                                OptionNode(letter="C", text="80"),
                                OptionNode(letter="D", text="90"),
                            ],
                        ),
                    ],
                )
            ],
        )
        annotate_project_quality(project)
        section = project.sections[0]
        current_question = section.questions[1]

        changes, reason = repair_question_boundary(project, section, None, current_question)

        self.assertEqual(changes, 2)
        self.assertIn("上一题 D 选项尾部", reason)
        self.assertEqual([question.source_number for question in section.questions], ["71", "72", "73"])
        self.assertEqual(section.questions[0].options[3].text, "40")
        inserted_question = section.questions[1]
        self.assertEqual(inserted_question.stem, "某工程队 5 天完成任务的 1/2，问全部完成需要几天？")
        self.assertEqual([option.text for option in inserted_question.options], ["6天", "8天", "10天", "12天"])

    def test_boundary_repair_reconstructs_missing_question_after_current_from_stem_tail(self):
        project = ExamProject(
            title="当前题题干尾部漏题示例",
            sections=[
                Section(
                    kind="quant",
                    title="数量关系",
                    questions=[
                        QuestionNode(
                            source_number="72",
                            stem=(
                                "甲乙两地相距 100 千米。"
                                "73. 某工程队 5 天完成任务的 1/2，问全部完成需要几天？ "
                                "A. 6天 B. 8天 C. 10天 D. 12天 "
                                "74. 某商品原价 100 元，现打八折出售，现价是多少？"
                            ),
                            options=[
                                OptionNode(letter="A", text="10"),
                                OptionNode(letter="B", text="20"),
                                OptionNode(letter="C", text="30"),
                                OptionNode(letter="D", text="40"),
                            ],
                        ),
                        QuestionNode(
                            source_number="74",
                            stem="某商品原价 100 元，现打八折出售，现价是多少？",
                            options=[
                                OptionNode(letter="A", text="60"),
                                OptionNode(letter="B", text="70"),
                                OptionNode(letter="C", text="80"),
                                OptionNode(letter="D", text="90"),
                            ],
                        ),
                    ],
                )
            ],
        )
        annotate_project_quality(project)
        section = project.sections[0]
        current_question = section.questions[0]

        changes, reason = repair_question_boundary(project, section, None, current_question)

        self.assertEqual(changes, 2)
        self.assertIn("当前题题干尾部", reason)
        self.assertEqual([question.source_number for question in section.questions], ["72", "73", "74"])
        self.assertEqual(section.questions[0].stem, "甲乙两地相距 100 千米。")
        inserted_question = section.questions[1]
        self.assertEqual(inserted_question.stem, "某工程队 5 天完成任务的 1/2，问全部完成需要几天？")
        self.assertEqual([option.text for option in inserted_question.options], ["6天", "8天", "10天", "12天"])

    def test_boundary_repair_reconstructs_missing_question_after_current_from_option_tail(self):
        project = ExamProject(
            title="当前题选项尾部漏题示例",
            sections=[
                Section(
                    kind="quant",
                    title="数量关系",
                    questions=[
                        QuestionNode(
                            source_number="72",
                            stem="甲乙两地相距 100 千米。",
                            options=[
                                OptionNode(letter="A", text="10"),
                                OptionNode(letter="B", text="20"),
                                OptionNode(letter="C", text="30"),
                                OptionNode(
                                    letter="D",
                                    text=(
                                        "40 73. 某工程队 5 天完成任务的 1/2，问全部完成需要几天？ "
                                        "A. 6天 B. 8天 C. 10天 D. 12天 "
                                        "74. 某商品原价 100 元，现打八折出售，现价是多少？"
                                    ),
                                ),
                            ],
                        ),
                        QuestionNode(
                            source_number="74",
                            stem="某商品原价 100 元，现打八折出售，现价是多少？",
                            options=[
                                OptionNode(letter="A", text="60"),
                                OptionNode(letter="B", text="70"),
                                OptionNode(letter="C", text="80"),
                                OptionNode(letter="D", text="90"),
                            ],
                        ),
                    ],
                )
            ],
        )
        annotate_project_quality(project)
        section = project.sections[0]
        current_question = section.questions[0]

        changes, reason = repair_question_boundary(project, section, None, current_question)

        self.assertEqual(changes, 2)
        self.assertIn("当前题 D 选项尾部", reason)
        self.assertEqual([question.source_number for question in section.questions], ["72", "73", "74"])
        self.assertEqual(section.questions[0].options[3].text, "40")
        inserted_question = section.questions[1]
        self.assertEqual(inserted_question.stem, "某工程队 5 天完成任务的 1/2，问全部完成需要几天？")
        self.assertEqual([option.text for option in inserted_question.options], ["6天", "8天", "10天", "12天"])

    def test_boundary_repair_reconstructed_question_merges_cross_page_numbers(self):
        project = ExamProject(
            title="跨页漏题示例",
            sections=[
                Section(
                    kind="quant",
                    title="数量关系",
                    questions=[
                        QuestionNode(
                            source_number="72",
                            stem="甲乙两地相距 100 千米。",
                            options=[
                                OptionNode(letter="A", text="10"),
                                OptionNode(letter="B", text="20"),
                                OptionNode(letter="C", text="30"),
                                OptionNode(
                                    letter="D",
                                    text=(
                                        "40 73. 某工程队 5 天完成任务的 1/2，问全部完成需要几天？ "
                                        "A. 6天 B. 8天 C. 10天 D. 12天 "
                                        "74. 某商品原价 100 元，现打八折出售，现价是多少？"
                                    ),
                                ),
                            ],
                            page_numbers=[2],
                        ),
                        QuestionNode(
                            source_number="74",
                            stem="某商品原价 100 元，现打八折出售，现价是多少？",
                            options=[
                                OptionNode(letter="A", text="60"),
                                OptionNode(letter="B", text="70"),
                                OptionNode(letter="C", text="80"),
                                OptionNode(letter="D", text="90"),
                            ],
                            page_numbers=[3],
                        ),
                    ],
                )
            ],
        )
        annotate_project_quality(project)
        section = project.sections[0]

        changes, reason = repair_question_boundary(project, section, None, section.questions[0])

        self.assertEqual(changes, 2)
        self.assertIn("当前题 D 选项尾部", reason)
        self.assertEqual(section.questions[1].page_numbers, [2, 3])

    def test_boundary_repair_strips_prompt_suffix_from_current_stem(self):
        project = ExamProject(
            title="提示语污染示例",
            sections=[
                Section(
                    kind="quant",
                    title="数量关系",
                    questions=[
                        QuestionNode(
                            source_number="72",
                            stem="甲乙两地相距 100 千米。请根据以下材料回答下列问题。",
                            options=[
                                OptionNode(letter="A", text="10"),
                                OptionNode(letter="B", text="20"),
                                OptionNode(letter="C", text="30"),
                                OptionNode(letter="D", text="40"),
                            ],
                        ),
                        QuestionNode(
                            source_number="73",
                            stem="某工程队 5 天完成任务的 1/2，问全部完成需要几天？",
                            options=[
                                OptionNode(letter="A", text="6"),
                                OptionNode(letter="B", text="8"),
                                OptionNode(letter="C", text="10"),
                                OptionNode(letter="D", text="12"),
                            ],
                        ),
                    ],
                )
            ],
        )
        annotate_project_quality(project)
        section = project.sections[0]

        changes, reason = repair_question_boundary(project, section, None, section.questions[0])

        self.assertEqual(changes, 1)
        self.assertIn("提示语", reason)
        self.assertEqual(section.questions[0].stem, "甲乙两地相距 100 千米。")
        self.assertEqual(section.questions[1].stem, "某工程队 5 天完成任务的 1/2，问全部完成需要几天？")


if __name__ == "__main__":
    unittest.main()
