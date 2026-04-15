import unittest
from unittest.mock import patch

from core.subject_inference import (
    infer_document_subject,
    infer_pdf_filename_profile,
    infer_subject_diagnostics,
    resolve_objective_section_kinds,
)


class SubjectInferenceTest(unittest.TestCase):
    def test_infer_pdf_filename_profile_detects_single_subject_book(self):
        profile = infer_pdf_filename_profile("行测——判断推理（2000题）.pdf")

        self.assertEqual(profile.form, "single_subject_book")
        self.assertEqual(profile.subject_hint, "reasoning")
        self.assertGreaterEqual(profile.confidence, 0.9)

    def test_infer_pdf_filename_profile_prefers_explicit_subject_over_true_paper_wording(self):
        profile = infer_pdf_filename_profile("资料分析历年真题.pdf")

        self.assertEqual(profile.form, "single_subject_book")
        self.assertEqual(profile.subject_hint, "data")

    def test_infer_pdf_filename_profile_detects_politics_subject_book(self):
        profile = infer_pdf_filename_profile("政治理论题本.pdf")

        self.assertEqual(profile.form, "single_subject_book")
        self.assertEqual(profile.subject_hint, "politics")
        self.assertGreaterEqual(profile.confidence, 0.82)

    def test_infer_pdf_filename_profile_detects_set_paper(self):
        profile = infer_pdf_filename_profile("模拟卷十一.pdf")

        self.assertEqual(profile.form, "set_paper")
        self.assertIsNone(profile.subject_hint)

    def test_infer_pdf_filename_profile_detects_set_paper_for_history_bundle(self):
        profile = infer_pdf_filename_profile("AKA红猪·江苏历年真题.pdf")

        self.assertEqual(profile.form, "set_paper")
        self.assertIsNone(profile.subject_hint)

    def test_infer_document_subject_does_not_misclassify_image_heavy_quant_doc_as_data(self):
        kind, confidence = infer_document_subject(
            [
                "1.某商品上月售价为进价的1.4倍，本月进价下降20%，售价不变，本月销量为多少件？",
                "A.1.3m",
                "B.1.25m",
                "C.1.2m",
                "D.1.15m",
                "2.甲、乙两地相距240千米，两车相向而行，几小时后相遇？",
                "A.4",
                "B.5",
                "C.6",
                "D.8",
                "3.如下图所示，问矩形围栏面积是多少平方米？",
            ],
            image_count=18,
            material_header_count=0,
        )

        self.assertEqual(kind, "quant")
        self.assertGreater(confidence, 0.3)

    def test_infer_document_subject_keeps_data_when_material_headers_exist(self):
        kind, confidence = infer_document_subject(
            [
                "材料一",
                "2024年某市工业增加值同比增长8.3%，服务业增加值同比增长6.1%。",
                "111.根据上述材料，下列说法正确的是：",
                "A.甲",
                "B.乙",
                "C.丙",
                "D.丁",
            ],
            image_count=6,
            material_header_count=1,
        )

        self.assertEqual(kind, "data")
        self.assertGreaterEqual(confidence, 1.0)

    def test_diagnostics_returns_reasoning_with_calibrated_confidence(self):
        diagnostics = infer_subject_diagnostics(
            stem="根据上述定义，下列符合定义的是",
            options=["甲", "乙", "丙", "丁"],
            allow_data=False,
        )

        self.assertEqual(diagnostics.kind, "reasoning")
        self.assertGreater(diagnostics.margin, 1.0)
        self.assertGreaterEqual(diagnostics.confidence, 0.72)
        self.assertLessEqual(diagnostics.confidence, 1.0)
        self.assertEqual(diagnostics.subtype, "定义判断")
        self.assertIn("定义判断结构", diagnostics.matched_signals)

    def test_diagnostics_can_blend_trained_model_prediction(self):
        with patch(
            "core.subject_inference.predict_subject_distribution",
            return_value=type(
                "Prediction",
                (),
                {"probabilities": {"common_sense": 0.08, "reasoning": 0.86, "verbal": 0.06}},
            )(),
        ):
            diagnostics = infer_subject_diagnostics(
                stem="下列说法正确的是",
                options=["甲", "乙", "丙", "丁"],
                image_count=4,
                allow_data=False,
            )

        self.assertEqual(diagnostics.kind, "reasoning")
        self.assertTrue(any("学习模型" in signal for signal in diagnostics.matched_signals))

    def test_diagnostics_uses_knowledge_keywords_for_common_sense(self):
        diagnostics = infer_subject_diagnostics(
            stem="下列关于宪法的说法正确的是",
            options=["甲", "乙", "丙", "丁"],
            allow_data=False,
        )

        self.assertEqual(diagnostics.kind, "common_sense")
        self.assertGreaterEqual(diagnostics.confidence, 0.6)
        self.assertIn(diagnostics.subtype, {"法律常识", None})
        self.assertTrue(any(signal in diagnostics.matched_signals for signal in ("宪法", "下列关于", "下列说法")))

    def test_diagnostics_prefers_reasoning_when_definition_has_case_options(self):
        diagnostics = infer_subject_diagnostics(
            stem="行政指导是指行政机关在其职责范围内，通过非强制方式引导行政相对人作出或者不作出某种行为。下列属于行政指导的是",
            options=[
                "某市市场监管部门向餐饮企业发出规范经营建议书",
                "某法院依法作出生效判决要求立即履行",
                "某公安机关依法对违法人员处以罚款",
                "某税务机关强制扣缴欠税企业税款",
            ],
            allow_data=False,
        )

        self.assertEqual(diagnostics.kind, "reasoning")
        self.assertEqual(diagnostics.subtype, "定义判断")
        self.assertIn("案例型选项", diagnostics.matched_signals)

    def test_diagnostics_keeps_common_sense_for_fact_statement_options(self):
        diagnostics = infer_subject_diagnostics(
            stem="下列关于民法典相关规定的说法正确的是",
            options=[
                "民事主体从事民事活动，应当遵循自愿、公平、诚信原则。",
                "未成年人实施的民事法律行为一律无效。",
                "任何合同只要一方违约即当然无效。",
                "继承开始后，遗嘱继承必然优先于遗赠扶养协议。",
            ],
            allow_data=False,
        )

        self.assertEqual(diagnostics.kind, "common_sense")
        self.assertIn("知识判断型选项", diagnostics.matched_signals)

    def test_diagnostics_boosts_politics_for_policy_combo_options(self):
        diagnostics = infer_subject_diagnostics(
            stem="习近平总书记强调，要牢牢把握高质量发展这个首要任务。关于推进中国式现代化的相关表述，正确的有几项？",
            options=["1项", "2项", "3项", "4项"],
            allow_data=False,
        )

        self.assertEqual(diagnostics.kind, "politics")
        self.assertIn("政策理论判断", diagnostics.matched_signals)
        self.assertIn("组合项选项", diagnostics.matched_signals)

    def test_diagnostics_boosts_common_sense_for_law_title_questions(self):
        diagnostics = infer_subject_diagnostics(
            stem="根据《中华人民共和国监察法》及其实施条例，下列说法错误的是",
            options=[
                "监察机关可以依法采取谈话、讯问等措施。",
                "监察人员办理监察事项，应当回避与本人有利害关系的情形。",
                "所有调查措施均无需履行法定审批程序。",
                "监察工作应当依照法定权限和程序进行。",
            ],
            allow_data=False,
        )

        self.assertEqual(diagnostics.kind, "common_sense")
        self.assertIn("法律条文题", diagnostics.matched_signals)

    def test_diagnostics_uses_policy_combo_shape_for_politics(self):
        diagnostics = infer_subject_diagnostics(
            stem="2025年政府工作报告指出，要扎实推进重点领域改革，创造更加公平、更有活力的市场环境。下列与之相关的具体举措表述正确的有几项？1实施国有企业改革深化提升行动 2纵深推进全国统一大市场建设 3健全民营企业公平参与市场竞争制度 4完善要素市场化配置机制",
            options=["1项", "2项", "3项", "4项"],
            allow_data=False,
        )

        self.assertEqual(diagnostics.kind, "politics")
        self.assertIn("组合项选项", diagnostics.matched_signals)

    def test_diagnostics_does_not_treat_plain_large_numbers_as_combo_options(self):
        diagnostics = infer_subject_diagnostics(
            stem="有一个四位数，各位数字之和为一个两位数的偶数，将后三位上的数字从小到大排列，恰好构成等差数列，则这个四位数可能是：",
            options=["6765", "7675", "8978", "9873"],
            allow_data=False,
        )

        self.assertEqual(diagnostics.kind, "quant")
        self.assertNotIn("组合项选项", diagnostics.matched_signals)

    def test_diagnostics_prefers_common_sense_for_legal_scenarios(self):
        diagnostics = infer_subject_diagnostics(
            stem="根据我国行政诉讼法相关规定，下列案件中，被告正确的是",
            options=[
                "甲市民不服区政府处罚，经市政府复议后仍提起诉讼，市政府为被告。",
                "乙村民认为补偿决定违法，直接起诉村委会，村委会为被告。",
                "丙公司不服行业协会处分，起诉行业协会，行业协会为行政诉讼被告。",
                "丁市民不服街道办作出的治安处罚，起诉派出所，派出所为被告。",
            ],
            allow_data=False,
        )

        self.assertEqual(diagnostics.kind, "common_sense")
        self.assertIn("法条情景判断", diagnostics.matched_signals)

    def test_diagnostics_detects_sentence_expression_subtype(self):
        diagnostics = infer_subject_diagnostics(
            stem="①要提高服务质量 ②群众满意度才能提升 ③基层治理要更精细 ④机制建设也要同步推进 将以上4个句子重新排列，语序正确的是",
            options=["④①③②", "③①④②", "①③④②", "②④①③"],
            allow_data=False,
        )

        self.assertEqual(diagnostics.kind, "verbal")
        self.assertEqual(diagnostics.subtype, "语句表达")
        self.assertIn("语句表达结构", diagnostics.matched_signals)

    def test_diagnostics_keeps_sentence_order_questions_out_of_quant(self):
        diagnostics = infer_subject_diagnostics(
            stem="1原子时能够满足对时间间隔均匀性的要求 2它通过闰秒保证时刻与世界时在一定程度上相符 3世界时的时刻对应太阳在天空中的位置 4协调世界时综合了两者优点 将以上句子重新排列，语序正确的是",
            options=["431625", "415326", "361524", "312654"],
            allow_data=False,
        )

        self.assertEqual(diagnostics.kind, "verbal")
        self.assertIn("排序型选项", diagnostics.matched_signals)

    def test_diagnostics_detects_fill_blank_structure_without_explicit_marker(self):
        diagnostics = infer_subject_diagnostics(
            stem="互联网不是法外之地，当公平交易秩序被践踏时，司法力量必须()维护消费者权益。",
            options=["挺身而出", "杀一儆百", "严惩不贷", "破釜沉舟"],
            allow_data=False,
        )

        self.assertEqual(diagnostics.kind, "verbal")
        self.assertEqual(diagnostics.subtype, "逻辑填空")
        self.assertIn("填空结构", diagnostics.matched_signals)

    def test_diagnostics_prefers_verbal_for_fill_blank_even_with_law_terms(self):
        diagnostics = infer_subject_diagnostics(
            stem="法律自其诞生之日起，便已落后于时代，这是成文法无法避免的()。但是法律必须保证其()，因此对刑法的解释更显重要。",
            options=["局限性", "稳定性", "严厉", "谦抑"],
            allow_data=False,
        )

        self.assertEqual(diagnostics.kind, "verbal")
        self.assertEqual(diagnostics.subtype, "逻辑填空")

    def test_diagnostics_detects_logic_fill_blank_subtype(self):
        diagnostics = infer_subject_diagnostics(
            stem="依次填入画横线部分最恰当的一项是",
            options=["精细", "精致", "精准", "精巧"],
            allow_data=False,
        )

        self.assertEqual(diagnostics.kind, "verbal")
        self.assertEqual(diagnostics.subtype, "逻辑填空")

    def test_diagnostics_detects_logic_fill_blank_with_public_markers(self):
        diagnostics = infer_subject_diagnostics(
            stem="填入横线处最合适的一项是",
            options=["稳步", "稳健", "稳妥", "平稳"],
            allow_data=False,
        )

        self.assertEqual(diagnostics.kind, "verbal")
        self.assertEqual(diagnostics.subtype, "逻辑填空")
        self.assertIn("逻辑填空设问", diagnostics.matched_signals)

    def test_diagnostics_detects_reading_comprehension_with_tail_ask_markers(self):
        diagnostics = infer_subject_diagnostics(
            stem="近年来，基层治理正在从粗放管理走向精细治理，更强调资源协同、空间更新和居民参与。最能概括这段文字的是",
            options=["基层治理只需增加财政投入", "基层治理更强调精细治理", "社区更新应完全市场化", "居民参与会降低治理效率"],
            allow_data=False,
        )

        self.assertEqual(diagnostics.kind, "verbal")
        self.assertEqual(diagnostics.subtype, "片段阅读")
        self.assertIn("片段阅读设问", diagnostics.matched_signals)

    def test_diagnostics_detects_analogy_reasoning_subtype(self):
        diagnostics = infer_subject_diagnostics(
            stem="下列词项关系中，最贴近的一项是",
            options=["工人之于工厂相当于教师之于学校", "医生之于医院相当于护士之于药房", "树木之于森林相当于水滴之于海洋", "铁路之于列车相当于公路之于飞机"],
            allow_data=False,
        )

        self.assertEqual(diagnostics.kind, "reasoning")
        self.assertEqual(diagnostics.subtype, "类比推理")
        self.assertIn("类比推理结构", diagnostics.matched_signals)

    def test_diagnostics_detects_logic_reasoning_subtype(self):
        diagnostics = infer_subject_diagnostics(
            stem="如果甲成立，那么乙成立。只有丙成立，丁才成立。以下哪项最能削弱上述论证？",
            options=["甲", "乙", "丙", "丁"],
            allow_data=False,
        )

        self.assertEqual(diagnostics.kind, "reasoning")
        self.assertEqual(diagnostics.subtype, "逻辑判断")
        self.assertTrue(any(signal in diagnostics.matched_signals for signal in ("如果", "那么", "最能削弱", "逻辑判断结构")))

    def test_diagnostics_does_not_treat_plain_then_connector_as_reasoning_in_quant(self):
        diagnostics = infer_subject_diagnostics(
            stem="某商品上月售价为进价的1.4倍，销售m件。本月该商品进价下降20%，售价不变，销售利润为上月的1.8倍。那么本月销量为多少件？",
            options=["60", "72", "80", "90"],
            allow_data=False,
        )

        self.assertEqual(diagnostics.kind, "quant")
        self.assertNotEqual(diagnostics.subtype, "逻辑判断")

    def test_diagnostics_boosts_quant_for_quantity_prompt_and_numeric_options(self):
        diagnostics = infer_subject_diagnostics(
            stem="某超市设有10个人工收银台。若撤去4个改为6个自助收银台，预计顾客平均排队时间约为多少分钟？",
            options=["10", "12", "15", "18"],
            allow_data=False,
        )

        self.assertEqual(diagnostics.kind, "quant")
        self.assertIn("数量问法", diagnostics.matched_signals)

    def test_diagnostics_detects_constraint_reasoning(self):
        diagnostics = infer_subject_diagnostics(
            stem="某学校新来了三位年轻老师，分别教授生物、物理、英语、政治、历史和数学六门课程中的两门，每人所授科目均不相同。已知：(1)蔡老师年龄最小；(2)孙老师、生物老师和历史老师三人年龄各不相同；(3)朱老师不教英语。下列哪项正确？",
            options=[
                "朱老师教政治和生物",
                "孙老师教英语和数学",
                "朱老师教物理和历史",
                "蔡老师教物理和数学",
            ],
            allow_data=False,
        )

        self.assertEqual(diagnostics.kind, "reasoning")
        self.assertIn("约束条件推理", diagnostics.matched_signals)

    def test_diagnostics_detects_set_relation_reasoning(self):
        diagnostics = infer_subject_diagnostics(
            stem="如果用一个圆来表示词语所指称对象的集合，那么以下哪项中画横线词语之间的关系符合下图？",
            options=[
                "军用船舶与民用船舶",
                "落叶树与水杉",
                "少儿读物与成人图书",
                "哺乳动物与大熊猫",
            ],
            allow_data=False,
        )

        self.assertEqual(diagnostics.kind, "reasoning")
        self.assertIn("集合关系推理", diagnostics.matched_signals)

    def test_diagnostics_detects_grouping_figure_reasoning(self):
        diagnostics = infer_subject_diagnostics(
            stem="把下面六个图形分为两类，使每一类图形都有各自的共同特征或规律，分类正确的一项是：",
            options=["135,246", "123,456", "156,234", "124,356"],
            image_count=6,
            allow_data=False,
        )

        self.assertEqual(diagnostics.kind, "reasoning")
        self.assertIn("分组分类选项", diagnostics.matched_signals)

    def test_diagnostics_detects_relational_reasoning_from_people_and_attributes(self):
        diagnostics = infer_subject_diagnostics(
            stem="小明、小军、小花和小雅分别拿着不同颜色和形状的气球。已知：小花和小雅的气球都不是圆形，也不是蓝色；小军的气球与小雅的颜色和形状都不同。下列说法正确的是：",
            options=[
                "小明拿的是红色圆形气球",
                "小军拿的是蓝色五角星气球",
                "小花拿的是黄色心形气球",
                "小雅拿的是红色心形气球",
            ],
            allow_data=False,
        )

        self.assertEqual(diagnostics.kind, "reasoning")
        self.assertIn("关系约束题干", diagnostics.matched_signals)

    def test_diagnostics_prefers_verbal_for_long_semantic_paragraph_without_explicit_prompt(self):
        diagnostics = infer_subject_diagnostics(
            stem="朱熹在南宋士林中颇有影响，但他对坚持抗金、力主北伐的虞允文颇有微词，说他是“轻薄巧言之士”。显然，朱熹所论有失公允。但即便如此，虞允文对朱熹还是非常敬重。当孝宗皇帝问及朱熹时，虞允文给予高度评价。根据上述材料可知虞允文是一位的人。",
            options=["宽厚包容", "恃才傲物", "优柔寡断", "苛刻保守"],
            allow_data=False,
        )

        self.assertEqual(diagnostics.kind, "verbal")
        self.assertIn("长文段语义题", diagnostics.matched_signals)

    def test_diagnostics_prefers_verbal_for_long_reading_with_according_to_prompt(self):
        diagnostics = infer_subject_diagnostics(
            stem="根据某机构对2000名受访者的调查显示，69.9%的受访者有过在朋友圈购物的经历，39.8%的受访者对微商表示信赖，41.0%的受访者认为朋友圈购物方便了生活。据此，下列说法与原文相符的是：",
            options=[
                "大多数受访者从未在朋友圈购物",
                "所有受访者都信赖微商",
                "部分受访者认为朋友圈购物方便生活",
                "朋友圈购物必然降低消费风险",
            ],
            allow_data=False,
        )

        self.assertEqual(diagnostics.kind, "verbal")
        self.assertIn("阅读判断问法", diagnostics.matched_signals)

    def test_diagnostics_prefers_verbal_for_original_text_comparison_even_with_country_keyword(self):
        diagnostics = infer_subject_diagnostics(
            stem="根据某机构对2000名受访者的调查显示，我国消费者中有69.9%有过在朋友圈购物的经历，39.8%表示信赖微商。据此，下列说法与原文相符的是：",
            options=[
                "微商将逐步代替实体商店",
                "绝大多数微商都不可信",
                "我国对微商没有任何监管制度",
                "在微商购物时需谨慎核实可信度",
            ],
            allow_data=False,
        )

        self.assertEqual(diagnostics.kind, "verbal")
        self.assertTrue(any(signal in diagnostics.matched_signals for signal in ("阅读判断问法", "原文比对问法")))

    def test_diagnostics_prefers_verbal_for_summary_prompt_with_policy_words(self):
        diagnostics = infer_subject_diagnostics(
            stem="有人认为科技强国的主要标志是军事强国，但真正的国家安全也建立在掌握竞争和发展主动权的基础上。产业自主性强了，国家安全才真正有保障。上述材料重在强调：",
            options=[
                "网络信息安全是国家安全的保障",
                "先进科技是国家安全保障的基础",
                "国防科技是科技强国的主要指标",
                "科技强国必须以产业自主为前提",
            ],
            allow_data=False,
        )

        self.assertEqual(diagnostics.kind, "verbal")
        self.assertIn("主旨概括问法", diagnostics.matched_signals)

    def test_diagnostics_prefers_verbal_for_implicit_semantic_fill_blank(self):
        diagnostics = infer_subject_diagnostics(
            stem="“三苏祠”这座祠堂历史悠久，历朝历代都曾扩建。读书人把它奉为圣地，当地人也以此为荣。祠堂古迹在明末一度毁于战火，在清代之后又涅槃新生。如今“三苏祠”蜚声海内外，游人纷纷前来参观。",
            options=["修缮寄托瞻仰", "修理交托视察", "修复寄予景仰", "修葺付托游览"],
            allow_data=False,
        )

        self.assertEqual(diagnostics.kind, "verbal")
        self.assertIn("隐式语义填空", diagnostics.matched_signals)

    def test_diagnostics_prefers_reasoning_for_support_premise_questions(self):
        diagnostics = infer_subject_diagnostics(
            stem="某地区中小学教师中，毕业于师范类院校的女教师多于毕业于非师范类院校的男教师，所以该地区中小学女教师比男教师多。要使上述推理成立，最适合填入画横线位置的是：",
            options=[
                "毕业于师范类院校的教师少于毕业于非师范类院校的教师",
                "毕业于师范类院校的教师多于毕业于非师范类院校的教师",
                "毕业于师范类院校的女教师比毕业于非师范类院校的男教师多",
                "毕业于非师范类院校的女教师比毕业于非师范类院校的男教师多",
            ],
            allow_data=False,
        )

        self.assertEqual(diagnostics.kind, "reasoning")
        self.assertIn("支持前提问法", diagnostics.matched_signals)

    def test_diagnostics_prefers_verbal_for_concept_explanation_reading(self):
        diagnostics = infer_subject_diagnostics(
            stem="“沉默的螺旋”是指优势性意见占据主导位置，其他意见则逐渐从公共图景中消失。网络社会的匿名化又使更多人敢于发表不同见解。根据上述材料，下列说法正确的是：",
            options=[
                "“沉默的螺旋”对网络社会没有影响",
                "网络可在一定程度上减弱“沉默的螺旋”效应",
                "网络的兼容性必然导致意见完全一致",
                "反沉默螺旋只来源于害怕被孤立",
            ],
            allow_data=False,
        )

        self.assertEqual(diagnostics.kind, "verbal")
        self.assertIn("概念阐释型文段", diagnostics.matched_signals)

    def test_diagnostics_prefers_quant_for_graph_calculation_scene(self):
        diagnostics = infer_subject_diagnostics(
            stem="小李父子俩佩戴智能手表跑步锻炼，手表给出如下锻炼时间—配速图线。根据图线分析，在下列哪个时刻两人的跑动距离相同？",
            options=["A时刻", "B时刻", "C时刻", "D时刻"],
            image_count=1,
            allow_data=False,
        )

        self.assertEqual(diagnostics.kind, "quant")
        self.assertIn("图表计算场景", diagnostics.matched_signals)

    def test_diagnostics_prefers_quant_for_cycle_position_problem(self):
        diagnostics = infer_subject_diagnostics(
            stem="某班有48位同学，教室里有6排，每排8个座位。若每周一按规则轮换座位，那么坐在第一排最左边的同学经过多少周后首次回到原位？",
            options=["12周", "24周", "36周", "48周"],
            allow_data=False,
        )

        self.assertEqual(diagnostics.kind, "quant")

    def test_diagnostics_prefers_reasoning_for_science_figure_judgment(self):
        diagnostics = infer_subject_diagnostics(
            stem="催化加氢可制取汽油，转化过程示意图如下。根据图中的信息，下列判断错误的是：",
            options=[
                "反应1的生成物中可能还有水",
                "反应过程中元素化合价不发生改变",
                "小球代表氢原子，大球代表碳原子",
                "该技术有利于实现碳中和目标",
            ],
            allow_data=False,
        )

        self.assertEqual(diagnostics.kind, "reasoning")
        self.assertIn("图示信息判断", diagnostics.matched_signals)

    def test_diagnostics_does_not_use_relational_signal_for_neuroscience_reading(self):
        diagnostics = infer_subject_diagnostics(
            stem="像人类一样，老鼠会在大脑海马体中存储这个世界的心理地图。在老鼠探索周围环境时，不同的地方被一起放电的不同海马体神经元组合所记录和记忆。由上述材料可知，作者最想要传递的信息是：",
            options=[
                "老鼠像人类一样会做梦",
                "老鼠会梦到以前去过的地方",
                "老鼠会梦到想去的地方",
                "老鼠以做梦来构建心理地图",
            ],
            allow_data=False,
        )

        self.assertEqual(diagnostics.kind, "verbal")
        self.assertNotIn("关系约束题干", diagnostics.matched_signals)

    def test_resolve_objective_section_kinds_uses_neighbor_context_for_isolated_unknown(self):
        resolved = resolve_objective_section_kinds(
            default_kind="unknown",
            inferred_pairs=[
                ("verbal", 0.82),
                ("unknown", 0.34),
                ("verbal", 0.79),
            ],
            source_numbers=["21", "22", "23"],
            strong_text_signals=[True, True, True],
        )

        self.assertEqual(resolved, ["verbal", "verbal", "verbal"])

    def test_diagnostics_detects_short_analogy_prompts(self):
        diagnostics = infer_subject_diagnostics(
            stem="羊毛:戳针:羊毛毡",
            options=[
                "纸浆:压模:纸盒",
                "木材:切割:木板",
                "面团:烘烤:面包",
                "泥土:烧制:陶器",
            ],
            allow_data=False,
        )

        self.assertEqual(diagnostics.kind, "reasoning")
        self.assertEqual(diagnostics.subtype, "类比推理")

    def test_diagnostics_detects_table_data_subtype(self):
        diagnostics = infer_subject_diagnostics(
            stem="根据以下资料，回答问题",
            options=["10%", "12%", "14%", "16%"],
            material_header="材料",
            material_text="下表为2021年、2022年和2023年各地区固定资产投资额统计表，表中同比增速如下。",
            image_count=1,
            allow_data=True,
        )

        self.assertEqual(diagnostics.kind, "data")
        self.assertEqual(diagnostics.subtype, "表格型资料分析")
        self.assertTrue(any(signal in diagnostics.matched_signals for signal in ("下表", "表中", "同比", "根据以下资料")))

    def test_diagnostics_detects_mixed_data_subtype(self):
        diagnostics = infer_subject_diagnostics(
            stem="根据以下资料，回答问题",
            options=["10%", "12%", "14%", "16%"],
            material_header="材料",
            material_text="如下表所示为各地区投资额统计表，图中同时给出了同比增速变化趋势。",
            image_count=2,
            allow_data=True,
        )

        self.assertEqual(diagnostics.kind, "data")
        self.assertEqual(diagnostics.subtype, "综合型资料分析")


if __name__ == "__main__":
    unittest.main()
