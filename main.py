import argparse
import logging
import os
import sys


def configure_logging(level_name: str) -> None:
    level = getattr(logging, (level_name or "WARNING").upper(), logging.WARNING)
    logging.basicConfig(level=level, format="%(levelname)s: %(message)s")


def maybe_run_ai_repair(args, project) -> None:
    if not getattr(args, "ai_repair", False):
        return

    from core.ai_repair import AIRepairService, repair_project_questions

    service = AIRepairService()

    summary = repair_project_questions(
        project,
        service=service,
        only_flagged=not getattr(args, "ai_all_questions", False),
        limit=args.ai_limit,
    )
    print(
        "AI修复："
        f"处理 {summary.attempted_questions} 题，"
        f"写回 {summary.changed_questions} 题，"
        f"共 {summary.total_field_changes} 处字段"
        + (f"，科目调整 {summary.subject_changes} 处" if summary.subject_changes else "")
    )
    if summary.errors:
        print(f"AI修复：另有 {len(summary.errors)} 题未成功")
        for line in summary.errors[:5]:
            print(f"  - {line}")


def run_gui():
    from gui.app import PPTConvertApp

    app = PPTConvertApp()
    app.run()


def run_cli(args):
    from core.project_quality import annotate_project_quality
    from core.ppt_generator import PPTConfig
    from pptx.util import Pt
    from workflows.project_flow import build_word_project, export_project_outputs

    if not os.path.exists(args.input):
        print(f"错误：文件不存在 - {args.input}")
        sys.exit(1)

    base_path = os.path.splitext(args.input)[0]
    ppt_output = args.output or base_path + ".pptx"
    docx_output = args.docx_output
    manifest_output = args.manifest_output
    document_subject_hint = None if args.document_subject == "auto" else args.document_subject

    print(f"正在解析：{args.input}")
    project, _questions, asset_dir = build_word_project(
        args.input,
        document_subject_hint=document_subject_hint,
    )
    quality = annotate_project_quality(project)
    print(f"共解析到 {project.question_count} 道题目")
    print(
        "AI质检："
        f"待确认 {quality.flagged_questions} 题，"
        f"高风险 {quality.severe_questions} 题，"
        f"共 {quality.total_issue_count} 项提醒"
    )

    if not project.question_count:
        print("未找到任何题目，请检查 Word 文档格式。")
        sys.exit(1)

    maybe_run_ai_repair(args, project)
    quality = annotate_project_quality(project)
    if getattr(args, "ai_repair", False):
        print(
            "AI修复后质检："
            f"待确认 {quality.flagged_questions} 题，"
            f"高风险 {quality.severe_questions} 题，"
            f"共 {quality.total_issue_count} 项提醒"
        )

    ppt_config = PPTConfig()
    if args.layout:
        ppt_config.option_layout = args.layout
    if args.font_size:
        ppt_config.stem_font_size = Pt(args.font_size)
        ppt_config.option_font_size = Pt(max(args.font_size - 2, 1))

    outputs = export_project_outputs(
        project,
        asset_dir=asset_dir,
        docx_output=docx_output,
        ppt_output=ppt_output,
        manifest_output=manifest_output,
        template_path=args.template or None,
        ppt_config=ppt_config,
    )

    print(f"素材目录：{outputs.asset_dir}")
    if outputs.docx_path:
        print(f"题本 Word：{outputs.docx_path}")
    if outputs.pptx_path:
        print(f"授课 PPT：{outputs.pptx_path}")
    if outputs.manifest_path:
        print(f"工程清单：{outputs.manifest_path}")


def run_pdf_workflow(args):
    from core.project_quality import annotate_project_quality
    from core.ppt_generator import PPTConfig
    from pptx.util import Pt
    from workflows.project_flow import process_pdf_project

    if not os.path.exists(args.pdf_input):
        print(f"错误：文件不存在 - {args.pdf_input}")
        sys.exit(1)

    base_path = os.path.splitext(args.pdf_input)[0]
    docx_output = args.docx_output
    ppt_output = args.ppt_output or args.output
    manifest_output = args.manifest_output

    if not any([docx_output, ppt_output, manifest_output]):
        docx_output = base_path + "_题本.docx"

    ppt_config = None
    if ppt_output:
        ppt_config = PPTConfig()
        if args.layout:
            ppt_config.option_layout = args.layout
        if args.font_size:
            ppt_config.stem_font_size = Pt(args.font_size)
            ppt_config.option_font_size = Pt(max(args.font_size - 2, 1))

    document_subject_hint = None if args.document_subject == "auto" else args.document_subject
    print(f"正在导入 PDF：{args.pdf_input}")
    project, outputs = process_pdf_project(
        args.pdf_input,
        mode=args.subject,
        question_range_spec=args.question_range or "",
        docx_output=docx_output,
        ppt_output=ppt_output,
        manifest_output=manifest_output,
        template_path=args.template or None,
        ppt_config=ppt_config,
        document_subject_hint=document_subject_hint,
    )
    quality = annotate_project_quality(project)

    maybe_run_ai_repair(args, project)
    quality = annotate_project_quality(project)

    print(f"共整理到 {project.question_count} 道题目")
    print(
        "AI质检："
        f"待确认 {quality.flagged_questions} 题，"
        f"高风险 {quality.severe_questions} 题，"
        f"共 {quality.total_issue_count} 项提醒"
    )
    print(f"素材目录：{outputs.asset_dir}")
    if outputs.docx_path:
        print(f"题本 Word：{outputs.docx_path}")
    if outputs.pptx_path:
        print(f"授课 PPT：{outputs.pptx_path}")
    if outputs.manifest_path:
        print(f"工程清单：{outputs.manifest_path}")


def main():
    arg_parser = argparse.ArgumentParser(
        description="公考试卷整理与导出工具",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用示例:
  python main.py
  python main.py --pdf-input exam.pdf --docx-output exam_题本.docx
  python main.py --pdf-input exam.pdf --ppt-output exam.pptx --subject data
  python main.py --pdf-input exam.pdf --ppt-output exam.pptx --subject politics,common_sense,verbal
  python main.py -i exam.docx
  python main.py -i exam.docx -o out.pptx -t template.pptx
        """,
    )
    arg_parser.add_argument("-i", "--input", help="输入 Word 文件路径 (.docx)")
    arg_parser.add_argument("--pdf-input", help="输入 PDF 试卷路径 (.pdf)")
    arg_parser.add_argument("-o", "--output", help="输出 PPT 文件路径 (.pptx)")
    arg_parser.add_argument("--docx-output", help="输出题本 Word 文件路径 (.docx)")
    arg_parser.add_argument("--ppt-output", help="输出授课 PPT 文件路径 (.pptx)")
    arg_parser.add_argument("--manifest-output", help="输出工程清单 JSON 路径")
    arg_parser.add_argument("-t", "--template", help="PPT 模板文件路径 (.pptx)")
    arg_parser.add_argument(
        "--layout",
        choices=["grid", "list", "one_row"],
        help="选项排列：grid / list / one_row",
    )
    arg_parser.add_argument("--font-size", type=int, help="题干字号 (pt)")
    arg_parser.add_argument(
        "--subject",
        default="all",
        help=(
            "PDF 处理科目：all / politics / common_sense / verbal / quant / reasoning / data；"
            "也支持逗号组合，如 politics,common_sense,verbal"
        ),
    )
    arg_parser.add_argument(
        "--question-range",
        help="按题号筛选，例如 66-85,111-115",
    )
    arg_parser.add_argument(
        "--document-subject",
        choices=["auto", "politics", "common_sense", "verbal", "quant", "reasoning", "data"],
        default="auto",
        help="整份文档科目提示：适合单科题库或没有大标题的 PDF / Word",
    )
    arg_parser.add_argument(
        "--log-level",
        choices=["DEBUG", "INFO", "WARNING", "ERROR"],
        default="WARNING",
        help="CLI 日志级别",
    )
    arg_parser.add_argument(
        "--ai-repair",
        action="store_true",
        help="启用本地 AI 修复器，对题目工程做结构化修复并直接写回后再导出",
    )
    arg_parser.add_argument(
        "--ai-limit",
        type=int,
        default=12,
        help="单次本地 AI 修复最多处理多少题，默认 12",
    )
    arg_parser.add_argument(
        "--ai-all-questions",
        action="store_true",
        help="默认只修复 AI 质检标出的待确认题；启用后对当前工程里的所有题尝试 AI 修复",
    )

    args = arg_parser.parse_args()
    configure_logging(args.log_level)

    if args.pdf_input:
        run_pdf_workflow(args)
    elif args.input:
        run_cli(args)
    else:
        run_gui()


if __name__ == "__main__":
    main()
