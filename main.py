import argparse
import logging
import os
import sys


def configure_logging(level_name: str) -> None:
    level = getattr(logging, (level_name or "WARNING").upper(), logging.WARNING)
    logging.basicConfig(level=level, format="%(levelname)s: %(message)s")


def run_gui():
    from gui.app import PPTConvertApp

    app = PPTConvertApp()
    app.run()


def run_cli(args):
    from core.ppt_generator import PPTConfig, PPTGenerator
    from core.word_parser import WordParser
    from pptx.util import Pt

    if not os.path.exists(args.input):
        print(f"错误：文件不存在 - {args.input}")
        sys.exit(1)

    output = args.output or (os.path.splitext(args.input)[0] + ".pptx")
    parser = WordParser()

    try:
        print(f"正在解析：{args.input}")
        questions = parser.parse(args.input)
        print(f"共解析到 {len(questions)} 道题目")

        if not questions:
            print("未找到任何题目，请检查 Word 文档格式。")
            sys.exit(1)

        config = PPTConfig()
        if args.layout:
            config.option_layout = args.layout
        if args.font_size:
            config.stem_font_size = Pt(args.font_size)
            config.option_font_size = Pt(max(args.font_size - 2, 1))

        template = args.template or None

        def on_progress(current, total):
            pct = int(current / total * 100) if total else 100
            bar = "#" * (pct // 2) + "-" * (50 - pct // 2)
            print(
                f"\r生成进度：[{bar}] {pct}% ({current}/{total})",
                end="",
                flush=True,
            )

        print("正在生成 PPT...")
        generator = PPTGenerator(config=config)
        generator.generate(
            questions,
            output,
            template_path=template,
            progress_callback=on_progress,
        )

        print(f"\n完成！输出文件：{output}")
    finally:
        parser.cleanup()


def main():
    arg_parser = argparse.ArgumentParser(
        description="Word 选择题转 PPT 工具",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用示例:
  python main.py
  python main.py -i exam.docx
  python main.py -i exam.docx -o out.pptx -t template.pptx
        """,
    )
    arg_parser.add_argument("-i", "--input", help="输入 Word 文件路径 (.docx)")
    arg_parser.add_argument("-o", "--output", help="输出 PPT 文件路径 (.pptx)")
    arg_parser.add_argument("-t", "--template", help="PPT 模板文件路径 (.pptx)")
    arg_parser.add_argument(
        "--layout",
        choices=["grid", "list", "one_row"],
        help="选项排列：grid / list / one_row",
    )
    arg_parser.add_argument("--font-size", type=int, help="题干字号 (pt)")
    arg_parser.add_argument(
        "--log-level",
        choices=["DEBUG", "INFO", "WARNING", "ERROR"],
        default="WARNING",
        help="CLI 日志级别",
    )

    args = arg_parser.parse_args()
    configure_logging(args.log_level)

    if args.input:
        run_cli(args)
    else:
        run_gui()


if __name__ == "__main__":
    main()
