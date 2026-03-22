import sys
import argparse
import os


def run_gui():
    from gui.app import PPTConvertApp
    app = PPTConvertApp()
    app.run()


def run_cli(args):
    from core.word_parser import WordParser
    from core.ppt_generator import PPTGenerator, PPTConfig
    from pptx.util import Pt

    if not os.path.exists(args.input):
        print(f"错误：文件不存在 - {args.input}")
        sys.exit(1)

    output = args.output
    if not output:
        output = os.path.splitext(args.input)[0] + ".pptx"

    print(f"正在解析: {args.input}")
    parser = WordParser()
    questions = parser.parse(args.input)
    print(f"共解析到 {len(questions)} 道题目")

    if not questions:
        print("未找到任何题目，请检查 Word 文件格式")
        sys.exit(1)

    config = PPTConfig()
    if args.layout:
        config.option_layout = args.layout
    if args.font_size:
        config.stem_font_size = Pt(args.font_size)
        config.option_font_size = Pt(args.font_size - 2)

    template = args.template if args.template else None

    def on_progress(current, total):
        pct = int(current / total * 100)
        bar = "█" * (pct // 2) + "░" * (50 - pct // 2)
        print(f"\r生成进度: [{bar}] {pct}% ({current}/{total})", end="", flush=True)

    print("正在生成 PPT...")
    generator = PPTGenerator(config=config)
    generator.generate(questions, output, template_path=template,
                       progress_callback=on_progress)

    print(f"\n完成！输出文件: {output}")
    parser.cleanup()


def main():
    arg_parser = argparse.ArgumentParser(
        description="Word 选择题转 PPT 工具",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用示例:
  python main.py                        启动图形界面
  python main.py -i exam.docx           命令行转换（默认输出 exam.pptx）
  python main.py -i exam.docx -o out.pptx -t template.pptx
  （使用 -t 时，样式以模板为准，忽略 --layout、--font-size 等参数）
        """
    )
    arg_parser.add_argument("-i", "--input", help="输入 Word 文件路径 (.docx)")
    arg_parser.add_argument("-o", "--output", help="输出 PPT 文件路径 (.pptx)")
    arg_parser.add_argument("-t", "--template", help="PPT 模板文件路径 (.pptx)")
    arg_parser.add_argument(
        "--layout",
        choices=["grid", "list", "one_row"],
        help="选项排列: grid(2x2) / list(竖排) / one_row(一行四选项)",
    )
    arg_parser.add_argument("--font-size", type=int,
                            help="题干字体大小 (pt)")

    args = arg_parser.parse_args()

    if args.input:
        run_cli(args)
    else:
        run_gui()


if __name__ == "__main__":
    main()
