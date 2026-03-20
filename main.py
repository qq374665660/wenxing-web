# -*- coding: utf-8 -*-
"""
成都地区地基基础分析系统 - 主入口

用法：
    python main.py          # 启动 GUI 界面
    python main.py --help   # 显示帮助信息
"""

import sys
import argparse

from wenxing.interfaces.desktop import ModernUI
from wenxing.core import run_analysis_with_ui


def main():
    """主函数"""
    parser = argparse.ArgumentParser(
        description='成都地区地基基础分析系统',
        formatter_class=argparse.RawDescriptionHelpFormatter
    )
    parser.add_argument(
        '--cli', action='store_true',
        help='使用命令行模式（需要指定输入输出文件）'
    )
    parser.add_argument(
        '-i', '--input', type=str,
        help='输入 Excel 文件路径'
    )
    parser.add_argument(
        '-o', '--output', type=str,
        help='输出 Word 文件路径'
    )
    parser.add_argument(
        '-t', '--template', type=str, default='',
        help='Word 模板文件路径（可选）'
    )

    args = parser.parse_args()

    if args.cli:
        # 命令行模式
        if not args.input or not args.output:
            print("错误：命令行模式需要指定 --input 和 --output 参数")
            sys.exit(1)
        
        try:
            from wenxing.core import run_analysis
            run_analysis(args.template, args.input, args.output)
            print(f"分析完成，报告已保存至: {args.output}")
        except Exception as e:
            print(f"分析失败: {e}")
            sys.exit(1)
    else:
        # GUI 模式
        app = ModernUI(run_analysis_func=run_analysis_with_ui)
        app.run()


if __name__ == "__main__":
    main()
