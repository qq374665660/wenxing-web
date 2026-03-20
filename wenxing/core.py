# -*- coding: utf-8 -*-
"""
核心分析入口模块

采用桥接策略：调用原始 wenxing2.py 中的完整实现，确保功能一致性。
后续可逐步将逻辑迁移到模块化组件中。
"""

import sys
import os
from tkinter import messagebox

# 将项目根目录添加到路径，以便导入 wenxing2
_project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _project_root not in sys.path:
    sys.path.insert(0, _project_root)

# 导入原始实现
from wenxing2 import run_analysis as _original_run_analysis

from .logging_config import setup_logging

# 初始化日志
setup_logging()


def run_analysis(template_path, input_path, output_path, ask_func=None):
    """
    运行地基基础分析（桥接到原始实现）
    
    Args:
        template_path: Word 模板路径（可选）
        input_path: Excel 输入文件路径
        output_path: Word 输出文件路径
        ask_func: 用户交互函数（用于询问问题）
    """
    if ask_func is None:
        ask_func = messagebox.askyesno
    
    # 调用原始实现
    _original_run_analysis(template_path, input_path, output_path, ask_func)


def run_analysis_with_ui(template_path, input_path, output_path):
    """带 UI 交互的分析入口"""
    try:
        run_analysis(template_path, input_path, output_path, messagebox.askyesno)
    except FileNotFoundError as e:
        messagebox.showerror("错误", f"文件未找到: {e}")
    except KeyError as e:
        messagebox.showerror("错误", f"Excel中缺少工作表或列: {e}")
    except Exception as e:
        messagebox.showerror("错误", f"发生未知错误: {e}")
