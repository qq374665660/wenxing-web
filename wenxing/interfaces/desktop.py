# -*- coding: utf-8 -*-
"""
桌面 GUI 界面（Tkinter）
"""

import os
import tkinter as tk
from tkinter import filedialog, messagebox


class ModernUI:
    """现代化 UI 界面"""
    
    # 颜色主题
    PRIMARY_BLUE = "#2563EB"
    PRIMARY_GREEN = "#16A34A"
    LIGHT_BLUE = "#EFF6FF"
    LIGHT_GREEN = "#F0FDF4"
    WHITE = "#FFFFFF"
    GRAY_50 = "#F9FAFB"
    GRAY_100 = "#F3F4F6"
    GRAY_200 = "#E5E7EB"
    GRAY_400 = "#9CA3AF"
    GRAY_600 = "#4B5563"
    GRAY_800 = "#1F2937"

    def __init__(self, run_analysis_func=None):
        self.root = tk.Tk()
        self.root.title("成都地区地基基础分析及选型")
        self.root.geometry("420x580")
        self.root.configure(bg=self.WHITE)
        self.root.resizable(False, False)

        self.input_path_var = tk.StringVar()
        self.output_path_var = tk.StringVar()
        self.template_path_var = tk.StringVar()
        self.current_step = 1
        self.run_analysis_func = run_analysis_func

        self.setup_ui()

    def setup_ui(self):
        main_frame = tk.Frame(self.root, bg=self.WHITE, padx=25, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        self.create_header(main_frame)
        self.create_step_indicator(main_frame)
        self.create_cards(main_frame)
        self.create_footer(main_frame)

    def create_header(self, parent):
        header_frame = tk.Frame(parent, bg=self.WHITE)
        header_frame.pack(fill=tk.X, pady=(0, 15))

        title_label = tk.Label(
            header_frame,
            text="成都地区地基基础分析及选型",
            font=("Microsoft YaHei UI", 16, "bold"),
            fg=self.GRAY_800,
            bg=self.WHITE
        )
        title_label.pack(anchor='w')

        subtitle_label = tk.Label(
            header_frame,
            text="Foundation Analysis & Selection System",
            font=("Segoe UI", 9),
            fg=self.GRAY_400,
            bg=self.WHITE
        )
        subtitle_label.pack(anchor='w', pady=(2, 0))

        close_btn = tk.Label(
            header_frame,
            text="×",
            font=("Arial", 18),
            fg=self.GRAY_400,
            bg=self.WHITE,
            cursor="hand2"
        )
        close_btn.place(relx=1.0, rely=0, anchor='ne')
        close_btn.bind("<Button-1>", lambda e: self.root.destroy())
        close_btn.bind("<Enter>", lambda e: close_btn.configure(fg=self.GRAY_600))
        close_btn.bind("<Leave>", lambda e: close_btn.configure(fg=self.GRAY_400))

    def create_step_indicator(self, parent):
        step_frame = tk.Frame(parent, bg=self.WHITE)
        step_frame.pack(fill=tk.X, pady=(0, 25))

        steps = [("01", "上传数据"), ("02", "开始分析"), ("03", "获取结果")]
        self.step_circles = []
        self.step_labels = []

        for i, (num, text) in enumerate(steps):
            step_container = tk.Frame(step_frame, bg=self.WHITE)
            step_container.pack(side=tk.LEFT, expand=True)

            is_active = (i + 1) == self.current_step
            circle_color = self.PRIMARY_BLUE if is_active else self.WHITE
            text_color = self.WHITE if is_active else self.GRAY_400
            border_color = self.PRIMARY_BLUE if is_active else self.GRAY_200

            circle = tk.Frame(step_container, bg=circle_color, width=40, height=40)
            circle.pack()
            circle.pack_propagate(False)
            circle.configure(highlightbackground=border_color, highlightthickness=2)

            num_label = tk.Label(
                circle, text=num, font=("Segoe UI", 11, "bold"),
                fg=text_color, bg=circle_color
            )
            num_label.place(relx=0.5, rely=0.5, anchor='center')

            step_text = tk.Label(
                step_container, text=text, font=("Microsoft YaHei UI", 9),
                fg=self.GRAY_600 if is_active else self.GRAY_400, bg=self.WHITE
            )
            step_text.pack(pady=(5, 0))

            self.step_circles.append((circle, num_label))
            self.step_labels.append(step_text)

            if i < len(steps) - 1:
                line_frame = tk.Frame(step_frame, bg=self.WHITE, width=40)
                line_frame.pack(side=tk.LEFT, fill=tk.Y, pady=20)
                line = tk.Frame(line_frame, bg=self.GRAY_200, height=2)
                line.pack(fill=tk.X, pady=18)

    def update_step(self, step):
        self.current_step = step
        for i, ((circle, num_label), step_label) in enumerate(zip(self.step_circles, self.step_labels)):
            is_active = (i + 1) <= step
            is_current = (i + 1) == step

            circle_color = self.PRIMARY_BLUE if is_active else self.WHITE
            text_color = self.WHITE if is_active else self.GRAY_400
            border_color = self.PRIMARY_BLUE if is_active else self.GRAY_200

            circle.configure(bg=circle_color, highlightbackground=border_color)
            num_label.configure(bg=circle_color, fg=text_color)
            step_label.configure(fg=self.GRAY_600 if is_current else self.GRAY_400)

    def create_cards(self, parent):
        cards_frame = tk.Frame(parent, bg=self.WHITE)
        cards_frame.pack(fill=tk.BOTH, expand=True)

        self.create_card(
            cards_frame, icon="", title="选择Excel文件",
            subtitle="点击上传地质勘察数据", bg_color=self.PRIMARY_BLUE,
            command=self.select_input_file, path_var=self.input_path_var
        )

        self.create_card(
            cards_frame, icon="", title="输出文件位置",
            subtitle="结果将自动下载到默认文件夹", bg_color=None,
            command=self.select_output_file, path_var=self.output_path_var, is_outline=True
        )

        self.create_card(
            cards_frame, icon="", title="开始分析",
            subtitle="点击运行地基基础智能分析", bg_color=self.PRIMARY_GREEN,
            command=self.run_analysis
        )

    def create_card(self, parent, icon, title, subtitle, bg_color, command, path_var=None, is_outline=False):
        if is_outline:
            card = tk.Frame(parent, bg=self.WHITE, highlightbackground=self.GRAY_200, highlightthickness=1)
        else:
            card = tk.Frame(parent, bg=bg_color)

        card.pack(fill=tk.X, pady=8, ipady=12)
        card.configure(cursor="hand2")

        content_bg = self.WHITE if is_outline else bg_color
        content = tk.Frame(card, bg=content_bg, padx=15)
        content.pack(fill=tk.BOTH, expand=True)

        icon_frame = tk.Frame(content, bg=content_bg, width=45, height=45)
        icon_frame.pack(side=tk.LEFT, padx=(0, 12))
        icon_frame.pack_propagate(False)

        if is_outline:
            icon_frame.configure(highlightbackground=self.GRAY_200, highlightthickness=1)

        icon_label = tk.Label(
            icon_frame, text=icon, font=("Segoe UI Emoji", 16),
            bg=content_bg, fg=self.PRIMARY_BLUE if is_outline else self.WHITE
        )
        icon_label.place(relx=0.5, rely=0.5, anchor='center')

        text_frame = tk.Frame(content, bg=content_bg)
        text_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        title_color = self.GRAY_800 if is_outline else self.WHITE
        subtitle_color = self.GRAY_400 if is_outline else "#E0E7FF"

        title_label = tk.Label(
            text_frame, text=title, font=("Microsoft YaHei UI", 12, "bold"),
            fg=title_color, bg=content_bg, anchor='w'
        )
        title_label.pack(fill=tk.X)

        if path_var:
            if not hasattr(self, 'subtitle_labels'):
                self.subtitle_labels = {}
            sub_label = tk.Label(
                text_frame, text=subtitle, font=("Microsoft YaHei UI", 9),
                fg=subtitle_color, bg=content_bg, anchor='w'
            )
            sub_label.pack(fill=tk.X)
            self.subtitle_labels[str(id(path_var))] = (sub_label, subtitle, subtitle_color, content_bg)
        else:
            tk.Label(
                text_frame, text=subtitle, font=("Microsoft YaHei UI", 9),
                fg=subtitle_color, bg=content_bg, anchor='w'
            ).pack(fill=tk.X)

        arrow_label = tk.Label(
            content, text=">", font=("Arial", 16, "bold"),
            fg=self.GRAY_400 if is_outline else "#E0E7FF", bg=content_bg
        )
        arrow_label.pack(side=tk.RIGHT)

        def on_click(e):
            command()

        for widget in [card, content, icon_frame, icon_label, text_frame, title_label, arrow_label]:
            widget.bind("<Button-1>", on_click)

        def on_enter(e):
            if not is_outline:
                new_color = self._lighten_color(bg_color)
                card.configure(bg=new_color)
                for w in [content, icon_frame, icon_label, text_frame, title_label, arrow_label]:
                    try:
                        w.configure(bg=new_color)
                    except Exception:
                        pass

        def on_leave(e):
            if not is_outline:
                card.configure(bg=bg_color)
                for w in [content, icon_frame, icon_label, text_frame, title_label, arrow_label]:
                    try:
                        w.configure(bg=bg_color)
                    except Exception:
                        pass

        for widget in [card, content, icon_frame, icon_label, text_frame, title_label, arrow_label]:
            widget.bind("<Enter>", on_enter)
            widget.bind("<Leave>", on_leave)

    def _lighten_color(self, hex_color):
        if not hex_color:
            return self.GRAY_100
        hex_color = hex_color.lstrip('#')
        r, g, b = int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16)
        r, g, b = min(255, int(r * 1.1)), min(255, int(g * 1.1)), min(255, int(b * 1.1))
        return f"#{r:02x}{g:02x}{b:02x}"

    def create_footer(self, parent):
        footer = tk.Label(
            parent, text="中建西勘院 文兴", font=("Microsoft YaHei UI", 10),
            fg=self.GRAY_400, bg=self.WHITE
        )
        footer.pack(side=tk.BOTTOM, pady=(15, 0))

    def select_input_file(self):
        path = filedialog.askopenfilename(
            title="选择数据 Excel 文件",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if path:
            self.input_path_var.set(path)
            self.update_step(1)
            key = str(id(self.input_path_var))
            if hasattr(self, 'subtitle_labels') and key in self.subtitle_labels:
                label, _, _, _ = self.subtitle_labels[key]
                filename = os.path.basename(path)
                label.configure(text=f"已选择: {filename}")

    def select_output_file(self):
        path = filedialog.asksaveasfilename(
            title="保存输出 Word 文件",
            defaultextension=".docx",
            filetypes=[("Word files", "*.docx")]
        )
        if path:
            self.output_path_var.set(path)
            self.update_step(2)
            key = str(id(self.output_path_var))
            if hasattr(self, 'subtitle_labels') and key in self.subtitle_labels:
                label, _, _, _ = self.subtitle_labels[key]
                filename = os.path.basename(path)
                label.configure(text=f"保存至: {filename}")

    def run_analysis(self):
        input_path = self.input_path_var.get()
        output_path = self.output_path_var.get()

        if not input_path or not output_path:
            messagebox.showwarning("警告", "请先选择数据Excel文件和输出Word路径")
            return

        self.update_step(3)

        if self.run_analysis_func:
            try:
                self.run_analysis_func(self.template_path_var.get(), input_path, output_path)
                messagebox.showinfo("完成", "分析完成，Word 报告已生成。")
            except FileNotFoundError as e:
                messagebox.showerror("错误", f"文件未找到: {e}")
            except KeyError as e:
                messagebox.showerror("错误", f"Excel中缺少工作表或列: {e}")
            except Exception as e:
                messagebox.showerror("错误", f"发生未知错误: {e}")
        else:
            messagebox.showwarning("警告", "分析函数未配置")

    def run(self):
        self.root.mainloop()
