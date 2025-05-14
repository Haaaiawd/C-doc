"""
提供字数检测选项相关的界面组件
"""
import tkinter as tk
import os
import sys
from gui.utils.ttk_compat import *  # 导入兼容性组件
from gui.utils.ui_utils import create_tooltip

# 检查是否安装了openpyxl库
try:
    import openpyxl
    EXCEL_AVAILABLE = True
except ImportError:
    EXCEL_AVAILABLE = False

class WordcountFrame(ttk.LabelFrame):
    """字数检测设置框架"""
    
    def __init__(self, parent):
        """
        初始化字数检测框架
        
        参数:
            parent: 父容器
        """
        # 根据ttkbootstrap可用性选择构造方式
        if TTKBOOTSTRAP_AVAILABLE:
            super().__init__(parent, text="字数检测设置", padding="10")
        else:
            super().__init__(parent, text="字数检测设置", padding="10")
            
        # 创建变量
        self.enable_wordcount = tk.BooleanVar(value=False)  # 是否启用字数检测
        self.min_words = tk.IntVar(value=800)  # 最小字数要求
        self.mark_files = tk.BooleanVar(value=True)  # 是否标记不合格文件
        self.move_files = tk.BooleanVar(value=False)  # 是否移动不合格文件
        self.generate_excel = tk.BooleanVar(value=True)  # 是否生成统计表格
        self.strict_count = tk.BooleanVar(value=False)  # 是否启用严格字数统计
        
        # 创建界面组件
        self._create_widgets()
        
    def _create_widgets(self):
        """创建界面组件"""
        # 设置网格布局
        self.columnconfigure(0, weight=1)
        
        # 基本设置组
        if TTKBOOTSTRAP_AVAILABLE:
            basic_group = ttk.LabelFrame(self, text="基本设置", padding=10)
        else:
            basic_group = ttk.LabelFrame(self, text="基本设置", padding=10)
        basic_group.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=5)
        
        # 启用字数检测复选框
        if TTKBOOTSTRAP_AVAILABLE:
            wordcount_check = ttk.Checkbutton(
                basic_group,
                text="启用字数检测",
                variable=self.enable_wordcount
            )
        else:
            wordcount_check = ttk.Checkbutton(
                basic_group,
                text="启用字数检测",
                variable=self.enable_wordcount
            )
        wordcount_check.grid(row=0, column=0, sticky=tk.W, pady=5)
        create_tooltip(wordcount_check, "选中时将检测文档字数")
        
        # 最小字数要求
        if TTKBOOTSTRAP_AVAILABLE:
            min_words_label = ttk.Label(basic_group, text="最小字数要求:")
        else:
            min_words_label = ttk.Label(basic_group, text="最小字数要求:")
        min_words_label.grid(row=1, column=0, sticky=tk.W, pady=5)
        
        # 最小字数输入框和单位
        words_frame = ttk.Frame(basic_group)
        words_frame.grid(row=1, column=1, sticky=tk.W, pady=5)
        
        # 减少按钮
        if TTKBOOTSTRAP_AVAILABLE:
            minus_btn = ttk.Button(
                words_frame, 
                text="-", 
                width=2,
                command=self._decrease_min_words
            )
        else:
            minus_btn = ttk.Button(
                words_frame, 
                text="-", 
                width=2,
                command=self._decrease_min_words
            )
        minus_btn.pack(side=tk.LEFT, padx=2)
        
        # 字数输入框
        if TTKBOOTSTRAP_AVAILABLE:
            words_entry = ttk.Entry(
                words_frame, 
                textvariable=self.min_words, 
                width=6
            )
        else:
            words_entry = ttk.Entry(
                words_frame, 
                textvariable=self.min_words, 
                width=6
            )
        words_entry.pack(side=tk.LEFT, padx=2)
        
        # 字数单位
        if TTKBOOTSTRAP_AVAILABLE:
            ttk.Label(words_frame, text="字").pack(side=tk.LEFT, padx=2)
        else:
            ttk.Label(words_frame, text="字").pack(side=tk.LEFT, padx=2)
            
        # 增加按钮
        if TTKBOOTSTRAP_AVAILABLE:
            plus_btn = ttk.Button(words_frame, text="+", width=2, command=self._increase_min_words)
        else:
            plus_btn = ttk.Button(words_frame, text="+", width=2, command=self._increase_min_words)
        plus_btn.pack(side=tk.LEFT, padx=2)
        
        # 处理选项组
        if TTKBOOTSTRAP_AVAILABLE:
            action_group = ttk.LabelFrame(self, text="处理选项", padding=10)
        else:
            action_group = ttk.LabelFrame(self, text="处理选项", padding=10)
        action_group.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=5)
        
        if TTKBOOTSTRAP_AVAILABLE:
            action_label = ttk.Label(action_group, text="字数不足的文件:")
        else:
            action_label = ttk.Label(action_group, text="字数不足的文件:")
        action_label.grid(row=0, column=0, sticky=tk.W, columnspan=2, pady=5)
        
        # 标记不合格文件复选框
        if TTKBOOTSTRAP_AVAILABLE:
            mark_check = ttk.Checkbutton(
                action_group,
                text="在文件名前添加[字数不足]标记",
                variable=self.mark_files,
                command=self._handle_mark_changed
            )
        else:
            mark_check = ttk.Checkbutton(
                action_group,
                text="在文件名前添加[字数不足]标记",
                variable=self.mark_files,
                command=self._handle_mark_changed
            )
        mark_check.grid(row=1, column=0, sticky=tk.W, padx=20, pady=5)
        create_tooltip(mark_check, "选中时将在字数不足的文件名前添加标记")
        
        # 移动不合格文件复选框
        if TTKBOOTSTRAP_AVAILABLE:
            move_check = ttk.Checkbutton(
                action_group,
                text="移动到字数不足子文件夹",
                variable=self.move_files
            )
        else:
            move_check = ttk.Checkbutton(
                action_group,
                text="移动到字数不足子文件夹",
                variable=self.move_files
            )
        move_check.grid(row=2, column=0, sticky=tk.W, padx=20, pady=5)
        create_tooltip(move_check, "选中时将字数不足的文件移动到独立文件夹")
        
        # 严格字数统计复选框
        if TTKBOOTSTRAP_AVAILABLE:
            strict_check = ttk.Checkbutton(
                action_group,
                text="启用严格字数统计(不计算图片、页眉页脚等)",
                variable=self.strict_count
            )
        else:
            strict_check = ttk.Checkbutton(
                action_group,
                text="启用严格字数统计(不计算图片、页眉页脚等)",
                variable=self.strict_count
            )
        strict_check.grid(row=3, column=0, sticky=tk.W, padx=20, pady=5)
        create_tooltip(strict_check, "选中时将使用更严格的字数统计方式，可能更准确但会更慢")
        
        # 输出选项组
        if TTKBOOTSTRAP_AVAILABLE:
            output_group = ttk.LabelFrame(self, text="输出选项", padding=10)
        else:
            output_group = ttk.LabelFrame(self, text="输出选项", padding=10)
        output_group.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=5)
        
        # 生成统计表格选项
        if TTKBOOTSTRAP_AVAILABLE:
            excel_check = ttk.Checkbutton(
                output_group,
                text="生成字数统计Excel表格",
                variable=self.generate_excel
            )
        else:
            excel_check = ttk.Checkbutton(
                output_group,
                text="生成字数统计Excel表格",
                variable=self.generate_excel
            )
        excel_check.pack(anchor=tk.W, pady=5)
        create_tooltip(excel_check, "选中时将生成包含所有文档字数统计的Excel表格")
        
        # 注意事项
        if TTKBOOTSTRAP_AVAILABLE:
            note_check = ttk.Checkbutton(
                output_group,
                text="启用实时统计(处理大量文件时可能会变慢)"
            )
        else:
            note_check = ttk.Checkbutton(
                output_group,
                text="启用实时统计(处理大量文件时可能会变慢)"
            )
        note_check.pack(anchor=tk.W, pady=5)
        create_tooltip(note_check, "选中时将在处理过程中实时显示字数统计")
        
        # 添加警告图标和文本
        warning_frame = ttk.Frame(self)
        warning_frame.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=10)
        
        if TTKBOOTSTRAP_AVAILABLE:
            warning_label = ttk.Label(
                warning_frame,
                text="⚠️ 注意：字数统计可能与Word显示的略有差异"
            )
        else:
            warning_label = ttk.Label(
                warning_frame,
                text="⚠️ 注意：字数统计可能与Word显示的略有差异"
            )
        warning_label.pack(side=tk.LEFT)
        
        # 添加"字数不足"示例图标
        example_frame = ttk.Frame(self)
        example_frame.grid(row=4, column=0, sticky=(tk.W, tk.E), pady=5)
        
        if TTKBOOTSTRAP_AVAILABLE:
            example_label = ttk.Label(
                example_frame,
                text="示例: [字数不足]文档名.docx",
                relief="solid",
                padding=5
            )
        else:
            example_label = ttk.Label(
                example_frame,
                text="示例: [字数不足]文档名.docx",
                relief="solid",
                padding=5
            )
        example_label.pack(side=tk.LEFT)
        
    def _increase_min_words(self):
        """增加最小字数要求"""
        current = self.min_words.get()
        if current < 10000:
            self.min_words.set(current + 100)
            
    def _decrease_min_words(self):
        """减少最小字数要求"""
        current = self.min_words.get()
        if current > 100:
            self.min_words.set(current - 100)
            
    def _handle_mark_changed(self):
        """处理标记复选框状态变化"""
        if not self.mark_files.get():
            self.move_files.set(False)
    
    def get_wordcount_config(self):
        """获取字数检测配置"""
        return {
            "enabled": self.enable_wordcount.get(),
            "min_words": self.min_words.get(),
            "mark_files": self.mark_files.get(),
            "move_files": self.move_files.get(),
            "generate_excel": self.generate_excel.get() and EXCEL_AVAILABLE,
            "strict_count": self.strict_count.get()
        }