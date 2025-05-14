"""
提供文档格式选择相关的界面组件
"""
import tkinter as tk
from tkinter import messagebox # Import messagebox
# 使用兼容性模块替代直接导入ttkbootstrap
from gui.utils.ttk_compat import *
# 不再直接导入ttkbootstrap
# import ttkbootstrap as ttk
# from ttkbootstrap.constants import *
from gui.utils.ui_utils import create_tooltip

class FormatFrame(ttk.LabelFrame):
    """格式选择框架，提供文档格式选择功能"""
    
    def __init__(self, parent):
        """初始化格式选择框架"""
        # 根据ttkbootstrap可用性选择构造方式
        if TTKBOOTSTRAP_AVAILABLE:
            super().__init__(parent, text="文档格式设置", padding="10", bootstyle="info")
        else:
            super().__init__(parent, text="文档格式设置", padding="10")
        
        # 创建变量
        self.format_var = tk.StringVar(value="default")
        
        # 设置网格布局
        self.columnconfigure(0, weight=1)
        
        # 创建界面组件
        self._create_widgets()
        
    def _create_widgets(self):
        """创建界面组件"""
        # 创建格式选择分组框架
        if TTKBOOTSTRAP_AVAILABLE:
            format_group = ttk.LabelFrame(self, text="文档格式选择", padding=10, bootstyle="info")
        else:
            format_group = ttk.LabelFrame(self, text="文档格式选择", padding=10)
        format_group.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=5)
        format_group.columnconfigure(0, weight=1)
        format_group.columnconfigure(1, weight=1)
        
        # 创建格式单选按钮
        if TTKBOOTSTRAP_AVAILABLE:
            default_radio = ttk.Radiobutton(
                format_group, 
                text="默认格式", 
                variable=self.format_var, 
                value="default",
                bootstyle="info-toolbutton"
            )
        else:
            default_radio = ttk.Radiobutton(
                format_group, 
                text="默认格式", 
                variable=self.format_var, 
                value="default"
            )
        default_radio.grid(row=0, column=0, padx=10, sticky=tk.W)
        create_tooltip(default_radio, "使用默认文档格式")
        
        if TTKBOOTSTRAP_AVAILABLE:
            chinese_radio = ttk.Radiobutton(
                format_group, 
                text="中文标准格式", 
                variable=self.format_var, 
                value="chinese",
                bootstyle="info-toolbutton"
            )
        else:
            chinese_radio = ttk.Radiobutton(
                format_group, 
                text="中文标准格式", 
                variable=self.format_var, 
                value="chinese"
            )
        chinese_radio.grid(row=0, column=1, padx=10, sticky=tk.W)
        create_tooltip(chinese_radio, "使用中文标准格式")
        
        # 中文标准格式说明
        format_desc_frame = ttk.Frame(format_group, padding=5)
        format_desc_frame.grid(row=1, column=0, columnspan=2, padx=5, pady=5, sticky=(tk.W, tk.E))
        
        # 使用ttkbootstrap的滚动文本框
        if TTKBOOTSTRAP_AVAILABLE:
            format_desc = ttk.ScrolledText(
                format_desc_frame, 
                height=6,
                wrap=tk.WORD,
                bootstyle="info"
            )
        else:
            # Use standard ScrolledText directly
            format_desc = ScrolledText( 
                format_desc_frame, 
                height=6,
                wrap=tk.WORD
            )
        format_desc.pack(fill=tk.BOTH, expand=True)
        format_desc.insert(tk.END, "中文标准格式详情：\n"
                               "主标题：小二号方正小标宋简体\n"
                               "副标题：小三号仿宋_GB2312\n"
                               "正文：小三号仿宋_GB2312\n"
                               "文内一级标题：黑体小三，二级标题：楷体_GB2312小三，三四级标题：仿宋_GB2312小三")
        format_desc.config(state=tk.DISABLED)  # 设为只读
        
        # 创建预览按钮
        if TTKBOOTSTRAP_AVAILABLE:
            preview_btn = ttk.Button(self, text="预览样式", command=self.preview_format, bootstyle="info-outline", width=10)
        else:
            preview_btn = ttk.Button(self, text="预览样式", command=self.preview_format, width=10)
        preview_btn.grid(row=1, column=0, pady=10, sticky=tk.E)
    
    def get_format_config(self):
        """获取格式配置"""
        return {
            "use_chinese_format": self.format_var.get() == "chinese"
        }

    def preview_format(self):
        """预览所选格式的样式（占位符）"""
        selected_format = self.format_var.get()
        messagebox.showinfo("预览样式", f"预览功能待实现。\n当前选择: {selected_format}")