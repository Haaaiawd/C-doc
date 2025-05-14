"""
提供文件选择相关的界面组件
"""
import tkinter as tk
import os
from gui.utils.ttk_compat import *  # 导入兼容性组件
from tkinter import filedialog
from tkinter import messagebox as Messagebox  # 替换ttkbootstrap.dialogs
from gui.utils.ui_utils import create_tooltip, setup_drag_drop

class FileFrame(ttk.LabelFrame):
    """文件选择框架，提供输入和输出文件夹选择功能"""
    
    def __init__(self, parent):
        """初始化文件选择框架"""
        super().__init__(parent, text="文件选择", padding="10")
        
        # 创建变量
        self.input_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.suffix_enabled = tk.BooleanVar(value=True)
        self.suffix_text = tk.StringVar(
            value="——福州大学先进制造学院与海洋学院关工委2023年'中华魂'（毛泽东伟大精神品格）主题教育征文"
        )
        
        # 设置网格布局
        self.columnconfigure(0, weight=0)  # 标签列不伸展
        self.columnconfigure(1, weight=1)  # 内容列伸展
        self.columnconfigure(2, weight=0)  # 按钮列不伸展
        
        # 创建界面组件
        self._create_widgets()
        
    def _create_widgets(self):
        """创建界面组件"""
        # 输入文件夹选择
        input_label = ttk.Label(self, text="输入文件夹:")
        input_label.grid(row=0, column=0, sticky=tk.W, pady=5)
        
        input_entry = ttk.Entry(self, textvariable=self.input_path, width=50)
        input_entry.grid(row=0, column=1, padx=5, sticky=(tk.W, tk.E), pady=5)
        create_tooltip(input_entry, "存放Word文件的文件夹")
        
        input_btn = ttk.Button(
            self, 
            text="浏览",
            command=self._choose_input_dir,
        )
        input_btn.grid(row=0, column=2, pady=5)
        
        # 添加拖放支持提示
        drop_frame = ttk.Frame(self, padding=5)
        drop_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=3)
        
        drop_label = ttk.Label(
            drop_frame,
            text="提示: 您也可以直接拖拽文件夹到输入框",
            font=("微软雅黑", 9)
        )
        drop_label.pack(anchor=tk.W)
        
        # 输出文件夹选择
        output_label = ttk.Label(self, text="输出文件夹:")
        output_label.grid(row=2, column=0, sticky=tk.W, pady=5)
        
        output_entry = ttk.Entry(self, textvariable=self.output_path, width=50)
        output_entry.grid(row=2, column=1, padx=5, sticky=(tk.W, tk.E), pady=5)
        create_tooltip(output_entry, "处理后的文件将保存到这个文件夹")
        
        output_btn = ttk.Button(
            self, 
            text="浏览",
            command=self._choose_output_dir,
        )
        output_btn.grid(row=2, column=2, pady=5)
        
        # 添加"使用相同文件夹"快捷按钮
        same_dir_btn = ttk.Button(
            self, 
            text="使用相同文件夹",
            command=lambda: self.output_path.set(self.input_path.get()),
            width=12
        )
        same_dir_btn.grid(row=2, column=3, padx=5, pady=5)
        create_tooltip(same_dir_btn, "将输出文件夹设置为与输入文件夹相同")
        
        # 标题后缀选项
        suffix_label = ttk.Label(self, text="标题后缀:")
        suffix_label.grid(row=3, column=0, sticky=tk.W, pady=5)
        
        suffix_entry = ttk.Entry(self, textvariable=self.suffix_text, width=50)
        suffix_entry.grid(row=3, column=1, padx=5, sticky=(tk.W, tk.E), pady=5)
        create_tooltip(suffix_entry, "将添加到标题后的文本")
        
        suffix_check = ttk.Checkbutton(
            self, 
            text="启用标题后缀", 
            variable=self.suffix_enabled,
        )
        suffix_check.grid(row=3, column=2, pady=5)
        
        # 添加拖放支持
        self._setup_drag_drop(input_entry)
        self._setup_drag_drop(output_entry)
        
    def _setup_drag_drop(self, entry):
        """设置拖放支持"""
        # 使用工具函数实现拖放支持
        def handle_drop(path):
            if os.path.isdir(path):
                entry.delete(0, tk.END)
                entry.insert(0, path)
                
                # 如果是输入路径，且输出路径为空，则自动设置相同的输出路径
                if entry == self.nametowidget(entry.winfo_parent()).grid_slaves(row=0, column=1)[0]:
                    if not self.output_path.get():
                        self.output_path.set(path)
                
        # 调用ui_utils.py中的setup_drag_drop函数
        setup_drag_drop(entry, handle_drop, types="dirs")
        
    def _choose_input_dir(self):
        """选择输入文件夹"""
        directory = filedialog.askdirectory()
        if directory:
            self.input_path.set(directory)
            # 默认设置相同的输出目录
            if not self.output_path.get():
                self.output_path.set(directory)
    
    def _choose_output_dir(self):
        """选择输出文件夹"""
        directory = filedialog.askdirectory()
        if directory:
            self.output_path.set(directory)
    
    def get_paths(self):
        """获取输入和输出路径"""
        input_dir = self.input_path.get().strip()
        output_dir = self.output_path.get().strip()
        
        # 验证路径
        if not input_dir:
            Messagebox.showerror("输入错误", "请选择输入文件夹")
            return None
        
        if not output_dir:
            Messagebox.showerror("输入错误", "请选择输出文件夹")
            return None
        
        if not os.path.isdir(input_dir):
            Messagebox.showerror("路径错误", f"输入文件夹不存在: {input_dir}")
            return None
        
        if not os.path.exists(output_dir):
            try:
                os.makedirs(output_dir)
            except Exception as e:
                Messagebox.showerror("路径错误", f"无法创建输出文件夹: {str(e)}")
                return None
        
        return {
            "input_dir": input_dir,
            "output_dir": output_dir
        }
    
    def get_suffix_config(self):
        """获取后缀配置"""
        return {
            "enabled": self.suffix_enabled.get(),
            "text": self.suffix_text.get() if self.suffix_enabled.get() else ""
        }