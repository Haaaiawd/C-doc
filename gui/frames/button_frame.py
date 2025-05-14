"""
提供功能按钮相关的界面组件
"""
import tkinter as tk
import os
import sys
from gui.utils.ttk_compat import *  # 导入兼容性组件
from gui.utils.ui_utils import create_tooltip

class ButtonFrame(ttk.Frame):
    """按钮框架，提供各种功能按钮"""
    
    def __init__(self, parent, bootstyle=PRIMARY):
        """
        初始化按钮框架
        
        参数:
            parent: 父容器
            bootstyle: 按钮样式（仅ttkbootstrap可用时使用）
        """
        # 根据ttkbootstrap可用性决定是否使用bootstyle参数
        if TTKBOOTSTRAP_AVAILABLE:
            super().__init__(parent, padding="5", bootstyle=bootstyle)
        else:
            super().__init__(parent, padding="5")
        
        # 保存样式
        self.bootstyle = bootstyle
        
        # 创建按钮
        self.convert_btn = None
        self.pack_error_btn = None
        self.extract_images_btn = None
        self.extract_titles_btn = None
        self.check_wordcount_btn = None
        
        # 创建界面组件
        self._create_widgets()
        
    def _load_emoji_icons(self):
        """创建基于emoji的图标字典"""
        # 使用emoji作为简单图标
        return {
            'convert': '📄',       # 文档图标
            'error': '⚠️',        # 警告图标
            'images': '🖼️',       # 图片图标
            'titles': '📑',       # 标题图标
            'wordcount': '📊',    # 统计图标
            'settings': '⚙️'      # 设置图标
        }
        
    def _create_widgets(self):
        """创建界面组件"""
        # 加载图标
        icons = self._load_emoji_icons()
        
        # 转换按钮
        if TTKBOOTSTRAP_AVAILABLE:
            self.convert_btn = ttk.Button(
                self, 
                text=f"{icons['convert']} 开始转换",
                bootstyle=f"{self.bootstyle}-outline",
                width=12
            )
        else:
            self.convert_btn = ttk.Button(
                self, 
                text=f"{icons['convert']} 开始转换",
                width=12
            )
        self.convert_btn.pack(side=tk.LEFT, padx=5)
        create_tooltip(self.convert_btn, "开始处理Word文档")
        
        # 打包错误文件按钮
        if TTKBOOTSTRAP_AVAILABLE:
            self.pack_error_btn = ttk.Button(
                self, 
                text=f"{icons['error']} 打包错误文件", 
                bootstyle=f"warning-outline",
                width=15
            )
        else:
            self.pack_error_btn = ttk.Button(
                self, 
                text=f"{icons['error']} 打包错误文件",
                width=15
            )
        self.pack_error_btn.pack(side=tk.LEFT, padx=5)
        create_tooltip(self.pack_error_btn, "将处理失败的文件复制到单独文件夹")
        self.pack_error_btn.state(['disabled'])  # 初始状态为禁用
        
        # 提取图片按钮
        if TTKBOOTSTRAP_AVAILABLE:
            self.extract_images_btn = ttk.Button(
                self, 
                text=f"{icons['images']} 提取图片",
                bootstyle=f"info-outline",
                width=12
            )
        else:
            self.extract_images_btn = ttk.Button(
                self, 
                text=f"{icons['images']} 提取图片",
                width=12
            )
        self.extract_images_btn.pack(side=tk.LEFT, padx=5)
        create_tooltip(self.extract_images_btn, "从所有文档中提取图片")
        
        # 提取标题按钮
        if TTKBOOTSTRAP_AVAILABLE:
            self.extract_titles_btn = ttk.Button(
                self, 
                text=f"{icons['titles']} 提取标题",
                bootstyle=f"success-outline",
                width=12
            )
        else:
            self.extract_titles_btn = ttk.Button(
                self, 
                text=f"{icons['titles']} 提取标题",
                width=12
            )
        self.extract_titles_btn.pack(side=tk.LEFT, padx=5)
        create_tooltip(self.extract_titles_btn, "从所有文档中提取标题")
        
        # 检测字数按钮
        if TTKBOOTSTRAP_AVAILABLE:
            self.check_wordcount_btn = ttk.Button(
                self, 
                text=f"{icons['wordcount']} 检测字数",
                bootstyle=f"secondary-outline", 
                width=12
            )
        else:
            self.check_wordcount_btn = ttk.Button(
                self, 
                text=f"{icons['wordcount']} 检测字数",
                width=12
            )
        self.check_wordcount_btn.pack(side=tk.LEFT, padx=5)
        create_tooltip(self.check_wordcount_btn, "检测所有文档的字数")
        
    def set_command(self, button_name, command):
        """
        设置按钮的命令
        
        参数:
            button_name: 按钮名称 (convert, pack_error, extract_images, extract_titles, check_wordcount)
            command: 要绑定的命令
        """
        button_map = {
            'convert': self.convert_btn,
            'pack_error': self.pack_error_btn,
            'extract_images': self.extract_images_btn,
            'extract_titles': self.extract_titles_btn,
            'check_wordcount': self.check_wordcount_btn
        }
        
        if button_name in button_map and button_map[button_name]:
            button_map[button_name].configure(command=command)
        
    def enable_buttons(self, enable_all=True, error_files_exist=False):
        """
        启用或禁用按钮
        
        参数:
            enable_all: 是否启用所有按钮
            error_files_exist: 是否存在错误文件(决定是否启用打包错误文件按钮)
        """
        state = ['!disabled'] if enable_all else ['disabled']
        
        self.convert_btn.state(state)
        self.extract_images_btn.state(state)
        self.extract_titles_btn.state(state)
        self.check_wordcount_btn.state(state)
        
        # 打包错误文件按钮只在有错误文件时启用
        if enable_all and error_files_exist:
            self.pack_error_btn.state(['!disabled'])
        else:
            self.pack_error_btn.state(['disabled'])