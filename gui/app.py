"""
Word文档批量处理工具的主应用程序类
"""
import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox
import webbrowser

# 使用兼容性模块替代直接导入ttkbootstrap
from gui.utils.ttk_compat import *  # 导入所有兼容性组件

from gui.frames.file_frame import FileFrame
from gui.frames.format_frame import FormatFrame
from gui.frames.wordcount_frame import WordcountFrame
from gui.frames.button_frame import ButtonFrame
from gui.handlers.conversion_handler import ConversionHandler
from gui.handlers.image_handler import ImageHandler
from gui.handlers.wordcount_handler import WordcountHandler
from gui.handlers.title_handler import TitleHandler
from gui.utils.redirect_text import RedirectText
from gui.utils.ui_utils import set_window_icon, center_window, create_tooltip

class App:
    """Word文档批量处理工具的主应用程序类"""
    
    def __init__(self, root):
        """
        初始化应用程序
        
        参数:
            root: tkinter主窗口
        """
        self.root = root
        self.root.title("Word文档批量处理工具")
        self.root.minsize(900, 700)  # 增加最小高度，从800x600到900x700
        
        # 设置窗口图标
        set_window_icon(self.root)
        
        # 创建变量
        self.processing = False
        self.redirect = None
        self.log_text = None
        self.progress_bar = None
        self.status_label = None
        
        # 创建UI组件
        self._create_widgets()
        
        # 创建处理器
        self.conversion_handler = ConversionHandler(self)
        self.image_handler = ImageHandler(self)
        self.wordcount_handler = WordcountHandler(self)
        self.title_handler = TitleHandler(self)
        
        # 连接UI事件
        self._connect_events()
        
        # 设置键盘快捷键
        self._setup_keyboard_shortcuts()
        
        # 初始化状态
        self.set_status("就绪")
        
        # 窗口居中显示
        center_window(self.root, 900, 700)
        
    def _create_widgets(self):
        """创建界面组件"""
        # 配置根窗口网格布局
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)  # 让根窗口的行可以扩展

        # 创建主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建水平方向的PanedWindow，左侧功能区，右侧日志区
        h_paned_window = tk.PanedWindow(main_frame, orient=tk.HORIZONTAL)
        h_paned_window.pack(fill=tk.BOTH, expand=True)
        
        # 创建左侧功能区框架
        left_frame = ttk.Frame(h_paned_window)
        left_frame.columnconfigure(0, weight=1)
        left_frame.rowconfigure(1, weight=1)  # 让选项卡区域可以扩展
        left_frame.rowconfigure(2, weight=0)  # 状态和按钮区域不扩展
        
        # 创建右侧日志区域框架
        right_frame = ttk.Frame(h_paned_window)
        right_frame.columnconfigure(0, weight=1)
        right_frame.rowconfigure(0, weight=1)  # 让日志区域可以扩展
        
        # 将两个框架添加到水平分隔窗口
        h_paned_window.add(left_frame)
        h_paned_window.add(right_frame)
        
        # 设置初始分割位置，分配约70%的空间给左侧功能区
        h_paned_window.paneconfigure(left_frame, minsize=500)
        h_paned_window.paneconfigure(right_frame, minsize=250)
        
        # 在左侧框架中添加文件选择框架
        self.file_frame = FileFrame(left_frame)
        self.file_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=5)
        
        # 创建选项卡控件
        self.notebook = ttk.Notebook(left_frame)
        self.notebook.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        
        # 创建各个选项卡内容框架
        self._create_format_tab()
        self._create_images_tab()
        self._create_titles_tab()
        self._create_wordcount_tab()
        self._create_properties_tab()  # 文档属性选项卡
        self._create_settings_tab()
        
        # 创建状态和控制区域
        controls_frame = ttk.Frame(left_frame)
        controls_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=5)
        controls_frame.columnconfigure(0, weight=1)
        
        # 创建进度条框架
        progress_frame = ttk.Frame(controls_frame, padding="5")
        progress_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=5)
        progress_frame.columnconfigure(1, weight=1)
        
        # 创建状态标签
        ttk.Label(progress_frame, text="状态:").grid(row=0, column=0, sticky=tk.W)
        self.status_label = ttk.Label(progress_frame, text="就绪")
        self.status_label.grid(row=0, column=1, sticky=tk.W)
        
        # 创建进度条
        ttk.Label(progress_frame, text="进度:").grid(row=1, column=0, sticky=tk.W)
        self.progress_bar = ttk.Progressbar(progress_frame, mode="determinate")
        self.progress_bar.grid(row=1, column=1, sticky=(tk.W, tk.E))
        
        # 创建按钮框架
        self.button_frame = ButtonFrame(controls_frame)
        self.button_frame.grid(row=1, column=0, sticky=tk.E, pady=10)
        
        # 创建版权信息标签
        copyright_frame = ttk.Frame(controls_frame)
        copyright_frame.grid(row=2, column=0, sticky=(tk.W, tk.E))
        
        copyright_label = ttk.Label(
            copyright_frame, 
            text="© 2023-2025 先进制造与海洋学院关工委", 
            foreground="gray"
        )
        copyright_label.pack(side=tk.LEFT)
        
        # 添加"关于"链接和快捷键提示
        about_label = ttk.Label(
            copyright_frame, 
            text="关于", 
            foreground="blue", 
            cursor="hand2"
        )
        about_label.pack(side=tk.RIGHT)
        about_label.bind("<Button-1>", self._show_about)
        create_tooltip(about_label, "显示关于信息")
        
        shortcut_label = ttk.Label(
            copyright_frame,
            text="[F1: 帮助]",
            foreground="gray",
        )
        shortcut_label.pack(side=tk.RIGHT, padx=15)
        create_tooltip(shortcut_label, "按F1查看快捷键帮助")
        
        # 在右侧框架中添加日志区域
        log_frame = ttk.LabelFrame(right_frame, text="处理日志", padding="5")
        log_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
        # 创建日志工具栏
        log_toolbar = ttk.Frame(log_frame)
        log_toolbar.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
        
        # 添加日志操作按钮
        if TTKBOOTSTRAP_AVAILABLE:
            clear_log_btn = ttk.Button(log_toolbar, text="清空", 
                                     command=self._shortcut_clear_log, 
                                     bootstyle="info-outline", width=8)
            view_log_btn = ttk.Button(log_toolbar, text="大窗口", 
                                     command=self._open_log_window, 
                                     bootstyle="info-outline", width=8)
        else:
            clear_log_btn = ttk.Button(log_toolbar, text="清空", 
                                     command=self._shortcut_clear_log, 
                                     width=8)
            view_log_btn = ttk.Button(log_toolbar, text="大窗口", 
                                     command=self._open_log_window, 
                                     width=8)
        
        clear_log_btn.pack(side=tk.LEFT, padx=5)
        view_log_btn.pack(side=tk.LEFT, padx=5)
        create_tooltip(clear_log_btn, "清空日志内容")
        create_tooltip(view_log_btn, "在新窗口查看完整日志")
        
        # 创建日志文本框
        self.log_text = ScrolledText(
            log_frame,
            wrap=tk.WORD, 
            width=35,  # 设置适合右侧区域的宽度
            height=25,  
            font=("Consolas", 10)
        )
        self.log_text.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 创建重定向文本
        self.redirect = RedirectText(self.log_text)
        sys.stdout = self.redirect
        sys.stderr = self.redirect
        
    def _create_format_tab(self):
        """创建格式转换选项卡"""
        format_tab = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(format_tab, text=" 文档格式 ")
        
        # 设置网格布局
        format_tab.columnconfigure(0, weight=1)
        
        # 创建格式选择框架
        self.format_frame = FormatFrame(format_tab)
        self.format_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=5)
        
        # 添加转换说明
        desc_frame = ttk.LabelFrame(format_tab, text="功能说明", padding=10)
        desc_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=10)
        
        desc_text = ttk.Label(
            desc_frame,
            text="此功能可以批量转换Word文档格式，支持中文标准格式和默认格式。\n"
                 "转换过程将保留文档内容并应用所选格式设置。",
            wraplength=700,
            justify=tk.LEFT
        )
        desc_text.pack(fill=tk.X, expand=True)
    
    def _create_images_tab(self):
        """创建图片提取选项卡"""
        images_tab = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(images_tab, text=" 图片提取 ")
        
        # 设置网格布局
        images_tab.columnconfigure(0, weight=1)
        images_tab.columnconfigure(1, weight=1)
        
        # 添加图片提取说明
        desc_frame = ttk.LabelFrame(images_tab, text="功能说明", padding=10)
        desc_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        desc_text = ttk.Label(
            desc_frame,
            text="此功能可以从所选文件夹中的Word文档提取所有图片。\n"
                 "提取的图片将保存在输出文件夹中的images子文件夹内，并按文档名称分类。",
            wraplength=700,
            justify=tk.LEFT
        )
        desc_text.pack(fill=tk.X, expand=True)
        
        # 添加图片提取选项
        extract_options_frame = ttk.LabelFrame(images_tab, text="提取选项", padding=10)
        extract_options_frame.grid(row=1, column=0, sticky=(tk.N, tk.W, tk.E, tk.S), pady=5, padx=(0, 5))
        extract_options_frame.columnconfigure(0, weight=1)
        
        self.preserve_names = tk.BooleanVar(value=True)
        name_check = ttk.Checkbutton(
            extract_options_frame,
            text="保留原始图片名称",
            variable=self.preserve_names
        )
        name_check.grid(row=0, column=0, sticky=tk.W, pady=5)
        create_tooltip(name_check, "选中时尝试保留原始图片名称，否则使用自动生成的名称")
        
        # 添加图片格式选项
        formats_frame = ttk.Frame(extract_options_frame)
        formats_frame.grid(row=1, column=0, sticky=tk.W, pady=5)
        
        ttk.Label(formats_frame, text="保存格式:").grid(row=0, column=0, sticky=tk.W)
        
        self.image_format_var = tk.StringVar(value="original")
        if TTKBOOTSTRAP_AVAILABLE:
            format_combo = ttk.Combobox(
                formats_frame, 
                textvariable=self.image_format_var,
                values=["original", "png", "jpg", "webp"],
                width=10,
                bootstyle="info"
            )
        else:
            format_combo = ttk.Combobox(
                formats_frame, 
                textvariable=self.image_format_var,
                values=["original", "png", "jpg", "webp"],
                width=10
            )
        format_combo.grid(row=0, column=1, padx=5)
        create_tooltip(format_combo, "选择提取图片的保存格式，original表示保持原格式")
        
        # 添加图片品质设置
        quality_frame = ttk.Frame(extract_options_frame)
        quality_frame.grid(row=2, column=0, sticky=tk.W, pady=5)
        
        ttk.Label(quality_frame, text="图片质量:").grid(row=0, column=0, sticky=tk.W)
        
        self.image_quality_var = tk.IntVar(value=90)
        if TTKBOOTSTRAP_AVAILABLE:
            quality_spin = ttk.Spinbox(
                quality_frame,
                from_=10,
                to=100,
                textvariable=self.image_quality_var,
                width=5,
                bootstyle="info"
            )
        else:
            quality_spin = ttk.Spinbox(
                quality_frame,
                from_=10,
                to=100,
                textvariable=self.image_quality_var,
                width=5
            )
        quality_spin.grid(row=0, column=1, padx=5)
        create_tooltip(quality_spin, "设置图片保存质量，对JPG和WEBP格式有效")
        
        # 添加图片处理选项
        processing_frame = ttk.LabelFrame(images_tab, text="图片处理选项", padding=10)
        processing_frame.grid(row=1, column=1, sticky=(tk.N, tk.W, tk.E, tk.S), pady=5, padx=(5, 0))
        processing_frame.columnconfigure(0, weight=1)
        
        # 提取图片到单独文件夹选项
        self.extract_to_folder_var = tk.BooleanVar(value=False)
        if TTKBOOTSTRAP_AVAILABLE:
            extract_check = ttk.Checkbutton(
                processing_frame,
                text="提取图片到单独文件夹",
                variable=self.extract_to_folder_var,
                bootstyle="round-toggle"
            )
        else:
            extract_check = ttk.Checkbutton(
                processing_frame,
                text="提取图片到单独文件夹",
                variable=self.extract_to_folder_var
            )
        extract_check.grid(row=0, column=0, sticky=tk.W, pady=5)
        
        # 添加图片位置说明
        if TTKBOOTSTRAP_AVAILABLE:
            image_help = ttk.Label(
                processing_frame,
                text="提取的图片将保存在与输出文件相同目录下的images文件夹中",
                bootstyle="secondary",
                font=("", 9)
            )
        else:
            image_help = ttk.Label(
                processing_frame,
                text="提取的图片将保存在与输出文件相同目录下的images文件夹中",
                font=("", 9)
            )
        image_help.grid(row=1, column=0, sticky=tk.W, pady=2)
        
        # 添加图片大小限制选项
        self.resize_images_var = tk.BooleanVar(value=False)
        if TTKBOOTSTRAP_AVAILABLE:
            resize_check = ttk.Checkbutton(
                processing_frame,
                text="限制图片最大尺寸",
                variable=self.resize_images_var,
                bootstyle="round-toggle"
            )
        else:
            resize_check = ttk.Checkbutton(
                processing_frame,
                text="限制图片最大尺寸",
                variable=self.resize_images_var
            )
        resize_check.grid(row=2, column=0, sticky=tk.W, pady=(10, 2))
        
        # 添加尺寸设置框架
        size_frame = ttk.Frame(processing_frame)
        size_frame.grid(row=3, column=0, sticky=tk.W, padx=20, pady=2)
        
        ttk.Label(size_frame, text="最大宽度:").grid(row=0, column=0, sticky=tk.W)
        self.max_width_var = tk.IntVar(value=1024)
        width_spin = ttk.Spinbox(
            size_frame,
            from_=100,
            to=4000,
            textvariable=self.max_width_var,
            width=5
        )
        width_spin.grid(row=0, column=1, padx=5)
        
        ttk.Label(size_frame, text="最大高度:").grid(row=0, column=2, sticky=tk.W, padx=(10, 0))
        self.max_height_var = tk.IntVar(value=768)
        height_spin = ttk.Spinbox(
            size_frame,
            from_=100,
            to=4000,
            textvariable=self.max_height_var,
            width=5
        )
        height_spin.grid(row=0, column=3, padx=5)
    
    def _create_titles_tab(self):
        """创建标题提取选项卡"""
        titles_tab = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(titles_tab, text=" 标题提取 ")
        
        # 设置网格布局
        titles_tab.columnconfigure(0, weight=1)
        
        # 添加标题提取说明
        desc_frame = ttk.LabelFrame(titles_tab, text="功能说明", padding=10)
        desc_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=10)
        
        desc_text = ttk.Label(
            desc_frame,
            text="此功能可以从所选文件夹中的Word文档提取所有标题。\n"
                 "提取的标题将保存为Excel文件，包含文件名、标题文本和作者信息。",
            wraplength=700,
            justify=tk.LEFT
        )
        desc_text.pack(fill=tk.X, expand=True)
        
        # 添加标题提取选项
        options_frame = ttk.LabelFrame(titles_tab, text="提取选项", padding=10)
        options_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=10)
        
        self.include_content = tk.BooleanVar(value=False)
        # 使用标准ttk
        content_check = ttk.Checkbutton(
            options_frame,
            text="包含正文摘要",
            variable=self.include_content
        )
        content_check.pack(anchor=tk.W, pady=5)
        create_tooltip(content_check, "选中时将包含文档正文的前100个字符作为摘要")
    
    def _create_wordcount_tab(self):
        """创建字数检测选项卡"""
        wordcount_tab = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(wordcount_tab, text=" 字数检测 ")
        
        # 设置网格布局
        wordcount_tab.columnconfigure(0, weight=1)
        
        # 创建字数检测框架
        self.wordcount_frame = WordcountFrame(wordcount_tab)
        self.wordcount_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=5)
    
    def _create_properties_tab(self):
        """创建文档属性选项卡"""
        properties_tab = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(properties_tab, text=" 文档属性 ")
        
        # 设置网格布局
        properties_tab.columnconfigure(0, weight=1)
        
        # 添加作者信息选项框
        author_frame = ttk.LabelFrame(properties_tab, text="作者信息设置", padding=10)
        author_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=10)
        author_frame.columnconfigure(0, weight=1)
        
        # 创建作者信息控制选项
        self.author_var = tk.BooleanVar(value=False)
        if TTKBOOTSTRAP_AVAILABLE:
            author_check = ttk.Checkbutton(
                author_frame,
                text="移除Word文档中的作者和公司信息",
                variable=self.author_var,
                bootstyle="round-toggle"
            )
        else:
            author_check = ttk.Checkbutton(
                author_frame,
                text="移除Word文档中的作者和公司信息",
                variable=self.author_var
            )
        author_check.grid(row=0, column=0, sticky=tk.W, pady=5)
        
        # 添加自定义作者信息选项
        custom_author_frame = ttk.Frame(author_frame)
        custom_author_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Label(custom_author_frame, text="自定义作者:").grid(row=0, column=0, padx=5, sticky=tk.W)
        self.author_name_var = tk.StringVar()
        author_entry = ttk.Entry(custom_author_frame, textvariable=self.author_name_var, width=25)
        author_entry.grid(row=0, column=1, padx=5, sticky=tk.W)
        
        ttk.Label(custom_author_frame, text="单位/机构:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        self.company_var = tk.StringVar()
        company_entry = ttk.Entry(custom_author_frame, textvariable=self.company_var, width=25)
        company_entry.grid(row=1, column=1, padx=5, pady=5, sticky=tk.W)
        
        # 添加功能说明
        desc_frame = ttk.LabelFrame(properties_tab, text="功能说明", padding=10)
        desc_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=10)
        
        desc_text = ttk.Label(
            desc_frame,
            text="此选项卡用于管理文档的属性信息，如作者信息和公司/机构信息。\n"
                 "您可以选择移除原有信息或者添加自定义的作者和单位信息。",
            wraplength=700,
            justify=tk.LEFT
        )
        desc_text.pack(fill=tk.X, expand=True)
    
    def _create_settings_tab(self):
        """创建设置选项卡"""
        settings_tab = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(settings_tab, text=" 设置 ")
        
        # 设置网格布局
        settings_tab.columnconfigure(0, weight=1)
        
        # 主题设置部分被移除，因为已经不再支持ttkbootstrap
        
        # 添加性能设置
        perf_frame = ttk.LabelFrame(settings_tab, text="性能设置", padding=10)
        perf_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=10)
        
        # 线程数设置
        thread_label = ttk.Label(perf_frame, text="处理线程数:")
        thread_label.grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        
        self.thread_var = tk.IntVar(value=4)
        thread_spin = ttk.Spinbox(
            perf_frame,
            from_=1,
            to=16,
            textvariable=self.thread_var,
            width=5
        )
        thread_spin.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        create_tooltip(thread_spin, "设置处理文档时使用的线程数，数值越大处理速度越快，但会占用更多系统资源")

    def _setup_keyboard_shortcuts(self):
        """设置键盘快捷键"""
        # 全局快捷键
        self.root.bind("<F1>", self._show_shortcuts_help)  # F1显示帮助
        self.root.bind("<Control-o>", self._shortcut_open_input)  # Ctrl+O打开输入文件夹
        self.root.bind("<Control-s>", self._shortcut_open_output)  # Ctrl+S选择输出文件夹
        self.root.bind("<Control-Return>", self._shortcut_convert)  # Ctrl+Enter开始转换
        self.root.bind("<Control-1>", lambda e: self.notebook.select(0))  # Ctrl+1切换到格式选项卡
        self.root.bind("<Control-2>", lambda e: self.notebook.select(1))  # Ctrl+2切换到图片选项卡
        self.root.bind("<Control-3>", lambda e: self.notebook.select(2))  # Ctrl+3切换到标题选项卡
        self.root.bind("<Control-4>", lambda e: self.notebook.select(3))  # Ctrl+4切换到字数选项卡
        self.root.bind("<Control-5>", lambda e: self.notebook.select(4))  # Ctrl+5切换到设置选项卡
        self.root.bind("<Control-i>", self._shortcut_extract_images)  # Ctrl+I提取图片
        self.root.bind("<Control-t>", self._shortcut_extract_titles)  # Ctrl+T提取标题
        self.root.bind("<Control-w>", self._shortcut_check_wordcount)  # Ctrl+W检测字数
        self.root.bind("<Control-l>", self._shortcut_clear_log)  # Ctrl+L清空日志
        
    def _shortcut_open_input(self, event=None):
        """快捷键：打开输入文件夹"""
        self.file_frame._choose_input_dir()
        
    def _shortcut_open_output(self, event=None):
        """快捷键：选择输出文件夹"""
        self.file_frame._choose_output_dir()
        
    def _shortcut_convert(self, event=None):
        """快捷键：开始转换"""
        if not self.processing:
            self.conversion_handler.start_conversion()
            
    def _shortcut_extract_images(self, event=None):
        """快捷键：提取图片"""
        if not self.processing:
            self.image_handler.extract_images()
            
    def _shortcut_extract_titles(self, event=None):
        """快捷键：提取标题"""
        if not self.processing:
            self.title_handler.extract_titles()
            
    def _shortcut_check_wordcount(self, event=None):
        """快捷键：检测字数"""
        if not self.processing:
            self.wordcount_handler.check_wordcount()
            
    def _shortcut_clear_log(self, event=None):
        """快捷键：清空日志"""
        if self.log_text:
            self.log_text.delete(1.0, tk.END)
            self.set_status("日志已清空")
            
    def _show_shortcuts_help(self, event=None):
        """显示快捷键帮助"""
        help_text = """
键盘快捷键:

文件操作:
  Ctrl+O - 选择输入文件夹
  Ctrl+S - 选择输出文件夹
  
功能操作:
  Ctrl+Enter - 开始转换
  Ctrl+I - 提取图片
  Ctrl+T - 提取标题
  Ctrl+W - 检测字数
  
界面导航:
  Ctrl+1 - 切换到"文档格式"选项卡
  Ctrl+2 - 切换到"图片提取"选项卡
  Ctrl+3 - 切换到"标题提取"选项卡
  Ctrl+4 - 切换到"字数检测"选项卡
  Ctrl+5 - 切换到设置选项卡
  
其他:
  F1 - 显示此帮助
  Ctrl+L - 清空日志
        """
        messagebox.showinfo("快捷键帮助", help_text)

    # 移除_change_theme方法，因为已经不再支持ttkbootstrap主题切换
        
    def _connect_events(self):
        """连接UI事件与处理函数"""
        # 连接按钮事件
        self.button_frame.set_command("convert", self.conversion_handler.start_conversion)
        self.button_frame.set_command("pack_error", self.conversion_handler.pack_error_files)
        self.button_frame.set_command("extract_images", self.image_handler.extract_images)
        self.button_frame.set_command("extract_titles", self.title_handler.extract_titles)
        self.button_frame.set_command("check_wordcount", self.wordcount_handler.check_wordcount)
        
        # 连接窗口关闭事件
        self.root.protocol("WM_DELETE_WINDOW", self._on_closing)
    
    def set_status(self, message):
        """
        设置状态信息
        
        参数:
            message: 状态信息
        """
        if self.status_label:
            self.status_label.config(text=message)
        
    def reset_for_processing(self):
        """准备开始处理"""
        self.processing = True
        self.progress_bar.configure(value=0)
        self.log_text.delete(1.0, tk.END)
        self.button_frame.enable_buttons(False)
        
    def enable_buttons(self):
        """启用按钮"""
        self.processing = False
        error_files = self.redirect.get_error_files() if self.redirect else []
        self.button_frame.enable_buttons(True, len(error_files) > 0)
        
    def conversion_complete(self):
        """转换完成后的处理"""
        self.set_status("处理完成！")
        self.enable_buttons()
        
    def conversion_error(self, error_message):
        """
        转换出错时的处理
        
        参数:
            error_message: 错误信息
        """
        self.log_text.insert('end', f"\n\n错误: {error_message}\n")
        self.log_text.see('end')
        self.set_status(f"处理出错: {error_message}")
        self.enable_buttons()
        
    def _on_closing(self):
        """窗口关闭时的处理"""
        if self.processing:
            if not messagebox.askokcancel("确认退出", "任务正在进行中，确定要退出吗？"):
                return
        
        # 恢复标准输出和错误输出
        sys.stdout = sys.__stdout__
        sys.stderr = sys.__stderr__
        
        # 销毁窗口
        self.root.destroy()
        
    def _show_about(self, event=None):
        """显示关于对话框"""
        about_text = """Word文档批量处理工具 v2.1
        
本工具用于批量处理Word文档，支持以下功能:
• 转换文档格式
• 提取图片
• 提取标题
• 检测字数

版权所有 © 2023-2025 先进制造与海洋学院关工委
        """
        messagebox.showinfo("关于", about_text)

    def _open_log_window(self):
        """在新窗口中查看日志"""
        log_window = tk.Toplevel(self.root)
        log_window.title("查看日志")
        log_window.minsize(800, 600)
        
        log_frame = ttk.Frame(log_window, padding=10)
        log_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建工具栏框架
        toolbar_frame = ttk.Frame(log_frame)
        toolbar_frame.pack(fill=tk.X, pady=(0, 5))
        
        # 添加复制按钮
        if TTKBOOTSTRAP_AVAILABLE:
            copy_btn = ttk.Button(toolbar_frame, text="复制全部", 
                                  command=lambda: self._copy_log_to_clipboard(log_text),
                                  bootstyle="info-outline")
        else:
            copy_btn = ttk.Button(toolbar_frame, text="复制全部", 
                                  command=lambda: self._copy_log_to_clipboard(log_text))
        copy_btn.pack(side=tk.LEFT, padx=5)
        
        # 添加保存按钮
        if TTKBOOTSTRAP_AVAILABLE:
            save_btn = ttk.Button(toolbar_frame, text="保存日志", 
                                 command=lambda: self._save_log_to_file(log_text),
                                 bootstyle="info-outline")
        else:
            save_btn = ttk.Button(toolbar_frame, text="保存日志", 
                                 command=lambda: self._save_log_to_file(log_text))
        save_btn.pack(side=tk.LEFT, padx=5)
        
        # 创建日志文本框
        log_text = ScrolledText(
            log_frame,
            wrap=tk.WORD, 
            width=100, 
            height=30, 
            font=("Consolas", 10)
        )
        log_text.pack(fill=tk.BOTH, expand=True)
        
        # 插入主窗口中的日志内容
        log_text.insert(tk.END, self.log_text.get(1.0, tk.END))
        
        # 居中显示窗口
        center_window(log_window, 800, 600)
        
    def _copy_log_to_clipboard(self, log_text):
        """复制日志内容到剪贴板"""
        self.root.clipboard_clear()
        self.root.clipboard_append(log_text.get(1.0, tk.END))
        self.set_status("日志已复制到剪贴板")
        
    def _save_log_to_file(self, log_text):
        """保存日志到文件"""
        file_path = filedialog.asksaveasfilename(
            defaultextension=".log",
            filetypes=[("日志文件", "*.log"), ("文本文件", "*.txt"), ("所有文件", "*.*")]
        )
        if file_path:
            try:
                with open(file_path, "w", encoding="utf-8") as f:
                    f.write(log_text.get(1.0, tk.END))
                self.set_status(f"日志已保存到 {file_path}")
            except Exception as e:
                messagebox.showerror("保存错误", f"保存日志时出错：{str(e)}")

def run_app():
    """运行应用程序"""
    # 使用兼容性层创建主窗口，不再指定主题
    root = create_window(
        title="Word文档批量处理工具",
        size=(900, 700), 
        resizable=True
    )
    
    # 初始化应用程序
    app = App(root)
    root.mainloop()
    
if __name__ == "__main__":
    run_app()