"""
Word文档批量处理工具GUI模块
提供图形用户界面和相关组件
"""

# 导出主应用程序
from gui.app import App, run_app

# 导出框架组件
from gui.frames.file_frame import FileFrame
from gui.frames.format_frame import FormatFrame
from gui.frames.wordcount_frame import WordcountFrame
from gui.frames.button_frame import ButtonFrame

# 导出处理器
from gui.handlers.conversion_handler import ConversionHandler
from gui.handlers.image_handler import ImageHandler
from gui.handlers.wordcount_handler import WordcountHandler
from gui.handlers.title_handler import TitleHandler

# 导出工具
from gui.utils.redirect_text import RedirectText
from gui.utils.ui_utils import set_window_icon, center_window, create_tooltip

# 导出主应用类
__all__ = ['App', 'run_app', 'FileFrame', 'FormatFrame', 'WordcountFrame', 'ButtonFrame', 'ConversionHandler', 'ImageHandler', 'WordcountHandler', 'TitleHandler', 'RedirectText', 'set_window_icon', 'center_window', 'create_tooltip']