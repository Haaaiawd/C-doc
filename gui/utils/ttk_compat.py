"""
标准ttk模块
使用标准ttk组件替代ttkbootstrap
"""
import tkinter as tk
from tkinter import ttk

# 不再尝试导入ttkbootstrap
TTKBOOTSTRAP_AVAILABLE = False
print("使用标准ttk组件")

# 定义常量，用于替代ttkbootstrap.constants
PRIMARY = "primary"
SECONDARY = "secondary"
SUCCESS = "success"
INFO = "info"
WARNING = "warning"
DANGER = "danger"
LIGHT = "light"
DARK = "dark"

# 创建样式类 - 使用标准ttk.Style
class Style(ttk.Style):
    """兼容的样式类，用于标准ttk"""
    def __init__(self, theme=None):
        super().__init__()
        if theme:
            try:
                self.theme_use(theme)
            except:
                pass

# Window类 - 兼容性包装器
class Window(tk.Tk):
    """兼容的窗口类，用于标准tk"""
    def __init__(self, title="Application", themename=None, size=None, resizable=True):
        super().__init__()
        self.title(title)
        if size:
            self.geometry(f"{size[0]}x{size[1]}")
        if isinstance(resizable, tuple) and len(resizable) == 2:
            self.resizable(resizable[0], resizable[1])
        else:
            self.resizable(resizable, resizable)

# 直接使用标准ttk组件
Notebook = ttk.Notebook
Frame = ttk.Frame
LabelFrame = ttk.LabelFrame
Label = ttk.Label
Button = ttk.Button
Checkbutton = ttk.Checkbutton
Radiobutton = ttk.Radiobutton
Progressbar = ttk.Progressbar
Combobox = ttk.Combobox
Spinbox = ttk.Spinbox
Entry = ttk.Entry
Treeview = ttk.Treeview
Separator = ttk.Separator
Scale = ttk.Scale
Scrollbar = ttk.Scrollbar
PanedWindow = ttk.PanedWindow
Sizegrip = ttk.Sizegrip

# 使用标准ScrolledText
from tkinter.scrolledtext import ScrolledText

# 创建主窗口的函数
def create_window(title="Application", themename=None, size=None, resizable=True):
    """创建主窗口，使用标准tk.Tk"""
    # 使用标准Tk
    root = tk.Tk()
    root.title(title)
    if size:
        root.geometry(f"{size[0]}x{size[1]}")
    if isinstance(resizable, tuple) and len(resizable) == 2:
        root.resizable(resizable[0], resizable[1])
    else:
        root.resizable(resizable, resizable)
    
    return root