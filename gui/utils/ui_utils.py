"""
提供GUI相关的通用工具函数
"""
import tkinter as tk
import os
import sys
from gui.utils.ttk_compat import *  # 使用兼容性模块
from tkinter import TclError

def create_tooltip(widget, text):
    """
    为控件创建悬浮提示，使用ttkbootstrap样式
    
    参数:
        widget: 目标控件
        text: 提示文字
    """
    tooltip = None
    
    def enter(event):
        nonlocal tooltip
        try:
            x, y, _, _ = widget.bbox("insert")
        except (TclError, TypeError):
            # 非文本控件处理
            x = 0
            y = widget.winfo_height()
            
        x += widget.winfo_rootx() + 25
        y += widget.winfo_rooty() + 25
        
        # 创建提示窗口
        tooltip = tk.Toplevel(widget)
        tooltip.wm_overrideredirect(True)
        tooltip.wm_geometry(f"+{x}+{y}")
        
        # 使用ttkbootstrap样式
        frame = ttk.Frame(tooltip, relief="raised") # Removed bootstyle="light"
        frame.pack(fill=tk.BOTH, expand=True)
        
        label = ttk.Label(
            frame, 
            text=text, 
            # Removed bootstyle="light"
            wraplength=300,
            justify=tk.LEFT,
            padding=(6, 4)
        )
        label.pack(fill=tk.BOTH, padx=2, pady=2)
        
        # 设置圆角边框
        tooltip.update_idletasks()
        tooltip.attributes("-alpha", 0.9)  # 设置半透明效果
        
    def leave(event):
        nonlocal tooltip
        if tooltip:
            tooltip.destroy()
            tooltip = None
            
    widget.bind("<Enter>", enter)
    widget.bind("<Leave>", leave)

def setup_drag_drop(widget, callback, types=("files", "dirs")):
    """
    设置控件的拖放支持
    
    参数:
        widget: 目标控件（通常是Entry）
        callback: 拖放回调函数，接收路径参数
        types: 允许拖放的类型，可以是"files"、"dirs"或两者都有
    """
    if not isinstance(types, (list, tuple)):
        types = [types]
    
    def drop(event):
        # 提取拖放的数据
        try:
            data = event.data
            # Windows系统下可能包含花括号和多个文件，需要处理
            if data.startswith("{") and data.endswith("}"):
                data = data[1:-1]
            
            # 处理拖放的文件或目录
            if "files" in types and os.path.isfile(data):
                callback(data)
            elif "dirs" in types and os.path.isdir(data):
                callback(data)
            
        except Exception as e:
            print(f"! 拖放处理出错: {str(e)}")
    
    # 绑定拖放事件
    try:
        # 为控件注册拖放功能
        widget.drop_target_register("*")
        widget.dnd_bind("<<Drop>>", drop)
        # 改变控件样式提示可拖放
        widget.bind("<DragEnter>", lambda e: widget.state(["focus"]) if hasattr(widget, "state") else None)
        widget.bind("<DragLeave>", lambda e: widget.state(["!focus"]) if hasattr(widget, "state") else None)
        
        return True
    except (TclError, AttributeError):
        # 控件不支持拖放或TkDnD未安装
        print("! 控件不支持拖放功能，可能需要安装TkDnD")
        return False

def set_icon(window):
    """
    设置窗口图标
    
    参数:
        window: 目标窗口
    """
    try:
        # 获取应用程序目录
        if getattr(sys, 'frozen', False):
            # 如果是打包后的可执行文件
            base_dir = os.path.dirname(sys.executable)
        else:
            # 如果是脚本运行
            base_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
            
        icon_path = os.path.join(base_dir, "app.ico")
        if os.path.exists(icon_path):
            window.iconbitmap(icon_path)
    except Exception as e:
        print(f"! 设置应用程序图标时出错: {str(e)}")
        
def center_window(window, width, height):
    """
    将窗口居中显示
    
    参数:
        window: 目标窗口
        width: 窗口宽度
        height: 窗口高度
    """
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    
    x = (screen_width - width) // 2
    y = (screen_height - height) // 2
    
    window.geometry(f'{width}x{height}+{x}+{y}')

def show_module_requirement(module_name, feature_name=None):
    """
    显示模块需求信息并提供安装提示
    
    参数:
        module_name: 模块名称
        feature_name: 功能名称（可选）
    """
    feature_text = f"{feature_name}功能" if feature_name else "此功能"
    
    msg = (
        f"{feature_text}需要安装{module_name}模块。\n\n"
        f"您可以使用以下命令安装：\n"
        f"pip install {module_name}"
    )
    
    top = tk.Toplevel()
    top.title("需要安装模块")
    top.resizable(False, False)
    
    frame = ttk.Frame(top, padding=20)
    frame.pack(fill=tk.BOTH, expand=True)
    
    ttk.Label(
        frame, 
        text=msg, 
        wraplength=300,
        justify=tk.LEFT
    ).pack(pady=(0, 10))
    
    btn_frame = ttk.Frame(frame)
    btn_frame.pack(fill=tk.X)
    
    ttk.Button(
        btn_frame,
        text="确定",
        command=top.destroy
    ).pack(side=tk.RIGHT)
    
    center_window(top, 350, 180)
    
    return top

# 为了兼容性，创建别名
set_window_icon = set_icon