#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Word文档批量处理工具启动脚本
此脚本检查必要的依赖项并启动应用程序
"""
import sys
import os
import subprocess
import importlib.util
import tkinter as tk
import tkinter.messagebox as messagebox

# 必需的依赖包列表
REQUIRED_PACKAGES = [
    ("python-docx", "docx", "处理Word文档"),
    ("Pillow", "PIL", "图像处理"),
]
# 可选的依赖包
OPTIONAL_PACKAGES = [
    ("ttkbootstrap", "ttkbootstrap", "美化界面(可选)"),
    ("openpyxl", "openpyxl", "Excel报告生成(可选)")
]

def check_dependency(package_name, import_name):
    """检查依赖包是否已安装"""
    spec = importlib.util.find_spec(import_name)
    return spec is not None

def install_package(package_name):
    """安装依赖包"""
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", package_name])
        return True
    except subprocess.CalledProcessError:
        return False

def check_dependencies():
    """检查所有依赖包，并尝试安装缺少的包"""
    missing_packages = []
    optional_missing = []
    
    # 检查必需的包
    for package_info in REQUIRED_PACKAGES:
        package_name, import_name, description = package_info
        if not check_dependency(import_name, import_name):
            missing_packages.append((package_name, description))
    
    # 检查可选的包
    for package_info in OPTIONAL_PACKAGES:
        package_name, import_name, description = package_info
        if not check_dependency(import_name, import_name):
            optional_missing.append((package_name, description))
    
    # 如果有缺少的必需包，尝试安装
    if missing_packages:
        print("正在检查必要依赖项...")
        for package_name, description in missing_packages:
            print(f"正在安装 {package_name} ({description})...")
            if not install_package(package_name):
                return False, f"无法安装依赖项: {package_name}"
    
    # 显示可选包的安装信息
    if optional_missing:
        for package_name, description in optional_missing:
            print(f"提示: 可选依赖 {package_name} ({description}) 未安装")
            print(f"  您可以使用命令安装: pip install {package_name}")
    
    return True, "所有必需的依赖项已安装"

def show_gui_error(message):
    """显示图形界面错误消息"""
    try:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("启动错误", message)
        root.destroy()
    except:
        print(f"错误: {message}")

def start_app():
    """启动应用程序"""
    try:
        # 检查依赖项
        success, message = check_dependencies()
        if not success:
            show_gui_error(message)
            return
        
        # 导入主模块
        try:
            from gui.app import run_app
            # 启动应用程序
            run_app()
        except ImportError as e:
            show_gui_error(f"无法导入应用程序模块: {str(e)}")
        except Exception as e:
            show_gui_error(f"启动应用程序时出错: {str(e)}")
            
    except Exception as e:
        show_gui_error(f"检查依赖项时出错: {str(e)}")

if __name__ == "__main__":
    start_app()