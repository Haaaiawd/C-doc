"""
提供文件操作和辅助功能的模块
"""
import os
import re
import shutil

def extract_author_number(filename):
    """从文件名中提取作者数字"""
    match = re.search(r'(\d+)', filename)
    return int(match.group(1)) if match else float('inf')

def extract_author_from_filename(filename):
    """从文件名中提取作者名"""
    try:
        # 匹配文件名中852后面的数字，然后后面的2-4个汉字
        author_match = re.search(r'852\d+[^一-龥]*([一-龥]{2,4})', filename)
        if author_match:
            return author_match.group(1).strip()
    except Exception:
        pass
    return None

def sanitize_filename(filename):
    """清理文件名中的非法字符并限制长度"""
    # Windows非法字符: \ / : * ? " < > | 以及换行符
    invalid_chars = r'\/:*?"<>|\n'
    for char in invalid_chars:
        filename = filename.replace(char, "_")
    
    # 移除重复的连续下划线
    while '__' in filename:
        filename = filename.replace('__', '_')
    
    # 移除重复的——符号
    while '————' in filename:
        filename = filename.replace('————', '——')
    while '——' in filename and '——' != filename[:2]:
        filename = filename.replace('——', '—')
    
    # 限制文件名长度为180个字符(Windows路径最大长度为260,预留一些空间给路径)
    if len(filename) > 180:
        base, ext = os.path.splitext(filename)
        filename = base[:176] + ext  # 截断名称但保留扩展名
    
    return filename.strip()

def create_output_directories(output_dir):
    """创建输出目录结构"""
    # 创建成功文件文件夹
    success_dir = os.path.join(output_dir, "成功文件")
    no_image_dir = os.path.join(output_dir, "无图片成功文件")
    temp_dir = os.path.join(output_dir, "temp")
    
    os.makedirs(success_dir, exist_ok=True)
    os.makedirs(no_image_dir, exist_ok=True)
    os.makedirs(temp_dir, exist_ok=True)
    
    return {
        "success_dir": success_dir,
        "no_image_dir": no_image_dir,
        "temp_dir": temp_dir
    }

def cleanup_temp_directory(temp_dir):
    """清理临时目录"""
    try:
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)
    except Exception as e:
        print(f"! 清理临时文件时出错: {str(e)}")

def generate_output_filename(author_name, title, suffix_enabled, suffix_text, show_author_info, mark_low_wordcount=False): # Added mark_low_wordcount
    """
    根据作者名、标题和配置生成输出文件名
    
    参数:
        author_name: 作者名
        title: 原始标题
        suffix_enabled: 是否启用后缀
        suffix_text: 后缀文本
        show_author_info: 是否在文件名中包含作者信息
        mark_low_wordcount: 是否标记低字数文档
        
    返回:
        str: 清理和格式化后的文件名
    """
    # 清理标题和作者名中的非法字符
    clean_title = sanitize_filename(title)
    clean_author = sanitize_filename(author_name)
    
    # 添加字数不足标记
    prefix = "[字数不足]" if mark_low_wordcount else ""
    
    # 构建文件名
    if show_author_info:
        filename_base = f"{prefix}{clean_author}-{clean_title}"
    else:
        filename_base = f"{prefix}{clean_title}"
        
    # 添加后缀
    if suffix_enabled and suffix_text:
        clean_suffix = sanitize_filename(suffix_text)
        filename_base += clean_suffix
        
    # 限制文件名长度（例如，Windows通常限制为260个字符，包括路径）
    # 保留一些空间给路径和扩展名
    max_len = 200 
    if len(filename_base) > max_len:
        filename_base = filename_base[:max_len] + "..."
        
    return filename_base + ".docx"