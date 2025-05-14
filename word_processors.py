"""
Word文档批量处理工具 - 主处理模块

提供统一的接口来处理Word文档，整合各个功能模块
"""
import os
import sys
import shutil

# 导入处理模块
from doc_processor import process_doc_file 
from docx_processor import process_docx_file
from image_extractor import extract_images_from_doc
from file_utils import extract_author_number, extract_author_from_filename, cleanup_temp_directory

def process_word_file(input_file, output_dir, suffix_enabled=True, 
                     suffix_text="——福州大学先进制造学院与海洋学院关工委2023年'中华魂'（毛泽东伟大精神品格）主题教育征文", 
                     use_chinese_format=False, keep_image_position=True, show_author_info=True,
                     mark_low_wordcount=False): # Added mark_low_wordcount parameter
    """
    处理单个Word文件，自动识别.doc或.docx格式
    
    参数:
        input_file: 输入文件路径
        output_dir: 输出目录
        suffix_enabled: 是否启用标题后缀
        suffix_text: 标题后缀内容
        use_chinese_format: 是否使用中文格式
        keep_image_position: 是否保持图片位置
        show_author_info: 是否显示作者信息
        mark_low_wordcount: 是否标记低字数文档
        
    返回:
        bool: 处理成功返回True，否则返回False
    """
    # 检查文件是否存在
    if not os.path.exists(input_file):
        print(f"× 错误：输入文件 '{input_file}' 不存在")
        return False
    
    # 根据文件扩展名选择相应的处理函数
    if input_file.lower().endswith('.doc'):
        # TODO: Update process_doc_file similarly if needed
        return process_doc_file(
            input_file, output_dir, suffix_enabled, suffix_text, 
            use_chinese_format, keep_image_position, show_author_info
            # Pass mark_low_wordcount=mark_low_wordcount when process_doc_file is updated
        )
    elif input_file.lower().endswith('.docx'):
        return process_docx_file(
            input_file, output_dir, suffix_enabled, suffix_text, 
            use_chinese_format, keep_image_position, show_author_info,
            mark_low_wordcount=mark_low_wordcount # Pass parameter
        )
    else:
        print(f"× 错误：不支持的文件格式 '{input_file}'")
        return False

def process_folder(input_folder, output_folder, suffix_enabled=True, 
                  suffix_text="——福州大学先进制造学院与海洋学院关工委2023年'中华魂'（毛泽东伟大精神品格）主题教育征文",
                  use_chinese_format=False, keep_image_position=True, show_author_info=True):
    """
    处理文件夹中的所有Word文档
    
    参数:
        input_folder: 输入文件夹
        output_folder: 输出文件夹
        suffix_enabled: 是否启用标题后缀
        suffix_text: 标题后缀内容
        use_chinese_format: 是否使用中文格式
        keep_image_position: 是否保持图片位置
        show_author_info: 是否显示作者信息
        
    返回:
        tuple: (成功处理文件数, 失败文件数)
    """
    # 确保输出文件夹存在
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    
    # 创建临时目录
    temp_dir = os.path.join(output_folder, "temp")
    os.makedirs(temp_dir, exist_ok=True)
    
    success_count = 0
    failure_count = 0
    
    try:
        # 获取所有Word文件
        word_files = [f for f in os.listdir(input_folder) 
                     if not f.startswith('~$') and (f.endswith('.docx') or f.endswith('.doc'))]
        
        # 按作者编号排序
        sorted_files = sorted(word_files, key=extract_author_number)
        
        print(f"找到 {len(sorted_files)} 个Word文件需要处理")
        
        for filename in sorted_files:
            # 设置当前处理的文件（用于日志）
            if hasattr(sys.stdout, 'set_current_file'):
                sys.stdout.set_current_file(filename)
            if hasattr(sys.stderr, 'set_current_file'):
                sys.stderr.set_current_file(filename)
                
            input_path = os.path.join(input_folder, filename)
            print(f"\n处理文件：{filename}")
            
            # 处理文档
            success = process_word_file(
                input_path, output_folder, suffix_enabled, suffix_text,
                use_chinese_format, keep_image_position, show_author_info
            )
            
            if success:
                success_count += 1
            else:
                failure_count += 1
                
    except Exception as e:
        print(f"× 处理文件夹时发生错误：{str(e)}")
    finally:
        # 清理临时文件
        cleanup_temp_directory(temp_dir)
        
        # 重置当前处理的文件
        if hasattr(sys.stdout, 'set_current_file'):
            sys.stdout.set_current_file(None)
        if hasattr(sys.stderr, 'set_current_file'):
            sys.stderr.set_current_file(None)
    
    return (success_count, failure_count)

# 导出需要的函数，使其他模块可以直接从 process_word 导入
__all__ = [
    'process_word_file', 
    'process_folder', 
    'extract_author_number',
    'extract_author_from_filename',
    'extract_images_from_doc'
]

# 使用示例
if __name__ == "__main__":
    input_folder = "G:\\PROJECTALL\\C-doc\\example"
    output_folder = "G:\\PROJECTALL\\C-doc\\output"
    
    process_folder(input_folder, output_folder)