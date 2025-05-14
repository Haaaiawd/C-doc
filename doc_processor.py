"""
提供.doc格式Word文件的处理功能
"""
import os
import sys
import shutil
from docx import Document

def process_doc_file(input_file, output_dir, suffix_enabled=True, 
                    suffix_text="——福州大学先进制造学院与海洋学院关工委2023年'中华魂'（毛泽东伟大精神品格）主题教育征文", 
                    use_chinese_format=False, keep_image_position=True, show_author_info=True):
    """
    处理 doc 格式的 Word 文件
    
    参数:
        input_file: 输入文件路径
        output_dir: 输出目录
        suffix_enabled: 是否启用标题后缀
        suffix_text: 标题后缀内容
        use_chinese_format: 是否使用中文格式
        keep_image_position: 是否保持图片位置
        show_author_info: 是否显示作者信息
    
    返回:
        bool: 处理成功返回True，否则返回False
    """
    try:
        # 导入文件工具模块
        from file_utils import extract_author_number, extract_author_from_filename
        
        # 检查依赖
        try:
            import win32com.client
            import pythoncom
        except ImportError:
            print(f"× 错误：处理doc文件需要pywin32库，请安装：pip install pywin32")
            return False
            
        # 在新线程中调用COM对象前需初始化COM
        pythoncom.CoInitialize()
        
        # 确保输出目录存在
        os.makedirs(output_dir, exist_ok=True)
        
        # 提取作者信息
        filename = os.path.basename(input_file)
        author_num = extract_author_number(filename)
        author_name = extract_author_from_filename(filename)
        
        print(f"DEBUG: 开始处理doc文件 {input_file}")
        
        word = None
        doc = None
        
        try:
            # 创建临时目录
            temp_dir = os.path.join(output_dir, "temp")
            os.makedirs(temp_dir, exist_ok=True)
            
            # 创建 Word 应用程序 COM 对象
            word = win32com.client.Dispatch("Word.Application")
            
            # 尝试设置可见性
            try:
                word.Visible = 0
            except:
                print("! 无法设置Word应用程序可见性，继续处理...")
                pass
            
            # 打开文档
            try:
                doc = word.Documents.Open(os.path.abspath(input_file))
            except Exception as e:
                print(f"× 打开doc文件失败: {str(e)}")
                if word:
                    word.Quit()
                pythoncom.CoUninitialize()
                return False
            
            # 构建输出文件路径
            output_filename = f"temp_{author_num}.docx"
            output_file_path = os.path.join(temp_dir, output_filename)
            
            # 另存为 docx 格式
            try:
                doc.SaveAs2(os.path.abspath(output_file_path), FileFormat=16)  # 16 代表 docx 格式
            except Exception as e:
                print(f"× 转换doc文件失败: {str(e)}")
                if doc:
                    doc.Close(SaveChanges=0)
                if word:
                    word.Quit()
                pythoncom.CoUninitialize()
                return False
            
            # 关闭文档和Word应用程序
            doc.Close(SaveChanges=0)  # 不保存更改，因为已经另存为
            doc = None
            word.Quit()
            word = None
            
            print(f"✓ 已成功将 doc 文件转换为 docx: {output_filename}")
            
            # 导入 docx 处理模块
            from docx_processor import process_docx_file
            
            # 继续处理转换后的 docx 文件
            success = process_docx_file(output_file_path, output_dir, suffix_enabled, suffix_text, 
                                       use_chinese_format, keep_image_position, show_author_info)
            
            # 删除临时文件
            try:
                if os.path.exists(output_file_path):
                    os.remove(output_file_path)
            except:
                pass  # 忽略临时文件删除错误
            
            pythoncom.CoUninitialize()  # 释放COM资源
            return success
            
        except Exception as e:
            # 处理可能的错误
            print(f"× 处理doc文件时出错：{str(e)}")
            
            # 确保资源被释放
            if doc:
                try:
                    doc.Close(SaveChanges=0)
                except:
                    pass
            if word:
                try:
                    word.Quit()
                except:
                    pass
                    
            pythoncom.CoUninitialize()  # 释放COM资源
            return False
    except Exception as e:
        print(f"× 处理doc文件时出现错误：{str(e)}")
        return False