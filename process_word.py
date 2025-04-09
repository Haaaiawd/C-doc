from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import re
import os
import sys
from zipfile import BadZipFile
from docx.oxml.ns import qn
import shutil
from docx2python import docx2python
try:
    from PIL import Image
except ImportError:
    # PIL 模块可选，用于图片处理
    pass

# 导入样式模块
from document_styles import (
    create_document_with_styles, 
    apply_title_style, 
    apply_subtitle_style, 
    apply_author_style, 
    apply_body_style,
    # 导入中文格式样式函数
    create_chinese_formal_document,
    apply_chinese_main_title,
    apply_chinese_subtitle,
    apply_chinese_body,
    apply_chinese_heading1,
    apply_chinese_heading2,
    apply_chinese_heading34
)

# 导入简单的标题识别功能
from heading_utils import is_heading1

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

def has_images_in_doc(doc):
    """
    检查Word文档中是否包含图片
    """
    try:
        for para in doc.paragraphs:
            for run in para.runs:
                try:
                    # 检查内联图片
                    if run._element.findall('.//wp:inline', namespaces={'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'}):
                        return True
                    # 检查锚定图片
                    if run._element.findall('.//wp:anchor', namespaces={'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'}):
                        return True
                    # 检查图形对象
                    if run._element.findall('.//w:drawing', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}):
                        return True
                except Exception:
                    # 忽略单个图片的检查错误
                    continue
        return False
    except Exception:
        # 如果整个检查过程出错，假设文档包含图片
        return True

def extract_images_from_doc(input_path, temp_dir):
    """
    直接从Word文档关系中提取图片
    """
    try:
        doc = Document(input_path)
        temp_image_files = []
        
        # 获取文档中的所有关系
        rels = doc.part.rels
        
        for rel in rels.values():
            # 检查是否是图片
            if "image" in rel.target_ref:
                try:
                    # 获取图片数据
                    image_part = rel.target_part
                    image_data = image_part.blob
                    
                    # 从目标引用中获取图片扩展名
                    image_ext = os.path.splitext(rel.target_ref)[1]
                    if not image_ext or image_ext.lower() not in ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tif', '.tiff']:
                        image_ext = '.png'  # 默认使用png
                    
                    # 保存图片
                    temp_image_path = os.path.join(temp_dir, f"image_{len(temp_image_files)}{image_ext}")
                    with open(temp_image_path, 'wb') as f:
                        f.write(image_data)
                    temp_image_files.append(temp_image_path)
                except Exception as e:
                    print(f"保存图片时出错：{str(e)}")
                    continue
        
        return temp_image_files
    except Exception as e:
        print(f"提取图片时出错：{str(e)}")
        return []

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

def process_word_file(input_file, output_dir, suffix_enabled=True, suffix_text="——福州大学先进制造学院与海洋学院关工委2023年'中华魂'（毛泽东伟大精神品格）主题教育征文", use_chinese_format=False):
    """处理单个Word文件"""
    print(f"DEBUG: 开始处理文件 {input_file}")
    try:
        # 检查文件是否存在
        if not os.path.exists(input_file):
            print(f"× 错误：输入文件 '{input_file}' 不存在")
            return False

        # 创建成功文件文件夹
        success_dir = os.path.join(output_dir, "成功文件")
        no_image_dir = os.path.join(output_dir, "无图片成功文件")
        os.makedirs(success_dir, exist_ok=True)
        os.makedirs(no_image_dir, exist_ok=True)

        # 创建临时文件夹存储图片
        temp_dir = os.path.join(output_dir, "temp_images")
        os.makedirs(temp_dir, exist_ok=True)

        # 打开文档
        try:
            doc = Document(input_file)
        except BadZipFile:
            print(f"× 错误：文件 '{input_file}' 可能已损坏或不是有效的Word文档")
            return False

        # 根据格式选择创建文档方式
        if use_chinese_format:
            new_doc = create_chinese_formal_document()
        else:
            new_doc = create_document_with_styles()
            
        temp_image_files = []
        author_name = ""
        original_title = ""
        title_found = False
        author_added = False  # 添加标志，防止重复添加作者信息
        used_default_author = False  # 添加标志，表示是否使用了默认作者名

        # 从文件名中提取作者名
        filename = os.path.basename(input_file)
        author_name = extract_author_from_filename(filename)

        # 提取图片
        temp_image_files = extract_images_from_doc(input_file, temp_dir)
        has_images = len(temp_image_files) > 0

        # 处理文档内容
        for para in doc.paragraphs:
            try:
                text = para.text.strip()
                if not text:
                    continue

                # 提取标题（第一个非空段落）
                if not title_found:
                    original_title = text
                    title_para = new_doc.add_paragraph(original_title)
                    if use_chinese_format:
                        apply_chinese_main_title(new_doc, title_para)
                    else:
                        apply_title_style(new_doc, title_para)
                    
                    # 根据用户配置添加副标题
                    if suffix_enabled and suffix_text:
                        subtitle_para = new_doc.add_paragraph(suffix_text)
                        if use_chinese_format:
                            apply_chinese_subtitle(new_doc, subtitle_para)
                        else:
                            apply_subtitle_style(new_doc, subtitle_para)
                    
                    title_found = True
                    continue

                # 如果从文件名中没有提取到作者名，则尝试从文档内容中提取
                if not author_name and text.startswith('852'):
                    author_match = re.search(r'852\d*[^-]*-([^-\d\W]+)', text)
                    if not author_match:
                        author_match = re.search(r'852\d*[\s-]*([^\d\W]+)', text)
                    if author_match:
                        author_name = author_match.group(1).strip()

                # 添加作者信息（只添加一次）
                if author_name and not author_added and not text.startswith('852'):
                    author_text = f"（先进制造学院与海洋学院关工委通讯员{author_name}）"
                    author_para = new_doc.add_paragraph(author_text)
                    if use_chinese_format:
                        apply_chinese_subtitle(new_doc, author_para)  # 使用副标题样式
                    else:
                        apply_author_style(new_doc, author_para)
                    author_added = True
                    continue

                # 处理正文（跳过学号行）
                if title_found and not text.startswith('852'):
                    body_para = new_doc.add_paragraph(text)
                    if use_chinese_format:
                        # 简化标题处理，只处理主标题和正文
                        # 不再使用复杂的标题识别函数
                        apply_chinese_body(new_doc, body_para)
                    else:
                        apply_body_style(new_doc, body_para)

            except Exception as e:
                print(f"× 处理段落时出错：{str(e)}")
                continue

        # 在文档末尾添加图片
        if temp_image_files:
            new_doc.add_paragraph()  # 添加空行
            for image_path in temp_image_files:
                try:
                    img_para = new_doc.add_paragraph()
                    img_para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                    run = img_para.add_run()
                    
                    # 调整图片大小到适合的宽度，同时保持比例
                    try:
                        from PIL import Image
                        with Image.open(image_path) as img:
                            width, height = img.size
                            max_width = Inches(6)
                            if width > max_width.pt:
                                ratio = max_width.pt / width
                                run.add_picture(image_path, width=max_width)
                            else:
                                run.add_picture(image_path)
                    except ImportError:
                        # 如果PIL没有安装，使用固定宽度
                        run.add_picture(image_path, width=Inches(6))
                except Exception as e:
                    print(f"× 添加图片时出错：{str(e)}")
                    continue

        # 如果没有提取到作者名，使用默认值"佚名"
        if not author_name:
            author_name = "佚名"
            used_default_author = True
            print(f"! 警告：未能提取作者名，使用默认值\"{author_name}\"")
            
            # 如果还没有添加作者信息，添加默认作者信息
            if not author_added:
                author_text = f"（先进制造学院与海洋学院关工委通讯员{author_name}）"
                author_para = new_doc.add_paragraph(author_text)
                apply_author_style(new_doc, author_para)
                author_added = True

        # 保存新文档
        if original_title:
            # 防止副标题在原标题中重复出现
            if suffix_enabled and suffix_text and suffix_text in original_title:
                original_title = original_title.split(suffix_text)[0].strip()
            
            # 根据用户配置生成文件名
            if suffix_enabled and suffix_text:
                new_filename = f"({author_name}){original_title}{suffix_text}.docx"
            else:
                new_filename = f"({author_name}){original_title}.docx"
            
            # 清理文件名
            new_filename = sanitize_filename(new_filename)
            # 根据是否有图片选择保存目录
            output_dir_final = success_dir if has_images else no_image_dir
            output_file = os.path.join(output_dir_final, new_filename)
            
            try:
                new_doc.save(output_file)
                if used_default_author:
                    if has_images:
                        print(f"✓ 文件处理完成（使用默认作者名）：{new_filename}")
                    else:
                        print(f"✓ 文件处理完成（使用默认作者名，无图片）：{new_filename}")
                else:
                    if has_images:
                        print(f"✓ 文件处理完成：{new_filename}")
                    else:
                        print(f"✓ 文件处理完成（无图片）：{new_filename}")
            except Exception as e:
                print(f"× 保存文件时出错：{str(e)}")
                return False
        else:
            print("× 未能提取标题")
            return False

        # 清理临时文件
        for temp_file in temp_image_files:
            try:
                os.remove(temp_file)
            except:
                pass
        try:
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir)
        except:
            pass

        return True

    except Exception as e:
        print(f"× 处理文件时出现错误：{str(e)}")
        return False

# 新增：判断标题级别的辅助函数
def is_heading1(text):
    """判断文本是否为一级标题"""
    # 判断依据可以根据实际情况调整，比如检查是否有特定前缀，是否较短等
    return text.strip().startswith(('一、', '二、', '三、', '四、', '五、', '六、', '七、', '八、', '九、', '十、')) and len(text) < 40

def is_heading2(text):
    """判断文本是否为二级标题"""
    return text.strip().startswith(('(一)', '(二)', '(三)', '(四)', '(五)', '(六)', '(七)', '(八)', '(九)', '(十)', 
                                   '（一）', '（二）', '（三）', '（四）', '（五）', '（六）', '（七）', '（八）', '（九）', '（十）')) and len(text) < 40

def is_heading3_or_4(text):
    """判断文本是否为三级或四级标题"""
    return text.strip().startswith(('1.', '2.', '3.', '4.', '5.', '6.', '7.', '8.', '9.', '10.', 
                                   '1、', '2、', '3、', '4、', '5、', '6、', '7、', '8、', '9、', '10、')) and len(text) < 40

def open_word_doc(input_path):
    """
    安全地打开Word文档
    返回：(doc对象, 错误信息) 元组
    """
    try:
        # 首先尝试直接打开
        doc = Document(input_path)
        return doc, None
    except Exception as e:
        try:
            # 如果直接打开失败，尝试以二进制模式打开
            with open(input_path, 'rb') as f:
                doc = Document(f)
                return doc, None
        except Exception as e:
            error_msg = str(e)
            if "Bad CRC-32" in error_msg:
                return None, "文件中的图片可能已损坏"
            elif "Package not found" in error_msg:
                return None, "文件格式错误或已损坏"
            else:
                return None, f"无法打开文件：{error_msg}"

def process_folder(input_folder, output_folder):
    """处理文件夹中的所有Word文档"""
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    
    temp_dir = os.path.join(output_folder, "temp_images")
    
    try:
        for filename in os.listdir(input_folder):
            if filename.startswith('~$') or not filename.endswith('.docx'):
                continue
                
            if hasattr(sys.stdout, 'set_current_file'):
                sys.stdout.set_current_file(filename)
            if hasattr(sys.stderr, 'set_current_file'):
                sys.stderr.set_current_file(filename)
                
            input_path = os.path.join(input_folder, filename)
            print(f"\n处理文件：{filename}")
            
            # 处理文档
            process_word_file(input_path, output_folder)
                    
    except Exception as e:
        print(f"× 处理文件夹时发生错误：{str(e)}")
    finally:
        if os.path.exists(temp_dir):
            try:
                shutil.rmtree(temp_dir)
            except:
                pass
        if hasattr(sys.stdout, 'set_current_file'):
            sys.stdout.set_current_file(None)
        if hasattr(sys.stderr, 'set_current_file'):
            sys.stderr.set_current_file(None)

def process_doc_file(input_file, output_dir, suffix_enabled=True, suffix_text="——福州大学先进制造学院与海洋学院关工委2023年'中华魂'（毛泽东伟大精神品格）主题教育征文", use_chinese_format=False):
    """
    处理 doc 格式的 Word 文件
    """
    import win32com.client
    import os
    import pythoncom
    
    # 在新线程中调用COM对象前需初始化COM
    pythoncom.CoInitialize()
    
    # 确保输出目录存在
    os.makedirs(output_dir, exist_ok=True)
    
    # 提取作者信息
    author_num = extract_author_number(os.path.basename(input_file))
    author_name = extract_author_from_filename(os.path.basename(input_file))
    
    word = None
    doc = None
    
    try:
        # 创建 Word 应用程序 COM 对象
        word = win32com.client.Dispatch("Word.Application")
        
        # 尝试设置可见性，但如果失败则忽略该错误
        try:
            word.Visible = 0
        except:
            print("! 无法设置Word应用程序可见性，继续处理...")
            pass
        
        # 打开文档
        doc = word.Documents.Open(os.path.abspath(input_file))
        
        # 构建输出文件路径
        temp_dir = os.path.join(output_dir, "temp")
        os.makedirs(temp_dir, exist_ok=True)
        output_filename = f"作者{author_num}({author_name}).docx"
        output_file_path = os.path.join(temp_dir, output_filename)
        
        # 另存为 docx 格式
        doc.SaveAs2(os.path.abspath(output_file_path), FileFormat=16)  # 16 代表 docx 格式
        
        # 关闭文档
        doc.Close(SaveChanges=0)  # 不保存更改，因为已经另存为
        doc = None
        
        # 关闭Word应用程序
        word.Quit()
        word = None
        
        print(f"✓ 已成功将 doc 文件转换为 docx: {output_filename}")
        
        # 继续处理转换后的 docx 文件，并传递标题后缀配置和格式选择
        success = process_word_file(output_file_path, output_dir, suffix_enabled, suffix_text, use_chinese_format)
        
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
        error_msg = str(e)
        if "Property 'Word.Application.Visible' can not be set" in error_msg:
            # 如果错误是关于Visible属性，尝试不设置可见性继续处理
            try:
                if doc is None and word is not None:
                    # 尝试打开文档
                    doc = word.Documents.Open(os.path.abspath(input_file))
                    
                    # 构建输出文件路径
                    temp_dir = os.path.join(output_dir, "temp")
                    os.makedirs(temp_dir, exist_ok=True)
                    output_filename = f"作者{author_num}({author_name}).docx"
                    output_file_path = os.path.join(temp_dir, output_filename)
                    
                    # 另存为 docx 格式
                    doc.SaveAs2(os.path.abspath(output_file_path), FileFormat=16)
                    
                    # 关闭文档
                    doc.Close(SaveChanges=0)
                    word.Quit()
                    
                    print(f"✓ 已成功将 doc 文件转换为 docx: {output_filename}")
                    
                    # 继续处理转换后的 docx 文件，并传递标题后缀配置
                    success = process_word_file(output_file_path, output_dir, suffix_enabled, suffix_text, use_chinese_format)
                    
                    # 删除临时文件
                    try:
                        if os.path.exists(output_file_path):
                            os.remove(output_file_path)
                    except:
                        pass
                    
                    pythoncom.CoUninitialize()
                    return success
            except Exception as inner_e:
                error_msg = f"{error_msg}\n尝试继续处理失败: {str(inner_e)}"
        
        # 清理资源
        try:
            # 尝试关闭可能打开的Word实例
            if doc is not None:
                try:
                    doc.Close(SaveChanges=0)
                except:
                    pass
            if word is not None:
                try:
                    word.Quit()
                except:
                    pass
        except:
            pass
            
        # 保证COM资源释放
        try:
            pythoncom.CoUninitialize()
        except:
            pass
        
        print(f"× 处理 doc 文件时出错: {error_msg}")
        return False

# 使用示例
if __name__ == "__main__":
    input_folder = r"C:\Users\86159\Desktop\新建文件夹 (2)"
    output_folder = r"C:\Users\86159\Desktop\新建文件夹 (2)"
    
    process_folder(input_folder, output_folder)