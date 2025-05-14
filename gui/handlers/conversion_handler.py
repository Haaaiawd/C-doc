"""
提供文件转换相关的业务逻辑处理
"""
import os
import threading
import shutil
import tkinter as tk
from docx import Document
from word_processors import (
    process_word_file, 
    extract_author_number,
    extract_author_from_filename
)
from heading_utils import count_document_words

class ConversionHandler:
    """文件转换处理器，处理文件转换相关的业务逻辑"""
    
    def __init__(self, app):
        """
        初始化转换处理器
        
        参数:
            app: 主应用程序实例
        """
        self.app = app
        
    def validate_paths(self, input_dir, output_dir):
        """
        验证输入和输出路径是否有效
        
        参数:
            input_dir: 输入目录
            output_dir: 输出目录
            
        返回:
            tuple: (是否有效, 错误信息)
        """
        if not input_dir or not output_dir:
            return False, "请选择输入和输出文件夹"
            
        # 检查目录是否存在
        if not os.path.exists(input_dir):
            return False, "输入目录不存在"
            
        # 检查是否有 Word 文件
        input_files = [f for f in os.listdir(input_dir) if f.endswith(('.doc', '.docx'))]
        if not input_files:
            return False, "输入目录中没有 Word 文件"
            
        return True, ""
        
    def start_conversion(self):
        """开始文件转换处理"""
        # 获取路径配置
        paths = self.app.file_frame.get_paths()
        input_dir = paths["input_dir"]
        output_dir = paths["output_dir"]
        
        # 验证路径
        valid, error_msg = self.validate_paths(input_dir, output_dir)
        if not valid:
            self.app.set_status(error_msg)
            return
            
        # 获取标题后缀配置
        suffix_config = self.app.file_frame.get_suffix_config()
        suffix_enabled = suffix_config["enabled"]
        suffix_text = suffix_config["text"]
        
        # 获取格式配置 - 从新的位置获取各项配置
        format_config = self.app.format_frame.get_format_config()
        use_chinese_format = format_config["use_chinese_format"]
        
        # 从图片提取选项卡获取图片处理选项
        keep_image_position = not hasattr(self.app, 'extract_to_folder_var') or not self.app.extract_to_folder_var.get()
        
        # 从文档属性选项卡获取作者信息选项
        show_author_info = not hasattr(self.app, 'author_var') or not self.app.author_var.get()
        
        # 获取字数检测配置
        wordcount_config = self.app.wordcount_frame.get_wordcount_config()
        wordcount_enabled = wordcount_config["enabled"]
        min_words = wordcount_config["min_words"]
        mark_files = wordcount_config["mark_files"]
        move_files = wordcount_config["move_files"]
        
        # 准备开始转换处理
        self.app.reset_for_processing()
        
        # 向用户显示所选配置信息
        format_info = "中文标准格式" if use_chinese_format else "默认格式"
        self.app.log_text.insert('end', f"使用{format_info}处理文档\n")
        
        if not show_author_info:
            self.app.log_text.insert('end', "不添加作者信息\n")
            
        if keep_image_position:
            self.app.log_text.insert('end', "保持图片在原文位置\n")
        else:
            self.app.log_text.insert('end', "所有图片将移至文末\n")
        
        self.app.log_text.insert('end', "\n开始处理文件...\n\n")
        
        # 创建临时目录
        temp_dir = os.path.join(output_dir, "temp")
        os.makedirs(temp_dir, exist_ok=True)
        
        # 初始化字数不足文件夹路径为None
        low_wordcount_dir = None
        
        # 创建字数不足文件夹
        if wordcount_enabled and move_files: # Use move_files instead of wordcount_action
            low_wordcount_dir = os.path.join(output_dir, f"字数不足{min_words}字")
            os.makedirs(low_wordcount_dir, exist_ok=True)
        
        # 在新线程中运行转换
        thread = threading.Thread(
            target=self._conversion_thread, 
            args=(
                input_dir, output_dir, temp_dir,
                suffix_enabled, suffix_text,
                use_chinese_format, keep_image_position, show_author_info,
                wordcount_enabled, min_words, mark_files, move_files, low_wordcount_dir # Pass mark_files and move_files
            ),
            daemon=True
        )
        thread.start()
    
    def _conversion_thread(self, input_dir, output_dir, temp_dir,
                          suffix_enabled, suffix_text,
                          use_chinese_format, keep_image_position, show_author_info,
                          wordcount_enabled, min_words, mark_files, move_files, low_wordcount_dir): # Receive mark_files and move_files
        """转换处理线程"""
        try:
            # 获取目录中的所有文件并排序
            input_files = [f for f in os.listdir(input_dir) if f.endswith(('.doc', '.docx'))]
            sorted_files = sorted(input_files, key=lambda x: extract_author_number(x))
            
            total_files = len(sorted_files)
            self.app.set_status(f"开始处理，共 {total_files} 个文件...")
            
            # 配置进度条最大值
            self.app.root.after(0, lambda: self.app.progress_bar.configure(maximum=total_files))
            
            # 字数统计结果
            low_wordcount_files = []
            
            # 处理每个文件
            for index, filename in enumerate(sorted_files, 1):
                # 更新进度条
                current_progress = index / total_files * 100
                self.app.root.after(0, lambda p=index: self.app.progress_bar.configure(value=p))
                self.app.root.after(0, lambda: self.app.progress_bar.update())
                
                self.app.set_status(f"正在处理第 {index}/{total_files} 个文件 ({int(current_progress)}%)...")
                input_file = os.path.join(input_dir, filename)
                
                # 设置当前处理的文件
                self.app.redirect.set_current_file(filename)
                
                # 如果启用了字数检测，先检查字数
                word_count = None # Initialize word_count
                should_process = True # Flag to determine if process_word_file should run
                if wordcount_enabled:
                    try:
                        doc = Document(input_file)
                        word_count = count_document_words(doc)
                        
                        # 字数不足
                        if word_count < min_words:
                            self.app.log_text.insert('end', f"! {filename}: 字数为 {word_count}，不足 {min_words} 字\n")
                            low_wordcount_files.append((filename, word_count))
                            
                            # 如果选择移动字数不足文件到单独文件夹
                            if move_files: # Use move_files instead of wordcount_action
                                target_file = os.path.join(low_wordcount_dir, filename)
                                shutil.copy2(input_file, target_file)
                                should_process = False # Don't process further if moved
                                continue  # 跳过后续处理
                            # Note: Marking logic is handled within process_word_file based on parameters
                        else:
                            self.app.log_text.insert('end', f"✓ {filename}: 字数为 {word_count}，满足要求\n")
                    except Exception as e:
                        self.app.log_text.insert('end', f"! {filename}: 字数检测失败 - {str(e)}\n")
                
                # 根据文件扩展名选择处理方法 (only if not moved)
                if should_process:
                    # Determine if marking is needed for this file
                    mark_low_wordcount = (wordcount_enabled and 
                                          word_count is not None and 
                                          word_count < min_words and 
                                          mark_files and 
                                          not move_files) # Only mark if enabled, below threshold, marking is on, and not moving

                    process_word_file(
                        input_file, output_dir, suffix_enabled, suffix_text,
                        use_chinese_format, keep_image_position, show_author_info,
                        mark_low_wordcount=mark_low_wordcount # Pass the marking flag
                    )
            
            # 处理完成后显示统计信息
            self._show_summary(total_files, low_wordcount_files, wordcount_enabled)
            
            # 清理临时文件
            self._cleanup_temp_files(temp_dir)
            
        except Exception as e:
            # Capture the current value of e using a default argument
            self.app.root.after(0, lambda err=e: self.app.conversion_error(str(err)))
    
    def _show_summary(self, total_files, low_wordcount_files, wordcount_enabled):
        """显示处理结果摘要"""
        success_count = len(self.app.redirect.get_success_files())
        failed_count = len(self.app.redirect.get_error_files())
        low_count = len(low_wordcount_files)
        
        summary = f"\n处理完成！\n总计: {total_files} 个文件\n成功: {success_count} 个\n失败: {failed_count} 个\n"
        
        if wordcount_enabled:
            summary += f"字数不足: {low_count} 个\n"
            if low_count > 0:
                summary += "\n字数不足的文件:\n"
                for filename, count in low_wordcount_files:
                    summary += f"{filename}: {count} 字\n"
        
        if failed_count > 0:
            summary += "\n失败的文件作者数字:\n"
            for failed_file in self.app.redirect.get_error_files():
                author_num = extract_author_number(failed_file)
                summary += f"作者{author_num}\n"
        
        self.app.log_text.insert('end', summary)
        self.app.log_text.see('end')
        
        # 设置进度条为100%完成
        self.app.root.after(0, lambda: self.app.progress_bar.configure(value=total_files))
        
        # 更新状态
        self.app.root.after(0, lambda: self.app.conversion_complete())
    
    def _cleanup_temp_files(self, temp_dir):
        """清理临时文件"""
        try:
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir)
        except Exception as e:
            print(f"! 清理临时文件时出错: {str(e)}")
    
    def pack_error_files(self):
        """打包错误文件"""
        error_files = self.app.redirect.get_error_files()
        if not error_files:
            self.app.set_status("没有错误文件需要打包")
            return
        
        try:
            # 获取输出路径
            output_dir = self.app.file_frame.get_paths()["output_dir"]
            input_dir = self.app.file_frame.get_paths()["input_dir"]
            
            # 创建错误文件文件夹
            error_dir = os.path.join(output_dir, "错误文件")
            os.makedirs(error_dir, exist_ok=True)
            
            # 检查磁盘空间
            if shutil.disk_usage(error_dir).free < 1024 * 1024 * 100:  # 100MB
                self.app.set_status("磁盘空间不足")
                return
            
            # 复制错误文件
            copied_count = 0
            for filename in error_files:
                try:
                    src_path = os.path.join(input_dir, filename)
                    if os.path.exists(src_path):
                        dst_path = os.path.join(error_dir, filename)
                        shutil.copy2(src_path, dst_path)
                        copied_count += 1
                except Exception as e:
                    print(f"× 复制文件 {filename} 时出错：{str(e)}")
                
            self.app.set_status(f"已将 {copied_count} 个错误文件复制到 {error_dir}")
            # 复制完成后禁用打包按钮
            self.app.button_frame.pack_error_btn.state(['disabled'])
            
        except PermissionError:
            self.app.set_status("没有足够的权限创建或访问目录")
        except Exception as e:
            self.app.set_status(f"打包错误文件时出错: {str(e)}")