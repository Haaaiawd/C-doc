"""
提供增强版的Word文档处理功能，包括多线程处理和更好的用户界面反馈
"""
import os
import time
import tkinter as tk
from tkinter import messagebox
import threading

from gui.handlers.parallel_processor import ParallelProcessor
from word_processors import process_word_file
from file_utils import get_output_file_path

class EnhancedConversionHandler:
    """增强版Word文档转换处理器"""
    
    def __init__(self, app):
        """
        初始化增强版转换处理器
        
        参数:
            app: 主应用程序实例
        """
        self.app = app
        self.processor = ParallelProcessor()
        self.monitor_thread = None
        self.is_running = False
        
    def start_conversion(self):
        """开始转换处理"""
        # 获取UI组件
        file_frame = self.app.file_frame
        format_frame = self.app.format_frame
        
        # 获取源目录和文件
        input_dir = file_frame.get_source_dir()
        files = file_frame.get_selected_files()
        
        # 检查源目录和文件
        if not input_dir:
            messagebox.showerror("错误", "请先选择源文件目录！")
            return
            
        if not files or len(files) == 0:
            messagebox.showerror("错误", "请先选择要处理的文件！")
            return
            
        # 获取输出目录
        output_dir = file_frame.get_output_dir()
        if not output_dir:
            messagebox.showerror("错误", "请选择输出目录！")
            return
            
        # 确保输出目录存在
        os.makedirs(output_dir, exist_ok=True)
        
        # 获取处理配置
        process_config = {
            "suffix_enabled": format_frame.suffix_var.get(),
            "suffix_text": format_frame.suffix_entry.get(),
            "use_chinese_format": format_frame.chinese_format_var.get(),
            "keep_image_position": format_frame.keep_images_var.get(),
            "show_author_info": format_frame.show_author_var.get()
        }
        
        # 准备处理界面
        self.app.reset_for_processing()
        self.app.set_status("正在初始化处理...")
        self.app.log_text.insert('end', f"开始处理 {len(files)} 个文件...\n")
        self.app.log_text.see('end')
        
        # 显示暂停和取消按钮
        self.app.show_process_controls(True)
        self.is_running = True
        
        # 启动处理
        if self.processor.process_files(files, input_dir, output_dir, process_config):
            # 启动监控线程
            self.monitor_thread = threading.Thread(
                target=self._monitor_progress,
                daemon=True
            )
            self.monitor_thread.start()
        else:
            self.app.conversion_error("无法启动处理，可能有其他任务正在进行")
    
    def _monitor_progress(self):
        """监控处理进度的线程函数"""
        last_processed = 0
        update_interval = 0.5  # 更新界面的间隔时间（秒）
        
        try:
            while self.is_running:
                # 获取进度
                progress = self.processor.get_progress()
                processed = progress["processed"]
                total = progress["total"]
                percentage = progress["percentage"]
                
                # 更新进度条和状态
                self.app.root.after(0, lambda p=percentage: 
                    self.app.progress_bar.configure(value=p))
                
                status_text = f"正在处理... {processed}/{total} ({percentage}%)"
                if progress["paused"]:
                    status_text = f"已暂停 {processed}/{total} ({percentage}%)"
                    
                self.app.root.after(0, lambda s=status_text: 
                    self.app.set_status(s))
                
                # 检查是否有新的结果
                result = self.processor.get_next_result(timeout=update_interval)
                if result:
                    # 更新日志
                    if result["success"]:
                        log_text = f"√ 成功处理: {result['filename']}\n"
                    else:
                        error_msg = result["error"] or "未知错误"
                        log_text = f"× 处理失败: {result['filename']} - {error_msg}\n"
                        
                    self.app.root.after(0, lambda t=log_text: 
                        self._append_log(t))
                
                # 如果处理完成或已停止
                if not progress["running"]:
                    break
                    
                # 避免过于频繁的界面更新
                if processed == last_processed:
                    time.sleep(update_interval)
                    
                last_processed = processed
                
            # 处理完成
            self._process_completed()
                
        except Exception as e:
            self.app.root.after(0, lambda e=str(e): 
                self.app.conversion_error(f"监控进度时出错: {e}"))
                
    def _process_completed(self):
        """处理完成后的操作"""
        # 获取结果摘要
        summary = self.processor.get_result_summary()
        
        # 更新界面
        self.app.root.after(0, lambda: self._update_completion_ui(summary))
        
    def _update_completion_ui(self, summary):
        """更新完成后的界面"""
        # 更新进度条到100%
        self.app.progress_bar.configure(value=100)
        
        # 添加完成日志
        total = summary["total"]
        success = summary["success"]
        error = summary["error"]
        
        completion_text = (
            f"\n处理完成! 总计 {total} 文件，"
            f"成功 {success} 个, 失败 {error} 个。\n"
        )
        
        if error > 0:
            completion_text += "\n失败的文件:\n"
            for filename in summary["error_files"]:
                completion_text += f"- {filename}\n"
        
        self._append_log(completion_text)
        
        # 隐藏进程控制按钮
        self.app.show_process_controls(False)
        
        # 设置状态
        self.app.set_status(f"处理完成! 成功: {success}, 失败: {error}")
        
        # 恢复按钮状态
        self.is_running = False
        self.app.enable_buttons()
        
    def _append_log(self, text):
        """添加日志文本"""
        self.app.log_text.insert('end', text)
        self.app.log_text.see('end')
        
    def pause_conversion(self):
        """暂停处理"""
        if self.is_running:
            self.processor.pause()
            self.app.set_status("处理已暂停")
            self._append_log("处理已暂停...\n")
            
    def resume_conversion(self):
        """恢复处理"""
        if self.is_running:
            self.processor.resume()
            self.app.set_status("继续处理...")
            self._append_log("继续处理...\n")
            
    def cancel_conversion(self):
        """取消处理"""
        if self.is_running:
            if messagebox.askyesno("确认", "确定要取消当前处理吗?"):
                self.processor.stop()
                self.is_running = False
                self.app.set_status("处理已取消")
                self._append_log("\n处理已被用户取消\n")
                # 隐藏进程控制按钮
                self.app.show_process_controls(False)
                # 恢复按钮状态
                self.app.enable_buttons()