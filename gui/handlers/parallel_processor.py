"""
提供并行处理功能的处理器，用于多线程处理多个Word文档
"""

import os
import queue
import threading
import time
from concurrent.futures import ThreadPoolExecutor
from typing import List, Dict, Any, Optional

from word_processors import process_word_file
from file_utils import get_output_file_path

class ParallelProcessor:
    """
    并行处理器类，提供多线程并行处理功能
    """
    
    def __init__(self, max_workers=None):
        """
        初始化并行处理器
        
        参数:
            max_workers: 最大工作线程数，默认为CPU核心数的2倍
        """
        self.max_workers = max_workers or (os.cpu_count() * 2)
        self._executor = None
        self._lock = threading.Lock()
        self._results_queue = queue.Queue()
        self._total_files = 0
        self._processed_files = 0
        self._paused = False
        self._running = False
        self._success_count = 0
        self._error_count = 0
        self._error_files = []
        
    def process_files(self, files: List[str], input_dir: str, 
                      output_dir: str, config: Dict[Any, Any]) -> bool:
        """
        开始并行处理文件
        
        参数:
            files: 需要处理的文件列表
            input_dir: 输入目录
            output_dir: 输出目录
            config: 处理配置
            
        返回:
            是否成功启动处理
        """
        with self._lock:
            # 如果已经在运行，则不再启动新的处理
            if self._running:
                return False
                
            self._running = True
            self._paused = False
            
            # 重置处理状态
            self._total_files = len(files)
            self._processed_files = 0
            self._success_count = 0
            self._error_count = 0
            self._error_files = []
            
            # 清空结果队列
            while not self._results_queue.empty():
                self._results_queue.get()
            
            # 创建线程池
            self._executor = ThreadPoolExecutor(max_workers=self.max_workers)
            
            # 提交所有任务到线程池
            for file_name in files:
                self._executor.submit(
                    self._process_single_file,
                    file_name,
                    input_dir,
                    output_dir,
                    config
                )
            
            # 启动监控线程
            monitor_thread = threading.Thread(
                target=self._monitor_executor,
                daemon=True
            )
            monitor_thread.start()
            
            return True
    
    def _process_single_file(self, file_name: str, input_dir: str, 
                            output_dir: str, config: Dict[Any, Any]):
        """处理单个文件的工作函数"""
        # 检查是否需要暂停
        while self._paused and self._running:
            time.sleep(0.5)
            
        # 如果已停止，则不再处理
        if not self._running:
            return
            
        result = {
            "filename": file_name,
            "success": False,
            "error": None
        }
        
        try:
            # 构建输入和输出文件路径
            input_path = os.path.join(input_dir, file_name)
            
            # 根据后缀配置构建输出文件名
            if config.get("suffix_enabled", False):
                suffix = config.get("suffix_text", "")
                base_name, ext = os.path.splitext(file_name)
                output_file = f"{base_name}{suffix}{ext}"
            else:
                output_file = file_name
                
            output_path = os.path.join(output_dir, output_file)
            
            # 调用处理函数
            process_word_file(
                input_path, 
                output_path,
                use_chinese_format=config.get("use_chinese_format", False),
                keep_image_position=config.get("keep_image_position", True),
                show_author_info=config.get("show_author_info", False)
            )
            
            # 标记处理成功
            result["success"] = True
            
            # 更新成功计数
            with self._lock:
                self._success_count += 1
            
        except Exception as e:
            # 处理失败，记录错误
            result["error"] = str(e)
            
            # 更新失败计数
            with self._lock:
                self._error_count += 1
                self._error_files.append(file_name)
        
        finally:
            # 更新处理计数并将结果放入队列
            with self._lock:
                self._processed_files += 1
                self._results_queue.put(result)
    
    def _monitor_executor(self):
        """监控线程池执行状态"""
        # 等待线程池任务完成
        try:
            self._executor.shutdown(wait=True)
        finally:
            with self._lock:
                self._running = False
                
    def pause(self):
        """暂停处理"""
        with self._lock:
            self._paused = True
            
    def resume(self):
        """继续处理"""
        with self._lock:
            self._paused = False
            
    def stop(self):
        """停止处理"""
        with self._lock:
            self._running = False
            # 不阻塞立即关闭线程池
            if self._executor:
                self._executor.shutdown(wait=False, cancel_futures=True)
                
    def get_progress(self) -> Dict[str, Any]:
        """
        获取处理进度
        
        返回:
            包含进度信息的字典
        """
        with self._lock:
            total = self._total_files
            processed = self._processed_files
            percentage = int((processed / total) * 100) if total > 0 else 0
            
            return {
                "total": total,
                "processed": processed,
                "percentage": percentage,
                "running": self._running,
                "paused": self._paused
            }
            
    def get_next_result(self, timeout=0) -> Optional[Dict[str, Any]]:
        """
        获取下一个处理结果
        
        参数:
            timeout: 等待超时时间，0表示不等待
            
        返回:
            处理结果，如果没有则返回None
        """
        try:
            return self._results_queue.get(block=timeout > 0, timeout=timeout)
        except queue.Empty:
            return None
            
    def get_result_summary(self) -> Dict[str, Any]:
        """
        获取处理结果摘要
        
        返回:
            包含处理结果摘要的字典
        """
        with self._lock:
            return {
                "total": self._total_files,
                "success": self._success_count,
                "error": self._error_count,
                "error_files": self._error_files.copy()
            }