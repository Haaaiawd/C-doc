"""
提供图片提取相关的业务逻辑处理
"""
import os
import threading
from image_extractor import extract_images_from_doc

class ImageHandler:
    """图片提取处理器，处理从文档中提取图片的相关业务逻辑"""
    
    def __init__(self, app):
        """
        初始化图片提取处理器
        
        参数:
            app: 主应用程序实例
        """
        self.app = app
        
    def validate_paths(self, input_dir):
        """
        验证输入路径是否有效
        
        参数:
            input_dir: 输入目录
            
        返回:
            tuple: (是否有效, 错误信息)
        """
        if not input_dir:
            return False, "请选择输入文件夹"
            
        # 检查目录是否存在
        if not os.path.exists(input_dir):
            return False, "输入目录不存在"
            
        return True, ""
        
    def extract_images(self):
        """开始从文档中提取图片"""
        # 获取路径配置
        paths = self.app.file_frame.get_paths()
        input_dir = paths["input_dir"]
        output_dir = paths["output_dir"]
        
        # 验证输入路径
        valid, error_msg = self.validate_paths(input_dir)
        if not valid:
            self.app.set_status(error_msg)
            return
            
        if not output_dir:
            self.app.set_status("请选择输出文件夹")
            return
            
        # 准备开始提取图片
        self.app.reset_for_processing()
        self.app.set_status("正在提取图片...")
        
        # 在新线程中运行提取
        thread = threading.Thread(
            target=self._extract_thread, 
            args=(input_dir, output_dir),
            daemon=True
        )
        thread.start()
    
    def _extract_thread(self, input_dir, output_dir):
        """提取图片线程"""
        try:
            # 创建图片输出目录
            images_dir = os.path.join(output_dir, "提取的图片")
            os.makedirs(images_dir, exist_ok=True)
            
            # 获取所有Word文件
            docx_files = [f for f in os.listdir(input_dir) if f.endswith(('.doc', '.docx'))]
            total_files = len(docx_files)
            
            # 配置进度条最大值
            self.app.root.after(0, lambda: self.app.progress_bar.configure(maximum=total_files))
            
            total_images = 0
            
            for index, filename in enumerate(docx_files, 1):
                # 更新进度条
                current_progress = index / total_files * 100
                self.app.root.after(0, lambda p=index: self.app.progress_bar.configure(value=p))
                self.app.root.after(0, lambda: self.app.progress_bar.update())
                
                self.app.set_status(f"正在提取图片: {index}/{total_files} ({int(current_progress)}%)...")
                
                input_file = os.path.join(input_dir, filename)
                # 为每个文件创建子文件夹
                file_images_dir = os.path.join(images_dir, os.path.splitext(filename)[0])
                os.makedirs(file_images_dir, exist_ok=True)
                
                # 提取图片
                temp_images = extract_images_from_doc(input_file, file_images_dir)
                if temp_images:
                    self.app.log_text.insert('end', f"✓ {filename}: 提取了 {len(temp_images)} 张图片\n")
                    total_images += len(temp_images)
                else:
                    self.app.log_text.insert('end', f"! {filename}: 未找到图片\n")
            
            summary = f"\n提取完成！\n总计处理 {len(docx_files)} 个文件\n共提取 {total_images} 张图片\n"
            self.app.log_text.insert('end', summary)
            self.app.log_text.see('end')
            
            # 设置进度条为100%完成
            self.app.root.after(0, lambda: self.app.progress_bar.configure(value=total_files))
            
            # 更新状态
            self.app.root.after(0, lambda: self.app.set_status(f"图片提取完成，共 {total_images} 张"))
            self.app.root.after(0, self.app.enable_buttons)
            
        except Exception as e:
            self.app.root.after(0, lambda: self.app.set_status(f"提取图片时出错: {str(e)}"))
            self.app.root.after(0, self.app.enable_buttons)