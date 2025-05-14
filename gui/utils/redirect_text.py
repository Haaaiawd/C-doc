"""
提供文本重定向功能，用于将程序输出重定向到GUI文本框
"""
from word_processors import extract_author_number, extract_author_from_filename

class RedirectText:
    """
    将标准输出重定向到tkinter文本控件
    """
    def __init__(self, text_widget, error_only=False):
        """
        初始化文本重定向器
        
        参数:
            text_widget: 目标文本控件
            error_only: 是否只显示错误信息
        """
        self.text_widget = text_widget
        self.error_only = error_only
        self.error_files = set()  # 存储错误文件路径
        self.success_files = set()  # 存储成功文件路径
        self.current_file = None  # 当前正在处理的文件

    def set_current_file(self, filename):
        """设置当前正在处理的文件名"""
        self.current_file = filename

    def write(self, string):
        """写入文本到目标控件"""
        if not string.strip():
            return
            
        # 检查是否是错误信息
        if string.strip().startswith('×'):
            # 如果有当前文件，添加到错误文件集合中
            if self.current_file:
                self.error_files.add(self.current_file)
                
        # 检查是否是成功信息（包括使用默认作者名的成功处理）
        elif string.strip().startswith('✓'):
            # 如果有当前文件，添加到成功文件集合中
            if self.current_file:
                self.success_files.add(self.current_file)
            
        # 显示错误信息，每个文件占一行
        if string.strip().startswith('×'):
            if self.current_file:
                if not string.strip().endswith('.docx'):
                    # 提取作者行数字和作者名
                    author_num = extract_author_number(self.current_file)
                    author_name = extract_author_from_filename(self.current_file)
                    error_msg = f"作者{author_num}({author_name}): {string.strip()}\n"
                else:
                    error_msg = string
            else:
                error_msg = string
            
            self.text_widget.insert('end', error_msg)
            self.text_widget.see('end')
            self.text_widget.update()
            
        # 显示警告信息，每个文件占一行
        elif string.strip().startswith('!'):
            if self.current_file:
                if not string.strip().endswith('.docx'):
                    # 提取作者行数字
                    author_num = extract_author_number(self.current_file)
                    warning_msg = f"作者{author_num}: {string.strip()}\n"
                else:
                    warning_msg = string
            else:
                warning_msg = string
            
            self.text_widget.insert('end', warning_msg)
            self.text_widget.see('end')
            self.text_widget.update()
        
        # 显示普通信息（非错误、非警告）
        elif not self.error_only:
            self.text_widget.insert('end', string)
            self.text_widget.see('end')
            self.text_widget.update()

    def flush(self):
        """刷新缓冲区，用于兼容sys.stdout"""
        pass

    def get_error_files(self):
        """获取处理出错的文件列表"""
        return self.error_files

    def get_success_files(self):
        """获取处理成功的文件列表"""
        return self.success_files

    def clear_files(self):
        """清空记录的文件列表"""
        self.error_files.clear()
        self.success_files.clear()
        self.current_file = None