"""
提供简单的标题识别功能和字数统计功能
"""

def is_heading1(text):
    """判断文本是否为一级标题"""
    # 检查常见标题前缀，避免复杂的识别逻辑
    heading_prefixes = ['一、', '二、', '三、', '四、', '五、', '六、', '七、', '八、', '九、', '十、']
    common_headings = ['摘要', '引言', '前言', '背景', '介绍', '结论', '总结', '参考文献']
    
    # 如果行很短且是关键词，可能是标题
    if any(text.startswith(prefix) for prefix in heading_prefixes) and len(text) < 40:
        return True
    
    # 如果是常见标题关键词
    if any(heading in text for heading in common_headings) and len(text) < 20:
        return True
    
    return False

def count_document_words(doc):
    """
    计算Word文档的总字数(包括中英文字符)
    
    参数:
        doc: 已加载的docx文档对象
        
    返回:
        int: 文档字数
    """
    total_words = 0
    
    # 遍历所有段落
    for para in doc.paragraphs:
        text = para.text.strip()
        if text:
            # 统计中英文字符(不包括空格和标点)
            for char in text:
                # 判断是否是字母、数字或中文字符
                if char.isalnum() or '\u4e00' <= char <= '\u9fff':
                    total_words += 1
    
    # 遍历所有表格中的文本
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    text = para.text.strip()
                    if text:
                        # 统计中英文字符(不包括空格和标点)
                        for char in text:
                            # 判断是否是字母、数字或中文字符
                            if char.isalnum() or '\u4e00' <= char <= '\u9fff':
                                total_words += 1
    
    return total_words

def get_document_stats(doc):
    """
    获取文档统计信息
    
    参数:
        doc: 已加载的docx文档对象
        
    返回:
        dict: 包含文档统计信息的字典
    """
    stats = {
        'word_count': count_document_words(doc),
        'paragraph_count': len(doc.paragraphs),
        'character_count': sum(len(para.text) for para in doc.paragraphs),
    }
    return stats
