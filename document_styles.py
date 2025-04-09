from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml.ns import qn
from docx.enum.style import WD_STYLE_TYPE

def create_document_with_styles():
    """创建一个带有预定义样式的文档"""
    doc = Document()
    
    # 创建标题样式
    title_style = doc.styles.add_style('Custom Title', WD_STYLE_TYPE.PARAGRAPH)
    font = title_style.font
    font.name = '黑体'
    font.size = Pt(16)
    font.bold = True
    font._element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')
    title_style.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    
    # 创建副标题样式
    subtitle_style = doc.styles.add_style('Custom Subtitle', WD_STYLE_TYPE.PARAGRAPH)
    font = subtitle_style.font
    font.name = '黑体'
    font.size = Pt(16)
    font.bold = True
    font._element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')
    subtitle_style.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    
    # 创建作者样式
    author_style = doc.styles.add_style('Custom Author', WD_STYLE_TYPE.PARAGRAPH)
    font = author_style.font
    font.name = '宋体'
    font.size = Pt(14)
    font.bold = True
    font._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
    author_style.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    
    # 创建正文样式
    body_style = doc.styles.add_style('Custom Body', WD_STYLE_TYPE.PARAGRAPH)
    font = body_style.font
    font.name = '宋体'
    font.size = Pt(12)
    font._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
    body_style.paragraph_format.first_line_indent = Pt(24)  # 首行缩进2字符
    
    return doc

def apply_title_style(doc, paragraph):
    """应用标题样式"""
    if 'Custom Title' in doc.styles:
        paragraph.style = doc.styles['Custom Title']
    else:
        # 手动设置格式
        for run in paragraph.runs:
            run.font.name = '黑体'
            run.font.size = Pt(16)
            run.font.bold = True
            run._element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')
        paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

def apply_subtitle_style(doc, paragraph):
    """应用副标题样式"""
    if 'Custom Subtitle' in doc.styles:
        paragraph.style = doc.styles['Custom Subtitle']
    else:
        # 手动设置格式
        for run in paragraph.runs:
            run.font.name = '黑体'
            run.font.size = Pt(16)
            run.font.bold = True
            run._element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')
        paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

def apply_author_style(doc, paragraph):
    """应用作者样式"""
    if 'Custom Author' in doc.styles:
        paragraph.style = doc.styles['Custom Author']
    else:
        # 手动设置格式
        for run in paragraph.runs:
            run.font.name = '宋体'
            run.font.size = Pt(14)
            run.font.bold = True
            run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
        paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

def apply_body_style(doc, paragraph):
    """应用正文样式"""
    if 'Custom Body' in doc.styles:
        paragraph.style = doc.styles['Custom Body']
    else:
        # 手动设置格式
        for run in paragraph.runs:
            run.font.name = '宋体'
            run.font.size = Pt(12)
            run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
        paragraph.paragraph_format.first_line_indent = Pt(24)  # 首行缩进2字符

# 新增样式函数 - 创建中文正式文档格式
def create_chinese_formal_document():
    """创建一个带有中文正式格式的文档"""
    doc = Document()
    
    # 创建主标题样式 - 小二号方正小标宋简体
    main_title_style = doc.styles.add_style('Chinese Main Title', WD_STYLE_TYPE.PARAGRAPH)
    font = main_title_style.font
    font.name = '方正小标宋简体'
    font.size = Pt(18)  # 小二号字体约为18磅
    font._element.rPr.rFonts.set(qn('w:eastAsia'), '方正小标宋简体')
    main_title_style.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    
    # 创建副标题样式 - 小三号仿宋_GB2312
    subtitle_style = doc.styles.add_style('Chinese Subtitle', WD_STYLE_TYPE.PARAGRAPH)
    font = subtitle_style.font
    font.name = '仿宋_GB2312'
    font.size = Pt(15)  # 小三号字体约为15磅
    font._element.rPr.rFonts.set(qn('w:eastAsia'), '仿宋_GB2312')
    subtitle_style.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    
    # 创建正文样式 - 小三号仿宋_GB2312
    body_style = doc.styles.add_style('Chinese Body', WD_STYLE_TYPE.PARAGRAPH)
    font = body_style.font
    font.name = '仿宋_GB2312'
    font.size = Pt(15)  # 小三号字体约为15磅
    font._element.rPr.rFonts.set(qn('w:eastAsia'), '仿宋_GB2312')
    body_style.paragraph_format.first_line_indent = Pt(30)  # 首行缩进2字符
    
    # 移除不需要的标题样式
    
    return doc

# 应用中文主标题样式
def apply_chinese_main_title(doc, paragraph):
    """应用中文主标题样式"""
    if 'Chinese Main Title' in doc.styles:
        paragraph.style = doc.styles['Chinese Main Title']
    else:
        # 手动设置格式
        for run in paragraph.runs:
            run.font.name = '方正小标宋简体'
            run.font.size = Pt(18)
            run._element.rPr.rFonts.set(qn('w:eastAsia'), '方正小标宋简体')
        paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

# 应用中文副标题样式
def apply_chinese_subtitle(doc, paragraph):
    """应用中文副标题样式"""
    if 'Chinese Subtitle' in doc.styles:
        paragraph.style = doc.styles['Chinese Subtitle']
    else:
        # 手动设置格式
        for run in paragraph.runs:
            run.font.name = '仿宋_GB2312'
            run.font.size = Pt(15)
            run._element.rPr.rFonts.set(qn('w:eastAsia'), '仿宋_GB2312')
        paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

# 应用中文正文样式
def apply_chinese_body(doc, paragraph):
    """应用中文正文样式"""
    if 'Chinese Body' in doc.styles:
        paragraph.style = doc.styles['Chinese Body']
    else:
        # 手动设置格式
        for run in paragraph.runs:
            run.font.name = '仿宋_GB2312'
            run.font.size = Pt(15)
            run._element.rPr.rFonts.set(qn('w:eastAsia'), '仿宋_GB2312')
        paragraph.paragraph_format.first_line_indent = Pt(30)  # 首行缩进2字符

# 应用中文一级标题样式
def apply_chinese_heading1(doc, paragraph):
    """应用中文一级标题样式"""
    if 'Chinese Heading 1' in doc.styles:
        paragraph.style = doc.styles['Chinese Heading 1']
    else:
        # 手动设置格式
        for run in paragraph.runs:
            run.font.name = '黑体'
            run.font.size = Pt(15)
            run._element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')

# 应用中文二级标题样式
def apply_chinese_heading2(doc, paragraph):
    """应用中文二级标题样式"""
    if 'Chinese Heading 2' in doc.styles:
        paragraph.style = doc.styles['Chinese Heading 2']
    else:
        # 手动设置格式
        for run in paragraph.runs:
            run.font.name = '楷体_GB2312'
            run.font.size = Pt(15)
            run._element.rPr.rFonts.set(qn('w:eastAsia'), '楷体_GB2312')

# 应用中文三四级标题样式
def apply_chinese_heading34(doc, paragraph):
    """应用中文三四级标题样式"""
    if 'Chinese Heading 3-4' in doc.styles:
        paragraph.style = doc.styles['Chinese Heading 3-4']
    else:
        # 手动设置格式
        for run in paragraph.runs:
            run.font.name = '仿宋_GB2312'
            run.font.size = Pt(15)
            run._element.rPr.rFonts.set(qn('w:eastAsia'), '仿宋_GB2312')

# 改进的标题识别系统
def identify_heading_level(text, doc=None, paragraph=None):
    """
    更智能地识别标题层级
    
    参数:
        text: 段落文本
        doc: 文档对象(可选)
        paragraph: 段落对象(可选)
    
    返回:
        识别到的标题级别(1-4)，如果不是标题则返回0
    """
    text = text.strip()
    
    # 如果文本为空，则不是标题
    if not text:
        return 0
        
    # 识别特征
    features = {
        "length": len(text),
        "has_number_prefix": False,
        "prefix_type": None,
        "ends_with_punct": text[-1] in "。，；：！？,.;:!?",
    }
    
    # 检查一级标题特征（中文数字 + 顿号）
    level1_prefixes = ['一、', '二、', '三、', '四、', '五、', '六、', '七、', '八、', '九、', '十、',
                       '十一、', '十二、', '十三、', '十四、', '十五、', '十六、', '十七、', '十八、', '十九、', '二十、']
    
    # 检查二级标题特征（括号中文数字）
    level2_prefixes = []
    for num in ['一', '二', '三', '四', '五', '六', '七', '八', '九', '十',
               '十一', '十二', '十三', '十四', '十五', '十六', '十七', '十八', '十九', '二十']:
        level2_prefixes.append(f'({num})')
        level2_prefixes.append(f'（{num}）')
    
    # 检查三级标题特征（阿拉伯数字 + 点）
    level3_prefixes = [f'{i}.' for i in range(1, 31)] + [f'{i}、' for i in range(1, 31)]
    
    # 检查四级标题特征（括号阿拉伯数字）
    level4_prefixes = [f'({i})' for i in range(1, 31)] + [f'（{i}）' for i in range(1, 31)]
    
    # 非编号标题关键词
    common_headings = ['摘要', '引言', '前言', '背景', '介绍', '结论', '总结', '参考文献', 
                       '致谢', '附录', '问题', '方法', '研究方法', '实验', '实验结果', 
                       '讨论', '建议', '展望']
    
    # 判断前缀类型
    if any(text.startswith(prefix) for prefix in level1_prefixes):
        features["has_number_prefix"] = True
        features["prefix_type"] = "level1"
    elif any(text.startswith(prefix) for prefix in level2_prefixes):
        features["has_number_prefix"] = True
        features["prefix_type"] = "level2"
    elif any(text.startswith(prefix) for prefix in level3_prefixes):
        features["has_number_prefix"] = True
        features["prefix_type"] = "level3"
    elif any(text.startswith(prefix) for prefix in level4_prefixes):
        features["has_number_prefix"] = True
        features["prefix_type"] = "level4"
    
    # 启发式规则识别标题
    if features["prefix_type"] == "level1" and features["length"] < 50:
        return 1
    elif features["prefix_type"] == "level2" and features["length"] < 50:
        return 2
    elif features["prefix_type"] == "level3" and features["length"] < 50:
        return 3
    elif features["prefix_type"] == "level4" and features["length"] < 50:
        return 4
    
    # 识别无编号常见标题
    if features["length"] < 20 and not features["ends_with_punct"]:
        # 检查是否为常见标题词
        for heading in common_headings:
            if heading in text:
                return 1  # 默认作为一级标题处理
    
    # 如果提供了段落对象，尝试从格式中识别
    if paragraph is not None:
        try:
            # 如果段落已有标题样式，提取级别
            style_name = paragraph.style.name
            if 'heading' in style_name.lower() or '标题' in style_name:
                for i in range(1, 5):
                    if str(i) in style_name:
                        return i
        except:
            pass
    
    # 默认不是标题
    return 0

# 判断段落是否为一级标题
def is_heading1(text, doc=None, paragraph=None):
    return identify_heading_level(text, doc, paragraph) == 1

# 判断段落是否为二级标题
def is_heading2(text, doc=None, paragraph=None):
    return identify_heading_level(text, doc, paragraph) == 2

# 判断段落是否为三级标题
def is_heading3(text, doc=None, paragraph=None):
    return identify_heading_level(text, doc, paragraph) == 3

# 判断段落是否为四级标题
def is_heading4(text, doc=None, paragraph=None):
    return identify_heading_level(text, doc, paragraph) == 4

# 判断段落是否为三级或四级标题
def is_heading3_or_4(text, doc=None, paragraph=None):
    level = identify_heading_level(text, doc, paragraph)
    return level == 3 or 4
