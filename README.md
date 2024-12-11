<<<<<<< HEAD
# Word文档批量处理工具 v1.1.0

## 更新说明
- 新增提取图片功能
- 新增提取标题功能
- 优化文件处理逻辑
- 改进错误处理和显示
- 分离有图片和无图片的成功文件
- 正文段落首行缩进两个字符

=======
# Word文档处理脚本
> 为福州大学关工委设计
>>>>>>> 569b49ee032f5b58ec96de57af1f499111c175d9
## 功能说明
这个脚本用于处理Word文档，主要功能包括：
1. 将第一行设置为标题（黑体三号加粗）
2. 在标题后添加指定的副标题文本（黑体三号加粗）
3. 处理作者信息（852开头的行，宋体四号）
4. 正文内容设置为宋体小四号，段落首行缩进两个字符
5. 检测并收集文档中的所有图片，将其移动到文档末尾
6. 批量处理文件夹中的所有Word文档
7. 自动重命名处理后的文件，格式为：《作者名》原标题——福州大学先进制造学院与海洋学院关工委2023年"中华魂"（毛泽东伟大精神品格）主题教育征文
8. 图形界面支持：
   - 选择输入输出文件夹
   - 显示处理错误信息
   - 打包错误文件
   - 实时显示处理进度
   - 提取所有文档中的图片
   - 提取并显示所有文档的标题

## 使用方法
### 方式一：直接运行可执行文件
1. 下载并运行 "Word文档批量处理工具.exe"
2. 在界面上选择输入文件夹（包含要处理的Word文档）
3. 选择输出文件夹（处理后的文件保存位置）
4. 可以使用以下功能：
   - 点击"开始转换"进行文档处理
   - 点击"提取图片"提取所有文档中的图片
   - 点击"提取标题"查看所有文档的标题
   - 点击"打包错误文件"收集处理失败的文件

### 方式二：运行Python脚本
1. 安装Python 3.6+
2. 安装依赖库：
   - python-docx 库
   - docx2python 库
   - tkinter 库（Python标准库）

## 输出说明
1. 成功处理的文件会保存在"成功文件"文件夹中
2. 无图片的成功文件会保存在"无图片成功文件"文件夹中
3. 提取的图片会保存在"提取的图片"文件夹中，每个文档的图片单独存放
4. 处理失败的文件会被收集到"错误文件"文件夹中

<<<<<<< HEAD
## 注意事项
1. 确保Word文档格式正确，避免损坏的文件
2. 建议在处理前备份原始文件
3. 如果提取图片时遇到问题，可能是文档中的图片格式不支持
4. 标题提取功能会获取文档的第一个非空段落作为标题
=======
## 主要改进：
重新组织了内容结构，更清晰易读
添加了简介部分
将功能分类展示
详细说明了使用方法
补充了更多常见问题
优化了开发者信息的展示
添加了环境要求说明
>>>>>>> 569b49ee032f5b58ec96de57af1f499111c175d9
