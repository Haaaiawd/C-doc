import PyInstaller.__main__

PyInstaller.__main__.run([
    'gui.py',
    '--name=Word文档批量处理工具',
    '--onefile',
    '--windowed',
    '--clean',
    '--add-data=README.md;.',
    # Exclude large modules that aren't needed
    '--exclude-module=scipy',
    '--exclude-module=pandas',
    '--exclude-module=matplotlib',
    '--exclude-module=tkinter.test',
    '--exclude-module=unittest',
    # 添加 pywin32 相关模块，以支持 doc 文件处理
    '--hidden-import=win32com',
    '--hidden-import=win32com.client',
    '--hidden-import=win32com.client.gencache',
    '--hidden-import=pythoncom',
    '--hidden-import=pywintypes',
    # Only collect specific modules that are needed
    '--collect-submodules=docx',
    '--collect-submodules=docx2python',
    '--collect-submodules=win32com',
    '--collect-submodules=pythoncom',
])
