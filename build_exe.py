# -*- coding: utf-8 -*-
import PyInstaller.__main__
import os
import sys

# 项目根目录
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# 需要包含的隐藏导入
hidden_imports = [
    'tkinter',
    'tkinter.ttk',
    'tkinter.messagebox',
    'tkinter.filedialog',
    'tkinter.scrolledtext',
    'openpyxl',
    'openpyxl.workbook',
    'openpyxl.styles',
    'openpyxl.utils',
    'pandas',
    'pandas.core.frame',
    'pandas.core.series',
    'dateutil',
    'dateutil.parser',
    'dateutil.relativedelta',
    'docx',
    'docx.api',
    'docx.document',
    'docx.table',
    'docx.text.paragraph',
    'sqlite3',
    'json',
    'glob',
    'shutil',
    'typing',
    'datetime',
    'config',
    'data_models',
    'member_manager',
    'fee_calculator',
    'excel_generator',
    'statistics',
    'database_manager',
    'excel_importer',
    'word_importer',
]

# PyInstaller参数
args = [
    # 主入口文件
    os.path.join(BASE_DIR, 'gui_main.py'),
    
    # 打包为单个可执行文件
    '--onefile',
    
    # 不显示控制台窗口（GUI应用）
    '--windowed',
    
    # 应用程序名称
    '--name', '党费收取系统',
    
    # 清理临时文件
    '--clean',
    
    # 输出目录
    '--distpath', os.path.join(BASE_DIR, 'dist'),
    '--workpath', os.path.join(BASE_DIR, 'build'),
    '--specpath', BASE_DIR,
    
    # 图标（如果有的话）
    # '--icon', os.path.join(BASE_DIR, 'app.ico'),
    
    # 隐藏导入
    *[f'--hidden-import={imp}' for imp in hidden_imports],
    
    # 收集子模块
    '--collect-submodules', 'openpyxl',
    '--collect-submodules', 'pandas',
    '--collect-submodules', 'dateutil',
    '--collect-submodules', 'docx',
    
    # 额外的包
    '--collect-all', 'tkinter',
]

print("=" * 60)
print("开始打包党费收取系统...")
print("=" * 60)
print(f"Python版本: {sys.version}")
print(f"项目目录: {BASE_DIR}")
print("=" * 60)

# 运行PyInstaller
PyInstaller.__main__.run(args)

print("\n" + "=" * 60)
print("打包完成！")
print("=" * 60)
print(f"输出目录: {os.path.join(BASE_DIR, 'dist')}")
print("\n重要提示：")
print("1. 请将 'data' 目录复制到与EXE文件相同的目录")
print("2. 程序运行时会在EXE所在目录创建 'output' 目录")
print("3. 确保目标机器安装了 Visual C++ Redistributable")
print("=" * 60)
