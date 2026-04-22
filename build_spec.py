# -*- mode: python ; coding: utf-8 -*-
import PyInstaller.__main__
import os

# 项目根目录
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# 精简的隐藏导入 - 只包含必要的模块
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
    'openpyxl.cell',
    'openpyxl.worksheet',
    'openpyxl.reader',
    'openpyxl.writer',
    'pandas',
    'pandas.core',
    'pandas.core.frame',
    'pandas.core.series',
    'pandas.core.indexes',
    'pandas.core.dtypes',
    'pandas.io',
    'pandas.io.excel',
    'pandas._libs',
    'dateutil',
    'dateutil.parser',
    'dateutil.relativedelta',
    'dateutil.tz',
    'docx',
    'docx.api',
    'docx.document',
    'docx.table',
    'docx.text',
    'docx.oxml',
    'docx.shared',
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
    'database_manager',
    'excel_importer',
    'word_importer',
    'numpy',
    'numpy.core',
    'numpy.lib',
]

# 排除不需要的模块，减小体积
excludes = [
    'matplotlib',
    'scipy',
    'torch',
    'tensorflow',
    'IPython',
    'jupyter',
    'jupyter_client',
    'jupyter_core',
    'notebook',
    'jupyterlab',
    'ipykernel',
    'ipywidgets',
    'pytest',
    'pandas.tests',
    'openpyxl.tests',
    'PIL',
    'lxml',
    'Crypto',
    'cryptography',
    'sqlalchemy',
    'tables',
    'h5py',
    'bokeh',
    'dask',
    'distributed',
    'numba',
    'llvmlite',
    'sphinx',
    'docutils',
    'babel',
    'jedi',
    'parso',
    'nbformat',
    'jsonschema',
    'urllib3',
    'requests',
    'charset_normalizer',
    'certifi',
    'idna',
    'certifi',
    'nacl',
    'zmq',
    'tornado',
    'pywin32',
    'win32com',
    'pythoncom',
    'pywintypes',
    'PyQt5',
    'PyQt5.QtCore',
    'PyQt5.QtGui',
    'PyQt5.QtWidgets',
    'qtpy',
    'fsspec',
    'cloudpickle',
    'panel',
    'pyviz_comms',
    'markdown',
    'plotly',
    'black',
    'flake8',
    'pip',
    'setuptools',
    'conda',
    'anaconda',
]

args = [
    os.path.join(BASE_DIR, 'gui_main.py'),
    '--onefile',
    '--windowed',
    '--name', '党费收取系统',
    '--clean',
    '--distpath', os.path.join(BASE_DIR, 'dist'),
    '--workpath', os.path.join(BASE_DIR, 'build'),
    '--specpath', BASE_DIR,
]

# 添加隐藏导入
for imp in hidden_imports:
    args.extend(['--hidden-import', imp])

# 添加排除模块
for exc in excludes:
    args.extend(['--exclude-module', exc])

print("=" * 60)
print("开始精简打包党费收取系统...")
print("=" * 60)
print(f"项目目录: {BASE_DIR}")
print(f"隐藏导入数量: {len(hidden_imports)}")
print(f"排除模块数量: {len(excludes)}")
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
print("=" * 60)
