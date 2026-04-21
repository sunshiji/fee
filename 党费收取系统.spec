# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['D:\\project\\work\\fee\\gui_main.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=['tkinter', 'tkinter.ttk', 'tkinter.messagebox', 'tkinter.filedialog', 'tkinter.scrolledtext', 'openpyxl', 'openpyxl.workbook', 'openpyxl.styles', 'openpyxl.utils', 'openpyxl.cell', 'openpyxl.worksheet', 'openpyxl.reader', 'openpyxl.writer', 'pandas', 'pandas.core', 'pandas.core.frame', 'pandas.core.series', 'pandas.core.indexes', 'pandas.core.dtypes', 'pandas.io', 'pandas.io.excel', 'pandas._libs', 'dateutil', 'dateutil.parser', 'dateutil.relativedelta', 'dateutil.tz', 'docx', 'docx.api', 'docx.document', 'docx.table', 'docx.text', 'docx.oxml', 'docx.shared', 'sqlite3', 'json', 'glob', 'shutil', 'typing', 'datetime', 'config', 'data_models', 'member_manager', 'fee_calculator', 'excel_generator', 'statistics', 'database_manager', 'database_manager', 'excel_importer', 'word_importer', 'numpy', 'numpy.core', 'numpy.lib'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['matplotlib', 'scipy', 'torch', 'tensorflow', 'IPython', 'jupyter', 'jupyter_client', 'jupyter_core', 'notebook', 'jupyterlab', 'ipykernel', 'ipywidgets', 'pytest', 'pandas.tests', 'openpyxl.tests', 'PIL', 'lxml', 'Crypto', 'cryptography', 'sqlalchemy', 'tables', 'h5py', 'bokeh', 'dask', 'distributed', 'numba', 'llvmlite', 'sphinx', 'docutils', 'babel', 'jedi', 'parso', 'nbformat', 'jsonschema', 'urllib3', 'requests', 'charset_normalizer', 'certifi', 'idna', 'certifi', 'nacl', 'zmq', 'tornado', 'pywin32', 'win32com', 'pythoncom', 'pywintypes', 'PyQt5', 'PyQt5.QtCore', 'PyQt5.QtGui', 'PyQt5.QtWidgets', 'qtpy', 'fsspec', 'cloudpickle', 'panel', 'pyviz_comms', 'markdown', 'plotly', 'black', 'flake8', 'pip', 'setuptools', 'conda', 'anaconda'],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='党费收取系统',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
