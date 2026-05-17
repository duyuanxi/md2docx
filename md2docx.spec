# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_submodules, collect_data_files

hiddenimports = [
    # Core deps
    "markdown", "bs4", "docx", "lxml",
    # MarkItDown + all converters
    "markitdown", "markitdown._markitdown", "markitdown._base_converter",
    "markitdown._stream_info", "markitdown._uri_utils", "markitdown._exceptions",
    "markitdown.converters", "markitdown.converters._docx_converter",
    "markitdown.converters._pdf_converter", "markitdown.converters._pptx_converter",
    "markitdown.converters._xlsx_converter", "markitdown.converters._html_converter",
    "markitdown.converters._epub_converter", "markitdown.converters._image_converter",
    "markitdown.converters._audio_converter", "markitdown.converters._csv_converter",
    "markitdown.converters._ipynb_converter", "markitdown.converters._plain_text_converter",
    "markitdown.converters._zip_converter", "markitdown.converters._outlook_msg_converter",
    "markitdown.converters._rss_converter", "markitdown.converters._wikipedia_converter",
    "markitdown.converters._youtube_converter", "markitdown.converters._bing_serp_converter",
    "markitdown.converters._markdownify",
    # MarkItDown optional deps
    "mammoth", "pdfminer", "pdfminer.high_level", "pdfminer.layout",
    "pdfminer.pdfparser", "pdfminer.pdfdocument", "pdfminer.pdfpage",
    "pdfminer.pdfinterp", "pdfminer.converter", "pdfminer.cmapdb",
    "openpyxl", "xlrd", "pptx", "python.pptx", "python.pptx.parts",
    "speechrecognition", "pydub",
    # MarkItDown utils
    "magika", "charset_normalizer",
    "markdownify", "puremagic",
    # Image handling
    "PIL", "PIL.Image",
    # Jupyter
    "nbformat", "nbconvert",
    # Network
    "requests", "urllib3",
    # Misc
    "json", "csv", "io", "re", "pathlib",
    "xml", "xml.etree", "xml.etree.ElementTree",
    "html", "html.parser",
]

# Collect submodules for packages PyInstaller might miss
for pkg in ["markitdown", "docx", "bs4", "lxml", "PIL", "openpyxl", "pdfminer"]:
    try:
        hiddenimports.extend(collect_submodules(pkg))
    except Exception:
        pass

a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'tkinter.test', 'unittest', 'email', 'http.server',
        'pydoc', 'distutils', 'setuptools', 'pip',
    ],
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
    name='md2docx',
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
    icon=None,
)
