"""Multi-format file importer powered by Microsoft MarkItDown.

Full API for the GUI: converter list, format info, conversion with metadata.
"""

import re
import unicodedata
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

from markitdown import MarkItDown
from markitdown._markitdown import PRIORITY_SPECIFIC_FILE_FORMAT


# ── Converter registry ───────────────────────────────────────────

# Map file extensions to human-readable descriptions
FORMAT_MAP: dict[str, tuple[str, str]] = {
    ".docx":  ("Word 文档", "Microsoft Word (.docx)"),
    ".pdf":   ("PDF", "Adobe PDF (.pdf)"),
    ".pptx":  ("PowerPoint", "演示文稿 (.pptx)"),
    ".xlsx":  ("Excel", "电子表格 (.xlsx)"),
    ".xls":   ("Excel 97", "旧版电子表格 (.xls)"),
    ".csv":   ("CSV", "逗号分隔值 (.csv)"),
    ".html":  ("HTML", "网页 (.html/.htm)"),
    ".htm":   ("HTML", "网页 (.html/.htm)"),
    ".epub":  ("EPUB", "电子书 (.epub)"),
    ".md":    ("Markdown", "Markdown 文本 (.md)"),
    ".markdown": ("Markdown", "Markdown 文本 (.markdown)"),
    ".txt":   ("纯文本", "纯文本 (.txt)"),
    ".jpg":   ("JPEG 图片", "JPEG 图片 (.jpg)"),
    ".jpeg":  ("JPEG 图片", "JPEG 图片 (.jpeg)"),
    ".png":   ("PNG 图片", "PNG 图片 (.png)"),
    ".gif":   ("GIF 图片", "GIF 图片 (.gif)"),
    ".bmp":   ("BMP 图片", "位图 (.bmp)"),
    ".webp":  ("WebP 图片", "WebP 图片 (.webp)"),
    ".mp3":   ("MP3 音频", "MP3 音频 (.mp3)"),
    ".wav":   ("WAV 音频", "WAV 音频 (.wav)"),
    ".m4a":   ("M4A 音频", "M4A 音频 (.m4a)"),
    ".ogg":   ("OGG 音频", "OGG 音频 (.ogg)"),
    ".ipynb": ("Jupyter", "Jupyter Notebook (.ipynb)"),
    ".msg":   ("Outlook 邮件", "Outlook 邮件 (.msg)"),
    ".xml":   ("XML", "XML 文档 (.xml)"),
    ".json":  ("JSON", "JSON 文档 (.json)"),
    ".zip":   ("ZIP 压缩包", "ZIP 压缩包 (.zip)"),
}

SUPPORTED_EXTENSIONS = set(FORMAT_MAP.keys())

FILE_FILTERS = [
    ("所有支持的文件",
     "*.md;*.docx;*.pdf;*.pptx;*.xlsx;*.html;*.htm;*.epub;*.txt;*.csv;*.jpg;*.png;*.ipynb;*.xml;*.json;*.zip"),
    ("文档", "*.docx;*.pdf;*.md;*.txt;*.epub;*.csv"),
    ("表格/演示", "*.xlsx;*.xls;*.pptx"),
    ("网页", "*.html;*.htm;*.xml;*.json"),
    ("图片/音频", "*.jpg;*.jpeg;*.png;*.gif;*.bmp;*.webp;*.mp3;*.wav;*.m4a;*.ogg"),
    ("所有文件", "*.*"),
]


@dataclass
class ImportResult:
    """Result of importing a file."""
    text: str                    # Clean markdown text
    source_path: str             # Original file path
    source_format: str           # Detected format name
    source_ext: str              # File extension
    file_size_kb: float          # File size in KB
    converter_used: str          # Name of converter that handled it
    success: bool = True
    error: str = ""


# ── Engine ───────────────────────────────────────────────────────

class ImportEngine:
    """Wraps MarkItDown with format detection, batch support, and metadata."""

    def __init__(self):
        self._md = MarkItDown()

    def convert(self, path: str) -> ImportResult:
        """Convert a file to Markdown, returning full metadata."""
        p = Path(path)
        ext = p.suffix.lower()
        fmt_name, fmt_desc = FORMAT_MAP.get(ext, ("未知", f"未知格式 ({ext})"))
        size_kb = p.stat().st_size / 1024 if p.exists() else 0

        try:
            if ext in {".md", ".markdown", ".txt"}:
                text = _read_text_file(path)
                converter = "PlainText"
            else:
                result = self._md.convert(path)
                text = _clean_text(result.text_content)
                converter = "MarkItDown"

            return ImportResult(
                text=text,
                source_path=path,
                source_format=fmt_name,
                source_ext=ext,
                file_size_kb=size_kb,
                converter_used=converter,
            )
        except Exception as e:
            # Try fallback for text files
            if ext in {".md", ".markdown", ".txt", ".csv", ".xml", ".json"}:
                try:
                    text = _read_text_file(path)
                    return ImportResult(
                        text=text, source_path=path,
                        source_format=fmt_name, source_ext=ext,
                        file_size_kb=size_kb, converter_used="PlainText(fallback)",
                    )
                except Exception:
                    pass
            err_msg = str(e)
            # Translate common errors
            if "password" in err_msg.lower() or "encrypt" in err_msg.lower():
                err_msg = "文件已加密或受密码保护，无法读取。请先解密文件后再导入。"
            elif "truncated" in err_msg.lower() or "corrupt" in err_msg.lower():
                err_msg = "文件已损坏或不完整，无法解析。请确认文件来源可靠。"
            elif "not supported" in err_msg.lower() or "unsupported" in err_msg.lower():
                err_msg = f"不支持此文件格式或版本。当前支持 27 种格式（PDF/DOCX/PPTX/XLSX/EPUB/HTML/图片/音频等）。"
            elif "module" in err_msg.lower() or "import" in err_msg.lower():
                err_msg = f"缺少必要的组件来解析此文件。\n请运行: pip install markitdown[all]\n\n原始错误: {err_msg}"
            elif "missingdependency" in err_msg.lower() or "dependencies" in err_msg.lower():
                # MarkItDown MissingDependencyException
                err_msg = ("缺少解析此文件格式所需的依赖包。\n\n"
                           "请运行以下命令安装全部可选依赖:\n"
                           "  pip install \"markitdown[all]\"\n\n"
                           "或仅安装所需组件:\n"
                           "  pip install \"markitdown[docx]\"  (Word 文档)\n"
                           "  pip install \"markitdown[pptx]\"  (PowerPoint)\n"
                           "  pip install \"markitdown[xlsx]\"  (Excel)\n"
                           "  pip install \"markitdown[pdf]\"   (PDF)\n\n"
                           f"原始错误: {err_msg}")
            return ImportResult(
                text="", source_path=path, source_format=fmt_name,
                source_ext=ext, file_size_kb=size_kb,
                converter_used="", success=False, error=err_msg,
            )

    @staticmethod
    def get_format_list() -> list[tuple[str, str, str]]:
        """Return [(ext, short_name, description), ...] for supported formats."""
        return [(ext, name, desc) for ext, (name, desc) in sorted(FORMAT_MAP.items())]


# ── Internal helpers ─────────────────────────────────────────────

def _read_text_file(path: str) -> str:
    """Read text file with encoding detection."""
    for enc in ["utf-8", "utf-8-sig", "gbk", "gb2312", "gb18030", "latin-1"]:
        try:
            with open(path, "r", encoding=enc) as f:
                return _clean_text(f.read())
        except (UnicodeDecodeError, UnicodeError):
            continue
    with open(path, "r", encoding="utf-8", errors="replace") as f:
        return _clean_text(f.read())


def _clean_text(text: str) -> str:
    """Normalize: strip control chars, normalize Unicode, fix line endings."""
    text = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f]", "", text)
    text = unicodedata.normalize("NFKC", text)
    text = text.replace("�", "").replace("﻿", "").replace("​", "")
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = "\n".join(line.rstrip() for line in text.split("\n"))
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def is_supported(path: str) -> bool:
    return Path(path).suffix.lower() in SUPPORTED_EXTENSIONS


def import_any(path: str) -> str:
    """Quick one-shot import → clean markdown string."""
    return ImportEngine().convert(path).text
