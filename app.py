"""Markdown to Word Converter — Two-tab GUI with full MarkItDown integration.

Tab 1 "导入": Convert any file (PDF/DOCX/PPTX/XLSX/EPUB/HTML/IMG...) → Markdown
Tab 2 "编辑": Edit Markdown → live Word preview → export .docx
"""

import os
import re
import threading
import tkinter as tk
from tkinter import colorchooser, filedialog, font as tkfont
from tkinter import messagebox, ttk

import markdown
from bs4 import BeautifulSoup, NavigableString, Tag

from converter import convert_text, _clean_ai_blank_lines
from importer import (FILE_FILTERS, FORMAT_MAP, SUPPORTED_EXTENSIONS,
                      ImportEngine, ImportResult)
from templates import (COLOR_PRESETS, PAPER_SIZES, SIZE_LIST,
                       TemplateManager, BUILTIN_TEMPLATES)

DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_TEMPLATE = BUILTIN_TEMPLATES[0]
SIZE_STRS = [str(s) for s in SIZE_LIST]

# ── Palette ─────────────────────────────────────────────────────
C_HEADER      = "#4a9de0"
C_HEADER_TEXT = "#ffffff"
C_BG          = "#f0f5f3"
C_CARD        = "#ffffff"
C_CARD_WARM   = "#faf7ed"
C_ACCENT      = "#3b82c4"
C_TEXT        = "#1e293b"
C_MUTED       = "#64748b"
C_BORDER      = "#e2e8f0"
C_INPUT_BG    = "#f8fafc"
C_EDITOR_BG   = "#fafbfc"

PAPER_NAMES = list(PAPER_SIZES.keys())


def _translate_error(err: str) -> str:
    """Translate common Python/file errors to Chinese."""
    err_lower = err.lower()
    if "permission" in err_lower or "denied" in err_lower:
        return f"权限不足，无法写入文件。\n请检查文件是否被其他程序占用，或选择其他保存位置。\n\n详细信息: {err}"
    if "no such file" in err_lower or "not found" in err_lower or "does not exist" in err_lower:
        return f"找不到指定路径。\n请确认文件夹存在且可访问。\n\n详细信息: {err}"
    if "disk" in err_lower or "space" in err_lower or "full" in err_lower:
        return f"磁盘空间不足，无法写入文件。\n请清理磁盘后重试。\n\n详细信息: {err}"
    if "path" in err_lower and ("too long" in err_lower or "long" in err_lower):
        return f"文件路径过长。\n请选择较短的保存路径。\n\n详细信息: {err}"
    if "invalid" in err_lower or "illegal" in err_lower:
        return f"文件名包含无效字符，请使用合法文件名。\n\n详细信息: {err}"
    if "read-only" in err_lower or "readonly" in err_lower:
        return f"目标位置为只读，无法写入。\n请选择其他可写位置。\n\n详细信息: {err}"
    return f"转换过程中发生错误。\n\n详细信息: {err}"


def _get_system_fonts() -> list[str]:
    try:
        raw = sorted(set(tkfont.families()))
    except Exception:
        raw = []
    priority = [
        "微软雅黑", "宋体", "黑体", "仿宋", "楷体", "方正小标宋简体",
        "Arial", "Times New Roman", "Calibri", "Courier New",
        "Consolas", "Segoe UI", "Verdana", "Georgia",
    ]
    seen, result = set(), []
    for p in priority:
        if p in raw and p not in seen:
            result.append(p); seen.add(p)
    for f in raw:
        if f not in seen:
            result.append(f); seen.add(f)
    return result


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Markdown → Word 转换器 (Powered by MarkItDown)")
        self.root.geometry("1200x850")
        self.root.minsize(800, 560)
        self.root.configure(bg=C_BG)

        self._fonts = _get_system_fonts()
        self.tm = TemplateManager(DATA_DIR)
        self._import_engine = ImportEngine()
        self._current_template_name: str | None = None
        self._suppress_template_callback = False
        self._prev_body_size: float = 0
        self._raw_markdown: str = ""
        self._preview_job: str | None = None
        self._page_size = tk.StringVar(value="A4")

        self._setup_theme()
        self._build_ui()
        self._select_template(DEFAULT_TEMPLATE["name"])

    def _setup_theme(self) -> None:
        style = ttk.Style()
        available = style.theme_names()
        if "vista" in available:
            style.theme_use("vista")
        elif "clam" in available:
            style.theme_use("clam")

        style.configure("TFrame", background=C_BG)
        style.configure("Card.TLabelframe", background=C_CARD, relief="solid",
                        borderwidth=1, bordercolor=C_BORDER, padding=8)
        style.configure("Card.TLabelframe.Label", background=C_CARD,
                        foreground=C_ACCENT, font=("微软雅黑", 9, "bold"))
        style.configure("Warm.TLabelframe", background=C_CARD_WARM, relief="solid",
                        borderwidth=1, bordercolor=C_BORDER, padding=8)
        style.configure("Warm.TLabelframe.Label", background=C_CARD_WARM,
                        foreground="#b8861c", font=("微软雅黑", 9, "bold"))
        style.configure("TLabel", background=C_CARD, foreground=C_TEXT,
                        font=("微软雅黑", 9))
        style.configure("TButton", font=("微软雅黑", 9), padding=(12, 4))
        style.map("TButton", background=[("active", "#cce8e2")])
        style.configure("Small.TButton", padding=(6, 3), font=("微软雅黑", 8))
        style.configure("Convert.TButton", font=("微软雅黑", 11, "bold"),
                        padding=(36, 10))
        style.configure("Accent.TButton", font=("微软雅黑", 10, "bold"),
                        padding=(20, 8))
        style.configure("TRadiobutton", background=C_CARD, foreground=C_TEXT,
                        font=("微软雅黑", 9))
        style.map("TRadiobutton", background=[("active", C_CARD)])
        style.configure("TCheckbutton", background=C_CARD, foreground=C_TEXT,
                        font=("微软雅黑", 9))
        style.map("TCheckbutton", background=[("active", C_CARD)])
        style.configure("TCombobox", font=("微软雅黑", 9))
        style.configure("TEntry", font=("微软雅黑", 9))
        style.configure("TSpinbox", font=("微软雅黑", 9))
        style.configure("TNotebook", background=C_BG, borderwidth=0)
        style.configure("TNotebook.Tab", font=("微软雅黑", 10, "bold"),
                        padding=(20, 8))

    # ── Build UI ─────────────────────────────────────────────────

    def _build_ui(self) -> None:
        self._build_header()
        self._build_toolbar()
        self._build_notebook()
        self._build_footer()

    def _build_header(self) -> None:
        hdr = tk.Frame(self.root, bg=C_HEADER, height=44)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        inner = tk.Frame(hdr, bg=C_HEADER)
        inner.pack(fill="both", padx=16)
        tk.Label(inner, text="Markdown ↔ Word 转换器",
                 font=("微软雅黑", 14, "bold"),
                 fg=C_HEADER_TEXT, bg=C_HEADER).pack(side="left", pady=8)
        tk.Label(inner, text="Powered by MarkItDown · 30+ 格式导入",
                 font=("微软雅黑", 8), fg="#c8ddf8", bg=C_HEADER
                 ).pack(side="left", padx=(14, 0), pady=10)
        self._header_status = tk.Label(inner, text="● 就绪",
                                       font=("微软雅黑", 9),
                                       fg="#a8f0c8", bg=C_HEADER)
        self._header_status.pack(side="right", pady=8)

    def _build_toolbar(self) -> None:
        bar = tk.Frame(self.root, bg=C_CARD, height=38)
        bar.pack(fill="x")
        bar.pack_propagate(False)
        inner = tk.Frame(bar, bg=C_CARD)
        inner.pack(fill="both", padx=10, pady=3)

        tk.Label(inner, text="模板", font=("微软雅黑", 9, "bold"),
                 fg=C_MUTED, bg=C_CARD).pack(side="left", padx=(0, 6))
        self.template_var = tk.StringVar()
        self.template_combo = ttk.Combobox(
            inner, textvariable=self.template_var, state="readonly",
            values=self.tm.get_template_names(), font=("微软雅黑", 9), width=16)
        self.template_combo.pack(side="left")
        self.template_combo.bind("<<ComboboxSelected>>", self._on_template_select)
        ttk.Button(inner, text="保存", command=self._save_template,
                   style="Small.TButton").pack(side="left", padx=(6, 0))
        ttk.Button(inner, text="删除", command=self._delete_template,
                   style="Small.TButton").pack(side="left", padx=(4, 0))

        # Right side
        self.merge_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(inner, text="过滤AI空行", variable=self.merge_var,
                        command=self._schedule_word_preview
                        ).pack(side="right", padx=(0, 10))
        tk.Label(inner, text="纸张", font=("微软雅黑", 9, "bold"),
                 fg=C_MUTED, bg=C_CARD).pack(side="right", padx=(10, 4))
        ps = ttk.Combobox(inner, textvariable=self._page_size,
                          values=PAPER_NAMES, state="readonly",
                          width=5, font=("微软雅黑", 9))
        ps.pack(side="right")
        ps.bind("<<ComboboxSelected>>", lambda e: self._schedule_word_preview())

    def _build_notebook(self) -> None:
        """Two tabs: Import (full MarkItDown) and Editor (Markdown→Word)."""
        self._notebook = ttk.Notebook(self.root)
        self._notebook.pack(fill="both", expand=True, padx=8, pady=(4, 0))

        # ── Tab 1: Import ──
        self._tab_import = ttk.Frame(self._notebook)
        self._notebook.add(self._tab_import, text="  📥  导入转换  ")
        self._build_import_tab()

        # ── Tab 2: Editor ──
        self._tab_editor = ttk.Frame(self._notebook)
        self._notebook.add(self._tab_editor, text="  ✏️  Markdown 编辑  ")
        self._build_editor_tab()

    # ═══════════════════════════════════════════════════════════════
    # TAB 1: Import (Full MarkItDown)
    # ═══════════════════════════════════════════════════════════════

    def _build_import_tab(self) -> None:
        """File picker → convert → markdown preview → send to editor."""
        # ── Top: file selection ──
        top_bar = tk.Frame(self._tab_import, bg=C_BG)
        top_bar.pack(fill="x", padx=8, pady=(8, 4))

        tk.Label(top_bar, text="选择文件",
                 font=("微软雅黑", 11, "bold"), fg=C_TEXT, bg=C_BG
                 ).pack(side="left")

        self._import_path = tk.StringVar()
        ttk.Entry(top_bar, textvariable=self._import_path,
                  font=("微软雅黑", 9), width=50).pack(side="left", padx=(10, 6),
                                                        fill="x", expand=True)
        ttk.Button(top_bar, text="浏览...", command=self._browse_import,
                   style="Small.TButton").pack(side="left")
        ttk.Button(top_bar, text="转换为 Markdown", command=self._do_import,
                   style="Accent.TButton").pack(side="left", padx=(6, 0))

        # ── File info bar ──
        self._import_info = tk.Label(self._tab_import, text="",
                                     font=("微软雅黑", 8), fg=C_MUTED, bg=C_BG)
        self._import_info.pack(fill="x", padx=12, pady=(0, 2))

        # ── Format list (compact grid) ──
        fmt_frm = ttk.LabelFrame(self._tab_import, text="  支持格式  ", padding=6)
        fmt_frm.pack(fill="x", padx=8, pady=(0, 4))

        fmt_inner = tk.Frame(fmt_frm, bg=C_CARD)
        fmt_inner.pack(fill="x")
        # Build a compact text list of supported formats
        formats_by_cat: dict[str, list[str]] = {}
        for ext, (name, _) in sorted(FORMAT_MAP.items()):
            cat = {"Word 文档": "📄 文档", "PDF": "📄 文档",
                   "PowerPoint": "📊 演示/表格", "Excel": "📊 演示/表格",
                   "Excel 97": "📊 演示/表格", "CSV": "📊 演示/表格",
                   "HTML": "🌐 网页/电子书", "EPUB": "🌐 网页/电子书",
                   "Markdown": "📝 文本", "纯文本": "📝 文本",
                   "Jupyter": "📝 文本", "XML": "📝 文本", "JSON": "📝 文本",
                   "JPEG 图片": "🖼️ 媒体", "PNG 图片": "🖼️ 媒体",
                   "GIF 图片": "🖼️ 媒体", "BMP 图片": "🖼️ 媒体",
                   "WebP 图片": "🖼️ 媒体",
                   "MP3 音频": "🖼️ 媒体", "WAV 音频": "🖼️ 媒体",
                   "M4A 音频": "🖼️ 媒体", "OGG 音频": "🖼️ 媒体",
                   "Outlook 邮件": "📧 其他", "ZIP 压缩包": "📧 其他",
            }.get(name, "📧 其他")
            formats_by_cat.setdefault(cat, []).append(f"{ext} ({name})")

        for cat_name, items in sorted(formats_by_cat.items()):
            cat_frame = tk.Frame(fmt_inner, bg=C_CARD)
            cat_frame.pack(fill="x", pady=1)
            tk.Label(cat_frame, text=cat_name, font=("微软雅黑", 9, "bold"),
                     fg=C_ACCENT, bg=C_CARD, width=14, anchor="e"
                     ).pack(side="left", padx=(0, 6))
            tk.Label(cat_frame, text="  ".join(items),
                     font=("微软雅黑", 8), fg=C_MUTED, bg=C_CARD, anchor="w"
                     ).pack(side="left")

        # ── Main area: preview + actions ──
        panes = ttk.PanedWindow(self._tab_import, orient="horizontal")
        panes.pack(fill="both", expand=True, padx=8, pady=(2, 4))

        # Left: drag-drop area + file info
        left_frm = ttk.Frame(panes)
        panes.add(left_frm, weight=1)

        self._import_drop_label = tk.Label(
            left_frm,
            text="拖放文件到此处\n或点击上方「浏览...」选择文件\n\n支持 PDF · DOCX · PPTX · XLSX · EPUB\nHTML · 图片 · 音频 · CSV · ZIP 等",
            font=("微软雅黑", 10), fg=C_MUTED, bg=C_CARD,
            relief="solid", borderwidth=1, justify="center")
        self._import_drop_label.pack(fill="both", expand=True, padx=4, pady=4)

        # Right: markdown result preview
        right_frm = ttk.LabelFrame(panes, text="  Markdown 结果  ", padding=4)
        panes.add(right_frm, weight=1)

        self._import_result = tk.Text(
            right_frm, wrap="word", font=("Consolas", 10),
            relief="flat", borderwidth=0,
            bg=C_CARD, fg=C_TEXT,
            padx=10, pady=8,
            state="disabled")
        imp_scroll = ttk.Scrollbar(right_frm, orient="vertical",
                                   command=self._import_result.yview)
        self._import_result.configure(yscrollcommand=imp_scroll.set)
        self._import_result.pack(side="left", fill="both", expand=True)
        imp_scroll.pack(side="right", fill="y")

        # ── Bottom action bar ──
        bot = tk.Frame(self._tab_import, bg=C_BG)
        bot.pack(fill="x", padx=8, pady=(0, 8))

        self._send_btn = ttk.Button(
            bot, text="发送到 Markdown 编辑器 →",
            command=self._send_to_editor, style="Accent.TButton")
        self._send_btn.pack(side="right")
        self._send_btn.configure(state="disabled")

        self._import_status = tk.Label(bot, text="",
                                       font=("微软雅黑", 9), fg=C_MUTED, bg=C_BG)
        self._import_status.pack(side="right", padx=(0, 16))

    def _browse_import(self) -> None:
        path = filedialog.askopenfilename(
            title="选择要导入的文件", filetypes=FILE_FILTERS)
        if path:
            self._import_path.set(path)
            self._show_file_info(path)

    def _show_file_info(self, path: str) -> None:
        from pathlib import Path
        ext = Path(path).suffix.lower()
        name, desc = FORMAT_MAP.get(ext, ("未知", "未知格式"))
        size = Path(path).stat().st_size
        if size < 1024:
            size_str = f"{size} B"
        elif size < 1024 * 1024:
            size_str = f"{size / 1024:.1f} KB"
        else:
            size_str = f"{size / 1024 / 1024:.1f} MB"
        self._import_info.configure(
            text=f"  {Path(path).name}  |  格式: {name} ({ext})  |  大小: {size_str}  |  {desc}")

    def _do_import(self) -> None:
        path = self._import_path.get().strip()
        if not path or not os.path.exists(path):
            messagebox.showwarning("提示", "请先选择一个文件")
            return

        self._import_status.configure(text="正在转换...")
        self._import_drop_label.configure(text="转换中，请稍候...")
        self.root.update()

        def _run():
            result = self._import_engine.convert(path)
            self.root.after(0, lambda: self._on_import_done(result))

        threading.Thread(target=_run, daemon=True).start()

    def _on_import_done(self, result: ImportResult) -> None:
        if result.success:
            text = result.text
            self._import_result.configure(state="normal")
            self._import_result.delete("1.0", "end")
            self._import_result.insert("1.0", text)
            self._import_result.configure(state="disabled")

            self._import_drop_label.configure(
                text=f"✅ 转换成功\n\n{result.source_format}\n"
                     f"文件大小: {result.file_size_kb:.1f} KB\n"
                     f"输出: {len(text)} 字符",
                font=("微软雅黑", 10), fg="#16a34a")

            self._import_status.configure(
                text=f"转换完成 · {len(text)} 字符 · {result.converter_used}")
            self._send_btn.configure(state="normal")

            # Store for send-to-editor
            self._imported_text = text
            self._imported_path = result.source_path
        else:
            self._import_drop_label.configure(
                text=f"❌ 转换失败\n\n{result.source_format}\n{result.error[:200]}",
                font=("微软雅黑", 10), fg="#dc2626")
            self._import_status.configure(text="转换失败")
            cn_err = _translate_error(result.error)
            messagebox.showerror("导入失败", f"无法导入此文件。\n\n{cn_err}")

    def _send_to_editor(self) -> None:
        """Send imported Markdown to the editor tab and switch to it."""
        if not hasattr(self, "_imported_text") or not self._imported_text.strip():
            return

        self._editor.delete("1.0", "end")
        self._editor.insert("1.0", self._imported_text)
        self._editor.edit_modified(False)
        self._raw_markdown = self._imported_text

        if hasattr(self, "_imported_path"):
            self.input_path.set(self._imported_path)

        self._schedule_word_preview()
        self._notebook.select(self._tab_editor)
        self.status_var.set(f"已导入 {len(self._imported_text)} 字符到编辑器")

    # ═══════════════════════════════════════════════════════════════
    # TAB 2: Markdown Editor + Word Preview (existing functionality)
    # ═══════════════════════════════════════════════════════════════

    def _build_editor_tab(self) -> None:
        # ── Input bar (inside editor tab) ──
        inp_bar = tk.Frame(self._tab_editor, bg=C_BG)
        inp_bar.pack(fill="x", padx=8, pady=(8, 4))

        tk.Label(inp_bar, text="Markdown 文件",
                 font=("微软雅黑", 9, "bold"), fg=C_MUTED, bg=C_BG
                 ).pack(side="left", padx=(0, 6))
        self.input_path = tk.StringVar()
        ttk.Entry(inp_bar, textvariable=self.input_path, width=40,
                  font=("微软雅黑", 9)).pack(side="left", fill="x", expand=True)
        ttk.Button(inp_bar, text="加载 .md", command=self._browse_input,
                   style="Small.TButton").pack(side="left", padx=(6, 0))

        # ── Panes: Editor | Preview ──
        pw = ttk.PanedWindow(self._tab_editor, orient="horizontal")
        pw.pack(fill="both", expand=True, padx=8, pady=4)

        # Left: Editor
        left = ttk.Frame(pw)
        pw.add(left, weight=1)
        ef = ttk.LabelFrame(left, text="  Markdown 编辑  ", style="Card.TLabelframe")
        ef.pack(fill="both", expand=True)
        self._editor = tk.Text(
            ef, wrap="word", font=("Consolas", 11),
            relief="flat", borderwidth=0,
            bg=C_EDITOR_BG, fg=C_TEXT,
            insertbackground=C_ACCENT,
            padx=12, pady=10, undo=True, maxundo=50)
        ed_scroll = ttk.Scrollbar(ef, orient="vertical", command=self._editor.yview)
        self._editor.configure(yscrollcommand=ed_scroll.set)
        self._editor.pack(side="left", fill="both", expand=True)
        ed_scroll.pack(side="right", fill="y")
        self._editor.bind("<<Modified>>", self._on_editor_changed)
        self._editor.bind("<KeyRelease>", self._on_editor_changed)

        # Right: Word Preview
        right = ttk.Frame(pw)
        pw.add(right, weight=1)
        rf = ttk.LabelFrame(right, text="  Word 预览  ", style="Card.TLabelframe")
        rf.pack(fill="both", expand=True)
        self._prev_canvas = tk.Canvas(rf, bg="#dce4e8", highlightthickness=0)
        pv_scroll = ttk.Scrollbar(rf, orient="vertical",
                                  command=self._prev_canvas.yview)
        self._prev_canvas.configure(yscrollcommand=pv_scroll.set)
        self._prev_canvas.pack(side="left", fill="both", expand=True)
        pv_scroll.pack(side="right", fill="y")
        self._prev_pages = tk.Frame(self._prev_canvas, bg="#dce4e8")
        self._prev_canvas.create_window((0, 0), window=self._prev_pages,
                                        anchor="nw", tags="pages")
        self._prev_pages.bind("<Configure>",
                              lambda e: self._prev_canvas.configure(
                                  scrollregion=self._prev_canvas.bbox("all")))
        self._prev_canvas.bind("<Configure>",
                               lambda e: self._prev_canvas.itemconfig(
                                   "pages", width=e.width))
        self._prev_canvas.bind("<Enter>",
                               lambda e: self._prev_canvas.bind_all(
                                   "<MouseWheel>",
                                   lambda ev: self._prev_canvas.yview_scroll(
                                       -1 if ev.delta > 0 else 1, "units")))
        self._prev_canvas.bind("<Leave>",
                               lambda e: self._prev_canvas.unbind_all("<MouseWheel>"))

        # ── Settings panel (collapsible) ──
        self._build_editor_settings()

    def _build_editor_settings(self) -> None:
        sf = ttk.LabelFrame(self._tab_editor, text="  格式设置  ", style="Card.TLabelframe")
        sf.pack(fill="x", padx=8, pady=(2, 0))

        self._toggle_btn = ttk.Button(
            sf, text="▶ 展开格式设置", command=self._toggle_settings,
            style="Small.TButton")
        self._toggle_btn.pack(anchor="w", padx=6, pady=4)

        self._settings_canvas = tk.Canvas(sf, bg=C_CARD, highlightthickness=0, height=260)
        self._settings_inner = tk.Frame(self._settings_canvas, bg=C_CARD)
        self._settings_canvas.create_window((0, 0), window=self._settings_inner,
                                            anchor="nw", tags="settings_win")
        self._settings_inner.bind(
            "<Configure>",
            lambda e: self._settings_canvas.configure(
                scrollregion=self._settings_canvas.bbox("all")))
        self._settings_canvas.bind(
            "<Configure>",
            lambda e: self._settings_canvas.itemconfig("settings_win", width=e.width))
        self._settings_canvas.bind("<Enter>",
                                   lambda e: self._settings_canvas.bind_all(
                                       "<MouseWheel>",
                                       lambda ev: self._settings_canvas.yview_scroll(
                                           -1 if ev.delta > 0 else 1, "units")))
        self._settings_canvas.bind("<Leave>",
                                   lambda e: self._settings_canvas.unbind_all("<MouseWheel>"))
        self._settings_visible = False

        # Body + Para row
        r1 = tk.Frame(self._settings_inner, bg=C_CARD)
        r1.pack(fill="x")
        r1.columnconfigure(0, weight=1, uniform="s")
        r1.columnconfigure(1, weight=1, uniform="s")
        bf = tk.Frame(r1, bg=C_CARD_WARM, relief="solid", borderwidth=1)
        bf.grid(row=0, column=0, sticky="nsew", padx=(0, 3))
        self._build_body_settings(bf)
        pf = tk.Frame(r1, bg=C_CARD, relief="solid", borderwidth=1)
        pf.grid(row=0, column=1, sticky="nsew", padx=(3, 0))
        self._build_para_settings(pf)
        # Heading
        hf = tk.Frame(self._settings_inner, bg=C_CARD_WARM, relief="solid", borderwidth=1)
        hf.pack(fill="x", pady=(4, 0))
        self._build_heading_settings(hf)

    def _build_body_settings(self, parent: tk.Frame) -> None:
        inner = tk.Frame(parent, bg=C_CARD_WARM)
        inner.pack(fill="x", padx=8, pady=6)
        tk.Label(inner, text="正文", font=("微软雅黑", 9, "bold"),
                 fg=C_ACCENT, bg=C_CARD_WARM).pack(side="left", padx=(0, 8))
        for label, var_attr, width in [("字体", "body_font_var", 10), ("字号", "body_size_var", 4)]:
            ttk.Label(inner, text=label, background=C_CARD_WARM).pack(side="left")
            sv = tk.StringVar()
            setattr(self, var_attr, sv)
            cb = ttk.Combobox(inner, textvariable=sv,
                              values=(self._fonts if "font" in var_attr else SIZE_STRS),
                              width=width, state="readonly")
            cb.pack(side="left", padx=(2, 10))
            if "size" in var_attr:
                cb.bind("<<ComboboxSelected>>", self._on_body_size_changed)
            else:
                cb.bind("<<ComboboxSelected>>", self._mark_custom)
        ttk.Label(inner, text="颜色", background=C_CARD_WARM).pack(side="left")
        self.body_color_var = tk.StringVar()
        cf = tk.Frame(inner, bg=C_CARD_WARM)
        cf.pack(side="left", padx=(2, 0))
        self._build_color_presets(cf, self.body_color_var, C_CARD_WARM)

    def _build_para_settings(self, parent: tk.Frame) -> None:
        inner = tk.Frame(parent, bg=C_CARD)
        inner.pack(fill="x", padx=8, pady=6)
        tk.Label(inner, text="段落", font=("微软雅黑", 9, "bold"),
                 fg=C_ACCENT, bg=C_CARD).pack(side="left", padx=(0, 8))
        for label, var_name, frm in [("行距", "line_spacing_var", 6),
                                      ("段前", "para_before_var", 0),
                                      ("段后", "para_after_var", 0)]:
            ttk.Label(inner, text=label).pack(side="left")
            sv = tk.StringVar()
            setattr(self, var_name, sv)
            ttk.Spinbox(inner, textvariable=sv, from_=frm, to=100,
                        width=4).pack(side="left", padx=(2, 10))
        for v in [self.line_spacing_var, self.para_before_var, self.para_after_var]:
            v.trace_add("write", lambda *_: self._mark_custom())
        # Ensure they exist

    def _build_heading_settings(self, parent: tk.Frame) -> None:
        inner = tk.Frame(parent, bg=C_CARD_WARM)
        inner.pack(fill="x", padx=8, pady=6)
        self._heading_vars: dict[str, dict[str, tk.Variable]] = {}
        for lbl in ["H1", "H2", "H3", "H4", "H5", "H6"]:
            row_frm = tk.Frame(inner, bg=C_CARD_WARM)
            row_frm.pack(fill="x", pady=2)
            tk.Label(row_frm, text=lbl, width=3, font=("微软雅黑", 9, "bold"),
                     fg=C_ACCENT, bg=C_CARD_WARM).pack(side="left", padx=(0, 8))
            font_var = tk.StringVar()
            ttk.Combobox(row_frm, textvariable=font_var, values=self._fonts,
                         width=10, state="readonly").pack(side="left", padx=(0, 4))
            size_var = tk.StringVar()
            ttk.Combobox(row_frm, textvariable=size_var, values=SIZE_STRS,
                         width=4, state="readonly").pack(side="left", padx=(0, 4))
            color_var = tk.StringVar()
            cf = tk.Frame(row_frm, bg=C_CARD_WARM)
            cf.pack(side="left", padx=(0, 4))
            self._build_color_presets(cf, color_var, C_CARD_WARM)
            bold_var = tk.BooleanVar()
            ttk.Checkbutton(row_frm, text="粗体", variable=bold_var,
                            command=self._mark_custom).pack(side="left", padx=(6, 0))
            self._heading_vars[lbl.lower()] = {
                "font": font_var, "size": size_var,
                "color": color_var, "bold": bold_var,
            }

    def _build_footer(self) -> None:
        frm = tk.Frame(self.root, bg=C_BG, padx=12, pady=10)
        frm.pack(side="bottom", fill="x")
        self.convert_btn = ttk.Button(
            frm, text="转  换", command=self._save_docx, style="Convert.TButton")
        self.convert_btn.pack(side="right")
        self.status_var = tk.StringVar(value="就绪")
        tk.Label(frm, textvariable=self.status_var, font=("微软雅黑", 9),
                 fg=C_MUTED, bg=C_BG).pack(side="right", padx=(0, 16))

    def _build_color_presets(self, parent, target_var, bg=C_CARD):
        btn_frame = tk.Frame(parent, bg=bg)
        btn_frame.pack(side="left")
        for c in COLOR_PRESETS:
            swatch = tk.Label(btn_frame, bg=c, width=2, height=1,
                              relief="flat", borderwidth=1, cursor="hand2")
            swatch.pack(side="left", padx=1)
            swatch.bind("<Button-1>", lambda e, clr=c, tv=target_var: self._set_color(tv, clr))
        ttk.Button(btn_frame, text="⋯", width=3, style="Small.TButton",
                   command=lambda tv=target_var: self._pick_color(tv)).pack(side="left", padx=(4, 0))
        hex_entry = tk.Entry(btn_frame, textvariable=target_var, width=7,
                             font=("Consolas", 8), fg=C_TEXT,
                             bg=C_INPUT_BG, relief="solid", borderwidth=1)
        hex_entry.pack(side="left", padx=(4, 0))
        hex_entry.bind("<FocusOut>", lambda e: self._mark_custom())
        hex_entry.bind("<Return>", lambda e: self._mark_custom())

    def _toggle_settings(self):
        if self._settings_visible:
            self._settings_canvas.pack_forget()
            self._toggle_btn.configure(text="▶ 展开格式设置")
        else:
            self._settings_canvas.pack(fill="x", padx=4, pady=(2, 6))
            self._toggle_btn.configure(text="▼ 收起格式设置")
        self._settings_visible = not self._settings_visible

    # ── Shared: template, color, settings ────────────────────────

    def _select_template(self, name):
        t = self.tm.get_template(name)
        if not t: return
        self._current_template_name = name
        self._suppress_template_callback = True; s = t
        self._prev_body_size = float(s["body_size"])
        self.body_font_var.set(s["body_font"])
        self.body_size_var.set(str(s["body_size"]))
        self.body_color_var.set(s["body_color"])
        for lv in ["h1","h2","h3","h4","h5","h6"]:
            hv = self._heading_vars[lv]
            hv["font"].set(s[f"{lv}_font"]); hv["size"].set(str(s[f"{lv}_size"]))
            hv["color"].set(s[f"{lv}_color"]); hv["bold"].set(s[f"{lv}_bold"])
        self.line_spacing_var.set(str(s["line_spacing_pt"]))
        self.para_before_var.set(str(s["para_spacing_before"]))
        self.para_after_var.set(str(s["para_spacing_after"]))
        self.merge_var.set(s["merge_blank_lines"])
        self.template_var.set(name)
        self._suppress_template_callback = False
        self._schedule_word_preview()

    def _on_template_select(self, event=None):
        name = self.template_var.get()
        if name and name != self._current_template_name:
            self._select_template(name)

    def _mark_custom(self, event=None):
        if self._suppress_template_callback: return
        if self._current_template_name and not self._current_template_name.startswith("(自定义)"):
            self._current_template_name = None
            vals = list(self.template_combo["values"])
            if "(自定义)" not in vals:
                self.template_combo["values"] = ["(自定义)"] + vals
            self.template_var.set("(自定义)")
        try: self._prev_body_size = float(self.body_size_var.get())
        except ValueError: pass
        self._schedule_word_preview()

    def _on_body_size_changed(self, event=None):
        try: new_size = float(self.body_size_var.get())
        except ValueError: return
        if self._prev_body_size > 0 and self._prev_body_size != new_size:
            ratio = new_size / self._prev_body_size
            self._suppress_template_callback = True
            for lv in ["h1","h2","h3","h4","h5","h6"]:
                hv = self._heading_vars[lv]
                try:
                    old = float(hv["size"].get())
                    hv["size"].set(str(round(old * ratio)))
                except ValueError: pass
            self._suppress_template_callback = False
        self._prev_body_size = new_size; self._mark_custom()

    def _save_template(self):
        name = tk.simpledialog.askstring("保存模板", "请输入模板名称:", parent=self.root)
        if not name: return
        self.tm.save_template(name, self._collect_settings())
        self._refresh_template_list()
        self._select_template(name)
        self.status_var.set(f"模板「{name}」已保存")

    def _delete_template(self):
        name = self.template_var.get()
        if not name: return
        t = self.tm.get_template(name)
        if t and t.get("is_builtin"):
            messagebox.showwarning("提示", "内置模板不能删除"); return
        if not messagebox.askyesno("确认", f"确定要删除模板「{name}」吗?"): return
        self.tm.delete_template(name)
        self._refresh_template_list()
        self._select_template(DEFAULT_TEMPLATE["name"])
        self.status_var.set(f"模板「{name}」已删除")

    def _refresh_template_list(self):
        self.template_combo["values"] = self.tm.get_template_names()

    def _collect_settings(self) -> dict:
        def _f(v): return v.get()
        def _fl(v): return float(v.get())
        s = {
            "body_font": _f(self.body_font_var), "body_size": _fl(self.body_size_var),
            "body_color": _f(self.body_color_var),
            "line_spacing_pt": _fl(self.line_spacing_var),
            "para_spacing_before": _fl(self.para_before_var),
            "para_spacing_after": _fl(self.para_after_var),
            "merge_blank_lines": bool(self.merge_var.get()),
            "page_size": self._page_size.get(),
            "margin_top_cm": 2.54, "margin_bottom_cm": 2.54,
            "margin_left_cm": 3.18, "margin_right_cm": 3.18,
        }
        for lv in ["h1","h2","h3","h4","h5","h6"]:
            hv = self._heading_vars[lv]
            s[f"{lv}_font"] = _f(hv["font"]); s[f"{lv}_size"] = _fl(hv["size"])
            s[f"{lv}_color"] = _f(hv["color"]); s[f"{lv}_bold"] = bool(hv["bold"].get())
        ct = self.tm.get_template(self.template_var.get())
        if ct:
            for k in ("margin_top_cm","margin_bottom_cm","margin_left_cm","margin_right_cm"):
                s[k] = ct[k]
        return s

    def _set_color(self, tv, clr): tv.set(clr); self._mark_custom()
    def _pick_color(self, tv):
        c = colorchooser.askcolor(initialcolor=tv.get(), title="选择颜色")
        if c and c[1]: tv.set(c[1]); self._mark_custom()

    # ── Editor events ────────────────────────────────────────────

    def _browse_input(self):
        path = filedialog.askopenfilename(
            title="加载 Markdown", filetypes=[("Markdown", "*.md"), ("所有文件", "*.*")])
        if path:
            self.input_path.set(path)
            try:
                with open(path, "r", encoding="utf-8") as f:
                    content = f.read()
                self._editor.delete("1.0", "end")
                self._editor.insert("1.0", content)
                self._editor.edit_modified(False)
                self._raw_markdown = content
                self._schedule_word_preview()
                self.status_var.set(f"已加载 {os.path.basename(path)}")
            except OSError: pass

    def _on_editor_changed(self, event=None):
        """Handle both keystrokes and programmatic edits. Debounces preview."""
        # Skip pure navigation keys
        if event and event.type == tk.EventType.KeyRelease:
            if event.keysym in (
                "Left","Right","Up","Down","Home","End",
                "Page_Up","Page_Down","Shift_L","Shift_R",
                "Control_L","Control_R","Alt_L","Alt_R",
                "Caps_Lock","Num_Lock","Scroll_Lock"):
                return
        # Read editor content once
        if self._editor.edit_modified():
            self._editor.edit_modified(False)
        self._raw_markdown = self._editor.get("1.0", "end-1c")
        self._schedule_word_preview()

    # ── Word Preview ─────────────────────────────────────────────

    def _schedule_word_preview(self):
        if self._preview_job is not None:
            self.root.after_cancel(self._preview_job)
        self._preview_job = self.root.after(40, self._render_word_preview)

    def _render_word_preview(self):
        self._preview_job = None
        md_text = self._raw_markdown
        for w in self._prev_pages.winfo_children(): w.destroy()
        if not md_text.strip():
            tk.Label(self._prev_pages, text="在左侧编辑器中编写或加载 Markdown → 右边实时预览",
                     font=("微软雅黑", 11), fg=C_MUTED, bg="#dce4e8").pack(pady=40)
            return
        if self.merge_var.get():
            md_text = _clean_ai_blank_lines(md_text)
        s = self._collect_settings()
        ps_name = self._page_size.get()
        pw_cm, ph_cm = PAPER_SIZES.get(ps_name, PAPER_SIZES["A4"])
        usable_cm = ph_cm - s["margin_top_cm"] - s["margin_bottom_cm"]
        usable_pt = usable_cm / 2.54 * 72
        line_spacing = max(s["line_spacing_pt"], 1)
        lpp = max(int(usable_pt / line_spacing), 10)
        # Dynamic page width + font scaling
        canvas_width = self._prev_canvas.winfo_width()
        page_width = max(260, canvas_width - 40) if canvas_width > 10 else 480
        # Scale preview font: 480px→12pt base, floor 9pt
        preview_scale = max(0.75, min(1.0, page_width / 480))
        base_font_size = int(12 * preview_scale)
        page_ratio = pw_cm / ph_cm
        page_height = int(page_width / page_ratio)

        html = markdown.markdown(md_text, extensions=["tables", "fenced_code"])
        soup = BeautifulSoup(html, "html.parser")
        body = soup.find("body")
        blocks = list(body.children if body else soup.children)
        page_blocks = self._split_blocks_to_pages(blocks, lpp)

        for pg_idx, pg_blocks in enumerate(page_blocks):
            pf = tk.Frame(self._prev_pages, bg="#dce4e8"); pf.pack(pady=12)
            shadow = tk.Frame(pf, bg="#c8d0d8", width=page_width+4, height=page_height+4)
            shadow.pack(); shadow.pack_propagate(False)
            page_bg = tk.Frame(shadow, bg="#ffffff", width=page_width, height=page_height)
            page_bg.place(relx=0.5, rely=0.5, anchor="center"); page_bg.pack_propagate(False)
            pt = tk.Text(page_bg, wrap="word", font=("宋体", base_font_size),
                         relief="flat", borderwidth=0, bg="#ffffff", fg=C_TEXT,
                         padx=int(20 * preview_scale), pady=int(14 * preview_scale),
                         state="disabled")
            pt.pack(fill="both", expand=True)
            self._configure_preview_tags_for(pt, s, preview_scale)
            for element in pg_blocks:
                if isinstance(element, NavigableString):
                    text = str(element).strip()
                    if text:
                        pt.configure(state="normal"); pt.insert("end", text+"\n", "p"); pt.configure(state="disabled")
                elif isinstance(element, Tag):
                    self._render_block_to(pt, element)
            tk.Label(page_bg, text=f"第 {pg_idx+1} 页  |  {ps_name}  {pw_cm}×{ph_cm}cm",
                     font=("微软雅黑", 7), fg=C_MUTED, bg="#ffffff").pack(side="bottom", pady=(0,4))

    def _split_blocks_to_pages(self, blocks, lpp):
        pages, current, lc = [], [], 0
        cpl = 480 // 7
        for blk in blocks:
            if isinstance(blk, Tag):
                text = blk.get_text()
                lines = [text[i:i+cpl] for i in range(0, len(text), cpl)]
                bl = max(1, len(lines))
                if blk.name in ("h1","h2"): bl += 2
            else: bl = 1
            if lc + bl > lpp and current:
                pages.append(current); current, lc = [], 0
            current.append(blk); lc += bl
        if current: pages.append(current)
        return pages if pages else [blocks]

    def _configure_preview_tags_for(self, pt, s, scale=1.0):
        def _sc(size_pt):
            return max(7, int(int(size_pt) * scale))
        for lv in range(1,7):
            pfx = f"h{lv}"
            pt.tag_configure(f"h{lv}", font=(s[f"{pfx}_font"], _sc(s[f"{pfx}_size"]),
                             "bold" if s[f"{pfx}_bold"] else "normal"),
                             foreground=s[f"{pfx}_color"],
                             spacing1=int((14 if lv<=2 else 6)*scale), spacing3=int(4*scale))
        pt.tag_configure("p", font=(s["body_font"], _sc(s["body_size"])),
                         foreground=s["body_color"], spacing1=int(2*scale), spacing3=int(2*scale))
        pt.tag_configure("bold", font=(s["body_font"], _sc(s["body_size"]), "bold"),
                         foreground=s["body_color"])
        pt.tag_configure("italic", font=(s["body_font"], _sc(s["body_size"]), "italic"),
                         foreground=s["body_color"])
        pt.tag_configure("code", font=("Courier New", _sc(int(s["body_size"])-1)),
                         foreground="#c0392b")
        pt.tag_configure("pre", font=("Courier New", _sc(int(s["body_size"])-1)),
                         foreground=C_TEXT, background=C_INPUT_BG,
                         lmargin1=int(16*scale), lmargin2=int(16*scale),
                         spacing1=int(4*scale), spacing3=int(4*scale))
        pt.tag_configure("quote", font=(s["body_font"], _sc(s["body_size"]), "italic"),
                         foreground=C_MUTED, lmargin1=int(24*scale), lmargin2=int(24*scale),
                         background=C_INPUT_BG)
        pt.tag_configure("li", font=(s["body_font"], _sc(s["body_size"])),
                         foreground=s["body_color"],
                         lmargin1=int(20*scale), lmargin2=int(20*scale))
        pt.tag_configure("link", font=(s["body_font"], _sc(s["body_size"]), "underline"),
                         foreground=C_ACCENT)
        pt.tag_configure("hr", font=("微软雅黑", _sc(6)), foreground=C_BORDER, justify="center")

    def _render_block_to(self, pt, element):
        tag = element.name
        pt.configure(state="normal")
        if tag in ("h1","h2","h3","h4","h5","h6"):
            pt.insert("end", element.get_text(strip=True)+"\n", tag)
        elif tag == "p":
            self._render_inline_to(pt, element, "p"); pt.insert("end", "\n")
        elif tag in ("ul","ol"):
            for li in element.find_all("li", recursive=False):
                pt.insert("end", "  -  ", "li")
                self._render_inline_to(pt, li, "li"); pt.insert("end", "\n")
        elif tag == "pre":
            for line in element.get_text().rstrip("\n").split("\n"):
                pt.insert("end", "  "+line+"\n", "pre")
        elif tag == "blockquote":
            self._render_inline_to(pt, element, "quote"); pt.insert("end", "\n")
        elif tag == "hr": pt.insert("end", "─"*40+"\n", "hr")
        elif tag == "table":
            rows = element.find_all("tr")
            for i, row in enumerate(rows):
                parts = [c.get_text(strip=True) for c in row.find_all(["td","th"])]
                pt.insert("end", "  "+"  │  ".join(parts)+"\n", "bold" if i==0 else "p")
            pt.insert("end", "\n")
        pt.configure(state="disabled")

    def _render_inline_to(self, pt, element, default_tag):
        for child in element.children:
            if isinstance(child, NavigableString): pt.insert("end", str(child), default_tag)
            elif isinstance(child, Tag):
                tn = child.name
                if tn in ("strong","b"): pt.insert("end", child.get_text(), "bold")
                elif tn in ("em","i"): pt.insert("end", child.get_text(), "italic")
                elif tn == "code": pt.insert("end", f"`{child.get_text()}`", "code")
                elif tn == "a": pt.insert("end", child.get_text(), "link")
                elif tn == "br": pt.insert("end", "\n", default_tag)
                else: self._render_inline_to(pt, child, default_tag)

    # ── Export ───────────────────────────────────────────────────

    def _save_docx(self):
        md_text = self._raw_markdown
        if not md_text.strip():
            messagebox.showwarning("提示", "没有可转换的内容"); return
        from pathlib import Path
        init = ""
        if self.input_path.get().strip():
            init = os.path.splitext(self.input_path.get())[0] + ".docx"
        else:
            init = os.path.join(os.path.expanduser("~"), "output.docx")
        out = filedialog.asksaveasfilename(
            title="保存 Word 文件", defaultextension=".docx",
            initialfile=Path(init).name,
            initialdir=Path(init).parent if Path(init).parent.exists() else None,
            filetypes=[("Word 文档", "*.docx")])
        if not out: return
        template = self._collect_settings()
        merge = self.merge_var.get()
        self.convert_btn.configure(state="disabled")
        self.status_var.set("正在转换...")
        self._header_status.configure(text="● 转换中", fg="#fbbf24")
        def _run():
            try:
                convert_text(md_text, out, template, merge)
                self.root.after(0, lambda: self._on_save_done(out))
            except Exception as e:
                self.root.after(0, lambda: self._on_save_error(str(e)))
        threading.Thread(target=_run, daemon=True).start()

    def _on_save_done(self, out):
        self.convert_btn.configure(state="normal")
        self.status_var.set(f"转换完成 → {out}")
        self._header_status.configure(text="● 就绪", fg="#a8f0c8")
        if messagebox.askyesno("完成", f"转换成功!\n\n文件: {out}\n\n是否打开文件?"):
            os.startfile(out)

    def _on_save_error(self, err_msg):
        self.convert_btn.configure(state="normal")
        self.status_var.set("转换失败")
        self._header_status.configure(text="● 错误", fg="#f87171")
        cn_msg = _translate_error(err_msg)
        messagebox.showerror("转换失败", cn_msg)


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
