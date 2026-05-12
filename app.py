"""Markdown to Word Converter — Side-by-side editor & preview."""

import os
import re
import threading
import tkinter as tk
from tkinter import colorchooser, filedialog, font as tkfont
from tkinter import messagebox, ttk

import markdown
from bs4 import BeautifulSoup, NavigableString, Tag

from converter import convert_text, _clean_ai_blank_lines
from templates import (COLOR_PRESETS, SIZE_LIST, TemplateManager,
                       BUILTIN_TEMPLATES)

import sys as _sys
if getattr(_sys, 'frozen', False):
    DATA_DIR = os.path.dirname(_sys.executable)
else:
    DATA_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_TEMPLATE = BUILTIN_TEMPLATES[0]

SIZE_STRS = [str(s) for s in SIZE_LIST]
PAD = {"padx": 8, "pady": (3, 3)}

# Paper sizes: (width_cm, height_cm)
PAPER_SIZES = {
    "A4":  (21.0, 29.7),
    "A3":  (29.7, 42.0),
    "8K":  (26.0, 36.8),
    "16K": (18.4, 26.0),
}
PAPER_NAMES = list(PAPER_SIZES.keys())

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


def _get_system_fonts() -> list[str]:
    """Return all available system fonts, with CJK/common fonts first."""
    try:
        raw = sorted(set(tkfont.families()))
    except Exception:
        raw = []

    # Prioritize CJK + common fonts at the top
    priority = [
        "微软雅黑", "宋体", "黑体", "仿宋", "楷体", "方正小标宋简体",
        "Arial", "Times New Roman", "Calibri", "Courier New",
        "Consolas", "Segoe UI", "Verdana", "Georgia",
    ]
    seen = set()
    result = []
    for p in priority:
        if p in raw and p not in seen:
            result.append(p)
            seen.add(p)
    for f in raw:
        if f not in seen:
            result.append(f)
            seen.add(f)
    return result


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Markdown → Word 转换器")
        self.root.geometry("1100x820")
        self.root.minsize(700, 520)
        self.root.configure(bg=C_BG)

        self._fonts = _get_system_fonts()
        self.tm = TemplateManager(DATA_DIR)
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
        style.configure("Small.TButton", padding=(6, 3), font=("微软雅黑", 8))
        style.configure("Convert.TButton", font=("微软雅黑", 11, "bold"),
                        padding=(36, 10))
        style.configure("Save.TButton", font=("微软雅黑", 8), padding=(10, 3))
        style.configure("TRadiobutton", background=C_CARD, foreground=C_TEXT,
                        font=("微软雅黑", 9))
        style.map("TRadiobutton", background=[("active", C_CARD)])
        style.configure("TCheckbutton", background=C_CARD, foreground=C_TEXT,
                        font=("微软雅黑", 9))
        style.map("TCheckbutton", background=[("active", C_CARD)])
        style.configure("TCombobox", font=("微软雅黑", 9))
        style.configure("TEntry", font=("微软雅黑", 9))
        style.configure("TSpinbox", font=("微软雅黑", 9))

    # ── Build UI ─────────────────────────────────────────────────

    def _build_ui(self) -> None:
        self._build_header()
        self._build_toolbar()
        self._build_footer()               # pack first so always visible
        self._build_settings_panel()
        self._build_main_panes()           # fills remaining space

    def _build_header(self) -> None:
        hdr = tk.Frame(self.root, bg=C_HEADER, height=48)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        inner = tk.Frame(hdr, bg=C_HEADER)
        inner.pack(fill="both", padx=16)
        tk.Label(inner, text="Markdown → Word 转换器",
                 font=("微软雅黑", 14, "bold"),
                 fg=C_HEADER_TEXT, bg=C_HEADER).pack(side="left", pady=10)
        tk.Label(inner, text="左边编辑 · 右边预览 · 系统字体 · 一键导出",
                 font=("微软雅黑", 8), fg="#c8ddf8", bg=C_HEADER
                 ).pack(side="left", padx=(14, 0), pady=12)
        self._header_status = tk.Label(inner, text="● 就绪",
                                       font=("微软雅黑", 9),
                                       fg="#a8f0c8", bg=C_HEADER)
        self._header_status.pack(side="right", pady=10)

    def _build_toolbar(self) -> None:
        """Compact toolbar: template selector + load/save buttons."""
        bar = tk.Frame(self.root, bg=C_CARD, height=40)
        bar.pack(fill="x")
        bar.pack_propagate(False)
        inner = tk.Frame(bar, bg=C_CARD)
        inner.pack(fill="both", padx=10, pady=4)

        tk.Label(inner, text="模板", font=("微软雅黑", 9, "bold"),
                 fg=C_MUTED, bg=C_CARD).pack(side="left", padx=(0, 6))

        self.template_var = tk.StringVar()
        self.template_combo = ttk.Combobox(
            inner, textvariable=self.template_var, state="readonly",
            values=self.tm.get_template_names(),
            font=("微软雅黑", 9), width=16)
        self.template_combo.pack(side="left")
        self.template_combo.bind("<<ComboboxSelected>>", self._on_template_select)

        ttk.Button(inner, text="保存模板", command=self._save_template,
                   style="Save.TButton").pack(side="left", padx=(8, 0))
        ttk.Button(inner, text="删除", command=self._delete_template,
                   style="Save.TButton").pack(side="left", padx=(4, 0))

        # File load on right side
        self.input_path = tk.StringVar()
        ttk.Entry(inner, textvariable=self.input_path, width=24,
                  font=("微软雅黑", 9)).pack(side="right", padx=(4, 0))
        ttk.Button(inner, text="加载文件", command=self._browse_input,
                   style="Small.TButton").pack(side="right")

        self.merge_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(inner, text="过滤AI空行", variable=self.merge_var
                        ).pack(side="right", padx=(0, 10))

        # Paper size
        tk.Label(inner, text="纸张", font=("微软雅黑", 9, "bold"),
                 fg=C_MUTED, bg=C_CARD).pack(side="right", padx=(10, 4))
        ps = ttk.Combobox(inner, textvariable=self._page_size,
                          values=PAPER_NAMES, state="readonly",
                          width=5, font=("微软雅黑", 9))
        ps.pack(side="right")
        ps.bind("<<ComboboxSelected>>", lambda e: self._schedule_word_preview())

    def _build_main_panes(self) -> None:
        """Left: markdown editor. Right: Word preview. Resizable split."""
        pw = ttk.PanedWindow(self.root, orient="horizontal")
        pw.pack(fill="both", expand=True, padx=8, pady=(4, 0))

        # ── Left pane: Markdown editor ──
        left = ttk.Frame(pw)
        pw.add(left, weight=1)

        editor_frm = ttk.LabelFrame(left, text="  Markdown 编辑  ",
                                    style="Card.TLabelframe")
        editor_frm.pack(fill="both", expand=True)

        self._editor = tk.Text(
            editor_frm, wrap="word",
            font=("Consolas", 11),
            relief="flat", borderwidth=0,
            bg=C_EDITOR_BG, fg=C_TEXT,
            insertbackground=C_ACCENT,
            padx=12, pady=10,
            undo=True, maxundo=50)
        ed_scroll = ttk.Scrollbar(editor_frm, orient="vertical",
                                  command=self._editor.yview)
        self._editor.configure(yscrollcommand=ed_scroll.set)
        self._editor.pack(side="left", fill="both", expand=True)
        ed_scroll.pack(side="right", fill="y")
        self._editor.bind("<<Modified>>", self._on_editor_modified)
        self._editor.bind("<KeyRelease>", self._on_editor_key)

        # ── Right pane: Word preview (page-style) ──
        right = ttk.Frame(pw)
        pw.add(right, weight=1)

        prev_frm = ttk.LabelFrame(right, text="  Word 预览  ",
                                  style="Card.TLabelframe")
        prev_frm.pack(fill="both", expand=True)

        # Canvas for scrollable page-style preview with gray background
        self._prev_canvas = tk.Canvas(prev_frm, bg="#dce4e8",
                                      highlightthickness=0)
        pv_scroll = ttk.Scrollbar(prev_frm, orient="vertical",
                                  command=self._prev_canvas.yview)
        self._prev_canvas.configure(yscrollcommand=pv_scroll.set)
        self._prev_canvas.pack(side="left", fill="both", expand=True)
        pv_scroll.pack(side="right", fill="y")

        # Frame holding all page widgets
        self._prev_pages = tk.Frame(self._prev_canvas, bg="#dce4e8")
        self._prev_canvas.create_window((0, 0), window=self._prev_pages,
                                        anchor="nw", tags="pages")
        self._prev_pages.bind(
            "<Configure>",
            lambda e: self._prev_canvas.configure(
                scrollregion=self._prev_canvas.bbox("all")))
        self._prev_canvas.bind(
            "<Configure>",
            lambda e: self._prev_canvas.itemconfig("pages", width=e.width))

        # Mouse wheel
        self._prev_canvas.bind("<Enter>",
                               lambda e: self._prev_canvas.bind_all(
                                   "<MouseWheel>",
                                   lambda ev: self._prev_canvas.yview_scroll(
                                       -1 if ev.delta > 0 else 1, "units")))
        self._prev_canvas.bind("<Leave>",
                               lambda e: self._prev_canvas.unbind_all(
                                   "<MouseWheel>"))

    def _build_settings_panel(self) -> None:
        """Collapsible settings section below the editor/preview panes."""
        self._settings_frm = ttk.LabelFrame(self.root, text="  格式设置  ",
                                            style="Card.TLabelframe")
        self._settings_frm.pack(fill="x", padx=8, pady=(4, 0))

        self._toggle_btn = ttk.Button(
            self._settings_frm, text="▶ 展开格式设置",
            command=self._toggle_settings, style="Small.TButton")
        self._toggle_btn.pack(anchor="w", padx=6, pady=4)

        # Scrollable canvas wrapper
        self._settings_canvas = tk.Canvas(self._settings_frm, bg=C_CARD,
                                          highlightthickness=0, height=260)
        self._settings_inner = tk.Frame(self._settings_canvas, bg=C_CARD)
        self._settings_win = self._settings_canvas.create_window(
            (0, 0), window=self._settings_inner, anchor="nw")
        self._settings_inner.bind(
            "<Configure>",
            lambda e: self._settings_canvas.configure(
                scrollregion=self._settings_canvas.bbox("all")))
        self._settings_canvas.bind(
            "<Configure>",
            lambda e: self._settings_canvas.itemconfig(
                self._settings_win, width=e.width))
        self._settings_canvas.bind("<Enter>",
                                   lambda e: self._settings_canvas.bind_all(
                                       "<MouseWheel>",
                                       lambda ev: self._settings_canvas.yview_scroll(
                                           -1 if ev.delta > 0 else 1, "units")))
        self._settings_canvas.bind("<Leave>",
                                   lambda e: self._settings_canvas.unbind_all(
                                       "<MouseWheel>"))
        self._settings_visible = False

        # ── Row 1: Body + Paragraph side by side ──
        row1 = tk.Frame(self._settings_inner, bg=C_CARD)
        row1.pack(fill="x")
        row1.columnconfigure(0, weight=1, uniform="s")
        row1.columnconfigure(1, weight=1, uniform="s")

        bf = tk.Frame(row1, bg=C_CARD_WARM, relief="solid", borderwidth=1)
        bf.grid(row=0, column=0, sticky="nsew", padx=(0, 3))
        self._build_body_settings(bf)

        pf = tk.Frame(row1, bg=C_CARD, relief="solid", borderwidth=1)
        pf.grid(row=0, column=1, sticky="nsew", padx=(3, 0))
        self._build_para_settings(pf)

        # ── Row 2: Heading settings (one row per level) ──
        hf = tk.Frame(self._settings_inner, bg=C_CARD_WARM,
                      relief="solid", borderwidth=1)
        hf.pack(fill="x", pady=(4, 0))
        self._build_heading_settings(hf)

    def _build_body_settings(self, parent: tk.Frame) -> None:
        inner = tk.Frame(parent, bg=C_CARD_WARM)
        inner.pack(fill="x", padx=8, pady=6)
        tk.Label(inner, text="正文", font=("微软雅黑", 9, "bold"),
                 fg=C_ACCENT, bg=C_CARD_WARM).pack(side="left", padx=(0, 8))

        for label, var_attr, width in [("字体", "body_font_var", 13),
                                        ("字号", "body_size_var", 4)]:
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

        setattr(self, "body_font_var", getattr(self, "body_font_var", tk.StringVar(value="宋体")))
        setattr(self, "body_size_var", getattr(self, "body_size_var", tk.StringVar(value="12")))

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
            ttk.Combobox(row_frm, textvariable=font_var,
                         values=self._fonts, width=14,
                         state="readonly").pack(side="left", padx=(0, 4))

            size_var = tk.StringVar()
            ttk.Combobox(row_frm, textvariable=size_var,
                         values=SIZE_STRS, width=4, state="readonly"
                         ).pack(side="left", padx=(0, 4))

            color_var = tk.StringVar()
            cf = tk.Frame(row_frm, bg=C_CARD_WARM)
            cf.pack(side="left", padx=(0, 4))
            self._build_color_presets(cf, color_var, C_CARD_WARM)

            bold_var = tk.BooleanVar()
            ttk.Checkbutton(row_frm, text="粗体", variable=bold_var,
                            command=self._mark_custom
                            ).pack(side="left", padx=(6, 0))

            self._heading_vars[lbl.lower()] = {
                "font": font_var, "size": size_var,
                "color": color_var, "bold": bold_var,
            }

    def _build_footer(self) -> None:
        frm = tk.Frame(self.root, bg=C_BG, padx=12, pady=10)
        frm.pack(side="bottom", fill="x")

        self.convert_btn = ttk.Button(
            frm, text="转  换", command=self._save_docx,
            style="Convert.TButton")
        self.convert_btn.pack(side="right")

        self.status_var = tk.StringVar(value="就绪")
        tk.Label(frm, textvariable=self.status_var, font=("微软雅黑", 9),
                 fg=C_MUTED, bg=C_BG).pack(side="right", padx=(0, 16))

    def _toggle_settings(self) -> None:
        if self._settings_visible:
            self._settings_canvas.pack_forget()
            self._toggle_btn.configure(text="▶ 展开格式设置")
        else:
            self._settings_canvas.pack(fill="x", padx=4, pady=(2, 6))
            self._toggle_btn.configure(text="▼ 收起格式设置")
        self._settings_visible = not self._settings_visible

    # ── Color presets ────────────────────────────────────────────

    def _build_color_presets(self, parent: tk.Frame, target_var: tk.StringVar,
                             bg: str = C_CARD) -> None:
        btn_frame = tk.Frame(parent, bg=bg)
        btn_frame.pack(side="left")
        for c in COLOR_PRESETS:
            swatch = tk.Label(btn_frame, bg=c, width=2, height=1,
                              relief="flat", borderwidth=1, cursor="hand2")
            swatch.pack(side="left", padx=1)
            swatch.bind("<Button-1>",
                        lambda e, clr=c, tv=target_var: self._set_color(tv, clr))
        ttk.Button(btn_frame, text="⋯", width=3, style="Small.TButton",
                   command=lambda tv=target_var: self._pick_color(tv)
                   ).pack(side="left", padx=(4, 0))
        hex_entry = tk.Entry(btn_frame, textvariable=target_var, width=7,
                             font=("Consolas", 8), fg=C_TEXT,
                             bg=C_INPUT_BG, relief="solid", borderwidth=1)
        hex_entry.pack(side="left", padx=(4, 0))
        hex_entry.bind("<FocusOut>", lambda e: self._mark_custom())
        hex_entry.bind("<Return>", lambda e: self._mark_custom())

    def _set_color(self, target_var: tk.StringVar, color: str) -> None:
        target_var.set(color)
        self._mark_custom()

    def _pick_color(self, target_var: tk.StringVar) -> None:
        color = colorchooser.askcolor(initialcolor=target_var.get(), title="选择颜色")
        if color and color[1]:
            target_var.set(color[1])
            self._mark_custom()

    # ── Editor events ────────────────────────────────────────────

    def _on_editor_modified(self, event=None) -> None:
        """Handle programmatic changes (paste, undo, file load)."""
        if self._editor.edit_modified():
            self._editor.edit_modified(False)
            self._raw_markdown = self._editor.get("1.0", "end-1c")
            self._schedule_word_preview()

    def _on_editor_key(self, event=None) -> None:
        """Handle each keystroke for true real-time preview."""
        # Skip navigation keys (arrows, home, end, etc.)
        if event and event.keysym in (
            "Left", "Right", "Up", "Down", "Home", "End",
            "Page_Up", "Page_Down", "Shift_L", "Shift_R",
            "Control_L", "Control_R", "Alt_L", "Alt_R",
            "Caps_Lock", "Num_Lock", "Scroll_Lock",
        ):
            return
        self._raw_markdown = self._editor.get("1.0", "end-1c")
        self._schedule_word_preview()

    # ── Template operations ──────────────────────────────────────

    def _select_template(self, name: str) -> None:
        t = self.tm.get_template(name)
        if not t:
            return
        self._current_template_name = name
        self._suppress_template_callback = True
        s = t

        self._prev_body_size = float(s["body_size"])
        self.body_font_var.set(s["body_font"])
        self.body_size_var.set(str(s["body_size"]))
        self.body_color_var.set(s["body_color"])

        for level in ["h1", "h2", "h3", "h4", "h5", "h6"]:
            hv = self._heading_vars[level]
            hv["font"].set(s[f"{level}_font"])
            hv["size"].set(str(s[f"{level}_size"]))
            hv["color"].set(s[f"{level}_color"])
            hv["bold"].set(s[f"{level}_bold"])

        self.line_spacing_var.set(str(s["line_spacing_pt"]))
        self.para_before_var.set(str(s["para_spacing_before"]))
        self.para_after_var.set(str(s["para_spacing_after"]))
        self.merge_var.set(s["merge_blank_lines"])

        self.template_var.set(name)
        self._suppress_template_callback = False
        self._schedule_word_preview()

    def _on_template_select(self, event=None) -> None:
        name = self.template_var.get()
        if name and name != self._current_template_name:
            self._select_template(name)

    def _mark_custom(self, event=None) -> None:
        if self._suppress_template_callback:
            return
        if self._current_template_name and not self._current_template_name.startswith("(自定义)"):
            self._current_template_name = None
            custom_label = "(自定义)"
            current_values = list(self.template_combo["values"])
            if custom_label not in current_values:
                self.template_combo["values"] = [custom_label] + current_values
            self.template_var.set(custom_label)
        try:
            self._prev_body_size = float(self.body_size_var.get())
        except ValueError:
            pass
        self._schedule_word_preview()

    def _on_body_size_changed(self, event=None) -> None:
        try:
            new_size = float(self.body_size_var.get())
        except ValueError:
            return
        if self._prev_body_size > 0 and self._prev_body_size != new_size:
            ratio = new_size / self._prev_body_size
            self._suppress_template_callback = True
            for level in ["h1", "h2", "h3", "h4", "h5", "h6"]:
                hv = self._heading_vars[level]
                try:
                    old = float(hv["size"].get())
                    hv["size"].set(str(round(old * ratio)))
                except ValueError:
                    pass
            self._suppress_template_callback = False
        self._prev_body_size = new_size
        self._mark_custom()

    def _save_template(self) -> None:
        name = tk.simpledialog.askstring("保存模板", "请输入模板名称:", parent=self.root)
        if not name:
            return
        settings = self._collect_settings()
        actual_name = self.tm.save_template(name, settings)
        self._refresh_template_list()
        self._select_template(actual_name)
        self.status_var.set(f"模板「{actual_name}」已保存")

    def _delete_template(self) -> None:
        name = self.template_var.get()
        if not name:
            return
        t = self.tm.get_template(name)
        if t and t.get("is_builtin"):
            messagebox.showwarning("提示", "内置模板不能删除")
            return
        if not messagebox.askyesno("确认", f"确定要删除模板「{name}」吗?"):
            return
        self.tm.delete_template(name)
        self._refresh_template_list()
        self._select_template(DEFAULT_TEMPLATE["name"])
        self.status_var.set(f"模板「{name}」已删除")

    def _refresh_template_list(self) -> None:
        self.template_combo["values"] = self.tm.get_template_names()

    # ── Settings ─────────────────────────────────────────────────

    def _collect_settings(self) -> dict:
        def _f(v): return v.get()
        def _fl(v): return float(v.get())

        s: dict = {
            "body_font": _f(self.body_font_var),
            "body_size": _fl(self.body_size_var),
            "body_color": _f(self.body_color_var),
            "line_spacing_pt": _fl(self.line_spacing_var),
            "para_spacing_before": _fl(self.para_before_var),
            "para_spacing_after": _fl(self.para_after_var),
            "merge_blank_lines": bool(self.merge_var.get()),
            "page_size": self._page_size.get(),
            "margin_top_cm": 2.54, "margin_bottom_cm": 2.54,
            "margin_left_cm": 3.18, "margin_right_cm": 3.18,
        }
        for level in ["h1", "h2", "h3", "h4", "h5", "h6"]:
            hv = self._heading_vars[level]
            s[f"{level}_font"] = _f(hv["font"])
            s[f"{level}_size"] = _fl(hv["size"])
            s[f"{level}_color"] = _f(hv["color"])
            s[f"{level}_bold"] = bool(hv["bold"].get())

        current_name = self.template_var.get()
        current_t = self.tm.get_template(current_name)
        if current_t:
            for k in ("margin_top_cm", "margin_bottom_cm",
                      "margin_left_cm", "margin_right_cm"):
                s[k] = current_t[k]
        return s

    # ── File I/O ─────────────────────────────────────────────────

    def _browse_input(self) -> None:
        path = filedialog.askopenfilename(
            title="选择 Markdown 文件",
            filetypes=[("Markdown", "*.md"), ("所有文件", "*.*")])
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
            except OSError:
                pass

    # ── Word Preview ─────────────────────────────────────────────

    def _schedule_word_preview(self) -> None:
        if self._preview_job is not None:
            self.root.after_cancel(self._preview_job)
        self._preview_job = self.root.after(40, self._render_word_preview)

    def _render_word_preview(self) -> None:
        self._preview_job = None
        md_text = self._raw_markdown

        # Clear old page widgets
        for w in self._prev_pages.winfo_children():
            w.destroy()

        if not md_text.strip():
            empty = tk.Label(self._prev_pages, text="在左侧编辑器中编写或加载 Markdown 文件即可预览",
                             font=("微软雅黑", 11), fg=C_MUTED, bg="#dce4e8")
            empty.pack(pady=40)
            return

        if self.merge_var.get():
            md_text = _clean_ai_blank_lines(md_text)

        # ── Build content in temp Text to measure ──
        tmp = tk.Text(self._prev_pages, wrap="word", font=("宋体", 12))
        tmp.insert("1.0", md_text)

        s = self._collect_settings()
        ps_name = self._page_size.get()
        pw_cm, ph_cm = PAPER_SIZES.get(ps_name, PAPER_SIZES["A4"])

        # Page dimensions for preview (width in pixels, height proportional)
        usable_cm = ph_cm - s["margin_top_cm"] - s["margin_bottom_cm"]
        usable_pt = usable_cm / 2.54 * 72
        line_spacing = max(s["line_spacing_pt"], 1)
        lines_per_page = max(int(usable_pt / line_spacing), 10)

        # Page width: A4 proportion = ph/pw ≈ 1.414. Use a reasonable width
        page_ratio = pw_cm / ph_cm  # e.g. 21/29.7 ≈ 0.707
        page_width = 480  # fixed pixel width for readability
        page_height = int(page_width / page_ratio)  # taller for portrait

        # Split raw text into pages
        raw_lines = md_text.split("\n")
        pages_content = []
        current_page = []
        line_count = 0
        for raw_line in raw_lines:
            # Estimate display lines for this line based on width
            char_width = 7  # approximate pixel width per CJK char at 12pt
            chars_per_line = max(1, page_width // char_width)
            display_lines = max(1, -(-len(raw_line) // chars_per_line))  # ceil division
            if line_count + display_lines > lines_per_page and current_page:
                pages_content.append("\n".join(current_page))
                current_page = []
                line_count = 0
            current_page.append(raw_line)
            line_count += display_lines
        if current_page:
            pages_content.append("\n".join(current_page))

        if not pages_content:
            pages_content = [md_text]

        # Render HTML once
        html = markdown.markdown(md_text, extensions=["tables", "fenced_code"])
        soup = BeautifulSoup(html, "html.parser")
        body = soup.find("body")

        # Split blocks into pages
        blocks = list(body.children if body else soup.children)
        page_blocks = self._split_blocks_to_pages(blocks, lines_per_page)

        # ── Create page widgets ──
        for pg_idx, pg_blocks in enumerate(page_blocks):
            # Page frame with shadow
            page_frame = tk.Frame(self._prev_pages, bg="#dce4e8")
            page_frame.pack(pady=12)

            # Shadow (dark rectangle slightly offset)
            shadow = tk.Frame(page_frame, bg="#c8d0d8",
                              width=page_width + 4, height=page_height + 4)
            shadow.pack()
            shadow.pack_propagate(False)

            # White page
            page_bg = tk.Frame(shadow, bg="#ffffff",
                               width=page_width, height=page_height)
            page_bg.place(relx=0.5, rely=0.5, anchor="center")
            page_bg.pack_propagate(False)

            # Text widget for page content
            pt = tk.Text(page_bg, wrap="word", font=("宋体", 12),
                         relief="flat", borderwidth=0,
                         bg="#ffffff", fg=C_TEXT,
                         padx=20, pady=14,
                         state="disabled")
            pt.pack(fill="both", expand=True)

            self._configure_preview_tags_for(pt, s)

            # Render blocks for this page
            for element in pg_blocks:
                if isinstance(element, NavigableString):
                    text = str(element).strip()
                    if text:
                        pt.configure(state="normal")
                        pt.insert("end", text + "\n", "p")
                        pt.configure(state="disabled")
                    continue
                if isinstance(element, Tag):
                    self._render_block_to(pt, element)

            # Page number label
            tk.Label(page_bg, text=f"第 {pg_idx + 1} 页  |  {ps_name}  {pw_cm}×{ph_cm}cm",
                     font=("微软雅黑", 7), fg=C_MUTED, bg="#ffffff"
                     ).pack(side="bottom", pady=(0, 4))

        tmp.destroy()

    def _split_blocks_to_pages(self, blocks, lines_per_page: int) -> list:
        """Split rendered blocks into pages based on estimated line count."""
        pages = []
        current = []
        line_count = 0
        char_width = 7
        page_chars = 480 // char_width

        for blk in blocks:
            if isinstance(blk, Tag):
                text = blk.get_text()
                lines = [text[i:i+page_chars] for i in range(0, len(text), page_chars)]
                blk_lines = max(1, len(lines))
                if blk.name in ("h1", "h2"):
                    blk_lines += 2  # heading spacing
            else:
                blk_lines = 1

            if line_count + blk_lines > lines_per_page and current:
                pages.append(current)
                current = []
                line_count = 0

            current.append(blk)
            line_count += blk_lines

        if current:
            pages.append(current)
        return pages if pages else [blocks]

    def _configure_preview_tags_for(self, pt: tk.Text, s: dict) -> None:
        """Configure tags on a specific preview text widget."""
        for level in range(1, 7):
            prefix = f"h{level}"
            pt.tag_configure(f"h{level}",
                             font=(s[f"{prefix}_font"], int(s[f"{prefix}_size"]),
                                   "bold" if s[f"{prefix}_bold"] else "normal"),
                             foreground=s[f"{prefix}_color"],
                             spacing1=14 if level <= 2 else 6, spacing3=4)
        pt.tag_configure("p", font=(s["body_font"], int(s["body_size"])),
                         foreground=s["body_color"], spacing1=2, spacing3=2)
        pt.tag_configure("bold", font=(s["body_font"], int(s["body_size"]), "bold"),
                         foreground=s["body_color"])
        pt.tag_configure("italic", font=(s["body_font"], int(s["body_size"]), "italic"),
                         foreground=s["body_color"])
        pt.tag_configure("code", font=("Courier New", max(int(s["body_size"]) - 1, 9)),
                         foreground="#c0392b")
        pt.tag_configure("pre", font=("Courier New", max(int(s["body_size"]) - 1, 9)),
                         foreground=C_TEXT, background=C_INPUT_BG,
                         lmargin1=16, lmargin2=16, spacing1=4, spacing3=4)
        pt.tag_configure("quote", font=(s["body_font"], int(s["body_size"]), "italic"),
                         foreground=C_MUTED, lmargin1=24, lmargin2=24,
                         background=C_INPUT_BG)
        pt.tag_configure("li", font=(s["body_font"], int(s["body_size"])),
                         foreground=s["body_color"], lmargin1=20, lmargin2=20)
        pt.tag_configure("link", font=(s["body_font"], int(s["body_size"]), "underline"),
                         foreground=C_ACCENT)
        pt.tag_configure("hr", font=("微软雅黑", 6), foreground=C_BORDER,
                         justify="center")

    def _render_block_to(self, pt: tk.Text, element: Tag) -> None:
        """Render a block element into a specific preview text widget."""
        tag = element.name
        pt.configure(state="normal")
        if tag in ("h1", "h2", "h3", "h4", "h5", "h6"):
            pt.insert("end", element.get_text(strip=True) + "\n", tag)
        elif tag == "p":
            self._render_inline_to(pt, element, "p")
            pt.insert("end", "\n")
        elif tag in ("ul", "ol"):
            for li in element.find_all("li", recursive=False):
                pt.insert("end", "  -  ", "li")
                self._render_inline_to(pt, li, "li")
                pt.insert("end", "\n")
        elif tag == "pre":
            for line in element.get_text().rstrip("\n").split("\n"):
                pt.insert("end", "  " + line + "\n", "pre")
        elif tag == "blockquote":
            self._render_inline_to(pt, element, "quote")
            pt.insert("end", "\n")
        elif tag == "hr":
            pt.insert("end", "─" * 40 + "\n", "hr")
        elif tag == "table":
            rows = element.find_all("tr")
            for i, row in enumerate(rows):
                parts = [c.get_text(strip=True) for c in row.find_all(["td", "th"])]
                pt.insert("end", "  " + "  │  ".join(parts) + "\n",
                          "bold" if i == 0 else "p")
            pt.insert("end", "\n")
        pt.configure(state="disabled")

    def _render_inline_to(self, pt: tk.Text, element, default_tag: str) -> None:
        for child in element.children:
            if isinstance(child, NavigableString):
                pt.insert("end", str(child), default_tag)
            elif isinstance(child, Tag):
                tn = child.name
                if tn in ("strong", "b"):
                    pt.insert("end", child.get_text(), "bold")
                elif tn in ("em", "i"):
                    pt.insert("end", child.get_text(), "italic")
                elif tn == "code":
                    pt.insert("end", f"`{child.get_text()}`", "code")
                elif tn == "a":
                    pt.insert("end", child.get_text(), "link")
                elif tn == "br":
                    pt.insert("end", "\n", default_tag)
                else:
                    self._render_inline_to(pt, child, default_tag)

    # ── Save ─────────────────────────────────────────────────────

    def _save_docx(self) -> None:
        md_text = self._raw_markdown
        if not md_text.strip():
            messagebox.showwarning("提示", "没有可转换的内容")
            return

        initial_name = ""
        if self.input_path.get().strip():
            initial_name = os.path.splitext(self.input_path.get())[0] + ".docx"
        else:
            initial_name = os.path.join(os.path.expanduser("~"), "output.docx")

        out_path = filedialog.asksaveasfilename(
            title="保存 Word 文件", defaultextension=".docx",
            initialfile=os.path.basename(initial_name),
            initialdir=os.path.dirname(initial_name) if os.path.dirname(initial_name) else None,
            filetypes=[("Word 文档", "*.docx")])
        if not out_path:
            return

        template = self._collect_settings()
        merge = self.merge_var.get()

        self.convert_btn.configure(state="disabled")
        self.status_var.set("正在转换...")
        self._header_status.configure(text="● 转换中", fg="#fbbf24")

        def _run() -> None:
            try:
                convert_text(md_text, out_path, template, merge)
                self.root.after(0, lambda: self._on_save_done(out_path))
            except Exception as e:
                self.root.after(0, lambda: self._on_save_error(str(e)))

        threading.Thread(target=_run, daemon=True).start()

    def _on_save_done(self, out: str) -> None:
        self.convert_btn.configure(state="normal")
        self.status_var.set(f"转换完成 → {out}")
        self._header_status.configure(text="● 就绪", fg="#a8f0c8")
        if messagebox.askyesno("完成", f"转换成功!\n\n文件: {out}\n\n是否打开文件?"):
            os.startfile(out)

    def _on_save_error(self, err: str) -> None:
        self.convert_btn.configure(state="normal")
        self.status_var.set("转换失败")
        self._header_status.configure(text="● 错误", fg="#f87171")
        messagebox.showerror("错误", f"转换失败:\n{err}")


def main() -> None:
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
