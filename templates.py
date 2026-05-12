"""Template management for Markdown to Word converter."""

import json
import os
from typing import Any

# ---- Built-in template definitions ----

BUILTIN_TEMPLATES = [
    {
        "name": "公务员",
        "is_builtin": True,
        "body_font": "仿宋",
        "body_size": 16,
        "body_color": "#000000",
        "h1_font": "方正小标宋简体",
        "h1_size": 22,
        "h1_bold": False,
        "h1_color": "#000000",
        "h2_font": "黑体",
        "h2_size": 16,
        "h2_bold": False,
        "h2_color": "#000000",
        "h3_font": "楷体",
        "h3_size": 16,
        "h3_bold": False,
        "h3_color": "#000000",
        "h4_font": "仿宋",
        "h4_size": 16,
        "h4_bold": True,
        "h4_color": "#000000",
        "h5_font": "仿宋",
        "h5_size": 16,
        "h5_bold": True,
        "h5_color": "#000000",
        "h6_font": "仿宋",
        "h6_size": 16,
        "h6_bold": True,
        "h6_color": "#000000",
        "line_spacing_pt": 28,
        "para_spacing_before": 0,
        "para_spacing_after": 0,
        "margin_top_cm": 3.7,
        "margin_bottom_cm": 3.5,
        "margin_left_cm": 2.8,
        "margin_right_cm": 2.6,
        "merge_blank_lines": True,
    },
    {
        "name": "学生紧凑",
        "is_builtin": True,
        "body_font": "宋体",
        "body_size": 12,
        "body_color": "#000000",
        "h1_font": "黑体",
        "h1_size": 16,
        "h1_bold": True,
        "h1_color": "#000000",
        "h2_font": "黑体",
        "h2_size": 14,
        "h2_bold": True,
        "h2_color": "#000000",
        "h3_font": "宋体",
        "h3_size": 12,
        "h3_bold": True,
        "h3_color": "#000000",
        "h4_font": "宋体",
        "h4_size": 12,
        "h4_bold": True,
        "h4_color": "#000000",
        "h5_font": "宋体",
        "h5_size": 12,
        "h5_bold": True,
        "h5_color": "#000000",
        "h6_font": "宋体",
        "h6_size": 12,
        "h6_bold": True,
        "h6_color": "#000000",
        "line_spacing_pt": 22,
        "para_spacing_before": 0,
        "para_spacing_after": 6,
        "margin_top_cm": 2.54,
        "margin_bottom_cm": 2.54,
        "margin_left_cm": 3.18,
        "margin_right_cm": 3.18,
        "merge_blank_lines": True,
    },
]

# Common Chinese + English fonts for dropdown
FONT_LIST = [
    "宋体", "黑体", "仿宋", "楷体", "微软雅黑", "方正小标宋简体",
    "Arial", "Times New Roman", "Calibri", "Courier New",
]

# Common font sizes in pt
SIZE_LIST = [9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 42, 48]

# Preset color swatches
COLOR_PRESETS = [
    "#000000", "#333333", "#666666", "#999999",
    "#FF0000", "#CC0000", "#990000",
    "#0000FF", "#0000CC", "#000099",
    "#008000", "#006600", "#004400",
]


class TemplateManager:
    """Manages built-in and user templates with JSON persistence."""

    def __init__(self, data_dir: str):
        self.data_dir = data_dir
        self.json_path = os.path.join(data_dir, "templates.json")
        self._templates: list[dict[str, Any]] = []
        self._load()

    def _load(self) -> None:
        """Load built-in + user templates into memory."""
        self._templates = [dict(t) for t in BUILTIN_TEMPLATES]
        if os.path.exists(self.json_path):
            try:
                with open(self.json_path, "r", encoding="utf-8") as f:
                    user_templates = json.load(f)
                for t in user_templates:
                    t["is_builtin"] = False
                self._templates.extend(user_templates)
            except (json.JSONDecodeError, OSError):
                pass

    def _save_user_templates(self) -> None:
        """Persist only user (non-builtin) templates to JSON."""
        user_templates = [t for t in self._templates if not t.get("is_builtin")]
        os.makedirs(self.data_dir, exist_ok=True)
        with open(self.json_path, "w", encoding="utf-8") as f:
            json.dump(user_templates, f, ensure_ascii=False, indent=2)

    def list_templates(self) -> list[dict[str, Any]]:
        return list(self._templates)

    def get_template(self, name: str) -> dict[str, Any] | None:
        for t in self._templates:
            if t["name"] == name:
                return dict(t)
        return None

    def get_template_names(self) -> list[str]:
        return [t["name"] for t in self._templates]

    def save_template(self, name: str, settings: dict[str, Any]) -> str:
        """Save a user template. Returns the actual saved name."""
        # Check if name conflicts with a builtin
        for bt in BUILTIN_TEMPLATES:
            if bt["name"] == name:
                name = name + "(自定义)"
                break

        settings["name"] = name
        settings["is_builtin"] = False

        # Update existing or append
        for i, t in enumerate(self._templates):
            if t["name"] == name and not t.get("is_builtin"):
                self._templates[i] = settings
                self._save_user_templates()
                return name

        self._templates.append(settings)
        self._save_user_templates()
        return name

    def delete_template(self, name: str) -> bool:
        """Delete a user template. Returns False for builtins."""
        for i, t in enumerate(self._templates):
            if t["name"] == name:
                if t.get("is_builtin"):
                    return False
                del self._templates[i]
                self._save_user_templates()
                return True
        return False

    @staticmethod
    def template_to_settings(t: dict[str, Any]) -> dict[str, Any]:
        """Extract GUI-settable fields from a template dict."""
        return {
            "body_font": t["body_font"],
            "body_size": t["body_size"],
            "body_color": t["body_color"],
            "h1_font": t["h1_font"], "h1_size": t["h1_size"],
            "h1_bold": t["h1_bold"], "h1_color": t["h1_color"],
            "h2_font": t["h2_font"], "h2_size": t["h2_size"],
            "h2_bold": t["h2_bold"], "h2_color": t["h2_color"],
            "h3_font": t["h3_font"], "h3_size": t["h3_size"],
            "h3_bold": t["h3_bold"], "h3_color": t["h3_color"],
            "h4_font": t["h4_font"], "h4_size": t["h4_size"],
            "h4_bold": t["h4_bold"], "h4_color": t["h4_color"],
            "h5_font": t["h5_font"], "h5_size": t["h5_size"],
            "h5_bold": t["h5_bold"], "h5_color": t["h5_color"],
            "h6_font": t["h6_font"], "h6_size": t["h6_size"],
            "h6_bold": t["h6_bold"], "h6_color": t["h6_color"],
            "line_spacing_pt": t["line_spacing_pt"],
            "para_spacing_before": t["para_spacing_before"],
            "para_spacing_after": t["para_spacing_after"],
            "margin_top_cm": t["margin_top_cm"],
            "margin_bottom_cm": t["margin_bottom_cm"],
            "margin_left_cm": t["margin_left_cm"],
            "margin_right_cm": t["margin_right_cm"],
            "merge_blank_lines": t["merge_blank_lines"],
        }
