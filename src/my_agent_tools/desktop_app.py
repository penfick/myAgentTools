from __future__ import annotations

import json
import os
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from my_agent_tools.app_settings import (
    PROVIDER_PRESETS,
    DesktopPlannerSettings,
    load_desktop_settings,
    save_desktop_settings,
)
from my_agent_tools.ai_models import OutlinePlan, PlannerSettings
from my_agent_tools.openai_planner import OpenAIPlanner, PlannerError
from my_agent_tools.specs import BlockSpec, DeckSpec, SlideSpec
from my_agent_tools.tools.ppt import generate_ppt


APP_TITLE = "My Agent Tools 工作台"
APP_BG = "#F3F0E8"
SURFACE_BG = "#FFFDF8"
SURFACE_ALT_BG = "#F8F4EC"
HEADER_BG = "#133A5E"
HEADER_MUTED = "#D8E2ED"
TEXT_MAIN = "#1C2733"
TEXT_MUTED = "#66758A"
ACCENT = "#C96A2B"
ACCENT_SOFT = "#F7E2D3"
ACCENT_DEEP = "#A94F18"
SUCCESS = "#0F8B6D"
ERROR = "#B42318"
WARNING = "#B54708"
BORDER = "#D9D1C4"
CARD_SHADOW = "#E6DED2"
CANVAS_BG = "#F4F1EA"

FONT_UI = ("Microsoft YaHei UI", 10)
FONT_UI_BOLD = ("Microsoft YaHei UI", 10, "bold")
FONT_TITLE = ("Microsoft YaHei UI", 21, "bold")
FONT_SECTION = ("Microsoft YaHei UI", 14, "bold")
FONT_MONO = ("Consolas", 10)

SLIDE_KIND_LABELS = {
    "title": "封面",
    "section": "章节",
    "content": "内容",
    "two_column": "双栏",
    "table": "表格",
    "image": "图片",
    "metrics": "指标",
    "chart": "图表",
    "closing": "结尾",
}

AI_STAGE_ORDER = ["requirement", "outline", "clarification", "ready_spec", "spec_ready"]
AI_STAGE_LABELS = {
    "requirement": "1. 输入需求",
    "outline": "2. 生成大纲",
    "clarification": "3. 等待补充",
    "ready_spec": "4. 确认大纲",
    "spec_ready": "5. 生成规格",
}


class DesktopApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("1520x940")
        self.root.minsize(1260, 780)
        self.root.configure(bg=APP_BG)

        self.project_root = Path(__file__).resolve().parents[2]
        self.sample_spec_path = self.project_root / "examples" / "sample_deck.json"
        self.output_dir = self.project_root / "out"
        self.config_dir = self.project_root / "config"
        self.desktop_settings_path = self.config_dir / "desktop_settings.json"
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.desktop_settings = self._load_desktop_settings()

        self.current_spec_path: Path | None = None
        self.current_spec: DeckSpec | None = None
        self.current_outline: OutlinePlan | None = None
        self.last_output_path: Path | None = None
        self._refresh_job = None

        self.status_var = tk.StringVar(value="就绪")
        self.meta_var = tk.StringVar(value="尚未加载规格文件")
        self.summary_var = tk.StringVar(value="先加载一个规格文件，或者先在 AI 规划页生成大纲。")
        self.ai_status_var = tk.StringVar(value="AI 工作流：先输入需求，先确认大纲，再生成规格。")
        self.ai_stage_var = tk.StringVar(value="requirement")
        self.provider_preset_var = tk.StringVar(value=self.desktop_settings.provider_preset)
        self.model_var = tk.StringVar(value=self.desktop_settings.model or os.getenv("OPENAI_MODEL", "gpt-5-mini"))
        self.base_url_var = tk.StringVar(value=self.desktop_settings.base_url or os.getenv("OPENAI_BASE_URL", ""))
        self.api_key_var = tk.StringVar(value=self.desktop_settings.api_key or os.getenv("OPENAI_API_KEY", ""))
        self.remember_api_key_var = tk.BooleanVar(value=self.desktop_settings.remember_api_key)

        self.stage_labels: dict[str, tk.Label] = {}

        self._configure_style()
        self._build_layout()
        self._apply_provider_preset(self.provider_preset_var.get(), overwrite_empty_only=True)
        self._set_questions_text("当前没有待确认问题。")
        self._refresh_ai_controls()
        self._render_ai_stage()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.load_sample_spec()

    def _load_desktop_settings(self) -> DesktopPlannerSettings:
        try:
            return load_desktop_settings(self.desktop_settings_path)
        except Exception:
            return DesktopPlannerSettings()

    def _save_desktop_settings(self) -> None:
        settings = DesktopPlannerSettings(
            provider_preset=self.provider_preset_var.get().strip() or "custom",
            model=self.model_var.get().strip() or "gpt-5-mini",
            base_url=self.base_url_var.get().strip(),
            api_key=self.api_key_var.get().strip() if self.remember_api_key_var.get() else "",
            remember_api_key=self.remember_api_key_var.get(),
        )
        save_desktop_settings(self.desktop_settings_path, settings)

    def _on_close(self) -> None:
        self._save_desktop_settings()
        self.root.destroy()

    def _configure_style(self) -> None:
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Root.TFrame", background=APP_BG)
        style.configure("Panel.TFrame", background=SURFACE_BG)
        style.configure("Muted.TFrame", background=SURFACE_ALT_BG)
        style.configure("Status.TLabel", background=APP_BG, foreground=TEXT_MUTED, font=FONT_UI)

        style.configure("Primary.TButton", font=FONT_UI_BOLD, padding=(14, 9), borderwidth=0)
        style.map(
            "Primary.TButton",
            background=[("active", ACCENT_DEEP), ("!disabled", ACCENT)],
            foreground=[("!disabled", "#FFFFFF")],
        )

        style.configure("Secondary.TButton", font=FONT_UI, padding=(12, 8), borderwidth=1)
        style.map(
            "Secondary.TButton",
            background=[("active", "#F2E5D8"), ("!disabled", SURFACE_BG)],
            foreground=[("!disabled", TEXT_MAIN)],
            bordercolor=[("!disabled", BORDER)],
        )

        style.configure("Plain.TNotebook", background=SURFACE_BG, borderwidth=0)
        style.configure(
            "Plain.TNotebook.Tab",
            font=FONT_UI_BOLD,
            padding=(18, 10),
            background="#ECE4D7",
            foreground=TEXT_MUTED,
        )
        style.map(
            "Plain.TNotebook.Tab",
            background=[("selected", SURFACE_BG)],
            foreground=[("selected", TEXT_MAIN)],
        )

    def _build_layout(self) -> None:
        root_frame = ttk.Frame(self.root, style="Root.TFrame", padding=18)
        root_frame.pack(fill="both", expand=True)

        self._build_header(root_frame)

        main_pane = ttk.Panedwindow(root_frame, orient=tk.HORIZONTAL)
        main_pane.pack(fill="both", expand=True, pady=(16, 10))

        left_panel = ttk.Frame(main_pane, style="Panel.TFrame", padding=16)
        right_panel = ttk.Frame(main_pane, style="Panel.TFrame", padding=16)
        main_pane.add(left_panel, weight=27)
        main_pane.add(right_panel, weight=73)

        self._build_left_panel(left_panel)
        self._build_right_panel(right_panel)

        status_row = ttk.Frame(root_frame, style="Root.TFrame")
        status_row.pack(fill="x")
        ttk.Label(status_row, textvariable=self.status_var, style="Status.TLabel").pack(side="left")

    def _build_header(self, parent) -> None:
        header = tk.Frame(parent, bg=HEADER_BG, padx=22, pady=18, highlightthickness=0)
        header.pack(fill="x")

        left = tk.Frame(header, bg=HEADER_BG)
        left.pack(side="left", fill="both", expand=True)

        tk.Label(left, text=APP_TITLE, bg=HEADER_BG, fg="#FFFFFF", font=FONT_TITLE).pack(anchor="w")
        tk.Label(
            left,
            text="先规划大纲，再确认规格，最后导出 PPT。当前重点仍是 PPT 路线。",
            bg=HEADER_BG,
            fg=HEADER_MUTED,
            font=FONT_UI,
        ).pack(anchor="w", pady=(6, 0))

        badge_row = tk.Frame(left, bg=HEADER_BG)
        badge_row.pack(anchor="w", pady=(12, 0))
        self._make_badge(badge_row, "本地桌面 GUI", "#28577F", "#EAF3FB").pack(side="left")
        self._make_badge(badge_row, "确认式 AI 工作流", "#28577F", "#EAF3FB").pack(side="left", padx=(8, 0))
        self._make_badge(badge_row, "支持第三方 API URL / Key", "#6D3A17", "#FCEFE6").pack(side="left", padx=(8, 0))

        actions = tk.Frame(header, bg=HEADER_BG)
        actions.pack(side="right", anchor="ne")

        button_specs = [
            ("加载示例", self.load_sample_spec, "Primary.TButton"),
            ("打开规格", self.open_spec, "Secondary.TButton"),
            ("另存规格", self.save_spec_as, "Secondary.TButton"),
            ("校验", self.validate_spec, "Secondary.TButton"),
            ("生成 PPT", self.render_ppt, "Primary.TButton"),
            ("打开输出目录", self.open_output_folder, "Secondary.TButton"),
        ]
        for index, (label, callback, style_name) in enumerate(button_specs):
            ttk.Button(actions, text=label, command=callback, style=style_name).grid(
                row=index // 3,
                column=index % 3,
                padx=(0, 10) if index % 3 != 2 else 0,
                pady=(0, 10) if index < 3 else 0,
                sticky="ew",
            )
        for column in range(3):
            actions.grid_columnconfigure(column, weight=1)

    def _make_badge(self, parent, text: str, bg: str, fg: str) -> tk.Frame:
        frame = tk.Frame(parent, bg=bg, padx=10, pady=4)
        tk.Label(frame, text=text, bg=bg, fg=fg, font=("Microsoft YaHei UI", 9, "bold")).pack()
        return frame

    def _build_left_panel(self, parent) -> None:
        tk.Label(parent, text="当前演示概览", bg=SURFACE_BG, fg=TEXT_MAIN, font=FONT_SECTION).pack(anchor="w")
        tk.Label(parent, textvariable=self.meta_var, bg=SURFACE_BG, fg=TEXT_MUTED, font=FONT_UI).pack(anchor="w", pady=(6, 14))

        summary_card = tk.Frame(parent, bg=ACCENT_SOFT, padx=14, pady=14, highlightthickness=1, highlightbackground=BORDER)
        summary_card.pack(fill="x")
        tk.Label(summary_card, text="摘要", bg=ACCENT_SOFT, fg=ACCENT_DEEP, font=("Microsoft YaHei UI", 11, "bold")).pack(anchor="w")
        tk.Label(
            summary_card,
            textvariable=self.summary_var,
            bg=ACCENT_SOFT,
            fg=TEXT_MAIN,
            justify="left",
            wraplength=320,
            font=FONT_UI,
            pady=8,
        ).pack(anchor="w")

        hint_card = tk.Frame(parent, bg=SURFACE_ALT_BG, padx=14, pady=12, highlightthickness=1, highlightbackground=BORDER)
        hint_card.pack(fill="x", pady=(14, 14))
        tk.Label(hint_card, text="当前流程", bg=SURFACE_ALT_BG, fg=TEXT_MAIN, font=("Microsoft YaHei UI", 11, "bold")).pack(anchor="w")
        tk.Label(
            hint_card,
            text="1. 输入需求\n2. 先出大纲\n3. 用户补充并确认\n4. 再生成规格与 PPT",
            bg=SURFACE_ALT_BG,
            fg=TEXT_MUTED,
            justify="left",
            font=FONT_UI,
            pady=6,
        ).pack(anchor="w")

        tk.Label(parent, text="页面列表", bg=SURFACE_BG, fg=TEXT_MAIN, font=("Microsoft YaHei UI", 11, "bold")).pack(anchor="w", pady=(0, 8))

        list_shell = tk.Frame(parent, bg=CARD_SHADOW, padx=1, pady=1)
        list_shell.pack(fill="both", expand=True)
        list_frame = tk.Frame(list_shell, bg=SURFACE_BG)
        list_frame.pack(fill="both", expand=True)

        self.slide_list = tk.Listbox(
            list_frame,
            borderwidth=0,
            activestyle="none",
            bg=SURFACE_BG,
            fg=TEXT_MAIN,
            selectbackground=ACCENT,
            selectforeground="#FFFFFF",
            highlightthickness=0,
            font=FONT_UI,
        )
        self.slide_list.pack(side="left", fill="both", expand=True)
        self.slide_list.bind("<<ListboxSelect>>", self._on_slide_selected)

        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.slide_list.yview)
        scrollbar.pack(side="right", fill="y")
        self.slide_list.configure(yscrollcommand=scrollbar.set)

    def _build_right_panel(self, parent) -> None:
        notebook = ttk.Notebook(parent, style="Plain.TNotebook")
        notebook.pack(fill="both", expand=True)

        ai_tab = ttk.Frame(notebook, style="Muted.TFrame", padding=16)
        preview_tab = ttk.Frame(notebook, style="Muted.TFrame", padding=16)
        json_tab = ttk.Frame(notebook, style="Muted.TFrame", padding=16)

        notebook.add(ai_tab, text="AI 规划")
        notebook.add(preview_tab, text="页面预览")
        notebook.add(json_tab, text="规格 JSON")

        self._build_ai_tab(ai_tab)
        self._build_preview_tab(preview_tab)
        self._build_json_tab(json_tab)

    def _build_ai_tab(self, parent) -> None:
        config_card = tk.Frame(parent, bg=SURFACE_BG, padx=16, pady=16, highlightthickness=1, highlightbackground=BORDER)
        config_card.pack(fill="x")
        preset_values = [self._provider_display_value(key) for key in PROVIDER_PRESETS]

        tk.Label(config_card, text="模型接入设置", bg=SURFACE_BG, fg=TEXT_MAIN, font=FONT_SECTION).grid(row=0, column=0, sticky="w")
        tk.Label(
            config_card,
            text="支持官方 OpenAI，也支持 OpenAI 兼容的第三方网关。留空时优先读取环境变量。",
            bg=SURFACE_BG,
            fg=TEXT_MUTED,
            font=FONT_UI,
        ).grid(row=1, column=0, columnspan=6, sticky="w", pady=(6, 12))

        self._form_label(config_card, "模型").grid(row=2, column=0, sticky="w")
        tk.Entry(config_card, textvariable=self.model_var, width=18, font=FONT_MONO, relief="flat").grid(row=2, column=1, sticky="we", padx=(8, 12))
        self._form_label(config_card, "API Base URL").grid(row=2, column=2, sticky="w")
        tk.Entry(config_card, textvariable=self.base_url_var, width=30, font=FONT_MONO, relief="flat").grid(row=2, column=3, sticky="we", padx=(8, 12))
        self._form_label(config_card, "API Key").grid(row=2, column=4, sticky="w")
        tk.Entry(config_card, textvariable=self.api_key_var, width=30, show="*", font=FONT_MONO, relief="flat").grid(row=2, column=5, sticky="we", padx=(8, 0))

        tk.Label(
            config_card,
            text="示例：官方可用 `https://api.openai.com/v1`；第三方则填写其兼容地址。",
            bg=SURFACE_BG,
            fg=TEXT_MUTED,
            font=("Microsoft YaHei UI", 9),
        ).grid(row=3, column=0, columnspan=6, sticky="w", pady=(10, 0))

        self._form_label(config_card, "接口预设").grid(row=4, column=0, sticky="w", pady=(12, 0))
        self.provider_combo = ttk.Combobox(config_card, state="readonly", values=preset_values, font=FONT_UI)
        self.provider_combo.grid(row=4, column=1, columnspan=2, sticky="we", padx=(8, 12), pady=(12, 0))
        self.provider_combo.bind("<<ComboboxSelected>>", self._on_provider_preset_selected)
        self.provider_combo.set(self._provider_display_value(self.provider_preset_var.get()))

        tk.Checkbutton(
            config_card,
            text="记住 API Key（本地明文保存）",
            variable=self.remember_api_key_var,
            bg=SURFACE_BG,
            fg=TEXT_MUTED,
            activebackground=SURFACE_BG,
            activeforeground=TEXT_MAIN,
            selectcolor=SURFACE_BG,
            font=FONT_UI,
        ).grid(row=4, column=3, columnspan=2, sticky="w", pady=(12, 0))
        ttk.Button(config_card, text="保存模型配置", command=self._save_desktop_settings, style="Secondary.TButton").grid(row=4, column=5, sticky="e", pady=(12, 0))

        config_card.grid_columnconfigure(1, weight=1)
        config_card.grid_columnconfigure(3, weight=2)
        config_card.grid_columnconfigure(5, weight=2)

        stage_card = tk.Frame(parent, bg=SURFACE_ALT_BG, padx=16, pady=12, highlightthickness=1, highlightbackground=BORDER)
        stage_card.pack(fill="x", pady=(14, 14))
        tk.Label(stage_card, text="确认式生成流程", bg=SURFACE_ALT_BG, fg=TEXT_MAIN, font=("Microsoft YaHei UI", 11, "bold")).pack(anchor="w")
        tk.Label(stage_card, textvariable=self.ai_status_var, bg=SURFACE_ALT_BG, fg=TEXT_MUTED, font=FONT_UI).pack(anchor="w", pady=(4, 10))

        stage_row = tk.Frame(stage_card, bg=SURFACE_ALT_BG)
        stage_row.pack(fill="x")
        for index, stage in enumerate(AI_STAGE_ORDER):
            card = tk.Frame(stage_row, bg="#EFE7DA", padx=12, pady=10)
            card.pack(side="left", fill="x", expand=True, padx=(0, 10) if index < len(AI_STAGE_ORDER) - 1 else 0)
            label = tk.Label(card, text=AI_STAGE_LABELS[stage], bg="#EFE7DA", fg=TEXT_MUTED, font=FONT_UI_BOLD)
            label.pack(anchor="w")
            self.stage_labels[stage] = label

        editors = tk.Frame(parent, bg=SURFACE_ALT_BG)
        editors.pack(fill="both", expand=True, pady=(0, 14))
        editors.grid_columnconfigure(0, weight=1)
        editors.grid_columnconfigure(1, weight=1)
        editors.grid_rowconfigure(1, weight=1)
        editors.grid_rowconfigure(4, weight=2)

        tk.Label(editors, text="用户需求", bg=SURFACE_ALT_BG, fg=TEXT_MAIN, font=("Microsoft YaHei UI", 11, "bold")).grid(row=0, column=0, sticky="w")
        tk.Label(editors, text="补充说明 / 用户修正", bg=SURFACE_ALT_BG, fg=TEXT_MAIN, font=("Microsoft YaHei UI", 11, "bold")).grid(row=0, column=1, sticky="w", padx=(14, 0))

        self.requirement_editor = self._build_text_editor(editors, row=1, column=0, padx=(0, 0), pady=(8, 12))
        self.feedback_editor = self._build_text_editor(editors, row=1, column=1, padx=(14, 0), pady=(8, 12))
        self.requirement_editor.bind("<KeyRelease>", self._on_ai_input_changed)
        self.feedback_editor.bind("<KeyRelease>", self._on_ai_input_changed)

        button_row = tk.Frame(editors, bg=SURFACE_ALT_BG)
        button_row.grid(row=2, column=0, columnspan=2, sticky="w", pady=(0, 10))
        self.generate_outline_button = self._solid_button(button_row, "生成大纲", ACCENT, self.generate_outline)
        self.generate_outline_button.pack(side="left")
        self.regenerate_outline_button = self._solid_button(button_row, "结合补充重生成大纲", "#1F6A5A", self.regenerate_outline)
        self.regenerate_outline_button.pack(side="left", padx=(10, 0))
        self.confirm_outline_button = self._solid_button(button_row, "确认大纲并生成规格", "#8E4B15", self.confirm_outline_generate_spec)
        self.confirm_outline_button.pack(side="left", padx=(10, 0))

        tk.Label(editors, text="可编辑大纲 JSON", bg=SURFACE_ALT_BG, fg=TEXT_MAIN, font=("Microsoft YaHei UI", 11, "bold")).grid(row=3, column=0, columnspan=2, sticky="w")
        outline_shell = tk.Frame(editors, bg=CARD_SHADOW, padx=1, pady=1)
        outline_shell.grid(row=4, column=0, columnspan=2, sticky="nsew")
        outline_frame = tk.Frame(outline_shell, bg=SURFACE_BG)
        outline_frame.pack(fill="both", expand=True)

        self.outline_editor = tk.Text(
            outline_frame,
            wrap="none",
            bg=SURFACE_BG,
            fg=TEXT_MAIN,
            insertbackground=TEXT_MAIN,
            borderwidth=0,
            padx=14,
            pady=12,
            font=FONT_MONO,
        )
        self.outline_editor.pack(side="left", fill="both", expand=True)
        self.outline_editor.bind("<KeyRelease>", self._on_outline_changed)
        y_scroll = ttk.Scrollbar(outline_frame, orient="vertical", command=self.outline_editor.yview)
        y_scroll.pack(side="right", fill="y")
        self.outline_editor.configure(yscrollcommand=y_scroll.set)

        question_card = tk.Frame(parent, bg=SURFACE_BG, padx=16, pady=14, highlightthickness=1, highlightbackground=BORDER)
        question_card.pack(fill="x")
        tk.Label(question_card, text="待确认问题", bg=SURFACE_BG, fg=TEXT_MAIN, font=("Microsoft YaHei UI", 11, "bold")).pack(anchor="w")
        tk.Label(
            question_card,
            text="当模型认为信息不足时，会把追问列在这里。你可以一键生成答复模板，再回填到“补充说明 / 用户修正”。",
            bg=SURFACE_BG,
            fg=TEXT_MUTED,
            font=("Microsoft YaHei UI", 9),
        ).pack(anchor="w", pady=(4, 10))

        question_inner = tk.Frame(question_card, bg=CARD_SHADOW, padx=1, pady=1)
        question_inner.pack(fill="x")
        self.questions_view = tk.Text(
            question_inner,
            height=4,
            wrap="word",
            bg=SURFACE_BG,
            fg=TEXT_MAIN,
            borderwidth=0,
            padx=12,
            pady=10,
            font=FONT_UI,
            state="disabled",
        )
        self.questions_view.pack(fill="x")

        action_row = tk.Frame(question_card, bg=SURFACE_BG)
        action_row.pack(fill="x", pady=(10, 0))
        self.template_button = self._solid_button(action_row, "写入答复模板", "#375A7F", self._insert_clarification_template)
        self.template_button.pack(side="left")

    def _build_text_editor(self, parent, row: int, column: int, padx: tuple[int, int], pady: tuple[int, int]) -> tk.Text:
        shell = tk.Frame(parent, bg=CARD_SHADOW, padx=1, pady=1)
        shell.grid(row=row, column=column, sticky="nsew", padx=padx, pady=pady)
        editor = tk.Text(
            shell,
            height=8,
            wrap="word",
            bg=SURFACE_BG,
            fg=TEXT_MAIN,
            borderwidth=0,
            padx=12,
            pady=12,
            font=FONT_UI,
            insertbackground=TEXT_MAIN,
        )
        editor.pack(fill="both", expand=True)
        return editor

    def _solid_button(self, parent, text: str, bg: str, command) -> tk.Button:
        return tk.Button(
            parent,
            text=text,
            command=command,
            bg=bg,
            fg="#FFFFFF",
            activebackground=bg,
            activeforeground="#FFFFFF",
            relief="flat",
            padx=16,
            pady=8,
            font=FONT_UI,
            cursor="hand2",
        )

    def _form_label(self, parent, text: str) -> tk.Label:
        return tk.Label(parent, text=text, bg=SURFACE_BG, fg=TEXT_MUTED, font=FONT_UI)

    def _provider_display_value(self, key: str) -> str:
        preset = PROVIDER_PRESETS.get(key, PROVIDER_PRESETS["custom"])
        return f"{key} | {preset['label']}"

    def _on_provider_preset_selected(self, _event=None) -> None:
        raw_value = self.provider_combo.get().strip()
        preset_key = raw_value.split("|", 1)[0].strip() if raw_value else "custom"
        self.provider_preset_var.set(preset_key)
        self._apply_provider_preset(preset_key, overwrite_empty_only=False)
        self._set_status(f"已切换接口预设：{PROVIDER_PRESETS.get(preset_key, PROVIDER_PRESETS['custom'])['label']}")

    def _apply_provider_preset(self, preset_key: str, overwrite_empty_only: bool) -> None:
        preset = PROVIDER_PRESETS.get(preset_key)
        if not preset:
            return
        if preset.get("model") and (not overwrite_empty_only or not self.model_var.get().strip()):
            self.model_var.set(preset["model"])
        if not overwrite_empty_only or not self.base_url_var.get().strip():
            self.base_url_var.set(preset.get("base_url", ""))

    def _set_ai_stage(self, stage: str) -> None:
        self.ai_stage_var.set(stage)
        self._render_ai_stage()

    def _render_ai_stage(self) -> None:
        current_stage = self.ai_stage_var.get()
        current_index = AI_STAGE_ORDER.index(current_stage) if current_stage in AI_STAGE_ORDER else 0
        for index, stage in enumerate(AI_STAGE_ORDER):
            label = self.stage_labels.get(stage)
            if label is None:
                continue
            bg = "#EFE7DA"
            fg = TEXT_MUTED
            if index < current_index:
                bg = "#DDF3EA"
                fg = SUCCESS
            elif index == current_index:
                bg = ACCENT_SOFT
                fg = ACCENT_DEEP
            label.master.configure(bg=bg)
            label.configure(bg=bg, fg=fg)

    def _set_questions_text(self, text: str) -> None:
        self.questions_view.configure(state="normal")
        self.questions_view.delete("1.0", tk.END)
        self.questions_view.insert("1.0", text)
        self.questions_view.configure(state="disabled")

    def _insert_clarification_template(self) -> None:
        outline = self.current_outline
        if outline is None or not outline.clarification_questions:
            messagebox.showinfo("没有待确认问题", "当前没有需要补充回答的问题。")
            return

        existing = self.feedback_editor.get("1.0", tk.END).strip()
        template = "\n\n".join(
            f"问题 {index + 1}: {item.question}\n回答："
            for index, item in enumerate(outline.clarification_questions)
        )
        merged = f"{existing}\n\n{template}".strip() if existing else template
        self.feedback_editor.delete("1.0", tk.END)
        self.feedback_editor.insert("1.0", merged)
        self._set_status("已将待确认问题写入补充说明模板")
        self._refresh_ai_controls()

    def _on_ai_input_changed(self, _event=None) -> None:
        if self.current_outline is None:
            self._set_ai_stage("requirement")
        self._refresh_ai_controls()

    def _on_outline_changed(self, _event=None) -> None:
        self._refresh_ai_controls()

    def _refresh_ai_controls(self) -> None:
        requirement = self.requirement_editor.get("1.0", tk.END).strip()
        feedback = self.feedback_editor.get("1.0", tk.END).strip()
        has_requirement = bool(requirement)
        has_feedback = bool(feedback)

        self.generate_outline_button.configure(state="normal" if has_requirement else "disabled")

        can_regenerate = has_requirement and (has_feedback or (self.current_outline is not None and self.current_outline.needs_clarification))
        self.regenerate_outline_button.configure(state="normal" if can_regenerate else "disabled")

        outline = self._parse_outline(show_errors=False)
        can_confirm = has_requirement and outline is not None and not outline.needs_clarification
        self.confirm_outline_button.configure(state="normal" if can_confirm else "disabled")
        self.template_button.configure(state="normal" if self.current_outline and self.current_outline.clarification_questions else "disabled")

    def _build_preview_tab(self, parent) -> None:
        header = tk.Frame(parent, bg=SURFACE_ALT_BG)
        header.pack(fill="x")

        self.preview_title_label = tk.Label(header, text="预览", bg=SURFACE_ALT_BG, fg=TEXT_MAIN, font=FONT_SECTION)
        self.preview_title_label.pack(anchor="w")

        self.preview_meta_label = tk.Label(header, text="尚未选择页面", bg=SURFACE_ALT_BG, fg=TEXT_MUTED, font=FONT_UI)
        self.preview_meta_label.pack(anchor="w", pady=(6, 0))

        preview_shell = tk.Frame(parent, bg=CARD_SHADOW, padx=1, pady=1)
        preview_shell.pack(fill="both", expand=True, pady=(14, 0))
        preview_frame = tk.Frame(preview_shell, bg=CANVAS_BG)
        preview_frame.pack(fill="both", expand=True)

        self.preview_canvas = tk.Canvas(preview_frame, bg=CANVAS_BG, highlightthickness=0, relief="flat")
        self.preview_canvas.pack(fill="both", expand=True)
        self.preview_canvas.bind("<Configure>", self._on_preview_resize)

    def _build_json_tab(self, parent) -> None:
        top = tk.Frame(parent, bg=SURFACE_ALT_BG)
        top.pack(fill="x", pady=(0, 10))

        text_group = tk.Frame(top, bg=SURFACE_ALT_BG)
        text_group.pack(side="left")
        tk.Label(text_group, text="可编辑规格 JSON", bg=SURFACE_ALT_BG, fg=TEXT_MAIN, font=FONT_SECTION).pack(anchor="w")
        tk.Label(text_group, text="这里是最终可修改的 DeckSpec。修改后左侧概览和预览会自动刷新。", bg=SURFACE_ALT_BG, fg=TEXT_MUTED, font=FONT_UI).pack(anchor="w", pady=(4, 0))

        ttk.Button(top, text="格式化 JSON", command=self.format_json, style="Primary.TButton").pack(side="right")

        editor_shell = tk.Frame(parent, bg=CARD_SHADOW, padx=1, pady=1)
        editor_shell.pack(fill="both", expand=True)
        editor_frame = tk.Frame(editor_shell, bg=SURFACE_BG)
        editor_frame.pack(fill="both", expand=True)

        self.json_editor = tk.Text(
            editor_frame,
            wrap="none",
            bg=SURFACE_BG,
            fg=TEXT_MAIN,
            insertbackground=TEXT_MAIN,
            borderwidth=0,
            padx=14,
            pady=12,
            font=FONT_MONO,
        )
        self.json_editor.pack(side="left", fill="both", expand=True)
        self.json_editor.bind("<KeyRelease>", self._schedule_preview_refresh)

        y_scroll = ttk.Scrollbar(editor_frame, orient="vertical", command=self.json_editor.yview)
        y_scroll.pack(side="right", fill="y")
        self.json_editor.configure(yscrollcommand=y_scroll.set)

    def load_sample_spec(self) -> None:
        self._load_spec_file(self.sample_spec_path)

    def open_spec(self) -> None:
        path = filedialog.askopenfilename(
            title="打开规格文件",
            filetypes=[("JSON 文件", "*.json"), ("所有文件", "*.*")],
            initialdir=str(self.project_root),
        )
        if path:
            self._load_spec_file(Path(path))

    def save_spec_as(self) -> None:
        payload = self._current_editor_text()
        if payload is None:
            return

        path = filedialog.asksaveasfilename(
            title="另存规格文件",
            defaultextension=".json",
            filetypes=[("JSON 文件", "*.json")],
            initialdir=str(self.project_root),
        )
        if not path:
            return

        Path(path).write_text(payload, encoding="utf-8")
        self.current_spec_path = Path(path)
        self._set_status(f"已保存规格文件：{path}")

    def validate_spec(self) -> None:
        spec = self._parse_editor_spec(show_errors=True)
        if spec is None:
            return
        self._sync_ui_from_spec(spec)
        self._set_status(f"规格校验通过，共 {len(spec.slides)} 页")
        messagebox.showinfo("校验结果", f"规格校验通过。\n页面数：{len(spec.slides)}")

    def generate_outline(self) -> None:
        self._run_outline_generation(use_feedback=False)

    def regenerate_outline(self) -> None:
        self._run_outline_generation(use_feedback=True)

    def _current_planner_settings(self) -> PlannerSettings:
        return PlannerSettings(
            model=self.model_var.get().strip() or "gpt-5-mini",
            base_url=self.base_url_var.get().strip() or None,
            api_key=self.api_key_var.get().strip() or None,
        )

    def _run_outline_generation(self, use_feedback: bool) -> None:
        requirement = self.requirement_editor.get("1.0", tk.END).strip()
        feedback = self.feedback_editor.get("1.0", tk.END).strip() if use_feedback else ""
        if not requirement:
            messagebox.showwarning("缺少需求", "请先输入用户需求，再生成大纲。")
            return

        self._save_desktop_settings()
        self._set_ai_stage("outline")
        self.ai_status_var.set("AI 工作流：正在生成大纲，请稍候。")
        self.root.update_idletasks()

        planner = OpenAIPlanner(self._current_planner_settings())
        try:
            outline = planner.generate_outline(requirement=requirement, feedback=feedback)
        except PlannerError as exc:
            self.ai_status_var.set("AI 工作流：生成大纲失败。")
            messagebox.showerror("生成大纲失败", str(exc))
            self._set_ai_stage("requirement")
            return

        self.current_outline = outline
        self._replace_outline_text(outline.model_dump_json(indent=2))

        if outline.needs_clarification:
            questions = "\n".join(f"{index + 1}. {item.question}" for index, item in enumerate(outline.clarification_questions))
            self._set_questions_text(questions or "当前没有待确认问题。")
            self._set_ai_stage("clarification")
            self.ai_status_var.set("AI 工作流：信息仍不足，请先回答问题，再结合补充重新生成大纲。")
            messagebox.showinfo("需要确认", f"当前信息不足，AI 需要继续确认：\n\n{questions}")
        else:
            self._set_questions_text("当前没有待确认问题。")
            self._set_ai_stage("ready_spec")
            self.ai_status_var.set("AI 工作流：大纲已生成，请检查并确认后再生成规格。")
            self._set_status(f"已生成大纲，预计 {outline.estimated_slides} 页")

        self._refresh_ai_controls()

    def confirm_outline_generate_spec(self) -> None:
        requirement = self.requirement_editor.get("1.0", tk.END).strip()
        feedback = self.feedback_editor.get("1.0", tk.END).strip()
        outline = self._parse_outline(show_errors=True)
        if outline is None:
            return
        if outline.needs_clarification:
            messagebox.showwarning("仍需确认", "当前大纲仍标记为需要继续确认，请先补充信息并重生成大纲。")
            return
        if not requirement:
            messagebox.showwarning("缺少需求", "请先输入用户需求。")
            return

        confirmed = messagebox.askyesno("确认大纲", "是否使用当前大纲生成最终规格？\n确认后才会进入 PPT 规格生成。")
        if not confirmed:
            return

        self._save_desktop_settings()
        self._set_ai_stage("spec_ready")
        self.ai_status_var.set("AI 工作流：正在根据已确认大纲生成规格，请稍候。")
        self.root.update_idletasks()

        planner = OpenAIPlanner(self._current_planner_settings())
        try:
            spec = planner.generate_deck_spec(requirement=requirement, feedback=feedback, outline=outline)
        except PlannerError as exc:
            self.ai_status_var.set("AI 工作流：生成规格失败。")
            messagebox.showerror("生成规格失败", str(exc))
            self._set_ai_stage("ready_spec")
            return

        self.current_outline = outline
        self._replace_editor_text(spec.model_dump_json(indent=2))
        self._sync_ui_from_spec(spec)
        self._set_questions_text("当前没有待确认问题。")
        self.ai_status_var.set("AI 工作流：规格已生成。你可以继续微调，确认后再生成 PPT。")
        self._set_status(f"AI 已生成规格，共 {len(spec.slides)} 页")
        self._refresh_ai_controls()

    def render_ppt(self) -> None:
        spec = self._parse_editor_spec(show_errors=True)
        if spec is None:
            return

        base_name = (self.current_spec_path.stem if self.current_spec_path else "deck").replace(" ", "_")
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        default_output = self.output_dir / f"{base_name}_{timestamp}.pptx"

        output_path = filedialog.asksaveasfilename(
            title="保存生成的 PPT",
            defaultextension=".pptx",
            filetypes=[("PowerPoint", "*.pptx")],
            initialdir=str(self.output_dir),
            initialfile=default_output.name,
        )
        if not output_path:
            return

        result = generate_ppt(spec, output_path=output_path)
        if result.ok:
            self.last_output_path = Path(output_path)
            self._set_status(f"PPT 已生成：{output_path}")
            open_now = messagebox.askyesno("生成完成", f"PPT 已成功生成。\n\n{output_path}\n\n现在打开吗？")
            if open_now:
                os.startfile(output_path)
            return

        self._set_status(result.message or "PPT 生成失败")
        messagebox.showerror("生成失败", result.message or "PPT 生成失败")

    def open_output_folder(self) -> None:
        folder = self.last_output_path.parent if self.last_output_path else self.output_dir
        folder.mkdir(parents=True, exist_ok=True)
        os.startfile(str(folder))

    def format_json(self) -> None:
        text = self._current_editor_text()
        if text is None:
            return
        try:
            payload = json.loads(text)
        except json.JSONDecodeError as exc:
            self._set_status(f"JSON 格式错误，第 {exc.lineno} 行")
            messagebox.showerror("JSON 错误", str(exc))
            return

        pretty = json.dumps(payload, ensure_ascii=False, indent=2)
        self._replace_editor_text(pretty)
        self._set_status("JSON 已格式化")

    def _load_spec_file(self, path: Path) -> None:
        try:
            payload = path.read_text(encoding="utf-8")
            self._replace_editor_text(payload)
            self.current_spec_path = path
            spec = DeckSpec.model_validate_json(payload)
        except Exception as exc:
            self._set_status(f"加载规格文件失败：{path.name}")
            messagebox.showerror("加载失败", str(exc))
            return

        self._sync_ui_from_spec(spec)
        self._set_status(f"已加载规格文件：{path}")

    def _parse_editor_spec(self, show_errors: bool) -> DeckSpec | None:
        text = self._current_editor_text()
        if text is None:
            return None
        try:
            return DeckSpec.model_validate_json(text)
        except Exception as exc:
            self._set_status("规格校验失败")
            if show_errors:
                messagebox.showerror("规格错误", str(exc))
            return None

    def _parse_outline(self, show_errors: bool) -> OutlinePlan | None:
        text = self.outline_editor.get("1.0", tk.END).strip()
        if not text:
            if show_errors:
                messagebox.showwarning("缺少大纲", "请先生成大纲。")
            return None
        try:
            return OutlinePlan.model_validate_json(text)
        except Exception as exc:
            if show_errors:
                messagebox.showerror("大纲错误", str(exc))
            return None

    def _sync_ui_from_spec(self, spec: DeckSpec) -> None:
        self.current_spec = spec
        self.meta_var.set(f"{spec.meta.title}  |  {len(spec.slides)} 页  |  比例: {spec.meta.ratio}  |  主题: {spec.theme.name}")
        self.summary_var.set(self._build_deck_summary(spec))

        self.slide_list.delete(0, tk.END)
        for index, slide in enumerate(spec.slides, start=1):
            kind_label = SLIDE_KIND_LABELS.get(slide.kind, slide.kind)
            self.slide_list.insert(tk.END, f"{index:02d}. [{kind_label}] {slide.title}")

        if spec.slides:
            self.slide_list.selection_clear(0, tk.END)
            self.slide_list.selection_set(0)
            self.slide_list.activate(0)
            self._render_slide_preview(spec.slides[0], 0)
        else:
            self._clear_preview()

    def _build_deck_summary(self, spec: DeckSpec) -> str:
        kinds: dict[str, int] = {}
        for slide in spec.slides:
            label = SLIDE_KIND_LABELS.get(slide.kind, slide.kind)
            kinds[label] = kinds.get(label, 0) + 1
        kind_summary = "，".join(f"{key}:{value}" for key, value in sorted(kinds.items()))
        return (
            f"作者：{spec.meta.author or '未填写'}\n"
            f"公司：{spec.meta.company or '未填写'}\n"
            f"页面构成：{kind_summary or '暂无'}"
        )

    def _on_slide_selected(self, _event=None) -> None:
        if not self.current_spec:
            return
        selection = self.slide_list.curselection()
        if not selection:
            return
        index = selection[0]
        if index < len(self.current_spec.slides):
            self._render_slide_preview(self.current_spec.slides[index], index)

    def _schedule_preview_refresh(self, _event=None) -> None:
        if self._refresh_job is not None:
            self.root.after_cancel(self._refresh_job)
        self._refresh_job = self.root.after(350, self._refresh_preview_from_editor)

    def _refresh_preview_from_editor(self) -> None:
        self._refresh_job = None
        spec = self._parse_editor_spec(show_errors=False)
        if spec is None:
            return

        selected = self.slide_list.curselection()
        self.current_spec = spec
        self.meta_var.set(f"{spec.meta.title}  |  {len(spec.slides)} 页  |  比例: {spec.meta.ratio}  |  主题: {spec.theme.name}")
        self.summary_var.set(self._build_deck_summary(spec))

        self.slide_list.delete(0, tk.END)
        for index, slide in enumerate(spec.slides, start=1):
            kind_label = SLIDE_KIND_LABELS.get(slide.kind, slide.kind)
            self.slide_list.insert(tk.END, f"{index:02d}. [{kind_label}] {slide.title}")

        new_index = min(selected[0], len(spec.slides) - 1) if selected and spec.slides else 0
        if spec.slides:
            self.slide_list.selection_set(new_index)
            self.slide_list.activate(new_index)
            self._render_slide_preview(spec.slides[new_index], new_index)
        else:
            self._clear_preview()

    def _clear_preview(self) -> None:
        self.preview_title_label.config(text="预览")
        self.preview_meta_label.config(text="尚未选择页面")
        self.preview_canvas.delete("all")

    def _render_slide_preview(self, slide: SlideSpec, index: int) -> None:
        if not self.current_spec:
            return
        total = len(self.current_spec.slides)
        kind_label = SLIDE_KIND_LABELS.get(slide.kind, slide.kind)
        self.preview_title_label.config(text=slide.title)
        self.preview_meta_label.config(text=f"第 {index + 1}/{total} 页  |  类型: {kind_label}")
        self._draw_slide(slide)

    def _on_preview_resize(self, _event=None) -> None:
        if not self.current_spec:
            return
        selection = self.slide_list.curselection()
        if selection and selection[0] < len(self.current_spec.slides):
            self._draw_slide(self.current_spec.slides[selection[0]])

    def _draw_slide(self, slide: SlideSpec) -> None:
        canvas = self.preview_canvas
        canvas.delete("all")

        width = max(canvas.winfo_width(), 640)
        height = max(canvas.winfo_height(), 480)
        margin = 28
        shell_w = width - margin * 2
        shell_h = min(height - margin * 2, int(shell_w * 9 / 16))
        shell_x = (width - shell_w) / 2
        shell_y = (height - shell_h) / 2

        canvas.create_rectangle(shell_x + 10, shell_y + 12, shell_x + shell_w + 14, shell_y + shell_h + 16, fill="#DDD4C8", outline="")
        canvas.create_rectangle(shell_x, shell_y, shell_x + shell_w, shell_y + shell_h, fill="#FAF8F4", outline="#D8D0C4", width=2)

        def rx(value: float) -> float:
            return shell_x + shell_w * value

        def ry(value: float) -> float:
            return shell_y + shell_h * value

        if slide.kind == "title":
            canvas.create_rectangle(rx(0.0), ry(0.0), rx(0.03), ry(1.0), fill=ACCENT, outline="")
            canvas.create_text(rx(0.09), ry(0.28), text=slide.title, anchor="w", font=("Microsoft YaHei UI", 24, "bold"), fill=TEXT_MAIN)
            subtitle = slide.subtitle or (self.current_spec.meta.subtitle if self.current_spec else "")
            canvas.create_text(rx(0.09), ry(0.40), text=subtitle, anchor="w", font=("Microsoft YaHei UI", 13), fill=TEXT_MUTED)
        elif slide.kind == "section":
            canvas.create_rectangle(rx(0.08), ry(0.23), rx(0.92), ry(0.61), fill=ACCENT_SOFT, outline="")
            canvas.create_text(rx(0.13), ry(0.42), text=slide.title, anchor="w", font=("Microsoft YaHei UI", 22, "bold"), fill=TEXT_MAIN)
            canvas.create_text(rx(0.13), ry(0.52), text=slide.subtitle or "章节页", anchor="w", font=("Microsoft YaHei UI", 12), fill=TEXT_MUTED)
        else:
            canvas.create_text(rx(0.06), ry(0.08), text=slide.title, anchor="w", font=("Microsoft YaHei UI", 18, "bold"), fill=TEXT_MAIN)
            if slide.subtitle:
                canvas.create_text(rx(0.06), ry(0.14), text=slide.subtitle, anchor="w", font=("Microsoft YaHei UI", 10), fill=TEXT_MUTED)

            if slide.kind == "two_column":
                self._draw_panel(canvas, rx(0.06), ry(0.20), rx(0.47), ry(0.88))
                self._draw_panel(canvas, rx(0.53), ry(0.20), rx(0.94), ry(0.88))
                self._draw_block_summary(canvas, slide.blocks[:1], rx(0.09), ry(0.25), rx(0.43), ry(0.80))
                self._draw_block_summary(canvas, slide.blocks[1:2], rx(0.56), ry(0.25), rx(0.90), ry(0.80))
            elif slide.kind == "metrics":
                metrics = slide.blocks[0].metrics if slide.blocks else []
                card_w = 0.26
                start_x = 0.06
                gap = 0.035
                for idx, metric in enumerate(metrics or []):
                    left = start_x + idx * (card_w + gap)
                    self._draw_panel(canvas, rx(left), ry(0.30), rx(left + card_w), ry(0.74))
                    canvas.create_text(rx(left + 0.03), ry(0.38), text=metric.label, anchor="w", font=("Microsoft YaHei UI", 10), fill=TEXT_MUTED)
                    canvas.create_text(rx(left + 0.03), ry(0.49), text=metric.value, anchor="w", font=("Microsoft YaHei UI", 18, "bold"), fill=TEXT_MAIN)
                    if metric.delta:
                        delta_color = ACCENT if not metric.delta.startswith("-") else ERROR
                        canvas.create_text(rx(left + 0.03), ry(0.60), text=metric.delta, anchor="w", font=("Microsoft YaHei UI", 10, "bold"), fill=delta_color)
            elif slide.kind == "chart":
                self._draw_panel(canvas, rx(0.06), ry(0.22), rx(0.94), ry(0.86))
                self._draw_chart_preview(canvas, slide.blocks, rx(0.09), ry(0.30), rx(0.91), ry(0.78))
            elif slide.kind == "table":
                self._draw_panel(canvas, rx(0.06), ry(0.22), rx(0.94), ry(0.86))
                self._draw_table_preview(canvas, slide.blocks, rx(0.08), ry(0.28), rx(0.92), ry(0.80))
            else:
                self._draw_panel(canvas, rx(0.06), ry(0.20), rx(0.94), ry(0.88))
                self._draw_block_summary(canvas, slide.blocks, rx(0.09), ry(0.26), rx(0.90), ry(0.82))

        if slide.speaker_notes:
            canvas.create_text(rx(0.06), ry(0.94), text=f"备注: {slide.speaker_notes[:100]}", anchor="w", font=("Microsoft YaHei UI", 9), fill=TEXT_MUTED)

    def _draw_panel(self, canvas: tk.Canvas, x1: float, y1: float, x2: float, y2: float) -> None:
        canvas.create_rectangle(x1, y1, x2, y2, fill=SURFACE_BG, outline=BORDER, width=1)

    def _draw_block_summary(self, canvas: tk.Canvas, blocks: list[BlockSpec], x1: float, y1: float, x2: float, y2: float) -> None:
        cursor_y = y1
        for block in blocks:
            if block.type == "paragraph":
                canvas.create_text(x1, cursor_y, text=self._truncate(block.text or "", 180), anchor="nw", width=x2 - x1, font=("Microsoft YaHei UI", 11), fill=TEXT_MAIN)
                cursor_y += 48
            elif block.type == "bullet_list":
                for item in block.items or []:
                    canvas.create_text(x1, cursor_y, text=f"- {self._truncate(item, 84)}", anchor="nw", width=x2 - x1, font=("Microsoft YaHei UI", 11), fill=TEXT_MAIN)
                    cursor_y += 28
            elif block.type == "code":
                canvas.create_rectangle(x1, cursor_y, x2, cursor_y + 88, fill="#F3EEE6", outline="")
                canvas.create_text(x1 + 12, cursor_y + 12, text=self._truncate(block.content or "", 180), anchor="nw", width=x2 - x1 - 24, font=FONT_MONO, fill=TEXT_MAIN)
                cursor_y += 100
            elif block.type == "image":
                canvas.create_rectangle(x1, cursor_y, x2, min(cursor_y + 140, y2), fill="#EFE7DA", outline="")
                canvas.create_text((x1 + x2) / 2, cursor_y + 68, text=f"图片\n{Path(block.path or '').name}", justify="center", font=FONT_UI, fill=TEXT_MUTED)
                cursor_y += 154
            elif block.type == "table":
                canvas.create_text(x1, cursor_y, text="表格内容", anchor="nw", font=("Microsoft YaHei UI", 11), fill=TEXT_MUTED)
                cursor_y += 26

    def _draw_table_preview(self, canvas: tk.Canvas, blocks: list[BlockSpec], x1: float, y1: float, x2: float, y2: float) -> None:
        table_block = next((block for block in blocks if block.type == "table"), None)
        if table_block is None:
            canvas.create_text(x1, y1, text="未找到表格块", anchor="nw", font=("Microsoft YaHei UI", 11), fill=ERROR)
            return

        columns = table_block.columns or []
        rows = table_block.rows or []
        total_rows = min(len(rows), 5) + 1
        if not columns:
            return

        col_width = (x2 - x1) / len(columns)
        row_height = min((y2 - y1) / total_rows, 44)

        for col_index, heading in enumerate(columns):
            cx1 = x1 + col_index * col_width
            cx2 = cx1 + col_width
            canvas.create_rectangle(cx1, y1, cx2, y1 + row_height, fill=ACCENT, outline="#FFFFFF")
            canvas.create_text(cx1 + 8, y1 + row_height / 2, text=heading, anchor="w", font=("Microsoft YaHei UI", 10, "bold"), fill="#FFFFFF")

        for row_index, row in enumerate(rows[:5], start=1):
            top = y1 + row_index * row_height
            fill = "#FFFFFF" if row_index % 2 == 1 else ACCENT_SOFT
            for col_index, value in enumerate(row):
                cx1 = x1 + col_index * col_width
                cx2 = cx1 + col_width
                canvas.create_rectangle(cx1, top, cx2, top + row_height, fill=fill, outline=BORDER)
                canvas.create_text(cx1 + 8, top + row_height / 2, text=str(value), anchor="w", font=("Microsoft YaHei UI", 10), fill=TEXT_MAIN)

    def _draw_chart_preview(self, canvas: tk.Canvas, blocks: list[BlockSpec], x1: float, y1: float, x2: float, y2: float) -> None:
        chart_block = next((block for block in blocks if block.type == "chart" and block.chart is not None), None)
        if chart_block is None:
            canvas.create_text(x1, y1, text="未找到图表块", anchor="nw", font=("Microsoft YaHei UI", 11), fill=ERROR)
            return

        chart = chart_block.chart
        if chart.title:
            canvas.create_text(x1, y1 - 14, text=chart.title, anchor="nw", font=("Microsoft YaHei UI", 11, "bold"), fill=TEXT_MAIN)

        base_y = y2
        canvas.create_line(x1, base_y, x2, base_y, fill=BORDER, width=1)
        canvas.create_line(x1, y1, x1, base_y, fill=BORDER, width=1)

        categories = chart.categories
        if not categories:
            return

        band_width = (x2 - x1) / len(categories)
        max_value = max((max(series.values) for series in chart.series), default=1) or 1
        series_count = max(1, len(chart.series))
        usable_width = band_width * 0.7
        bar_gap = usable_width / max(4, series_count * 2)
        bar_width = max(8, (usable_width - bar_gap * (series_count - 1)) / series_count)

        palette = [ACCENT, "#127A6A", "#D99A2B", "#7C4D23"]
        for cat_index, category in enumerate(categories):
            group_x = x1 + cat_index * band_width + band_width * 0.15
            canvas.create_text(x1 + cat_index * band_width + band_width / 2, base_y + 16, text=category, anchor="n", font=("Microsoft YaHei UI", 9), fill=TEXT_MUTED)
            for series_index, series in enumerate(chart.series):
                value = series.values[cat_index]
                height = 0 if max_value == 0 else (value / max_value) * (base_y - y1 - 20)
                left = group_x + series_index * (bar_width + bar_gap)
                top = base_y - height
                canvas.create_rectangle(left, top, left + bar_width, base_y, fill=palette[series_index % len(palette)], outline="")

    def _truncate(self, text: str, limit: int) -> str:
        compact = " ".join(text.split())
        return compact if len(compact) <= limit else compact[: limit - 1] + "…"

    def _replace_editor_text(self, text: str) -> None:
        self.json_editor.delete("1.0", tk.END)
        self.json_editor.insert("1.0", text)

    def _replace_outline_text(self, text: str) -> None:
        self.outline_editor.delete("1.0", tk.END)
        self.outline_editor.insert("1.0", text)
        self._refresh_ai_controls()

    def _current_editor_text(self) -> str | None:
        return self.json_editor.get("1.0", tk.END).strip()

    def _set_status(self, message: str) -> None:
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.status_var.set(f"[{timestamp}] {message}")


def main() -> None:
    root = tk.Tk()
    DesktopApp(root)
    root.mainloop()
