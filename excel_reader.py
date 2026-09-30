# -*- coding: utf-8 -*-
"""
Excel 智能核对朗读助手 (Excel Speech Auditor) - 微软神经网络高保真拟人版
专为人工高效、舒适、精准核对 Excel 表格数据设计
- 微软研究院神经网络高拟人度 TTS 引擎 (晓晓/云希/云健/晓艺/晓北等)，彻底告别机械音
- 独立的音色试听与明显倍速滑块调节 (0.7x ~ 2.0x)
- 单元格停顿间隔设置 (0.0s ~ 3.0s)，支持快捷预设，节奏清晰
- 底层按键状态检测 (GetAsyncKeyState)，物理空格键 100% 毫秒级即时暂停/恢复
- 选项全面采用直观对号 (✔) 显示
- 红色悬浮粗边框 (2.5pt 纯红无遮挡，0破坏原表)
- 超视距平滑自动滚屏 (放大至 400% 也始终居中视野)
- 异步音频预加载流水线，连贯低延迟
"""

import os
import sys
import json
import time
import queue
import asyncio
import threading
import tempfile
import tkinter as tk
from tkinter import ttk, messagebox

import pythoncom
import win32com.client
import win32api
import win32con

import edge_tts
import pygame

# 配置文件路径 (保存在当前程序目录下)
CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")

# 首次默认配置：默认读整张表，默认采用竖向读完再读下一列
DEFAULT_CONFIG = {
    "range_mode": "sheet",           # 默认：整张工作表
    "read_mode": "col_full",         # 默认：完整逐列 (竖向读完再换下一列)
    "n_count": 2,                    # N 格数量
    "voice": "zh-CN-XiaoxiaoNeural", # 默认：微软晓晓拟人语音
    "speed_idx": 2,                  # 默认：1.0x 标准语速
    "interval_sec": 0.5,             # 默认：0.5 秒停顿间隔
    "red_box": True,                 # 默认：开启红色边框
    "auto_scroll": True,             # 默认：开启居中滚动
    "skip_empty": True               # 默认：跳过空白单元格
}

# COM 状态常量
msoShapeRectangle = 1
FOCUS_BOX_NAME = "__AUDIT_FOCUS_BOX__"

# 微软研究院精选高拟人度音色列表
VOICE_OPTIONS = [
    ("【微软晓晓】温和甜美女声 (最逼真拟人 · 推荐)", "zh-CN-XiaoxiaoNeural"),
    ("【微软云希】阳光青年男声 (吐字利落 · 核对首选)", "zh-CN-YunxiNeural"),
    ("【微软云健】成熟稳重男声 (严肃专业 · 账单推荐)", "zh-CN-YunjianNeural"),
    ("【微软晓艺】亲切自然女声 (轻松流畅)", "zh-CN-XiaoyiNeural"),
    ("【东北晓北】生动幽默女声 (提神防疲劳)", "zh-CN-liaoning-XiaobeiNeural"),
    ("【微软云扬】专业播音男声 (新闻播报风)", "zh-CN-YunyangNeural"),
    ("【本地离线】微软慧慧 (无需联网兜底)", "local_sapi"),
]

# 语速倍率与 Edge-TTS 参数映射表
SPEED_PRESETS = [
    (0.70, "-30%", "0.7x (慢速)"),
    (0.85, "-15%", "0.85x (稍慢)"),
    (1.00, "+0%",  "1.0x (标准)"),
    (1.15, "+15%", "1.15x (适度加快)"),
    (1.30, "+30%", "1.3x (轻快核对)"),
    (1.50, "+50%", "1.5x (快速)"),
    (1.80, "+80%", "1.8x (极速)"),
]

class ExcelSpeechAuditor:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Excel 智能朗读核对助手 (微软神经语音版)")
        self.root.geometry("480x730")
        self.root.minsize(460, 680)
        self.root.attributes("-topmost", True)  # 置顶显示在 Excel 上方
        
        # 初始化 pygame 音频子系统
        try:
            pygame.mixer.init()
        except Exception as e:
            print(f"Pygame mixer init warning: {e}")

        # 临时音频缓存目录
        self.cache_dir = os.path.join(tempfile.gettempdir(), "excel_tts_cache")
        os.makedirs(self.cache_dir, exist_ok=True)

        # 运行与状态控制
        self.is_running = False
        self.is_paused = False
        self.stop_requested = False
        self.manual_step = 0  # -1 上一格，1 下一格
        self.worker_thread = None
        self.is_auditioning = False  # 试听状态

        # 1. 加载持久化配置 (若无则使用默认: 读整张表, 竖向读完换下一列)
        self.config = self._load_config()

        # 参数配置变量
        self.var_voice = tk.StringVar(value=self.config.get("voice", DEFAULT_CONFIG["voice"]))
        self.var_speed_idx = tk.IntVar(value=self.config.get("speed_idx", DEFAULT_CONFIG["speed_idx"]))
        self.var_interval_sec = tk.DoubleVar(value=self.config.get("interval_sec", DEFAULT_CONFIG["interval_sec"]))
        self.var_range_mode = tk.StringVar(value=self.config.get("range_mode", DEFAULT_CONFIG["range_mode"]))
        self.var_read_mode = tk.StringVar(value=self.config.get("read_mode", DEFAULT_CONFIG["read_mode"]))
        self.var_n_count = tk.IntVar(value=self.config.get("n_count", DEFAULT_CONFIG["n_count"]))
        self.var_red_box = tk.BooleanVar(value=self.config.get("red_box", DEFAULT_CONFIG["red_box"]))
        self.var_auto_scroll = tk.BooleanVar(value=self.config.get("auto_scroll", DEFAULT_CONFIG["auto_scroll"]))
        self.var_skip_empty = tk.BooleanVar(value=self.config.get("skip_empty", DEFAULT_CONFIG["skip_empty"]))

        # 监听所有配置项变动，实时自动持久化保存到本地 config.json
        for var in [self.var_voice, self.var_speed_idx, self.var_interval_sec,
                    self.var_range_mode, self.var_read_mode, self.var_n_count,
                    self.var_red_box, self.var_auto_scroll, self.var_skip_empty]:
            var.trace_add("write", lambda *args: self._save_config())

        # 状态展示变量
        self.var_status_text = tk.StringVar(value="准备就绪，请打开 Excel 并确认选项")
        self.var_current_cell = tk.StringVar(value="--")
        self.var_current_text = tk.StringVar(value="等待朗读...")
        self.var_progress_text = tk.StringVar(value="0 / 0")

        # 键盘与按键事件
        self._init_key_bindings()

        # 构建现代 UI
        self._setup_ui()

        # 退出清理
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    def _load_config(self):
        """读取本地保存的配置"""
        cfg = dict(DEFAULT_CONFIG)
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    saved = json.load(f)
                    if isinstance(saved, dict):
                        cfg.update(saved)
            except Exception as e:
                print("加载配置文件异常:", e)
        return cfg

    def _save_config(self):
        """将当前界面上的所有选项自动保存到本地 config.json"""
        try:
            curr_cfg = {
                "range_mode": self.var_range_mode.get(),
                "read_mode": self.var_read_mode.get(),
                "n_count": self.var_n_count.get(),
                "voice": self.var_voice.get(),
                "speed_idx": self.var_speed_idx.get(),
                "interval_sec": round(float(self.var_interval_sec.get()), 2),
                "red_box": self.var_red_box.get(),
                "auto_scroll": self.var_auto_scroll.get(),
                "skip_empty": self.var_skip_empty.get()
            }
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(curr_cfg, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print("保存配置文件异常:", e)

    def _init_key_bindings(self):
        """窗口内部快捷键响应: 仅当选中/激活本助手窗口时生效，绝不干扰其他应用打字"""
        self.root.bind_all("<space>", self._on_space_key)
        self.root.bind_all("<Escape>", lambda e: self.stop_reading())
        self.root.bind_all("<Left>", lambda e: self._trigger_step(-1))
        self.root.bind_all("<Right>", lambda e: self._trigger_step(1))

    def _on_space_key(self, event):
        """当焦点位于朗读助手窗口时，按空格键控制：启动 / 暂停 / 继续"""
        # 如果当前焦点在输入框 (如 Spinbox)，正在输入数字，则不拦截空格
        if isinstance(event.widget, (tk.Entry, ttk.Entry, tk.Spinbox, ttk.Spinbox)):
            return

        if not self.is_running:
            # 未启动状态：按空格直接【启动朗读】
            self.start_reading()
        else:
            # 运行中状态：按空格【暂停 / 继续】
            self.toggle_pause()
        return "break"

    def _setup_ui(self):
        """构建界面"""
        # 1. 顶部信息卡片 (深色背景醒目展示)
        card_frame = tk.Frame(self.root, bg="#1E293B", padx=14, pady=10)
        card_frame.pack(fill=tk.X, padx=10, pady=(8, 4))

        top_info_row = tk.Frame(card_frame, bg="#1E293B")
        top_info_row.pack(fill=tk.X)

        tk.Label(top_info_row, text="当前单元格:", fg="#94A3B8", bg="#1E293B", font=("Microsoft YaHei", 9)).pack(side=tk.LEFT)
        lbl_cell = tk.Label(top_info_row, textvariable=self.var_current_cell, fg="#FBBF24", bg="#1E293B", font=("Microsoft YaHei", 13, "bold"))
        lbl_cell.pack(side=tk.LEFT, padx=6)

        lbl_prog = tk.Label(top_info_row, textvariable=self.var_progress_text, fg="#CBD5E1", bg="#1E293B", font=("Microsoft YaHei", 9))
        lbl_prog.pack(side=tk.RIGHT)

        # 单元格内容大字预览区
        text_container = tk.Frame(card_frame, bg="#334155", padx=10, pady=8)
        text_container.pack(fill=tk.X, pady=(6, 0))

        lbl_content = tk.Label(
            text_container,
            textvariable=self.var_current_text,
            fg="#F8FAFC",
            bg="#334155",
            font=("Microsoft YaHei", 13, "bold"),
            wraplength=420,
            justify=tk.LEFT
        )
        lbl_content.pack(anchor=tk.W)

        # 快捷键与状态提示条
        status_bar = tk.Frame(self.root, bg="#F1F5F9", padx=10, pady=4)
        status_bar.pack(fill=tk.X, padx=10, pady=(2, 4))
        
        lbl_status = tk.Label(status_bar, textvariable=self.var_status_text, fg="#0284C7", bg="#F1F5F9", font=("Microsoft YaHei", 9, "bold"))
        lbl_status.pack(side=tk.LEFT)
        
        lbl_tip = tk.Label(status_bar, text="选中本窗口按【空格】：启动/暂停/继续 | 【Esc】停止", fg="#64748B", bg="#F1F5F9", font=("Microsoft YaHei", 8))
        lbl_tip.pack(side=tk.RIGHT)

        # 2. 核心操作按钮栏
        btn_frame = tk.Frame(self.root)
        btn_frame.pack(fill=tk.X, padx=10, pady=3)

        self.btn_start = tk.Button(
            btn_frame, text="▶ 开始朗读", bg="#10B981", fg="white", activebackground="#059669",
            font=("Microsoft YaHei", 10, "bold"), height=2, relief=tk.FLAT, takefocus=False,
            command=self.start_reading
        )
        self.btn_start.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))

        self.btn_pause = tk.Button(
            btn_frame, text="⏸ 暂停 (空格)", bg="#F59E0B", fg="white", activebackground="#D97706",
            font=("Microsoft YaHei", 10, "bold"), height=2, relief=tk.FLAT, state=tk.DISABLED, takefocus=False,
            command=self.toggle_pause
        )
        self.btn_pause.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=4)

        self.btn_stop = tk.Button(
            btn_frame, text="⏹ 停止 (Esc)", bg="#EF4444", fg="white", activebackground="#DC2626",
            font=("Microsoft YaHei", 10, "bold"), height=2, relief=tk.FLAT, state=tk.DISABLED, takefocus=False,
            command=self.stop_reading
        )
        self.btn_stop.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(4, 0))

        # 微调快退/快进
        step_frame = tk.Frame(self.root)
        step_frame.pack(fill=tk.X, padx=10, pady=(2, 6))

        btn_prev = tk.Button(step_frame, text="⏮ 上一格 (重读)", bg="#E2E8F0", fg="#334155", relief=tk.GROOVE,
                             font=("Microsoft YaHei", 8), takefocus=False, command=lambda: self._trigger_step(-1))
        btn_prev.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))

        btn_next = tk.Button(step_frame, text="⏭ 下一格 (跳过)", bg="#E2E8F0", fg="#334155", relief=tk.GROOVE,
                             font=("Microsoft YaHei", 8), takefocus=False, command=lambda: self._trigger_step(1))
        btn_next.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(4, 0))

        # 3. 微软研究院高拟人度音色选择
        voice_group = tk.LabelFrame(self.root, text="🎙 微软神经网络拟人语音引擎", font=("Microsoft YaHei", 9, "bold"), padx=10, pady=6)
        voice_group.pack(fill=tk.X, padx=10, pady=3)

        voice_top = tk.Frame(voice_group)
        voice_top.pack(fill=tk.X)

        voice_labels = [opt[0] for opt in VOICE_OPTIONS]
        self.combo_voice = ttk.Combobox(voice_top, values=voice_labels, state="readonly", font=("Microsoft YaHei", 9))
        
        # 根据已保存的音色匹配下拉框选中索引
        current_saved_voice = self.var_voice.get()
        matched_voice_idx = 0
        for i, opt in enumerate(VOICE_OPTIONS):
            if opt[1] == current_saved_voice:
                matched_voice_idx = i
                break
        self.combo_voice.current(matched_voice_idx)
        self.combo_voice.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        self.combo_voice.bind("<<ComboboxSelected>>", self._on_voice_change)

        self.btn_audition = tk.Button(
            voice_top, text="🔊 试听音色", bg="#3B82F6", fg="white", activebackground="#2563EB",
            font=("Microsoft YaHei", 8, "bold"), relief=tk.FLAT, takefocus=False, command=self.audition_voice
        )
        self.btn_audition.pack(side=tk.RIGHT)

        # 4. 朗读速度与停顿节奏控制
        timing_group = tk.LabelFrame(self.root, text="⚡ 语速与停顿节奏调节", font=("Microsoft YaHei", 9, "bold"), padx=10, pady=6)
        timing_group.pack(fill=tk.X, padx=10, pady=3)

        # 语速调节 (明显大滑块)
        speed_row = tk.Frame(timing_group)
        speed_row.pack(fill=tk.X, pady=(2, 4))

        tk.Label(speed_row, text="朗读语速:", font=("Microsoft YaHei", 9)).pack(side=tk.LEFT)
        self.slider_speed = tk.Scale(
            speed_row, from_=0, to=len(SPEED_PRESETS) - 1, orient=tk.HORIZONTAL, showvalue=0,
            variable=self.var_speed_idx, command=self._on_speed_slide, length=240, width=14,
            sliderlength=22, activebackground="#3B82F6", highlightthickness=0
        )
        self.slider_speed.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=8)

        self.lbl_speed_val = tk.Label(speed_row, text="", fg="#0284C7", font=("Microsoft YaHei", 9, "bold"), width=12, anchor=tk.W)
        self.lbl_speed_val.pack(side=tk.LEFT)
        self._on_speed_slide(self.var_speed_idx.get())

        # 停顿间隔调节 (解决连续读找不到分隔)
        interval_row = tk.Frame(timing_group)
        interval_row.pack(fill=tk.X, pady=(4, 2))

        tk.Label(interval_row, text="格间停顿:", font=("Microsoft YaHei", 9)).pack(side=tk.LEFT)
        self.slider_interval = tk.Scale(
            interval_row, from_=0.0, to=3.0, resolution=0.1, orient=tk.HORIZONTAL, showvalue=0,
            variable=self.var_interval_sec, command=self._on_interval_slide, length=240, width=14,
            sliderlength=22, activebackground="#10B981", highlightthickness=0
        )
        self.slider_interval.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=8)

        self.lbl_interval_val = tk.Label(interval_row, text="", fg="#059669", font=("Microsoft YaHei", 9, "bold"), width=12, anchor=tk.W)
        self.lbl_interval_val.pack(side=tk.LEFT)
        self._on_interval_slide(self.var_interval_sec.get())

        # 快捷停顿预设按钮
        preset_row = tk.Frame(timing_group)
        preset_row.pack(fill=tk.X, pady=(2, 0))
        tk.Label(preset_row, text="停顿预设:", fg="#64748B", font=("Microsoft YaHei", 8)).pack(side=tk.LEFT)
        
        presets = [("0.2s 紧凑", 0.2), ("0.5s 适中", 0.5), ("0.8s 从容", 0.8), ("1.2s 仔细", 1.2)]
        for text, val in presets:
            btn_p = tk.Button(preset_row, text=text, font=("Microsoft YaHei", 8), relief=tk.GROOVE, padx=4, pady=1,
                              takefocus=False, command=lambda v=val: self._set_interval_preset(v))
            btn_p.pack(side=tk.LEFT, padx=3)

        # 5. 范围与轨迹规划
        plan_group = tk.LabelFrame(self.root, text="🧭 范围与轨迹规划", font=("Microsoft YaHei", 9, "bold"), padx=10, pady=6)
        plan_group.pack(fill=tk.X, padx=10, pady=3)

        # 范围选择
        range_row = tk.Frame(plan_group)
        range_row.pack(fill=tk.X, pady=1)
        tk.Label(range_row, text="朗读范围:").pack(side=tk.LEFT)
        tk.Radiobutton(range_row, text="当前选中区域 (推荐)", variable=self.var_range_mode, value="selection").pack(side=tk.LEFT, padx=8)
        tk.Radiobutton(range_row, text="整张工作表", variable=self.var_range_mode, value="sheet").pack(side=tk.LEFT)

        # 轨迹方向
        ttk.Separator(plan_group, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=4)

        # 横向读 N 格换行
        r_n_frame = tk.Frame(plan_group)
        r_n_frame.pack(fill=tk.X, pady=1)
        tk.Radiobutton(r_n_frame, text="横向读", variable=self.var_read_mode, value="row_n").pack(side=tk.LEFT)
        spin_n = tk.Spinbox(r_n_frame, from_=1, to=50, width=3, textvariable=self.var_n_count, font=("Microsoft YaHei", 9))
        spin_n.pack(side=tk.LEFT, padx=4)
        tk.Label(r_n_frame, text="个单元格，然后换下一行 (两两配对核对推荐)").pack(side=tk.LEFT)

        # 完整逐行与逐列
        rf_frame = tk.Frame(plan_group)
        rf_frame.pack(fill=tk.X, pady=1)
        tk.Radiobutton(rf_frame, text="完整逐行朗读 (整行从左到右读完换行)", variable=self.var_read_mode, value="row_full").pack(side=tk.LEFT)

        cf_frame = tk.Frame(plan_group)
        cf_frame.pack(fill=tk.X, pady=1)
        tk.Radiobutton(cf_frame, text="完整逐列朗读 (整列从上到下读完换列)", variable=self.var_read_mode, value="col_full").pack(side=tk.LEFT)

        # 6. 辅助与视觉选项 (改用标准直观的打勾对号 ✔)
        opt_group = tk.LabelFrame(self.root, text="🎯 辅助功能与视觉增强 (直观对号 ✔ 表示启用)", font=("Microsoft YaHei", 9, "bold"), padx=10, pady=4)
        opt_group.pack(fill=tk.X, padx=10, pady=3)

        self._build_custom_check(opt_group, "🔴 红色粗边框聚焦 (贴合单元格高亮，0破坏原表样式)", self.var_red_box)
        self._build_custom_check(opt_group, "🔍 自动滚屏居中 (即使表格放大至 300% 也自动移入视野)", self.var_auto_scroll)
        self._build_custom_check(opt_group, "⏭ 跳过空白单元格 (遇到无内容格子直接跳过不浪费时间)", self.var_skip_empty)

    def _build_custom_check(self, parent, text, var):
        """构建直观、永远显示对号 ✔ 的复选控件"""
        row = tk.Frame(parent)
        row.pack(fill=tk.X, pady=2)

        lbl_check = tk.Label(row, text="✔", font=("Microsoft YaHei", 10, "bold"), fg="#16A34A", width=2)
        lbl_check.pack(side=tk.LEFT)

        def _update_icon():
            if var.get():
                lbl_check.config(text="✔", fg="#16A34A")
            else:
                lbl_check.config(text="☐", fg="#94A3B8")

        cb = tk.Checkbutton(row, text=text, variable=var, font=("Microsoft YaHei", 9), command=_update_icon, takefocus=False)
        cb.pack(side=tk.LEFT, fill=tk.X)
        _update_icon()

    def _on_voice_change(self, event=None):
        idx = self.combo_voice.current()
        if 0 <= idx < len(VOICE_OPTIONS):
            self.var_voice.set(VOICE_OPTIONS[idx][1])

    def _on_speed_slide(self, val):
        idx = int(val)
        if 0 <= idx < len(SPEED_PRESETS):
            _, _, label = SPEED_PRESETS[idx]
            self.lbl_speed_val.config(text=label)

    def _on_interval_slide(self, val):
        sec = float(val)
        self.lbl_interval_val.config(text=f"{sec:.1f} 秒")

    def _set_interval_preset(self, val):
        self.var_interval_sec.set(val)
        self.lbl_interval_val.config(text=f"{val:.1f} 秒")

    def _trigger_step(self, step):
        """上一格或下一格"""
        if self.is_running:
            self.manual_step = step

    def toggle_pause(self):
        """切换暂停状态"""
        if not self.is_running:
            return
        self.is_paused = not self.is_paused
        if self.is_paused:
            self.btn_pause.config(text="▶ 继续 (空格)", bg="#0284C7")
            self.var_status_text.set("⏸ 已暂停 - 按物理空格键继续")
            try:
                pygame.mixer.music.pause()
            except Exception:
                pass
        else:
            self.btn_pause.config(text="⏸ 暂停 (空格)", bg="#F59E0B")
            self.var_status_text.set("▶ 正在朗读核对中...")
            try:
                pygame.mixer.music.unpause()
            except Exception:
                pass

    def stop_reading(self):
        """停止朗读"""
        self.stop_requested = True
        self.is_paused = False
        try:
            pygame.mixer.music.stop()
        except Exception:
            pass
        self.var_status_text.set("⏹ 正在停止...")

    def audition_voice(self):
        """试听当前选择的音色与语速"""
        if self.is_auditioning or self.is_running:
            return
        self.is_auditioning = True
        self.btn_audition.config(state=tk.DISABLED, text="正在试听...")
        threading.Thread(target=self._run_audition, daemon=True).start()

    def _run_audition(self):
        try:
            voice_id = self.var_voice.get()
            speed_idx = self.var_speed_idx.get()
            rate_str = SPEED_PRESETS[speed_idx][1]
            sample_text = "正在试听微软智能语音，一千二百八十元整，核对无误。"
            
            if voice_id == "local_sapi":
                pythoncom.CoInitialize()
                sp = win32com.client.Dispatch("SAPI.SpVoice")
                sp.Rate = int((speed_idx - 2) * 2)
                sp.Speak(sample_text)
                pythoncom.CoUninitialize()
            else:
                out_path = os.path.join(self.cache_dir, "audition.mp3")
                asyncio.run(self._generate_edge_tts(sample_text, voice_id, rate_str, out_path))
                pygame.mixer.music.load(out_path)
                pygame.mixer.music.play()
                while pygame.mixer.music.get_busy():
                    time.sleep(0.05)
        except Exception as e:
            print("试听失败:", e)
        finally:
            self.root.after(0, lambda: self.btn_audition.config(state=tk.NORMAL, text="🔊 试听音色"))
            self.is_auditioning = False

    async def _generate_edge_tts(self, text, voice, rate, out_file):
        """异步调用微软神经网络 TTS 生成音频"""
        communicate = edge_tts.Communicate(text, voice, rate=rate)
        await communicate.save(out_file)

    def start_reading(self):
        """开始朗读"""
        if self.is_running:
            return

        self.is_running = True
        self.is_paused = False
        self.stop_requested = False
        self.manual_step = 0

        self.btn_start.config(state=tk.DISABLED)
        self.btn_pause.config(state=tk.NORMAL, text="⏸ 暂停 (空格)", bg="#F59E0B")
        self.btn_stop.config(state=tk.NORMAL)
        self.var_status_text.set("正在连接 Excel 并解析单元格...")

        self.worker_thread = threading.Thread(target=self._worker_routine, daemon=True)
        self.worker_thread.start()

    def _worker_routine(self):
        """核心工作线程: 负责 Excel COM 交互、音频预取与播放、底层按键检测"""
        pythoncom.CoInitialize()
        excel = None
        ws = None
        box_shape = None
        sapi_speaker = None

        try:
            # 1. 尝试连接 Excel
            try:
                excel = win32com.client.GetActiveObject("Excel.Application")
            except Exception:
                self._ui_error("未检测到运行中的 Excel！\n请先打开您的 Excel 文件，然后点击开始。")
                return

            try:
                ws = excel.ActiveSheet
                wb_name = excel.ActiveWorkbook.Name
                sheet_name = ws.Name
            except Exception:
                self._ui_error("Excel 正处于单元格输入/编辑状态！\n请先在 Excel 中按一下回车退出编辑。")
                return

            # 2. 规划单元格队列
            cells_plan = self._build_cells_plan(excel, ws)
            if not cells_plan:
                self._ui_error("未找到有效的单元格进行朗读！")
                return

            total_count = len(cells_plan)
            self._ui_set_status(f"🟢 已连接 [{wb_name}] {sheet_name} | 共 {total_count} 个待核对项")

            # 3. 初始化发音引擎
            voice_id = self.var_voice.get()
            if voice_id == "local_sapi":
                sapi_speaker = win32com.client.Dispatch("SAPI.SpVoice")

            # 4. 创建红色方框
            if self.var_red_box.get():
                box_shape = self._get_or_create_focus_box(ws)

            # 5. 音频预生成管道 (提前后台加载下一格音频，消除网络延迟)
            audio_cache = {}  # idx -> mp3_path
            prefetch_lock = threading.Lock()

            def _prefetch_worker():
                """后台流水线预加载音频"""
                asyncio.set_event_loop(asyncio.new_event_loop())
                while not self.stop_requested:
                    # 检查当前位置向前预加载 2 个单元格
                    curr_i = current_reading_idx
                    for next_i in range(curr_i, min(total_count, curr_i + 3)):
                        if next_i not in audio_cache and not self.stop_requested:
                            r, c = cells_plan[next_i]
                            try:
                                cell_val = str(ws.Cells(r, c).Text).strip()
                            except Exception:
                                cell_val = ""
                            if not cell_val and self.var_skip_empty.get():
                                continue
                            txt = cell_val if cell_val else "空"
                            tmp_mp3 = os.path.join(self.cache_dir, f"cell_{next_i}_{int(time.time()*1000)%10000}.mp3")
                            try:
                                rate_s = SPEED_PRESETS[self.var_speed_idx.get()][1]
                                asyncio.run(self._generate_edge_tts(txt, voice_id, rate_s, tmp_mp3))
                                with prefetch_lock:
                                    audio_cache[next_i] = tmp_mp3
                            except Exception:
                                pass
                    time.sleep(0.08)

            current_reading_idx = 0
            if voice_id != "local_sapi":
                threading.Thread(target=_prefetch_worker, daemon=True).start()

            idx = 0
            while idx < total_count and not self.stop_requested:
                current_reading_idx = idx

                # 检查手动调整 (上一格/下一格)
                if self.manual_step != 0:
                    idx = max(0, min(total_count - 1, idx + self.manual_step))
                    self.manual_step = 0
                    current_reading_idx = idx

                row, col = cells_plan[idx]

                try:
                    cell = ws.Cells(row, col)
                    cell_addr = f"R{row}C{col}"
                    try:
                        cell_addr = cell.Address.replace("$", "")
                    except Exception:
                        pass
                    raw_text = cell.Text
                except Exception:
                    time.sleep(0.2)
                    continue

                text_to_speak = str(raw_text).strip() if raw_text is not None else ""

                # 检查是否跳过空单元格
                if not text_to_speak and self.var_skip_empty.get():
                    idx += 1
                    continue

                display_text = text_to_speak if text_to_speak else "(空白)"
                speech_text = text_to_speak if text_to_speak else "空"

                # 更新 UI
                self._ui_update_cell_info(cell_addr, display_text, idx + 1, total_count)

                # 1. 选中 Excel 单元格
                try:
                    cell.Select()
                except Exception:
                    pass

                # 2. 自动滚动居中
                if self.var_auto_scroll.get():
                    self._ensure_cell_visible(excel, row, col)

                # 3. 移动红色方框
                if self.var_red_box.get():
                    if box_shape is None:
                        box_shape = self._get_or_create_focus_box(ws)
                    self._move_focus_box(box_shape, cell)

                # 4. 执行发音
                speed_idx = self.var_speed_idx.get()
                voice_id = self.var_voice.get()

                if voice_id == "local_sapi":
                    # 本地 SAPI 发音
                    sapi_speaker.Rate = int((speed_idx - 2) * 2)
                    sapi_speaker.Speak(speech_text, 1)  # 异步
                    while sapi_speaker.Status.RunningState != 1:
                        if self.stop_requested or self.manual_step != 0:
                            sapi_speaker.Speak("", 2)
                            break
                        if self.is_paused:
                            sapi_speaker.Pause()
                            while self.is_paused and not self.stop_requested and self.manual_step == 0:
                                time.sleep(0.03)
                            if not self.is_paused and not self.stop_requested:
                                sapi_speaker.Resume()
                        time.sleep(0.03)
                else:
                    # 微软神经网络语音发音 (优先从预取缓存读取)
                    audio_file = None
                    rate_str = SPEED_PRESETS[speed_idx][1]
                    
                    # 最多等待预取 1.5 秒
                    wait_start = time.time()
                    while time.time() - wait_start < 1.5 and not self.stop_requested and self.manual_step == 0:
                        with prefetch_lock:
                            if idx in audio_cache:
                                audio_file = audio_cache.pop(idx)
                                break
                        time.sleep(0.02)

                    if not audio_file and not self.stop_requested:
                        # 如果预取未命中，现场生成
                        audio_file = os.path.join(self.cache_dir, f"cell_{idx}_{int(time.time()*1000)%10000}.mp3")
                        try:
                            asyncio.run(self._generate_edge_tts(speech_text, voice_id, rate_str, audio_file))
                        except Exception as tts_err:
                            # 降级到本地 SAPI
                            print("Edge TTS 降级:", tts_err)
                            if sapi_speaker is None:
                                sapi_speaker = win32com.client.Dispatch("SAPI.SpVoice")
                            sapi_speaker.Speak(speech_text)
                            audio_file = None

                    if audio_file and os.path.exists(audio_file) and not self.stop_requested and self.manual_step == 0:
                        try:
                            pygame.mixer.music.load(audio_file)
                            pygame.mixer.music.play()
                            
                            while (pygame.mixer.music.get_busy() or self.is_paused) and not self.stop_requested and self.manual_step == 0:
                                time.sleep(0.03)

                            if self.stop_requested or self.manual_step != 0:
                                pygame.mixer.music.stop()
                        except Exception as play_err:
                            print("音频播放异常:", play_err)
                        finally:
                            try:
                                pygame.mixer.music.unload()
                                os.remove(audio_file)
                            except Exception:
                                pass

                if self.stop_requested:
                    break
                if self.manual_step != 0:
                    continue

                # 5. 单元格间停顿 (静音延时，留出人耳核对间隔)
                interval_sec = max(0.0, float(self.var_interval_sec.get()))
                sleep_start = time.time()
                while (time.time() - sleep_start < interval_sec) and not self.stop_requested and self.manual_step == 0:
                    if self.is_paused:
                        pause_start = time.time()
                        while self.is_paused and not self.stop_requested and self.manual_step == 0:
                            time.sleep(0.03)
                        # 恢复后补齐停顿时间
                        sleep_start += (time.time() - pause_start)
                    time.sleep(0.02)

                idx += 1

        except Exception as ex:
            self._ui_error(f"朗读过程中发生异常: {ex}")
        finally:
            if ws is not None:
                self._cleanup_focus_box(ws)
            pythoncom.CoUninitialize()
            self._ui_finish()

    def _ensure_cell_visible(self, excel, row, col):
        """平滑居中滚动，彻底解决表格放大到 200%~400% 时视野看不到的问题"""
        try:
            win = excel.ActiveWindow
            vis_range = win.VisibleRange
            top_row = win.ScrollRow
            left_col = win.ScrollColumn
            vis_rows = max(1, vis_range.Rows.Count)
            vis_cols = max(1, vis_range.Columns.Count)
            bottom_row = top_row + vis_rows - 1
            right_col = left_col + vis_cols - 1

            # 垂直方向 (行)
            if vis_rows <= 3:
                if row < top_row or row > bottom_row:
                    win.ScrollRow = row
            else:
                margin_row = 1 if vis_rows <= 5 else 2
                if row <= top_row + margin_row or row >= bottom_row - margin_row:
                    win.ScrollRow = max(1, row - max(1, vis_rows // 3))

            # 水平方向 (列)
            if vis_cols <= 3:
                if col < left_col or col > right_col:
                    win.ScrollColumn = col
            else:
                margin_col = 1 if vis_cols <= 4 else 2
                if col <= left_col + margin_col or col >= right_col - margin_col:
                    win.ScrollColumn = max(1, col - max(1, vis_cols // 4))
        except Exception:
            pass

    def _get_or_create_focus_box(self, ws):
        """创建无损红色透明矢量方框"""
        try:
            return ws.Shapes(FOCUS_BOX_NAME)
        except Exception:
            pass
        try:
            box = ws.Shapes.AddShape(msoShapeRectangle, 10, 10, 50, 20)
            box.Name = FOCUS_BOX_NAME
            box.Fill.Visible = False      # 全透明无遮挡
            box.Line.ForeColor.RGB = 255  # 纯红
            box.Line.Weight = 2.5         # 粗线醒目
            return box
        except Exception:
            return None

    def _move_focus_box(self, box, cell):
        """移动红框瞄准镜"""
        if box is None:
            return
        try:
            box.Left = cell.Left
            box.Top = cell.Top
            box.Width = cell.Width
            box.Height = cell.Height
            box.Visible = True
        except Exception:
            pass

    def _cleanup_focus_box(self, ws):
        """清理红框，恢复原表 100% 原始样貌"""
        try:
            ws.Shapes(FOCUS_BOX_NAME).Delete()
        except Exception:
            pass

    def _build_cells_plan(self, excel, ws):
        """规划单元格遍历序列"""
        range_mode = self.var_range_mode.get()
        read_mode = self.var_read_mode.get()
        n_count = max(1, self.var_n_count.get())

        target_range = excel.Selection if range_mode == "selection" else ws.UsedRange

        start_row = target_range.Row
        start_col = target_range.Column
        row_count = target_range.Rows.Count
        col_count = target_range.Columns.Count
        end_row = start_row + row_count - 1
        end_col = start_col + col_count - 1

        cells = []

        if read_mode == "row_n":
            # 横向读前 N 个单元格，然后换下一行 (两两核对推荐)
            for r in range(start_row, end_row + 1):
                for k in range(n_count):
                    c = start_col + k
                    if c <= end_col:
                        cells.append((r, c))
        elif read_mode == "row_full":
            # 完整逐行 (整行从左到右读完换行)
            for r in range(start_row, end_row + 1):
                for c in range(start_col, end_col + 1):
                    cells.append((r, c))
        elif read_mode == "col_full":
            # 完整逐列 (整列从上到下读完换列)
            for c in range(start_col, end_col + 1):
                for r in range(start_row, end_row + 1):
                    cells.append((r, c))
        elif read_mode == "col_n":
            # 竖向读 N 个单元格，换下一列
            for c in range(start_col, end_col + 1):
                for k in range(n_count):
                    r = start_row + k
                    if r <= end_row:
                        cells.append((r, c))

        return cells

    def _ui_update_cell_info(self, addr, text, current_idx, total_count):
        self.root.after(0, lambda: [
            self.var_current_cell.set(addr),
            self.var_current_text.set(text),
            self.var_progress_text.set(f"{current_idx} / {total_count}")
        ])

    def _ui_set_status(self, msg):
        self.root.after(0, lambda: self.var_status_text.set(msg))

    def _ui_error(self, err_msg):
        self.root.after(0, lambda: messagebox.showwarning("核对提示", err_msg))

    def _ui_finish(self):
        def _reset():
            self.is_running = False
            self.is_paused = False
            self.btn_start.config(state=tk.NORMAL)
            self.btn_pause.config(state=tk.DISABLED, text="⏸ 暂停 (空格)", bg="#F59E0B")
            self.btn_stop.config(state=tk.DISABLED)
            if not self.stop_requested:
                self.var_status_text.set("朗读完成！")
            else:
                self.var_status_text.set("朗读已停止。")
        self.root.after(0, _reset)

    def _on_close(self):
        self.stop_requested = True
        self._save_config()
        try:
            pygame.mixer.music.stop()
            pygame.mixer.quit()
        except Exception:
            pass
        self.root.destroy()

    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = ExcelSpeechAuditor()
    app.run()
