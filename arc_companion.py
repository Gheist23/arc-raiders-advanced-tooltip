from __future__ import annotations

import csv
import cv2
import numpy as np
import time
import ctypes
import os
from pathlib import Path
import requests
from mss import mss
import json
import difflib
import tkinter as tk
import re
import threading
import queue
import sys
import subprocess
import ast  # needed for robust Crafting parsing

from PIL import Image, ImageDraw, ImageFont, ImageTk
from tesserocr import PyTessBaseAPI, PSM

try:
    from pynput import keyboard as pynput_keyboard, mouse as pynput_mouse
    PYNPUT_AVAILABLE = True
except ImportError:
    pynput_keyboard = None
    pynput_mouse = None
    PYNPUT_AVAILABLE = False

from PySide6.QtCore import Qt
from PySide6.QtGui import QFont, QPalette, QColor, QIcon, QAction
from PySide6.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QLabel,
    QVBoxLayout,
    QHBoxLayout,
    QPushButton,
    QLineEdit,
    QCheckBox,
    QFrame,
    QDialog,
    QDialogButtonBox,
    QSystemTrayIcon,
    QMenu,
    QStyle,
    QSpinBox,
    QDoubleSpinBox,
    QColorDialog,
)


# ---------- NEW: per-user config directory (no extra dependency) ----------

def get_config_dir() -> Path:
    """
    Return a per-user configuration directory for ARC Companion.

    Windows: %APPDATA%\\ARC_Companion
    macOS:  ~/Library/Application Support/ARC_Companion
    Linux:  $XDG_CONFIG_HOME/ARC_Companion or ~/.config/ARC_Companion
    """
    if sys.platform.startswith("win"):
        base = os.getenv("APPDATA")
        if not base:
            base = str(Path.home())
        return Path(base) / "ARC_Companion"
    elif sys.platform == "darwin":
        # macOS
        return Path.home() / "Library" / "Application Support" / "ARC_Companion"
    else:
        # Linux / Unix
        base = os.getenv("XDG_CONFIG_HOME")
        if base:
            return Path(base) / "ARC_Companion"
        else:
            return Path.home() / ".config" / "ARC_Companion"


CONFIG_DIR = get_config_dir()
CONFIG_DIR.mkdir(parents=True, exist_ok=True)

SETTINGS_PATH = CONFIG_DIR / "arc_tooltip_settings.json"
VERDICTS_PATH = CONFIG_DIR / "arc_tooltip_verdicts.json"

# ------------------------------------------------------------------------

DEFAULT_SETTINGS = {
    "always_on": False,
    "hotkey": {
        "device": "keyboard",
        "key": "^",
    },
    "cycle_hotkey": {
        "device": "keyboard",
        "key": "space",
    },
    # base font size for tooltip label/body text
    "tooltip_font_size": 14,

    # show/hide Reverse Recycle and Crafting info
    "show_rr_and_crafting": True,

    # tooltip window opacity
    "tooltip_alpha": 0.94,

    # tooltip panel and text colors (hex; alpha optional)
    "tooltip_panel_color": "#3737370f",
    "tooltip_text_primary_color": "#141414ff",
    "tooltip_text_secondary_color": "#505050ff",
    "tooltip_keep_color": "#ff2828ff",
    "tooltip_recycle_color": "#28ffffff",
    "tooltip_sell_color": "#28ff28ff",
}


def load_settings() -> dict:
    """
    Load settings from JSON, merging with DEFAULT_SETTINGS and validating
    basic structure.
    """
    if not SETTINGS_PATH.is_file():
        return DEFAULT_SETTINGS.copy()

    try:
        with open(SETTINGS_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception:
        return DEFAULT_SETTINGS.copy()

    if not isinstance(data, dict):
        return DEFAULT_SETTINGS.copy()

    merged = DEFAULT_SETTINGS.copy()
    merged.update(
        {
            k: v
            for k, v in data.items()
            if k
               in (
                   "always_on",
                   "hotkey",
                   "cycle_hotkey",
                   "tooltip_font_size",
                   "show_rr_and_crafting",
                   "tooltip_alpha",
                   "tooltip_panel_color",
                   "tooltip_text_primary_color",
                   "tooltip_text_secondary_color",
                   "tooltip_keep_color",
                   "tooltip_recycle_color",
                   "tooltip_sell_color",
               )
        }
    )

    if isinstance(data.get("hotkey"), dict):
        hk = DEFAULT_SETTINGS["hotkey"].copy()
        hk.update({k: v for k, v in data["hotkey"].items() if k in ("device", "key")})
        merged["hotkey"] = hk

    if isinstance(data.get("cycle_hotkey"), dict):
        chk = DEFAULT_SETTINGS["cycle_hotkey"].copy()
        chk.update(
            {k: v for k, v in data["cycle_hotkey"].items() if k in ("device", "key")}
        )
        merged["cycle_hotkey"] = chk

    # Basic validation for tooltip_font_size
    tfs = data.get("tooltip_font_size", merged.get("tooltip_font_size"))
    try:
        tfs_int = int(tfs)
        merged["tooltip_font_size"] = max(10, min(32, tfs_int))
    except (TypeError, ValueError):
        merged["tooltip_font_size"] = DEFAULT_SETTINGS["tooltip_font_size"]

    # tooltip_alpha
    alpha_raw = data.get("tooltip_alpha", merged.get("tooltip_alpha"))
    try:
        alpha_val = float(alpha_raw)
        if not (0.1 <= alpha_val <= 1.0):
            raise ValueError
    except (TypeError, ValueError):
        alpha_val = DEFAULT_SETTINGS["tooltip_alpha"]
    merged["tooltip_alpha"] = alpha_val

    # show_rr_and_crafting
    merged["show_rr_and_crafting"] = bool(
        data.get("show_rr_and_crafting", merged.get("show_rr_and_crafting", True))
    )

    # normalize color strings
    color_keys = (
        "tooltip_panel_color",
        "tooltip_text_primary_color",
        "tooltip_text_secondary_color",
        "tooltip_keep_color",
        "tooltip_recycle_color",
        "tooltip_sell_color",
    )
    for ck in color_keys:
        val = data.get(ck, merged.get(ck, DEFAULT_SETTINGS.get(ck)))
        if isinstance(val, str) and val.strip():
            merged[ck] = val.strip()
        else:
            merged[ck] = DEFAULT_SETTINGS.get(ck)

    return merged


def save_settings(settings: dict) -> None:
    """
    Save settings dict to JSON.
    """
    with open(SETTINGS_PATH, "w", encoding="utf-8") as f:
        json.dump(settings, f, indent=2)


SETTINGS = DEFAULT_SETTINGS.copy()
HOTKEY_HELD = False
HOTKEY_LISTENERS_STARTED = False

USER_VERDICTS: dict[str, str] = {}
TOOLTIP_NEEDS_REFRESH = False

# alpha used by Tk overlay
TOOLTIP_ALPHA = DEFAULT_SETTINGS["tooltip_alpha"]


def refresh_settings():
    """
    Refresh SETTINGS from the shared settings file and update tooltip alpha.
    """
    global SETTINGS, TOOLTIP_ALPHA, TOOLTIP_ROOT
    try:
        SETTINGS = load_settings()
    except Exception as e:
        print(f"[helper] Failed to load settings from {SETTINGS_PATH}: {e}")
        SETTINGS = DEFAULT_SETTINGS.copy()

    # keep TOOLTIP_ALPHA in sync
    try:
        alpha = float(SETTINGS.get("tooltip_alpha", DEFAULT_SETTINGS["tooltip_alpha"]))
    except (TypeError, ValueError):
        alpha = DEFAULT_SETTINGS["tooltip_alpha"]
    alpha = max(0.1, min(1.0, alpha))
    TOOLTIP_ALPHA = alpha

    # if overlay already exists, update its alpha
    if "TOOLTIP_ROOT" in globals():
        try:
            if TOOLTIP_ROOT is not None:
                TOOLTIP_ROOT.attributes("-alpha", TOOLTIP_ALPHA)
        except Exception:
            pass


class HotkeyCaptureDialog(QDialog):
    """
    Minimal, clean dialog that waits for *one* key or mouse button press
    and returns (device, key) or (None, None) if cancelled.
    """

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Set Hotkey")
        self.setWindowFlag(Qt.WindowContextHelpButtonHint, False)
        self.setModal(True)

        self.setMinimumSize(340, 180)
        self.resize(420, 220)
        self.setSizeGripEnabled(True)

        self.device: str | None = None
        self.key: str | None = None

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 16, 20, 16)
        layout.setSpacing(14)

        title = QLabel("Press a key or mouse button")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("font-size: 15px; font-weight: 600;")
        layout.addWidget(title)

        subtitle = QLabel("Press ESC to cancel.")
        subtitle.setAlignment(Qt.AlignCenter)
        subtitle.setStyleSheet("color: #9aa0a6; font-size: 12px;")
        layout.addWidget(subtitle)

        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setFrameShadow(QFrame.Sunken)
        line.setStyleSheet("color: #5f6368;")
        layout.addWidget(line)

        button_box = QDialogButtonBox(QDialogButtonBox.Cancel)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)

    def keyPressEvent(self, event):
        k = event.key()
        if k == Qt.Key_Escape:
            self.device = None
            self.key = None
            self.reject()
            return

        text = event.text()
        if text and text.strip():
            key_str = text.lower()
        else:
            key_map = {
                Qt.Key_F1: "f1",
                Qt.Key_F2: "f2",
                Qt.Key_F3: "f3",
                Qt.Key_F4: "f4",
                Qt.Key_F5: "f5",
                Qt.Key_F6: "f6",
                Qt.Key_F7: "f7",
                Qt.Key_F8: "f8",
                Qt.Key_F9: "f9",
                Qt.Key_F10: "f10",
                Qt.Key_F11: "f11",
                Qt.Key_F12: "f12",
                Qt.Key_Tab: "tab",
                Qt.Key_Shift: "shift",
                Qt.Key_Control: "ctrl",
                Qt.Key_Alt: "alt",
                Qt.Key_Space: "space",
            }
            key_str = key_map.get(k)
            if key_str is None:
                key_str = event.text().lower() or f"key_{k}"

        self.device = "keyboard"
        self.key = key_str
        self.accept()

    def mousePressEvent(self, event):
        btn = event.button()
        btn_map = {
            Qt.LeftButton: "left",
            Qt.RightButton: "right",
            Qt.MiddleButton: "middle",
            Qt.XButton1: "x1",
            Qt.XButton2: "x2",
        }
        key_str = btn_map.get(btn)
        if not key_str:
            return
        self.device = "mouse"
        self.key = key_str
        self.accept()


class SettingsWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("ARC Companion")
        self.setWindowFlag(Qt.WindowContextHelpButtonHint, False)

        self._allow_close = False

        self.helper_process: subprocess.Popen | None = None

        self._init_responsive_size()

        self.settings = load_settings()

        if not SETTINGS_PATH.is_file():
            try:
                save_settings(self.settings)
            except Exception:
                pass

        central = QWidget()
        self.setCentralWidget(central)

        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(20, 20, 20, 16)
        main_layout.setSpacing(18)

        header_layout = QVBoxLayout()
        title = QLabel("ARC Companion")
        title.setStyleSheet("font-size: 20px; font-weight: 600;")
        subtitle = QLabel("Configure how the overlay is triggered while you play.")
        subtitle.setStyleSheet("color: #9aa0a6; font-size: 12px;")
        header_layout.addWidget(title)
        header_layout.addWidget(subtitle)
        main_layout.addLayout(header_layout)

        card = QFrame()
        card.setObjectName("Card")
        card.setStyleSheet(
            """
            QFrame#Card {
                background-color: #202124;
                border-radius: 10px;
                border: 1px solid #3c4043;
            }
            """
        )
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(16, 14, 16, 14)
        card_layout.setSpacing(12)

        # Hotkey row
        hotkey_row = QHBoxLayout()
        lbl_hotkey = QLabel("Hold to show tooltip:")
        lbl_hotkey.setMinimumWidth(160)

        self.hotkey_edit = QLineEdit()
        self.hotkey_edit.setReadOnly(True)
        self.hotkey_edit.setObjectName("HotkeyEdit")
        self.hotkey_edit.setFixedHeight(32)
        self.hotkey_edit.setStyleSheet(
            """
            QLineEdit#HotkeyEdit {
                border-radius: 6px;
                border: 1px solid #5f6368;
                padding: 4px 8px;
                background-color: #171717;
                color: #e8eaed;
                font-size: 12px;
            }
            QLineEdit#HotkeyEdit:disabled {
                color: #5f6368;
            }
            """
        )

        self.update_hotkey_display()

        btn_change = QPushButton("Change")
        btn_change.setFixedHeight(32)
        btn_change.setCursor(Qt.PointingHandCursor)
        btn_change.clicked.connect(self.on_change_hotkey)

        hotkey_row.addWidget(lbl_hotkey)
        hotkey_row.addWidget(self.hotkey_edit, stretch=1)
        hotkey_row.addWidget(btn_change)
        card_layout.addLayout(hotkey_row)

        # Cycle hotkey row
        cycle_row = QHBoxLayout()
        lbl_cycle = QLabel("Cycle suggested action:")
        lbl_cycle.setMinimumWidth(160)

        self.cycle_hotkey_edit = QLineEdit()
        self.cycle_hotkey_edit.setReadOnly(True)
        self.cycle_hotkey_edit.setObjectName("CycleHotkeyEdit")
        self.cycle_hotkey_edit.setFixedHeight(32)
        self.cycle_hotkey_edit.setStyleSheet(
            """
            QLineEdit#CycleHotkeyEdit {
                border-radius: 6px;
                border: 1px solid #5f6368;
                padding: 4px 8px;
                background-color: #171717;
                color: #e8eaed;
                font-size: 12px;
            }
            QLineEdit#CycleHotkeyEdit:disabled {
                color: #5f6368;
            }
            """
        )

        self.update_cycle_hotkey_display()

        btn_cycle_change = QPushButton("Change")
        btn_cycle_change.setFixedHeight(32)
        btn_cycle_change.setCursor(Qt.PointingHandCursor)
        btn_cycle_change.clicked.connect(self.on_change_cycle_hotkey)

        cycle_row.addWidget(lbl_cycle)
        cycle_row.addWidget(self.cycle_hotkey_edit, stretch=1)
        cycle_row.addWidget(btn_cycle_change)
        card_layout.addLayout(cycle_row)

        # Tooltip font size row
        font_row = QHBoxLayout()
        lbl_font = QLabel("Tooltip text size:")
        lbl_font.setMinimumWidth(160)

        self.font_spin = QSpinBox()
        self.font_spin.setRange(4, 64)
        self.font_spin.setFixedHeight(32)
        self.font_spin.setValue(
            int(
                self.settings.get(
                    "tooltip_font_size", DEFAULT_SETTINGS["tooltip_font_size"]
                )
            )
        )
        self.font_spin.valueChanged.connect(self.on_any_setting_changed)

        font_row.addWidget(lbl_font)
        font_row.addWidget(self.font_spin, stretch=1)
        card_layout.addLayout(font_row)

        # Show / hide Reverse Recycle & Crafting info
        rr_row = QHBoxLayout()
        lbl_rr = QLabel("Show RR & Crafting info:")
        lbl_rr.setMinimumWidth(160)

        self.chk_show_rr_crafting = QCheckBox()
        self.chk_show_rr_crafting.setChecked(
            bool(self.settings.get("show_rr_and_crafting", True))
        )
        self.chk_show_rr_crafting.stateChanged.connect(self.on_any_setting_changed)

        rr_row.addWidget(lbl_rr)
        rr_row.addWidget(self.chk_show_rr_crafting)
        rr_row.addStretch()
        card_layout.addLayout(rr_row)

        # Tooltip opacity
        opacity_row = QHBoxLayout()
        lbl_opacity = QLabel("Tooltip opacity:")
        lbl_opacity.setMinimumWidth(160)

        self.opacity_spin = QDoubleSpinBox()
        self.opacity_spin.setRange(0.30, 1.00)
        self.opacity_spin.setSingleStep(0.05)
        self.opacity_spin.setDecimals(2)
        self.opacity_spin.setFixedHeight(32)
        try:
            alpha_val = float(
                self.settings.get(
                    "tooltip_alpha", DEFAULT_SETTINGS.get("tooltip_alpha", 0.94)
                )
            )
        except (TypeError, ValueError):
            alpha_val = DEFAULT_SETTINGS.get("tooltip_alpha", 0.94)
        self.opacity_spin.setValue(alpha_val)
        self.opacity_spin.valueChanged.connect(self.on_any_setting_changed)

        opacity_row.addWidget(lbl_opacity)
        opacity_row.addWidget(self.opacity_spin, stretch=1)
        card_layout.addLayout(opacity_row)

        # Color rows use line edit + color picker button
        color_hint = "#RRGGBB or #RRGGBBAA"

        # Panel color
        panel_row = QHBoxLayout()
        lbl_panel = QLabel("Panel color:")
        lbl_panel.setMinimumWidth(160)

        self.panel_color_edit = QLineEdit()
        self.panel_color_edit.setFixedHeight(32)
        self.panel_color_edit.setPlaceholderText(color_hint)
        self.panel_color_edit.setText(
            str(
                self.settings.get(
                    "tooltip_panel_color",
                    DEFAULT_SETTINGS.get("tooltip_panel_color", "#3737370f"),
                )
            )
        )
        self.panel_color_edit.textChanged.connect(self.on_any_setting_changed)

        btn_panel = QPushButton("Pick")
        btn_panel.setFixedHeight(32)
        btn_panel.setCursor(Qt.PointingHandCursor)
        btn_panel.clicked.connect(
            lambda _=False, e=self.panel_color_edit, key="tooltip_panel_color": self.open_color_dialog(
                e, key
            )
        )

        panel_row.addWidget(lbl_panel)
        panel_row.addWidget(self.panel_color_edit, stretch=1)
        panel_row.addWidget(btn_panel)
        card_layout.addLayout(panel_row)

        # Primary text color
        primary_row = QHBoxLayout()
        lbl_primary = QLabel("Primary text color:")
        lbl_primary.setMinimumWidth(160)

        self.text_primary_color_edit = QLineEdit()
        self.text_primary_color_edit.setFixedHeight(32)
        self.text_primary_color_edit.setPlaceholderText(color_hint)
        self.text_primary_color_edit.setText(
            str(
                self.settings.get(
                    "tooltip_text_primary_color",
                    DEFAULT_SETTINGS.get(
                        "tooltip_text_primary_color", "#141414ff"
                    ),
                )
            )
        )
        self.text_primary_color_edit.textChanged.connect(
            self.on_any_setting_changed
        )

        btn_primary = QPushButton("Pick")
        btn_primary.setFixedHeight(32)
        btn_primary.setCursor(Qt.PointingHandCursor)
        btn_primary.clicked.connect(
            lambda _=False, e=self.text_primary_color_edit, key="tooltip_text_primary_color": self.open_color_dialog(
                e, key
            )
        )

        primary_row.addWidget(lbl_primary)
        primary_row.addWidget(self.text_primary_color_edit, stretch=1)
        primary_row.addWidget(btn_primary)
        card_layout.addLayout(primary_row)

        # Secondary text color
        secondary_row = QHBoxLayout()
        lbl_secondary = QLabel("Secondary text color:")
        lbl_secondary.setMinimumWidth(160)

        self.text_secondary_color_edit = QLineEdit()
        self.text_secondary_color_edit.setFixedHeight(32)
        self.text_secondary_color_edit.setPlaceholderText(color_hint)
        self.text_secondary_color_edit.setText(
            str(
                self.settings.get(
                    "tooltip_text_secondary_color",
                    DEFAULT_SETTINGS.get(
                        "tooltip_text_secondary_color", "#505050ff"
                    ),
                )
            )
        )
        self.text_secondary_color_edit.textChanged.connect(
            self.on_any_setting_changed
        )

        btn_secondary = QPushButton("Pick")
        btn_secondary.setFixedHeight(32)
        btn_secondary.setCursor(Qt.PointingHandCursor)
        btn_secondary.clicked.connect(
            lambda _=False, e=self.text_secondary_color_edit, key="tooltip_text_secondary_color": self.open_color_dialog(
                e, key
            )
        )

        secondary_row.addWidget(lbl_secondary)
        secondary_row.addWidget(self.text_secondary_color_edit, stretch=1)
        secondary_row.addWidget(btn_secondary)
        card_layout.addLayout(secondary_row)

        # KEEP color
        keep_row = QHBoxLayout()
        lbl_keep = QLabel("KEEP color:")
        lbl_keep.setMinimumWidth(160)

        self.keep_color_edit = QLineEdit()
        self.keep_color_edit.setFixedHeight(32)
        self.keep_color_edit.setPlaceholderText("#ff2828ff")
        self.keep_color_edit.setText(
            str(
                self.settings.get(
                    "tooltip_keep_color",
                    DEFAULT_SETTINGS.get("tooltip_keep_color", "#ff2828ff"),
                )
            )
        )
        self.keep_color_edit.textChanged.connect(self.on_any_setting_changed)

        btn_keep = QPushButton("Pick")
        btn_keep.setFixedHeight(32)
        btn_keep.setCursor(Qt.PointingHandCursor)
        btn_keep.clicked.connect(
            lambda _=False, e=self.keep_color_edit, key="tooltip_keep_color": self.open_color_dialog(
                e, key
            )
        )

        keep_row.addWidget(lbl_keep)
        keep_row.addWidget(self.keep_color_edit, stretch=1)
        keep_row.addWidget(btn_keep)
        card_layout.addLayout(keep_row)

        # RECYCLE color
        recycle_row = QHBoxLayout()
        lbl_recycle = QLabel("RECYCLE color:")
        lbl_recycle.setMinimumWidth(160)

        self.recycle_color_edit = QLineEdit()
        self.recycle_color_edit.setFixedHeight(32)
        self.recycle_color_edit.setPlaceholderText("#28ffffff")
        self.recycle_color_edit.setText(
            str(
                self.settings.get(
                    "tooltip_recycle_color",
                    DEFAULT_SETTINGS.get("tooltip_recycle_color", "#28ffffff"),
                )
            )
        )
        self.recycle_color_edit.textChanged.connect(self.on_any_setting_changed)

        btn_recycle = QPushButton("Pick")
        btn_recycle.setFixedHeight(32)
        btn_recycle.setCursor(Qt.PointingHandCursor)
        btn_recycle.clicked.connect(
            lambda _=False, e=self.recycle_color_edit, key="tooltip_recycle_color": self.open_color_dialog(
                e, key
            )
        )

        recycle_row.addWidget(lbl_recycle)
        recycle_row.addWidget(self.recycle_color_edit, stretch=1)
        recycle_row.addWidget(btn_recycle)
        card_layout.addLayout(recycle_row)

        # SELL color
        sell_row = QHBoxLayout()
        lbl_sell = QLabel("SELL color:")
        lbl_sell.setMinimumWidth(160)

        self.sell_color_edit = QLineEdit()
        self.sell_color_edit.setFixedHeight(32)
        self.sell_color_edit.setPlaceholderText("#28ff28ff")
        self.sell_color_edit.setText(
            str(
                self.settings.get(
                    "tooltip_sell_color",
                    DEFAULT_SETTINGS.get("tooltip_sell_color", "#28ff28ff"),
                )
            )
        )
        self.sell_color_edit.textChanged.connect(self.on_any_setting_changed)

        btn_sell = QPushButton("Pick")
        btn_sell.setFixedHeight(32)
        btn_sell.setCursor(Qt.PointingHandCursor)
        btn_sell.clicked.connect(
            lambda _=False, e=self.sell_color_edit, key="tooltip_sell_color": self.open_color_dialog(
                e, key
            )
        )

        sell_row.addWidget(lbl_sell)
        sell_row.addWidget(self.sell_color_edit, stretch=1)
        sell_row.addWidget(btn_sell)
        card_layout.addLayout(sell_row)

        # Always-on checkbox (still not added to layout, to preserve behavior)
        self.chk_always_on = QCheckBox("Always on")
        self.chk_always_on.setChecked(bool(self.settings.get("always_on", False)))
        self.chk_always_on.setStyleSheet(
            """
            QCheckBox {
                font-size: 13px;
            }
            """
        )
        self.chk_always_on.stateChanged.connect(self.on_any_setting_changed)

        # Note: not added to layout to keep behaviour identical
        # card_layout.addWidget(self.chk_always_on)

        main_layout.addWidget(card)

        footer = QLabel(
            "The tooltip helper runs automatically in the background. "
            "Changes take effect immediately."
        )
        footer.setStyleSheet("color: #9aa0a6; font-size: 11px;")
        footer.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        main_layout.addWidget(footer)

        self.start_helper_if_needed()

    def closeEvent(self, event):
        if self._allow_close:
            super().closeEvent(event)
        else:
            event.ignore()
            self.hide()

    def _init_responsive_size(self):
        """
        Set a resolution-aware initial size and a minimum size for the whole UI.
        """
        app = QApplication.instance()
        screen = self.screen() or (app.primaryScreen() if app else None)

        if screen:
            geo = screen.availableGeometry()
            w = max(int(geo.width() * 0.35), 800)
            h = max(int(geo.height() * 0.35), 500)
            self.resize(w, h)
        else:
            self.resize(900, 550)

        # minimum size of the whole UI
        self.setMinimumSize(720, 660)

    def _normalize_hex_for_qt(self, s: str, setting_key: str) -> str:
        """
        Ensure a hex string is in '#RRGGBB' form for QColor.
        Takes default from DEFAULT_SETTINGS if needed.
        """
        default_hex = DEFAULT_SETTINGS.get(setting_key, "#ffffff")
        if not isinstance(default_hex, str) or not default_hex:
            default_hex = "#ffffff"
        if not default_hex.startswith("#"):
            default_hex = "#" + default_hex
        if len(default_hex) >= 7:
            default_hex = default_hex[:7]

        if not isinstance(s, str):
            return default_hex
        s = s.strip()
        if not s:
            return default_hex
        if not s.startswith("#"):
            s = "#" + s
        if len(s) >= 7:
            return s[:7]
        return default_hex

    def open_color_dialog(self, line_edit: QLineEdit, setting_key: str):
        """
        Open a QColorDialog and store the chosen color as hex in the line_edit.
        """
        current = self._normalize_hex_for_qt(line_edit.text(), setting_key)
        col = QColor(current)
        if not col.isValid():
            col = QColor(self._normalize_hex_for_qt("", setting_key))

        chosen = QColorDialog.getColor(col, self, "Pick color")
        if chosen.isValid():
            hex_rgb = chosen.name(QColor.HexRgb)  # "#RRGGBB"
            line_edit.setText(hex_rgb)
            # textChanged -> on_any_setting_changed will handle saving

    def update_hotkey_display(self):
        hk = self.settings.get("hotkey") or {}
        device = hk.get("device", "keyboard")
        key = hk.get("key", "")

        if not key:
            text = "Not set"
        else:
            if device == "mouse":
                text = f"Mouse: {key.capitalize()}"
            else:
                text = f"Key: {key.upper()}"

        self.hotkey_edit.setText(text)

    def update_cycle_hotkey_display(self):
        chk = self.settings.get("cycle_hotkey") or {}
        device = chk.get("device", "keyboard")
        key = chk.get("key", "")

        if not key:
            text = "Not set"
        else:
            if device == "mouse":
                text = f"Mouse: {key.capitalize()}"
            else:
                text = f"Key: {key.upper()}"

        self.cycle_hotkey_edit.setText(text)

    def on_change_hotkey(self):
        dlg = HotkeyCaptureDialog(self)
        if dlg.exec() == QDialog.Accepted and dlg.device and dlg.key:
            self.settings["hotkey"] = {
                "device": dlg.device,
                "key": dlg.key,
            }
            self.update_hotkey_display()
            self.on_any_setting_changed()

    def on_change_cycle_hotkey(self):
        dlg = HotkeyCaptureDialog(self)
        if dlg.exec() == QDialog.Accepted and dlg.device and dlg.key:
            self.settings["cycle_hotkey"] = {
                "device": dlg.device,
                "key": dlg.key,
            }
            self.update_cycle_hotkey_display()
            self.on_any_setting_changed()

    def _save_current_settings(self) -> bool:
        """
        Save current settings; return True on success, False on error.
        """
        self.settings["always_on"] = bool(self.chk_always_on.isChecked())
        if hasattr(self, "font_spin"):
            self.settings["tooltip_font_size"] = int(self.font_spin.value())

        # show RR & Crafting
        if hasattr(self, "chk_show_rr_crafting"):
            self.settings["show_rr_and_crafting"] = bool(
                self.chk_show_rr_crafting.isChecked()
            )

        # opacity
        if hasattr(self, "opacity_spin"):
            try:
                alpha_val = float(self.opacity_spin.value())
            except (TypeError, ValueError):
                alpha_val = DEFAULT_SETTINGS.get("tooltip_alpha", 0.94)
            alpha_val = max(0.1, min(1.0, alpha_val))
            self.settings["tooltip_alpha"] = alpha_val

        # color helpers
        def _read_color(edit_attr, setting_key, default_key):
            edit = getattr(self, edit_attr, None)
            if edit is None:
                return
            txt = edit.text().strip()
            if not txt:
                txt = DEFAULT_SETTINGS.get(default_key, "")
            self.settings[setting_key] = txt

        _read_color("panel_color_edit", "tooltip_panel_color", "tooltip_panel_color")
        _read_color(
            "text_primary_color_edit",
            "tooltip_text_primary_color",
            "tooltip_text_primary_color",
        )
        _read_color(
            "text_secondary_color_edit",
            "tooltip_text_secondary_color",
            "tooltip_text_secondary_color",
        )
        _read_color("keep_color_edit", "tooltip_keep_color", "tooltip_keep_color")
        _read_color(
            "recycle_color_edit",
            "tooltip_recycle_color",
            "tooltip_recycle_color",
        )
        _read_color("sell_color_edit", "tooltip_sell_color", "tooltip_sell_color")

        try:
            save_settings(self.settings)
            return True
        except Exception as e:
            self.hotkey_edit.setText(f"Error saving settings: {e}")
            return False

    def on_any_setting_changed(self):
        self._save_current_settings()

    def start_helper_if_needed(self):
        """
        Automatically start helper process if it's not running yet.
        """
        if self.helper_process is not None and self.helper_process.poll() is None:
            return

        script_path = os.path.abspath(sys.argv[0])
        try:
            creation_flags = 0
            if os.name == "nt":
                creation_flags = subprocess.CREATE_NEW_PROCESS_GROUP

            self.helper_process = subprocess.Popen(
                [sys.executable, script_path, "--run-helper"],
                creationflags=creation_flags,
            )
        except Exception as e:
            pass
            # self.hotkey_edit.setText(f"Error starting helper: {e}")
            # self.helper_process = None


def create_dark_palette() -> QPalette:
    palette = QPalette()
    palette.setColor(QPalette.Window, QColor("#171717"))
    palette.setColor(QPalette.WindowText, QColor("#e8eaed"))
    palette.setColor(QPalette.Base, QColor("#202124"))
    palette.setColor(QPalette.AlternateBase, QColor("#202124"))
    palette.setColor(QPalette.ToolTipBase, QColor("#202124"))
    palette.setColor(QPalette.ToolTipText, QColor("#e8eaed"))
    palette.setColor(QPalette.Text, QColor("#e8eaed"))
    palette.setColor(QPalette.Button, QColor("#303134"))
    palette.setColor(QPalette.ButtonText, QColor("#e8eaed"))
    palette.setColor(QPalette.BrightText, Qt.red)
    palette.setColor(QPalette.Highlight, QColor("#8ab4f8"))
    palette.setColor(QPalette.HighlightedText, QColor("#202124"))
    return palette


def run_settings_ui():
    app = QApplication.instance()
    if app is None:
        app = QApplication(sys.argv)

    app.setStyle("Fusion")
    app.setPalette(create_dark_palette())

    app.setFont(QFont("Segoe UI", 11))

    app.setQuitOnLastWindowClosed(False)

    app.setStyleSheet(
        """
        QPushButton {
            background-color: #3c4043;
            color: #e8eaed;
            border-radius: 6px;
            padding: 6px 12px;
            border: 1px solid #5f6368;
            font-size: 12px;
        }
        QPushButton:hover {
            background-color: #5f6368;
        }
        QPushButton:pressed {
            background-color: #8ab4f8;
            border-color: #8ab4f8;
            color: #202124;
        }
        QCheckBox::indicator {
            width: 16px;
            height: 16px;
        }
        QCheckBox::indicator:unchecked {
            border-radius: 3px;
            border: 1px solid #5f6368;
            background-color: #171717;
        }
        QCheckBox::indicator:checked {
            border-radius: 3px;
            border: 1px solid #8ab4f8;
            background-color: #8ab4f8;
        }
        """
    )

    if not QSystemTrayIcon.isSystemTrayAvailable():
        print("System tray not available, running with normal window.")
        window = SettingsWindow()
        window.show()
        sys.exit(app.exec())

    window = SettingsWindow()
    window.hide()

    tray = QSystemTrayIcon()

    icon = QIcon("arc_tooltip_icon.png")
    if icon.isNull():
        icon = app.style().standardIcon(QStyle.SP_ComputerIcon)
    tray.setIcon(icon)
    tray.setToolTip("ARC Companion")

    menu = QMenu()
    action_settings = QAction("Settings", menu)
    action_quit = QAction("Quit", menu)

    menu.addAction(action_settings)
    menu.addSeparator()
    menu.addAction(action_quit)

    tray.setContextMenu(menu)

    def show_settings():
        if window.isMinimized():
            window.showNormal()
        if not window.isVisible():
            window.show()
        window.raise_()
        window.activateWindow()

    action_settings.triggered.connect(show_settings)

    def on_tray_activated(reason):
        if reason in (QSystemTrayIcon.Trigger, QSystemTrayIcon.DoubleClick):
            show_settings()

    tray.activated.connect(on_tray_activated)

    def quit_app():
        if window.helper_process is not None and window.helper_process.poll() is None:
            try:
                window.helper_process.terminate()
            except Exception:
                pass

        window._allow_close = True
        window.close()

        tray.hide()
        app.quit()

    action_quit.triggered.connect(quit_app)

    tray.show()

    sys.exit(app.exec())


TESSDATA_PATH = r"tessdata"

try:
    import mss.windows as mss_win

    mss_win.CAPTUREBLT = 0
except Exception:
    mss_win = None

try:
    response = requests.get("https://ghostworld073.pythonanywhere.com/arc_raiders_items")
    if response.status_code == 200:
        arc_raider_item_names = response.json()
    else:
        with open("arc_raiders_items.csv", newline="", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            arc_raider_item_names = list(reader)
except Exception:
    with open("arc_raiders_items.csv", newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        arc_raider_item_names = list(reader)

ITEM_LOOKUP = {}
ITEM_ORDER: dict[str, int] = {}


def safe_str(val, default=""):
    if val is None:
        return default
    return str(val)


VERDICT_CYCLE = ["KEEP", "RECYCLE", "SELL"]


def load_user_verdicts() -> None:
    """
    Load per-item verdict overrides from JSON.
    """
    global USER_VERDICTS
    if not VERDICTS_PATH.is_file():
        USER_VERDICTS = {}
        return

    try:
        with open(VERDICTS_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception as e:
        print(f"[helper] Failed to load verdict overrides from {VERDICTS_PATH}: {e}")
        USER_VERDICTS = {}
        return

    if not isinstance(data, dict):
        USER_VERDICTS = {}
        return

    USER_VERDICTS = {safe_str(k): safe_str(v).upper() for k, v in data.items()}


def save_user_verdicts() -> None:
    """
    Save per-item verdict overrides to JSON.
    """
    try:
        with open(VERDICTS_PATH, "w", encoding="utf-8") as f:
            json.dump(USER_VERDICTS, f, indent=2)
    except Exception as e:
        print(f"[helper] Failed to save verdict overrides to {VERDICTS_PATH}: {e}")


def get_effective_verdict(row, detected_name: str | None) -> str:
    """
    Return the verdict that should be shown for a row, taking the user
    override (if any) into account.
    """
    base = safe_str(row.get("Verdict", "") if row else "").upper()
    name_key = safe_str(row.get("Name") if row else detected_name).strip()

    override = USER_VERDICTS.get(name_key) if name_key else None
    eff = (override or base).upper()
    return eff or "UNKNOWN"


def cycle_verdict_for_current_item(direction: int = 1) -> None:
    """
    Cycle the verdict for the currently shown tooltip item and persist it.
    direction: +1 to go forward in VERDICT_CYCLE.
    """
    global TOOLTIP_NEEDS_REFRESH

    if LAST_SHOWN_ROW is None:
        return

    name = safe_str(LAST_SHOWN_ROW.get("Name"))
    if not name:
        return

    base = safe_str(LAST_SHOWN_ROW.get("Verdict", "")).upper()
    current = USER_VERDICTS.get(name) or base or VERDICT_CYCLE[0]

    if current not in VERDICT_CYCLE:
        current = VERDICT_CYCLE[0]

    idx = VERDICT_CYCLE.index(current)
    if direction >= 0:
        idx = (idx + 1) % len(VERDICT_CYCLE)
    else:
        idx = (idx - 1) % len(VERDICT_CYCLE)

    new_verdict = VERDICT_CYCLE[idx]
    USER_VERDICTS[name] = new_verdict
    save_user_verdicts()
    TOOLTIP_NEEDS_REFRESH = True
    print(f"[helper] Override for '{name}' -> {new_verdict}")


def normalize_name_for_match(name: str) -> str:
    """
    Normalize strings so OCR quirks (I vs l, 0 vs O, etc.) still match.

    Additionally normalizes trailing roman numerals / digits so that
    e.g. "Extended Light Mag II" and "EXTENDED LIGHT MAG 2" end up
    with the same normalized key.

    Special case: for very short names (<= 3 non-space chars, e.g. "OIL"),
    we do NOT apply the I/|/1 -> l mapping, so "OIL" and "Oil" both become "oil".
    """
    s = safe_str(name).strip()

    if s:
        m = re.search(r"(\b[IVXLCDM]+\b|\b\d+\b)$", s, re.IGNORECASE)
        if m:
            token = m.group(1)
            token_clean = token.upper().replace("|", "I").replace("L", "I")

            roman_map = {
                "I": "1",
                "II": "2",
                "III": "3",
                "IV": "4",
            }

            digit = None
            if token_clean in roman_map:
                digit = roman_map[token_clean]
            elif token_clean.isdigit():
                digit = token_clean

            if digit is not None:
                s = s[: m.start(1)] + digit

    core = re.sub(r"\s+", "", s)
    if len(core) <= 3:
        trans = str.maketrans(
            {
                "0": "o",
            }
        )
    else:
        trans = str.maketrans(
            {
                "I": "l",
                "|": "l",
                "1": "l",
                "0": "o",
            }
        )

    s = s.translate(trans).lower()

    s = re.sub(r"\s+", " ", s)
    return s


def build_item_lookup():
    """
    Build:
      - ITEM_LOOKUP: normalized name -> row
      - ITEM_ORDER:  normalized name -> index in arc_raider_item_names
    """
    global ITEM_LOOKUP, ITEM_ORDER
    ITEM_LOOKUP = {}
    ITEM_ORDER = {}

    for idx, row in enumerate(arc_raider_item_names):
        name = str(row.get("Name", "")).strip()
        if not name:
            continue

        norm = normalize_name_for_match(name)

        # For lookup, keep the last row as before
        ITEM_LOOKUP[norm] = row

        # For ordering, remember the first occurrence
        if norm not in ITEM_ORDER:
            ITEM_ORDER[norm] = idx


def get_csv_index_for_name(name: str) -> int | None:
    """
    Return the index of an item in the original CSV, or None if not found.
    Matching uses the same normalization as ITEM_LOOKUP.
    """
    norm = normalize_name_for_match(name)
    return ITEM_ORDER.get(norm)


build_item_lookup()


def find_item_row_by_name(name: str):
    if not name or not ITEM_LOOKUP:
        return None

    norm = normalize_name_for_match(name)
    keys = list(ITEM_LOOKUP.keys())

    if norm in ITEM_LOOKUP:
        return ITEM_LOOKUP[norm]

    tokens = [t for t in norm.split() if len(t) >= 3]

    if tokens:
        strict_candidates = [k for k in keys if all(t in k for t in tokens)]
        if strict_candidates:
            best_key = max(
                strict_candidates,
                key=lambda k: difflib.SequenceMatcher(None, norm, k).ratio(),
            )
            ratio = difflib.SequenceMatcher(None, norm, best_key).ratio()
            if ratio >= 0.6:
                return ITEM_LOOKUP[best_key]

    partial_candidates = [k for k in keys if norm in k or k in norm]
    if partial_candidates:
        best_key = max(
            partial_candidates,
            key=lambda k: difflib.SequenceMatcher(None, norm, k).ratio(),
        )
        ratio = difflib.SequenceMatcher(None, norm, best_key).ratio()
        if ratio >= 0.70:
            return ITEM_LOOKUP[best_key]

    best = difflib.get_close_matches(norm, keys, n=1, cutoff=0.70)
    if best:
        return ITEM_LOOKUP[best[0]]

    return None


def format_percentage(value):
    if value is None:
        return "N/A"

    if isinstance(value, (int, float)):
        v = float(value)
    else:
        s = str(value).strip()
        if not s or s in ("-", "nan", "NaN"):
            return "N/A"
        try:
            v = float(s)
        except ValueError:
            return s

    sign = "+" if v > 0 else ""
    if abs(v - int(v)) < 1e-6:
        return f"{sign}{int(v)}%"
    return f"{sign}{v:.1f}%"


def parse_reverse_recycle(row):
    """
    Parse the 'Reverse Recycle' column.

    - Lines are grouped for names with trailing numbers (e.g. "Anvil 1/2/3")
      into a single line like "3-9x Anvil" (if multiple counts exist).
    - Sorting now uses:
        1) Category (alphabetically; unknown categories last)
        2) Amount (descending)
        3) Line text (alphabetically)
    """
    raw = safe_str(row.get("Reverse Recycle", "")).strip()
    if not raw or raw == "[]":
        return []

    try:
        data = json.loads(raw)
    except Exception:
        # Fallback: just show the raw content
        return [raw]

    # Collect raw entries with their CSV index & category
    item_entries: list[dict] = []
    for entry in data:
        if not entry:
            continue

        if len(entry) >= 2:
            item, count = entry[0], entry[1]
            try:
                c = int(count)
            except (ValueError, TypeError):
                c = 0
        else:
            item = entry[0]
            c = 0

        item_name = safe_str(item).strip()
        idx = get_csv_index_for_name(item_name)

        category = ""
        if idx is not None and 0 <= idx < len(arc_raider_item_names):
            try:
                category = safe_str(
                    arc_raider_item_names[idx].get("Category", "")
                )
            except Exception:
                category = ""

        item_entries.append(
            {
                "name": item_name,
                "count": c,
                "index": idx,
                "category": category,
            }
        )

    if not item_entries:
        return []

    # --- Group items that end with a trailing number: "Anvil 1", "Anvil 2", ...
    suffix_re = re.compile(r"^(.*\D)(\d+)$")

    numeric_groups: dict[str, dict] = {}
    no_suffix_entries: list[dict] = []

    for ent in item_entries:
        name = ent["name"]
        m = suffix_re.match(name)
        if m:
            base = m.group(1).strip()  # "Anvil"
            key = base.lower()
            g = numeric_groups.get(key)
            if g is None:
                g = {"base": base, "entries": []}
                numeric_groups[key] = g
            g["entries"].append(ent)
        else:
            no_suffix_entries.append(ent)

    groups_out: list[dict] = []

    # Summarise groups with >= 2 numeric-suffix variants; singletons fall back
    for g in numeric_groups.values():
        entries = g["entries"]
        if len(entries) == 1:
            # Only one "Anvil N" -> keep as-is, including the "N"
            no_suffix_entries.append(entries[0])
            continue

        counts = [e["count"] for e in entries if isinstance(e["count"], (int, float))]
        if not counts:
            # No usable counts, degrade to plain entries
            no_suffix_entries.extend(entries)
            continue

        min_c = min(counts)
        max_c = max(counts)

        if min_c == max_c:
            count_text = f"{min_c}x"
        else:
            count_text = f"{min_c}-{max_c}x"

        base_name = g["base"]  # e.g. "Anvil"

        idxs = [e["index"] for e in entries if e["index"] is not None]
        group_index = min(idxs) if idxs else None

        # Choose category from first non-empty
        cat = ""
        for e in entries:
            if e.get("category"):
                cat = e["category"]
                break

        groups_out.append(
            {
                "line": f"{count_text} {base_name}",
                "index": group_index,
                "count_key": max_c,
                "category": cat,
            }
        )

    # Non-grouped entries (including numeric singletons and non-numeric names)
    for ent in no_suffix_entries:
        c = ent["count"]
        name = ent["name"]
        idx = ent["index"]
        cat = ent.get("category", "")

        if c > 0:
            line = f"{c}x {name}"
        else:
            line = name

        groups_out.append(
            {
                "line": line,
                "index": idx,
                "count_key": c,
                "category": cat,
            }
        )

    # Sort by:
    #   1) Known categories before unknown
    #   2) "Loot" category first among known categories
    #   3) Category name (alphabetically)
    #   4) Amount (descending)
    #   5) Line text (alphabetically)
    def sort_key(g: dict):
        cat = (g.get("category") or "").strip()
        is_unknown = 1 if not cat else 0
        cat_norm = cat.lower()

        loot_priority = 0 if cat_norm == "loot" else 1

        count_val = int(g.get("count_key", 0) or 0)
        return (is_unknown, loot_priority, cat_norm, -count_val, g["line"].lower())

    groups_out.sort(key=sort_key)

    return [g["line"] for g in groups_out]


def parse_crafting(row, return_meta: bool = False):
    """
    Parse the 'Crafting' column.

    - Robust to JSON or Python-list representations.
    - Returns either:
        * list[str]: "count x item" strings   (default)
        * list[dict]: when return_meta=True, each dict has:
            {
                "line": "3x Vulcano I",
                "name": "Vulcano I",
                "count": 3,
                "index": <csv index or None>,
                "category": <Category from DB or "">
            }
    - Sorted primarily by CSV index (ITEM_ORDER), unknown items last.
    """
    raw = safe_str(row.get("Crafting", "")).strip()
    if not raw or raw == "[]":
        return [] if not return_meta else []

    data = None

    # 1) Try JSON first
    try:
        data = json.loads(raw)
    except Exception:
        data = None

    # 2) Try Python list literal if JSON failed
    if data is None:
        try:
            lit = ast.literal_eval(raw)
        except Exception:
            lit = None
        if isinstance(lit, (list, tuple)):
            data = lit

    # 3) Fallback: comma-separated list or raw string
    if data is None:
        parts = [p.strip() for p in raw.split(",") if p.strip()]
        if not return_meta:
            return parts if parts else [raw]
        else:
            if not parts:
                return [{"line": raw, "name": raw, "count": 0, "index": None, "category": ""}]
            out_meta = []
            for p in parts:
                idx = get_csv_index_for_name(p)
                category = ""
                if idx is not None and 0 <= idx < len(arc_raider_item_names):
                    try:
                        category = safe_str(
                            arc_raider_item_names[idx].get("Category", "")
                        )
                    except Exception:
                        category = ""
                out_meta.append(
                    {
                        "line": p,
                        "name": p,
                        "count": 0,
                        "index": idx,
                        "category": category,
                    }
                )
            return out_meta

    entries: list[dict] = []

    for entry in data:
        if not entry:
            continue

        if isinstance(entry, (list, tuple)) and len(entry) >= 2:
            item, count = entry[0], entry[1]
            try:
                c = int(count)
            except (ValueError, TypeError):
                c = 0
        else:
            item = entry[0] if isinstance(entry, (list, tuple)) else entry
            c = 0

        item_name = safe_str(item).strip()
        idx = get_csv_index_for_name(item_name)

        category = ""
        if idx is not None and 0 <= idx < len(arc_raider_item_names):
            try:
                category = safe_str(
                    arc_raider_item_names[idx].get("Category", "")
                )
            except Exception:
                category = ""

        entries.append(
            {
                "name": item_name,
                "count": c,
                "index": idx,
                "category": category,
            }
        )

    if not entries:
        return [] if not return_meta else []

    def sort_key(e: dict):
        idx = e["index"]
        if idx is None:
            # Unknown index -> after known ones, by descending count then name
            return (10**9, -int(e["count"] or 0), e["name"].lower())
        return (idx, 0, e["name"].lower())

    entries.sort(key=sort_key)

    lines: list[str] = []
    meta_out: list[dict] = []
    for e in entries:
        c = e["count"]
        name = e["name"]
        if c > 0:
            line = f"{c}x {name}"
        else:
            line = name

        lines.append(line)
        meta_out.append(
            {
                "line": line,
                "name": name,
                "count": c,
                "index": e["index"],
                "category": e["category"],
            }
        )

    return meta_out if return_meta else lines


ROMAN_SUFFIXES: tuple[str, ...] = ("I", "II", "III", "IV")
ROMAN_ORDER = {r: i for i, r in enumerate(ROMAN_SUFFIXES)}


def condense_roman_variants(lines: list[str], drop_suffix: bool = False) -> list[str]:
    """
    Group lines that differ only by a trailing roman numeral (I–IV) into
    a single line like 'Anvil I, II, III, IV'.

    Quantities ('3x Anvil I') are preserved as long as they are the same
    for all grouped variants. If counts differ, the original lines are kept.

    If drop_suffix=True, the condensed line omits the roman numerals entirely,
    e.g. 'Anvil' or '3x Anvil'.
    """
    groups: dict[tuple[str, str], dict] = {}
    passthrough: list[tuple[int, str]] = []

    for idx, original in enumerate(lines):
        line = original.strip()
        if not line:
            continue

        m = re.match(r"^\s*(\d+)\s*[x×]\s*(.+)$", line)
        if m:
            qty_str = m.group(1)
            rest = m.group(2).strip()
        else:
            qty_str = None
            rest = line

        rest_no_punct = rest.rstrip(",.;: ")
        tokens = rest_no_punct.split()
        if len(tokens) >= 2:
            last = tokens[-1]
            last_norm = last.upper().replace("|", "I").replace("L", "I")
            if last_norm in ROMAN_SUFFIXES:
                base = " ".join(tokens[:-1])
                roman = last_norm

                key = (base.lower(), qty_str or "")
                g = groups.get(key)
                if g is None:
                    g = {"base": base, "qty": qty_str, "romans": [roman], "first_index": idx}
                    groups[key] = g
                else:
                    if roman not in g["romans"]:
                        g["romans"].append(roman)
                continue

        passthrough.append((idx, original))

    out_entries: list[tuple[int, str]] = []

    for g in groups.values():
        romans_sorted = sorted(g["romans"], key=lambda r: ROMAN_ORDER.get(r, 99))

        if drop_suffix:
            if g["qty"] is not None:
                line_text = f"{g['qty']}x {g['base']}"
            else:
                line_text = g["base"]
        else:
            roman_part = ", ".join(romans_sorted)
            if g["qty"] is not None:
                line_text = f"{g['qty']}x {g['base']} {roman_part}"
            else:
                line_text = f"{g['base']} {roman_part}"

        out_entries.append((g["first_index"], line_text))

    out_entries.extend(passthrough)
    out_entries.sort(key=lambda t: t[0])

    return [line for _, line in out_entries]


def condense_numeric_suffixes_for_crafting(lines: list[str]) -> list[str]:
    """
    Extra compaction pass for the 'Used for Crafting/Upgrading:' section.

    Lines that differ only by a trailing integer on the item name
    (e.g. 'Vulcano 1', 'Vulcano 3') are grouped into a single base name.
    """
    groups: dict[str, dict] = {}
    passthrough: list[tuple[int, str]] = []

    for idx, original in enumerate(lines):
        line = original.strip()
        if not line:
            continue

        m_qty = re.match(r"^\s*(\d+)\s*[x×]\s*(.+)$", line)
        if m_qty:
            qty_str = m_qty.group(1)
            rest = m_qty.group(2).strip()
            try:
                qty = int(qty_str)
            except (ValueError, TypeError):
                qty = 0
        else:
            qty = 0
            rest = line

        m_suffix = re.match(r"^(.*\D)(\d+)$", rest)
        if not m_suffix:
            passthrough.append((idx, original))
            continue

        base = m_suffix.group(1).strip()
        number = int(m_suffix.group(2))

        key = base.lower()
        g = groups.get(key)
        if g is None:
            g = {"base": base, "entries": [], "first_index": idx}
            groups[key] = g
        g["entries"].append({"qty": qty, "suffix": number})
        g["first_index"] = min(g["first_index"], idx)

    out_entries: list[tuple[int, str]] = []

    for g in groups.values():
        base = g["base"]
        entries = g["entries"]
        qties = [e["qty"] for e in entries if e["qty"] > 0]

        if qties:
            mn = min(qties)
            mx = max(qties)
            if mn == mx:
                line_text = f"{mn}x {base}"
            else:
                line_text = f"{mn}-{mx}x {base}"
        else:
            line_text = base

        out_entries.append((g["first_index"], line_text))

    out_entries.extend(passthrough)
    out_entries.sort(key=lambda t: t[0])

    result: list[str] = []
    seen: set[str] = set()
    for _, line in out_entries:
        if line in seen:
            continue
        seen.add(line)
        result.append(line)

    return result


def parse_workshop_requirements(row):
    raw = safe_str(row.get("Workshop Requirement", "")).strip()
    if not raw or raw == "[]":
        return []

    try:
        data = json.loads(raw)
    except Exception:
        return [raw]

    lines = []
    for entry in data:
        if not entry:
            continue
        if len(entry) == 3:
            station, level, count = entry
            lines.append(f"{count}x {station} - Level {level}")
        elif len(entry) == 2:
            station, count = entry
            lines.append(f"{count}x {station}")
        else:
            lines.append(" ".join(str(x) for x in entry))
    return lines


def parse_keep_for_quests_workshop(row):
    raw = safe_str(row.get("Keep for Quests/Workshop", "")).strip()
    if not raw:
        return []

    s = " ".join(raw.split())

    bullets = []

    def normalize_item_name(name: str) -> str:
        name_lower = name.lower()
        if name_lower.startswith("expedition"):
            return "Expedition"
        if name_lower.startswith("scrappy"):
            return "Scrappy"
        return name

    pattern = re.compile(r"(\d+)\s*[x×]\s*(Expedition[s]?|Scrappy)", re.IGNORECASE)
    for count, item in pattern.findall(s):
        normalized = normalize_item_name(item)
        bullets.append(f"{count}x {normalized}")

    if not bullets:
        pattern2 = re.compile(r"(\d+)\s+(Expedition[s]?|Scrappy)", re.IGNORECASE)
        for count, item in pattern2.findall(s):
            normalized = normalize_item_name(item)
            bullets.append(f"{count}x {normalized}")

    if not bullets:
        if re.search(r"Expedition", s, re.IGNORECASE):
            bullets.append("Expedition")
        if re.search(r"Scrappy", s, re.IGNORECASE):
            bullets.append("Scrappy")

    return bullets


def parse_quest_usage(row):
    raw = safe_str(row.get("Quest Usage", "")).strip()
    if not raw or raw == "[]":
        return []

    try:
        data = json.loads(raw)
    except Exception:
        return [raw]

    lines = []
    for entry in data:
        if not entry:
            continue
        if len(entry) == 2:
            count, quest_name = entry
            lines.append(f"{count}x Quest - {quest_name}")
        else:
            lines.append(" ".join(str(x) for x in entry))
    return lines


ROI_REL = (0.06, 0.04, 0.94, 0.92)

DETECTION_INTERVAL = 0.10
MISSING_FRAMES_BEFORE_HIDE = 2

REF_W = 1920
REF_H = 1080

NAME_REF_X = 15
NAME_REF_Y = 54
NAME_REF_W = 340
NAME_REF_H = 90

NAME_REF_X2 = 15
NAME_REF_Y2 = 4
NAME_REF_W2 = 340
NAME_REF_H2 = 90

HELPER_GAP_X_REF = 4
HELPER_GAP_Y_REF = 46

cv2.setUseOptimized(True)
cv2.setNumThreads(0)

COMPACT_MAX_WIDTH = 260
COMPACT_PADDING = 14
COMPACT_LINE_GAP = 8

OCR_MIN_INTERVAL = 0.35
LAST_OCR_TIME = 0.0

OCR_API = None

ocr_task_queue = queue.Queue(maxsize=4)
ocr_result_queue = queue.Queue()

HELPER_SCREEN_RECT = None


def init_ocr():
    global OCR_API
    if OCR_API is None:
        OCR_API = PyTessBaseAPI(
            path=TESSDATA_PATH,
            lang="eng",
            psm=PSM.SINGLE_BLOCK,
        )
        OCR_API.SetVariable(
            "tessedit_char_whitelist",
            "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789 ",
        )
        OCR_API.SetVariable("load_system_dawg", "F")
        OCR_API.SetVariable("load_freq_dawg", "F")


def find_tooltip_panel_by_color(
        frame_bgr, min_area=30000, min_fill_ratio=0.80, max_vertices=6
):
    global HELPER_SCREEN_RECT

    hsv = cv2.cvtColor(frame_bgr, cv2.COLOR_BGR2HSV)

    lower = np.array([10, 5, 200], dtype=np.uint8)
    upper = np.array([30, 80, 255], dtype=np.uint8)

    mask = cv2.inRange(hsv, lower, upper)

    if HELPER_SCREEN_RECT is not None:
        hx1, hy1, hx2, hy2 = HELPER_SCREEN_RECT
        h, w = mask.shape[:2]

        hx1 = max(0, min(w - 1, hx1))
        hx2 = max(0, min(w, hx2))
        hy1 = max(0, min(h - 1, hy1))
        hy2 = max(0, min(h, hy2))

        if hx2 > hx1 and hy2 > hy1:
            cv2.rectangle(mask, (hx1, hy1), (hx2, hy2), 0, thickness=-1)

    kernel = np.ones((12, 12), np.uint8)
    mask = cv2.morphologyEx(mask, cv2.MORPH_OPEN, kernel)

    contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    if not contours:
        return None

    candidates = []

    for c in contours:
        x, y, w, h = cv2.boundingRect(c)
        if w <= 0 or h <= 0:
            continue

        area = cv2.contourArea(c)
        if area < min_area:
            continue

        rect_area = float(w * h)
        fill_ratio = area / rect_area
        if fill_ratio < min_fill_ratio:
            continue

        peri = cv2.arcLength(c, True)
        approx = cv2.approxPolyDP(c, 0.03 * peri, True)
        if len(approx) > max_vertices:
            continue

        aspect = w / float(h)
        if aspect < 0.4 or aspect > 1.8:
            continue

        candidates.append((x, y, x + w, y + h, area, fill_ratio))

    if not candidates:
        return None

    candidates.sort(key=lambda b: (b[0], -(b[4] * b[5])))

    x1, y1, x2, y2, _, _ = candidates[0]
    return (x1, y1, x2, y2)


def _crop_name_region_from_panel_generic(
        frame_bgr, panel_box, ref_x, ref_y, ref_w, ref_h
):
    if frame_bgr is None or frame_bgr.size == 0 or panel_box is None:
        return None

    px1, py1, px2, py2 = panel_box

    tooltip = frame_bgr[py1:py2, px1:px2]

    th, tw = frame_bgr.shape[:2]
    if th == 0 or tw == 0:
        return None

    scale_x = tw / float(REF_W)
    scale_y = th / float(REF_H)

    nx1 = int(ref_x * scale_x)
    ny1 = int(ref_y * scale_y)
    nw = int(ref_w * scale_x)
    nh = int(ref_h * scale_y)

    nx2 = nx1 + nw
    ny2 = ny1 + nh

    nx1 = max(0, min(nx1, tw - 1))
    ny1 = max(0, min(ny1, th - 1))
    nx2 = max(nx1 + 1, min(nx2, tw))
    ny2 = max(ny1 + 1, min(ny2, th))

    return tooltip[ny1:ny2, nx1:nx2]


def crop_name_region_from_panel(frame_bgr, panel_box):
    return _crop_name_region_from_panel_generic(
        frame_bgr, panel_box, NAME_REF_X, NAME_REF_Y, NAME_REF_W, NAME_REF_H
    )


def crop_name_region_from_panel_alt(frame_bgr, panel_box):
    return _crop_name_region_from_panel_generic(
        frame_bgr, panel_box, NAME_REF_X2, NAME_REF_Y2, NAME_REF_W2, NAME_REF_H2
    )


def convert_trailing_roman_numeral(name: str) -> str:
    if not name:
        return name

    s = name.rstrip()
    if not s:
        return name

    if len(s.split()) < 2:
        return name

    patterns = [
        ("IV", "4"),
        ("III", "3"),
        ("II", "2"),
        ("I", "1"),
    ]

    for roman, digit in patterns:
        L = len(roman)
        if len(s) < L:
            continue

        tail_original = s[-L:]

        tail_norm = tail_original.upper()
        tail_norm = tail_norm.replace("|", "I").replace("L", "I")

        if tail_norm == roman:
            new_s = s[:-L] + digit
            return new_s

    return name


def ocr_item_lines(name_roi_bgr) -> list[str]:
    """
    OCR the ROI and return a list of cleaned lines (top to bottom).
    Each line keeps its own content so we can match individually
    against the item DB.
    """
    if name_roi_bgr is None or name_roi_bgr.size == 0:
        return []

    init_ocr()
    global OCR_API

    # Preprocess (same as before)
    gray = cv2.cvtColor(name_roi_bgr, cv2.COLOR_BGR2GRAY)

    target_h = 40
    h, w = gray.shape[:2]
    if h > target_h:
        scale = target_h / float(h)
        gray = cv2.resize(gray, (int(w * scale), target_h), interpolation=cv2.INTER_AREA)

    pil_img = Image.fromarray(gray)

    OCR_API.SetImage(pil_img)
    text = OCR_API.GetUTF8Text() or ""

    lines: list[str] = []
    for raw_line in text.splitlines():

        # collapse internal whitespace and strip
        line = " ".join(raw_line.split())
        if not line:
            continue

        # fix trailing roman numerals -> digits
        line = convert_trailing_roman_numeral(line)

        lines.append(line)

    return lines


def ocr_item_name(name_roi_bgr) -> str:
    """
    Backward-compatible helper:
    - OCR all lines
    - Try to find a DB match for each line, in order
    - Return the line that matches, or the longest line as a fallback
    """
    candidates = ocr_item_lines(name_roi_bgr)
    if not candidates:
        return ""

    # Try each line against the DB, in order
    for ln in candidates:
        if find_item_row_by_name(ln) is not None:
            return ln

    # Fallback: just return the longest line
    return max(candidates, key=len, default="")


def compute_name_roi_hash(name_roi_bgr, diff_threshold=3.0, _cache={}):
    if name_roi_bgr is None or name_roi_bgr.size == 0:
        return None

    try:
        small = cv2.resize(name_roi_bgr, (64, 16), interpolation=cv2.INTER_AREA)
        gray = cv2.cvtColor(small, cv2.COLOR_BGR2GRAY)
    except Exception:
        return None

    prev_gray = _cache.get("prev_gray")
    if prev_gray is not None and prev_gray.shape == gray.shape:
        diff = np.mean(np.abs(gray.astype(np.int16) - prev_gray.astype(np.int16)))
        if diff < diff_threshold:
            return _cache.get("prev_hash")

    h = gray.tobytes()
    _cache["prev_gray"] = gray
    _cache["prev_hash"] = h
    return h


def set_current_thread_lowest_priority():
    if os.name != "nt":
        return
    try:
        THREAD_PRIORITY_LOWEST = -2
        kernel32 = ctypes.windll.kernel32
        thread_handle = kernel32.GetCurrentThread()
        kernel32.SetThreadPriority(thread_handle, THREAD_PRIORITY_LOWEST)
    except Exception as e:
        print("Could not set thread priority:", e)


def ocr_db_worker():
    global OCR_API
    try:
        init_ocr()
        set_current_thread_lowest_priority()

        while True:
            task = ocr_task_queue.get()
            if task is None:
                ocr_task_queue.task_done()
                break

            task_id = task.get("task_id")
            roi_primary = task.get("roi_primary")
            roi_secondary = task.get("roi_secondary")
            panel_box = task.get("panel_box")

            try:
                name = None
                row = None
                used_secondary = False

                if roi_primary is not None:
                    # Get all OCR lines, try each one against the DB
                    primary_lines = ocr_item_lines(roi_primary)
                    for ln in primary_lines:
                        r = find_item_row_by_name(ln)
                        if r is not None:
                            name = ln
                            row = r
                            break

                if (row is None) and (roi_secondary is not None):
                    # Fallback: try the secondary ROI (alt name area)
                    secondary_lines = ocr_item_lines(roi_secondary)
                    for ln in secondary_lines:
                        r = find_item_row_by_name(ln)
                        if r is not None:
                            name = ln
                            row = r
                            used_secondary = True
                            break

                ocr_result_queue.put(
                    {
                        "task_id": task_id,
                        "name": name,
                        "row": row,
                        "panel_box": panel_box,
                        "secondary_used": used_secondary,
                    }
                )
            except Exception as e:
                ocr_result_queue.put(
                    {
                        "task_id": task_id,
                        "name": None,
                        "row": None,
                        "panel_box": panel_box,
                        "secondary_used": False,
                        "error": str(e),
                    }
                )
            finally:
                ocr_task_queue.task_done()
    finally:
        if OCR_API is not None:
            OCR_API.End()
            OCR_API = None


def start_ocr_worker():
    t = threading.Thread(target=ocr_db_worker, daemon=True)
    t.start()
    return t


TOOLTIP_ROOT = None
TOOLTIP_LABEL = None
TOOLTIP_PHOTO = None
SCREEN_W = 0
SCREEN_H = 0
TOOLTIP_VISIBLE = False
TOOLTIP_CACHE_KEY = None

TOOLTIP_IMAGE_CACHE = {}

LAST_SHOWN_ROW = None
LAST_SHOWN_PANEL_BOX = None

# offsets of the game monitor in the virtual desktop
MONITOR_LEFT = 0
MONITOR_TOP = 0


def init_overlay_window():
    global TOOLTIP_ROOT, TOOLTIP_LABEL, SCREEN_W, SCREEN_H, TOOLTIP_VISIBLE
    if TOOLTIP_ROOT is not None:
        return

    TRANSP_COLOR = "#f9eedf"

    TOOLTIP_ROOT = tk.Tk()
    TOOLTIP_ROOT.overrideredirect(True)
    TOOLTIP_ROOT.attributes("-topmost", True)
    TOOLTIP_ROOT.attributes("-alpha", TOOLTIP_ALPHA)

    TOOLTIP_ROOT.config(bg=TRANSP_COLOR)

    try:
        TOOLTIP_ROOT.attributes("-transparentcolor", TRANSP_COLOR)
    except tk.TclError:
        pass

    TOOLTIP_LABEL = tk.Label(TOOLTIP_ROOT, bd=0, bg=TRANSP_COLOR)
    TOOLTIP_LABEL.pack()

    # default: use actual screen size; will be overridden with game monitor
    SCREEN_W = TOOLTIP_ROOT.winfo_screenwidth()
    SCREEN_H = TOOLTIP_ROOT.winfo_screenheight()

    TOOLTIP_ROOT.withdraw()
    TOOLTIP_VISIBLE = False


def parse_color_hex(value: str, default_rgba: tuple[int, int, int, int]) -> tuple[int, int, int, int]:
    """
    Parse '#RRGGBB' or '#RRGGBBAA' into (r, g, b, a).

    - If the string is 6 hex digits (no alpha), we use alpha=255 (fully opaque).
    - If the string is 8 hex digits, the last two are used as alpha.
    - On any error, we fall back to default_rgba.
    """
    s = safe_str(value).strip()
    if not s:
        return default_rgba
    if s.startswith("#"):
        s = s[1:]

    if len(s) not in (6, 8):
        return default_rgba

    try:
        r = int(s[0:2], 16)
        g = int(s[2:4], 16)
        b = int(s[4:6], 16)

        if len(s) == 8:
            a = int(s[6:8], 16)
        else:
            # IMPORTANT CHANGE: 6-digit colors are now fully opaque.
            a = 255

        r = max(0, min(255, r))
        g = max(0, min(255, g))
        b = max(0, min(255, b))
        a = max(0, min(255, a))

        return (r, g, b, a)
    except ValueError:
        return default_rgba


def create_helper_tooltip_image(
        row, detected_name, percent_in_second_column=False
):
    padding = COMPACT_PADDING
    line_gap = COMPACT_LINE_GAP
    indent = 14
    column_gap = 10

    BG_COLOR = (0, 0, 0, 255)

    PANEL_COLOR = parse_color_hex(
        SETTINGS.get(
            "tooltip_panel_color",
            DEFAULT_SETTINGS.get("tooltip_panel_color", "#3737370f"),
        ),
        (55, 55, 55, 15),
    )
    TEXT_PRIMARY = parse_color_hex(
        SETTINGS.get(
            "tooltip_text_primary_color",
            DEFAULT_SETTINGS.get("tooltip_text_primary_color", "#141414ff"),
        ),
        (20, 20, 20, 255),
    )
    TEXT_SECONDARY = parse_color_hex(
        SETTINGS.get(
            "tooltip_text_secondary_color",
            DEFAULT_SETTINGS.get("tooltip_text_secondary_color", "#505050ff"),
        ),
        (80, 80, 80, 255),
    )

    COLOR_KEEP = parse_color_hex(
        SETTINGS.get(
            "tooltip_keep_color",
            DEFAULT_SETTINGS.get("tooltip_keep_color", "#ff2828ff"),
        ),
        (255, 40, 40, 255),
    )
    COLOR_RECYCLE = parse_color_hex(
        SETTINGS.get(
            "tooltip_recycle_color",
            DEFAULT_SETTINGS.get("tooltip_recycle_color", "#28ffffff"),
        ),
        (40, 255, 255, 255),
    )
    COLOR_SELL = parse_color_hex(
        SETTINGS.get(
            "tooltip_sell_color",
            DEFAULT_SETTINGS.get("tooltip_sell_color", "#28ff28ff"),
        ),
        (40, 255, 40, 255),
    )

    show_rr_crafting = bool(SETTINGS.get("show_rr_and_crafting", True))

    # font sizes
    try:
        base_size = SETTINGS.get(
            "tooltip_font_size", DEFAULT_SETTINGS.get("tooltip_font_size", 14)
        )
        try:
            base_size = int(base_size)
        except (TypeError, ValueError):
            base_size = DEFAULT_SETTINGS.get("tooltip_font_size", 14)

        base_size = max(10, min(32, base_size))

        font_title = ImageFont.truetype("arialbd.ttf", 17)
        font_label = ImageFont.truetype("arialbd.ttf", base_size)
        font_body = ImageFont.truetype("arialbd.ttf", base_size)
    except Exception:
        font_title = ImageFont.load_default()
        font_label = font_title
        font_body = font_title

    measure_img = Image.new("RGB", (1, 1))
    measure_draw = ImageDraw.Draw(measure_img)

    def text_h(font, txt):
        try:
            bbox = font.getbbox(txt)
            return bbox[3] - bbox[1]
        except Exception:
            return font.getsize(txt)[1]

    def text_w(font, txt):
        try:
            return measure_draw.textlength(txt, font=font)
        except Exception:
            return font.getsize(txt)[0]

    def split_items_list(s: str):
        s = (s or "").strip()
        if not s:
            return []

        pattern = re.compile(r"\d+\s*[x×]\s+")
        matches = list(pattern.finditer(s))

        if matches:
            parts = []
            for i, m in enumerate(matches):
                start = m.start()
                end = matches[i + 1].start() if i + 1 < len(matches) else len(s)
                part = s[start:end].strip().strip(",")
                if part:
                    parts.append(part)
            return parts

        parts = [p.strip() for p in s.split(",") if p.strip()]
        return parts if parts else [s]

    # basic item + verdict info
    name_to_show = safe_str(row.get("Name") if row else detected_name)
    if not name_to_show:
        name_to_show = detected_name or "Unknown Item"

    name_key = safe_str(row.get("Name") if row else detected_name).strip()
    base_verdict = safe_str(row.get("Verdict", "") if row else "").upper()
    override_verdict = USER_VERDICTS.get(name_key)
    override_verdict_up = safe_str(override_verdict).upper() if override_verdict is not None else ""

    verdict_raw = get_effective_verdict(row, detected_name)
    verdict = verdict_raw if verdict_raw and verdict_raw != "UNKNOWN" else "Unknown"

    is_my_suggestion = False
    if override_verdict_up:
        if not base_verdict or override_verdict_up != base_verdict:
            if safe_str(verdict_raw).upper() == override_verdict_up:
                is_my_suggestion = True

    verdict_col = TEXT_SECONDARY
    if verdict_raw == "KEEP":
        verdict_col = COLOR_KEEP
    elif verdict_raw == "RECYCLE":
        verdict_col = COLOR_RECYCLE
    elif verdict_raw == "SELL":
        verdict_col = COLOR_SELL

    # Needed for tasks
    needed_bullets = []
    if row:
        needed_bullets.extend(parse_keep_for_quests_workshop(row))
        needed_bullets.extend(parse_workshop_requirements(row))
        needed_bullets.extend(parse_quest_usage(row))

    if needed_bullets:
        needed_lines = needed_bullets
    else:
        needed_lines = ["No uses known"]

    # recycle / salvage / % / sell price
    recycles_to = safe_str(row.get("Recycles To", "") if row else "")
    salvage_to = safe_str(row.get("Salvages To", "") if row else "")

    raw_rec_gain = format_percentage(
        row.get("Recycle Value Gain %") if row else None
    )
    raw_sell_gain = format_percentage(
        row.get("Sell Value Gain %") if row else None
    )

    rec_gain_text = "-" if raw_rec_gain == "N/A" else raw_rec_gain
    sell_gain_text = "-" if raw_sell_gain == "N/A" else raw_sell_gain

    if row:
        sell_price_val = (
                row.get("Sell Price")
                or row.get("Sell Value")
                or row.get("Sell Price (Base)")
        )
        try:
            sell_price_val = int(sell_price_val)
        except Exception:
            pass
    else:
        sell_price_val = None

    sell_price_text = "-"
    if sell_price_val is not None:
        sp = safe_str(sell_price_val).strip()
        if sp:
            sell_price_text = sp

    rec_items = split_items_list(recycles_to) if recycles_to else []
    sal_items = split_items_list(salvage_to) if salvage_to else []

    rec_lines = []
    if rec_items:
        rec_lines.extend(rec_items)
    elif recycles_to and "cannot" in recycles_to.lower():
        rec_lines.append("Cannot be recycled")
    else:
        rec_lines.append("Cannot be recycled")

    sal_lines = []
    if sal_items:
        sal_lines.extend(sal_items)
    elif salvage_to and "cannot" in salvage_to.lower():
        sal_lines.append("Cannot be salvaged")
    else:
        sal_lines.append("Cannot be salvaged")

    # Reverse Recycle
    if show_rr_crafting and row:
        rr_list = parse_reverse_recycle(row)
        rr_list = condense_roman_variants(rr_list)
    else:
        rr_list = []

    has_rr = bool(rr_list)

    MAX_RR_LINES = 14
    if has_rr:
        if len(rr_list) > MAX_RR_LINES:
            display_rr_list = rr_list[:MAX_RR_LINES]
            hidden_rr = len(rr_list) - MAX_RR_LINES
            display_rr_list.append(f"+{hidden_rr} more items...")
        else:
            display_rr_list = rr_list
            hidden_rr = 0
    else:
        display_rr_list = []
        hidden_rr = 0

    rr_display_line_count = len(display_rr_list)

    # Crafting
    if show_rr_crafting and row:
        craft_entries = parse_crafting(row, return_meta=True)
    else:
        craft_entries = []

    if craft_entries:
        weapon_lines = [
            e["line"]
            for e in craft_entries
            if (e.get("category") or "").lower() == "weapon"
        ]
        other_lines = [
            e["line"]
            for e in craft_entries
            if (e.get("category") or "").lower() != "weapon"
        ]

        if weapon_lines:
            weapon_lines = condense_roman_variants(
                weapon_lines, drop_suffix=True
            )
            weapon_lines = condense_numeric_suffixes_for_crafting(weapon_lines)

        craft_list = weapon_lines + other_lines
    else:
        craft_list = []

    has_craft = bool(craft_list)

    MAX_CRAFT_LINES = 14
    if has_craft:
        if len(craft_list) > MAX_CRAFT_LINES:
            display_craft_list = craft_list[:MAX_CRAFT_LINES]
            hidden_craft = len(craft_list) - MAX_CRAFT_LINES
            display_craft_list.append(f"+{hidden_craft} more items...")
        else:
            display_craft_list = craft_list
            hidden_craft = 0
    else:
        display_craft_list = []
        hidden_craft = 0

    if display_rr_list or display_craft_list:
        percent_in_second_column = False

    items = []
    header_max_width = 0
    left_col_max_width = 0
    rr_col_max_width = 0
    craft_col_max_width = 0
    needed_label_y = None

    y = COMPACT_PADDING
    title_y = y

    header_max_width = max(header_max_width, text_w(font_title, name_to_show))
    y = COMPACT_LINE_GAP * 2

    needed_label = "Needed for Tasks:"
    needed_label_y = y
    header_max_width = max(header_max_width, text_w(font_label, needed_label))
    items.append(
        dict(
            kind="header",
            x_off=0,
            y=needed_label_y,
            text=needed_label,
            font=font_label,
            fill=TEXT_SECONDARY,
        )
    )
    y += text_h(font_label, needed_label) + COMPACT_LINE_GAP

    for line in needed_lines:
        line_w = indent + text_w(font_body, line)
        header_max_width = max(header_max_width, line_w)
        items.append(
            dict(
                kind="header",
                x_off=indent,
                y=y,
                text=line,
                font=font_body,
                fill=TEXT_PRIMARY,
            )
        )
        y += text_h(font_body, line) + COMPACT_LINE_GAP

    y += COMPACT_LINE_GAP

    if needed_bullets:
        verdict_label = (
            "My Suggested action: (When Tasks done)"
            if is_my_suggestion
            else "Suggested action: (When Tasks done)"
        )
    else:
        verdict_label = "My Suggested action:" if is_my_suggestion else "Suggested action:"

    label_w = text_w(font_label, verdict_label)
    verdict_w = text_w(font_label, verdict)

    header_max_width = max(
        header_max_width,
        label_w,
        indent + verdict_w,
        )

    items.append(
        dict(
            kind="header",
            x_off=0,
            y=y,
            text=verdict_label,
            font=font_label,
            fill=TEXT_SECONDARY,
        )
    )
    y += text_h(font_label, verdict_label) + COMPACT_LINE_GAP

    items.append(
        dict(
            kind="header",
            x_off=indent,
            y=y,
            text=verdict,
            font=font_label,
            fill=verdict_col,
        )
    )
    y += text_h(font_label, verdict) + COMPACT_LINE_GAP * 2

    columns_top_y = y
    y_left = columns_top_y
    y_rr = columns_top_y
    y_craft = columns_top_y

    if needed_label_y is None:
        needed_label_y = columns_top_y

    if display_rr_list or display_craft_list:
        y_rr = title_y
        y_craft = title_y

        if display_rr_list:
            label = "Reverse Recycle:"
            rr_col_max_width = max(rr_col_max_width, text_w(font_label, label))
            items.append(
                dict(
                    kind="right",
                    x_off=0,
                    y=y_rr,
                    text=label,
                    font=font_label,
                    fill=TEXT_SECONDARY,
                )
            )
            y_rr += text_h(font_label, label) + COMPACT_LINE_GAP

            for line in display_rr_list:
                w_line = indent + text_w(font_body, line)
                rr_col_max_width = max(rr_col_max_width, w_line)
                items.append(
                    dict(
                        kind="right",
                        x_off=indent,
                        y=y_rr,
                        text=line,
                        font=font_body,
                        fill=TEXT_PRIMARY,
                    )
                )
                y_rr += text_h(font_body, line) + COMPACT_LINE_GAP

        if display_craft_list:
            if display_rr_list and rr_display_line_count >= 8:
                label = "Used for Crafting/Upgrading:"
                craft_col_max_width = max(craft_col_max_width, text_w(font_label, label))
                items.append(
                    dict(
                        kind="craft",
                        x_off=0,
                        y=y_craft,
                        text=label,
                        font=font_label,
                        fill=TEXT_SECONDARY,
                    )
                )
                y_craft += text_h(font_label, label) + COMPACT_LINE_GAP

                for line in display_craft_list:
                    w_line = indent + text_w(font_body, line)
                    craft_col_max_width = max(craft_col_max_width, w_line)
                    items.append(
                        dict(
                            kind="craft",
                            x_off=indent,
                            y=y_craft,
                            text=line,
                            font=font_body,
                            fill=TEXT_PRIMARY,
                        )
                    )
                    y_craft += text_h(font_body, line) + COMPACT_LINE_GAP
            else:
                if display_rr_list:
                    y_rr += COMPACT_LINE_GAP
                label = "Used for Crafting/Upgrading:"
                rr_col_max_width = max(rr_col_max_width, text_w(font_label, label))
                items.append(
                    dict(
                        kind="right",
                        x_off=0,
                        y=y_rr,
                        text=label,
                        font=font_label,
                        fill=TEXT_SECONDARY,
                    )
                )
                y_rr += text_h(font_label, label) + COMPACT_LINE_GAP

                for line in display_craft_list:
                    w_line = indent + text_w(font_body, line)
                    rr_col_max_width = max(rr_col_max_width, w_line)
                    items.append(
                        dict(
                            kind="right",
                            x_off=indent,
                            y=y_rr,
                            text=line,
                            font=font_body,
                            fill=TEXT_PRIMARY,
                        )
                    )
                    y_rr += text_h(font_body, line) + COMPACT_LINE_GAP

    elif percent_in_second_column:
        y_rr = needed_label_y

        label = "Recycle Value Gain (vs Salvage):"
        rr_col_max_width = max(rr_col_max_width, text_w(font_label, label))
        items.append(
            dict(
                kind="right",
                x_off=0,
                y=y_rr,
                text=label,
                font=font_label,
                fill=TEXT_SECONDARY,
            )
        )
        y_rr += text_h(font_label, label) + COMPACT_LINE_GAP

        w_line = indent + text_w(font_body, rec_gain_text)
        rr_col_max_width = max(rr_col_max_width, w_line)
        items.append(
            dict(
                kind="right",
                x_off=indent,
                y=y_rr,
                text=rec_gain_text,
                font=font_body,
                fill=TEXT_PRIMARY,
            )
        )
        y_rr += text_h(font_body, rec_gain_text) + COMPACT_LINE_GAP * 2

        label = "Sell Value Gain (vs Recycle):"
        rr_col_max_width = max(rr_col_max_width, text_w(font_label, label))
        items.append(
            dict(
                kind="right",
                x_off=0,
                y=y_rr,
                text=label,
                font=font_label,
                fill=TEXT_SECONDARY,
            )
        )
        y_rr += text_h(font_label, label) + COMPACT_LINE_GAP

        w_line = indent + text_w(font_body, sell_gain_text)
        rr_col_max_width = max(rr_col_max_width, w_line)
        items.append(
            dict(
                kind="right",
                x_off=indent,
                y=y_rr,
                text=sell_gain_text,
                font=font_body,
                fill=TEXT_PRIMARY,
            )
        )
        y_rr += text_h(font_body, sell_gain_text) + COMPACT_LINE_GAP * 2

    # Left column
    label = "Recycle:"
    left_col_max_width = max(left_col_max_width, text_w(font_label, label))
    items.append(
        dict(
            kind="left",
            x_off=0,
            y=y_left,
            text=label,
            font=font_label,
            fill=TEXT_SECONDARY,
        )
    )
    y_left += text_h(font_label, label) + COMPACT_LINE_GAP

    for line in rec_lines:
        w_line = indent + text_w(font_body, line)
        left_col_max_width = max(left_col_max_width, w_line)
        items.append(
            dict(
                kind="left",
                x_off=indent,
                y=y_left,
                text=line,
                font=font_body,
                fill=TEXT_PRIMARY,
            )
        )
        y_left += text_h(font_body, line) + COMPACT_LINE_GAP
    y_left += COMPACT_LINE_GAP

    label = "Salvage:"
    left_col_max_width = max(left_col_max_width, text_w(font_label, label))
    items.append(
        dict(
            kind="left",
            x_off=0,
            y=y_left,
            text=label,
            font=font_label,
            fill=TEXT_SECONDARY,
        )
    )
    y_left += text_h(font_label, label) + COMPACT_LINE_GAP

    for line in sal_lines:
        w_line = indent + text_w(font_body, line)
        left_col_max_width = max(left_col_max_width, w_line)
        items.append(
            dict(
                kind="left",
                x_off=indent,
                y=y_left,
                text=line,
                font=font_body,
                fill=TEXT_PRIMARY,
            )
        )
        y_left += text_h(font_body, line) + COMPACT_LINE_GAP
    y_left += COMPACT_LINE_GAP

    if not percent_in_second_column:
        label = "Recycle Value Gain (vs Salvage):"
        left_col_max_width = max(left_col_max_width, text_w(font_label, label))
        items.append(
            dict(
                kind="left",
                x_off=0,
                y=y_left,
                text=label,
                font=font_label,
                fill=TEXT_SECONDARY,
            )
        )
        y_left += text_h(font_label, label) + COMPACT_LINE_GAP

        w_line = indent + text_w(font_body, rec_gain_text)
        left_col_max_width = max(left_col_max_width, w_line)
        items.append(
            dict(
                kind="left",
                x_off=indent,
                y=y_left,
                text=rec_gain_text,
                font=font_body,
                fill=TEXT_PRIMARY,
            )
        )
        y_left += text_h(font_body, rec_gain_text) + COMPACT_LINE_GAP * 2

        label = "Sell Value Gain (vs Recycle):"
        left_col_max_width = max(left_col_max_width, text_w(font_label, label))
        items.append(
            dict(
                kind="left",
                x_off=0,
                y=y_left,
                text=label,
                font=font_label,
                fill=TEXT_SECONDARY,
            )
        )
        y_left += text_h(font_label, label) + COMPACT_LINE_GAP

        w_line = indent + text_w(font_body, sell_gain_text)
        left_col_max_width = max(left_col_max_width, w_line)
        items.append(
            dict(
                kind="left",
                x_off=indent,
                y=y_left,
                text=sell_gain_text,
                font=font_body,
                fill=TEXT_PRIMARY,
            )
        )
        y_left += text_h(font_body, sell_gain_text) + COMPACT_LINE_GAP * 2

    label = "Sell Price per Item:"
    left_col_max_width = max(left_col_max_width, text_w(font_label, label))
    items.append(
        dict(
            kind="left",
            x_off=0,
            y=y_left,
            text=label,
            font=font_label,
            fill=TEXT_SECONDARY,
        )
    )
    y_left += text_h(font_label, label) + COMPACT_LINE_GAP

    w_line = indent + text_w(font_body, sell_price_text)
    left_col_max_width = max(left_col_max_width, w_line)
    items.append(
        dict(
            kind="left",
            x_off=indent,
            y=y_left,
            text=sell_price_text,
            font=font_body,
            fill=TEXT_PRIMARY,
        )
    )
    y_left += text_h(font_body, sell_price_text) + COMPACT_LINE_GAP * 2

    content_bottom = y_left
    if rr_col_max_width > 0:
        content_bottom = max(content_bottom, y_rr)
    if craft_col_max_width > 0:
        content_bottom = max(content_bottom, y_craft)

    left_col_width = max(header_max_width, left_col_max_width)
    if left_col_width <= 0:
        left_col_width = header_max_width or 50

    content_width = left_col_width
    rr_x_inner = None
    craft_x_inner = None

    if rr_col_max_width > 0:
        rr_x_inner = left_col_width + column_gap
        content_width = max(content_width, rr_x_inner + rr_col_max_width)

    if craft_col_max_width > 0:
        if display_rr_list and rr_display_line_count >= 8:
            craft_x_inner = (
                rr_x_inner + rr_col_max_width + column_gap
                if rr_x_inner is not None
                else (left_col_width + column_gap)
            )
        else:
            craft_x_inner = left_col_width + column_gap
        content_width = max(content_width, craft_x_inner + craft_col_max_width)

    width = int(COMPACT_PADDING + content_width + COMPACT_PADDING)
    used_height = int(content_bottom + COMPACT_PADDING)

    img = Image.new("RGBA", (width, used_height), BG_COLOR)
    draw = ImageDraw.Draw(img)

    radius = 8
    draw.rounded_rectangle(
        (0, 0, width - 1, used_height - 1),
        radius=radius,
        fill=PANEL_COLOR,
    )

    left_x = COMPACT_PADDING
    rr_x = left_x + rr_x_inner if rr_x_inner is not None else None
    craft_x = left_x + craft_x_inner if craft_x_inner is not None else None

    for it in items:
        kind = it["kind"]

        if kind == "right" and rr_x is None:
            continue
        if kind == "craft" and craft_x is None:
            continue

        if kind in ("header", "left"):
            x = left_x + it["x_off"]
        elif kind == "right":
            x = rr_x + it["x_off"]
        elif kind == "craft":
            x = craft_x + it["x_off"]
        else:
            x = left_x + it["x_off"]

        ty = it["y"]
        if 0 <= ty < used_height:
            draw.text((x, ty), it["text"], font=it["font"], fill=it["fill"])

    return img


def get_helper_gaps():
    """
    Return (gap_x, gap_y) scaled from the 1920x1080 reference
    to the current screen size (game monitor).
    """
    w = SCREEN_W or REF_W
    h = SCREEN_H or REF_H

    gap_x = int(round(HELPER_GAP_X_REF * w / REF_W))
    gap_y = int(round(HELPER_GAP_Y_REF * h / REF_H))

    return max(gap_x, 1), max(gap_y, 1)


def get_mouse_position():
    try:
        class POINT(ctypes.Structure):
            _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]

        pt = POINT()
        if ctypes.windll.user32.GetCursorPos(ctypes.byref(pt)):
            return pt.x, pt.y
    except Exception:
        pass
    return None, None


def show_helper_tooltip(
        row, detected_name, global_panel_box, used_secondary: bool = False
):
    """
    Show the helper tooltip next to the detected panel.

    NOTE: panel box and tooltip math is done in *monitor-local* coordinates,
    then converted to global desktop coordinates once for the Tk window.
    """
    global TOOLTIP_PHOTO, TOOLTIP_VISIBLE, TOOLTIP_CACHE_KEY, TOOLTIP_IMAGE_CACHE
    global HELPER_SCREEN_RECT, MONITOR_LEFT, MONITOR_TOP

    if global_panel_box is None:
        hide_helper_tooltip()
        return

    init_overlay_window()

    margin = 10
    min_gap = 4

    gap_x, gap_y = get_helper_gaps()
    side_gap = max(gap_x, min_gap)

    if len(global_panel_box) == 5:
        gx1, gy1, gx2, gy2, _ = global_panel_box
    else:
        gx1, gy1, gx2, gy2 = global_panel_box

    row_name = safe_str(row.get("Name") if row else detected_name)
    effective_verdict = get_effective_verdict(row, detected_name)

    def get_photo(percent_in_second_column_flag: bool):
        key = (row_name, effective_verdict, percent_in_second_column_flag)
        if key in TOOLTIP_IMAGE_CACHE:
            return TOOLTIP_IMAGE_CACHE[key], key

        img = create_helper_tooltip_image(
            row,
            detected_name,
            percent_in_second_column=percent_in_second_column_flag,
        )
        photo = ImageTk.PhotoImage(img)
        TOOLTIP_IMAGE_CACHE[key] = photo
        return photo, key

    test_photo, test_key = get_photo(False)
    test_w = test_photo.width()
    test_h = test_photo.height()

    x_right_for_test = int(gx2 + side_gap)
    right_fits_with_single_column = x_right_for_test + test_w <= SCREEN_W - margin

    if right_fits_with_single_column:
        TOOLTIP_PHOTO = test_photo
        TOOLTIP_CACHE_KEY = test_key
        w, h = test_w, test_h
    else:
        compact_photo, compact_key = get_photo(True)
        TOOLTIP_PHOTO = compact_photo
        TOOLTIP_CACHE_KEY = compact_key
        w = TOOLTIP_PHOTO.width()
        h = TOOLTIP_PHOTO.height()

    TOOLTIP_LABEL.configure(image=TOOLTIP_PHOTO)

    if used_secondary:
        y = int(gy1)
    else:
        y = int(gy1 + gap_y)

    if y < margin:
        y = margin
    if y + h > SCREEN_H - margin:
        y = SCREEN_H - h - margin
    if y < margin:
        y = margin

    x_right = int(gx2 + side_gap)
    x_left = int(gx1 - side_gap - w)

    right_fits = x_right + w <= SCREEN_W - margin

    mx, my = get_mouse_position()
    # convert mouse position from global desktop to monitor-local coordinates
    if mx is not None and my is not None:
        mx -= MONITOR_LEFT
        my -= MONITOR_TOP

    panel_center_x = gx1 + (gx2 - gx1) // 2

    def horizontal_distance(mx_, tx_, tw_):
        if mx_ < tx_:
            return tx_ - mx_
        elif mx_ > tx_ + tw_:
            return mx_ - (tx_ + tw_)
        else:
            return 0

    x = None
    place_left = False

    if right_fits:
        if mx is not None:
            d_right = horizontal_distance(mx, x_right, w)
            d_left = horizontal_distance(mx, x_left, w)
            x = x_right if d_right >= d_left else x_left
        else:
            x = x_right
    else:
        if mx is not None and panel_center_x < mx:
            place_left = True
            x = x_left
        else:
            if mx is not None:
                x = margin
            else:
                preferred_center = gx1 + (gx2 - gx1) // 2
                x = preferred_center - w // 2

    x = max(0, min(SCREEN_W - w, x))
    y = max(0, min(SCREEN_H - h, y))

    if not place_left and not right_fits:
        if mx is not None and my is not None and x <= mx:
            placed = False

            def clamp_to_screen(cx, cy):
                cx = max(margin, min(SCREEN_W - margin - w, cx))
                cy = max(margin, min(SCREEN_H - margin - h, cy))
                return cx, cy

            def try_place(cx, cy):
                cx, cy = clamp_to_screen(cx, cy)
                if cx <= mx <= cx + w and cy <= my <= cy + h:
                    return None
                return cx - 1, cy

            align_x = gx1

            under_y = gy2 + side_gap
            above_y = gy1 - side_gap - h

            fits_under = under_y + h <= SCREEN_H - margin
            fits_above = above_y >= margin

            if fits_under:
                pos = try_place(align_x, under_y)
                if pos is not None:
                    x, y = pos
                    placed = True

            if not placed and fits_above:
                pos = try_place(align_x, above_y)
                if pos is not None:
                    x, y = pos
                    placed = True

            if not placed:
                pos = try_place(align_x, under_y)
                if pos is not None:
                    x, y = pos

    if TOOLTIP_ROOT is not None:
        # convert monitor-local tooltip coords to global desktop coords for Tk
        global_x = int(x + MONITOR_LEFT)
        global_y = int(y + MONITOR_TOP)

        TOOLTIP_ROOT.geometry(f"{w}x{h}+{global_x}+{global_y}")

        # HELPER_SCREEN_RECT is stored in monitor-local coordinates
        HELPER_SCREEN_RECT = (int(x), int(y), int(x) + w, int(y) + h)

        if not TOOLTIP_VISIBLE:
            TOOLTIP_ROOT.deiconify()
            TOOLTIP_VISIBLE = True


def hide_helper_tooltip():
    global TOOLTIP_VISIBLE, HELPER_SCREEN_RECT
    if TOOLTIP_ROOT is not None and TOOLTIP_VISIBLE:
        try:
            TOOLTIP_ROOT.withdraw()
        except tk.TclError:
            pass
        TOOLTIP_VISIBLE = False

    HELPER_SCREEN_RECT = None


def set_low_priority():
    try:
        if os.name == "nt":
            PROCESS_PRIORITY_IDLE = 0x40
            handle = ctypes.windll.kernel32.GetCurrentProcess()
            ctypes.windll.kernel32.SetPriorityClass(handle, PROCESS_PRIORITY_IDLE)
    except Exception as e:
        print("Could not set low priority:", e)


def warm_up_tooltip_engine():
    try:
        init_overlay_window()
        dummy_row = {"Name": "Warmup Item", "Verdict": "KEEP"}
        img = create_helper_tooltip_image(dummy_row, "Warmup Item", False)
        _ = ImageTk.PhotoImage(img)
    except Exception as e:
        print("Warmup failed (non-fatal):", e)


def _keyboard_hotkey_matches(key) -> bool:
    if pynput_keyboard is None:
        return False

    cfg = SETTINGS.get("hotkey") or {}
    if cfg.get("device") != "keyboard":
        return False

    target = (cfg.get("key") or "").lower()
    if not target:
        return False

    try:
        if isinstance(key, pynput_keyboard.KeyCode):
            ch = (key.char or "").lower()
            return ch == target
        elif isinstance(key, pynput_keyboard.Key):
            name = (key.name or "").lower()
            return name == target
    except Exception:
        return False

    return False


def _mouse_hotkey_matches(button) -> bool:
    if pynput_mouse is None:
        return False

    cfg = SETTINGS.get("hotkey") or {}
    if cfg.get("device") != "mouse":
        return False

    target = (cfg.get("key") or "").lower()
    if not target:
        return False

    try:
        name = (button.name or "").lower()
        return name == target
    except Exception:
        return False


def _keyboard_cycle_hotkey_matches(key) -> bool:
    if pynput_keyboard is None:
        return False

    cfg = SETTINGS.get("cycle_hotkey") or {}
    if cfg.get("device") != "keyboard":
        return False

    target = (cfg.get("key") or "").lower()
    if not target:
        return False

    try:
        if isinstance(key, pynput_keyboard.KeyCode):
            ch = (key.char or "").lower()
            return ch == target
        elif isinstance(key, pynput_keyboard.Key):
            name = (key.name or "").lower()
            return name == target
    except Exception:
        return False

    return False


def _mouse_cycle_hotkey_matches(button) -> bool:
    if pynput_mouse is None:
        return False

    cfg = SETTINGS.get("cycle_hotkey") or {}
    if cfg.get("device") != "mouse":
        return False

    target = (cfg.get("key") or "").lower()
    if not target:
        return False

    try:
        name = (button.name or "").lower()
        return name == target
    except Exception:
        return False


def start_hotkey_listeners():
    global HOTKEY_LISTENERS_STARTED, HOTKEY_HELD

    if HOTKEY_LISTENERS_STARTED:
        return

    if not PYNPUT_AVAILABLE:
        print(
            "[helper] pynput not installed; hotkey support is disabled.\n"
            "         Install it with: pip install pynput\n"
            "         or use the 'Always on' setting."
        )
        return

    def on_press(key):
        global HOTKEY_HELD
        try:
            if _keyboard_hotkey_matches(key):
                HOTKEY_HELD = True

            if _keyboard_cycle_hotkey_matches(key):
                cycle_verdict_for_current_item(+1)
        except Exception:
            pass

    def on_release(key):
        global HOTKEY_HELD
        try:
            if _keyboard_hotkey_matches(key):
                HOTKEY_HELD = False
        except Exception:
            pass

    def on_click(x, y, button, pressed):
        global HOTKEY_HELD
        try:
            if _mouse_hotkey_matches(button):
                HOTKEY_HELD = pressed

            if pressed and _mouse_cycle_hotkey_matches(button):
                cycle_verdict_for_current_item(+1)
        except Exception:
            pass

    kb_listener = pynput_keyboard.Listener(on_press=on_press, on_release=on_release)
    ms_listener = pynput_mouse.Listener(on_click=on_click)
    kb_listener.daemon = True
    ms_listener.daemon = True
    kb_listener.start()
    ms_listener.start()

    HOTKEY_LISTENERS_STARTED = True


def get_arc_raiders_monitor(sct, game_title: str = "ARC Raiders"):
    """
    Try to find the monitor that contains the ARC Raiders window.

    - On Windows: enumerate top-level windows, find one whose title contains
      `game_title`, get its rect and select the monitor whose bounds contain
      the window center.
    - On other platforms or if not found: fall back to sct.monitors[1]
      (current behavior).
    """
    # Non-Windows: just use primary monitor as before
    if os.name != "nt":
        return sct.monitors[1]

    user32 = ctypes.windll.user32

    EnumWindows = user32.EnumWindows
    EnumWindowsProc = ctypes.WINFUNCTYPE(
        ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p
    )
    GetWindowTextLengthW = user32.GetWindowTextLengthW
    GetWindowTextW = user32.GetWindowTextW
    IsWindowVisible = user32.IsWindowVisible
    GetWindowRect = user32.GetWindowRect

    target_hwnd = ctypes.c_void_p()

    game_title_lower = game_title.lower()

    def callback(hwnd, lParam):
        nonlocal target_hwnd

        # Only visible windows
        if not IsWindowVisible(hwnd):
            return True

        length = GetWindowTextLengthW(hwnd)
        if length == 0:
            return True

        buf = ctypes.create_unicode_buffer(length + 1)
        GetWindowTextW(hwnd, buf, length + 1)
        title = buf.value.strip()

        if not title:
            return True

        if game_title_lower in title.lower():
            target_hwnd = ctypes.c_void_p(hwnd)
            # Stop enumeration
            return False

        return True

    try:
        EnumWindows(EnumWindowsProc(callback), 0)
    except Exception:
        # In case EnumWindows fails, just fall back
        target_hwnd = ctypes.c_void_p()

    # If we didn't find a matching window, use primary
    if not target_hwnd.value:
        return sct.monitors[1]

    # Get window rect
    rect = (ctypes.c_long * 4)()
    if not GetWindowRect(target_hwnd, rect):
        # If GetWindowRect fails, fall back
        return sct.monitors[1]

    left, top, right, bottom = rect
    win_cx = (left + right) // 2
    win_cy = (top + bottom) // 2

    # mss.monitors[0] is the virtual screen; real monitors start at index 1
    chosen_monitor = None
    for m in sct.monitors[1:]:
        m_left = m.get("left", 0)
        m_top = m.get("top", 0)
        m_width = m.get("width", 0)
        m_height = m.get("height", 0)

        m_right = m_left + m_width
        m_bottom = m_top + m_height

        if m_left <= win_cx < m_right and m_top <= win_cy < m_bottom:
            chosen_monitor = m
            break

    if chosen_monitor is not None:
        return chosen_monitor

    # If the window center didn't fall inside any monitor (weird edge case),
    # keep old behavior
    return sct.monitors[1]


def main_live():
    global LAST_OCR_TIME, LAST_SHOWN_ROW,LAST_SHOWN_PANEL_BOX,SETTINGS,TOOLTIP_NEEDS_REFRESH,TOOLTIP_IMAGE_CACHE,SCREEN_W,SCREEN_H,MONITOR_LEFT,MONITOR_TOP

    sct = mss()
    # Prefer the monitor ARC Raiders is on; fall back to primary
    monitor = get_arc_raiders_monitor(sct)

    # Use game monitor dimensions and offsets for all tooltip math
    MONITOR_LEFT = monitor.get("left", 0)
    MONITOR_TOP = monitor.get("top", 0)
    SCREEN_W = monitor.get("width", SCREEN_W)
    SCREEN_H = monitor.get("height", SCREEN_H)

    last_settings_mtime = None
    try:
        if SETTINGS_PATH.is_file():
            last_settings_mtime = SETTINGS_PATH.stat().st_mtime
    except Exception:
        pass

    last_name = None
    last_row = None
    last_panel_box = None
    last_name_roi_hash = None
    last_used_secondary = False

    missing_frames = 0

    next_task_id = 0
    latest_result_id = -1

    print("Starting live detection. Press Ctrl+C to stop.")
    try:
        while True:
            start = time.time()

            try:
                if SETTINGS_PATH.is_file():
                    mtime = SETTINGS_PATH.stat().st_mtime
                    if last_settings_mtime is None or mtime > last_settings_mtime:
                        last_settings_mtime = mtime
                        refresh_settings()
                        TOOLTIP_IMAGE_CACHE.clear()
                        TOOLTIP_NEEDS_REFRESH = True
            except Exception:
                pass

            always_on = bool(SETTINGS.get("always_on", False))
            gating_active = always_on or HOTKEY_HELD

            panel_box = None

            if gating_active:
                sct_img = sct.grab(monitor)
                # frame_full is in monitor-local coordinates (0..width, 0..height)
                frame_full = np.array(sct_img)[:, :, :3]

                panel_box = find_tooltip_panel_by_color(frame_full)

                if panel_box is not None:
                    missing_frames = 0
                    last_panel_box = panel_box

                    name_roi_primary = crop_name_region_from_panel(
                        frame_full, panel_box
                    )

                    name_roi_secondary = crop_name_region_from_panel_alt(
                        frame_full, panel_box
                    )

                    roi_hash = compute_name_roi_hash(name_roi_primary)

                    if roi_hash is not None and roi_hash != last_name_roi_hash:
                        now = time.time()
                        if now - LAST_OCR_TIME >= OCR_MIN_INTERVAL:
                            last_name_roi_hash = roi_hash
                            LAST_OCR_TIME = now

                            next_task_id += 1
                            task = {
                                "task_id": next_task_id,
                                "roi_primary": name_roi_primary,
                                "roi_secondary": name_roi_secondary,
                                "panel_box": panel_box,
                            }
                            try:
                                ocr_task_queue.put_nowait(task)
                            except queue.Full:
                                pass
                else:
                    missing_frames += 1
                    if missing_frames >= MISSING_FRAMES_BEFORE_HIDE:
                        if last_name is not None or last_row is not None:
                            print("Tooltip lost, hiding helper.")
                        last_name = None
                        last_row = None
                        last_panel_box = None
                        last_name_roi_hash = None
                        last_used_secondary = False
                        hide_helper_tooltip()
                        LAST_SHOWN_ROW = None
                        LAST_SHOWN_PANEL_BOX = None
            else:
                if (
                        last_name is not None
                        or last_row is not None
                        or last_panel_box is not None
                ):
                    last_name = None
                    last_row = None
                    last_panel_box = None
                    last_name_roi_hash = None
                    last_used_secondary = False
                    hide_helper_tooltip()
                    LAST_SHOWN_ROW = None
                    LAST_SHOWN_PANEL_BOX = None
                missing_frames = 0

            tooltip_active = gating_active and (last_panel_box is not None)

            while True:
                try:
                    res = ocr_result_queue.get_nowait()
                except queue.Empty:
                    break

                rid = res.get("task_id", -1)
                name = res.get("name")
                row = res.get("row")
                err = res.get("error")
                secondary_used = bool(res.get("secondary_used", False))

                ocr_result_queue.task_done()

                if rid < latest_result_id:
                    continue
                latest_result_id = rid

                if not tooltip_active:
                    continue

                if err:
                    print(f"OCR worker error on task {rid}: {err}")
                    continue

                if name and row is not None:
                    last_name = name
                    last_row = row
                    last_used_secondary = secondary_used
                    print(
                        f"Detected item (async): {name} -> matched '{last_row.get('Name', '')}' "
                        f"(secondary={secondary_used})"
                    )
                else:
                    if name:
                        print(f"Detected item (async): {name} -> no match in DB")
                    last_name = None
                    last_row = None
                    last_used_secondary = False
                    hide_helper_tooltip()
                    LAST_SHOWN_ROW = None
                    LAST_SHOWN_PANEL_BOX = None

            if tooltip_active and last_panel_box is not None and last_row is not None:
                if (
                        last_row is not LAST_SHOWN_ROW
                        or last_panel_box != LAST_SHOWN_PANEL_BOX
                        or TOOLTIP_NEEDS_REFRESH
                ):
                    x1, y1, x2, y2 = last_panel_box
                    # panel_box is already monitor-local; keep it that way
                    global_panel_box = (x1, y1, x2, y2, 1.0)
                    show_helper_tooltip(
                        last_row,
                        last_name,
                        global_panel_box,
                        used_secondary=last_used_secondary,
                    )
                    LAST_SHOWN_ROW = last_row
                    LAST_SHOWN_PANEL_BOX = last_panel_box
                    TOOLTIP_NEEDS_REFRESH = False
            else:
                if LAST_SHOWN_ROW is not None or LAST_SHOWN_PANEL_BOX is not None:
                    hide_helper_tooltip()
                    LAST_SHOWN_ROW = None
                    LAST_SHOWN_PANEL_BOX = None

            if TOOLTIP_ROOT is not None:
                try:
                    TOOLTIP_ROOT.update_idletasks()
                    TOOLTIP_ROOT.update()
                except tk.TclError:
                    pass

            elapsed = time.time() - start
            if elapsed < DETECTION_INTERVAL:
                time.sleep(DETECTION_INTERVAL - elapsed)

    except KeyboardInterrupt:
        print("Stopping live detection.")
    finally:
        cv2.destroyAllWindows()
        hide_helper_tooltip()
        if TOOLTIP_ROOT is not None:
            try:
                TOOLTIP_ROOT.destroy()
            except tk.TclError:
                pass
        try:
            ocr_task_queue.put_nowait(None)
        except queue.Full:
            pass


def run_helper():
    refresh_settings()
    load_user_verdicts()
    print(f"Loaded settings from {SETTINGS_PATH}: {SETTINGS}")
    if USER_VERDICTS:
        print(f"Loaded {len(USER_VERDICTS)} verdict override(s) from {VERDICTS_PATH}")

    if SETTINGS.get("always_on", False):
        print("Mode: Always on (continuous tooltip detection)")
    else:
        hk = SETTINGS.get("hotkey", {})
        chk = SETTINGS.get("cycle_hotkey", {})
        print(
            f"Mode: Hold-to-show. Hotkey device={hk.get('device', '?')}, "
            f"key={hk.get('key', '?')}"
        )
        print(
            f"Cycle suggested action hotkey: device={chk.get('device', '?')}, "
            f"key={chk.get('key', '?')}"
        )

    set_low_priority()
    warm_up_tooltip_engine()
    _ = start_ocr_worker()
    start_hotkey_listeners()

    main_live()


def run_settings():
    run_settings_ui()


if __name__ == "__main__":
    if "--run-helper" in sys.argv:
        run_helper()
    else:
        run_settings_ui()
