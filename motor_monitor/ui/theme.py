"""화면 색상/글꼴/이미지. 원본 XAML 의 색상값을 그대로 옮겼습니다.

스펙시트(docs/SPEC_SHEET.md) '5. 색상/스타일' 항목과 1:1 대응합니다.
"""
from __future__ import annotations

import os
import sys
import tkinter as tk
import tkinter.font as tkfont
from typing import Dict

ASSET_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "assets")
IS_WINDOWS = sys.platform.startswith("win")

# 제목표시줄 없는 대화상자(원본 WindowStyle=None). Windows 이외에서는 키보드 포커스 문제로 끔.
BORDERLESS_DIALOGS = IS_WINDOWS

# ---- 색상 ----
WHITE = "#ffffff"
TEXT = "#000000"
TEXT_DARK = "#363636"
LIGHT_GRAY = "#d3d3d3"          # WPF LightGray
DARK_GRAY = "#a9a9a9"           # WPF DarkGray
GROUP_BORDER = "#d5dfe5"        # WPF GroupBox 기본 테두리

ZONE_TITLE_BG = "#34495e"
SIDE_EVEN_BG = "#f39c12"        # 짝
SIDE_ODD_BG = "#2980b9"         # 홀

NET_OK = "#2980b9"              # RGB(41,128,185)
NET_ERR = "#c0392b"             # RGB(192,57,43)
NET_STOP = "#808080"            # RGB(128,128,128)

START_ICON = "#2980b9"
STOP_ICON = "#c0392b"

LINE_DIALOG_HEADER = "#2c3e50"
MOTOR_DIALOG_HEADER = "#16a085"
DIALOG_BORDER = "#e0e0e0"
ENTRY_BG = "#efefef"
SEPARATOR = "#e0e0e0"

LIST_HOVER = "#e5f3ff"
LIST_SELECTED = "#cce8ff"
LIST_HEADER_BG = "#ffffff"
LIST_HEADER_HOVER = "#d9ebf9"
LIST_HEADER_LINE = "#e5e5e5"
LIST_BORDER = "#828790"

STATUS_ERR_FG = "#ff0000"
STATUS_DIM_FG = LIGHT_GRAY

BUTTON_STYLES = {
    # 이름: (배경, 마우스오버 배경, 글자색, 테두리)
    "normal": ("#ecf0f1", "#d4d4d4", "#000000", "#e0e0e0"),   # NormalButtonTemplate
    "submit": ("#547791", "#3a5e79", "#ffffff", None),        # SubmitButtonTemplate
    "alert": ("#e74c3c", "#c0392b", "#ffffff", None),         # AlertButtonTemplate
}

# ---- 글꼴 ----
_PREFERRED_FAMILIES = ["Malgun Gothic", "맑은 고딕", "Noto Sans CJK KR", "NanumGothic",
                       "Apple SD Gothic Neo", "WenQuanYi Zen Hei", "Segoe UI"]
FONT_FAMILY = "TkDefaultFont"


def init_fonts(root: tk.Misc) -> None:
    global FONT_FAMILY
    fams = set(tkfont.families(root))
    for f in _PREFERRED_FAMILIES:
        if f in fams:
            FONT_FAMILY = f
            break
    for name in ("TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont"):
        try:
            tkfont.nametofont(name, root).configure(family=FONT_FAMILY, size=9)
        except tk.TclError:
            pass


def font(size: int = 9, bold: bool = False) -> tuple:
    return (FONT_FAMILY, size, "bold") if bold else (FONT_FAMILY, size)


# ---- 이미지 ----
def _scaled(img: tk.PhotoImage, target: int) -> tk.PhotoImage:
    """정수 배율 zoom/subsample 로 크기 조정 (외부 라이브러리 없이)."""
    w = img.width()
    if w == target:
        return img
    from math import gcd
    g = gcd(target, w)
    z, s = target // g, w // g
    return img.zoom(z, z).subsample(s, s) if z > 1 else img.subsample(s, s)


class Images:
    """원본 images 폴더의 그림을 원본 화면과 같은 크기로 준비."""

    def __init__(self, root: tk.Misc) -> None:
        self._raw: Dict[str, tk.PhotoImage] = {}
        self._cache: Dict[tuple, tk.PhotoImage] = {}
        self.root = root

    def get(self, name: str, size: int) -> tk.PhotoImage:
        key = (name, size)
        if key not in self._cache:
            if name not in self._raw:
                self._raw[name] = tk.PhotoImage(master=self.root, file=os.path.join(ASSET_DIR, f"{name}.png"))
            self._cache[key] = _scaled(self._raw[name], size)
        return self._cache[key]


def set_window_icon(win: tk.Tk) -> None:
    ico = os.path.join(ASSET_DIR, "motor.ico")
    try:
        if IS_WINDOWS:
            win.iconbitmap(default=ico)
            return
    except tk.TclError:
        pass
    try:
        win.iconphoto(True, tk.PhotoImage(master=win, file=os.path.join(ASSET_DIR, "oldmotor.png")))
    except tk.TclError:
        pass
