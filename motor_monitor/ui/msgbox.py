"""메시지 창 (운영체제 메시지 상자 대신).

운영체제 메시지 상자는 Windows 가 화면(모니터) 가운데에 띄우므로, 프로그램 창 안 가운데에 뜨는
MessageDialog(dialogs.py)를 씁니다. 함수 이름과 인자는 tkinter.messagebox 와 비슷하게 맞췄습니다.
(시험에서는 이 모듈의 함수를 바꿔 끼워 창 없이 확인합니다.)
"""
from __future__ import annotations

import tkinter as tk


def _show(title: str, message: str, parent: tk.Misc, kind: str, buttons) -> "int | None":
    from .dialogs import MessageDialog      # dialogs 가 이 모듈을 쓰므로 여기서 불러옴
    d = MessageDialog(parent, title, message, kind, buttons)
    d.show_modal()
    return d.result


def show_info(title: str, message: str, parent: tk.Misc) -> None:
    _show(title, message, parent, "info", ("확인",))


def show_warning(title: str, message: str, parent: tk.Misc) -> None:
    _show(title, message, parent, "warning", ("확인",))


def show_error(title: str, message: str, parent: tk.Misc) -> None:
    _show(title, message, parent, "error", ("확인",))


def ask_yes_no(title: str, message: str, parent: tk.Misc, icon: str = "question") -> bool:
    """[예]=True, [아니오]·닫기·Esc=False. Enter 는 [예]."""
    return _show(title, message, parent, icon, ("예", "아니오")) == 0
