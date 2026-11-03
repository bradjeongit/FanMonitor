"""대화상자: 구역설정 창(운영 wHOUSE), 모터 추가/수정 창(운영 wMotor), 메시지 창.

컴포넌트 이름(속성명)은 스펙시트 docs/SPEC_SHEET.md 의 '코드 객체명' 과 같습니다.
새 창은 모두 기존 창(부모 창) 안 가운데에 띄웁니다. 화면(모니터) 기준으로 옮기거나 자르지 않으므로,
프로그램 창이 다른 모니터에 있으면 대화상자·메시지 창도 그 모니터의 프로그램 창 안에 뜹니다.
"""
from __future__ import annotations

import re
import tkinter as tk
from tkinter import ttk
from typing import TYPE_CHECKING, Callable, List, Optional, Sequence, Tuple

from ..constants import (CTRL_MODE_MANUAL, CTRL_MODE_REF, MAX_RPM, MODBUS_PORT, MOTOR_ID_MAX, MOTOR_ID_MIN,
                         MotorType, id_range, pct_text, rpm_to_pct)
from ..helpers import check_integer, check_valid_ip, get_house_name
from ..house import motor_key
from . import msgbox
from . import theme as T
from .widgets import DottedSeparator, RoundButton

if TYPE_CHECKING:
    from ..plant import Plant

_INT_RE = re.compile(r"^\s*[+-]?\d+\s*$")


def try_parse_int(text: str) -> Optional[int]:
    """VB Integer.TryParse 와 같은 규칙(앞뒤 공백, 부호 허용). 실패 시 None."""
    if not _INT_RE.match(text or ""):
        return None
    v = int(text)
    if not -2147483648 <= v <= 2147483647:
        return None
    return v


# ======================================================================== 창 위치
Rect = Tuple[int, int, int, int]      # 화면 좌표 (x, y, 폭, 높이)


def inside_position(owner: Rect, size: Tuple[int, int], bounds: Optional[Rect] = None) -> Tuple[int, int]:
    """새 창의 왼쪽 위 좌표: 기존 창(owner) 가운데에 두고, 프로그램 창(bounds, 없으면 owner) 밖으로 나가지 않게 맞춤.
    새 창이 bounds 보다 크면 bounds 의 왼쪽 위에 맞춤. 화면(모니터) 크기로는 자르지 않음
    (프로그램 창이 왼쪽/위쪽 모니터에 있어 좌표가 음수여도 그대로 사용)."""
    ox, oy, ow, oh = owner
    w, h = size
    bx, by, bw, bh = bounds or owner
    x = max(bx, min(ox + (ow - w) // 2, bx + bw - w))
    y = max(by, min(oy + (oh - h) // 2, by + bh - h))
    return x, y


def owner_rect(owner: tk.Misc) -> Rect:
    """기존 창(owner 가 속한 최상위 창)의 화면 좌표 영역."""
    top = owner.winfo_toplevel()
    top.update_idletasks()
    return top.winfo_rootx(), top.winfo_rooty(), top.winfo_width(), top.winfo_height()


def app_rect(owner: tk.Misc) -> Optional[Rect]:
    """프로그램 메인 창의 영역 (보이지 않으면 None)."""
    root = owner._root()
    try:
        if root.winfo_viewable():
            return root.winfo_rootx(), root.winfo_rooty(), root.winfo_width(), root.winfo_height()
    except tk.TclError:
        pass
    return None


def geometry_text(width: int, height: int, x: int, y: int) -> str:
    """Tk 창 위치 문자열. 음수 좌표는 '+-100' 형식 (화면 왼쪽/위쪽 바깥 = 다른 모니터)."""
    return f"{width}x{height}+{x}+{y}"


def place_inside(win: tk.Toplevel, owner: tk.Misc, width: int, height: int) -> None:
    """새 창(win)을 기존 창(owner) 가운데, 프로그램 메인 창 안에 둠."""
    try:
        x, y = inside_position(owner_rect(owner), (width, height), app_rect(owner))
    except tk.TclError:
        win.geometry(f"{width}x{height}")
        return
    win.geometry(geometry_text(width, height, x, y))


class BaseDialog(tk.Toplevel):
    """원본 대화상자 공통 틀: 색 머리글(끌어서 이동) + 본문 + 아래쪽 버튼줄, 모달.
    기존 창(owner) 안 가운데에 뜸 (place_inside)."""

    def __init__(self, owner: tk.Misc, title: str, header_bg: str, width: int, height: int) -> None:
        super().__init__(owner)
        self.withdraw()
        self.title(title)
        self.configure(bg=T.DIALOG_BORDER)
        self.resizable(False, False)
        self.transient(owner)
        if T.BORDERLESS_DIALOGS:
            self.overrideredirect(True)
        self._dlg_w, self._dlg_h = width, height

        frame = tk.Frame(self, bg=T.WHITE)
        frame.pack(fill="both", expand=True, padx=1, pady=1)
        self.header = tk.Frame(frame, bg=header_bg, height=40)
        self.header.pack(side="top", fill="x")
        self.header.pack_propagate(False)
        self.title_label = tk.Label(self.header, text=title, bg=header_bg, fg=T.WHITE, font=T.font(13, bold=True))
        self.title_label.pack(expand=True)
        for w in (self.header, self.title_label):
            w.bind("<ButtonPress-1>", self._drag_start)
            w.bind("<B1-Motion>", self._drag_move)

        self.footer = tk.Frame(frame, bg=T.WHITE, height=50)
        self.footer.pack(side="bottom", fill="x")
        self.footer.pack_propagate(False)
        self.body = tk.Frame(frame, bg=T.WHITE)
        self.body.pack(side="top", fill="both", expand=True)

        self.bind("<Escape>", lambda e: self.on_cancel())
        self.bind("<Return>", lambda e: self.on_default())
        self.protocol("WM_DELETE_WINDOW", self.on_cancel)

    # 하위 클래스에서 재정의
    def on_default(self) -> None:
        pass

    def on_cancel(self) -> None:
        self.destroy()

    def _drag_start(self, e) -> None:
        self._dx, self._dy = e.x_root - self.winfo_x(), e.y_root - self.winfo_y()

    def _drag_move(self, e) -> None:
        self.geometry(f"+{e.x_root - self._dx}+{e.y_root - self._dy}")

    def row(self, title: str, indent: int = 10, title_size: int = 9) -> tk.Frame:
        """원본 'LineTitle' 줄: 높이 40, 제목 폭 100."""
        r = tk.Frame(self.body, bg=T.WHITE, height=40)
        r.pack(side="top", fill="x", padx=(indent, 10))
        r.pack_propagate(False)
        tk.Label(r, text=title, bg=T.WHITE, fg=T.TEXT_DARK, font=T.font(title_size, bold=True),
                 anchor="w", width=0).place(x=10, rely=0.5, anchor="w")
        inner = tk.Frame(r, bg=T.WHITE)
        inner.place(x=120, rely=0.5, anchor="w")
        return inner

    def separator(self) -> None:
        DottedSeparator(self.body).pack(side="top", fill="x")

    def entry(self, parent, width_px: int, size: int = 11, bold: bool = True) -> tk.Entry:
        holder = tk.Frame(parent, width=width_px, height=26, bg=T.ENTRY_BG,
                          highlightthickness=1, highlightbackground="#abadb3")
        holder.pack_propagate(False)
        e = tk.Entry(holder, bg=T.ENTRY_BG, fg=T.TEXT_DARK, relief="flat", bd=0,
                     font=T.font(size, bold=bold), insertbackground=T.TEXT_DARK)
        e.pack(fill="both", expand=True, padx=2)
        holder.pack(side="left")
        return e

    def show_modal(self) -> None:
        """기존 창 안 가운데에 띄우고 닫힐 때까지 기다림. 다른 모달 창 위에 떴으면 닫힌 뒤 그 창으로 되돌림."""
        self.update_idletasks()
        place_inside(self, self.master, self._dlg_w, self._dlg_h)
        self.deiconify()
        self.lift()
        try:
            prev_grab = self.grab_current()
        except (tk.TclError, KeyError):
            prev_grab = None
        try:
            self.grab_set()
        except tk.TclError:
            pass
        self.focus_force()
        self.wait_window(self)
        if prev_grab is not None and prev_grab is not self:
            try:
                if prev_grab.winfo_exists():
                    prev_grab.grab_set()
                    prev_grab.focus_force()
            except tk.TclError:
                pass

    def keep_inside_owner(self) -> None:
        """크기가 바뀐 뒤에도 프로그램 창 안에 머물도록 위치 조정 (지금 위치를 유지하되 넘치는 만큼만 옮김)."""
        try:
            bx, by, bw, bh = app_rect(self.master) or owner_rect(self.master)
            x, y = self.winfo_rootx(), self.winfo_rooty()
        except tk.TclError:
            self.geometry(f"{self._dlg_w}x{self._dlg_h}")
            return
        x = max(bx, min(x, bx + bw - self._dlg_w))
        y = max(by, min(y, by + bh - self._dlg_h))
        self.geometry(geometry_text(self._dlg_w, self._dlg_h, x, y))

    def error(self, msg: str, title: str) -> None:
        msgbox.show_error(title, msg, parent=self)

    def ask(self, msg: str, title: str) -> bool:
        return msgbox.ask_yes_no(title, msg, parent=self)


class MessageDialog(BaseDialog):
    """메시지/확인 창. 운영체제 메시지 상자(화면 가운데에 뜸) 대신 써서 기존 창 안 가운데에 띄움.
    result: 누른 버튼 번호 (첫 버튼 0), 닫거나 Esc 면 None. Enter 는 첫 버튼."""

    WIDTH = 460
    ICONS = {"info": ("i", "#2e86c1"), "question": ("?", "#2e86c1"),
             "warning": ("!", "#d68910"), "error": ("×", "#c0392b")}

    def __init__(self, owner: tk.Misc, title: str, message: str, kind: str = "info",
                 buttons: Sequence[str] = ("확인",)) -> None:
        super().__init__(owner, title, T.LINE_DIALOG_HEADER, self.WIDTH, 170)
        self.kind = kind if kind in self.ICONS else "info"
        self.message = message
        self.result: Optional[int] = None
        holder = tk.Frame(self.body, bg=T.WHITE)
        holder.pack(side="top", fill="both", expand=True, padx=16, pady=14)
        char, color = self.ICONS[self.kind]
        icon = tk.Canvas(holder, width=34, height=34, bg=T.WHITE, highlightthickness=0)
        icon.create_oval(2, 2, 32, 32, fill=color, outline=color)
        icon.create_text(17, 17, text=char, fill=T.WHITE, font=T.font(14, bold=True))
        icon.pack(side="left", anchor="n")
        self.lbl_msg = tk.Label(holder, text=message, bg=T.WHITE, fg=T.TEXT_DARK, font=T.font(10),
                                justify="left", anchor="nw", wraplength=self.WIDTH - 100)
        self.lbl_msg.pack(side="left", fill="both", expand=True, padx=(12, 0))
        btns = tk.Frame(self.footer, bg=T.WHITE)
        btns.pack(side="right", padx=5)
        self.buttons: List[RoundButton] = []
        for i, text in enumerate(buttons):
            b = RoundButton(btns, text, lambda i=i: self.press(i), style="submit" if i == 0 else "normal")
            b.pack(side="left", padx=(0, 5))
            self.buttons.append(b)
        # 키보드: 첫 버튼(예/확인) = Enter·Space·Y, 둘째 버튼(아니오) = Esc·N. 한글 입력 상태의 ㅛ/ㅜ 도 같음
        first = buttons[0] if buttons else "확인"
        hint = f"Enter·Space·Y = {first}" + (f"   Esc·N = {buttons[1]}" if len(buttons) > 1 else "")
        self.lbl_keys = tk.Label(self.footer, text=hint, bg=T.WHITE, fg="#8a8a8a", font=T.font(8))
        self.lbl_keys.pack(side="left", padx=(12, 0))
        for seq in ("<space>", "<KP_Enter>"):
            self.bind(seq, lambda e: self._key_default())
        self.bind("<Key>", self._on_key, add="+")
        self.update_idletasks()
        self._dlg_h = 40 + max(self.lbl_msg.winfo_reqheight(), 34) + 28 + 50 + 2

    def _key_default(self) -> str:
        self.on_default()
        return "break"

    def _on_key(self, e) -> Optional[str]:
        """Y(ㅛ) = 첫 버튼, N(ㅜ) = 둘째 버튼. Enter·Space·Esc 는 따로 연결."""
        ch = (e.char or "").lower()
        ks = (e.keysym or "").lower()
        if ks == "y" or ch in ("y", "ㅛ"):
            self.press(0)
            return "break"
        if (ks == "n" or ch in ("n", "ㅜ")) and len(self.buttons) > 1:
            self.press(1)
            return "break"
        return None

    def press(self, index: int) -> None:
        if not self.winfo_exists():
            return
        self.result = index
        self.destroy()

    def on_default(self) -> None:
        self.press(0)

    def on_cancel(self) -> None:
        self.result = None
        self.destroy()


# ======================================================================== 구역설정 창
# 참조 구역 선택 항목 (값은 'R1'~'R8', 저장 시 'H1' 을 앞에 붙여 'H1R1' 형식)
LINE_CHOICES: List[Tuple[str, str]] = [("", "구역선택")] + [(f"R{i}", f"{i}구역") for i in range(1, 9)]


class LineSettingsDialog(BaseDialog):
    """DLG-LINE 구역설정 창. 운영 wHOUSE"""

    def __init__(self, owner: tk.Misc, plant: "Plant", house_code: str,
                 on_saved: Optional[Callable[[], None]] = None) -> None:
        super().__init__(owner, "구역설정", T.LINE_DIALOG_HEADER, 400, 300)
        self.plant = plant
        self.house_code = house_code
        self.house_header = house_code[:2]
        self.on_saved = on_saved
        self._id_tags: List[str] = []

        r = self.row("위치")
        self.lbl_house = tk.Label(r, text="-", bg=T.WHITE, font=T.font(11, bold=True))
        self.lbl_house.pack(side="left")
        self.separator()
        r = self.row("기존모터 IP")         # PORT 는 502 고정이라 입력칸 없음
        self.txtSIP = self.entry(r, 140)
        r = self.row("신규모터 IP")
        self.txtTIP = self.entry(r, 140)
        self.separator()
        r = self.row("RPM 제어 소스")
        self.source_mode = tk.StringVar(value="")
        self.rb_source_manual = tk.Radiobutton(
            r, text="수동", value="manual", variable=self.source_mode, command=self._on_mode,
            bg=T.WHITE, activebackground=T.WHITE, font=T.font(11, bold=True), highlightthickness=0)
        self.rb_source_manual.pack(side="left")
        self.rb_source_ref = tk.Radiobutton(
            r, text="기존모터", value="ref", variable=self.source_mode, command=self._on_mode,
            bg=T.WHITE, activebackground=T.WHITE, font=T.font(11, bold=True), highlightthickness=0)
        self.rb_source_ref.pack(side="left", padx=20)

        self.mode_holder = tk.Frame(self.body, bg=T.WHITE, height=40)
        self.mode_holder.pack(side="top", fill="x")
        self.mode_holder.pack_propagate(False)

        self.manual_line = tk.Frame(self.mode_holder, bg=T.WHITE)
        tk.Label(self.manual_line, text="┖ 수동RPM값", bg=T.WHITE, fg=T.TEXT_DARK,
                 font=T.font(9, bold=True)).place(x=10, rely=0.5, anchor="w")
        inner = tk.Frame(self.manual_line, bg=T.WHITE)
        inner.place(x=120, rely=0.5, anchor="w")
        self.txtMRPM = self.entry(inner, 80)
        self.lbl_rpm_range = tk.Label(inner, text="-", bg=T.WHITE)
        self.lbl_rpm_range.pack(side="left", padx=5)

        self.ref_line = tk.Frame(self.mode_holder, bg=T.WHITE)
        tk.Label(self.ref_line, text="┖ 참조기존모터", bg=T.WHITE, fg=T.TEXT_DARK,
                 font=T.font(9, bold=True)).place(x=10, rely=0.5, anchor="w")
        inner = tk.Frame(self.ref_line, bg=T.WHITE)
        inner.place(x=120, rely=0.5, anchor="w")
        self.cb_lines = ttk.Combobox(inner, state="readonly", width=8, height=len(LINE_CHOICES),
                                     font=T.font(11, bold=True), values=[name for _, name in LINE_CHOICES])
        self.cb_lines.pack(side="left")
        self.cb_lines.bind("<<ComboboxSelected>>", lambda e: self._on_line_changed())
        tk.Label(inner, text=" > ", bg=T.WHITE).pack(side="left")
        self.cb_ids = ttk.Combobox(inner, state="readonly", width=4, font=T.font(11, bold=True))
        self.cb_ids.pack(side="left")
        self.separator()

        btns = tk.Frame(self.footer, bg=T.WHITE)
        btns.pack(side="right", padx=5)
        self.BtnSave = RoundButton(btns, "저장", self.on_save, style="submit")
        self.BtnSave.pack(side="left", padx=(0, 5))
        self.BtnCancel = RoundButton(btns, "취소", self.on_cancel, style="normal")
        self.BtnCancel.pack(side="left")

        self._load()

    def _load(self) -> None:
        """운영 wHOUSE_Loaded"""
        cfg = self.plant.configs[self.house_code]
        self.lbl_house.configure(text=get_house_name(self.house_code))
        for e, v in ((self.txtSIP, cfg.source_ip), (self.txtTIP, cfg.target_ip), (self.txtMRPM, cfg.manual_rpm)):
            e.delete(0, "end")
            e.insert(0, str(v))
        source_line = cfg.source_line or self.house_code
        if len(source_line) == 4:
            tag = source_line[2:4]
            self._select_line(tag)
            self._load_old_ids(self.house_header + tag, str(cfg.source_id))
        else:
            self._select_line("")
            self._load_old_ids(None)
        self.source_mode.set("manual" if cfg.source_manual else "ref")
        self._on_mode()
        self.lbl_rpm_range.configure(text=f"정수(0~{MAX_RPM})")

    def _select_line(self, tag: str) -> None:
        tags = [t for t, _ in LINE_CHOICES]
        self.cb_lines.current(tags.index(tag) if tag in tags else 0)

    def _line_tag(self) -> str:
        i = self.cb_lines.current()
        return LINE_CHOICES[i][0] if i >= 0 else ""

    def _load_old_ids(self, lcode: Optional[str], default: str = "") -> None:
        """운영 Load_OLD_IDS: 첫 항목 'ID' + 해당 구역의 기존모터 ID 목록."""
        self._id_tags = [""]
        names = ["ID"]
        house = self.plant.houses.get(lcode) if lcode else None
        if house is not None:
            for mid in house.get_old_motor_ids():
                self._id_tags.append(mid)
                names.append(mid)
        self.cb_ids.configure(values=names)
        sel = self._id_tags.index(default) if default and default in self._id_tags else 0
        self.cb_ids.current(sel)

    def _id_tag(self) -> str:
        i = self.cb_ids.current()
        return self._id_tags[i] if 0 <= i < len(self._id_tags) else ""

    def _on_line_changed(self) -> None:
        tag = self._line_tag()
        self._load_old_ids(self.house_header + tag if tag else None)

    def _on_mode(self) -> None:
        if self.source_mode.get() == "manual":
            self.ref_line.place_forget()
            self.manual_line.place(x=20, y=0, relwidth=1, relheight=1)
        else:
            self.manual_line.place_forget()
            self.ref_line.place(x=20, y=0, relwidth=1, relheight=1)

    def on_default(self) -> None:
        self.on_save()

    def on_save(self) -> None:
        """운영 BtnSave_Click"""
        sip = self.txtSIP.get()
        if sip == "":
            return self.error("기존모터 접속 IP를 입력해 주세요.", "에러")
        if not check_valid_ip(sip):
            return self.error("기존모터 접속 IP가 올바른 형식이 아닙니다..", "에러")
        tip = self.txtTIP.get()
        if tip == "":
            return self.error("신규 모터 접속 IP를 입력해 주세요.", "에러")
        if not check_valid_ip(tip):
            return self.error("신규모터 접속 IP가 올바른 형식이 아닙니다..", "에러")

        mode = self.source_mode.get()
        if mode == "manual":
            smanual = True
            mrpm = try_parse_int(self.txtMRPM.get())
            if mrpm is None:
                return self.error("수동 RPM속도를 숫자로 입력해주세요.", "에러")
            if mrpm < 0 or mrpm > MAX_RPM:
                return self.error(f"수동 RPM속도를 0~{MAX_RPM}범위내로 입력해주세요.", "에러")
            sline = self._line_tag()
            id_tag = self._id_tag()
            sid = int(id_tag) if id_tag not in ("", "ID") else -1
        elif mode == "ref":
            smanual = False
            mrpm = try_parse_int(self.txtMRPM.get()) or 0
            sline = self._line_tag()
            if sline == "":
                return self.error("참조할 모터가 있는 구역을 선택해주세요.", "에러")
            id_tag = self._id_tag()
            sid = try_parse_int(id_tag) if id_tag != "ID" else None
            if sid is None or sid < 0:
                return self.error("참조할 모터 ID를 선택해주세요.", "에러")
        else:
            return self.error("RPM제어 소스를 선택해주세요.", "")

        cfg = self.plant.configs[self.house_code].copy(
            source_ip=sip, source_port=MODBUS_PORT, target_ip=tip, target_port=MODBUS_PORT,
            source_manual=smanual, manual_rpm=mrpm, source_line=self.house_header + sline, source_id=sid)
        self.plant.set_line_config(self.house_code, cfg)
        if self.on_saved:
            self.on_saved()
        self.destroy()


# ======================================================================== 모터 추가/수정 창
class MotorDialog(BaseDialog):
    """DLG-MOTOR 모터 추가(ADD) / 모터 수정·삭제(MOD) 창. 운영 wMotor"""

    LIMIT_CHOICES = ("사용함", "사용안함")
    BASE_H = 340
    CTRL_H = 460

    def __init__(self, owner: tk.Misc, plant: "Plant", images: T.Images, house_code: str,
                 mode: str = "ADD", mod_dkey: Optional[str] = None,
                 on_changed: Optional[Callable[[], None]] = None) -> None:
        mode = mode if mode in ("ADD", "MOD") else "ADD"
        title = "모터 추가" if mode == "ADD" else "모터 수정/삭제"
        # 기존모터제어(3번째 종류)와 그 입력줄을 위해 운영 프로그램(400x300)보다 크게. 기존모터제어면 더 늘림
        super().__init__(owner, title, T.MOTOR_DIALOG_HEADER, 560, self.BASE_H)
        self.plant = plant
        self.images = images
        self.house_code = house_code
        self.wmode = mode
        self.mod_dkey = mod_dkey
        self.on_changed = on_changed
        self.motor_id = 0
        self.motor_type = MotorType.OLD_MOTOR

        r = self.row("위치", title_size=12)
        self.lbl_house = tk.Label(r, text="-", bg=T.WHITE, font=T.font(11, bold=True))
        self.lbl_house.pack(side="left")
        self.separator()
        r = self.row("모터종류", title_size=12)
        self.motor_type_var = tk.IntVar(value=-1)
        self.add_type_line = tk.Frame(r, bg=T.WHITE)
        self.rb_oldmotor = self._radio(self.add_type_line, "기존모터", "oldmotor",
                                       self.motor_type_var, int(MotorType.OLD_MOTOR), self._on_type, 10)
        self.rb_newmotor = self._radio(self.add_type_line, "신규모터", "newmotor",
                                       self.motor_type_var, int(MotorType.NEW_MOTOR), self._on_type, 10)
        self.rb_newmotor.pack_configure(padx=(6, 0))
        self.rb_ctrlmotor = self._radio(self.add_type_line, "기존모터제어", "oldmotor_ctrl",
                                        self.motor_type_var, int(MotorType.OLD_CTRL), self._on_type, 10)
        self.rb_ctrlmotor.pack_configure(padx=(6, 0))
        self.mod_type_line = tk.Frame(r, bg=T.WHITE)
        self.img_type = tk.Label(self.mod_type_line, bg=T.WHITE)
        self.img_type.pack(side="left", padx=(0, 5))
        self.lbl_type = tk.Label(self.mod_type_line, text="type", bg=T.WHITE, font=T.font(11))
        self.lbl_type.pack(side="left")
        self.separator()
        r = self.row("모터ID", title_size=12)
        self.add_id_line = tk.Frame(r, bg=T.WHITE)
        self.txtID = self.entry(self.add_id_line, 60, size=11, bold=False)
        self.lbl_id_range = tk.Label(self.add_id_line, text=f"정수({MOTOR_ID_MIN}~{MOTOR_ID_MAX}, 종류별)", bg=T.WHITE)
        self.lbl_id_range.pack(side="left", padx=10)
        self.mod_id_line = tk.Frame(r, bg=T.WHITE)
        self.lblID = tk.Label(self.mod_id_line, text="00", bg=T.WHITE, font=T.font(11, bold=True))
        self.lblID.pack(side="left")
        self.separator()
        r = self.row("사용여부", title_size=12)
        self.use_var = tk.IntVar(value=1)          # 운영: 기본 '사용' 선택
        self.rb_use = self._radio(r, "사용", "good", self.use_var, 1)
        self.rb_nouse = self._radio(r, "중지", "nouse", self.use_var, 0)
        self.rb_nouse.pack_configure(padx=(36, 0))
        self.separator()
        # 속도제한 줄 (신규모터만 표시)
        self.limit_holder = tk.Frame(self.body, bg=T.WHITE)
        self.limit_holder.pack(side="top", fill="x")
        self.limit_row = tk.Frame(self.limit_holder, bg=T.WHITE, height=40)
        self.limit_row.pack_propagate(False)
        tk.Label(self.limit_row, text="속도제한", bg=T.WHITE, fg=T.TEXT_DARK,
                 font=T.font(12, bold=True)).place(x=10, rely=0.5, anchor="w")
        inner = tk.Frame(self.limit_row, bg=T.WHITE)
        inner.place(x=130, rely=0.5, anchor="w")
        self.cbUSELIMIT = ttk.Combobox(inner, state="readonly", width=7, values=self.LIMIT_CHOICES,
                                       font=T.font(10))
        self.cbUSELIMIT.current(1)                 # 기본 '사용안함'
        self.cbUSELIMIT.pack(side="left")
        tk.Label(inner, text="최대속도:", bg=T.WHITE).pack(side="left", padx=(10, 2))
        self.txtLIMITRPM = self.entry(inner, 65, size=11, bold=False)
        tk.Label(inner, text="RPM", bg=T.WHITE).pack(side="left", padx=2)
        self.limit_row_line = DottedSeparator(self.limit_holder)
        # 기존모터제어 줄 (접속 IP, 운전속도, 구역 참조값) - 기존모터제어만 표시. PORT 는 502 고정이라 입력칸 없음
        self.ctrl_holder = tk.Frame(self.body, bg=T.WHITE)
        self.ctrl_holder.pack(side="top", fill="x")
        self.ctrl_ip_row = tk.Frame(self.ctrl_holder, bg=T.WHITE, height=40)
        self.ctrl_ip_row.pack_propagate(False)
        tk.Label(self.ctrl_ip_row, text="접속 IP", bg=T.WHITE, fg=T.TEXT_DARK,
                 font=T.font(12, bold=True)).place(x=10, rely=0.5, anchor="w")
        inner = tk.Frame(self.ctrl_ip_row, bg=T.WHITE)
        inner.place(x=130, rely=0.5, anchor="w")
        self.txtCIP = self.entry(inner, 140, size=11, bold=False)
        # 속도 지정 방법: 수동 지정(운전속도) / 구역 참조(이 구역 신규모터가 따르는 값을 1100 대비 % 로 같은 비율)
        self.ctrl_speed_row = tk.Frame(self.ctrl_holder, bg=T.WHITE, height=40)
        self.ctrl_speed_row.pack_propagate(False)
        tk.Label(self.ctrl_speed_row, text="운전속도", bg=T.WHITE, fg=T.TEXT_DARK,
                 font=T.font(12, bold=True)).place(x=10, rely=0.5, anchor="w")
        inner = tk.Frame(self.ctrl_speed_row, bg=T.WHITE)
        inner.place(x=130, rely=0.5, anchor="w")
        self.ctrl_mode_var = tk.StringVar(value=CTRL_MODE_MANUAL)
        self.rb_ctrl_manual = tk.Radiobutton(inner, text="수동 지정", value=CTRL_MODE_MANUAL,
                                             variable=self.ctrl_mode_var, command=self._on_ctrl_mode,
                                             bg=T.WHITE, activebackground=T.WHITE, font=T.font(10, bold=True),
                                             highlightthickness=0)
        self.rb_ctrl_manual.pack(side="left")
        self.txtSETRPM = self.entry(inner, 60, size=11, bold=False)
        self.txtSETRPM.bind("<KeyRelease>", lambda e: self._show_set_pct(), add="+")
        self.lbl_set_pct = tk.Label(inner, text="", bg=T.WHITE, width=17, anchor="w")
        self.lbl_set_pct.pack(side="left", padx=(4, 6))
        self.rb_ctrl_ref = tk.Radiobutton(inner, text="구역 참조", value=CTRL_MODE_REF,
                                          variable=self.ctrl_mode_var, command=self._on_ctrl_mode,
                                          bg=T.WHITE, activebackground=T.WHITE, font=T.font(10, bold=True),
                                          highlightthickness=0)
        self.rb_ctrl_ref.pack(side="left")
        self.ctrl_ref_row = tk.Frame(self.ctrl_holder, bg=T.WHITE, height=40)
        self.ctrl_ref_row.pack_propagate(False)
        tk.Label(self.ctrl_ref_row, text="┖ 구역 참조값", bg=T.WHITE, fg=T.TEXT_DARK,
                 font=T.font(11, bold=True)).place(x=10, rely=0.5, anchor="w")
        self.lbl_zone_ref = tk.Label(self.ctrl_ref_row, text="-", bg=T.WHITE, fg=T.TEXT_DARK,
                                     font=T.font(10, bold=True), anchor="w")
        self.lbl_zone_ref.place(x=130, rely=0.5, anchor="w")
        self.ctrl_ref_note = tk.Label(self.ctrl_holder, bg=T.WHITE, fg="#666666", font=T.font(8), anchor="w",
                                      justify="left", wraplength=500,
                                      text=f"구역 참조: 이 구역 신규모터가 따르는 값(구역설정의 RPM 제어 소스)을 최고속도 "
                                           f"{MAX_RPM} RPM 대비 %로 바꿔 같은 비율로 운전합니다. "
                                           "값이 없으면 이전 속도를 유지합니다.")
        # 모터 최고속도(h8) 기준 % (수정 창, 구역 제어 중 h8 을 읽은 뒤에만 표시)
        self.motor_max_rpm = 0
        self.lbl_motor_max = tk.Label(self.ctrl_holder, bg=T.WHITE, fg="#1f4e79", font=T.font(8, bold=True),
                                      anchor="w", justify="left", text="")
        self.ctrl_row_line = DottedSeparator(self.ctrl_holder)

        left = tk.Frame(self.footer, bg=T.WHITE)
        left.pack(side="left", padx=5)
        self.BtnDelete = RoundButton(left, "삭제", self.on_delete, style="alert")
        right = tk.Frame(self.footer, bg=T.WHITE)
        right.pack(side="right", padx=5)
        self.BtnSave = RoundButton(right, "저장", self.on_save, style="submit")
        self.BtnSave.pack(side="left", padx=(0, 5))
        self.BtnCancel = RoundButton(right, "취소", self.on_cancel, style="normal")
        self.BtnCancel.pack(side="left")

        self._load()

    def _radio(self, parent, text: str, image: str, var: tk.IntVar, value: int, command=None,
               size: int = 11) -> tk.Radiobutton:
        rb = tk.Radiobutton(parent, text=" " + text, image=self.images.get(image, 20), compound="left",
                            variable=var, value=value, bg=T.WHITE, activebackground=T.WHITE,
                            font=T.font(size, bold=True), highlightthickness=0, command=command)
        rb.pack(side="left")
        return rb

    def show_limit_row(self, show: bool) -> None:
        if show:
            self.limit_row.pack(side="top", fill="x", padx=(10, 10))
            self.limit_row_line.pack(side="top", fill="x")
        else:
            self.limit_row.pack_forget()
            self.limit_row_line.pack_forget()

    @property
    def limit_row_visible(self) -> bool:
        return bool(self.limit_row.winfo_manager())

    def show_ctrl_rows(self, show: bool) -> None:
        if show:
            self.ctrl_ip_row.pack(side="top", fill="x", padx=(10, 10))
            self.ctrl_speed_row.pack(side="top", fill="x", padx=(10, 10))
            self.ctrl_ref_row.pack(side="top", fill="x", padx=(10, 10))
            self.ctrl_ref_note.pack(side="top", fill="x", padx=(30, 10))
            if self.motor_max_rpm > 0:
                self.lbl_motor_max.pack(side="top", fill="x", padx=(30, 10), pady=(2, 0))
            self.ctrl_row_line.pack(side="top", fill="x", pady=(4, 0))
            self._on_ctrl_mode()
        else:
            for w in (self.ctrl_ip_row, self.ctrl_speed_row, self.ctrl_ref_row, self.ctrl_ref_note,
                      self.lbl_motor_max, self.ctrl_row_line):
                w.pack_forget()
        self._dlg_h = self.CTRL_H if show else self.BASE_H
        if self.winfo_ismapped():
            self.keep_inside_owner()          # 크기가 바뀌어도 기존 창 안에 머묾

    def _on_ctrl_mode(self) -> None:
        ref = self.ctrl_mode_var.get() == CTRL_MODE_REF
        self.txtSETRPM.configure(state="disabled" if ref else "normal", disabledbackground=T.LIGHT_GRAY)
        self.lbl_zone_ref.configure(fg=T.TEXT_DARK if ref else "#9a9a9a")
        self._show_set_pct()

    def _show_set_pct(self) -> None:
        """수동 지정 RPM 옆에 1100 대비 % 표시 (+ 모터 최고속도를 알면 그 대비 %)"""
        v = try_parse_int(self.txtSETRPM.get())
        ok = v is not None and 0 <= v <= MAX_RPM
        pct = f" = {pct_text(rpm_to_pct(v))}" if ok else ""
        self.lbl_set_pct.configure(text=f"RPM (0~{MAX_RPM}){pct}")
        if self.motor_max_rpm > 0:
            ref = self.ctrl_mode_var.get() == CTRL_MODE_REF
            want = self.plant.houses[self.house_code].zone_reference()[1] if ref else (v if ok else None)
            text = f"모터 최고속도(h8) {self.motor_max_rpm} RPM"
            if want is not None:
                text += f": 원하는 속도 {want} RPM = 최고속도의 {pct_text(want * 100 / self.motor_max_rpm)}"
                if want > self.motor_max_rpm:
                    text += " (최고속도보다 높아 최고속도로 운전)"
            self.lbl_motor_max.configure(text=text)

    def zone_reference_text(self) -> str:
        """이 구역 신규모터가 따르는 값과 그 % (예: '1구역 기존모터 155 → 지금 770 RPM = 70%')."""
        desc, value = self.plant.houses[self.house_code].zone_reference()
        return f"{desc} → {self._now_text(value)}"

    @staticmethod
    def _now_text(value: Optional[int]) -> str:
        if value is None:
            return "지금 값 없음 (참조 기존모터가 있는 구역을 실행하면 읽음)"
        return f"지금 {value} RPM = {pct_text(rpm_to_pct(value))}"

    @staticmethod
    def _set_text(e: tk.Entry, value) -> None:
        """입력칸 값 바꾸기 (비활성 칸도 바꿀 수 있게 잠시 활성화)."""
        state = str(e.cget("state"))
        e.configure(state="normal")
        e.delete(0, "end")
        e.insert(0, str(value))
        e.configure(state=state)

    @property
    def ctrl_rows_visible(self) -> bool:
        return bool(self.ctrl_ip_row.winfo_manager())

    def _on_type(self) -> None:
        """운영 rb_oldmotor_Checked / rb_newmotor_Checked + 기존모터제어"""
        t = self.motor_type_var.get()
        self.show_limit_row(t == int(MotorType.NEW_MOTOR))
        self.show_ctrl_rows(t == int(MotorType.OLD_CTRL))
        lo, hi = id_range(MotorType(t))
        self.lbl_id_range.configure(text=f"정수({lo}~{hi})")

    def _load(self) -> None:
        """운영 wMotor_Loaded"""
        if self.wmode == "ADD":
            self.BtnSave.set_text("추가")
            self.add_type_line.pack(side="left")
            self.add_id_line.pack(side="left")
            self.show_limit_row(False)
            self.show_ctrl_rows(False)
            self.txtID.delete(0, "end")
            self.txtLIMITRPM.insert(0, "0")
            cfg = self.plant.configs[self.house_code]       # 기존모터 게이트웨이 주소를 기본값으로
            self.txtCIP.insert(0, cfg.source_ip)
            self.lbl_zone_ref.configure(text=self.zone_reference_text())
            self.ctrl_mode_var.set(CTRL_MODE_MANUAL)
            self._on_ctrl_mode()
            self.after(50, self.txtID.focus_set)
        else:
            self.BtnDelete.pack(side="left")
            self.BtnSave.set_text("수정")
            self.mod_type_line.pack(side="left")
            self.mod_id_line.pack(side="left")
            m = self.plant.houses[self.house_code].get_motor(self.mod_dkey or "")
            if m is not None:
                old = m.mtype == MotorType.OLD_MOTOR
                ctrl = m.mtype == MotorType.OLD_CTRL
                img = "oldmotor_ctrl" if ctrl else ("oldmotor" if old else "newmotor")
                self.img_type.configure(image=self.images.get(img, 20))
                self.lbl_type.configure(text="기존모터제어" if ctrl else ("기존모터" if old else "신규모터"))
                self.motor_id = m.id
                self.motor_type = m.mtype
                self.lblID.configure(text=str(m.id))
                self.use_var.set(1 if m.use else 0)
                self.cbUSELIMIT.current(0 if m.use_limit else 1)
                self.txtLIMITRPM.delete(0, "end")
                self.txtLIMITRPM.insert(0, str(m.limit_rpm))
                self.show_limit_row(m.mtype == MotorType.NEW_MOTOR)
                if ctrl:
                    self.show_ctrl_rows(True)
                    self.txtCIP.insert(0, m.ip)
                    self.txtCIP.configure(state="readonly", readonlybackground=T.WHITE)  # 접속처는 수정 불가
                    self._set_text(self.txtSETRPM, m.set_rpm)
                    self.ctrl_ip = m.ip
                    self.ctrl_port = m.port
                    self.motor_max_rpm = m.max_rpm
                    if m.max_rpm > 0:
                        self.show_ctrl_rows(True)       # 최고속도 줄 표시
                    self.lbl_zone_ref.configure(text=self.zone_reference_text())
                    self.ctrl_mode_var.set(m.ctrl_mode)
                    self._on_ctrl_mode()
        self.lbl_house.configure(text=get_house_name(self.house_code))

    def on_default(self) -> None:
        self.on_save()

    def on_save(self) -> None:
        """운영 BtnSave_Click"""
        msg = "추가" if self.wmode == "ADD" else "수정"
        if self.wmode == "MOD":
            new_id, new_type = self.motor_id, self.motor_type
        else:
            t = self.motor_type_var.get()
            if t not in (0, 1, 2):
                return self.error("모터 종류을 선택해 주세요.", "오류")
            new_type = MotorType(t)
            txt = self.txtID.get()
            if txt == "" or not check_integer(txt):
                return self.error("모터 ID값을 숫자로 입력해주세요.", "오류")
            new_id = try_parse_int(txt) or 0
            lo, hi = id_range(new_type)
            if new_id < lo or new_id > hi:
                return self.error(f"모터ID값은 숫자로 {lo}-{hi}범위로 입력하셔야 합니다.", "오류")

        u = self.use_var.get()
        if u not in (0, 1):
            return self.error("사용 여부을 선택해 주세요.", "오류")
        new_use = u == 1

        use_limit = self.cbUSELIMIT.current() == 0
        if self.txtLIMITRPM.get() == "":
            self.txtLIMITRPM.insert(0, "0")
        limit_t = self.txtLIMITRPM.get()
        if not check_integer(limit_t):
            return self.error(f"제한모드 RPM값을 숫자 0~{MAX_RPM}범위로 입력해주세요.", "오류")
        limit_rpm = try_parse_int(limit_t) or 0
        # 운영 프로그램은 여기서 모터 ID 를 검사하는 오류가 있어 범위 밖 값이 저장됨 -> 제한값을 검사하도록 수정
        if limit_rpm < 0 or limit_rpm > MAX_RPM:
            return self.error(f"제한모드 RPM값은 숫자 0~{MAX_RPM}범위로 입력하셔야 합니다.", "오류")

        cip, cport, set_rpm = "", MODBUS_PORT, None
        ctrl_mode = CTRL_MODE_MANUAL
        if new_type == MotorType.OLD_CTRL:
            cip = self.txtCIP.get().strip()
            if cip == "":
                return self.error("기존모터제어 접속 IP를 입력해 주세요.", "오류")
            if not check_valid_ip(cip):
                return self.error("기존모터제어 접속 IP가 올바른 형식이 아닙니다..", "오류")
            if self.wmode == "MOD":
                cport = getattr(self, "ctrl_port", MODBUS_PORT)
            ctrl_mode = self.ctrl_mode_var.get()
            sr_t = self.txtSETRPM.get().strip()
            set_rpm = try_parse_int(sr_t) if check_integer(sr_t) and sr_t else None
            if ctrl_mode == CTRL_MODE_REF:
                if self.plant.reference_endpoint(self.house_code) == (cip, cport, new_id):
                    return self.error("구역 참조값이 이 모터 자신의 회전수입니다(구역설정의 참조기존모터). "
                                      "수동 지정으로 하거나 구역설정에서 다른 참조기존모터를 지정해 주세요.", "오류")
                if set_rpm is None or set_rpm < 0 or set_rpm > MAX_RPM:
                    set_rpm = 0                  # 수동 지정 값은 구역 참조 방식에서는 쓰지 않음
                _desc, value = self.plant.houses[self.house_code].zone_reference()
                how = (f"이 구역 신규모터가 따르는 값({_desc})을 최고속도 {MAX_RPM} RPM 대비 %로 바꿔 "
                       f"같은 비율로({self._now_text(value)})")
            else:
                if set_rpm is None or set_rpm < 0 or set_rpm > MAX_RPM:
                    return self.error(f"운전속도를 숫자 0~{MAX_RPM}범위로 입력해주세요.", "오류")
                how = f"지정한 속도({set_rpm} RPM = {pct_text(rpm_to_pct(set_rpm))})로"
            question = (f"해당 모터정보를 {msg}하시겠습니까?\n\n"
                        f"구역 제어를 시작하면 이 모터는 아날로그 제어신호 대신 통신으로 받은 속도로, "
                        f"{how} 계속 운전합니다.")
        else:
            question = f"해당 모터정보를 {msg}하시겠습니까?"

        if not self.ask(question, msg):
            return
        if new_type != MotorType.NEW_MOTOR:
            use_limit, limit_rpm = False, 0
        house = self.plant.houses[self.house_code]
        dkey = motor_key(new_type, new_id, cip, cport)
        if self.wmode == "ADD":
            if house.get_motor(dkey) is not None:
                return self.error("이미 해당 모터 종류와 ID를 갖는 모터가 있습니다. 확인바랍니다.", "오류")
            house.add_motor(new_id, new_type, new_use, use_limit, limit_rpm, cip, cport, set_rpm or 0, ctrl_mode)
        elif new_type == MotorType.OLD_CTRL:
            house.mod_motor(self.mod_dkey or dkey, new_use, use_limit, limit_rpm, set_rpm=set_rpm,
                            ctrl_mode=ctrl_mode)
        else:
            house.mod_motor(self.mod_dkey or dkey, new_use, use_limit, limit_rpm)
        self.plant.save()
        if self.on_changed:
            self.on_changed()
        self.destroy()

    def on_delete(self) -> None:
        """운영 BtnDelete_Click"""
        question = "해당 모터정보를 삭제하시겠습니까?"
        if self.motor_type == MotorType.OLD_CTRL:
            question += "\n\n삭제해도 모터는 마지막으로 받은 속도로 계속 운전합니다(통신 제어모드 유지)."
        if not self.ask(question, "삭제"):
            return
        if self.motor_type == MotorType.OLD_MOTOR:
            if self.plant.find_reference_users(self.house_code, self.motor_id):
                return self.error("해당 모터를 참조로 사용하는 구역이 있습니다. 변경 후, 삭제할 수 있습니다.", "오류")
        self.plant.houses[self.house_code].remove_motor(self.mod_dkey or "")
        self.plant.save()
        if self.on_changed:
            self.on_changed()
        self.destroy()
