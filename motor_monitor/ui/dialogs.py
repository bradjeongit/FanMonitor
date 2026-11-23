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

from ..constants import (CTRL_MODE_MANUAL, CTRL_MODE_REF, DEFAULT_CTRL_RPM, DEFAULT_MANUAL_RPM, MAX_RPM,
                         MODBUS_PORT, MOTOR_ID_MAX, MOTOR_ID_MIN, WD_SECONDS, WD_SPEED, MotorStatus, MotorType,
                         id_range, pct_text, rpm_to_pct)
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
    """DLG-LINE 구역설정 창. 운영 wHOUSE
    2.7.0: 기존모터 IP 를 2개(IP 2 는 비워 둘 수 있음)까지 등록. RPM 제어 소스는 모터마다 정하므로 이 창에서 뺌."""

    def __init__(self, owner: tk.Misc, plant: "Plant", house_code: str,
                 on_saved: Optional[Callable[[], None]] = None) -> None:
        super().__init__(owner, "구역설정", T.LINE_DIALOG_HEADER, 400, 300)
        self.plant = plant
        self.house_code = house_code
        self.on_saved = on_saved

        r = self.row("위치")
        self.lbl_house = tk.Label(r, text="-", bg=T.WHITE, font=T.font(11, bold=True))
        self.lbl_house.pack(side="left")
        self.separator()
        r = self.row("기존모터 IP 1")       # PORT 는 502 고정이라 입력칸 없음
        self.txtSIP = self.entry(r, 140)
        r = self.row("기존모터 IP 2")
        self.txtSIP2 = self.entry(r, 140)
        tk.Label(r, text="(없으면 비움)", bg=T.WHITE, fg="#7a7a7a", font=T.font(8)).pack(side="left", padx=6)
        r = self.row("신규모터 IP")
        self.txtTIP = self.entry(r, 140)
        self.separator()
        self.lbl_note = tk.Label(self.body, bg=T.WHITE, fg="#666666", font=T.font(8), anchor="w", justify="left",
                                 wraplength=370,
                                 text="RPM 제어 소스(수동 / 참조 기존모터)와 모니터링만 설정은 모터마다 "
                                      "모터 추가·수정 창에서 정합니다.")
        self.lbl_note.pack(side="top", fill="x", padx=20, pady=(8, 0))

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
        for e, v in ((self.txtSIP, cfg.source_ip), (self.txtSIP2, cfg.source_ip2), (self.txtTIP, cfg.target_ip)):
            e.delete(0, "end")
            e.insert(0, str(v))

    def on_default(self) -> None:
        self.on_save()

    def on_save(self) -> None:
        """운영 BtnSave_Click"""
        sip = self.txtSIP.get().strip()
        if sip == "":
            return self.error("기존모터 접속 IP 1을 입력해 주세요.", "에러")
        if not check_valid_ip(sip):
            return self.error("기존모터 접속 IP 1이 올바른 형식이 아닙니다..", "에러")
        sip2 = self.txtSIP2.get().strip()
        if sip2:
            if not check_valid_ip(sip2):
                return self.error("기존모터 접속 IP 2가 올바른 형식이 아닙니다..", "에러")
            if sip2 == sip:
                return self.error("기존모터 접속 IP 2가 IP 1과 같습니다. 다른 IP를 넣거나 비워 주세요.", "에러")
        elif self.plant.houses[self.house_code].get_old_motor_ids(gw=2):
            return self.error("기존모터 IP 2에 등록된 기존모터가 있어 IP 2를 비울 수 없습니다. "
                              "그 모터를 먼저 삭제해 주세요.", "에러")
        tip = self.txtTIP.get().strip()
        if tip == "":
            return self.error("신규 모터 접속 IP를 입력해 주세요.", "에러")
        if not check_valid_ip(tip):
            return self.error("신규모터 접속 IP가 올바른 형식이 아닙니다..", "에러")

        cfg = self.plant.configs[self.house_code].copy(
            source_ip=sip, source_port=MODBUS_PORT, source_ip2=sip2, target_ip=tip, target_port=MODBUS_PORT)
        self.plant.set_line_config(self.house_code, cfg)
        if self.on_saved:
            self.on_saved()
        self.destroy()


# ======================================================================== 모터 추가/수정 창
class MotorDialog(BaseDialog):
    """DLG-MOTOR 모터 추가(ADD) / 모터 수정·삭제(MOD) 창. 운영 wMotor

    종류별 줄 (2.7.0)
      기존모터    : 기존모터 IP (구역의 IP 1 / IP 2 중 드롭다운으로 선택, 수정 창에서는 바꿀 수 없음)
      신규모터    : 속도제한, 운전속도(수동 RPM / 참조 기존모터), 모니터링만
      기존모터제어: 접속 IP, 운전속도(수동 지정 / 참조 기존모터), 모니터링만
    """

    LIMIT_CHOICES = ("사용함", "사용안함")
    BASE_H = 340

    def __init__(self, owner: tk.Misc, plant: "Plant", images: T.Images, house_code: str,
                 mode: str = "ADD", mod_dkey: Optional[str] = None,
                 on_changed: Optional[Callable[[], None]] = None) -> None:
        mode = mode if mode in ("ADD", "MOD") else "ADD"
        title = "모터 추가" if mode == "ADD" else "모터 수정/삭제"
        # 기존모터제어(3번째 종류)와 그 입력줄을 위해 운영 프로그램(400x300)보다 크게. 종류에 따라 높이가 바뀜
        super().__init__(owner, title, T.MOTOR_DIALOG_HEADER, 560, self.BASE_H)
        self.plant = plant
        self.images = images
        self.house_code = house_code
        self.house_header = house_code[:2]
        self.wmode = mode
        self.mod_dkey = mod_dkey
        self.on_changed = on_changed
        self.motor_id = 0
        self.motor_type = MotorType.OLD_MOTOR
        self.motor_gw = 1
        self._ref_tags: List[Tuple[int, int]] = []

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

        # 기존모터 IP 줄 (기존모터만): 구역설정의 기존모터 IP 1 / IP 2 중 선택
        self.gw_holder = tk.Frame(self.body, bg=T.WHITE)
        self.gw_holder.pack(side="top", fill="x")
        self.gw_row = self._type_row(self.gw_holder, "기존모터 IP")
        self.cb_gw = ttk.Combobox(self.gw_row.inner, state="readonly", width=26, font=T.font(10))
        self.cb_gw.pack(side="left")
        self.gw_row_line = DottedSeparator(self.gw_holder)

        # 속도제한 줄 (신규모터만)
        self.limit_holder = tk.Frame(self.body, bg=T.WHITE)
        self.limit_holder.pack(side="top", fill="x")
        self.limit_row = self._type_row(self.limit_holder, "속도제한")
        inner = self.limit_row.inner
        self.cbUSELIMIT = ttk.Combobox(inner, state="readonly", width=7, values=self.LIMIT_CHOICES,
                                       font=T.font(10))
        self.cbUSELIMIT.current(1)                 # 기본 '사용안함'
        self.cbUSELIMIT.pack(side="left")
        tk.Label(inner, text="최대속도:", bg=T.WHITE).pack(side="left", padx=(10, 2))
        self.txtLIMITRPM = self.entry(inner, 65, size=11, bold=False)
        tk.Label(inner, text="RPM", bg=T.WHITE).pack(side="left", padx=2)
        self.limit_row_line = DottedSeparator(self.limit_holder)

        # 기존모터제어 접속 IP 줄 (기존모터제어만). PORT 는 502 고정이라 입력칸 없음
        self.ctrl_holder = tk.Frame(self.body, bg=T.WHITE)
        self.ctrl_holder.pack(side="top", fill="x")
        self.ctrl_ip_row = self._type_row(self.ctrl_holder, "접속 IP")
        self.txtCIP = self.entry(self.ctrl_ip_row.inner, 140, size=11, bold=False)

        # 운전속도 줄 (신규모터·기존모터제어): 수동 지정 / 참조 기존모터 + 모니터링만
        self.speed_holder = tk.Frame(self.body, bg=T.WHITE)
        self.speed_holder.pack(side="top", fill="x")
        self.ctrl_speed_row = self._type_row(self.speed_holder, "운전속도")
        inner = self.ctrl_speed_row.inner
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
        self.rb_ctrl_ref = tk.Radiobutton(inner, text="참조", value=CTRL_MODE_REF,
                                          variable=self.ctrl_mode_var, command=self._on_ctrl_mode,
                                          bg=T.WHITE, activebackground=T.WHITE, font=T.font(10, bold=True),
                                          highlightthickness=0)
        self.rb_ctrl_ref.pack(side="left")
        self.ctrl_ref_row = self._type_row(self.speed_holder, "┖ 참조 기존모터", size=10)
        inner = self.ctrl_ref_row.inner
        self.cb_ref_line = ttk.Combobox(inner, state="readonly", width=7, height=len(LINE_CHOICES),
                                        font=T.font(10, bold=True), values=[name for _, name in LINE_CHOICES])
        self.cb_ref_line.pack(side="left")
        self.cb_ref_line.bind("<<ComboboxSelected>>", lambda e: self._on_ref_line_changed())
        tk.Label(inner, text=" > ", bg=T.WHITE).pack(side="left")
        self.cb_ref_motor = ttk.Combobox(inner, state="readonly", width=11, font=T.font(10, bold=True))
        self.cb_ref_motor.pack(side="left")
        self.cb_ref_motor.bind("<<ComboboxSelected>>", lambda e: self._show_ref_now())
        self.lbl_zone_ref = tk.Label(inner, text="", bg=T.WHITE, fg=T.TEXT_DARK, font=T.font(9), anchor="w")
        self.lbl_zone_ref.pack(side="left", padx=(8, 0))
        self.ctrl_ref_note = tk.Label(self.speed_holder, bg=T.WHITE, fg="#666666", font=T.font(8), anchor="w",
                                      justify="left", wraplength=500, text="")
        # 모터 최고속도(h8) 기준 % (기존모터제어 수정 창, 구역 제어 중 h8 을 읽은 뒤에만 표시)
        self.motor_max_rpm = 0
        self.lbl_motor_max = tk.Label(self.speed_holder, bg=T.WHITE, fg="#1f4e79", font=T.font(8, bold=True),
                                      anchor="w", justify="left", text="")
        self.mon_row = self._type_row(self.speed_holder, "모니터링만")
        self.mon_var = tk.IntVar(value=0)
        self.chk_mon = tk.Checkbutton(self.mon_row.inner, text="이 모터는 읽기만 (전체 모드와 상관없이 제어 신호 안 보냄)",
                                      variable=self.mon_var, command=self._on_ctrl_mode, bg=T.WHITE,
                                      activebackground=T.WHITE, font=T.font(9, bold=True), highlightthickness=0)
        self.chk_mon.pack(side="left")
        self.ctrl_row_line = DottedSeparator(self.speed_holder)

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

    def _type_row(self, parent: tk.Frame, title: str, size: int = 12) -> tk.Frame:
        """종류별 줄 (높이 40, 제목 + 입력 영역 .inner). 보이기·숨기기는 show_* 가 함."""
        row = tk.Frame(parent, bg=T.WHITE, height=40)
        row.pack_propagate(False)
        tk.Label(row, text=title, bg=T.WHITE, fg=T.TEXT_DARK, font=T.font(size, bold=True)).place(
            x=10, rely=0.5, anchor="w")
        row.inner = tk.Frame(row, bg=T.WHITE)
        row.inner.place(x=130, rely=0.5, anchor="w")
        return row

    def _radio(self, parent, text: str, image: str, var: tk.IntVar, value: int, command=None,
               size: int = 11) -> tk.Radiobutton:
        rb = tk.Radiobutton(parent, text=" " + text, image=self.images.get(image, 20), compound="left",
                            variable=var, value=value, bg=T.WHITE, activebackground=T.WHITE,
                            font=T.font(size, bold=True), highlightthickness=0, command=command)
        rb.pack(side="left")
        return rb

    # --- 종류별 줄 보이기 ---
    def _resize(self) -> None:
        """보이는 줄에 맞춰 창 높이 조정."""
        self.update_idletasks()
        self._dlg_h = max(self.BASE_H, 40 + self.body.winfo_reqheight() + 50 + 2)
        if self.winfo_ismapped():
            self.keep_inside_owner()          # 크기가 바뀌어도 기존 창 안에 머묾

    def show_gw_row(self, show: bool) -> None:
        if show:
            self.gw_row.pack(side="top", fill="x", padx=(10, 10))
            self.gw_row_line.pack(side="top", fill="x")
        else:
            self.gw_row.pack_forget()
            self.gw_row_line.pack_forget()

    @property
    def gw_row_visible(self) -> bool:
        return bool(self.gw_row.winfo_manager())

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
        """기존모터제어 접속 IP 줄."""
        if show:
            self.ctrl_ip_row.pack(side="top", fill="x", padx=(10, 10))
        else:
            self.ctrl_ip_row.pack_forget()

    @property
    def ctrl_rows_visible(self) -> bool:
        return bool(self.ctrl_ip_row.winfo_manager())

    def show_speed_rows(self, show: bool) -> None:
        """운전속도·참조 기존모터·모니터링만 줄 (신규모터·기존모터제어)."""
        for w in (self.ctrl_speed_row, self.ctrl_ref_row, self.ctrl_ref_note, self.lbl_motor_max, self.mon_row,
                  self.ctrl_row_line):
            w.pack_forget()
        if not show:
            return
        self.ctrl_speed_row.pack(side="top", fill="x", padx=(10, 10))
        self.ctrl_ref_row.pack(side="top", fill="x", padx=(10, 10))
        if self.ctrl_ref_note.cget("text"):
            self.ctrl_ref_note.pack(side="top", fill="x", padx=(30, 10))
        if self.motor_max_rpm > 0:
            self.lbl_motor_max.pack(side="top", fill="x", padx=(30, 10), pady=(2, 0))
        self.mon_row.pack(side="top", fill="x", padx=(10, 10))
        self.ctrl_row_line.pack(side="top", fill="x", pady=(4, 0))
        self._on_ctrl_mode()

    @property
    def speed_rows_visible(self) -> bool:
        return bool(self.ctrl_speed_row.winfo_manager())

    def _apply_type(self, t: MotorType) -> None:
        ctrl = t == MotorType.OLD_CTRL
        self.show_gw_row(t == MotorType.OLD_MOTOR)
        self.show_limit_row(t == MotorType.NEW_MOTOR)
        self.show_ctrl_rows(ctrl)
        self.rb_ctrl_manual.configure(text="수동 지정" if ctrl else "수동 RPM")
        self.ctrl_ref_note.configure(
            text=(f"참조: 고른 기존모터의 회전수를 최고속도 {MAX_RPM} RPM 대비 %로 바꿔 같은 비율로 운전합니다. "
                  "값이 없으면 이전 속도를 유지합니다.") if ctrl else
            "참조: 고른 기존모터의 회전수를 따라 운전합니다 (그 모터가 정상이 아니면 속도를 바꾸지 않음).")
        self.show_speed_rows(t in (MotorType.NEW_MOTOR, MotorType.OLD_CTRL))
        self._resize()

    # --- 기존모터 IP 선택 ---
    def _gw_choices(self) -> List[Tuple[int, str]]:
        cfg = self.plant.configs[self.house_code]
        out = [(1, f"IP 1  ({cfg.source_ip})")]
        if cfg.source_ip2:
            out.append((2, f"IP 2  ({cfg.source_ip2})"))
        return out

    def _load_gw(self, gw: int = 1) -> None:
        choices = self._gw_choices()
        self.cb_gw.configure(values=[name for _, name in choices])
        gws = [g for g, _ in choices]
        self.cb_gw.current(gws.index(gw) if gw in gws else 0)

    def selected_gw(self) -> int:
        i = self.cb_gw.current()
        choices = self._gw_choices()
        return choices[i][0] if 0 <= i < len(choices) else 1

    # --- 참조 기존모터 선택 ---
    def _select_ref_line(self, code: str) -> None:
        tags = [t for t, _ in LINE_CHOICES]
        tag = code[2:4] if len(code) == 4 else ""
        self.cb_ref_line.current(tags.index(tag) if tag in tags else 0)

    def _ref_line(self) -> str:
        i = self.cb_ref_line.current()
        tag = LINE_CHOICES[i][0] if i >= 0 else ""
        return self.house_header + tag if tag else ""

    def _load_ref_motors(self, code: str, default: Optional[Tuple[int, int]] = None) -> None:
        """구역 code 의 기존모터 목록 ('155', 통신장치 2 는 '7 (IP 2)'). 첫 항목은 'ID'."""
        house = self.plant.houses.get(code) if code else None
        self._ref_tags = house.get_old_motors() if house is not None else []
        names = ["ID"] + [f"{mid}" + (" (IP 2)" if gw == 2 else "") for gw, mid in self._ref_tags]
        self.cb_ref_motor.configure(values=names)
        self.cb_ref_motor.current(self._ref_tags.index(default) + 1 if default in self._ref_tags else 0)

    def selected_ref(self) -> Optional[Tuple[str, int, int]]:
        """(참조 구역, 통신장치 번호, ID). 고르지 않았으면 None."""
        i = self.cb_ref_motor.current()
        line = self._ref_line()
        if not line or not 1 <= i <= len(self._ref_tags):
            return None
        gw, mid = self._ref_tags[i - 1]
        return line, gw, mid

    def _on_ref_line_changed(self) -> None:
        self._load_ref_motors(self._ref_line())
        self._show_ref_now()

    def _ref_value(self) -> Optional[int]:
        """고른 참조 기존모터의 지금 회전수 (정상일 때만)."""
        ref = self.selected_ref()
        if ref is None:
            return None
        house = self.plant.houses.get(ref[0])
        sm = house.get_motor(motor_key(MotorType.OLD_MOTOR, ref[2], gw=ref[1])) if house else None
        return sm.rpm if sm is not None and sm.status == MotorStatus.OK else None

    def _show_ref_now(self) -> None:
        self.lbl_zone_ref.configure(text=self._now_text(self._ref_value()) if self.selected_ref() else "")
        self._show_set_pct()

    @staticmethod
    def _now_text(value: Optional[int]) -> str:
        if value is None:
            return "지금 값 없음 (구역을 실행하면 읽음)"
        return f"지금 {value} RPM = {pct_text(rpm_to_pct(value))}"

    # --- 운전속도 ---
    def _on_ctrl_mode(self) -> None:
        ref = self.ctrl_mode_var.get() == CTRL_MODE_REF
        mon = self.mon_var.get() == 1
        self.txtSETRPM.configure(state="disabled" if ref else "normal", disabledbackground=T.LIGHT_GRAY)
        for w in (self.cb_ref_line, self.cb_ref_motor):
            w.configure(state="readonly" if ref else "disabled")
        self.lbl_zone_ref.configure(fg=T.TEXT_DARK if ref else "#9a9a9a")
        for w in (self.rb_ctrl_manual, self.rb_ctrl_ref):
            w.configure(fg="#9a9a9a" if mon else T.TEXT)
        self._show_set_pct()

    def _show_set_pct(self) -> None:
        """수동 지정 RPM 옆에 1100 대비 % 표시 (+ 모터 최고속도를 알면 그 대비 %)"""
        v = try_parse_int(self.txtSETRPM.get())
        ok = v is not None and 0 <= v <= MAX_RPM
        pct = f" = {pct_text(rpm_to_pct(v))}" if ok else ""
        self.lbl_set_pct.configure(text=f"RPM (0~{MAX_RPM}){pct}")
        if self.motor_max_rpm > 0:
            ref = self.ctrl_mode_var.get() == CTRL_MODE_REF
            want = self._ref_value() if ref else (v if ok else None)
            text = f"모터 최고속도(h8) {self.motor_max_rpm} RPM"
            if want is not None:
                text += f": 원하는 속도 {want} RPM = 최고속도의 {pct_text(want * 100 / self.motor_max_rpm)}"
                if want > self.motor_max_rpm:
                    text += " (최고속도보다 높아 최고속도로 운전)"
            self.lbl_motor_max.configure(text=text)

    @staticmethod
    def _set_text(e: tk.Entry, value) -> None:
        """입력칸 값 바꾸기 (비활성 칸도 바꿀 수 있게 잠시 활성화)."""
        state = str(e.cget("state"))
        e.configure(state="normal")
        e.delete(0, "end")
        e.insert(0, str(value))
        e.configure(state=state)

    def _on_type(self) -> None:
        """운영 rb_oldmotor_Checked / rb_newmotor_Checked + 기존모터제어"""
        t = MotorType(self.motor_type_var.get())
        self._apply_type(t)
        lo, hi = id_range(t)
        self.lbl_id_range.configure(text=f"정수({lo}~{hi})")
        # 새 모터 기본 속도: 신규모터 수동 770 RPM, 기존모터제어 수동 지정 70% (= 770 RPM)
        self._set_text(self.txtSETRPM, DEFAULT_CTRL_RPM if t == MotorType.OLD_CTRL else DEFAULT_MANUAL_RPM)
        self._show_set_pct()

    def _load(self) -> None:
        """운영 wMotor_Loaded"""
        self._select_ref_line(self.house_code)       # 참조 기존모터: 기본은 이 구역
        self._load_ref_motors(self.house_code)
        if self.wmode == "ADD":
            self.BtnSave.set_text("추가")
            self.add_type_line.pack(side="left")
            self.add_id_line.pack(side="left")
            self._load_gw(1)
            self.txtID.delete(0, "end")
            self.txtLIMITRPM.insert(0, "0")
            cfg = self.plant.configs[self.house_code]       # 기존모터 게이트웨이 주소를 기본값으로
            self.txtCIP.insert(0, cfg.source_ip)
            self.ctrl_mode_var.set(CTRL_MODE_MANUAL)              # 기본: 수동 770 RPM (기존모터제어 70%)
            self._set_text(self.txtSETRPM, DEFAULT_MANUAL_RPM)
            self.show_gw_row(False)
            self.show_limit_row(False)
            self.show_ctrl_rows(False)
            self.show_speed_rows(False)
            self._resize()
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
                self.motor_gw = m.gw
                self.lblID.configure(text=str(m.id))
                self.use_var.set(1 if m.use else 0)
                self.cbUSELIMIT.current(0 if m.use_limit else 1)
                self.txtLIMITRPM.delete(0, "end")
                self.txtLIMITRPM.insert(0, str(m.limit_rpm))
                self._load_gw(m.gw)
                self.cb_gw.configure(state="disabled")          # 통신장치는 수정 불가 (삭제 후 다시 추가)
                if ctrl:
                    self.txtCIP.insert(0, m.ip)
                    self.txtCIP.configure(state="readonly", readonlybackground=T.WHITE)  # 접속처는 수정 불가
                    self.ctrl_ip = m.ip
                    self.ctrl_port = m.port
                    self.motor_max_rpm = m.max_rpm
                if not old:
                    self._set_text(self.txtSETRPM, m.set_rpm)
                    self.ctrl_mode_var.set(m.ctrl_mode)
                    self.mon_var.set(1 if m.mon_only else 0)
                    if m.ref_line:
                        self._select_ref_line(m.ref_line)
                        self._load_ref_motors(m.ref_line, (m.ref_gw, m.ref_id))
                self._apply_type(m.mtype)
                self._show_ref_now()
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
        gw = self.selected_gw() if self.wmode == "ADD" else self.motor_gw
        ref: Optional[Tuple[str, int, int]] = None
        mon_only = False
        wd = (f"\n\n연결되면 '통신 끊김 시 자동 운전'을 확인해 꺼져 있으면 켭니다 "
              f"({WD_SECONDS}초 동안 통신이 없으면 {WD_SPEED} RPM 으로 운전, 모터 영구저장 설정).")
        if new_type == MotorType.OLD_CTRL:
            cip = self.txtCIP.get().strip()
            if cip == "":
                return self.error("기존모터제어 접속 IP를 입력해 주세요.", "오류")
            if not check_valid_ip(cip):
                return self.error("기존모터제어 접속 IP가 올바른 형식이 아닙니다..", "오류")
            if self.wmode == "MOD":
                cport = getattr(self, "ctrl_port", MODBUS_PORT)
        if new_type == MotorType.OLD_MOTOR:
            question = f"해당 모터정보를 {msg}하시겠습니까?\n\n기존모터는 회전수를 읽기만 합니다(모니터링)." + wd
        else:
            ctrl_mode = self.ctrl_mode_var.get()
            mon_only = self.mon_var.get() == 1
            ref = self.selected_ref()
            sr_t = self.txtSETRPM.get().strip()
            set_rpm = try_parse_int(sr_t) if check_integer(sr_t) and sr_t else None
            if ctrl_mode == CTRL_MODE_REF:
                if ref is None:
                    return self.error("참조할 기존모터(구역과 ID)를 선택해 주세요.", "오류")
                if (new_type == MotorType.OLD_CTRL
                        and self.plant.motor_endpoint(*ref) == (cip, cport, new_id)):
                    return self.error("참조 기존모터가 이 모터 자신입니다. "
                                      "수동 지정으로 하거나 다른 참조 기존모터를 지정해 주세요.", "오류")
                if set_rpm is None or set_rpm < 0 or set_rpm > MAX_RPM:
                    set_rpm = DEFAULT_MANUAL_RPM  # 수동 값은 참조 방식에서는 쓰지 않음 (나중에 수동으로 바꿀 때의 값)
                desc = f"{get_house_name(ref[0])} 기존모터{' 2' if ref[1] == 2 else ''} ID {ref[2]}"
                how = (f"참조 기존모터({desc})의 회전수와 같은 비율로({self._now_text(self._ref_value())})"
                       if new_type == MotorType.OLD_CTRL else f"참조 기존모터({desc})의 회전수를 따라")
            else:
                if set_rpm is None or set_rpm < 0 or set_rpm > MAX_RPM:
                    return self.error(f"운전속도를 숫자 0~{MAX_RPM}범위로 입력해주세요.", "오류")
                how = f"지정한 속도({set_rpm} RPM = {pct_text(rpm_to_pct(set_rpm))})로"
            if mon_only:
                question = (f"해당 모터정보를 {msg}하시겠습니까?\n\n모니터링만: 이 모터는 읽기만 하고 "
                            "제어 신호(속도 등)를 보내지 않습니다 (전체 모드와 상관없이).")
            elif new_type == MotorType.OLD_CTRL:
                question = (f"해당 모터정보를 {msg}하시겠습니까?\n\n"
                            f"구역 제어를 시작하면 이 모터는 아날로그 제어신호 대신 통신으로 받은 속도로, "
                            f"{how} 계속 운전합니다.")
            else:
                question = f"해당 모터정보를 {msg}하시겠습니까?\n\n구역 제어 중 {how} 운전합니다."
            if new_type == MotorType.OLD_CTRL:
                question += wd

        if not self.ask(question, msg):
            return
        if new_type != MotorType.NEW_MOTOR:
            use_limit, limit_rpm = False, 0
        house = self.plant.houses[self.house_code]
        dkey = motor_key(new_type, new_id, cip, cport, gw)
        if self.wmode == "ADD":
            if house.get_motor(dkey) is not None:
                return self.error("이미 해당 모터 종류와 ID를 갖는 모터가 있습니다. 확인바랍니다.", "오류")
            r_line, r_gw, r_id = ref or ("", 1, 0)
            house.add_motor(new_id, new_type, new_use, use_limit, limit_rpm, cip, cport,
                            set_rpm if set_rpm is not None else 0, ctrl_mode, gw, r_line, r_gw, r_id, mon_only)
        elif new_type == MotorType.OLD_MOTOR:
            house.mod_motor(self.mod_dkey or dkey, new_use, use_limit, limit_rpm)
        else:
            house.mod_motor(self.mod_dkey or dkey, new_use, use_limit, limit_rpm, set_rpm=set_rpm,
                            ctrl_mode=ctrl_mode, ref=ref, mon_only=mon_only)
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
            users = self.plant.find_reference_users(self.house_code, self.motor_id, self.motor_gw)
            if users:
                return self.error("해당 모터를 참조로 사용하는 모터가 있습니다. 변경 후, 삭제할 수 있습니다.\n"
                                  + ", ".join(users[:5]) + (" 외" if len(users) > 5 else ""), "오류")
        self.plant.houses[self.house_code].remove_motor(self.mod_dkey or "")
        self.plant.save()
        if self.on_changed:
            self.on_changed()
        self.destroy()
