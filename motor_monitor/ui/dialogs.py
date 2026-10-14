"""대화상자: 구역설정 창(운영 wHOUSE), 모터 추가/수정 창(운영 wMotor).

컴포넌트 이름(속성명)은 스펙시트 docs/SPEC_SHEET.md 의 '코드 객체명' 과 같습니다.
"""
from __future__ import annotations

import re
import tkinter as tk
from tkinter import messagebox, ttk
from typing import TYPE_CHECKING, Callable, List, Optional, Tuple

from ..constants import MAX_RPM, MOTOR_ID_MAX, MOTOR_ID_MIN, MotorType
from ..helpers import check_integer, check_valid_ip, get_house_name
from ..house import motor_key
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


class BaseDialog(tk.Toplevel):
    """원본 대화상자 공통 틀: 색 머리글(끌어서 이동) + 본문 + 아래쪽 버튼줄, 모달."""

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
        owner = self.master
        self.update_idletasks()
        try:
            ox, oy = owner.winfo_rootx(), owner.winfo_rooty()
            ow, oh = owner.winfo_width(), owner.winfo_height()
            x = ox + (ow - self._dlg_w) // 2
            y = oy + (oh - self._dlg_h) // 2
        except tk.TclError:
            x, y = 100, 100
        self.geometry(f"{self._dlg_w}x{self._dlg_h}+{max(x, 0)}+{max(y, 0)}")
        self.deiconify()
        self.lift()
        try:
            self.grab_set()
        except tk.TclError:
            pass
        self.focus_force()
        self.wait_window(self)

    def error(self, msg: str, title: str) -> None:
        messagebox.showerror(title, msg, parent=self)

    def ask(self, msg: str, title: str) -> bool:
        return messagebox.askyesno(title, msg, parent=self)


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
        r = self.row("기존모터 IP/PORT")
        self.txtSIP = self.entry(r, 120)
        tk.Label(r, text="/", bg=T.WHITE).pack(side="left", padx=4)
        self.txtSPORT = self.entry(r, 60)
        r = self.row("신규모터 IP/PORT")
        self.txtTIP = self.entry(r, 120)
        tk.Label(r, text="/", bg=T.WHITE).pack(side="left", padx=4)
        self.txtTPORT = self.entry(r, 60)
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
        for e, v in ((self.txtSIP, cfg.source_ip), (self.txtSPORT, cfg.source_port),
                     (self.txtTIP, cfg.target_ip), (self.txtTPORT, cfg.target_port),
                     (self.txtMRPM, cfg.manual_rpm)):
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
        sport_t = self.txtSPORT.get()
        if sport_t == "":
            return self.error("기존모터 접속 PORT를 입력해 주세요.", "에러")
        if not check_integer(sport_t):
            return self.error("기존모터 접속 PORT를 숫자로 입력해 주세요.", "에러")
        sport = try_parse_int(sport_t) or 0
        if sport < 1 or sport > 65535:
            return self.error("기존모터 접속 PORT를 숫자로 1~65535범위로 입력해 주세요.", "에러")

        tip = self.txtTIP.get()
        if tip == "":
            return self.error("신규 모터 접속 IP를 입력해 주세요.", "에러")
        if not check_valid_ip(tip):
            return self.error("신규모터 접속 IP가 올바른 형식이 아닙니다..", "에러")
        tport_t = self.txtTPORT.get()
        if tport_t == "":
            return self.error("신규모터 접속 PORT를 입력해 주세요.", "에러")
        if not check_integer(tport_t):
            return self.error("신규모터 접속 PORT를 숫자로 입력해 주세요.", "에러")
        tport = try_parse_int(tport_t) or 0
        if tport < 1 or tport > 65535:
            return self.error("신규모터 접속 PORT를 숫자로 1~65535범위로 입력해 주세요.", "에러")

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
            source_ip=sip, source_port=sport, target_ip=tip, target_port=tport,
            source_manual=smanual, manual_rpm=mrpm, source_line=self.house_header + sline, source_id=sid)
        self.plant.set_line_config(self.house_code, cfg)
        if self.on_saved:
            self.on_saved()
        self.destroy()


# ======================================================================== 모터 추가/수정 창
class MotorDialog(BaseDialog):
    """DLG-MOTOR 모터 추가(ADD) / 모터 수정·삭제(MOD) 창. 운영 wMotor"""

    LIMIT_CHOICES = ("사용함", "사용안함")

    def __init__(self, owner: tk.Misc, plant: "Plant", images: T.Images, house_code: str,
                 mode: str = "ADD", mod_dkey: Optional[str] = None,
                 on_changed: Optional[Callable[[], None]] = None) -> None:
        mode = mode if mode in ("ADD", "MOD") else "ADD"
        title = "모터 추가" if mode == "ADD" else "모터 수정/삭제"
        super().__init__(owner, title, T.MOTOR_DIALOG_HEADER, 400, 300)
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
                                       self.motor_type_var, int(MotorType.OLD_MOTOR), self._on_type)
        self.rb_newmotor = self._radio(self.add_type_line, "신규모터", "newmotor",
                                       self.motor_type_var, int(MotorType.NEW_MOTOR), self._on_type)
        self.rb_newmotor.pack_configure(padx=(8, 0))
        self.mod_type_line = tk.Frame(r, bg=T.WHITE)
        self.img_type = tk.Label(self.mod_type_line, bg=T.WHITE)
        self.img_type.pack(side="left", padx=(0, 5))
        self.lbl_type = tk.Label(self.mod_type_line, text="type", bg=T.WHITE, font=T.font(11))
        self.lbl_type.pack(side="left")
        self.separator()
        r = self.row("모터ID", title_size=12)
        self.add_id_line = tk.Frame(r, bg=T.WHITE)
        self.txtID = self.entry(self.add_id_line, 60, size=11, bold=False)
        tk.Label(self.add_id_line, text=f"정수({MOTOR_ID_MIN}~{MOTOR_ID_MAX})", bg=T.WHITE).pack(side="left", padx=10)
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

    def _radio(self, parent, text: str, image: str, var: tk.IntVar, value: int, command=None) -> tk.Radiobutton:
        rb = tk.Radiobutton(parent, text=" " + text, image=self.images.get(image, 20), compound="left",
                            variable=var, value=value, bg=T.WHITE, activebackground=T.WHITE,
                            font=T.font(11, bold=True), highlightthickness=0, command=command)
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

    def _on_type(self) -> None:
        """운영 rb_oldmotor_Checked / rb_newmotor_Checked"""
        self.show_limit_row(self.motor_type_var.get() == int(MotorType.NEW_MOTOR))

    def _load(self) -> None:
        """운영 wMotor_Loaded"""
        if self.wmode == "ADD":
            self.BtnSave.set_text("추가")
            self.add_type_line.pack(side="left")
            self.add_id_line.pack(side="left")
            self.show_limit_row(False)
            self.txtID.delete(0, "end")
            self.txtLIMITRPM.insert(0, "0")
            self.after(50, self.txtID.focus_set)
        else:
            self.BtnDelete.pack(side="left")
            self.BtnSave.set_text("수정")
            self.mod_type_line.pack(side="left")
            self.mod_id_line.pack(side="left")
            m = self.plant.houses[self.house_code].get_motor(self.mod_dkey or "")
            if m is not None:
                old = m.mtype == MotorType.OLD_MOTOR
                self.img_type.configure(image=self.images.get("oldmotor" if old else "newmotor", 20))
                self.lbl_type.configure(text="기존모터" if old else "신규모터")
                self.motor_id = m.id
                self.motor_type = m.mtype
                self.lblID.configure(text=str(m.id))
                self.use_var.set(1 if m.use else 0)
                self.cbUSELIMIT.current(0 if m.use_limit else 1)
                self.txtLIMITRPM.delete(0, "end")
                self.txtLIMITRPM.insert(0, str(m.limit_rpm))
                self.show_limit_row(not old)
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
            if t not in (0, 1):
                return self.error("모터 종류을 선택해 주세요.", "오류")
            new_type = MotorType(t)
            txt = self.txtID.get()
            if txt == "" or not check_integer(txt):
                return self.error("모터 ID값을 숫자로 입력해주세요.", "오류")
            new_id = try_parse_int(txt) or 0
            if new_id < MOTOR_ID_MIN or new_id > MOTOR_ID_MAX:
                return self.error(f"모터ID값은 숫자로 {MOTOR_ID_MIN}-{MOTOR_ID_MAX}범위로 입력하셔야 합니다.", "오류")

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

        if not self.ask(f"해당 모터정보를 {msg}하시겠습니까?", msg):
            return
        if new_type == MotorType.OLD_MOTOR:
            use_limit, limit_rpm = False, 0
        house = self.plant.houses[self.house_code]
        dkey = motor_key(new_type, new_id)
        if self.wmode == "ADD":
            if house.get_motor(dkey) is not None:
                return self.error("이미 해당 모터 종류와 ID를 갖는 모터가 있습니다. 확인바랍니다.", "오류")
            house.add_motor(new_id, new_type, new_use, use_limit, limit_rpm)
        else:
            house.mod_motor(dkey, new_use, use_limit, limit_rpm)
        self.plant.save()
        if self.on_changed:
            self.on_changed()
        self.destroy()

    def on_delete(self) -> None:
        """운영 BtnDelete_Click"""
        if not self.ask("해당 모터정보를 삭제하시겠습니까?", "삭제"):
            return
        if self.motor_type == MotorType.OLD_MOTOR:
            if self.plant.find_reference_users(self.house_code, self.motor_id):
                return self.error("해당 모터를 참조로 사용하는 구역이 있습니다. 변경 후, 삭제할 수 있습니다.", "오류")
        self.plant.houses[self.house_code].remove_motor(self.mod_dkey or "")
        self.plant.save()
        if self.on_changed:
            self.on_changed()
        self.destroy()
