"""공용 화면 위젯 (둥근 버튼, 상태 점, 툴팁, 모터 목록 등)."""
from __future__ import annotations

import time
import tkinter as tk
from typing import Callable, Dict, List, Optional

from ..constants import CTRL_MODE_REF, MotorStatus, MotorType, pct_text, rpm_to_pct, speed_within
from ..house import Motor
from . import theme as T


# ------------------------------------------------------------------ 툴팁
class Tooltip:
    """마우스를 올리면 잠시 후 설명을 보여줌. text 는 문자열 또는 (x, y) -> 문자열 함수."""

    DELAY_MS = 500

    def __init__(self, widget: tk.Widget, text) -> None:
        self.widget = widget
        self.text = text
        self._after = None
        self._tip: Optional[tk.Toplevel] = None
        self._shown_text = None
        self._xy = (0, 0)
        widget.bind("<Enter>", self._on_motion, add="+")
        widget.bind("<Motion>", self._on_motion, add="+")
        widget.bind("<Leave>", self._hide, add="+")
        widget.bind("<ButtonPress>", self._hide, add="+")

    def _resolve(self) -> str:
        if callable(self.text):
            return self.text(*self._xy) or ""
        return self.text or ""

    def _on_motion(self, e) -> None:
        self._xy = (e.x, e.y)
        if self._tip is not None:
            if self._resolve() != self._shown_text:
                self._hide()
            else:
                return
        if self._after:
            self.widget.after_cancel(self._after)
        self._after = self.widget.after(self.DELAY_MS, self._show)

    def _show(self) -> None:
        self._after = None
        text = self._resolve()
        if not text:
            return
        x = self.widget.winfo_pointerx() + 12
        y = self.widget.winfo_pointery() + 18
        tip = tk.Toplevel(self.widget)
        tip.wm_overrideredirect(True)
        tip.wm_geometry(f"+{x}+{y}")
        tk.Label(tip, text=text, bg="#ffffe1", fg="#000000", relief="solid", bd=1,
                 font=T.font(9), padx=4, pady=2).pack()
        self._tip = tip
        self._shown_text = text

    def _hide(self, _e=None) -> None:
        if self._after:
            self.widget.after_cancel(self._after)
            self._after = None
        if self._tip is not None:
            self._tip.destroy()
            self._tip = None
            self._shown_text = None


# ------------------------------------------------------------------ 버튼
def _round_rect(c: tk.Canvas, x1, y1, x2, y2, r, **kw):
    pts = [x1 + r, y1, x2 - r, y1, x2, y1, x2, y1 + r, x2, y2 - r, x2, y2,
           x2 - r, y2, x1 + r, y2, x1, y2, x1, y2 - r, x1, y1 + r, x1, y1]
    return c.create_polygon(pts, smooth=True, **kw)


class RoundButton(tk.Canvas):
    """원본 Normal/Submit/AlertButtonTemplate (모서리 5, 높이 30, 좌우 여백 30, 마우스오버 시 진하게+굵게)."""

    def __init__(self, master, text: str, command: Callable[[], None], style: str = "normal", **kw) -> None:
        self._bg, self._hover, self._fg, self._border = T.BUTTON_STYLES[style]
        self._text = text
        self._command = command
        self._font_n = T.font(9)
        self._font_b = T.font(9, bold=True)
        import tkinter.font as tkfont
        tw = tkfont.Font(font=self._font_b).measure(text)
        super().__init__(master, width=tw + 60, height=30, highlightthickness=0,
                         bg=master.cget("bg"), cursor="hand2", **kw)
        self._draw(False)
        self.bind("<Enter>", lambda e: self._draw(True))
        self.bind("<Leave>", lambda e: self._draw(False))
        self.bind("<ButtonRelease-1>", self._click)

    def set_text(self, text: str) -> None:
        self._text = text
        import tkinter.font as tkfont
        self.configure(width=tkfont.Font(font=self._font_b).measure(text) + 60)
        self._draw(False)

    def _draw(self, hover: bool) -> None:
        self.delete("all")
        w, h = int(self["width"]), int(self["height"])
        _round_rect(self, 1, 1, w - 1, h - 1, 8, fill=self._hover if hover else self._bg,
                    outline=self._border or "", width=1 if self._border else 0)
        self.create_text(w // 2, h // 2, text=self._text, fill=self._fg,
                         font=self._font_b if hover else self._font_n)

    def _click(self, e) -> None:
        if 0 <= e.x <= int(self["width"]) and 0 <= e.y <= int(self["height"]):
            self._command()


class ControlButton(tk.Canvas):
    """운영 StartButton / StopButton 템플릿 (높이 29, 구역 폭에 맞춰 늘어남, 파란 삼각형 / 빨간 사각형)."""

    def __init__(self, master, text: str, kind: str, command: Callable[[], None]) -> None:
        super().__init__(master, width=140, height=29, highlightthickness=0, bg=master.cget("bg"), cursor="hand2")
        self._text = text
        self._kind = kind  # "start" | "stop"
        self._command = command
        self._hover = False
        self.bind("<Configure>", lambda e: self._draw())
        self.bind("<Enter>", lambda e: self._set_hover(True))
        self.bind("<Leave>", lambda e: self._set_hover(False))
        self.bind("<ButtonRelease-1>", self._click)
        self._draw()

    @property
    def text(self) -> str:
        return self._text

    def set_text(self, text: str) -> None:
        self._text = text
        self._draw()

    def _set_hover(self, hover: bool) -> None:
        self._hover = hover
        self._draw()

    def _draw(self) -> None:
        self.delete("all")
        w = self.winfo_width() if self.winfo_width() > 1 else int(self["width"])
        accent = T.START_ICON if self._kind == "start" else T.STOP_ICON
        bg = accent if self._hover else T.WHITE
        icon = T.WHITE if self._hover else accent
        fg = T.WHITE if self._hover else T.TEXT
        _round_rect(self, 1, 1, w - 1, 28, 8, fill=bg, outline=T.DARK_GRAY)
        f = T.font(9, bold=True)
        import tkinter.font as tkfont
        tw = tkfont.Font(font=f).measure(self._text)
        x0 = max((w - (16 + tw)) // 2, 4)
        if self._kind == "start":
            self.create_polygon(x0, 9, x0 + 6, 14, x0, 20, fill=icon, outline="")
        else:
            self.create_rectangle(x0, 10, x0 + 10, 20, fill=icon, outline="")
        self.create_text(x0 + 16, 15, text=self._text, anchor="w", fill=fg, font=f)

    def _click(self, e) -> None:
        if 0 <= e.x <= self.winfo_width() and 0 <= e.y <= 29:
            self._command()


class ImageButton(tk.Label):
    """그림 버튼 (원본 SettingsTemplate / AddTemplate)."""

    def __init__(self, master, image: tk.PhotoImage, tooltip: str, command: Callable[[], None]) -> None:
        super().__init__(master, image=image, bg=master.cget("bg"), cursor="hand2", bd=0)
        self.bind("<Button-1>", lambda e: command())
        Tooltip(self, tooltip)


class StatusDot(tk.Canvas):
    """통신 상태 원형 표시(10x10). 파랑=정상, 빨강=이상, 회색=중지."""

    def __init__(self, master, size: int = 10) -> None:
        super().__init__(master, width=size, height=size, highlightthickness=0, bg=master.cget("bg"))
        self._oval = self.create_oval(0, 0, size - 1, size - 1, fill=T.NET_STOP, outline="")
        self._color = T.NET_STOP

    def set_color(self, color: str) -> None:
        if color != self._color:
            self._color = color
            self.itemconfigure(self._oval, fill=color)

    @property
    def color(self) -> str:
        return self._color


class DottedSeparator(tk.Canvas):
    """대화상자의 점선 구분선 (원본 StrokeDashArray='1 2')."""

    def __init__(self, master, padx: int = 10) -> None:
        super().__init__(master, height=1, highlightthickness=0, bg=master.cget("bg"))
        self._padx = padx
        self.bind("<Configure>", self._draw)

    def _draw(self, e) -> None:
        self.delete("all")
        self.create_line(self._padx, 0, e.width - self._padx, 0, fill=T.SEPARATOR, dash=(1, 2))


# ------------------------------------------------------------------ 모터 목록
class MotorListView(tk.Frame):
    """라인별 모터 목록 (원본 ListView + GridView).

    열: [종류 그림 35px] [ID 60px] [상태: 상태그림 + 수동/참조 표시 + RPM/이상/중지/대기]
    - 머리글 클릭: 정렬 (종류 / ID / 상태)
    - 항목 더블클릭: 모터 수정/삭제 창
    """

    HEADER_H = 25
    ROW_H = 26
    COL0 = 32      # 운영 ListView_SizeChanged: 32 / 42 / 나머지
    COL1 = 42
    COLUMNS = ("", "ID", "상태")
    SORT_KEYS = {"": "MTYPE", "ID": "ID", "상태": "STATUS"}
    STATUS_ICON_W = 22
    # 속도 지정 방식 표시 (상태 그림과 숫자 사이): 수동 / 참조(자동으로 참조값을 따름)
    SRC_BADGES = {"manual": ("수동", T.SRC_MANUAL_FG, T.SRC_MANUAL_BG), "ref": ("참조", T.SRC_REF_FG, T.SRC_REF_BG)}
    BADGE_X = 6            # 상태 그림 오른쪽 여백
    BADGE_W, BADGE_H = 28, 15
    # 속도 변경을 보내면 1초 동안 팬 그림을 붉은 그림으로, 그 행 바탕을 붉게 (깜빡이지 않고 1초 동안 유지)
    SENT_SEC = 1.0
    _clock = staticmethod(time.monotonic)

    @staticmethod
    def status_icon(m: Motor):
        """상태 그림 이름과 툴팁 (운영 StatusTemplate). 대기는 그림 없음.
        이상표시(접속 실패·무응답·제어 문제)는 다시 응답이 오면 바로 정상 그림으로 돌아감."""
        why = f"\n{m.err_note}" if m.err_note else ""
        speed = MotorListView.speed_text(m)
        if m.status == MotorStatus.ERR:
            return "bad_new", "이상 (오래 응답 없음)" + why
        if m.status == MotorStatus.NOTUSE:
            return "nouse_new", "중지(사용안함)"
        if m.had_err:
            return "bad_haderr", ("응답 대기 중 이상" if m.status == MotorStatus.READY else "통신 이상") + why
        if m.speed_note:
            return "bad_haderr", f"속도 미달\n{m.speed_note}" + speed
        if m.status == MotorStatus.OK:
            return "good", "정상" + speed
        return None, ""

    @staticmethod
    def speed_text(m: Motor) -> str:
        """설정·지정·실제 속도 한 줄 (신규모터·기존모터제어, 설정 속도를 알 때만). 예: '설정 770 · 지정 785 · 실제 768 RPM'"""
        if m.mtype == MotorType.OLD_MOTOR or m.target_rpm is None or m.status != MotorStatus.OK:
            return ""
        cmd = "-" if m.cmd_rpm is None else str(m.cmd_rpm)
        ok = " (오차 1% 이내)" if m.target_rpm > 0 and speed_within(m.rpm, m.target_rpm) else ""
        return f"\n설정 {m.target_rpm} · 지정 {cmd} · 실제 {m.rpm} RPM{ok}"

    @staticmethod
    def type_icon(m: Motor):
        """종류 그림 이름과 툴팁 (운영 MotorTypeTemplate)."""
        if m.mtype == MotorType.OLD_MOTOR:
            return "oldmotor", "기존모터"
        if m.mtype == MotorType.OLD_CTRL:
            if m.ctrl_mode == CTRL_MODE_REF:
                src = f" {m.ref_desc}" if m.ref_desc else ""
                now = ("값 없음" if m.ref_rpm is None
                       else f"{m.ref_rpm} RPM = {pct_text(rpm_to_pct(m.ref_rpm))}")
                t = m.target_rpm if m.target_rpm is not None else m.ctrl_target
                speed = (f"구역 참조{src} (지금 {now})"
                         + ("" if t is None else f" → 설정 {pct_text(rpm_to_pct(t))} = {t} RPM"))
            else:
                speed = f"수동 지정 {m.set_rpm} RPM ({pct_text(rpm_to_pct(m.set_rpm))})"
            tip = f"기존모터제어 {m.ip}\n운전속도: {speed}"
            if m.target_rpm is not None and m.cmd_rpm is not None and m.cmd_rpm != m.target_rpm:
                tip += f"\n모터 지정값 {m.cmd_rpm} RPM (설정과 다름)"
            if m.max_rpm > 0:
                tip += f"\n모터 최고속도(h8) {m.max_rpm} RPM"
                if m.target_rpm:
                    tip += f": 설정 {m.target_rpm} RPM = 최고속도의 {pct_text(m.target_rpm * 100 / m.max_rpm)}"
            return "oldmotor_ctrl", tip + (f"\n{m.ctrl_note}" if m.ctrl_note else "")
        if m.use_limit:
            return "newmotor_limit", "신규모터(제한모드)"
        return "newmotor", "신규모터"

    def __init__(self, master, images: T.Images,
                 on_header_click: Callable[[str], None],
                 on_double_click: Callable[[str], None]) -> None:
        super().__init__(master, bg=T.LIST_BORDER, padx=1, pady=1)
        self.images = images
        self._on_header_click = on_header_click
        self._on_double_click = on_double_click
        self._rows: List[Motor] = []
        self._selected_key: Optional[str] = None
        self._hover_idx: Optional[int] = None
        self._hover_hdr: Optional[int] = None
        self._width = 200
        self._sent_start: Dict[str, float] = {}    # 모터키 -> '속도 보냄' 붉은 표시 시작 시각
        self._seen_sent: Dict[str, float] = {}     # 모터키 -> 마지막으로 본 '속도 보낸 시각'
        self._sent_job = None
        self.drawn_icons: List[str] = []           # 마지막으로 그린 종류 그림 이름 (시험용)
        self.drawn_srcs: List[str] = []            # 마지막으로 그린 수동/참조 표시 글자 ('' = 없음, 시험용)

        inner = tk.Frame(self, bg=T.WHITE)
        inner.pack(fill="both", expand=True)
        self.header = tk.Canvas(inner, height=self.HEADER_H, bg=T.LIST_HEADER_BG, highlightthickness=0)
        self.header.pack(side="top", fill="x")
        body = tk.Frame(inner, bg=T.WHITE)
        body.pack(side="top", fill="both", expand=True)
        self.canvas = tk.Canvas(body, bg=T.WHITE, highlightthickness=0)
        self.vbar = tk.Scrollbar(body, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self._on_scroll)
        self.canvas.pack(side="left", fill="both", expand=True)
        self._vbar_visible = False

        self.canvas.bind("<Configure>", self._on_resize)
        self.canvas.bind("<Button-1>", self._on_click)
        self.canvas.bind("<Double-Button-1>", self._on_dclick)
        self.canvas.bind("<Motion>", self._on_motion)
        self.canvas.bind("<Enter>", self._on_enter)
        self.canvas.bind("<Leave>", self._on_leave)
        self.canvas.bind("<Button-4>", lambda e: self.canvas.yview_scroll(-1, "units"))
        self.canvas.bind("<Button-5>", lambda e: self.canvas.yview_scroll(1, "units"))
        self.header.bind("<Motion>", self._on_hdr_motion)
        self.header.bind("<Leave>", lambda e: self._set_hdr_hover(None))
        self.header.bind("<Button-1>", self._on_hdr_click)
        Tooltip(self.canvas, self._tooltip_text)
        self.bind("<Destroy>", self._on_destroy, add="+")
        self._draw_header()

    # --- 데이터 ---
    def set_rows(self, rows: List[Motor]) -> None:
        now = self._clock()
        for m in rows:
            if m.sent_at is not None and self._seen_sent.get(m.key) != m.sent_at:
                self._seen_sent[m.key] = m.sent_at
                if time.monotonic() - m.sent_at < 2.0:        # 방금 속도 변경을 보낸 모터만 붉게
                    self._sent_start[m.key] = now
        self._rows = rows
        self._redraw()
        self._schedule_sent()

    # --- 속도 보냄 표시 (1초 동안 붉은 팬 그림 + 붉은 행) ---
    def sent_on(self, key: str, now: Optional[float] = None) -> bool:
        """속도 변경을 보낸 뒤 1초(SENT_SEC) 안인지 (그동안 계속 붉게 표시)."""
        start = self._sent_start.get(key)
        if start is None:
            return False
        el = (self._clock() if now is None else now) - start
        return 0 <= el < self.SENT_SEC

    def _schedule_sent(self) -> None:
        """붉은 표시가 끝나는 때에 맞춰 다시 그리기 예약."""
        now = self._clock()
        for k, st in list(self._sent_start.items()):
            if now - st >= self.SENT_SEC:
                del self._sent_start[k]
        if self._sent_job is not None:
            try:
                self.after_cancel(self._sent_job)
            except tk.TclError:
                pass
            self._sent_job = None
        if not self._sent_start:
            return
        nxt = min(st + self.SENT_SEC for st in self._sent_start.values())
        self._sent_job = self.after(max(10, int((nxt - now) * 1000) + 5), self._sent_tick)

    def _sent_tick(self) -> None:
        self._sent_job = None
        if not self.winfo_exists():
            return
        self._redraw()
        self._schedule_sent()

    def _on_destroy(self, e) -> None:
        if e.widget is self and self._sent_job is not None:
            try:
                self.after_cancel(self._sent_job)
            except tk.TclError:
                pass
            self._sent_job = None

    @property
    def rows(self) -> List[Motor]:
        return self._rows

    @property
    def selected_key(self) -> Optional[str]:
        return self._selected_key

    # --- 그리기 ---
    def _col_x(self):
        w = self._width
        return 0, self.COL0, self.COL0 + self.COL1, w

    def _on_resize(self, e) -> None:
        # 원본 ListView_SizeChanged: 0열 35, 1열 60, 2열 = 나머지
        self._width = e.width
        self._draw_header()
        self._redraw()

    def _on_scroll(self, first, last) -> None:
        self.vbar.set(first, last)
        need = not (float(first) <= 0.0 and float(last) >= 1.0)
        if need != self._vbar_visible:
            self._vbar_visible = need
            if need:
                self.vbar.pack(side="right", fill="y", before=self.canvas)
            else:
                self.vbar.pack_forget()

    def _draw_header(self) -> None:
        h = self.header
        h.delete("all")
        x = self._col_x()
        w = max(self._width, 1)
        for i in range(3):
            x1, x2 = x[i], x[i + 1]
            if self._hover_hdr == i:
                h.create_rectangle(x1, 0, x2, self.HEADER_H, fill=T.LIST_HEADER_HOVER, outline="")
            h.create_line(x2 - 1, 2, x2 - 1, self.HEADER_H - 2, fill=T.LIST_HEADER_LINE)
            h.create_text((x1 + x2) // 2, self.HEADER_H // 2, text=self.COLUMNS[i], font=T.font(9))
        h.create_line(0, self.HEADER_H - 1, w, self.HEADER_H - 1, fill=T.LIST_HEADER_LINE)

    def _redraw(self) -> None:
        c = self.canvas
        c.delete("all")
        x = self._col_x()
        w = self._width
        now = self._clock()
        self.drawn_icons = []
        self.drawn_srcs = []
        for i, m in enumerate(self._rows):
            y = i * self.ROW_H
            sent = self.sent_on(m.key, now)          # 속도 변경을 보냄: 1초 동안 붉은 팬 그림 + 붉은 행
            fill = T.LIST_SENT_BG if sent else self.row_fill(m, m.key == self._selected_key, i == self._hover_idx)
            if fill:
                c.create_rectangle(0, y, w, y + self.ROW_H - 1, fill=fill, outline="", tags=("rowbg", f"row{i}"))
            cy = y + self.ROW_H // 2
            tname, _ = self.type_icon(m)
            if sent:
                tname += "_red"
            self.drawn_icons.append(tname)
            c.create_image((x[0] + x[1]) // 2, cy, image=self.images.get(tname, 20))
            c.create_text((x[1] + x[2]) // 2, cy, text=str(m.id), font=T.font(9))
            self._draw_status(c, m, x[2], x[3], cy)
        c.configure(scrollregion=(0, 0, w, max(len(self._rows) * self.ROW_H, 1)))

    @staticmethod
    def row_fill(m: Motor, selected: bool, hover: bool) -> Optional[str]:
        """행 바탕색. 문제 상태(이상표시·이상·속도 미달)인 행은 늘 붉은 음영 (선택·마우스오버는 더 진하게)."""
        if m.has_problem:
            return T.LIST_PROBLEM_SELECTED if selected else (T.LIST_PROBLEM_HOVER if hover else T.LIST_PROBLEM_BG)
        if selected:
            return T.LIST_SELECTED
        if hover:
            return T.LIST_HOVER
        return None

    @classmethod
    def src_badge(cls, m: Motor) -> Optional[tuple]:
        """수동/참조 표시 (글자, 글자색, 바탕색). 기존모터·사용안함은 없음."""
        if m.status == MotorStatus.NOTUSE:
            return None
        return cls.SRC_BADGES.get(m.speed_src)

    @staticmethod
    def src_tip(m: Motor) -> str:
        """수동/참조 표시 툴팁."""
        now = "" if m.target_rpm is None else f"\n지금 설정 {m.target_rpm} RPM ({pct_text(rpm_to_pct(m.target_rpm))})"
        if m.mtype == MotorType.OLD_CTRL:
            if m.speed_src == "ref":
                return "참조(자동): 구역 참조 - 이 구역 신규모터가 따르는 값을 같은 비율로 따라 운전" + now
            return f"수동: 수동 지정 {m.set_rpm} RPM ({pct_text(rpm_to_pct(m.set_rpm))})으로 운전"
        if m.speed_src == "ref":
            return "참조(자동): 구역설정의 참조 기존모터 회전수를 따라 운전" + now
        return "수동: 구역설정의 수동 RPM 으로 운전" + now

    def _draw_status(self, c: tk.Canvas, m: Motor, x1: int, x2: int, cy: int) -> None:
        """운영 StatusTemplate: 정상/이상은 RPM 숫자(이상은 빨강 굵게), 중지/대기는 연회색 굵은 글씨.
        상태 그림과 숫자 사이에 속도 지정 방식(수동 / 참조) 표시."""
        icon, _ = self.status_icon(m)
        if icon:
            c.create_image(x1 + self.STATUS_ICON_W // 2 + 1, cy, image=self.images.get(icon, 20))
        badge = self.src_badge(m)
        self.drawn_srcs.append(badge[0] if badge else "")
        if badge:
            text, fg, bg = badge
            bx = x1 + self.STATUS_ICON_W + self.BADGE_X
            hh = self.BADGE_H // 2
            _round_rect(c, bx, cy - hh, bx + self.BADGE_W, cy + hh, 4, fill=bg, outline=fg, tags=("srcbadge",))
            c.create_text(bx + self.BADGE_W // 2, cy, text=text, fill=fg, font=T.font(8, bold=True),
                          tags=("srcbadge",))
        if m.status == MotorStatus.OK:
            text, fg, bold = str(m.rpm), T.TEXT, False
        elif m.status == MotorStatus.ERR:
            text, fg, bold = str(m.rpm), T.STATUS_ERR_FG, True
        elif m.status == MotorStatus.NOTUSE:
            text, fg, bold = "중지", T.STATUS_DIM_FG, True
        else:
            text, fg, bold = "대기", T.STATUS_DIM_FG, True
        c.create_text(x2 - 8, cy, text=text, anchor="e", fill=fg, font=T.font(9, bold=bold))

    # --- 마우스 ---
    def _row_at(self, y: int) -> Optional[int]:
        idx = int(self.canvas.canvasy(y) // self.ROW_H)
        return idx if 0 <= idx < len(self._rows) else None

    def _on_click(self, e) -> None:
        idx = self._row_at(e.y)
        if idx is not None:
            self._selected_key = self._rows[idx].key
            self._redraw()

    def _on_dclick(self, e) -> None:
        idx = self._row_at(e.y)
        if idx is not None:
            self._selected_key = self._rows[idx].key
            self._redraw()
            self._on_double_click(self._rows[idx].key)

    def _on_motion(self, e) -> None:
        idx = self._row_at(e.y)
        if idx != self._hover_idx:
            self._hover_idx = idx
            self._redraw()

    def _on_enter(self, e) -> None:
        # Windows 는 휠 이벤트가 포커스 위젯으로 가므로, 마우스가 올라간 목록이 받도록 전역 연결
        self.canvas.bind_all("<MouseWheel>", self._on_wheel)

    def _on_leave(self, e) -> None:
        self.canvas.unbind_all("<MouseWheel>")
        if self._hover_idx is not None:
            self._hover_idx = None
            self._redraw()

    def _on_wheel(self, e) -> None:
        self.canvas.yview_scroll(int(-e.delta / 120) or (-1 if e.delta > 0 else 1), "units")

    def _hdr_col(self, x: int) -> Optional[int]:
        xs = self._col_x()
        for i in range(3):
            if xs[i] <= x < xs[i + 1]:
                return i
        return None

    def _set_hdr_hover(self, col: Optional[int]) -> None:
        if col != self._hover_hdr:
            self._hover_hdr = col
            self._draw_header()

    def _on_hdr_motion(self, e) -> None:
        self._set_hdr_hover(self._hdr_col(e.x))

    def _on_hdr_click(self, e) -> None:
        col = self._hdr_col(e.x)
        if col is not None:
            self._on_header_click(self.SORT_KEYS[self.COLUMNS[col]])

    def _tooltip_text(self, x: int, y: int) -> str:
        idx = self._row_at(y)
        if idx is not None:
            m = self._rows[idx]
            xs = self._col_x()
            if xs[0] <= x < xs[1]:
                return self.type_icon(m)[1]
            if xs[2] <= x < xs[2] + self.STATUS_ICON_W + 2:
                tip = self.status_icon(m)[1]
                if tip:
                    return tip
            bx = xs[2] + self.STATUS_ICON_W + self.BADGE_X
            if bx <= x < bx + self.BADGE_W and self.src_badge(m):
                return self.src_tip(m)
        return "항목 더블클릭시 수정/삭제"
