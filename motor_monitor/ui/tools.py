"""메인 화면 도구: 에러 로그 창, 다른 프로그램 설정 가져오기 창.

- ErrorLogWindow (SCR-ERRLOG): 최근 에러/복구 기록 (최신순). 창이 열려 있는 동안 2초마다 새 기록 반영
- ConfigImportDialog (DLG-IMPORT-CFG): 운영 프로그램 설정(Config.dll / uconfig.dll / Config.ini) 미리보기 후 적용
- MiniImportDialog (DLG-IMPORT-MINI): 미니 제어기 항목(mini_control.json)을 한 구역의 기존모터제어로 추가
모든 창은 메인 창 안 가운데에 뜸 (dialogs.place_inside).
"""
from __future__ import annotations

import tkinter as tk
from tkinter import ttk
from typing import TYPE_CHECKING, Callable, List, Optional

from ..config import MiniItem
from ..constants import HOUSE_LINES, MAX_RPM, pct_text
from ..errorlog import ERROR, RECOVER
from ..helpers import get_house_name
from . import msgbox
from . import theme as T
from .dialogs import BaseDialog, place_inside
from .widgets import RoundButton

if TYPE_CHECKING:
    from ..plant import ConfigImport, Plant

REFRESH_MS = 2000


# ======================================================================== 에러 로그 창
class ErrorLogWindow(tk.Toplevel):
    """SCR-ERRLOG 에러 로그 창 (모달 아님)."""

    COLS = (("time", "시각", 140), ("level", "구분", 50), ("zone", "구역", 55),
            ("target", "대상", 230), ("msg", "내용", 420))
    SIZE = (920, 480)

    def __init__(self, owner: tk.Misc, plant: "Plant", on_seen: Optional[Callable[[], None]] = None) -> None:
        super().__init__(owner)
        self.withdraw()
        self.plant = plant
        self.on_seen = on_seen
        self.title("에러 로그")
        self.minsize(640, 300)
        self.configure(bg=T.WHITE)
        self._version = -1

        head = tk.Frame(self, bg=T.LINE_DIALOG_HEADER)
        head.pack(side="top", fill="x")
        tk.Label(head, text="에러 로그", bg=T.LINE_DIALOG_HEADER, fg=T.WHITE,
                 font=T.font(12, bold=True)).pack(side="left", padx=10, pady=6)
        self.lbl_info = tk.Label(head, text="", bg=T.LINE_DIALOG_HEADER, fg="#d0d7de", font=T.font(9))
        self.lbl_info.pack(side="left", padx=6)

        bar = tk.Frame(self, bg=T.WHITE)
        bar.pack(side="bottom", fill="x", padx=8, pady=6)
        self.show_recover = tk.BooleanVar(value=True)
        tk.Checkbutton(bar, text="복구 기록도 보기", variable=self.show_recover, command=self.refresh_force,
                       bg=T.WHITE, activebackground=T.WHITE, font=T.font(9),
                       highlightthickness=0).pack(side="left")
        self.BtnClose = RoundButton(bar, "닫기", self.destroy, style="normal")
        self.BtnClose.pack(side="right")
        self.BtnClear = RoundButton(bar, "목록 지우기", self.on_clear, style="normal")
        self.BtnClear.pack(side="right", padx=5)
        self.BtnRefresh = RoundButton(bar, "새로고침", self.refresh_force, style="submit")
        self.BtnRefresh.pack(side="right")

        body = tk.Frame(self, bg=T.WHITE)
        body.pack(side="top", fill="both", expand=True, padx=8, pady=(6, 0))
        self.tree = ttk.Treeview(body, columns=[c for c, _, _ in self.COLS], show="headings")
        for c, t, w in self.COLS:
            self.tree.heading(c, text=t)
            self.tree.column(c, width=w, stretch=(c == "msg"), anchor="w")
        sb = ttk.Scrollbar(body, command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.tree.tag_configure("err", foreground="#c0392b")
        self.tree.tag_configure("ok", foreground="#1e8449")

        self.refresh_force()
        self._after = self.after(REFRESH_MS, self._tick)
        self.bind("<Destroy>", self._on_destroy, add="+")
        place_inside(self, owner, *self.SIZE)          # 메인 창 안 가운데 (모니터 기준 아님)
        self.deiconify()
        self.lift()

    def _on_destroy(self, e) -> None:
        if e.widget is self and self._after:
            try:
                self.after_cancel(self._after)
            except tk.TclError:
                pass
            self._after = None

    def _tick(self) -> None:
        self._after = None
        try:
            if self.plant.errors.version != self._version:
                self.refresh_force()
        finally:
            if self.winfo_exists():
                self._after = self.after(REFRESH_MS, self._tick)

    def refresh_force(self) -> None:
        el = self.plant.errors
        self._version = el.version
        rows = [e for e in el.recent() if self.show_recover.get() or e.level != RECOVER]
        self.tree.delete(*self.tree.get_children())
        for e in rows:
            tag = "err" if e.level == ERROR else "ok"
            self.tree.insert("", "end", values=(e.time.strftime("%Y-%m-%d %H:%M:%S"), e.level, e.zone, e.target,
                                                e.message), tags=(tag,))
        n_err = sum(1 for e in rows if e.level == ERROR)
        where = el.path if el.path else "파일 기록 안 함"
        self.lbl_info.configure(text=f"최근 {len(rows)}건 (에러 {n_err}건, 최신순)  ·  파일: {where}")
        el.mark_seen()
        if self.on_seen:
            self.on_seen()

    @property
    def rows(self) -> List[tuple]:
        return [tuple(self.tree.item(i, "values")) for i in self.tree.get_children()]

    def on_clear(self) -> None:
        if msgbox.ask_yes_no("목록 지우기", "화면의 에러 로그 목록을 지우시겠습니까?\n(파일 기록은 그대로 남습니다)",
                             parent=self):
            self.plant.errors.clear()
            self.refresh_force()


# ======================================================================== 설정 가져오기
class ConfigImportDialog(BaseDialog):
    """DLG-IMPORT-CFG 운영 프로그램 설정 가져오기: 내용 미리보기 후 [가져오기]."""

    def __init__(self, owner: tk.Misc, plant: "Plant", imp: "ConfigImport",
                 on_done: Optional[Callable[[], None]] = None) -> None:
        super().__init__(owner, "운영 프로그램 설정 가져오기", T.LINE_DIALOG_HEADER, 820, 470)
        self.plant = plant
        self.imp = imp
        self.on_done = on_done
        self.result_backup: Optional[str] = None
        tk.Label(self.body, text=f"파일: {imp.path}", bg=T.WHITE, font=T.font(9, bold=True), anchor="w",
                 wraplength=790, justify="left").pack(side="top", fill="x", padx=12, pady=(10, 4))
        tk.Label(self.body, text="가져오면 아래 내용으로 8개 구역의 통신 설정과 모터 목록을 바꿉니다. "
                                 "지금 설정 파일은 먼저 Config.dll.bak_날짜_시각 으로 보관합니다.",
                 bg=T.WHITE, fg=T.TEXT_DARK, font=T.font(9), anchor="w", wraplength=790,
                 justify="left").pack(side="top", fill="x", padx=12)
        box = tk.Frame(self.body, bg=T.WHITE)
        box.pack(side="top", fill="both", expand=True, padx=12, pady=6)
        self.txt = tk.Text(box, height=12, font=T.font(9), wrap="none", bg="#f8f9fa", relief="solid", bd=1)
        self.txt.insert("1.0", "\n".join(imp.summary_lines()))
        self.txt.configure(state="disabled")
        self.txt.pack(fill="both", expand=True)
        self.keep_ctrl = tk.BooleanVar(value=True)
        if not imp.has_ctrl:
            tk.Checkbutton(self.body, text="가져온 파일에 기존모터제어가 없으므로, 지금 등록된 기존모터제어는 그대로 유지",
                           variable=self.keep_ctrl, bg=T.WHITE, activebackground=T.WHITE, font=T.font(9),
                           highlightthickness=0).pack(side="top", anchor="w", padx=10)
        btns = tk.Frame(self.footer, bg=T.WHITE)
        btns.pack(side="right", padx=5)
        self.BtnImport = RoundButton(btns, "가져오기", self.on_import, style="submit")
        self.BtnImport.pack(side="left", padx=(0, 5))
        self.BtnCancel = RoundButton(btns, "취소", self.on_cancel, style="normal")
        self.BtnCancel.pack(side="left")

    def on_default(self) -> None:
        self.on_import()

    def on_import(self) -> None:
        if self.plant.any_running:
            return self.error("실행 중인 구역이 있습니다. 모든 구역을 종료한 뒤 가져오세요.", "가져오기")
        if not self.ask("설정을 가져오시겠습니까? 지금 설정 파일은 백업 후 바뀝니다.", "가져오기"):
            return
        try:
            self.result_backup = self.plant.apply_config_import(self.imp, keep_ctrl=self.keep_ctrl.get())
        except (OSError, RuntimeError) as e:
            return self.error(f"가져오지 못했습니다.\n{e}", "가져오기")
        if self.on_done:
            self.on_done()
        self.info("설정을 가져왔습니다." + (f"\n이전 설정 백업: {self.result_backup}" if self.result_backup else ""))
        self.destroy()

    def info(self, msg: str) -> None:
        msgbox.show_info("가져오기", msg, parent=self)


class MiniImportDialog(BaseDialog):
    """DLG-IMPORT-MINI 미니 제어기 항목을 선택한 구역의 기존모터제어로 추가."""

    def __init__(self, owner: tk.Misc, plant: "Plant", items: List[MiniItem], path: str,
                 on_done: Optional[Callable[[], None]] = None) -> None:
        super().__init__(owner, "미니 제어기 목록 가져오기", T.MOTOR_DIALOG_HEADER, 640, 430)
        self.plant = plant
        self.items = items
        self.on_done = on_done
        self.added: List[str] = []
        self.skipped: List[str] = []
        tk.Label(self.body, text=f"파일: {path}", bg=T.WHITE, font=T.font(9, bold=True), anchor="w",
                 wraplength=610, justify="left").pack(side="top", fill="x", padx=12, pady=(10, 4))
        tk.Label(self.body, text=f"항목 {len(items)}개를 기존모터제어로 추가합니다. 속도 %는 수동 지정 RPM"
                                 f"(100% = {MAX_RPM} RPM), 참조팬 항목은 '구역 참조'(그 구역 신규모터가 따르는 값을 "
                                 f"{MAX_RPM} RPM 대비 %로 같은 비율)로 바꿉니다. 미니 제어기에 지정한 참조팬은 쓰지 않습니다.",
                 bg=T.WHITE, fg=T.TEXT_DARK, font=T.font(9), anchor="w", wraplength=610,
                 justify="left").pack(side="top", fill="x", padx=12)
        box = tk.Frame(self.body, bg=T.WHITE)
        box.pack(side="top", fill="both", expand=True, padx=12, pady=6)
        self.tree = ttk.Treeview(box, columns=("ip", "id", "spec"), show="headings", height=7)
        for c, t, w in (("ip", "IP", 140), ("id", "ID", 50), ("spec", "운전속도", 360)):
            self.tree.heading(c, text=t)
            self.tree.column(c, width=w, anchor="w")
        for it in items:
            if it.is_ref:
                was = f" (미니 제어기 참조팬 {it.ref_ip} ID {it.ref_uid} 대신)" if it.ref_ip else ""
                spec = f"구역 참조{was}"
            else:
                spec = f"수동 지정 {it.rpm} RPM ({pct_text(it.pct)})"
            self.tree.insert("", "end", values=(it.ip, it.uid, spec))
        self.tree.pack(fill="both", expand=True)
        row = tk.Frame(self.body, bg=T.WHITE)
        row.pack(side="top", fill="x", padx=12, pady=(0, 6))
        tk.Label(row, text="추가할 구역", bg=T.WHITE, font=T.font(10, bold=True)).pack(side="left")
        self.zone_names = [get_house_name(s) for s in HOUSE_LINES]
        self.cb_zone = ttk.Combobox(row, state="readonly", width=8, values=self.zone_names, font=T.font(10))
        guess = plant.zone_for_ip(items[0].ip) if items else None
        self.cb_zone.current(HOUSE_LINES.index(guess) if guess else 0)
        self.cb_zone.pack(side="left", padx=8)
        tk.Label(row, text="(기존모터 통신장치 IP 가 같은 구역을 먼저 선택해 둠)", bg=T.WHITE, fg="#666666",
                 font=T.font(8)).pack(side="left")
        btns = tk.Frame(self.footer, bg=T.WHITE)
        btns.pack(side="right", padx=5)
        self.BtnImport = RoundButton(btns, "추가", self.on_import, style="submit")
        self.BtnImport.pack(side="left", padx=(0, 5))
        self.BtnCancel = RoundButton(btns, "취소", self.on_cancel, style="normal")
        self.BtnCancel.pack(side="left")

    @property
    def zone_code(self) -> str:
        return HOUSE_LINES[max(self.cb_zone.current(), 0)]

    def on_default(self) -> None:
        self.on_import()

    def on_import(self) -> None:
        name = get_house_name(self.zone_code)
        if not self.ask(f"{name}에 기존모터제어 {len(self.items)}개를 추가하시겠습니까?", "가져오기"):
            return
        try:
            self.added, self.skipped = self.plant.import_mini_items(self.items, self.zone_code)
        except OSError as e:
            return self.error(f"저장하지 못했습니다.\n{e}", "가져오기")
        if self.on_done:
            self.on_done()
        msg = f"{name}에 {len(self.added)}개를 추가했습니다."
        if self.skipped:
            msg += "\n\n건너뜀:\n" + "\n".join(self.skipped)
        msgbox.show_info("가져오기", msg, parent=self)
        self.destroy()
