"""메인 화면 (SCR-MAIN). 원본: MainWindow.xaml / MainWindow.xaml.vb

구성 (v1.1.0: 원본의 온실1/온실2 탭을 없애고 한 화면에 1~8구역 표시)
  상단: 1~4구역 / 5~8구역 제어 시작·종료 버튼
  본문: 4열x2행 구역 패널(윗줄 1~4구역=온실1, 아랫줄 5~8구역=온실2), 각 구역은 짝/홀 라인 패널 2개
  라인 패널: 머리 영역(통신상태 점, IP/PORT, RPM 소스, 설정/추가 버튼) + 모터 목록
"""
from __future__ import annotations

import logging
import tkinter as tk
from tkinter import messagebox
from typing import Dict

from .. import constants as C
from ..helpers import get_house_short_name, zone_number
from ..plant import Plant
from . import theme as T
from .dialogs import LineSettingsDialog, MotorDialog
from .widgets import ControlButton, ImageButton, MotorListView, StatusDot, Tooltip

log = logging.getLogger(__name__)

LIST_REFRESH_MS = 300   # 모터 목록 변경 확인 주기 (원본은 데이터 바인딩으로 즉시 반영)


class LinePanel(tk.Frame):
    """라인 1개(예: 1구역 짝) 영역. 스펙시트 LINE-{코드}"""

    def __init__(self, master, app: "MainWindow", code: str) -> None:
        super().__init__(master, bg=T.WHITE)
        self.app = app
        self.code = code
        img = app.images
        side_even = code[4] == "E"

        head = tk.Frame(self, bg=T.WHITE, height=60)
        head.pack(side="top", fill="x")
        head.grid_propagate(False)
        head.columnconfigure(1, weight=1)
        for r in range(3):
            head.rowconfigure(r, weight=1, uniform="h")

        # 짝/홀 색상 표시
        side = tk.Frame(head, bg=T.SIDE_EVEN_BG if side_even else T.SIDE_ODD_BG, width=21)
        side.grid(row=0, column=0, rowspan=3, sticky="ns", padx=(5, 0), pady=(0, 2))
        side.grid_propagate(False)
        tk.Label(side, text="짝" if side_even else "홀", fg=T.WHITE,
                 bg=side.cget("bg"), font=T.font(9)).place(relx=0.5, rely=0.5, anchor="center")

        def info_row(r: int, icon, tip: str):
            f = tk.Frame(head, bg=T.WHITE)
            f.grid(row=r, column=1, sticky="ew", padx=(5, 0))
            if icon is not None:
                lb = tk.Label(f, image=icon, bg=T.WHITE, bd=0)
                Tooltip(lb, tip)
            else:
                lb = tk.Label(f, text="┖", bg=T.WHITE, width=2, anchor="w", bd=0, font=T.font(9))
            lb.pack(side="left")
            dot = StatusDot(f)
            dot.pack(side="left", padx=(5, 0))
            lbl = tk.Label(f, text="-", bg=T.WHITE, bd=0, font=T.font(9), anchor="w")
            lbl.pack(side="left", padx=(5, 0), fill="x", expand=True)
            # 화면이 좁아 글자가 잘릴 때를 위해 전체 문구를 툴팁으로 표시
            Tooltip(lbl, lambda x, y, lbl=lbl: lbl.cget("text"))
            return dot, lbl

        self.rec_SOURCE_NETSTATUS, self.lbl_old_ip = info_row(0, img.get("oldmotor", 15), "기존모터")
        self.rec_TARGET_NETSTATUS, self.lbl_new_ip = info_row(1, img.get("newmotor", 15), "신규모터")
        self.rec_SID_NETSTATUS, self.lbl_source = info_row(2, None, "")

        btns = tk.Frame(head, bg=T.WHITE, width=30)
        btns.grid(row=0, column=2, rowspan=3, sticky="ns")
        self.Btn_Settings = ImageButton(btns, img.get("settings", 22), "설정", self.open_settings)
        self.Btn_Settings.pack(side="top", padx=4, pady=(2, 0))
        self.Btn_Add = ImageButton(btns, img.get("add", 22), "추가", self.open_add)
        self.Btn_Add.pack(side="bottom", padx=4, pady=(0, 2))

        self.LV = MotorListView(self, img, on_header_click=self.on_sort, on_double_click=self.open_modify)
        self.LV.pack(side="top", fill="both", expand=True, padx=5)
        self._version = -1

    # --- 동작 ---
    def open_settings(self) -> None:
        """설정(톱니) 버튼 -> 라인설정 창 (원본 Btn_Setting_Click)"""
        LineSettingsDialog(self.app.root, self.app.plant, self.code,
                           on_saved=self.app.update_header_box).show_modal()

    def open_add(self) -> None:
        """추가(+) 버튼 -> 모터 추가 창 (원본 Btn_Add_Click)"""
        MotorDialog(self.app.root, self.app.plant, self.app.images, self.code, "ADD",
                    on_changed=self.refresh).show_modal()

    def open_modify(self, dkey: str) -> None:
        """목록 더블클릭 -> 모터 수정/삭제 창 (원본 Lv_DBClick)"""
        MotorDialog(self.app.root, self.app.plant, self.app.images, self.code, "MOD", mod_dkey=dkey,
                    on_changed=self.refresh).show_modal()

    def on_sort(self, skey: str) -> None:
        """목록 머리글 클릭 -> 정렬 (원본 lvHeader_CLICK)"""
        self.app.plant.houses[self.code].sort(skey)
        self.refresh()

    def refresh(self, force: bool = True) -> None:
        h = self.app.plant.houses[self.code]
        if force or h.version != self._version:
            self._version = h.version
            self.LV.set_rows(h.get_list())


class ZonePanel(tk.Frame):
    """구역 1개(GroupBox): 제목줄 + 짝/홀 라인 패널. 스펙시트 ZONE-{1~8}"""

    def __init__(self, master, app: "MainWindow", house: str, zone: int) -> None:
        """house: 'H1'/'H2', zone: 온실 내 구역번호 1~4 (화면 표시는 1~8구역)"""
        super().__init__(master, bg=T.GROUP_BORDER, padx=1, pady=1)
        inner = tk.Frame(self, bg=T.WHITE, padx=3, pady=3)
        inner.pack(fill="both", expand=True)
        inner.columnconfigure(0, weight=1, uniform="c")
        inner.columnconfigure(1, weight=1, uniform="c")
        inner.rowconfigure(1, weight=1)
        title = tk.Frame(inner, bg=T.ZONE_TITLE_BG)
        title.grid(row=0, column=0, columnspan=2, sticky="ew", pady=3)
        tk.Label(title, text=f"{zone_number(f'{house}R{zone}E')}구역", bg=T.ZONE_TITLE_BG, fg=T.WHITE,
                 font=T.font(9, bold=True)).pack(pady=3)
        self.lines: Dict[str, LinePanel] = {}
        for col, side in enumerate(("E", "O")):
            code = f"{house}R{zone}{side}"
            lp = LinePanel(inner, app, code)
            lp.grid(row=1, column=col, sticky="nsew")
            self.lines[code] = lp


class MainWindow:
    """SCR-MAIN 메인 화면"""

    def __init__(self, root: tk.Tk, plant: Plant) -> None:
        self.root = root
        self.plant = plant
        T.init_fonts(root)
        self.images = T.Images(root)

        root.title(C.APP_TITLE)
        root.minsize(1280, 700)
        root.geometry("1600x900")
        if T.IS_WINDOWS:
            try:
                root.state("zoomed")     # 8개 구역을 넓게 보도록 최대화로 시작
            except tk.TclError:
                pass
        root.configure(bg=T.WHITE)
        T.set_window_icon(root)
        root.protocol("WM_DELETE_WINDOW", self.on_close)

        # --- 상단 제어 버튼 줄 ---
        top = tk.Frame(root, bg=T.WHITE, height=34)
        top.pack(side="top", fill="x", padx=5)
        top.pack_propagate(False)
        self.H1_StartBtn = ControlButton(top, "1~4구역 제어 시작", "start", self.on_h1_start)
        self.H1_StopBtn = ControlButton(top, "1~4구역 제어 종료", "stop", self.on_h1_stop)
        self.H2_StartBtn = ControlButton(top, "5~8구역 제어 시작", "start", self.on_h2_start)
        self.H2_StopBtn = ControlButton(top, "5~8구역 제어 종료", "stop", self.on_h2_stop)
        self._show_ctrl("H1", running=False)
        self._show_ctrl("H2", running=False)

        # --- 구역 패널 4열 x 2행 ---
        self.content = tk.Frame(root, bg=T.WHITE)
        self.content.pack(side="top", fill="both", expand=True, padx=2, pady=(0, 2))
        self.line_panels: Dict[str, LinePanel] = {}
        self.zone_panels: Dict[int, ZonePanel] = {}
        for c in range(4):
            self.content.columnconfigure(c, weight=1, uniform="c")
        for r in range(2):
            self.content.rowconfigure(r, weight=1, uniform="r")
        for row, house in enumerate(("H1", "H2")):
            for zone in range(1, 5):
                zp = ZonePanel(self.content, self, house, zone)
                zp.grid(row=row, column=zone - 1, sticky="nsew", padx=3, pady=3)
                self.zone_panels[row * 4 + zone] = zp
                self.line_panels.update(zp.lines)

        # --- 초기화 (원본 MainWindow_Loaded) ---
        self.update_header_box()
        self.refresh_lists(force=True)
        self.update_network_status()
        root.after(C.NETWORK_UPDATE_INTERVAL * 1000, self._tick)
        root.after(LIST_REFRESH_MS, self._list_tick)

    # ------------------------------------------------------------ 타이머
    def _tick(self) -> None:
        """5초 주기: 모터 상태 갱신(원본 TMR_STATUS_UPDATER) + 통신상태 표시(원본 TMR_UPDATE)"""
        try:
            self.plant.update_all_status()
            self.update_network_status()
        except Exception:
            log.exception("상태 갱신 오류")
        self.root.after(C.NETWORK_UPDATE_INTERVAL * 1000, self._tick)

    def _list_tick(self) -> None:
        try:
            self.refresh_lists(force=False)
        except Exception:
            log.exception("목록 갱신 오류")
        self.root.after(LIST_REFRESH_MS, self._list_tick)

    def refresh_lists(self, force: bool) -> None:
        for lp in self.line_panels.values():
            lp.refresh(force=force)

    # ------------------------------------------------------------ 표시
    def _show_ctrl(self, house: str, running: bool) -> None:
        start = getattr(self, f"{house}_StartBtn")
        stop = getattr(self, f"{house}_StopBtn")
        x = 0 if house == "H1" else 150
        if running:
            start.place_forget()
            stop.place(x=x, y=3)
        else:
            stop.place_forget()
            start.place(x=x, y=3)

    def update_header_box(self) -> None:
        """라인 머리 영역의 IP/PORT, RPM 소스 문구 갱신 (원본 Update_HeaderBox)"""
        for code, lp in self.line_panels.items():
            cfg = self.plant.configs[code]
            lp.lbl_old_ip.configure(text=f"{cfg.source_ip}/{cfg.source_port}")
            lp.lbl_new_ip.configure(text=f"{cfg.target_ip}/{cfg.target_port}")
            if cfg.source_manual:
                lp.lbl_source.configure(text=f"수동 > {cfg.manual_rpm} RPM")
            else:
                lp.lbl_source.configure(text=f"{get_house_short_name(cfg.source_line, ' > ')} > {cfg.source_id}")

    def update_network_status(self) -> None:
        """통신상태 점 색상 갱신 (원본 Update_Network_Status)"""
        for code, lp in self.line_panels.items():
            run = self.plant.network_run.get(code[:2], False)
            h = self.plant.houses[code]
            if run:
                lp.rec_SOURCE_NETSTATUS.set_color(T.NET_OK if h.source_net_status else T.NET_ERR)
                lp.rec_TARGET_NETSTATUS.set_color(T.NET_OK if h.target_net_status else T.NET_ERR)
                lp.rec_SID_NETSTATUS.set_color(T.NET_OK if self.plant.reference_ok(code) else T.NET_ERR)
            else:
                for d in (lp.rec_SOURCE_NETSTATUS, lp.rec_TARGET_NETSTATUS, lp.rec_SID_NETSTATUS):
                    d.set_color(T.NET_STOP)

    # ------------------------------------------------------------ 제어 버튼
    def _start(self, house: str, name: str) -> None:
        if messagebox.askyesno("제어", f"{name} 제어를 시작하시겠습니까?", parent=self.root):
            self.plant.start_house(house)
            self._show_ctrl(house, running=True)
            self.update_network_status()

    def _stop(self, house: str, name: str) -> None:
        if messagebox.askyesno("제어", f"{name} 제어를 종료하시겠습니까?", parent=self.root):
            self.plant.stop_house(house)
            self._show_ctrl(house, running=False)
            self.update_network_status()

    def on_h1_start(self) -> None:
        self._start("H1", "1~4구역")

    def on_h1_stop(self) -> None:
        self._stop("H1", "1~4구역")

    def on_h2_start(self) -> None:
        if C.H2_CONTROL_ENABLED:      # 원본(v1.0.0 이전): 버튼은 보이지만 동작하지 않음
            self._start("H2", "5~8구역")

    def on_h2_stop(self) -> None:
        if C.H2_CONTROL_ENABLED:
            self._stop("H2", "5~8구역")

    # ------------------------------------------------------------ 종료
    def on_close(self) -> None:
        """창 닫기 (원본 MainWindow_Closing)"""
        if messagebox.askyesno("종료", "종료하시겠습니까? 종료시 제어기능이 동작하지 않습니다.", parent=self.root):
            try:
                self.plant.shutdown()
            except OSError as e:
                messagebox.showerror("오류", f"설정 저장에 실패했습니다.\n{e}", parent=self.root)
            self.root.destroy()
