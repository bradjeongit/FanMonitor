"""메인 화면 (SCR-MAIN). 기준: 운영 실행파일 MainWindow

구성: 맨 위 모드 표시 띠(색상·문구 + 모드 전환 버튼) + 1~8구역을 가로 8열로 배치. 구역 패널마다
  제목줄(n구역) / 제어 시작·종료 버튼 / 머리 영역(통신상태 점, IP/PORT, RPM 소스, 설정·추가 버튼) / 모터 목록
"""
from __future__ import annotations

import logging
import tkinter as tk
from tkinter import messagebox
from typing import Dict

from .. import constants as C
from ..helpers import get_house_name, zone_number
from ..plant import Plant
from . import theme as T
from .dialogs import LineSettingsDialog, MotorDialog
from .widgets import ControlButton, ImageButton, MotorListView, StatusDot, Tooltip

log = logging.getLogger(__name__)

LIST_REFRESH_MS = 300   # 모터 목록 변경 확인 주기 (운영 프로그램은 데이터 바인딩으로 즉시 반영)


class ZonePanel(tk.Frame):
    """구역 1개(GroupBox). 스펙시트 ZONE-{n}"""

    def __init__(self, master, app: "MainWindow", code: str) -> None:
        super().__init__(master, bg=T.GROUP_BORDER, padx=1, pady=1)
        self.app = app
        self.code = code
        img = app.images
        inner = tk.Frame(self, bg=T.WHITE, padx=3, pady=3)
        inner.pack(fill="both", expand=True)

        # 제목줄
        title = tk.Frame(inner, bg=T.ZONE_TITLE_BG)
        title.pack(side="top", fill="x", pady=3)
        tk.Label(title, text=get_house_name(code), bg=T.ZONE_TITLE_BG, fg=T.WHITE,
                 font=T.font(9, bold=True)).pack(pady=3)

        # 제어 시작/종료 버튼 (같은 자리에서 번갈아 표시)
        btn_row = tk.Frame(inner, bg=T.WHITE, height=33)
        btn_row.pack(side="top", fill="x")
        btn_row.pack_propagate(False)
        self.StartBtn = ControlButton(btn_row, "", "start", self.on_start)
        self.StopBtn = ControlButton(btn_row, "", "stop", self.on_stop)
        self.apply_mode()
        self.show_running(False)

        # 머리 영역 (높이 60)
        head = tk.Frame(inner, bg=T.WHITE, height=60)
        head.pack(side="top", fill="x")
        head.grid_propagate(False)
        head.columnconfigure(0, weight=1)
        for r in range(3):
            head.rowconfigure(r, weight=1, uniform="h")

        def info_row(r: int, icon, tip: str):
            f = tk.Frame(head, bg=T.WHITE)
            f.grid(row=r, column=0, sticky="ew", padx=(5, 0))
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
            # 폭이 좁아 글자가 잘릴 때를 위해 전체 문구를 툴팁으로 표시
            Tooltip(lbl, lambda x, y, lbl=lbl: lbl.cget("text"))
            return dot, lbl

        self.rec_SOURCE_NETSTATUS, self.lbl_old_ip = info_row(0, img.get("oldmotor", 15), "기존모터")
        self.rec_TARGET_NETSTATUS, self.lbl_new_ip = info_row(1, img.get("newmotor", 15), "신규모터")
        self.rec_SID_NETSTATUS, self.lbl_source = info_row(2, None, "")

        btns = tk.Frame(head, bg=T.WHITE, width=30)
        btns.grid(row=0, column=1, rowspan=3, sticky="ns")
        self.Btn_Settings = ImageButton(btns, img.get("settings", 22), "설정", self.open_settings)
        self.Btn_Settings.pack(side="top", padx=4, pady=(2, 0))
        self.Btn_Add = ImageButton(btns, img.get("add", 22), "추가", self.open_add)
        self.Btn_Add.pack(side="bottom", padx=4, pady=(0, 2))

        # 모터 목록
        self.LV = MotorListView(inner, img, on_header_click=self.on_sort, on_double_click=self.open_modify)
        self.LV.pack(side="top", fill="both", expand=True, padx=2)
        self._version = -1

    # --- 제어 버튼 (운영 H1_StartBtn_Click / H1_StopBtn_Click) ---
    @property
    def action(self) -> str:
        return "모니터링" if self.app.plant.monitoring else "제어"

    def apply_mode(self) -> None:
        """현재 실행 모드에 맞게 버튼 문구 변경 (n구역 제어 시작 <-> n구역 모니터링 시작)"""
        n = zone_number(self.code)
        self.StartBtn.set_text(f"{n}구역 {self.action} 시작")
        self.StopBtn.set_text(f"{n}구역 {self.action} 종료")

    def show_running(self, running: bool) -> None:
        if running:
            self.StartBtn.place_forget()
            self.StopBtn.place(x=0, y=2, relwidth=1)
        else:
            self.StopBtn.place_forget()
            self.StartBtn.place(x=0, y=2, relwidth=1)

    def on_start(self) -> None:
        name = get_house_name(self.code)
        if messagebox.askyesno(self.action, f"{name} {self.action}을 시작하시겠습니까?" if self.action == "모니터링"
                               else f"{name} 제어를 시작하시겠습니까?", parent=self.app.root):
            self.app.plant.start_line(self.code)
            self.show_running(True)
            self.app.update_network_status()

    def on_stop(self) -> None:
        name = get_house_name(self.code)
        if messagebox.askyesno(self.action, f"{name} {self.action}을 종료하시겠습니까?" if self.action == "모니터링"
                               else f"{name} 제어를 종료하시겠습니까?", parent=self.app.root):
            self.app.plant.stop_line(self.code)
            self.show_running(False)
            self.app.update_network_status()

    # --- 대화상자 ---
    def open_settings(self) -> None:
        """설정(톱니) 버튼 -> 구역설정 창"""
        LineSettingsDialog(self.app.root, self.app.plant, self.code,
                           on_saved=self.app.update_header_box).show_modal()

    def open_add(self) -> None:
        """추가(+) 버튼 -> 모터 추가 창"""
        MotorDialog(self.app.root, self.app.plant, self.app.images, self.code, "ADD",
                    on_changed=self.refresh).show_modal()

    def open_modify(self, dkey: str) -> None:
        """목록 더블클릭 -> 모터 수정/삭제 창"""
        MotorDialog(self.app.root, self.app.plant, self.app.images, self.code, "MOD", mod_dkey=dkey,
                    on_changed=self.refresh).show_modal()

    def on_sort(self, skey: str) -> None:
        """목록 머리글 클릭 -> 정렬"""
        self.app.plant.houses[self.code].sort(skey)
        self.refresh()

    def refresh(self, force: bool = True) -> None:
        h = self.app.plant.houses[self.code]
        if force or h.version != self._version:
            self._version = h.version
            self.LV.set_rows(h.get_list())


class MainWindow:
    """SCR-MAIN 메인 화면"""

    def __init__(self, root: tk.Tk, plant: Plant) -> None:
        self.root = root
        self.plant = plant
        T.init_fonts(root)
        self.images = T.Images(root)

        root.minsize(1500, 450)          # 운영 MinWidth=1500, MinHeight=450
        root.geometry("1600x900")
        if T.IS_WINDOWS:
            try:
                root.state("zoomed")
            except tk.TclError:
                pass
        root.configure(bg=T.WHITE)
        T.set_window_icon(root)
        root.protocol("WM_DELETE_WINDOW", self.on_close)

        # 모드 표시 띠 (MAIN-MODE-BANNER): 모드에 따라 색상·문구가 바뀌고, 오른쪽 버튼으로 모드 전환
        self.mode_banner = tk.Frame(root)
        self.mode_banner.pack(side="top", fill="x")
        self.mode_label = tk.Label(self.mode_banner, fg=T.WHITE, font=T.font(10, bold=True), pady=4, anchor="w")
        self.mode_label.pack(side="left", fill="x", expand=True, padx=(8, 0))
        self.BtnMode = tk.Button(self.mode_banner, command=self.on_toggle_mode, font=T.font(9, bold=True),
                                 bg=T.WHITE, fg=T.MODE_BUTTON_FG, activebackground=T.LIGHT_GRAY,
                                 relief="flat", bd=0, padx=10, pady=1, cursor="hand2")
        self.BtnMode.pack(side="right", padx=8, pady=3)

        self.content = tk.Frame(root, bg=T.WHITE)
        self.content.pack(fill="both", expand=True, padx=2, pady=2)
        self.content.rowconfigure(0, weight=1)
        self.zone_panels: Dict[str, ZonePanel] = {}
        for i, code in enumerate(C.HOUSE_LINES):
            self.content.columnconfigure(i, weight=1, uniform="z")
            zp = ZonePanel(self.content, self, code)
            zp.grid(row=0, column=i, sticky="nsew", padx=3, pady=3)
            self.zone_panels[code] = zp

        self.apply_mode()
        self.update_header_box()
        self.refresh_lists(force=True)
        self.update_network_status()
        root.after(C.NETWORK_UPDATE_INTERVAL * 1000, self._tick)
        root.after(LIST_REFRESH_MS, self._list_tick)

    # ------------------------------------------------------------ 타이머
    def _tick(self) -> None:
        """5초 주기: 모터 상태 갱신(제어 중인 구역) + 통신상태 점 갱신"""
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
        for zp in self.zone_panels.values():
            zp.refresh(force=force)

    # ------------------------------------------------------------ 실행 모드
    def apply_mode(self) -> None:
        """현재 모드에 맞게 띠 색상·문구, 전환 버튼, 창 제목, 구역 버튼 문구 갱신"""
        if self.plant.monitoring:
            bg = T.MONITOR_BANNER_BG
            text = "모니터링 모드  |  모터 상태를 읽기만 하고, 제어 신호(수동·동작·RPM 설정)는 보내지 않습니다."
            btn = "제어 모드로 전환"
            self.root.title(f"{C.APP_TITLE} [모니터링 모드 - 제어 신호 전송 안 함]")
        else:
            bg = T.CONTROL_BANNER_BG
            text = "제어 모드  |  모터 상태를 읽고, 제어 신호(수동·동작·RPM 설정)를 보냅니다."
            btn = "모니터링 모드로 전환"
            self.root.title(C.APP_TITLE)
        self.mode_banner.configure(bg=bg)
        self.mode_label.configure(text=text, bg=bg)
        self.BtnMode.configure(text=btn)
        for zp in self.zone_panels.values():
            zp.apply_mode()

    def on_toggle_mode(self) -> None:
        """모드 전환 버튼 (MAIN-MODE-BTN). 확인 후 즉시 전환, 실행 중인 구역은 그대로 계속 동작."""
        running = [get_house_name(c) for c, h in self.plant.houses.items() if h.running]
        if self.plant.monitoring:
            msg = "제어 모드로 전환하시겠습니까?\n전환하면 제어 신호(수동·동작·RPM 설정)를 모터로 보냅니다."
            if running:
                msg += (f"\n\n실행 중인 구역({', '.join(running)})에 제어 신호가 바로 전송됩니다."
                        "\n운영 프로그램이 같은 모터를 제어하고 있다면 신호가 겹칠 수 있습니다.")
            if not messagebox.askyesno("제어 모드로 전환", msg, icon="warning", parent=self.root):
                return
            self.plant.set_monitoring(False)
        else:
            msg = ("모니터링 모드로 전환하시겠습니까?\n전환하면 제어 신호 전송을 즉시 멈추고, 모터 상태 읽기만 계속합니다."
                   "\n모터는 마지막으로 받은 설정대로 계속 운전합니다.")
            if not messagebox.askyesno("모니터링 모드로 전환", msg, parent=self.root):
                return
            self.plant.set_monitoring(True)
        self.apply_mode()

    # ------------------------------------------------------------ 표시
    def update_header_box(self) -> None:
        """구역 머리 영역의 IP/PORT, RPM 소스 문구 갱신 (운영 Update_HeaderBox)"""
        for code, zp in self.zone_panels.items():
            cfg = self.plant.configs[code]
            zp.lbl_old_ip.configure(text=f"{cfg.source_ip}/{cfg.source_port}")
            zp.lbl_new_ip.configure(text=f"{cfg.target_ip}/{cfg.target_port}")
            if cfg.source_manual:
                zp.lbl_source.configure(text=f"수동 > {cfg.manual_rpm} RPM")
            else:
                zp.lbl_source.configure(text=f"{get_house_name(cfg.source_line)} > {cfg.source_id}")

    def update_network_status(self) -> None:
        """통신상태 점 색상 갱신 (운영 Update_Network_Status)"""
        for code, zp in self.zone_panels.items():
            h = self.plant.houses[code]
            if h.running:
                zp.rec_SOURCE_NETSTATUS.set_color(T.NET_OK if h.source_net_status else T.NET_ERR)
                zp.rec_TARGET_NETSTATUS.set_color(T.NET_OK if h.target_net_status else T.NET_ERR)
                zp.rec_SID_NETSTATUS.set_color(T.NET_OK if self.plant.reference_ok(code) else T.NET_ERR)
            else:
                for d in (zp.rec_SOURCE_NETSTATUS, zp.rec_TARGET_NETSTATUS, zp.rec_SID_NETSTATUS):
                    d.set_color(T.NET_STOP)

    # ------------------------------------------------------------ 종료
    def on_close(self) -> None:
        """창 닫기 (운영 MainWindow_Closing)"""
        msg = ("종료하시겠습니까?" if self.plant.monitoring
               else "종료하시겠습니까? 종료시 제어기능이 동작하지 않습니다.")
        if messagebox.askyesno("종료", msg, parent=self.root):
            try:
                self.plant.shutdown()
            except OSError as e:
                messagebox.showerror("오류", f"설정 저장에 실패했습니다.\n{e}", parent=self.root)
            self.root.destroy()
