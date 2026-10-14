"""기존모터제어(ECblue) 현장 점검 도구.

현장 PC(모터 통신장치와 같은 네트워크)에서 실행합니다. 운영 Config.dll 은 읽지도 바꾸지도 않습니다.

  1단계 - 읽기 전용 (모터에 아무것도 쓰지 않음):
    python tools/ecblue_check.py --ip 172.27.0.95 --ids 180 176

  응답이 없을 때 - 같은 통신장치의 운영 기존모터와 비교, ID 범위 찾기, 응답 대기시간 늘리기:
    python tools/ecblue_check.py --ip 172.27.0.95 --ids 155 159 180 176
    python tools/ecblue_check.py --ip 172.27.0.95 --ids 170-190 247 --timeout 1

  2단계 - 실제 제어 시험 (프로그램의 기존모터제어 코드로 지정 속도 운전 -> 끝나면 원래 설정 복원):
    python tools/ecblue_check.py --ip 172.27.0.95 --ids 180 176 --rpm 500 --seconds 90

결과는 화면과 logs/ecblue_check_날짜_시간.txt 에 함께 남습니다.
"""
from __future__ import annotations

import argparse
import datetime
import os
import socket
import sys
import tempfile
import threading
import time
from typing import Dict, List, Optional

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from motor_monitor import constants as C  # noqa: E402
from motor_monitor import modbus  # noqa: E402
from motor_monitor.constants import MotorType  # noqa: E402
from motor_monitor.house import motor_key  # noqa: E402
from motor_monitor.plant import Plant  # noqa: E402

EXCEPTIONS = {
    1: "지원하지 않는 기능코드",
    2: "없는 레지스터 주소",
    3: "허용되지 않는 값",
    4: "장치 오류",
    6: "장치 사용 중",
    10: "게이트웨이 경로 없음",
    11: "게이트웨이: 해당 ID 모터 응답 없음 (ID/배선/버스 확인)",
}


class Report:
    def __init__(self, path: Optional[str]) -> None:
        self.f = open(path, "w", encoding="utf-8") if path else None
        self.path = path

    def __call__(self, text: str = "") -> None:
        print(text, flush=True)
        if self.f:
            self.f.write(text + "\n")
            self.f.flush()

    def close(self) -> None:
        if self.f:
            self.f.close()


class Link:
    """단순 Modbus TCP 요청/응답 (1건씩)."""

    def __init__(self, ip: str, port: int, timeout: float = 2.0) -> None:
        self.sock = socket.create_connection((ip, port), timeout=5.0)
        self.sock.settimeout(timeout)
        self.tid = 0
        self.raw = b""      # 마지막 요청 동안 받은 모든 바이트 (진단용)

    def request(self, build, *args) -> tuple:
        """(응답 Response 또는 None, 응답시간 ms)"""
        self.tid = (self.tid + 1) % 65536
        frame = build(self.tid, *args)
        t0 = time.monotonic()
        self.sock.sendall(frame)
        splitter = modbus.FrameSplitter()
        self.raw = b""
        deadline = t0 + self.sock.gettimeout()
        while time.monotonic() < deadline:
            try:
                data = self.sock.recv(513)
            except socket.timeout:
                break
            if not data:
                raise OSError("통신장치가 연결을 끊음")
            self.raw += data
            for f in splitter.feed(data):
                r = modbus.parse_response(f)
                if r is not None and r.tid == self.tid:
                    return r, (time.monotonic() - t0) * 1000
        return None, None

    def close(self) -> None:
        try:
            self.sock.close()
        except OSError:
            pass


TIMEOUT = 2.0


def describe(r, raw: bytes = b"") -> str:
    if r is None:
        extra = f", 다른 데이터 수신: {raw[:40].hex(' ')}" if raw else ", 받은 데이터 없음"
        return f"응답 없음 ({TIMEOUT:g}초{extra})"
    if r.exception is not None:
        return f"예외 응답 {r.exception}: {EXCEPTIONS.get(r.exception, '알 수 없음')}"
    return "정상"


def read_state(link: Link, uid: int) -> Dict[str, object]:
    hold, t_h = link.request(modbus.build_read_holding, uid, modbus.ECB_H_CONTROL, 4)
    raw_h = link.raw
    speed, t_i = link.request(modbus.build_read_input, uid, modbus.ECB_I_SPEED, 1)
    st: Dict[str, object] = {"hold": hold, "speed": speed, "t_h": t_h, "t_i": t_i,
                             "raw_h": raw_h, "raw_i": link.raw}
    if hold is not None and hold.exception is None and len(hold.values) >= 4:
        st["h1"], st["h2"], st["h3"], st["h4"] = hold.values[:4]
    if speed is not None and speed.exception is None and speed.values:
        st["i14"] = speed.values[0]
    return st


def fmt_state(st: Dict[str, object]) -> List[str]:
    lines = []
    t = f" ({st['t_h']:.0f} ms)" if st.get("t_h") is not None else ""
    lines.append(f"  h1~h4 읽기 (FC03 주소1, 4개): {describe(st['hold'], st.get('raw_h', b''))}{t}")
    if "h4" in st:
        h1, h2, h3, h4 = st["h1"], st["h2"], st["h3"], st["h4"]
        mode = h4 & modbus.ECB_MODE_MASK
        lines.append(f"    h1 Control  = {h1} (0x{h1:04X})  Enable(bit0)={'켜짐' if h1 & 1 else '꺼짐'}")
        lines.append(f"    h2 Speed    = {h2}")
        lines.append(f"    h3 COM 파라미터 = {h3} (0x{h3:04X})")
        lines.append(f"    h4 Controlmode = {h4} (0x{h4:04X})  속도제어방식(bit3~0)={mode}"
                     f"{' (h2 절대회전수)' if mode == 1 else ' (아날로그 E1 등)' if mode == 0 else ''}"
                     f", 버스 Enable(bit8)={'켜짐' if h4 & modbus.ECB_BUSCON_ENABLE else '꺼짐'}")
        lines.append(f"    -> 버스 속도제어 준비: {'됨' if modbus.ecb_mode_ok(h4) else '안 됨 (2단계에서 프로그램이 h4 를 설정)'}")
    t = f" ({st['t_i']:.0f} ms)" if st.get("t_i") is not None else ""
    lines.append(f"  i14 실제 회전수 읽기 (FC04 주소14): {describe(st['speed'], st.get('raw_i', b''))}{t}"
                 + (f"  -> {st['i14']} RPM" if "i14" in st else ""))
    return lines


# ------------------------------------------------------------------ 1단계
def stage_read(ip: str, port: int, ids: List[int], rep: Report) -> Dict[int, Dict[str, object]]:
    rep(f"[1단계] 읽기 전용 점검 - {ip}:{port} (응답 대기 {TIMEOUT:g}초)")
    try:
        link = Link(ip, port, TIMEOUT)
    except OSError as e:
        rep(f"  접속 실패: {e}")
        rep("  -> PC 와 통신장치의 네트워크(IP 대역, 케이블), 통신장치 전원, 포트 번호를 확인하세요.")
        return {}
    rep("  TCP 접속 성공")
    states = {}
    if len(ids) > 8:
        return scan_ids(link, ids, rep)
    try:
        for uid in ids:
            rep(f"- 모터 ID {uid}")
            st = read_state(link, uid)
            for line in fmt_state(st):
                rep(line)
            states[uid] = st
    except OSError as e:
        rep(f"  통신 오류: {e}")
    finally:
        link.close()
    return states


def scan_ids(link: Link, ids: List[int], rep: Report) -> Dict[int, Dict[str, object]]:
    """ID 가 많을 때: 실제 회전수(i14)만 읽어 응답하는 ID 를 찾음 (운영 기존모터 읽기와 같은 요청)."""
    rep(f"  ID {len(ids)}개 찾기 - 실제 회전수(i14)만 읽음")
    found, states = [], {}
    try:
        for uid in ids:
            r, ms = link.request(modbus.build_read_input, uid, modbus.ECB_I_SPEED, 1)
            if r is not None:
                found.append(uid)
                rep(f"  ID {uid}: {describe(r)}" + (f" -> {r.values[0]} RPM" if r.exception is None and r.values else "")
                    + f" ({ms:.0f} ms)")
            elif link.raw:
                rep(f"  ID {uid}: {describe(r, link.raw)}")
    except OSError as e:
        rep(f"  통신 오류: {e}")
    finally:
        link.close()
    rep(f"  응답한 ID: {found if found else '없음'}")
    return states


# ------------------------------------------------------------------ 2단계
def stage_control(ip: str, port: int, ids: List[int], rpm: int, seconds: int, rep: Report,
                  interval: float = 2.0) -> Dict[int, Optional[int]]:
    """프로그램의 실제 기존모터제어 코드(HouseLine._ctrl_job)로 지정 속도 운전. 마지막 실제 회전수 반환."""
    rep(f"[2단계] 제어 시험 - 운전속도 {rpm} RPM, {seconds}초 (프로그램 기존모터제어 코드 사용)")
    cfg = os.path.join(tempfile.mkdtemp(prefix="ecblue_check_"), "Config.dll")   # 운영 Config.dll 은 건드리지 않음
    plant = Plant(cfg, monitoring=False)
    plant.load()
    h = plant.houses["H1R1"]
    for uid in ids:
        h.add_motor(uid, MotorType.OLD_CTRL, True, ip=ip, port=port, set_rpm=rpm)
    # 구역 전체(기존/신규모터 통신)는 시작하지 않고 기존모터제어 쓰레드만 실행
    stop = threading.Event()
    h._stop = stop
    t = threading.Thread(target=h._ctrl_job, args=(stop,), daemon=True)
    t.start()
    last: Dict[int, Optional[int]] = {uid: None for uid in ids}
    t0 = time.monotonic()
    try:
        while time.monotonic() - t0 < seconds:
            time.sleep(interval)
            parts = []
            for uid in ids:
                m = h.get_motor(motor_key(MotorType.OLD_CTRL, uid, ip, port))
                if m is None:
                    continue
                last[uid] = m.rpm if m.rdate is not None else None
                note = f" [{m.ctrl_note}]" if m.ctrl_note else ""
                parts.append(f"ID{uid}: {last[uid] if last[uid] is not None else '-'} RPM "
                             f"({'응답' if m.rdate is not None else '응답없음'}){note}")
            rep(f"  {time.monotonic() - t0:5.0f}초  " + "  |  ".join(parts))
    except KeyboardInterrupt:
        rep("  (사용자가 중단)")
    finally:
        stop.set()
        t.join(timeout=5)
    return last


def stage_restore(ip: str, port: int, before: Dict[int, Dict[str, object]], rep: Report) -> None:
    rep("[복원] 시험 전 설정으로 되돌리기 (h2 -> h1 -> PIN -> h4)")
    try:
        link = Link(ip, port, TIMEOUT)
    except OSError as e:
        rep(f"  접속 실패로 복원하지 못함: {e}")
        return
    try:
        for uid, st in before.items():
            if "h4" not in st:
                rep(f"- ID {uid}: 시험 전 값을 읽지 못해 복원 생략")
                continue
            cur = read_state(link, uid)
            results = []
            for reg, name in ((modbus.ECB_H_SPEED, "h2"), (modbus.ECB_H_CONTROL, "h1")):
                if cur.get(name) != st[name]:
                    r, _ = link.request(modbus.build_new_write_single, uid, reg, st[name])
                    results.append(f"{name}={st[name]} {describe(r)}")
            if cur.get("h4") != st["h4"]:
                link.request(modbus.build_new_write_single, uid, modbus.ECB_H_PIN, modbus.ECB_PIN_LEVEL2)
                r, _ = link.request(modbus.build_new_write_single, uid, modbus.ECB_H_CONTROLMODE, st["h4"])
                results.append(f"h4={st['h4']} {describe(r)}")
            after = read_state(link, uid)
            same = all(after.get(k) == st[k] for k in ("h1", "h2", "h4"))
            rep(f"- ID {uid}: {', '.join(results) if results else '바뀐 값 없음'} -> "
                f"{'복원 확인' if same else '복원 확인 안 됨 (h1=%s h2=%s h4=%s)' % (after.get('h1'), after.get('h2'), after.get('h4'))}")
    except OSError as e:
        rep(f"  통신 오류: {e}")
    finally:
        link.close()


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="기존모터제어(ECblue) 현장 점검")
    ap.add_argument("--ip", required=True)
    ap.add_argument("--port", type=int, default=502)
    ap.add_argument("--ids", nargs="+", required=True, help="모터 ID. 범위는 170-190 처럼")
    ap.add_argument("--timeout", type=float, default=2.0, help="응답 대기 시간(초), 기본 2")
    ap.add_argument("--rpm", type=int, help="2단계 제어 시험 운전속도 (생략하면 1단계 읽기만)")
    ap.add_argument("--seconds", type=int, default=90, help="2단계 시험 시간(초)")
    ap.add_argument("--yes", action="store_true", help="2단계 시작 확인 질문 생략")
    ap.add_argument("--restore", choices=["ask", "yes", "no"], default="ask", help="2단계 후 원래 설정 복원")
    ap.add_argument("--no-log", action="store_true")
    a = ap.parse_args(argv)
    global TIMEOUT
    TIMEOUT = a.timeout
    ids: List[int] = []
    for part in a.ids:
        lo, _, hi = part.partition("-")
        try:
            ids.extend(range(int(lo), int(hi or lo) + 1))
        except ValueError:
            ap.error(f"--ids {part}: 숫자 또는 170-190 형식")
    a.ids = ids
    if a.rpm is not None and len(ids) > 8:
        ap.error("2단계(--rpm)는 ID 8개 이하로 지정하세요")
    for uid in a.ids:
        if not C.CTRL_ID_MIN <= uid <= C.CTRL_ID_MAX:
            ap.error(f"ID {uid}: {C.CTRL_ID_MIN}~{C.CTRL_ID_MAX} 범위여야 합니다")
    if a.rpm is not None and not 0 <= a.rpm <= C.MAX_RPM:
        ap.error(f"--rpm 은 0~{C.MAX_RPM}")

    path = None
    if not a.no_log:
        os.makedirs("logs", exist_ok=True)
        path = os.path.join("logs", datetime.datetime.now().strftime("ecblue_check_%Y%m%d_%H%M%S.txt"))
    rep = Report(path)
    rep(f"ECblue 점검 {datetime.datetime.now():%Y-%m-%d %H:%M:%S}  대상 {a.ip}:{a.port}  "
        f"ID {a.ids if len(a.ids) <= 8 else f'{len(a.ids)}개 ({a.ids[0]}~{a.ids[-1]})'}")
    before = stage_read(a.ip, a.port, a.ids, rep)
    ok_ids = [uid for uid, st in before.items() if "h4" in st]
    if a.rpm is None:
        rep("")
        rep("1단계 끝. 모터에는 아무것도 쓰지 않았습니다. 제어 시험은 --rpm 값을 붙여 다시 실행하세요.")
        rep.close()
        return 0 if ok_ids else 1
    if not ok_ids:
        rep("읽기에 성공한 모터가 없어 2단계를 진행하지 않습니다.")
        rep.close()
        return 1
    rep("")
    rep("주의: 2단계는 모터 설정을 바꿉니다. h4 제어모드는 전원을 꺼도 유지되는 값이며, 바뀐 동안 모터는")
    rep("      아날로그(0-10V) 신호 대신 통신으로 받은 속도로 돕니다. 시험이 끝나면 원래 값으로 되돌릴 수 있습니다.")
    rep("      운영 프로그램이 같은 모터를 제어 중이면 먼저 멈추세요.")
    if not a.yes and input(f"ID {ok_ids} 를 {a.rpm} RPM 으로 운전 시험할까요? 진행하려면 YES 입력: ").strip() != "YES":
        rep("취소했습니다.")
        rep.close()
        return 1
    last = stage_control(a.ip, a.port, ok_ids, a.rpm, a.seconds, rep)
    rep("")
    rep("[판정]")
    for uid in ok_ids:
        v = last.get(uid)
        if v is None:
            rep(f"- ID {uid}: 실제 회전수 응답 없음")
        else:
            diff = abs(v - a.rpm)
            ok = diff <= max(30, a.rpm * 0.05)
            rep(f"- ID {uid}: 지정 {a.rpm} / 실제 {v} RPM (차이 {diff}) -> "
                f"{'정상: 지정 속도 도달' if ok else '미도달: 가속 중이거나 제어모드 미설정, 위 로그의 [ ] 안내 확인'}")
    restore = a.restore
    if restore == "ask":
        ans = input("시험 전 설정으로 되돌릴까요? (Y/n): ").strip().lower()
        restore = "no" if ans in ("n", "no") else "yes"
    if restore == "yes":
        stage_restore(a.ip, a.port, {uid: before[uid] for uid in ok_ids}, rep)
    else:
        rep("[복원] 하지 않음 - 모터는 통신 제어모드로 남아 있으며, 운전속도 h2 는 전원이 꺼지면 사라집니다.")
    if path:
        rep(f"결과 파일: {os.path.abspath(path)}")
    rep.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
