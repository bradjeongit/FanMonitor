"""장비 없이 시험하기 위한 Modbus TCP 모터 게이트웨이 시뮬레이터.

사용 예)
  python -m motor_monitor.simulator --port 5020 --old 5:950,6:1000 --new 10,11
  -> 127.0.0.1:5020 에 기존모터 ID5(950RPM), ID6(1000RPM) 와 신규모터 ID10, ID11 이 있는 게이트웨이 생성

구역설정 창에서 기존모터/신규모터 IP 를 127.0.0.1, PORT 를 시뮬레이터 포트로 지정해 시험합니다.
"""
from __future__ import annotations

import argparse
import logging
import socket
import struct
import threading
import time
from typing import Dict, List, Optional

from . import modbus

log = logging.getLogger("motor_monitor.simulator")


class FakeGateway:
    """기존모터(FC04 레지스터14)와 신규모터(FC03 0x00~0x08, 0x13~0x19 / FC06)를 흉내냄.
    신규모터 현재 RPM(0x05)은 동작(0x14=1) 중이면 설정 RPM(0x16)을 따라가고, 정지면 0."""

    def __init__(self, host: str = "127.0.0.1", port: int = 0) -> None:
        self.old_rpm: Dict[int, int] = {}
        self.new_regs: Dict[int, Dict[int, int]] = {}
        self.writes: List[tuple] = []          # (unit_id, reg, value) 기록
        self.ecb: Dict[int, dict] = {}         # ECblue 기존모터 (기존모터제어 시험용)
        self.requests: List[bytes] = []
        self._lock = threading.Lock()
        self._srv = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        self._srv.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        self._srv.bind((host, port))
        self._srv.listen(5)
        self.host, self.port = self._srv.getsockname()[:2]
        self._stop = threading.Event()
        self._clients: List[socket.socket] = []
        self._thread: Optional[threading.Thread] = None

    def add_old(self, uid: int, rpm: int) -> None:
        with self._lock:
            self.old_rpm[uid] = rpm

    def add_new(self, uid: int, automanual: int = 0, startstop: int = 0, rpm: int = 0) -> None:
        with self._lock:
            regs = {r: 0 for r in range(0, modbus.NEW_READ_START + modbus.NEW_READ_COUNT)}
            regs[modbus.NEW_REG_AUTOMANUAL] = automanual
            regs[modbus.NEW_REG_STARTSTOP] = startstop
            regs[modbus.NEW_REG_RPM] = rpm
            self.new_regs[uid] = regs

    def add_ecblue(self, uid: int, analog_rpm: int = 770, h4: int = 0, pin_level: int = 2,
                   product: int = 0x0100, firmware: int = 1430) -> None:
        """ECblue 기존모터. h4 제어모드가 1(+Enable)이면 i14 실제속도 = h2, 아니면 아날로그(E1) 속도.
        pin_level < 2 이면 h0 에 0010 을 먼저 써야 h4 를 쓸 수 있음(아니면 예외 응답).
        레지스터 h0~h37, i0~i49 는 "ECblue MODBUS Description" L-BAL-E076-GB 의 기본값을 따름."""
        h = [0] * 38
        h[3], h[4] = 0xF702, h4
        h[11], h[12], h[14], h[15], h[16], h[18], h[19] = 5, 100, 1, 2, 513, 75, 9999
        h[25] = 0x1414
        h[32], h[33], h[34], h[35], h[36] = 100, 200, 300, 500, 600
        with self._lock:
            self.ecb[uid] = {"h": h, "analog": analog_rpm, "pin_level": pin_level, "pin_ok": False,
                             "pin_rel": 0, "h4_writes": 0, "nv_writes": {}, "product": product,
                             "firmware": firmware, "last_seen": time.monotonic(), "wd_tripped": False}

    def start(self) -> "FakeGateway":
        self._thread = threading.Thread(target=self._accept_loop, daemon=True)
        self._thread.start()
        return self

    def stop(self) -> None:
        self._stop.set()
        try:
            self._srv.close()
        except OSError:
            pass
        for c in list(self._clients):
            try:
                c.shutdown(socket.SHUT_RDWR)
                c.close()
            except OSError:
                pass

    def _accept_loop(self) -> None:
        self._srv.settimeout(0.2)
        while not self._stop.is_set():
            try:
                c, _ = self._srv.accept()
            except socket.timeout:
                continue
            except OSError:
                break
            self._clients.append(c)
            threading.Thread(target=self._client_loop, args=(c,), daemon=True).start()

    def _client_loop(self, c: socket.socket) -> None:
        splitter = modbus.FrameSplitter()
        c.settimeout(0.5)
        while not self._stop.is_set():
            try:
                data = c.recv(512)
            except socket.timeout:
                continue
            except OSError:
                break
            if not data:
                break
            for req in splitter.feed(data):
                resp = self.handle(req)
                if resp:
                    try:
                        c.sendall(resp)
                    except OSError:
                        return

    def handle(self, req: bytes) -> Optional[bytes]:
        if len(req) < 12:
            return None
        tid, _pid, _len, uid, fc = struct.unpack(">HHHBB", req[:8])
        addr, val = struct.unpack(">HH", req[8:12])
        with self._lock:
            self.requests.append(req)
            if uid in self.ecb:
                return self._handle_ecb(tid, uid, fc, addr, val, req)
            if fc == modbus.FC_READ_INPUT and uid in self.old_rpm and addr == modbus.OLD_RPM_REG:
                pdu = struct.pack(">BBH", fc, 2, self.old_rpm[uid] & 0xFFFF)
            elif fc == modbus.FC_READ_HOLDING and uid in self.new_regs:
                regs = self.new_regs[uid]
                running = regs[modbus.NEW_REG_STARTSTOP] == 1
                regs[modbus.NEW_REG_ACTUAL_RPM] = regs[modbus.NEW_REG_RPM] if running else 0
                values = [regs.get(addr + i, 0) for i in range(val)]
                pdu = struct.pack(">BB", fc, val * 2) + b"".join(struct.pack(">H", v) for v in values)
            elif fc == modbus.FC_WRITE_SINGLE and uid in self.new_regs:
                self.new_regs[uid][addr] = val
                self.writes.append((uid, addr, val))
                pdu = req[7:12]
            else:
                return None
        return struct.pack(">HHHB", tid, 0, len(pdu) + 1, uid) + pdu


    # 쓰기에 필요한 PIN 보호 단계 (문서 3.3~3.7). 그 밖은 0
    ECB_PIN_NEED = {3: 1, 5: 1, 6: 1, 7: 1, 8: 1, 9: 1, **{a: 2 for a in (4,) + tuple(range(10, 38))}}
    ECB_NV = {3, 4} | set(range(5, 38))

    def ecb_speed(self, uid: int) -> int:
        d = self.ecb[uid]
        h = d["h"]
        if d.get("wd_tripped") and h[17] & 0xFF in (2, 3):
            spd = h[5]                                   # 통신 끊김: 정속 1 (Set Intern1, h5)
        elif modbus.ecb_mode_ok(h[4]):
            spd = h[2] if h[1] & 1 else 0
        else:
            spd = d["analog"]
        return min(spd, h[8]) if h[8] > 0 else spd      # 최고속도(h8) 제한

    WD_HOLD = 1.0   # 통신이 다시 온 뒤에도 '통신 끊김' 상태를 이만큼 유지 (실제 복귀 동작은 문서에 없음)

    def _ecb_watchdog(self, d) -> None:
        """요청이 올 때 호출: 지난 요청 뒤 워치독 시간(h17 상위 바이트)이 지났으면 '통신 끊김' 상태가 되어
        속도 = Set Intern1(h5), 오류 i12 bit15. 통신이 다시 오면 WD_HOLD 초 뒤 정상으로 돌아감."""
        h = d["h"]
        wd_time, wd_mode = h[17] >> 8, h[17] & 0xFF
        now = time.monotonic()
        if wd_time and wd_mode and now - d["last_seen"] > wd_time:
            d["wd_at"] = now
        d["wd_tripped"] = bool(d.get("wd_at")) and now - d["wd_at"] < self.WD_HOLD
        d["last_seen"] = now

    def _ecb_inputs(self, uid: int, d) -> Dict[int, int]:
        spd = self.ecb_speed(uid)
        err = d.get("err", 0) | (0x8000 if d.get("wd_tripped") else 0)
        inputs = {0: d["firmware"], 1: d["product"], 2: 0, 3: 0x1111, 4: 0x2222, 5: 0x3333, 6: 0x4444,
                  7: 0x5555, 8: 0, 9: 0, 10: 0, 11: 0, 12: err, modbus.ECB_I_SPEED: spd,
                  15: d.get("current", spd // 4), 20: 560, 21: 325, 22: 412, 23: 356, 24: 398,
                  26: 0, 27: min(32767, spd * 20), 30: 0, 31: 0, 33: spd // 3, 49: max(d["pin_level"], d["pin_rel"])}
        if d["firmware"] >= 1400:                       # i13 오류상태 2 는 FW14 부터
            inputs[13] = 0
        return inputs

    def _handle_ecb(self, tid, uid, fc, addr, val, req) -> Optional[bytes]:
        d = self.ecb[uid]
        h = d["h"]
        self._ecb_watchdog(d)

        def frame(pdu):
            return struct.pack(">HHHB", tid, 0, len(pdu) + 1, uid) + pdu
        if fc == modbus.FC_READ_HOLDING and 1 <= val and addr + val <= len(h):
            return frame(struct.pack(">BB", fc, val * 2) + b"".join(struct.pack(">H", v) for v in h[addr:addr + val]))
        if fc == modbus.FC_READ_INPUT:
            # 없는 입력 레지스터(i13: FW14 미만, i16~i19 등)가 범위에 있으면 예외
            inputs = self._ecb_inputs(uid, d)
            regs = [inputs.get(a) for a in range(addr, addr + val)]
            if val >= 1 and None not in regs:
                return frame(struct.pack(">BB", fc, val * 2) + b"".join(struct.pack(">H", v) for v in regs))
        if fc == modbus.FC_WRITE_SINGLE and addr < len(h):
            if addr == modbus.ECB_H_PIN:
                d["pin_ok"] = val == modbus.ECB_PIN_LEVEL2
                d["pin_rel"] = 2 if val == modbus.ECB_PIN_LEVEL2 else (1 if val == 1234 else d["pin_rel"])
            else:
                need = self.ECB_PIN_NEED.get(addr, 0)
                if addr == modbus.ECB_H_CONTROLMODE and d.get("refuse_h4"):
                    return frame(struct.pack(">BB", fc | 0x80, 0x04))
                if max(d["pin_level"], d["pin_rel"]) < need:
                    return frame(struct.pack(">BB", fc | 0x80, 0x04))
                if addr in (5, 6, 7, 8, 9) and val > 6000 or addr in (11, 12, 18) and val > 100:
                    return frame(struct.pack(">BB", fc | 0x80, 0x03))
                if addr == modbus.ECB_H_CONTROLMODE:
                    d["h4_writes"] += 1
                if addr in self.ECB_NV:
                    d["nv_writes"][addr] = d["nv_writes"].get(addr, 0) + 1
            h[addr] = val
            self.writes.append((uid, addr, val))
            return frame(req[7:12])
        return frame(struct.pack(">BB", fc | 0x80, 0x02))

    def add_ecb_from_args(self, text: str) -> None:
        for part in filter(None, text.split(",")):
            uid, _, rpm = part.partition(":")
            self.add_ecblue(int(uid), int(rpm or 770))


def _parse_old(text: str) -> Dict[int, int]:
    out: Dict[int, int] = {}
    for part in filter(None, text.split(",")):
        uid, _, rpm = part.partition(":")
        out[int(uid)] = int(rpm or 0)
    return out


def main(argv=None) -> None:
    ap = argparse.ArgumentParser(description="BLDC 모터 게이트웨이 시뮬레이터")
    ap.add_argument("--host", default="127.0.0.1")
    ap.add_argument("--port", type=int, default=5020)
    ap.add_argument("--old", default="", help="기존모터 목록 'ID:RPM,ID:RPM'")
    ap.add_argument("--new", default="", help="신규모터 ID 목록 'ID,ID'")
    ap.add_argument("--ecblue", default="", help="기존모터제어 시험용 ECblue 'ID:아날로그RPM,...'")
    args = ap.parse_args(argv)
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(message)s")

    gw = FakeGateway(args.host, args.port)
    for uid, rpm in _parse_old(args.old).items():
        gw.add_old(uid, rpm)
    for uid in filter(None, args.new.split(",")):
        gw.add_new(int(uid))
    gw.add_ecb_from_args(args.ecblue)
    gw.start()
    log.info("시뮬레이터 실행 %s:%s  기존=%s 신규=%s", gw.host, gw.port, gw.old_rpm, list(gw.new_regs))
    seen = 0
    try:
        while True:
            time.sleep(1)
            with gw._lock:
                new_writes = gw.writes[seen:]
                seen = len(gw.writes)
            for uid, reg, val in new_writes:
                log.info("쓰기 수신: ID=%s 레지스터=0x%02X 값=%s", uid, reg, val)
    except KeyboardInterrupt:
        gw.stop()


if __name__ == "__main__":
    main()
