"""장비 없이 시험하기 위한 Modbus TCP 모터 게이트웨이 시뮬레이터.

사용 예)
  python -m motor_monitor.simulator --port 5020 --old 5:950,6:1000 --new 10,11
  -> 127.0.0.1:5020 에 기존모터 ID5(950RPM), ID6(1000RPM) 와 신규모터 ID10, ID11 이 있는 게이트웨이 생성

라인설정 창에서 기존모터/신규모터 IP 를 127.0.0.1, PORT 를 시뮬레이터 포트로 지정해 시험합니다.
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
    """기존모터(FC04 레지스터14)와 신규모터(FC03/FC06 레지스터 0x13~0x19)를 흉내냄."""

    def __init__(self, host: str = "127.0.0.1", port: int = 0) -> None:
        self.old_rpm: Dict[int, int] = {}
        self.new_regs: Dict[int, Dict[int, int]] = {}
        self.writes: List[tuple] = []          # (unit_id, reg, value) 기록
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
            regs = {r: 0 for r in range(modbus.NEW_READ_START, modbus.NEW_READ_START + modbus.NEW_READ_COUNT)}
            regs[modbus.NEW_REG_AUTOMANUAL] = automanual
            regs[modbus.NEW_REG_STARTSTOP] = startstop
            regs[modbus.NEW_REG_RPM] = rpm
            self.new_regs[uid] = regs

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
            if fc == modbus.FC_READ_INPUT and uid in self.old_rpm and addr == modbus.OLD_RPM_REG:
                pdu = struct.pack(">BBH", fc, 2, self.old_rpm[uid] & 0xFFFF)
            elif fc == modbus.FC_READ_HOLDING and uid in self.new_regs:
                regs = self.new_regs[uid]
                values = [regs.get(addr + i, 0) for i in range(val)]
                pdu = struct.pack(">BB", fc, val * 2) + b"".join(struct.pack(">H", v) for v in values)
            elif fc == modbus.FC_WRITE_SINGLE and uid in self.new_regs:
                self.new_regs[uid][addr] = val
                self.writes.append((uid, addr, val))
                pdu = req[7:12]
            else:
                return None
        return struct.pack(">HHHB", tid, 0, len(pdu) + 1, uid) + pdu


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
    args = ap.parse_args(argv)
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(message)s")

    gw = FakeGateway(args.host, args.port)
    for uid, rpm in _parse_old(args.old).items():
        gw.add_old(uid, rpm)
    for uid in filter(None, args.new.split(",")):
        gw.add_new(int(uid))
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
