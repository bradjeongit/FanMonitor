"""Modbus TCP 프레임 생성/해석. 원본: C_HOUSE.vb '#Region modbus_function' 및 수신 처리부

프레임 구조 (MBAP 헤더 7바이트 + PDU)
  [0-1] 트랜잭션 번호  [2-3] 프로토콜 ID(0)  [4-5] 이후 바이트 수  [6] Unit ID(모터 ID)  [7] 기능코드 ...

레지스터 맵
  기존모터 : FC04(Read Input Register) 주소 14(0x0E) 1개 = 현재 RPM
  신규모터 : FC03(Read Holding Register) 주소 0x13 부터 7개
             0x13 자동/수동(1=수동), 0x14 동작/정지(1=동작), 0x16 RPM 설정값
             FC06(Write Single Register) 로 각 레지스터를 씀
"""
from __future__ import annotations

import struct
from dataclasses import dataclass
from typing import List, Optional

FC_READ_HOLDING = 0x03
FC_READ_INPUT = 0x04
FC_WRITE_SINGLE = 0x06
FC_WRITE_MULTIPLE = 0x10

OLD_RPM_REG = 0x0E          # 기존모터 RPM 레지스터 (14)
NEW_READ_START = 0x13       # 신규모터 읽기 시작 주소
NEW_READ_COUNT = 7          # 신규모터 읽기 개수 (무조건 7개)
NEW_REG_AUTOMANUAL = 0x13   # 자동/수동
NEW_REG_STARTSTOP = 0x14    # 동작/정지
NEW_REG_RPM = 0x16          # RPM

OLD_RESPONSE_LEN = 11       # 기존모터 읽기 응답 길이
NEW_RESPONSE_LEN = 23       # 신규모터 읽기 응답 길이


def _frame(tid: int, unit_id: int, pdu: bytes) -> bytes:
    return struct.pack(">HHHB", tid & 0xFFFF, 0, len(pdu) + 1, unit_id & 0xFF) + pdu


def build_old_rpm_read(tid: int, unit_id: int) -> bytes:
    """원본 GET_MSG_SOURCE_RPM_READ"""
    return _frame(tid, unit_id, struct.pack(">BHH", FC_READ_INPUT, OLD_RPM_REG, 1))


def build_new_read(tid: int, unit_id: int) -> bytes:
    """원본 GET_MSG_TAGET_READ"""
    return _frame(tid, unit_id, struct.pack(">BHH", FC_READ_HOLDING, NEW_READ_START, NEW_READ_COUNT))


def build_new_write_single(tid: int, unit_id: int, reg: int, value: int) -> bytes:
    return _frame(tid, unit_id, struct.pack(">BHH", FC_WRITE_SINGLE, reg, value & 0xFFFF))


def build_new_rpm_write(tid: int, unit_id: int, rpm: int) -> bytes:
    """원본 GET_MSG_TAGET_RPM_WRITE"""
    return build_new_write_single(tid, unit_id, NEW_REG_RPM, rpm)


def build_new_manual_write(tid: int, unit_id: int) -> bytes:
    """원본 GET_MSG_TAGET_MANUAL_WRITE (자동/수동 레지스터 = 1:수동)"""
    return build_new_write_single(tid, unit_id, NEW_REG_AUTOMANUAL, 1)


def build_new_run_write(tid: int, unit_id: int) -> bytes:
    """원본 GET_MSG_TAGET_RUN_WRITE (동작/정지 레지스터 = 1:동작)"""
    return build_new_write_single(tid, unit_id, NEW_REG_STARTSTOP, 1)


def build_new_write_multiple(tid: int, unit_id: int, automanual: int, startstop: int, rpm: int) -> bytes:
    """원본 GET_MSG_TAGET_WRITE (원본에서도 사용하지 않음 - 참고용)"""
    pdu = struct.pack(">BHHB", FC_WRITE_MULTIPLE, NEW_READ_START, 4, 8)
    pdu += struct.pack(">HHHH", automanual, startstop, 0, rpm)
    return _frame(tid, unit_id, pdu)


@dataclass
class OldRpmResponse:
    unit_id: int
    rpm: int


@dataclass
class NewReadResponse:
    unit_id: int
    automanual: int
    startstop: int
    rpm: int


def parse_old_response(frame: bytes) -> Optional[OldRpmResponse]:
    """길이 11, FC04 응답만 인정 (원본 src_recv_job 조건과 동일)."""
    if len(frame) == OLD_RESPONSE_LEN and frame[7] == FC_READ_INPUT:
        return OldRpmResponse(unit_id=frame[6], rpm=frame[9] * 256 + frame[10])
    return None


def parse_new_response(frame: bytes) -> Optional[NewReadResponse]:
    """길이 23, FC03, 바이트수 14 응답만 인정 (원본 target_recv_job 조건과 동일)."""
    if len(frame) == NEW_RESPONSE_LEN and frame[7] == FC_READ_HOLDING and frame[8] == NEW_READ_COUNT * 2:
        return NewReadResponse(
            unit_id=frame[6],
            automanual=frame[9] * 256 + frame[10],
            startstop=frame[11] * 256 + frame[12],
            rpm=frame[15] * 256 + frame[16],
        )
    return None


class FrameSplitter:
    """TCP 수신 바이트 스트림을 MBAP 길이 필드 기준으로 프레임 단위로 분리.

    원본은 Receive() 1회 = 프레임 1개로 가정했으나, TCP 특성상 여러 응답이 붙어 오거나
    나뉘어 올 수 있어 길이 필드로 분리합니다. 응답이 1개씩 오면 원본과 결과가 같습니다.
    """

    MAX_BUFFER = 4096

    def __init__(self) -> None:
        self._buf = bytearray()

    def feed(self, data: bytes) -> List[bytes]:
        self._buf.extend(data)
        frames: List[bytes] = []
        while len(self._buf) >= 6:
            length = self._buf[4] * 256 + self._buf[5]
            if length < 2 or length > 260:
                # 헤더가 깨진 경우: 버퍼를 비우고 다음 수신부터 다시 맞춤
                self._buf.clear()
                break
            total = 6 + length
            if len(self._buf) < total:
                break
            frames.append(bytes(self._buf[:total]))
            del self._buf[:total]
        if len(self._buf) > self.MAX_BUFFER:
            self._buf.clear()
        return frames


def to_hex(data: bytes) -> str:
    return " ".join(f"{b:02X}" for b in data)
