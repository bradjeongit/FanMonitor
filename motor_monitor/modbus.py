"""Modbus TCP 프레임 생성/해석. 기준: 운영 실행파일 C_HOUSE (GET_MSG_*, 수신 처리부)

프레임 구조 (MBAP 헤더 7바이트 + PDU)
  [0-1] 트랜잭션 번호  [2-3] 프로토콜 ID(0)  [4-5] 이후 바이트 수  [6] Unit ID(모터 ID)  [7] 기능코드 ...

레지스터 맵
  기존모터 : FC04(Read Input Register) 주소 14(0x0E) 1개 = 현재 RPM
  신규모터 : FC03(Read Holding Register)
             1페이지: 주소 0x00 부터 9개  -> 0x05 = 현재(실제) RPM  (화면 표시용)
             2페이지: 주소 0x13 부터 7개  -> 0x13 자동/수동(1=수동), 0x14 동작/정지(1=동작), 0x16 RPM 설정값 (제어용)
             FC06(Write Single Register) 로 0x13, 0x14, 0x16 을 씀
"""
from __future__ import annotations

import struct
from dataclasses import dataclass
from typing import List, Optional

FC_READ_HOLDING = 0x03
FC_READ_INPUT = 0x04
FC_WRITE_SINGLE = 0x06

OLD_RPM_REG = 0x0E          # 기존모터 RPM 레지스터 (14)
NEW_PAGE1_START = 0x00      # 신규모터 1페이지 읽기 시작 주소
NEW_PAGE1_COUNT = 9         # 신규모터 1페이지 읽기 개수
NEW_REG_ACTUAL_RPM = 0x05   # 신규모터 현재 RPM (1페이지)
NEW_READ_START = 0x13       # 신규모터 2페이지 읽기 시작 주소
NEW_READ_COUNT = 7          # 신규모터 2페이지 읽기 개수 (무조건 7개)
NEW_REG_AUTOMANUAL = 0x13   # 자동/수동
NEW_REG_STARTSTOP = 0x14    # 동작/정지
NEW_REG_RPM = 0x16          # RPM

OLD_RESPONSE_LEN = 11       # 기존모터 읽기 응답 길이
NEW_PAGE1_RESPONSE_LEN = 27 # 신규모터 1페이지 응답 길이
NEW_RESPONSE_LEN = 23       # 신규모터 2페이지 응답 길이


def _frame(tid: int, unit_id: int, pdu: bytes) -> bytes:
    return struct.pack(">HHHB", tid & 0xFFFF, 0, len(pdu) + 1, unit_id & 0xFF) + pdu


def build_old_rpm_read(tid: int, unit_id: int) -> bytes:
    """원본 GET_MSG_SOURCE_RPM_READ"""
    return _frame(tid, unit_id, struct.pack(">BHH", FC_READ_INPUT, OLD_RPM_REG, 1))


def build_new_read_page1(tid: int, unit_id: int) -> bytes:
    """운영 GET_MSG_TAGET_READ_PAGE1 (주소 0x00, 9개)"""
    return _frame(tid, unit_id, struct.pack(">BHH", FC_READ_HOLDING, NEW_PAGE1_START, NEW_PAGE1_COUNT))


def build_new_read(tid: int, unit_id: int) -> bytes:
    """운영 GET_MSG_TAGET_READ_PAGE2 (주소 0x13, 7개)"""
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


@dataclass
class OldRpmResponse:
    unit_id: int
    rpm: int


@dataclass
class NewPage1Response:
    unit_id: int
    actual_rpm: int


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


def parse_new_page1(frame: bytes) -> Optional[NewPage1Response]:
    """길이 27, FC03, 바이트수 18 응답 -> 현재 RPM = 레지스터 0x05 (byte[19..20])."""
    if len(frame) == NEW_PAGE1_RESPONSE_LEN and frame[7] == FC_READ_HOLDING and frame[8] == NEW_PAGE1_COUNT * 2:
        off = 9 + NEW_REG_ACTUAL_RPM * 2
        return NewPage1Response(unit_id=frame[6], actual_rpm=frame[off] * 256 + frame[off + 1])
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


# ---------------------------------------------------------------- 기존모터제어 (ZIEHL-ABEGG ECblue)
# 출처: "ECblue MODBUS Description" L-BAL-E076-GB 1517 Index 013 (FW 1.03~14)
ECB_H_PIN = 0          # h0 PIN 입력 (0010 = PIN 보호 2단계 해제, 약 15분 유지)
ECB_H_CONTROL = 1      # h1 Control: bit0 Enable
ECB_H_SPEED = 2        # h2 Speed control: 제어모드 1 이면 절대 회전수(1/min), 휘발성
ECB_H_COMPARAM = 3     # h3 COM 파라미터 (읽기만)
ECB_H_CONTROLMODE = 4  # h4 Controlmode: bit[3:0] 속도제어 방식, bit8 Buscon enable (NV, PIN 2단계)
ECB_I_SPEED = 14       # i14 실제 회전수(1/min) - 기존모터 읽기와 같은 주소
ECB_PIN_LEVEL2 = 10    # PIN "0010"

ECB_MODE_MASK = 0x000F
ECB_MODE_ABS_SPEED = 1         # h4 bit[3:0] = 1 : h2 = 절대 회전수
ECB_BUSCON_ENABLE = 0x0100     # h4 bit8 : h1 bit0 Enable 사용
ECB_ENABLE_BIT = 0x0001        # h1 bit0


def ecb_mode_ok(h4: int) -> bool:
    """버스 속도제어(h2 절대 회전수) + 버스 Enable 이 켜져 있는지."""
    return (h4 & ECB_MODE_MASK) == ECB_MODE_ABS_SPEED and bool(h4 & ECB_BUSCON_ENABLE)


def ecb_mode_value(h4: int) -> int:
    """h4 의 다른 비트는 유지하고 속도제어 방식=1, Buscon enable=1 로 바꾼 값."""
    return (h4 & ~ECB_MODE_MASK & 0xFFFF) | ECB_MODE_ABS_SPEED | ECB_BUSCON_ENABLE


def build_read_holding(tid: int, unit_id: int, addr: int, count: int) -> bytes:
    return _frame(tid, unit_id, struct.pack(">BHH", FC_READ_HOLDING, addr, count))


def build_read_input(tid: int, unit_id: int, addr: int, count: int) -> bytes:
    return _frame(tid, unit_id, struct.pack(">BHH", FC_READ_INPUT, addr, count))


@dataclass
class Response:
    tid: int
    unit_id: int
    fc: int
    values: List[int]          # 읽기 응답의 레지스터 값 / 쓰기 응답은 [주소, 값]
    exception: Optional[int] = None


def parse_response(frame: bytes) -> Optional[Response]:
    """일반 응답 해석 (FC03/04 읽기, FC06 쓰기 echo, 예외 응답)."""
    if len(frame) < 9:
        return None
    tid = frame[0] * 256 + frame[1]
    uid, fc = frame[6], frame[7]
    if fc & 0x80:
        return Response(tid, uid, fc & 0x7F, [], exception=frame[8])
    if fc in (FC_READ_HOLDING, FC_READ_INPUT):
        n = frame[8]
        data = frame[9:9 + n]
        if len(data) != n or n % 2:
            return None
        return Response(tid, uid, fc, [data[i] * 256 + data[i + 1] for i in range(0, n, 2)])
    if fc == FC_WRITE_SINGLE and len(frame) >= 12:
        return Response(tid, uid, fc, [frame[8] * 256 + frame[9], frame[10] * 256 + frame[11]])
    return None
