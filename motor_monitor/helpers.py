"""공용 함수. 기준: 운영 실행파일 M_Functions"""
import ipaddress
import re

_NON_IP_CHARS = re.compile(r"[^0-9.\-]+")
_NON_DIGITS = re.compile(r"[^0-9]+")


def zone_number(house_code: str) -> int:
    """'H1R3' -> 3. 형식이 틀리면 0."""
    if len(house_code) != 4 or house_code[0] != "H" or house_code[2] != "R" or not house_code[3].isdigit():
        return 0
    return int(house_code[3])


def get_house_name(house_code: str) -> str:
    """'H1R3' -> '3구역' (형식이 틀리면 빈 문자열)"""
    z = zone_number(house_code)
    return f"{z}구역" if z else ""


def check_valid_ip(ip: str) -> bool:
    """IPv4 형식 검사 (숫자/점만, 7~15자, 점 3개, 실제 주소로 해석 가능)."""
    if _NON_IP_CHARS.search(ip):
        return False
    if not 7 <= len(ip) <= 15:
        return False
    if ip.count(".") != 3:
        return False
    try:
        ipaddress.IPv4Address(ip)
    except ValueError:
        return False
    return True


def check_integer(txt: str) -> bool:
    """숫자(0-9) 외 문자가 없으면 True (원본과 동일하게 빈 문자열도 True)."""
    return _NON_DIGITS.search(txt) is None
