"""공용 함수. 원본: M_Functions.vb"""
import ipaddress
import re

_NON_IP_CHARS = re.compile(r"[^0-9.\-]+")
_NON_DIGITS = re.compile(r"[^0-9]+")


def _is_line_code(code: str) -> bool:
    return (
        len(code) == 5
        and code[0] == "H"
        and code[2] == "R"
        and code[4] in ("E", "O")
    )


def get_house_name(house_code: str, sep: str = " / ") -> str:
    """'H1R2E' -> '온실1 / 2구역 / 짝'"""
    if not _is_line_code(house_code):
        return ""
    side = "짝" if house_code[4] == "E" else "홀"
    return f"온실{house_code[1]}{sep}{house_code[3]}구역{sep}{side}"


def get_house_short_name(house_code: str, sep: str = " / ") -> str:
    """'H1R2E' -> '2구역 / 짝'"""
    if not _is_line_code(house_code):
        return ""
    side = "짝" if house_code[4] == "E" else "홀"
    return f"{house_code[3]}구역{sep}{side}"


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
