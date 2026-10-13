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


def zone_number(house_code: str) -> int:
    """화면 구역 번호(1~8). 온실1 1~4구역 -> 1~4, 온실2 1~4구역 -> 5~8. 형식이 틀리면 0."""
    if not _is_line_code(house_code) or not house_code[1].isdigit() or not house_code[3].isdigit():
        return 0
    return (int(house_code[1]) - 1) * 4 + int(house_code[3])


def get_house_name(house_code: str, sep: str = " / ") -> str:
    """'H2R1E' -> '5구역 / 짝' (화면은 1~8구역 단일 체계로 표시)"""
    z = zone_number(house_code)
    if z == 0:
        return ""
    side = "짝" if house_code[4] == "E" else "홀"
    return f"{z}구역{sep}{side}"


def get_house_short_name(house_code: str, sep: str = " / ") -> str:
    """get_house_name 과 같음 (원본 호환용 이름 유지)"""
    return get_house_name(house_code, sep)


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
