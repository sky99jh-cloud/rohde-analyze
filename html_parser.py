"""
HTML 파라미터 스냅샷 파싱 모듈.
Rohde & Schwarz 송신기 로그 HTML에서 필요한 측정값을 추출한다.
"""

import re
from datetime import datetime
from bs4 import BeautifulSoup


def _extract_number(text: str) -> float | None:
    """'15.5 V', '45 dB', '42.4 °C' 등에서 숫자 부분만 추출."""
    text = text.strip()
    match = re.search(r"[-+]?\d+\.?\d*", text)
    if match:
        return float(match.group())
    return None


def _extract_value_and_unit(text: str) -> tuple[float | None, str]:
    """값과 단위를 분리하여 반환. 예: '45 dB' → (45.0, 'dB')"""
    text = text.strip()
    match = re.match(r"([-+]?\d+\.?\d*)\s*(.*)", text)
    if match:
        return float(match.group(1)), match.group(2).strip()
    return None, ""


def _get_section(soup: BeautifulSoup, caption_text: str) -> dict[str, str]:
    """caption 텍스트로 테이블 섹션을 찾아 key→value 딕셔너리로 반환."""
    for caption in soup.find_all("caption"):
        if caption.get_text(strip=True) == caption_text:
            table = caption.find_parent("table")
            result: dict[str, str] = {}
            for row in table.find_all("tr"):
                cells = row.find_all("td")
                if len(cells) >= 2:
                    key = cells[0].get_text(strip=True)
                    val = cells[1].get_text(strip=True)
                    # 중복 키는 첫 번째 값만 사용 (Limit 같은 경우)
                    if key not in result:
                        result[key] = val
            return result
    return {}


def parse_html(html_path: str) -> dict:
    """
    HTML 파일을 파싱하여 Excel에 입력할 데이터를 반환한다.

    반환 구조:
    {
        "created_on": datetime,
        "amp1": { label: value, ... },
        "amp2": { label: value, ... },
        "forward_power": float,
        "reflected_power": float,
        "shoulder_distance": (float, str),
        "shoulder_left": (float, str),
        "shoulder_right": (float, str),
        "measured_ripple": (float, str),
        "measured_group_delay": (float, str),
    }
    """
    with open(html_path, "r", encoding="utf-8") as f:
        content = f.read()

    soup = BeautifulSoup(content, "lxml")

    result: dict = {}

    # ── 생성 날짜 추출 ──────────────────────────────────────────────
    created_td = soup.find("td", class_="key", string=re.compile(r"Created on"))
    if created_td:
        match = re.search(r"(\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2})", created_td.get_text())
        if match:
            result["created_on"] = datetime.fromisoformat(match.group(1))

    # ── Power Limits ────────────────────────────────────────────────
    power_limits = _get_section(soup, "Power Limits » Power and Limits")
    result["forward_power"] = _extract_number(power_limits.get("Forward Power", ""))
    result["reflected_power"] = _extract_number(power_limits.get("Reflected Power", ""))

    # ── Pre-Correction Non Linear ───────────────────────────────────
    non_linear = _get_section(soup, "Exciter » Pre- Correction » Non Linear")
    result["shoulder_distance"] = _extract_value_and_unit(non_linear.get("Shoulder Distance", ""))
    result["shoulder_left"] = _extract_value_and_unit(non_linear.get("Shoulder Left", ""))
    result["shoulder_right"] = _extract_value_and_unit(non_linear.get("Shoulder Right", ""))

    # ── Pre-Correction Linear ───────────────────────────────────────
    linear = _get_section(soup, "Exciter » Pre- Correction » Linear")
    result["measured_ripple"] = _extract_value_and_unit(linear.get("Measured Ripple", ""))
    result["measured_group_delay"] = _extract_value_and_unit(linear.get("Measured Group Delay", ""))

    # ── Amplifier 데이터 추출 (1, 2) ───────────────────────────────
    for amp_n in (1, 2):
        status = _get_section(
            soup,
            f"Output Stage » Rack 1 Amplifiers » Amplifier {amp_n} » Status",
        )
        supply = _get_section(
            soup,
            f"Output Stage » Rack 1 Amplifiers » Amplifier {amp_n} » Supply",
        )
        transistors = _get_section(
            soup,
            f"Output Stage » Rack 1 Amplifiers » Amplifier {amp_n} » Transistors",
        )
        rf_levels = _get_section(
            soup,
            f"Output Stage » Rack 1 Amplifiers » Amplifier {amp_n} » RF Levels",
        )

        amp_data: dict[str, float | None] = {
            "AMP Temp [°C]":     _extract_number(status.get("Amplifier Temp.", "")),
            "V Aux in [V]":      _extract_number(supply.get("V Aux In", "")),
            "V+ Mon [V]":        _extract_number(supply.get("V+ Mon", "")),
            "I DC [A]":          _extract_number(supply.get("I DC", "")),
            "I Pre [A]":         _extract_number(supply.get("I Pre", "")),
            "V5V ACB [V]":       _extract_number(supply.get("V5V ACB", "")),
            "V 3V5 [V]":         _extract_number(supply.get("V 3V5", "")),
            "V 12Mon [V]":       _extract_number(supply.get("V 12 Mon", "")),
            "V Pre Mon [V]":     _extract_number(supply.get("V Pre Mon", "")),
            "I PRE [A]":         _extract_number(transistors.get("I Pre", "")),
            "I DRV [A]":         _extract_number(transistors.get("I Drv", "")),
            "I 1A [A]":          _extract_number(transistors.get("I 1A", "")),
            "I 2A [A]":          _extract_number(transistors.get("I 2A", "")),
            "I 3A [A]":          _extract_number(transistors.get("I 3A", "")),
            "I 1B [A]":          _extract_number(transistors.get("I 1B", "")),
            "I 2B [A]":          _extract_number(transistors.get("I 2B", "")),
            "I 3B [A]":          _extract_number(transistors.get("I 3B", "")),
            "Power A [V]":       _extract_number(rf_levels.get("Power A", "")),
            "Power B [V]":       _extract_number(rf_levels.get("Power B", "")),
            "Power V Ref [V]":   _extract_number(rf_levels.get("Power V Ref", "")),
            "Power Out [V]":     _extract_number(rf_levels.get("Power Out", "")),
            "Reflected Out [V]": _extract_number(rf_levels.get("Reflected Out", "")),
        }

        result[f"amp{amp_n}"] = amp_data

    return result
