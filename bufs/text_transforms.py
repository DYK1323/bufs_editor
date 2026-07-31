# -*- coding: utf-8 -*-
"""Pure text, number, date, and clipboard matrix transforms."""

from __future__ import annotations

import re
from datetime import datetime
from decimal import Decimal, ROUND_DOWN, ROUND_FLOOR, ROUND_HALF_UP, ROUND_UP


def add_thousand_commas(text: str) -> str:
    def repl(match: re.Match[str]) -> str:
        raw = match.group("number")
        sign = ""
        if raw.startswith("-"):
            sign, raw = "-", raw[1:]
        integer, dot, decimal = raw.partition(".")
        if len(integer) <= 3:
            return sign + raw
        return sign + f"{int(integer):,}" + (dot + decimal if dot else "")

    # 날짜/전화번호처럼 하이픈으로 이어지는 숫자와 2026년/2025학년도 같은 연도 표기는 피한다.
    return re.sub(r"(?<![\d,.-])(?P<number>-?\d{4,}(?:\.\d+)?)(?![\d,.-]|년|년도|학년도|월|일)", repl, text)


def remove_number_commas(text: str) -> str:
    def repl(match: re.Match[str]) -> str:
        return match.group("number").replace(",", "")

    return re.sub(r"(?<![\d,])(?P<number>-?\d{1,3}(?:,\d{3})+(?:\.\d+)?)(?![\d,])", repl, text)


class SuspiciousDecimalNumberError(ValueError):
    def __init__(self, values: list[str]) -> None:
        self.values = values
        preview = ", ".join(values[:5])
        if len(values) > 5:
            preview += f" 외 {len(values) - 5}개"
        super().__init__(preview)


class UnsafeUnitConversionNumberError(ValueError):
    def __init__(self, values: list[str]) -> None:
        self.values = values
        preview = ", ".join(values[:5])
        if len(values) > 5:
            preview += f" 외 {len(values) - 5}개"
        super().__init__(preview)


CURRENCY_UNIT_FACTORS = {
    "원": Decimal("1"),
    "천": Decimal("1000"),
    "천원": Decimal("1000"),
    "백만": Decimal("1000000"),
    "백만원": Decimal("1000000"),
    "억원": Decimal("100000000"),
}
UNIT_CONVERSION_TARGET_LABELS = {
    "원": "원",
    "천": "천원",
    "백만": "백만원",
}
NUMBER_PATTERN = r"-?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?"
UNSAFE_UNIT_PATTERN = r"%|％|㎡|m\^2|m2|m²|평|명|개|건|회"


def find_suspicious_decimal_numbers(text: str) -> list[str]:
    candidates = re.finditer(
        r"(?<![\d.-])(?P<number>-?\d[\d,]*(?:\.\d+)?)(?![\d.-]|년|년도|학년도|월|일)",
        text,
    )
    suspicious: list[str] = []
    seen: set[str] = set()
    valid_comma_number = re.compile(r"-?\d{1,3}(?:,\d{3})+(?:\.\d+)?")
    for match in candidates:
        value = match.group("number")
        if "," not in value:
            continue
        if valid_comma_number.fullmatch(value):
            continue
        if value not in seen:
            suspicious.append(value)
            seen.add(value)
    return suspicious


def ensure_no_suspicious_decimal_numbers(text: str) -> None:
    suspicious = find_suspicious_decimal_numbers(text)
    if suspicious:
        raise SuspiciousDecimalNumberError(suspicious)


def find_unsafe_unit_conversion_numbers(text: str) -> list[str]:
    candidates = re.finditer(
        rf"(?<![\d,.-])(?P<value>{NUMBER_PATTERN})(?P<space>[ \t]*)(?P<unit>{UNSAFE_UNIT_PATTERN})",
        text,
        flags=re.IGNORECASE,
    )
    unsafe: list[str] = []
    seen: set[str] = set()
    for match in candidates:
        value = match.group("value") + match.group("space") + match.group("unit")
        if value not in seen:
            unsafe.append(value)
            seen.add(value)
    return unsafe


def ensure_no_unsafe_unit_conversion_numbers(text: str) -> None:
    unsafe = find_unsafe_unit_conversion_numbers(text)
    if unsafe:
        raise UnsafeUnitConversionNumberError(unsafe)


def find_unexpected_currency_unit_numbers(text: str, expected_source_units: set[str]) -> list[str]:
    unit_names = "|".join(sorted(map(re.escape, CURRENCY_UNIT_FACTORS), key=len, reverse=True))
    candidates = re.finditer(
        rf"(?<![\d,.-])(?P<value>{NUMBER_PATTERN})(?P<space>[ \t]*)(?P<unit>{unit_names})",
        text,
    )
    unexpected: list[str] = []
    seen: set[str] = set()
    for match in candidates:
        unit = match.group("unit")
        if unit in expected_source_units:
            continue
        value = match.group("value") + match.group("space") + unit
        if value not in seen:
            unexpected.append(value)
            seen.add(value)
    return unexpected


def ensure_safe_decimal_unit_conversion_text(text: str, expected_source_units: set[str]) -> None:
    ensure_no_suspicious_decimal_numbers(text)
    ensure_no_unsafe_unit_conversion_numbers(text)
    unexpected = find_unexpected_currency_unit_numbers(text, expected_source_units)
    if unexpected:
        raise UnsafeUnitConversionNumberError(unexpected)


def format_decimal_value(value: Decimal, places: int, use_commas: bool) -> str:
    if places <= 0:
        text = f"{value:.0f}"
    else:
        text = f"{value:.{places}f}"
    if not use_commas:
        return text
    sign = ""
    if text.startswith("-"):
        sign, text = "-", text[1:]
    integer, dot, decimal_part = text.partition(".")
    return sign + f"{int(integer):,}" + (dot + decimal_part if dot else "")


def scale_decimal_numbers(
    text: str,
    multiplier: Decimal,
    places: int,
    mode: str,
    use_commas: bool = True,
) -> str:
    if places < 0:
        raise ValueError("소수 자릿수는 0 이상이어야 합니다.")
    rounding_by_mode = {
        "반올림": ROUND_HALF_UP,
        "올림": ROUND_UP,
        "버림": ROUND_DOWN,
        "내림": ROUND_FLOOR,
    }
    if mode not in rounding_by_mode:
        raise ValueError(f"알 수 없는 소수 처리 방식입니다: {mode}")
    quant = Decimal(1).scaleb(-places)
    rounding = rounding_by_mode[mode]

    def repl(match: re.Match[str]) -> str:
        raw = match.group("number")
        normalized = raw.replace(",", "")
        value = (Decimal(normalized) * multiplier).quantize(quant, rounding=rounding)
        if value == 0:
            value = abs(value)
        return format_decimal_value(value, places, use_commas)

    return re.sub(
        r"(?<![\d,.-])(?P<number>-?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?)(?![\d,.-]|년|년도|학년도|월|일)",
        repl,
        text,
    )


def scale_decimal_numbers_for_unit_conversion(
    text: str,
    multiplier: Decimal,
    unit_transitions: dict[str, str],
    places: int,
    mode: str,
    use_commas: bool = True,
) -> str:
    if places < 0:
        raise ValueError("소수 자릿수는 0 이상이어야 합니다.")
    rounding_by_mode = {
        "반올림": ROUND_HALF_UP,
        "올림": ROUND_UP,
        "버림": ROUND_DOWN,
        "내림": ROUND_FLOOR,
    }
    if mode not in rounding_by_mode:
        raise ValueError(f"알 수 없는 소수 처리 방식입니다: {mode}")
    quant = Decimal(1).scaleb(-places)
    rounding = rounding_by_mode[mode]

    unit_names = "|".join(sorted(map(re.escape, CURRENCY_UNIT_FACTORS), key=len, reverse=True))
    number_re = re.compile(
        rf"(?<![\d,.-])(?P<number>{NUMBER_PATTERN})"
        rf"(?:(?P<space>[ \t]*)(?P<unit>{unit_names})|"
        rf"(?![\d,.-]|년|년도|학년도|월|일|[A-Za-z가-힣%％㎡²]))"
    )

    def repl(match: re.Match[str]) -> str:
        raw = match.group("number")
        attached_unit = match.group("unit") or ""
        normalized = raw.replace(",", "")
        value = Decimal(normalized)
        if attached_unit:
            target_unit = unit_transitions.get(attached_unit)
            if target_unit is None:
                return match.group(0)
            value = value * multiplier
            suffix = target_unit
        else:
            value = value * multiplier
            suffix = ""
        value = value.quantize(quant, rounding=rounding)
        if value == 0:
            value = abs(value)
        return format_decimal_value(value, places, use_commas) + suffix

    return number_re.sub(repl, text)


def convert_decimal_numbers_to_currency_unit(
    text: str,
    target_unit: str,
    places: int,
    mode: str,
    use_commas: bool = True,
) -> str:
    if places < 0:
        raise ValueError("소수 자릿수는 0 이상이어야 합니다.")
    target_suffix = UNIT_CONVERSION_TARGET_LABELS.get(target_unit, target_unit)
    target_factor = CURRENCY_UNIT_FACTORS.get(target_suffix)
    if target_factor is None:
        raise ValueError(f"알 수 없는 대상 단위입니다: {target_unit}")
    rounding_by_mode = {
        "반올림": ROUND_HALF_UP,
        "올림": ROUND_UP,
        "버림": ROUND_DOWN,
        "내림": ROUND_FLOOR,
    }
    if mode not in rounding_by_mode:
        raise ValueError(f"알 수 없는 소수 처리 방식입니다: {mode}")
    quant = Decimal(1).scaleb(-places)
    rounding = rounding_by_mode[mode]

    unit_names = "|".join(sorted(map(re.escape, CURRENCY_UNIT_FACTORS), key=len, reverse=True))
    number_re = re.compile(
        rf"(?<![\d,.-])(?P<number>{NUMBER_PATTERN})"
        rf"(?:(?P<space>[ \t]*)(?P<unit>{unit_names})|"
        rf"(?![\d,.-]|년|년도|학년도|월|일|[A-Za-z가-힣%％㎡²]))"
    )

    def repl(match: re.Match[str]) -> str:
        raw = match.group("number")
        source_unit = match.group("unit") or "원"
        source_factor = CURRENCY_UNIT_FACTORS.get(source_unit)
        if source_factor is None or source_unit == "억원":
            return match.group(0)
        won_value = Decimal(raw.replace(",", "")) * source_factor
        converted = (won_value / target_factor).quantize(quant, rounding=rounding)
        if converted == 0:
            converted = abs(converted)
        return format_decimal_value(converted, places, use_commas) + target_suffix

    return number_re.sub(repl, text)


def normalize_decimal_numbers(text: str, places: int, mode: str, use_commas: bool = True) -> str:
    return scale_decimal_numbers(text, Decimal(1), places, mode, use_commas)


def strip_wrapping_blank_lines(text: str) -> str:
    return re.sub(r"\A(?:[ \t]*[\r\n]+)+|(?:[\r\n]+[ \t]*)+\Z", "", text)


DATE_RE = re.compile(
    r"(?<!\d)"
    r"(?P<year>\d{4})"
    r"(?:\s*년\s*|\s*\.\s*|\s*-\s*)"
    r"(?P<month>\d{1,2})"
    r"(?:\s*월\s*|\s*\.\s*|\s*-\s*)"
    r"(?P<day>\d{1,2})"
    r"(?:\s*일|\s*\.)?"
    r"(?!\d)"
)
COMPACT_DATE_RE = re.compile(
    r"(?<![\d,.-])"
    r"(?P<date>(?P<year>\d{4})(?P<month>\d{2})(?P<day>\d{2})|(?P<short_year>\d{2})(?P<short_month>\d{2})(?P<short_day>\d{2}))"
    r"(?![\d,.-])"
)
YEAR_ONLY_RE = re.compile(
    r"(?<![\d,.-])"
    r"(?P<quote>['’]?)"
    r"(?P<year>\d{4}|\d{2})"
    r"(?P<suffix>학년도|년)"
    r"(?!도)"
    r"(?!\s*\d{1,2}\s*[월일])"
)


WEEKDAY_DATE_RE = re.compile(
    r"(?<!\d)"
    r"(?P<date>"
    r"(?P<year>\d{4})"
    r"(?:\s*년\s*|\s*\.\s*|\s*-\s*)"
    r"(?P<month>\d{1,2})"
    r"(?:\s*월\s*|\s*\.\s*|\s*-\s*)"
    r"(?P<day>\d{1,2})"
    r"(?:\s*일|\s*\.)?"
    r")"
    r"(?:\s*\([월화수목금토일]\))?"
    r"(?!\d)"
)

KOREAN_WEEKDAYS = ("월", "화", "수", "목", "금", "토", "일")


def normalize_date_parts(year: int, month: int, day: int, style: str, original: str) -> str:
    try:
        datetime(year, month, day)
    except ValueError:
        return original
    if style == "korean":
        return f"{year}년 {month}월 {day}일"
    if style == "dot":
        return f"{year}. {month}. {day}."
    if style == "dot_padded":
        return f"{year}. {month:02d}. {day:02d}."
    return original


def normalize_dates(text: str, style: str) -> str:
    def repl(match: re.Match[str]) -> str:
        year = int(match.group("year"))
        month = int(match.group("month"))
        day = int(match.group("day"))
        if not (1 <= month <= 12 and 1 <= day <= 31):
            return match.group(0)
        return normalize_date_parts(year, month, day, style, match.group(0))

    def compact_repl(match: re.Match[str]) -> str:
        if match.group("short_year") is not None:
            short_year = int(match.group("short_year"))
            year = 2000 + short_year if short_year <= 49 else 1900 + short_year
            month = int(match.group("short_month"))
            day = int(match.group("short_day"))
        else:
            year = int(match.group("year"))
            month = int(match.group("month"))
            day = int(match.group("day"))
        return normalize_date_parts(year, month, day, style, match.group(0))

    return COMPACT_DATE_RE.sub(compact_repl, DATE_RE.sub(repl, text))


def full_year_from_short_or_full(value: str) -> int:
    year = int(value)
    if len(value) == 2:
        return 2000 + year if year <= 49 else 1900 + year
    return year


def normalize_years(text: str, style: str) -> str:
    def repl(match: re.Match[str]) -> str:
        year = full_year_from_short_or_full(match.group("year"))
        suffix = match.group("suffix")
        if style == "full":
            return f"{year}{suffix}"
        short_year = f"{year % 100:02d}"
        if style == "short":
            return f"{short_year}{suffix}"
        if style == "apostrophe":
            return f"'{short_year}{suffix}"
        return match.group(0)

    return YEAR_ONLY_RE.sub(repl, text)


def add_weekdays_to_dates(text: str) -> str:
    def repl(match: re.Match[str]) -> str:
        year = int(match.group("year"))
        month = int(match.group("month"))
        day = int(match.group("day"))
        try:
            weekday = KOREAN_WEEKDAYS[datetime(year, month, day).weekday()]
        except ValueError:
            return match.group(0)
        return f"{match.group('date')}({weekday})"

    return WEEKDAY_DATE_RE.sub(repl, text)


def remove_weekdays_from_dates(text: str) -> str:
    return re.sub(r"\s*\([월화수목금토일]\)", "", text)


def is_wrapped_regulation_name(text: str) -> bool:
    return (text.startswith(("｢", "「")) and text.endswith(("｣", "」")))


def is_standard_regulation_wrapper(text: str) -> bool:
    return text.startswith("｢") and text.endswith("｣")


def needs_regulation_wrapper_conversion(text: str) -> bool:
    return is_wrapped_regulation_name(text) and not is_standard_regulation_wrapper(text)


def looks_like_multiline_number_block(text: str) -> bool:
    lines = [line.strip() for line in text.replace("\r\n", "\n").replace("\r", "\n").split("\n")]
    values = [line for line in lines if line]
    if len(values) < 2:
        return False
    return all(re.fullmatch(r"-?\d{1,3}(?:,\d{3})*(?:\.\d+)?|-?\d{4,}(?:\.\d+)?", value) for value in values)


def looks_like_multiline_date_block(text: str) -> bool:
    lines = [line.strip() for line in text.replace("\r\n", "\n").replace("\r", "\n").split("\n")]
    values = [line for line in lines if line]
    if len(values) < 2:
        return False
    return all(DATE_RE.search(value) or COMPACT_DATE_RE.search(value) for value in values)


def parse_cell_clipboard_matrix(text: str) -> list[list[str]]:
    normalized = text.replace("\r\n", "\n").replace("\r", "\n")
    normalized = normalized.rstrip("\n")
    if not normalized:
        return []
    return [row.split("\t") for row in normalized.split("\n")]


def is_rectangular_cell_matrix(matrix: list[list[str]]) -> bool:
    if not matrix:
        return False
    width = len(matrix[0])
    return width > 0 and all(len(row) == width for row in matrix)


def cell_matrix_size(matrix: list[list[str]]) -> tuple[int, int]:
    if not matrix:
        return 0, 0
    return len(matrix), len(matrix[0])


def transform_cell_matrix(
    matrix: list[list[str]],
    transform,
    *,
    strip_wrapping_lines: bool = False,
) -> tuple[list[list[str]], int]:
    transformed_matrix: list[list[str]] = []
    changed = 0
    for row in matrix:
        transformed_row: list[str] = []
        for cell_text in row:
            source_text = strip_wrapping_blank_lines(cell_text) if strip_wrapping_lines else cell_text
            transformed = transform(source_text) if source_text else source_text
            if transformed != source_text:
                changed += 1
            transformed_row.append(transformed)
        transformed_matrix.append(transformed_row)
    return transformed_matrix, changed


def cell_matrix_to_tsv(matrix: list[list[str]]) -> str:
    return "\r\n".join("\t".join(row) for row in matrix)


def looks_like_single_copied_cell(text: str) -> bool:
    if "\t" in text:
        return False
    lines = [line.strip() for line in text.replace("\r\n", "\n").replace("\r", "\n").split("\n")]
    return sum(1 for line in lines if line) <= 1


def flatten_cell_matrix(matrix: list[list[str]], *, strip_wrapping_lines: bool = False) -> list[str]:
    values: list[str] = []
    for row in matrix:
        for cell_text in row:
            values.append(strip_wrapping_blank_lines(cell_text) if strip_wrapping_lines else cell_text.strip())
    return values


def normalize_clipboard_newlines(text: str) -> str:
    return text.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")


def looks_like_cell_clipboard_matrix(text: str) -> bool:
    rows = parse_cell_clipboard_matrix(text)
    if not rows:
        return False
    if any(len(row) > 1 for row in rows):
        return True
    return looks_like_multiline_number_block(text) or looks_like_multiline_date_block(text)
