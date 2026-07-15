# -*- coding: utf-8 -*-
"""
한글 문서 스타일 자동화 도구 MVP.

1차 목적:
- 한글 COM 연결 확인
- bufs 안의 스타일/표지/로고 자산 읽기
- 색상/선 프리셋 확인
- 숫자 쉼표 넣기/빼기 유틸 검증

주의:
- 이 MVP는 원본 한글 문서를 수정하지 않는 읽기 중심 도구다.
- 실제 스타일 적용/표 테두리 적용은 다음 단계에서 별도 검증 후 연결한다.
"""

from __future__ import annotations

import json
import re
import shutil
import sys
import threading
import time
import unicodedata
import zipfile
from dataclasses import dataclass
from datetime import datetime
from html import escape
from html import unescape
from pathlib import Path
from typing import Iterable

ROOT = Path(__file__).resolve().parent
if Path.cwd().resolve() != ROOT:
    # 설치 폴더 루트에는 쉽다한글이 번들링한 python312.dll이 있어
    # Python 3.11 가상환경의 pywin32 로딩과 충돌할 수 있다.
    import os

    os.chdir(ROOT)

try:
    import tkinter as tk
    from tkinter import ttk, messagebox
except Exception as exc:  # pragma: no cover
    raise SystemExit(f"Tkinter를 불러오지 못했습니다: {exc}")

try:
    import pythoncom
    import win32clipboard
    import win32com.client
    import win32con
    import win32gui
except Exception:
    win32com = None  # type: ignore[assignment]
    pythoncom = None  # type: ignore[assignment]
    win32clipboard = None  # type: ignore[assignment]
    win32con = None  # type: ignore[assignment]
    win32gui = None  # type: ignore[assignment]


STYLE_FILE = ROOT / "보고서 본문 서식.hwpx"
COVER_FILE = ROOT / "표지.hwpx"
LOGO_DIR = ROOT / "logos"
TEST_OUTPUT_DIR = ROOT / "test-output"
STYLE_ORDER_FILE = ROOT / "style-order.json"
TABLE_SETTINGS_FILE = ROOT / "table-settings.json"
LAST_HWP_CONNECTION_LOG: list[str] = []


@dataclass(frozen=True)
class PaletteColor:
    name: str
    rgb: tuple[int, int, int]
    hex_value: str
    purpose: str


@dataclass(frozen=True)
class StyleRecord:
    style_id: int
    style_type: str
    name: str


@dataclass
class HwpCandidate:
    display_name: str
    hwp: object
    score: int
    visible: bool
    path: str
    documents: str
    windows: str


DEFAULT_PALETTE = [
    {"name": "노랑", "rgb": [255, 203, 5], "purpose": "제목 셀, 강조 셀"},
    {"name": "진회색", "rgb": [85, 85, 85], "purpose": "보조 글자색, 진한 선"},
    {"name": "회색", "rgb": [150, 150, 150], "purpose": "보조선, 약한 글자색"},
    {"name": "연회색", "rgb": [204, 204, 204], "purpose": "표 셀 배경, 구분 배경"},
    {"name": "검정", "rgb": [0, 0, 0], "purpose": "기본 글자색"},
    {"name": "흰색", "rgb": [255, 255, 255], "purpose": "반전 글자색, 흰 배경"},
]


LINE_PRESETS = {
    "얇은 선": {"width": "0.12mm", "type": "Solid"},
    "굵은 선": {"width": "0.3mm", "type": "Solid"},
    "이중 선": {"width": "0.7mm", "type_candidates": ("DoubleSlim", "SlimThick", "ThickSlim")},
    "점선": {"width": "0.12mm", "type": "Dot"},
}

DEFAULT_TABLE_SETTINGS = {
    "palette": DEFAULT_PALETTE,
    "line_presets": {
        "thin": {"label": "얇은 선", "width": "0.12mm", "type": "solid", "color": "검정"},
        "thick": {"label": "굵은 선", "width": "0.3mm", "type": "solid", "color": "검정"},
        "double": {"label": "이중 선", "width": "0.7mm", "type": "double", "color": "검정"},
        "dotted": {"label": "점선", "width": "0.12mm", "type": "dot", "color": "검정"},
        "none": {"label": "선 없음", "width": "0.12mm", "type": "none", "color": "검정"},
    },
    "title_cell": {
        "background_color": "노랑",
        "text_color": "검정",
        "paragraph_style": "표내용-중간",
        "character_style": "표내용-굵게",
        "borders": {
            "top": "thick",
            "bottom": "thick",
            "left": "thin",
            "right": "thin",
            "inside_horz": "thin",
            "inside_vert": "thin",
        },
    },
    "table_margins": {
        "cell_left": "1.0",
        "cell_right": "1.0",
        "cell_top": "1.0",
        "cell_bottom": "1.0",
        "outside_left": "0.0",
        "outside_right": "0.0",
        "outside_top": "0.0",
        "outside_bottom": "0.0",
    },
}


def read_zip_text(path: Path, entry_name: str) -> str:
    with zipfile.ZipFile(path) as zf:
        with zf.open(entry_name) as fp:
            return fp.read().decode("utf-8", errors="ignore")


def read_style_records(path: Path = STYLE_FILE) -> list[StyleRecord]:
    if not path.exists():
        return []
    try:
        header = read_zip_text(path, "Contents/header.xml")
    except Exception:
        return []
    records: list[StyleRecord] = []
    pattern = re.compile(r'<hh:style\b([^>]*)/>')
    for match in pattern.finditer(header):
        attrs = dict(re.findall(r'(\w+)="([^"]*)"', match.group(1)))
        if "id" not in attrs or "name" not in attrs:
            continue
        records.append(
            StyleRecord(
                style_id=int(attrs["id"]),
                style_type=attrs.get("type", ""),
                name=unescape(attrs["name"]),
            )
        )
    return sorted(records, key=lambda item: (item.style_type != "PARA", item.style_id, item.name))


def load_style_order() -> list[str]:
    if not STYLE_ORDER_FILE.exists():
        return []
    try:
        data = json.loads(STYLE_ORDER_FILE.read_text(encoding="utf-8"))
    except Exception:
        return []
    if not isinstance(data, dict):
        return []
    names = data.get("style_order", [])
    if not isinstance(names, list):
        return []
    return [str(name) for name in names]


def save_style_order(records: list[StyleRecord]) -> None:
    data = {"style_order": [record.name for record in records]}
    STYLE_ORDER_FILE.write_text(
        json.dumps(data, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )


def apply_saved_style_order(records: list[StyleRecord]) -> list[StyleRecord]:
    saved_names = load_style_order()
    if not saved_names:
        return records

    by_name = {record.name: record for record in records}
    ordered: list[StyleRecord] = []
    used: set[str] = set()
    for name in saved_names:
        record = by_name.get(name)
        if record is None or name in used:
            continue
        ordered.append(record)
        used.add(name)
    ordered.extend(record for record in records if record.name not in used)
    return ordered


def merge_dict(default: dict, loaded: dict) -> dict:
    merged = dict(default)
    for key, value in loaded.items():
        if isinstance(value, dict) and isinstance(default.get(key), dict):
            merged[key] = merge_dict(default[key], value)
        else:
            merged[key] = value
    return merged


def normalize_table_settings(settings: dict) -> dict:
    title = settings.get("title_cell")
    if isinstance(title, dict):
        for key in ("apply_hwp_header", "apply_cell_background", "apply_cell_borders"):
            title.pop(key, None)
    settings["line_presets"]["none"]["label"] = "선 없음"
    return settings


def load_table_settings() -> dict:
    if not TABLE_SETTINGS_FILE.exists():
        settings = merge_dict(DEFAULT_TABLE_SETTINGS, {})
        return normalize_table_settings(settings)
    try:
        data = json.loads(TABLE_SETTINGS_FILE.read_text(encoding="utf-8"))
    except Exception:
        settings = merge_dict(DEFAULT_TABLE_SETTINGS, {})
        return normalize_table_settings(settings)
    if not isinstance(data, dict):
        settings = merge_dict(DEFAULT_TABLE_SETTINGS, {})
        return normalize_table_settings(settings)
    settings = merge_dict(DEFAULT_TABLE_SETTINGS, data)
    return normalize_table_settings(settings)


def save_table_settings(settings: dict) -> None:
    TABLE_SETTINGS_FILE.write_text(
        json.dumps(settings, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )


def normalize_style_name(name: str) -> str:
    text = unicodedata.normalize("NFKC", name)
    text = text.replace("\ufeff", "").replace("\u200b", "").replace("\u200c", "").replace("\u200d", "")
    text = text.replace("\u00a0", " ")
    text = re.sub(r"[\u2010-\u2015\u2212\uff0d]", "-", text)
    text = re.sub(r"\s+", "", text)
    return text.casefold()


def rgb_to_hex(rgb: tuple[int, int, int]) -> str:
    red, green, blue = rgb
    return f"#{red:02X}{green:02X}{blue:02X}"


def palette_from_settings(settings: dict) -> list[PaletteColor]:
    colors: list[PaletteColor] = []
    for item in settings.get("palette", []):
        if not isinstance(item, dict):
            continue
        name = str(item.get("name", "")).strip()
        rgb_raw = item.get("rgb", [])
        if not name or not isinstance(rgb_raw, list) or len(rgb_raw) != 3:
            continue
        try:
            rgb = tuple(max(0, min(255, int(value))) for value in rgb_raw)
        except Exception:
            continue
        colors.append(
            PaletteColor(
                name=name,
                rgb=(rgb[0], rgb[1], rgb[2]),
                hex_value=rgb_to_hex((rgb[0], rgb[1], rgb[2])),
                purpose=str(item.get("purpose", "")),
            )
        )
    if colors:
        return colors
    return palette_from_settings({"palette": DEFAULT_PALETTE})


def read_style_names(path: Path = STYLE_FILE) -> list[str]:
    return [record.name for record in read_style_records(path)]


def read_cover_placeholders(path: Path = COVER_FILE) -> dict[str, bool]:
    if not path.exists():
        return {}
    preview = read_zip_text(path, "Preview/PrvText.txt")
    return {
        "{{문서제목}}": "{{문서제목}}" in preview,
        "{{날짜}}": "{{날짜}}" in preview,
        "{{부서명}}": "{{부서명}}" in preview,
        "<대 외 비>": "대 외 비" in preview,
    }


def parse_cover_input(text: str) -> tuple[str, str, str]:
    values = split_cover_lines(text)
    if len(values) < 3:
        raise ValueError("표지로 만들 제목, 날짜, 부서명 3줄을 선택하세요.")
    return values[0], values[1], "\n".join(values[2:])


def hwp_xml_text(value: str) -> str:
    lines = [escape(line) for line in value.replace("\r\n", "\n").replace("\r", "\n").split("\n")]
    return "</hp:t><hp:lineBreak/><hp:t>".join(lines)


def split_cover_lines(value: str) -> list[str]:
    normalized = value.replace("\r\n", "\n").replace("\r", "\n").replace("\v", "\n").replace("\u2028", "\n")
    lines = [line.strip() for line in normalized.split("\n")]
    return [line for line in lines if line]


def replace_marker_with_paragraphs(section_xml: str, marker: str, value: str) -> str:
    index = section_xml.find(marker)
    if index < 0:
        return section_xml
    para_start = section_xml.rfind("<hp:p ", 0, index)
    para_end = section_xml.find("</hp:p>", index)
    if para_start < 0 or para_end < 0:
        return section_xml.replace(marker, hwp_xml_text(value))
    para_end += len("</hp:p>")
    paragraph = section_xml[para_start:para_end]
    lines = split_cover_lines(value)
    if not lines:
        lines = [""]
    replacement = "".join(paragraph.replace(marker, escape(line)) for line in lines)
    return section_xml[:para_start] + replacement + section_xml[para_end:]


def collapse_cover_paragraph_after(section_xml: str, start: int) -> str:
    para_end = section_xml.find("</hp:p>", start)
    if para_end < 0:
        return section_xml
    lineseg_start = section_xml.find("<hp:linesegarray>", start, para_end)
    lineseg_end = section_xml.find("</hp:linesegarray>", lineseg_start, para_end)
    if lineseg_start < 0 or lineseg_end < 0:
        return section_xml
    lineseg_end += len("</hp:linesegarray>")
    collapsed = (
        '<hp:linesegarray><hp:lineseg textpos="0" vertpos="0" vertsize="0" '
        'textheight="0" baseline="0" spacing="0" horzpos="0" horzsize="0" flags="393216"/></hp:linesegarray>'
    )
    return section_xml[:lineseg_start] + collapsed + section_xml[lineseg_end:]


def remove_confidential_cover_table(section_xml: str) -> str:
    needle = "대 외 비"
    index = section_xml.find(needle)
    if index < 0:
        return section_xml
    table_start = section_xml.rfind("<hp:tbl", 0, index)
    table_end = section_xml.find("</hp:tbl>", index)
    if table_start < 0 or table_end < 0:
        return section_xml.replace(needle, "")
    table_end += len("</hp:tbl>")
    section_xml = section_xml[:table_start] + section_xml[table_end:]
    section_xml = section_xml.replace("<hp:t/></hp:run>", "</hp:run>", 1)
    return collapse_cover_paragraph_after(section_xml, table_start)


def create_filled_cover_file(
    title: str,
    date_text: str,
    department: str,
    confidential: bool,
    template: Path = COVER_FILE,
) -> Path:
    if not template.exists():
        raise FileNotFoundError(f"표지 템플릿을 찾지 못했습니다: {template}")

    TEST_OUTPUT_DIR.mkdir(exist_ok=True)
    target = TEST_OUTPUT_DIR / f"cover-{datetime.now().strftime('%Y%m%d-%H%M%S')}.hwpx"
    section_replacements = {
        "{{문서제목}}": hwp_xml_text(title),
        "{{날짜}}": hwp_xml_text(date_text),
    }
    preview_replacements = {
        "{{문서제목}}": title,
        "{{날짜}}": date_text,
        "{{부서명}}": department,
        "대 외 비": "대 외 비" if confidential else "",
    }

    with zipfile.ZipFile(template, "r") as src, zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as dst:
        for info in src.infolist():
            data = src.read(info.filename)
            if info.filename == "Contents/section0.xml":
                text = data.decode("utf-8")
                for old, new in section_replacements.items():
                    text = text.replace(old, new)
                text = replace_marker_with_paragraphs(text, "{{부서명}}", department)
                if not confidential:
                    text = remove_confidential_cover_table(text)
                data = text.encode("utf-8")
            elif info.filename == "Preview/PrvText.txt":
                text = data.decode("utf-8")
                for old, new in preview_replacements.items():
                    text = text.replace(old, new)
                data = text.encode("utf-8")
            dst.writestr(info, data)

    return target


def list_logos(root: Path = LOGO_DIR) -> list[Path]:
    if not root.exists():
        return []
    return sorted(
        (p for p in root.rglob("*") if p.is_file() and p.suffix.lower() in {".png", ".jpg", ".jpeg", ".bmp"}),
        key=lambda p: str(p.relative_to(root)),
    )


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
    return re.sub(r"(?<=\d),(?=\d{3}\b)", "", text)


def strip_wrapping_blank_lines(text: str) -> str:
    return re.sub(r"\A(?:[ \t]*[\r\n]+)+|(?:[\r\n]+[ \t]*)+\Z", "", text)


DATE_RE = re.compile(
    r"(?<!\d)"
    r"(?P<year>\d{4})"
    r"(?:\s*년\s*|\s*\.\s*)"
    r"(?P<month>\d{1,2})"
    r"(?:\s*월\s*|\s*\.\s*)"
    r"(?P<day>\d{1,2})"
    r"(?:\s*일|\s*\.)?"
    r"(?!\d)"
)


WEEKDAY_DATE_RE = re.compile(
    r"(?<!\d)"
    r"(?P<date>"
    r"(?P<year>\d{4})"
    r"(?:\s*년\s*|\s*\.\s*)"
    r"(?P<month>\d{1,2})"
    r"(?:\s*월\s*|\s*\.\s*)"
    r"(?P<day>\d{1,2})"
    r"(?:\s*일|\s*\.)?"
    r")"
    r"(?:\s*\([월화수목금토일]\))?"
    r"(?!\d)"
)

KOREAN_WEEKDAYS = ("월", "화", "수", "목", "금", "토", "일")


def normalize_dates(text: str, style: str) -> str:
    def repl(match: re.Match[str]) -> str:
        year = int(match.group("year"))
        month = int(match.group("month"))
        day = int(match.group("day"))
        if not (1 <= month <= 12 and 1 <= day <= 31):
            return match.group(0)
        if style == "korean":
            return f"{year}년 {month}월 {day}일"
        if style == "dot":
            return f"{year}. {month}. {day}."
        if style == "dot_padded":
            return f"{year}. {month:02d}. {day:02d}."
        return match.group(0)

    return DATE_RE.sub(repl, text)


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
    return all(DATE_RE.search(value) for value in values)


def parse_cell_clipboard_matrix(text: str) -> list[list[str]]:
    normalized = text.replace("\r\n", "\n").replace("\r", "\n")
    normalized = normalized.rstrip("\n")
    if not normalized:
        return []
    return [row.split("\t") for row in normalized.split("\n")]


def looks_like_cell_clipboard_matrix(text: str) -> bool:
    rows = parse_cell_clipboard_matrix(text)
    if not rows:
        return False
    if any(len(row) > 1 for row in rows):
        return True
    return looks_like_multiline_number_block(text) or looks_like_multiline_date_block(text)


def clipboard_cell_values(text: str, *, preserve_tabs: bool = False) -> list[str]:
    if preserve_tabs:
        rows = parse_cell_clipboard_matrix(text)
        return [cell.strip() for row in rows for cell in row if cell.strip()]
    normalized = text.replace("\r\n", "\n").replace("\r", "\n")
    return [value.strip() for value in normalized.split("\n") if value.strip()]


def parse_markdown_table(text: str) -> list[list[str]] | None:
    lines = [line.strip() for line in text.replace("\r\n", "\n").replace("\r", "\n").split("\n")]
    lines = [line for line in lines if line]
    if len(lines) < 2:
        return None

    def split_row(line: str) -> list[str]:
        if line.startswith("|"):
            line = line[1:]
        if line.endswith("|"):
            line = line[:-1]
        return [cell.replace(r"\|", "|").strip() for cell in re.split(r"(?<!\\)\|", line)]

    def is_separator(cells: list[str]) -> bool:
        return bool(cells) and all(re.fullmatch(r":?-{3,}:?", cell.strip()) for cell in cells)

    start = None
    for index in range(len(lines) - 1):
        header = split_row(lines[index])
        separator = split_row(lines[index + 1])
        if len(header) >= 2 and len(separator) == len(header) and is_separator(separator):
            start = index
            break
    if start is None:
        return None

    rows: list[list[str]] = [split_row(lines[start])]
    width = len(rows[0])
    for line in lines[start + 2 :]:
        if "|" not in line:
            break
        row = split_row(line)
        if len(row) < width:
            row.extend([""] * (width - len(row)))
        rows.append(row[:width])
    return rows if len(rows) >= 2 else None


def markdown_rows_to_html_table(rows: list[list[str]]) -> str:
    header, body = rows[0], rows[1:]
    parts = [
        '<table border="1" cellspacing="0" cellpadding="4" '
        'style="border-collapse:collapse;">',
        "<thead><tr>",
    ]
    parts.extend(f"<th>{escape(cell)}</th>" for cell in header)
    parts.append("</tr></thead><tbody>")
    for row in body:
        parts.append("<tr>")
        parts.extend(f"<td>{escape(cell)}</td>" for cell in row)
        parts.append("</tr>")
    parts.append("</tbody></table>")
    return "".join(parts)


def rows_to_tsv(rows: list[list[str]]) -> str:
    return "\r\n".join("\t".join(cell for cell in row) for row in rows)


def build_cf_html(fragment: str) -> bytes:
    html = f"<html><body><!--StartFragment-->{fragment}<!--EndFragment--></body></html>"
    header_template = (
        "Version:0.9\r\n"
        "StartHTML:{start_html:010d}\r\n"
        "EndHTML:{end_html:010d}\r\n"
        "StartFragment:{start_fragment:010d}\r\n"
        "EndFragment:{end_fragment:010d}\r\n"
    )
    placeholder = header_template.format(start_html=0, end_html=0, start_fragment=0, end_fragment=0)
    prefix = placeholder.encode("utf-8")
    html_bytes = html.encode("utf-8")
    start_html = len(prefix)
    end_html = start_html + len(html_bytes)
    start_fragment = start_html + html.encode("utf-8").index(b"<!--StartFragment-->") + len(b"<!--StartFragment-->")
    end_fragment = start_html + html.encode("utf-8").index(b"<!--EndFragment-->")
    header = header_template.format(
        start_html=start_html,
        end_html=end_html,
        start_fragment=start_fragment,
        end_fragment=end_fragment,
    ).encode("utf-8")
    return header + html_bytes


def clean_manual_line_breaks(text: str) -> str:
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"[ \t]+\n", "\n", text)
    text = re.sub(r"\n[ \t]+", "\n", text)
    lines = text.split("\n")
    result: list[str] = []

    def should_keep_break(previous: str, current: str) -> bool:
        if not previous.strip() or not current.strip():
            return True
        if re.match(r"^\s*([①-⑳•ㆍ\-–—*]|\(?\d+[\).]|[가-힣]\)|[A-Za-z]\))", current):
            return True
        if re.search(r"[:：]$", previous.strip()):
            return True
        return False

    for raw_line in lines:
        line = re.sub(r"[ \t]+", " ", raw_line.strip())
        if not result:
            result.append(line)
            continue
        if should_keep_break(result[-1], line):
            result.append(line)
        else:
            separator = "" if re.search(r"[가-힣A-Za-z0-9]$", result[-1]) and re.match(r"^[,.;:!?)]", line) else " "
            result[-1] = (result[-1].rstrip() + separator + line.lstrip()).strip()

    cleaned = "\n".join(result)
    cleaned = re.sub(r"[ \t]{2,}", " ", cleaned)
    cleaned = re.sub(r"\n{3,}", "\n\n", cleaned)
    return cleaned


HORIZONTAL_SPACE = r"[^\S\r\n\u2028\u2029]"
INVISIBLE_PREFIX = r"[\ufeff\u200b\u200c\u200d]*"
OUTLINE_MARK = r"(?:[∙•·ㆍ◦○●■□▪▫▶▷▸▹►▻◆◇※*+\-–—]+|[^A-Za-z0-9가-힣ㄱ-ㅎㅏ-ㅣ\s])"
OUTLINE_PREFIX_RE = re.compile(rf"(?m)^{INVISIBLE_PREFIX}{HORIZONTAL_SPACE}*{OUTLINE_MARK}+{HORIZONTAL_SPACE}*")


def remove_manual_outline_prefix(text: str) -> str:
    # 수동으로 넣은 들여쓰기 공백 + 글머리 특수문자를 걷어낸다.
    # 예: "   ◦ 본문", "  -본문", "    >> 본문" -> "본문"
    return OUTLINE_PREFIX_RE.sub("", text, count=1)


def remove_manual_outline_prefixes(text: str) -> str:
    return OUTLINE_PREFIX_RE.sub("", text)


def connect_hwp():
    global LAST_HWP_CONNECTION_LOG
    LAST_HWP_CONNECTION_LOG = []
    if win32com is None or pythoncom is None:
        raise RuntimeError("pywin32(win32com)가 설치되어 있지 않습니다.")

    candidates = list_hwp_candidates()
    if candidates:
        candidates.sort(key=lambda item: item.score, reverse=True)
        for item in candidates[:8]:
            LAST_HWP_CONNECTION_LOG.append(
                "[hwp-candidate] "
                f"score={item.score}, visible={item.visible}, docs={item.documents}, "
                f"windows={item.windows}, path={item.path or '(없음)'}, rot={item.display_name}"
            )
        chosen = candidates[0]
        LAST_HWP_CONNECTION_LOG.append(
            f"[hwp-selected] {chosen.display_name} / visible={chosen.visible} / path={chosen.path or '(없음)'}"
        )
        return chosen.hwp

    try:
        hwp = win32com.client.GetActiveObject("HWPFrame.HwpObject")
        LAST_HWP_CONNECTION_LOG.append("[hwp-fallback] GetActiveObject 사용")
        return hwp
    except Exception:
        hwp = win32com.client.Dispatch("HWPFrame.HwpObject")
        LAST_HWP_CONNECTION_LOG.append("[hwp-fallback] 새 HwpObject Dispatch")
        return hwp


def safe_str(value) -> str:
    try:
        return str(value)
    except Exception:
        return ""


def get_foreground_title() -> str:
    if win32gui is None:
        return ""
    try:
        hwnd = win32gui.GetForegroundWindow()
        return win32gui.GetWindowText(hwnd) or ""
    except Exception:
        return ""


def list_visible_window_titles() -> list[tuple[int, str]]:
    if win32gui is None:
        return []

    windows: list[tuple[int, str]] = []

    def collect(hwnd, _extra):
        try:
            if not win32gui.IsWindowVisible(hwnd):
                return
            title = win32gui.GetWindowText(hwnd) or ""
            if title:
                windows.append((hwnd, title))
        except Exception:
            return

    try:
        win32gui.EnumWindows(collect, None)
    except Exception:
        return []
    return windows


def list_hwp_candidates() -> list[HwpCandidate]:
    if pythoncom is None or win32com is None:
        return []

    foreground_title = get_foreground_title()
    candidates: list[HwpCandidate] = []
    ctx = pythoncom.CreateBindCtx(0)
    rot = pythoncom.GetRunningObjectTable()
    for moniker in rot.EnumRunning():
        try:
            display_name = moniker.GetDisplayName(ctx, None)
        except Exception:
            continue
        if "HwpObject" not in display_name:
            continue

        try:
            raw = rot.GetObject(moniker)
            dispatch = raw.QueryInterface(pythoncom.IID_IDispatch)
            hwp = win32com.client.Dispatch(dispatch)
        except Exception:
            continue

        path = ""
        visible = False
        documents = ""
        windows = ""
        try:
            path = safe_str(hwp.Path)
        except Exception:
            path = ""
        try:
            documents = safe_str(hwp.XHwpDocuments.Count)
        except Exception as exc:
            documents = f"err:{type(exc).__name__}"
        try:
            windows_count = hwp.XHwpWindows.Count
            windows = safe_str(windows_count)
            for index in range(int(windows_count)):
                try:
                    if bool(hwp.XHwpWindows.Item(index).Visible):
                        visible = True
                        break
                except Exception:
                    continue
        except Exception as exc:
            windows = f"err:{type(exc).__name__}"

        score = 0
        if visible:
            score += 100
        if path:
            score += 30
            file_name = Path(path).name
            stem = Path(path).stem
            if file_name and file_name in foreground_title:
                score += 120
            elif stem and stem in foreground_title:
                score += 100
        if foreground_title and foreground_title.endswith("- 한글") and visible:
            score += 10
        if documents == "1":
            score += 5

        candidates.append(
            HwpCandidate(
                display_name=display_name,
                hwp=hwp,
                score=score,
                visible=visible,
                path=path,
                documents=documents,
                windows=windows,
            )
        )
    return candidates


def describe_hwp(hwp) -> list[str]:
    lines: list[str] = []
    try:
        lines.append(f"Version: {hwp.Version}")
    except Exception as exc:
        lines.append(f"Version 확인 실패: {exc}")
    try:
        lines.append(f"Documents: {hwp.XHwpDocuments.Count}")
    except Exception as exc:
        lines.append(f"Documents 확인 실패: {exc}")
    try:
        lines.append(f"Windows: {hwp.XHwpWindows.Count}")
    except Exception as exc:
        lines.append(f"Windows 확인 실패: {exc}")
    try:
        lines.append(f"Path: {hwp.Path}")
    except Exception as exc:
        lines.append(f"Path 확인 실패: {exc}")
    try:
        visible_values = []
        for index in range(int(hwp.XHwpWindows.Count)):
            visible_values.append(str(bool(hwp.XHwpWindows.Item(index).Visible)))
        lines.append(f"WindowVisible: {', '.join(visible_values)}")
    except Exception as exc:
        lines.append(f"WindowVisible 확인 실패: {exc}")
    try:
        key = hwp.KeyIndicator()
        lines.append(f"KeyIndicator: {key}")
    except Exception as exc:
        lines.append(f"KeyIndicator 확인 실패: {exc}")
    try:
        pos = hwp.GetPos()
        lines.append(f"CursorPos: {pos}")
    except Exception as exc:
        lines.append(f"CursorPos 확인 실패: {exc}")
    return lines


class MvpApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("한글 스타일 자동화 MVP")
        self.geometry("440x1000")
        self.minsize(420, 720)
        self.hwp = None
        self.style_records: list[StyleRecord] = []
        self.current_doc_style_path: Path | None = None
        self.current_doc_style_map: dict[str, StyleRecord] = {}
        self.current_doc_style_norm_map: dict[str, StyleRecord] = {}
        self.table_settings = load_table_settings()
        self.palette = palette_from_settings(self.table_settings)
        self.apply_on_select = tk.BooleanVar(value=True)
        self.cover_confidential = tk.BooleanVar(value=False)
        self._refreshing = False

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=6, pady=(6, 4))

        log_frame = ttk.LabelFrame(self, text="디버그 로그", padding=5)
        log_frame.pack(fill="both", expand=False, padx=6, pady=(0, 6))
        self.debug_text = tk.Text(log_frame, height=7, wrap="word")
        self.debug_text.pack(side="left", fill="both", expand=True)
        log_scroll = ttk.Scrollbar(log_frame, orient="vertical", command=self.debug_text.yview)
        log_scroll.pack(side="right", fill="y")
        self.debug_text.configure(yscrollcommand=log_scroll.set)

        self._build_status_tab()
        self._build_styles_tab()
        self._build_table_tab()
        self._build_cover_logo_tab()
        self.bind_all("<Control-z>", self.on_global_undo)
        self.bind_all("<Control-Z>", self.on_global_undo)
        self.refresh_all()

    def _build_status_tab(self) -> None:
        frame = ttk.Frame(self.notebook, padding=8)
        self.notebook.add(frame, text="상태")

        self.status_text = tk.Text(frame, height=30, wrap="word")
        self.status_text.pack(fill="both", expand=True)

        buttons = ttk.Frame(frame)
        buttons.pack(fill="x", pady=(8, 0))
        ttk.Button(buttons, text="한글 연결 확인", command=self.check_hwp).pack(side="left")
        ttk.Button(buttons, text="자산 다시 읽기", command=self.refresh_all).pack(side="left", padx=(8, 0))
        ttk.Button(buttons, text="실행취소", command=self.undo_hwp).pack(side="left", padx=(8, 0))

    def _build_styles_tab(self) -> None:
        frame = ttk.Frame(self.notebook, padding=8)
        self.notebook.add(frame, text="스타일/정리")

        style_group = ttk.LabelFrame(frame, text="스타일 적용", padding=8)
        style_group.pack(fill="both", expand=True)

        ttk.Label(style_group, text="bufs/보고서 본문 서식.hwpx에서 읽은 스타일 이름").pack(anchor="w")
        self.style_list = tk.Listbox(style_group, height=25)
        self.style_list.pack(fill="both", expand=True, pady=(8, 0))
        self.style_list.bind("<<ListboxSelect>>", self.on_style_selected)
        self.style_list.bind("<Double-Button-1>", lambda _event: self.apply_selected_style(reason="double-click"))

        ttk.Label(
            style_group,
            text="스타일을 클릭하면 이름 기준으로 현재 문서의 같은 스타일을 찾아 즉시 적용합니다.",
            wraplength=390,
        ).pack(anchor="w", pady=(8, 0))

        row1 = ttk.Frame(style_group)
        row1.pack(fill="x", pady=(8, 0))
        ttk.Checkbutton(row1, text="클릭 즉시 적용", variable=self.apply_on_select).pack(side="left")
        ttk.Button(row1, text="다시 적용", command=lambda: self.apply_selected_style(reason="button")).pack(
            side="left", padx=(8, 0)
        )
        ttk.Button(row1, text="연결 확인", command=self.check_hwp).pack(side="left", padx=(8, 0))

        row2 = ttk.Frame(style_group)
        row2.pack(fill="x", pady=(5, 0))
        ttk.Button(row2, text="대상 확인", command=self.show_current_target).pack(side="left")
        ttk.Button(row2, text="테스트 문서 열기", command=self.open_style_test_copy).pack(side="left", padx=(8, 0))

        row3 = ttk.Frame(style_group)
        row3.pack(fill="x", pady=(5, 0))
        ttk.Button(row3, text="위로", command=lambda: self.move_selected_style(-1)).pack(side="left")
        ttk.Button(row3, text="아래로", command=lambda: self.move_selected_style(1)).pack(side="left", padx=(8, 0))
        ttk.Button(row3, text="기본순서", command=self.reset_style_order).pack(side="left", padx=(8, 0))

        self.build_cleanup_controls(frame)

    def _build_table_tab(self) -> None:
        frame = ttk.Frame(self.notebook, padding=8)
        self.table_frame = frame
        self.notebook.add(frame, text="표")
        self.build_table_tab_content()

    def build_table_tab_content(self) -> None:
        frame = self.table_frame
        for child in frame.winfo_children():
            child.destroy()

        ttk.Button(frame, text="제목셀 자동화", command=self.apply_title_cell_preset).pack(fill="x", pady=(0, 12))
        table_action_row = ttk.Frame(frame)
        table_action_row.pack(fill="x", pady=(0, 12))
        ttk.Button(table_action_row, text="표 설정", command=self.open_table_settings_window).pack(
            side="left", fill="x", expand=True
        )
        ttk.Button(table_action_row, text="셀 속성 복사/적용", command=self.apply_cell_shape_copy_paste).pack(
            side="left", fill="x", expand=True, padx=(6, 0)
        )
        ttk.Button(table_action_row, text="표 바깥 여백 적용", command=self.apply_table_outside_margins).pack(
            side="left", fill="x", expand=True, padx=(6, 0)
        )
        ttk.Button(frame, text="Markdown 표 → 한글 표", command=self.convert_selected_markdown_table).pack(
            fill="x", pady=(0, 12)
        )

        ttk.Label(frame, text="셀 배경색").pack(anchor="w")
        bg_row = ttk.Frame(frame)
        bg_row.pack(fill="x", pady=(6, 12))
        self.create_color_chips(bg_row, self.apply_cell_fill_color)

        ttk.Label(frame, text="글자색").pack(anchor="w")
        text_row = ttk.Frame(frame)
        text_row.pack(fill="x", pady=(6, 12))
        self.create_color_chips(text_row, self.apply_text_color)

        ttk.Label(frame, text="표 스타일").pack(anchor="w")
        table_style_row1 = ttk.Frame(frame)
        table_style_row1.pack(fill="x", pady=(6, 0))
        ttk.Button(table_style_row1, text="얇은전체", command=lambda: self.apply_table_border_preset("thin_all")).pack(
            side="left", fill="x", expand=True
        )
        ttk.Button(table_style_row1, text="굵은외곽", command=lambda: self.apply_table_border_preset("thick_outer")).pack(
            side="left", fill="x", expand=True, padx=(6, 0)
        )

        table_style_row2 = ttk.Frame(frame)
        table_style_row2.pack(fill="x", pady=(6, 12))
        ttk.Button(table_style_row2, text="얇은상하", command=lambda: self.apply_table_border_preset("thin_top_bottom")).pack(
            side="left", fill="x", expand=True
        )
        ttk.Button(table_style_row2, text="굵은상하", command=lambda: self.apply_table_border_preset("thick_top_bottom")).pack(
            side="left", fill="x", expand=True, padx=(6, 0)
        )

        ttk.Label(frame, text="테두리").pack(anchor="w")
        border_row1 = ttk.Frame(frame)
        border_row1.pack(fill="x", pady=(6, 0))
        ttk.Button(border_row1, text="이중 아래", command=lambda: self.apply_table_border_preset("double_bottom")).pack(
            side="left", fill="x", expand=True
        )
        ttk.Button(border_row1, text="좌우 없음", command=lambda: self.apply_table_border_preset("no_left_right")).pack(
            side="left", fill="x", expand=True, padx=(6, 0)
        )

        border_row3 = ttk.Frame(frame)
        border_row3.pack(fill="x", pady=(0, 0))
        ttk.Button(border_row3, text="점선 가로", command=lambda: self.apply_table_border_preset("dotted_inside_horz")).pack(
            side="left", fill="x", expand=True
        )
        ttk.Button(border_row3, text="점선 세로", command=lambda: self.apply_table_border_preset("dotted_inside_vert")).pack(
            side="left", fill="x", expand=True, padx=(6, 0)
        )

        border_row4 = ttk.Frame(frame)
        border_row4.pack(fill="x", pady=(6, 12))
        ttk.Button(border_row4, text="얇은실선 가로", command=lambda: self.apply_table_border_preset("thin_inside_horz")).pack(
            side="left", fill="x", expand=True
        )
        ttk.Button(border_row4, text="얇은실선 세로", command=lambda: self.apply_table_border_preset("thin_inside_vert")).pack(
            side="left", fill="x", expand=True, padx=(6, 0)
        )

        ttk.Label(frame, text="색상 팔레트").pack(anchor="w")
        palette = ttk.Treeview(frame, columns=("rgb", "hex", "purpose"), show="headings", height=8)
        palette.heading("rgb", text="RGB")
        palette.heading("hex", text="HEX")
        palette.heading("purpose", text="기본 용도")
        palette.pack(fill="x", pady=(8, 0))
        for color in self.palette:
            palette.insert("", "end", text=color.name, values=(color.rgb, color.hex_value, color.purpose))

    def create_color_chips(self, parent: ttk.Frame, command) -> None:
        for index, color in enumerate(self.palette):
            button = tk.Button(
                parent,
                text=color.name,
                bg=color.hex_value,
                fg=self.contrast_text_color(color.rgb),
                width=8,
                relief="raised",
                command=lambda item=color: command(item),
            )
            button.grid(row=index // 3, column=index % 3, padx=(0, 6), pady=(0, 6), sticky="ew")
        for col in range(3):
            parent.columnconfigure(col, weight=1)

    def contrast_text_color(self, rgb: tuple[int, int, int]) -> str:
        red, green, blue = rgb
        brightness = (red * 299 + green * 587 + blue * 114) / 1000
        return "#000000" if brightness >= 140 else "#FFFFFF"

    def open_table_settings_window(self) -> None:
        window = tk.Toplevel(self)
        window.title("표 설정")
        window.geometry("420x720")
        window.transient(self)

        settings = merge_dict(DEFAULT_TABLE_SETTINGS, self.table_settings)
        color_names = [color.name for color in self.palette]
        para_styles = [record.name for record in self.style_records if record.style_type == "PARA"]
        char_styles = ["(적용 안 함)", *[record.name for record in self.style_records if record.style_type == "CHAR"]]
        line_type_values = ["solid", "dot", "double", "none"]
        width_values = ["0.12mm", "0.3mm", "0.7mm"]
        preset_keys = list(settings["line_presets"].keys())
        if settings["line_presets"].get("none", {}).get("label") == "적용 안 함":
            settings["line_presets"]["none"]["label"] = "선 없음"
        preset_labels = {key: settings["line_presets"][key]["label"] for key in preset_keys}
        preset_choices = [f"{key}: {preset_labels[key]}" for key in preset_keys]

        vars_by_path: dict[tuple[str, ...], tk.StringVar] = {}
        color_combos: list[ttk.Combobox] = []

        def var_for(path: tuple[str, ...], value: str) -> tk.StringVar:
            var = tk.StringVar(value=value)
            vars_by_path[path] = var
            return var

        def combo(parent, path: tuple[str, ...], value: str, values: list[str], width: int = 16):
            box = ttk.Combobox(parent, textvariable=var_for(path, value), values=values, state="readonly", width=width)
            box.pack(side="left", padx=(4, 0))
            if path[-1] in {"color", "background_color", "text_color"}:
                color_combos.append(box)
            return box

        def entry(parent, path: tuple[str, ...], value: str, width: int = 8):
            box = ttk.Entry(parent, textvariable=var_for(path, value), width=width)
            box.pack(side="left", padx=(4, 0))
            return box

        canvas = tk.Canvas(window, highlightthickness=0)
        scrollbar = ttk.Scrollbar(window, orient="vertical", command=canvas.yview)
        body = ttk.Frame(canvas, padding=10)
        body.bind("<Configure>", lambda _event: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=body, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        ttk.Label(body, text="색상 팔레트").pack(anchor="w")
        palette_tree = ttk.Treeview(body, columns=("rgb", "purpose"), show="tree headings", height=6)
        palette_tree.heading("#0", text="이름")
        palette_tree.heading("rgb", text="RGB")
        palette_tree.heading("purpose", text="용도")
        palette_tree.column("#0", width=82, stretch=False)
        palette_tree.column("rgb", width=92, stretch=False)
        palette_tree.pack(fill="x", pady=(6, 6))
        for color in self.palette:
            palette_tree.insert("", "end", text=color.name, values=(".".join(str(v) for v in color.rgb), color.purpose))

        color_form = ttk.Frame(body)
        color_form.pack(fill="x", pady=(0, 6))
        name_var = tk.StringVar()
        r_var = tk.StringVar(value="0")
        g_var = tk.StringVar(value="0")
        b_var = tk.StringVar(value="0")
        purpose_var = tk.StringVar()
        ttk.Entry(color_form, textvariable=name_var, width=8).pack(side="left")
        ttk.Entry(color_form, textvariable=r_var, width=4).pack(side="left", padx=(4, 0))
        ttk.Entry(color_form, textvariable=g_var, width=4).pack(side="left", padx=(4, 0))
        ttk.Entry(color_form, textvariable=b_var, width=4).pack(side="left", padx=(4, 0))
        ttk.Entry(color_form, textvariable=purpose_var, width=14).pack(side="left", padx=(4, 0))

        def update_color_combo_values() -> None:
            current_names = [palette_tree.item(item, "text") for item in palette_tree.get_children()]
            for box in color_combos:
                box.configure(values=current_names)
                if box.get() not in current_names and current_names:
                    box.set(current_names[0])

        def add_color() -> None:
            name = name_var.get().strip()
            if not name:
                messagebox.showwarning("색상 이름 필요", "색상 이름을 입력하세요.", parent=window)
                return
            if name in [palette_tree.item(item, "text") for item in palette_tree.get_children()]:
                messagebox.showwarning("색상 이름 중복", "이미 있는 색상 이름입니다.", parent=window)
                return
            try:
                rgb = [max(0, min(255, int(v.get()))) for v in (r_var, g_var, b_var)]
            except Exception:
                messagebox.showwarning("RGB 확인", "RGB는 0부터 255 사이 숫자로 입력하세요.", parent=window)
                return
            palette_tree.insert("", "end", text=name, values=(".".join(str(v) for v in rgb), purpose_var.get()))
            name_var.set("")
            purpose_var.set("")
            update_color_combo_values()

        def remove_color() -> None:
            selection = palette_tree.selection()
            if not selection:
                return
            if len(palette_tree.get_children()) <= 1:
                messagebox.showwarning("삭제 불가", "색상은 최소 1개가 필요합니다.", parent=window)
                return
            for item in selection:
                palette_tree.delete(item)
            update_color_combo_values()

        color_buttons = ttk.Frame(body)
        color_buttons.pack(fill="x", pady=(0, 10))
        ttk.Button(color_buttons, text="색 추가", command=add_color).pack(side="left")
        ttk.Button(color_buttons, text="선택 삭제", command=remove_color).pack(side="left", padx=(8, 0))

        ttk.Separator(body).pack(fill="x", pady=12)
        ttk.Label(body, text="선 프리셋").pack(anchor="w")
        for key in preset_keys:
            preset = settings["line_presets"][key]
            row = ttk.Frame(body)
            row.pack(fill="x", pady=(6, 0))
            ttk.Label(row, text=preset["label"], width=10).pack(side="left")
            combo(row, ("line_presets", key, "width"), preset["width"], width_values, width=8)
            combo(row, ("line_presets", key, "type"), preset["type"], line_type_values, width=8)
            combo(row, ("line_presets", key, "color"), preset["color"], color_names, width=8)

        ttk.Separator(body).pack(fill="x", pady=12)
        ttk.Label(body, text="제목셀").pack(anchor="w")

        title = settings["title_cell"]
        row = ttk.Frame(body)
        row.pack(fill="x", pady=(6, 0))
        ttk.Label(row, text="배경색", width=12).pack(side="left")
        combo(row, ("title_cell", "background_color"), title["background_color"], color_names)

        row = ttk.Frame(body)
        row.pack(fill="x", pady=(6, 0))
        ttk.Label(row, text="글자색", width=12).pack(side="left")
        combo(row, ("title_cell", "text_color"), title["text_color"], color_names)

        row = ttk.Frame(body)
        row.pack(fill="x", pady=(6, 0))
        ttk.Label(row, text="문단 스타일", width=12).pack(side="left")
        combo(row, ("title_cell", "paragraph_style"), title["paragraph_style"], para_styles)

        row = ttk.Frame(body)
        row.pack(fill="x", pady=(6, 0))
        ttk.Label(row, text="글자 스타일", width=12).pack(side="left")
        combo(row, ("title_cell", "character_style"), title["character_style"], char_styles)

        ttk.Separator(body).pack(fill="x", pady=12)
        ttk.Label(body, text="제목셀 테두리").pack(anchor="w")
        border_labels = {
            "top": "상",
            "bottom": "하",
            "left": "좌",
            "right": "우",
            "inside_horz": "안쪽 가로",
            "inside_vert": "안쪽 세로",
        }
        for key, label in border_labels.items():
            row = ttk.Frame(body)
            row.pack(fill="x", pady=(6, 0))
            ttk.Label(row, text=label, width=12).pack(side="left")
            current_key = title["borders"].get(key, "thin")
            current = f"{current_key}: {preset_labels.get(current_key, current_key)}"
            combo(row, ("title_cell", "borders", key), current, preset_choices)

        ttk.Separator(body).pack(fill="x", pady=12)
        ttk.Label(body, text="표/셀 여백(mm)").pack(anchor="w")
        margins = settings["table_margins"]
        margin_labels = [
            ("cell_left", "셀 안쪽 좌"),
            ("cell_right", "셀 안쪽 우"),
            ("cell_top", "셀 안쪽 상"),
            ("cell_bottom", "셀 안쪽 하"),
            ("outside_left", "표 바깥 좌"),
            ("outside_right", "표 바깥 우"),
            ("outside_top", "표 바깥 상"),
            ("outside_bottom", "표 바깥 하"),
        ]
        for index in range(0, len(margin_labels), 2):
            row = ttk.Frame(body)
            row.pack(fill="x", pady=(6, 0))
            for key, label in margin_labels[index : index + 2]:
                ttk.Label(row, text=label, width=12).pack(side="left")
                entry(row, ("table_margins", key), str(margins.get(key, "0.0")), width=7)

        def set_path(data: dict, path: tuple[str, ...], value: str) -> None:
            target = data
            for key in path[:-1]:
                target = target[key]
            if path[:2] == ("title_cell", "borders"):
                value = value.split(":", 1)[0]
            target[path[-1]] = value

        def save_and_close() -> None:
            new_settings = merge_dict(DEFAULT_TABLE_SETTINGS, settings)
            palette_items = []
            for item in palette_tree.get_children():
                rgb_text, purpose = palette_tree.item(item, "values")
                rgb = [int(value) for value in str(rgb_text).split(".")]
                palette_items.append(
                    {
                        "name": palette_tree.item(item, "text"),
                        "rgb": rgb,
                        "purpose": str(purpose),
                    }
                )
            new_settings["palette"] = palette_items
            valid_color_names = {item["name"] for item in palette_items}
            for path, var in vars_by_path.items():
                value = var.get()
                if path[-1] in {"color", "background_color", "text_color"} and value not in valid_color_names:
                    value = palette_items[0]["name"]
                if path[0] == "table_margins":
                    try:
                        number = float(value)
                    except ValueError:
                        messagebox.showwarning("여백 값 확인", "표/셀 여백은 mm 단위 숫자로 입력하세요.", parent=window)
                        return
                    if number < 0:
                        messagebox.showwarning("여백 값 확인", "표/셀 여백은 0 이상으로 입력하세요.", parent=window)
                        return
                    value = f"{number:g}"
                set_path(new_settings, path, value)
            normalize_table_settings(new_settings)
            self.table_settings = new_settings
            self.palette = palette_from_settings(new_settings)
            save_table_settings(new_settings)
            self.build_table_tab_content()
            self.debug(f"[table-settings] 저장: {TABLE_SETTINGS_FILE}")
            window.destroy()

        buttons = ttk.Frame(body)
        buttons.pack(fill="x", pady=(14, 0))
        ttk.Button(buttons, text="저장", command=save_and_close).pack(side="left")
        ttk.Button(buttons, text="취소", command=window.destroy).pack(side="left", padx=(8, 0))

    def build_cleanup_controls(self, parent: ttk.Frame) -> None:
        cleanup_group = ttk.LabelFrame(parent, text="문장 정리", padding=8)
        cleanup_group.pack(fill="x", pady=(8, 0))

        hwp_buttons1 = ttk.Frame(cleanup_group)
        hwp_buttons1.pack(fill="x", pady=(8, 0))
        ttk.Button(hwp_buttons1, text="줄바꿈/띄어쓰기 정리", command=self.clean_selected_line_breaks).pack(
            side="left", fill="x", expand=True
        )
        ttk.Button(hwp_buttons1, text="개요 기호 제거", command=self.clean_selected_outline_prefixes).pack(
            side="left", fill="x", expand=True, padx=(6, 0)
        )

        number_group = ttk.LabelFrame(parent, text="숫자 정리", padding=8)
        number_group.pack(fill="x", pady=(8, 0))
        number_buttons = ttk.Frame(number_group)
        number_buttons.pack(fill="x")
        ttk.Button(number_buttons, text="천원단위 쉼표 넣기", command=self.add_commas_to_selection).pack(
            side="left", fill="x", expand=True
        )
        ttk.Button(number_buttons, text="쉼표 빼기", command=self.remove_commas_from_selection).pack(
            side="left", fill="x", expand=True, padx=(6, 0)
        )

        date_group = ttk.LabelFrame(parent, text="날짜 정규화", padding=8)
        date_group.pack(fill="x", pady=(8, 0))
        date_buttons = ttk.Frame(date_group)
        date_buttons.pack(fill="x")
        ttk.Button(date_buttons, text="0000년 0월 0일", command=self.normalize_dates_to_korean).pack(
            side="left", fill="x", expand=True
        )
        ttk.Button(date_buttons, text="0000. 0. 0.", command=self.normalize_dates_to_dot).pack(
            side="left", fill="x", expand=True, padx=(6, 0)
        )
        ttk.Button(date_buttons, text="0000. 00. 00.", command=self.normalize_dates_to_dot_padded).pack(
            side="left", fill="x", expand=True, padx=(6, 0)
        )
        date_buttons2 = ttk.Frame(date_group)
        date_buttons2.pack(fill="x", pady=(6, 0))
        ttk.Button(date_buttons2, text="요일 추가 (월)", command=self.add_weekdays_to_selection).pack(
            side="left", fill="x", expand=True
        )

        table_group = ttk.LabelFrame(parent, text="표 정리/되돌리기", padding=8)
        table_group.pack(fill="x", pady=(8, 0))
        table_buttons = ttk.Frame(table_group)
        table_buttons.pack(fill="x")
        ttk.Button(table_buttons, text="엑셀표 정리", command=self.clean_excel_table).pack(
            side="left", fill="x", expand=True
        )
        ttk.Button(table_buttons, text="실행취소", command=self.undo_hwp).pack(
            side="left", fill="x", expand=True, padx=(6, 0)
        )

        ttk.Label(
            parent,
            text="정리 버튼은 한글에서 글자나 표 셀을 선택한 상태로 사용합니다.",
            wraplength=390,
        ).pack(anchor="w", pady=(8, 0))

    def _build_cover_logo_tab(self) -> None:
        frame = ttk.Frame(self.notebook, padding=8)
        self.notebook.add(frame, text="표지/로고")

        cover_group = ttk.LabelFrame(frame, text="표지 자동화", padding=8)
        cover_group.pack(fill="x", pady=(0, 10))
        ttk.Checkbutton(cover_group, text="대외비 표시", variable=self.cover_confidential).pack(anchor="w")
        ttk.Button(cover_group, text="선택 3줄로 표지 만들기", command=self.create_cover_from_selected_lines).pack(
            fill="x", pady=(6, 0)
        )
        ttk.Label(
            cover_group,
            text="한글에서 제목, 날짜, 부서명을 선택한 뒤 실행합니다. 부서명은 Shift+Enter 줄바꿈처럼 여러 줄이어도 됩니다.",
            wraplength=390,
        ).pack(anchor="w", pady=(6, 0))

        ttk.Label(frame, text="표지 자리표시자").pack(anchor="w")
        self.cover_list = tk.Listbox(frame, height=6)
        self.cover_list.pack(fill="x", pady=(8, 12))

        ttk.Label(frame, text="로고 파일").pack(anchor="w")
        self.logo_list = tk.Listbox(frame, height=18)
        self.logo_list.pack(fill="both", expand=True, pady=(8, 0))

    def log(self, lines: Iterable[str] | str) -> None:
        if isinstance(lines, str):
            lines = [lines]
        for line in lines:
            self.status_text.insert("end", line + "\n")
            self.debug_text.insert("end", line + "\n")
        self.status_text.see("end")
        self.debug_text.see("end")

    def debug(self, line: str) -> None:
        self.debug_text.insert("end", line + "\n")
        self.debug_text.see("end")

    def refresh_all(self) -> None:
        self._refreshing = True
        self.status_text.delete("1.0", "end")
        self.debug_text.delete("1.0", "end")
        self.log(f"작업 폴더: {ROOT}")
        self.log(f"스타일 파일: {STYLE_FILE.exists()} / {STYLE_FILE.name}")
        self.log(f"표지 파일: {COVER_FILE.exists()} / {COVER_FILE.name}")

        self.style_records = apply_saved_style_order(read_style_records())
        self.populate_style_list()
        self.log(f"스타일 이름 수: {len(self.style_records)}")

        placeholders = read_cover_placeholders()
        self.cover_list.delete(0, "end")
        for key, exists in placeholders.items():
            self.cover_list.insert("end", f"{key}: {exists}")

        logos = list_logos()
        self.logo_list.delete(0, "end")
        for logo in logos:
            self.logo_list.insert("end", str(logo.relative_to(LOGO_DIR)))
        self.log(f"로고 파일 수: {len(logos)}")
        self._refreshing = False

    def populate_style_list(self, selected_index: int | None = None) -> None:
        self.style_list.delete(0, "end")
        for record in self.style_records:
            self.style_list.insert("end", f"[{record.style_id:03d}] {record.style_type}  {record.name}")
        if selected_index is not None and self.style_records:
            index = max(0, min(selected_index, len(self.style_records) - 1))
            self.style_list.selection_clear(0, "end")
            self.style_list.selection_set(index)
            self.style_list.activate(index)
            self.style_list.see(index)

    def move_selected_style(self, delta: int) -> None:
        selection = self.style_list.curselection()
        if not selection:
            return
        old_index = selection[0]
        new_index = old_index + delta
        if new_index < 0 or new_index >= len(self.style_records):
            return
        self.style_records[old_index], self.style_records[new_index] = (
            self.style_records[new_index],
            self.style_records[old_index],
        )
        save_style_order(self.style_records)
        self._refreshing = True
        self.populate_style_list(new_index)
        self._refreshing = False
        self.debug(f"[style-order] 순서 변경 저장: {self.style_records[new_index].name}")

    def reset_style_order(self) -> None:
        if STYLE_ORDER_FILE.exists():
            STYLE_ORDER_FILE.unlink()
        self.style_records = read_style_records()
        self._refreshing = True
        self.populate_style_list(0)
        self._refreshing = False
        self.debug("[style-order] 기본순서로 복원")

    def ensure_hwp(self) -> bool:
        if self.hwp is not None:
            try:
                _ = self.hwp.Version
                return True
            except Exception as exc:
                self.debug(f"[hwp] 기존 COM 객체 무효, 재연결: {type(exc).__name__}: {exc}")
                self.hwp = None
                self.current_doc_style_path = None
                self.current_doc_style_map = {}
                self.current_doc_style_norm_map = {}
        try:
            self.debug("[hwp] COM 연결 시작")
            self.hwp = connect_hwp()
            self.debug("[hwp] COM 연결 성공")
            for line in LAST_HWP_CONNECTION_LOG:
                self.debug(line)
            return True
        except Exception as exc:
            self.log(f"한글 COM 연결 실패: {type(exc).__name__}: {exc}")
            messagebox.showerror("한글 연결 실패", str(exc))
            return False

    def hwp_rgb(self, color: PaletteColor) -> int:
        red, green, blue = color.rgb
        try:
            return int(self.hwp.RGBColor(red, green, blue))
        except Exception:
            return red | (green << 8) | (blue << 16)

    def palette_color(self, name: str) -> PaletteColor:
        for color in self.palette:
            if color.name == name:
                return color
        raise KeyError(name)

    def execute_first_hwp_action(self, action_names: tuple[str, ...], hset) -> tuple[str, bool]:
        last_action = action_names[-1]
        for action in action_names:
            last_action = action
            try:
                ok = bool(self.hwp.HAction.Execute(action, hset))
            except Exception as exc:
                self.debug(f"[hwp-action] {action} 실패: {type(exc).__name__}: {exc}")
                continue
            self.debug(f"[hwp-action] {action} result={ok}")
            if ok:
                return action, True
        return last_action, False

    def apply_cell_fill_color(self, color: PaletteColor) -> None:
        if not self.ensure_hwp():
            return
        try:
            action, ok, default_ok = self.set_cell_fill_color(color)
            self.log(f"셀 배경색 적용: {color.name} {color.rgb}, action={action}, GetDefault={default_ok}, result={ok}")
            if not ok:
                messagebox.showwarning(
                    "셀 배경색 적용 확인",
                    "한글이 셀 배경색 적용을 실패로 반환했습니다.\n표 안의 셀을 선택한 상태인지 확인하세요.",
                )
        except Exception as exc:
            self.log(f"셀 배경색 적용 실패: {type(exc).__name__}: {exc}")
            messagebox.showerror("셀 배경색 적용 실패", str(exc))

    def set_cell_fill_color(self, color: PaletteColor) -> tuple[str, bool, bool]:
        value = self.hwp_rgb(color)
        pset = self.hwp.HParameterSet.HCellBorderFill
        default_ok = bool(self.hwp.HAction.GetDefault("CellBorderFill", pset.HSet))

        fill_targets = [pset.FillAttr]
        try:
            fill_targets.append(pset.SelCellsBorderFill.FillAttr)
        except Exception:
            pass

        for fill in fill_targets:
            try:
                fill.WindowsBrush = 1
            except Exception:
                pass
            try:
                fill.WinBrushFaceStyle = 1
            except Exception:
                pass
            fill.WinBrushFaceColor = value
            fill.WinBrushHatchColor = value

        action, ok = self.execute_first_hwp_action(("CellBorderFill", "CellBorder"), pset.HSet)
        return action, ok, default_ok

    def apply_text_color(self, color: PaletteColor) -> None:
        if not self.ensure_hwp():
            return
        try:
            ok, default_ok = self.set_text_color(color)
            self.log(f"글자색 적용: {color.name} {color.rgb}, GetDefault={default_ok}, result={ok}")
            if not ok:
                messagebox.showwarning(
                    "글자색 적용 확인",
                    "한글이 글자색 적용을 실패로 반환했습니다. 글자나 셀을 선택한 상태인지 확인하세요.",
                )
        except Exception as exc:
            self.log(f"글자색 적용 실패: {type(exc).__name__}: {exc}")
            messagebox.showerror("글자색 적용 실패", str(exc))

    def set_text_color(self, color: PaletteColor) -> tuple[bool, bool]:
        value = self.hwp_rgb(color)
        pset = self.hwp.HParameterSet.HCharShape
        default_ok = bool(self.hwp.HAction.GetDefault("CharShape", pset.HSet))
        pset.TextColor = value
        ok = bool(self.hwp.HAction.Execute("CharShape", pset.HSet))
        return ok, default_ok

    def configure_style_apply_set(self, pset, style_id: int) -> None:
        pset.Apply = style_id

    def execute_style_record(self, record: StyleRecord) -> bool:
        pset = self.hwp.HParameterSet.HStyle
        default_ok = self.hwp.HAction.GetDefault("Style", pset.HSet)
        before_apply = pset.Apply
        self.debug(f"[style] GetDefault result={default_ok}, before Apply={before_apply}")
        self.configure_style_apply_set(pset, record.style_id)
        return bool(self.hwp.HAction.Execute("Style", pset.HSet))

    def apply_style_by_name(self, style_name: str) -> bool:
        self.refresh_current_doc_style_map()
        record = self.find_current_doc_style_record(style_name)
        if record is None:
            record = next((item for item in self.style_records if normalize_style_name(item.name) == normalize_style_name(style_name)), None)
            if record is None:
                self.debug(f"[title-cell] 스타일 이름 없음: {style_name}")
                return False
            self.debug(f"[title-cell] 현재 문서 스타일 맵 없음, 기준 스타일 ID fallback: {style_name}, id={record.style_id}")
        ok = self.execute_style_record(record)
        self.debug(
            f"[title-cell] 스타일 적용 {style_name}, id={record.style_id}, type={record.style_type}, "
            f"result={ok}"
        )
        return ok

    def line_type_value(self, kind: str) -> int:
        if kind == "none":
            return 0
        candidates = {
            "solid": ("Solid",),
            "dot": ("Dot",),
            "double": ("DoubleSlim", "SlimThick", "ThickSlim"),
        }[kind]
        for candidate in candidates:
            try:
                return int(self.hwp.HwpLineType(candidate))
            except Exception:
                continue
        return 1

    def line_width_value(self, width: str) -> int:
        return int(self.hwp.HwpLineWidth(width))

    def mm_to_hwpunit(self, value: str | float | int) -> int:
        mm = float(value)
        for method_name in ("MiliToHwpUnit", "MilliToHwpUnit"):
            try:
                method = getattr(self.hwp, method_name)
                return int(method(mm))
            except Exception:
                continue
        return int(round(mm * 283.465))

    def set_com_attr(self, target, attr: str, value) -> bool:
        try:
            setattr(target, attr, value)
            return True
        except Exception as attr_exc:
            hset_targets = []
            try:
                hset_targets.append(target.HSet)
            except Exception:
                pass
            if hasattr(target, "SetItem"):
                hset_targets.append(target)
            for hset in hset_targets:
                try:
                    hset.SetItem(attr, value)
                    self.debug(f"[table-border] {attr} HSet.SetItem fallback 성공")
                    return True
                except Exception as hset_exc:
                    self.debug(f"[table-border] {attr} HSet.SetItem 실패: {type(hset_exc).__name__}: {hset_exc}")
            self.debug(f"[table-border] {attr} 설정 건너뜀: {type(attr_exc).__name__}: {attr_exc}")
            return False

    def set_hset_item(self, hset, item: str, value) -> bool:
        try:
            hset.SetItem(item, value)
            return True
        except Exception as exc:
            self.debug(f"[table-border] HSet.{item} 설정 건너뜀: {type(exc).__name__}: {exc}")
            return False

    def set_parameter_item(self, target, item: str, value, log_prefix: str = "param") -> bool:
        try:
            target_hset = target.HSet
        except Exception:
            target_hset = None
        for candidate in (target, target_hset):
            if candidate is None:
                continue
            try:
                candidate.SetItem(item, value)
                return True
            except Exception as exc:
                self.debug(f"[{log_prefix}] {item} SetItem 실패: {type(exc).__name__}: {exc}")
        try:
            setattr(target, item, value)
            return True
        except Exception as exc:
            self.debug(f"[{log_prefix}] {item} 속성 설정 실패: {type(exc).__name__}: {exc}")
        return False

    def line_values_for_preset(self, preset_key: str) -> tuple[int, int, int] | None:
        preset = self.table_settings["line_presets"].get(preset_key)
        if not preset:
            return None
        line_type = self.line_type_value(preset["type"])
        width = self.line_width_value(preset["width"])
        color = self.hwp_rgb(self.palette_color(preset["color"]))
        return line_type, width, color

    def set_border_edge(self, target, edge: str, line_type: int, width: int, color: int) -> None:
        suffix = {"top": "Top", "bottom": "Bottom", "left": "Left", "right": "Right"}[edge]
        self.set_com_attr(target, f"BorderType{suffix}", line_type)
        self.set_com_attr(target, f"BorderWidth{suffix}", width)
        self.set_com_attr(target, f"BorderColor{suffix}", color)
        if edge == "left":
            self.set_com_attr(target, "BorderCorlorLeft", color)

    def set_border_edge_from_preset(self, target, edge: str, preset_key: str) -> None:
        values = self.line_values_for_preset(preset_key)
        if values is None:
            return
        self.set_border_edge(target, edge, *values)

    def set_inside_border_from_preset(self, target, orientation: str, preset_key: str) -> None:
        values = self.line_values_for_preset(preset_key)
        if values is None:
            return
        line_type, width, color = values
        if orientation == "horz":
            self.set_com_attr(target, "TypeHorz", line_type)
            self.set_com_attr(target, "WidthHorz", width)
            self.set_com_attr(target, "ColorHorz", color)
        else:
            self.set_com_attr(target, "TypeVert", line_type)
            self.set_com_attr(target, "WidthVert", width)
            self.set_com_attr(target, "ColorVert", color)

    def border_action_for_preset(self, preset_name: str) -> tuple[str, ...]:
        return ("CellBorderFill", "CellBorder")

    def apply_table_border_preset(self, preset_name: str) -> None:
        if not self.ensure_hwp():
            return
        try:
            action, ok = self.set_table_border_preset(preset_name)
            self.log(f"표 테두리 적용: {preset_name}, action={action}, result={ok}")
            if not ok:
                messagebox.showwarning("표 테두리 적용 확인", "표 셀을 선택한 상태인지 확인하세요.")
        except Exception as exc:
            self.log(f"표 테두리 적용 실패: {type(exc).__name__}: {exc}")
            messagebox.showerror("표 테두리 적용 실패", str(exc))

    def apply_table_cell_margins(self) -> None:
        if not self.ensure_hwp():
            return
        try:
            action, ok = self.set_table_cell_margins()
            margins = self.table_settings["table_margins"]
            self.log(
                "셀 여백 적용: "
                f"cell=좌{margins['cell_left']} 우{margins['cell_right']} "
                f"상{margins['cell_top']} 하{margins['cell_bottom']}mm, "
                f"action={action}, result={ok}"
            )
            if not ok:
                messagebox.showwarning("셀 여백 적용 확인", "표 안에 커서를 두거나 셀을 선택한 상태인지 확인하세요.")
        except Exception as exc:
            self.log(f"셀 여백 적용 실패: {type(exc).__name__}: {exc}")
            messagebox.showerror("셀 여백 적용 실패", str(exc))

    def apply_cell_shape_copy_paste(self) -> None:
        if not self.ensure_hwp():
            return
        try:
            pset = self.hwp.HParameterSet.HShapeCopyPaste
            default_ok = self.hwp.HAction.GetDefault("ShapeCopyPaste", pset.HSet)
            self.debug(f"[shape-copy-paste] GetDefault ShapeCopyPaste={default_ok}")
            self.set_com_attr(pset, "CellAttr", 1)
            action, ok = self.execute_first_hwp_action(("ShapeCopyPaste",), pset.HSet)
            self.log(f"셀 속성 복사/적용: action={action}, CellAttr=1, result={ok}")
            if not ok:
                messagebox.showwarning(
                    "셀 속성 복사/적용 확인",
                    "복사할 셀에 커서를 두거나, 적용할 셀 범위를 선택한 상태인지 확인하세요.",
                )
        except Exception as exc:
            self.log(f"셀 속성 복사/적용 실패: {type(exc).__name__}: {exc}")
            messagebox.showerror("셀 속성 복사/적용 실패", str(exc))

    def apply_table_outside_margins(self) -> None:
        if not self.ensure_hwp():
            return
        try:
            action, ok = self.set_table_outside_margins()
            margins = self.table_settings["table_margins"]
            self.log(
                "표 바깥 여백 적용: "
                f"outside=좌{margins['outside_left']} 우{margins['outside_right']} "
                f"상{margins['outside_top']} 하{margins['outside_bottom']}mm, "
                f"action={action}, result={ok}"
            )
            if not ok:
                messagebox.showwarning("표 바깥 여백 적용 확인", "표 안에 커서를 두거나 표 개체를 선택한 상태인지 확인하세요.")
        except Exception as exc:
            self.log(f"표 바깥 여백 적용 실패: {type(exc).__name__}: {exc}")
            messagebox.showerror("표 바깥 여백 적용 실패", str(exc))

    def select_current_table_object(self) -> bool:
        if self.run_hwp_command("SelectCtrlReverse"):
            time.sleep(0.03)
            return True
        return False

    def has_selected_object(self) -> bool:
        try:
            return self.hwp.CurSelectedCtrl is not None
        except Exception:
            return False

    def get_table_property_set(self):
        selected = False
        if not self.has_selected_object():
            selected = self.select_current_table_object()
        pset = self.hwp.HParameterSet.HShapeObject
        default_ok = self.hwp.HAction.GetDefault("TablePropertyDialog", pset.HSet)
        if not default_ok:
            default_ok = self.hwp.HAction.GetDefault("ShapeObjDialog", pset.HSet)
        self.debug(f"[table-property] GetDefault={default_ok}, SelectCtrlReverse={selected}")
        return pset

    def set_table_cell_margins(self) -> tuple[str, bool]:
        margins = self.table_settings["table_margins"]
        self.activate_hwp_window()
        time.sleep(0.05)

        pset = self.hwp.HParameterSet.HShapeObject
        default_ok = self.hwp.HAction.GetDefault("TablePropertyDialog", pset.HSet)
        self.debug(f"[table-cell-margin] HShapeObject GetDefault TablePropertyDialog={default_ok}")
        if not default_ok:
            return "TablePropertyDialog.GetDefault", False

        self.set_hset_item(pset.HSet, "ShapeType", 3)
        self.set_hset_item(pset.HSet, "ShapeCellSize", 0)

        cell = pset.ShapeTableCell
        values = {
            "MarginLeft": self.mm_to_hwpunit(margins["cell_left"]),
            "MarginRight": self.mm_to_hwpunit(margins["cell_right"]),
            "MarginTop": self.mm_to_hwpunit(margins["cell_top"]),
            "MarginBottom": self.mm_to_hwpunit(margins["cell_bottom"]),
        }
        self.set_parameter_item(cell, "HasMargin", 1, "table-cell-margin")
        for attr, value in values.items():
            applied = self.set_parameter_item(cell, attr, value, "table-cell-margin")
            self.debug(f"[table-cell-margin] ShapeTableCell.{attr}={value}, SetItem={applied}")
        cell_action, cell_ok = self.execute_first_hwp_action(("TablePropertyDialog",), pset.HSet)
        self.debug(f"[table-cell-margin] cell phase {cell_action}={cell_ok}")
        return cell_action, bool(cell_ok)

    def set_table_outside_margins(self) -> tuple[str, bool]:
        pset = self.get_table_property_set()
        margins = self.table_settings["table_margins"]
        values = {
            "OutsideMarginLeft": self.mm_to_hwpunit(margins["outside_left"]),
            "OutsideMarginRight": self.mm_to_hwpunit(margins["outside_right"]),
            "OutsideMarginTop": self.mm_to_hwpunit(margins["outside_top"]),
            "OutsideMarginBottom": self.mm_to_hwpunit(margins["outside_bottom"]),
        }
        for attr, value in values.items():
            applied = self.set_com_attr(pset, attr, value)
            self.debug(f"[table-outside-margin] {attr}={value}, set={applied}")
        return self.execute_first_hwp_action(("TablePropertyDialog", "ShapeObjDialog"), pset.HSet)

    def set_table_border_preset(self, preset_name: str) -> tuple[str, bool]:
        pset = self.hwp.HParameterSet.HCellBorderFill
        self.hwp.HAction.GetDefault("CellBorderFill", pset.HSet)
        self.prepare_border_scope(pset)
        targets = [pset]
        try:
            targets.append(pset.SelCellsBorderFill)
        except Exception:
            pass

        for target in targets:
            if preset_name == "thin_all":
                for edge in ("top", "bottom", "left", "right"):
                    self.set_border_edge_from_preset(target, edge, "thin")
                self.set_inside_border_from_preset(target, "horz", "thin")
                self.set_inside_border_from_preset(target, "vert", "thin")
            elif preset_name == "thick_outer":
                for edge in ("top", "bottom", "left", "right"):
                    self.set_border_edge_from_preset(target, edge, "thick")
                self.set_inside_border_from_preset(target, "horz", "thin")
                self.set_inside_border_from_preset(target, "vert", "thin")
            elif preset_name == "thin_top_bottom":
                self.set_border_edge_from_preset(target, "top", "thin")
                self.set_border_edge_from_preset(target, "bottom", "thin")
                self.set_border_edge_from_preset(target, "left", "none")
                self.set_border_edge_from_preset(target, "right", "none")
                self.set_inside_border_from_preset(target, "horz", "thin")
                self.set_inside_border_from_preset(target, "vert", "thin")
            elif preset_name == "thick_top_bottom":
                self.set_border_edge_from_preset(target, "top", "thick")
                self.set_border_edge_from_preset(target, "bottom", "thick")
                self.set_border_edge_from_preset(target, "left", "none")
                self.set_border_edge_from_preset(target, "right", "none")
                self.set_inside_border_from_preset(target, "horz", "thin")
                self.set_inside_border_from_preset(target, "vert", "thin")
            elif preset_name == "double_bottom":
                self.set_border_edge_from_preset(target, "bottom", "double")
            elif preset_name == "dotted_inside_horz":
                self.set_inside_border_from_preset(target, "horz", "dotted")
            elif preset_name == "dotted_inside_vert":
                self.set_inside_border_from_preset(target, "vert", "dotted")
            elif preset_name == "dotted_inside":
                self.set_inside_border_from_preset(target, "horz", "dotted")
                self.set_inside_border_from_preset(target, "vert", "dotted")
            elif preset_name == "thin_inside_horz":
                self.set_inside_border_from_preset(target, "horz", "thin")
            elif preset_name == "thin_inside_vert":
                self.set_inside_border_from_preset(target, "vert", "thin")
            elif preset_name == "no_left_right":
                self.set_border_edge_from_preset(target, "left", "none")
                self.set_border_edge_from_preset(target, "right", "none")

        return self.execute_first_hwp_action(self.border_action_for_preset(preset_name), pset.HSet)

    def prepare_border_scope(self, pset) -> None:
        # HWP 2026의 HCellBorderFill에는 ApplyBorderToEdge/NoNeighborCell 같은
        # 범위 제어 속성이 노출되지 않는다. 여기서 억지로 설정하면 로그만
        # 오염되고 동작 추적이 어려워져서 범위 제어는 매크로 기록으로 보완한다.
        return

    def border_scope_targets(self, pset) -> list:
        targets = [pset]
        try:
            targets.append(pset.SelCellsBorderFill)
        except Exception:
            pass
        try:
            targets.append(pset.HSet)
        except Exception:
            pass
        return targets

    def set_title_cell_borders(self) -> tuple[str, bool]:
        pset = self.hwp.HParameterSet.HCellBorderFill
        self.hwp.HAction.GetDefault("CellBorderFill", pset.HSet)
        self.prepare_border_scope(pset)
        targets = [pset]
        try:
            targets.append(pset.SelCellsBorderFill)
        except Exception:
            pass

        borders = self.table_settings["title_cell"]["borders"]
        for target in targets:
            self.set_border_edge_from_preset(target, "top", borders["top"])
            self.set_border_edge_from_preset(target, "bottom", borders["bottom"])
            self.set_border_edge_from_preset(target, "left", borders["left"])
            self.set_border_edge_from_preset(target, "right", borders["right"])
            self.set_inside_border_from_preset(target, "horz", borders["inside_horz"])
            self.set_inside_border_from_preset(target, "vert", borders["inside_vert"])

        return self.execute_first_hwp_action(("CellBorder",), pset.HSet)

    def apply_title_cell_preset(self) -> None:
        if not self.ensure_hwp():
            return
        title = self.table_settings["title_cell"]
        results: list[str] = []

        try:
            bg_action, bg_ok, bg_default = self.set_cell_fill_color(self.palette_color(title["background_color"]))
            results.append(f"배경={bg_ok}({bg_action}, default={bg_default})")
        except Exception as exc:
            results.append(f"배경 실패:{type(exc).__name__}")
            self.debug(f"[title-cell] 배경 실패: {exc}")

        try:
            border_action, border_ok = self.set_title_cell_borders()
            results.append(f"테두리={border_ok}({border_action})")
        except Exception as exc:
            results.append(f"테두리 실패:{type(exc).__name__}")
            self.debug(f"[title-cell] 테두리 실패: {exc}")

        try:
            text_ok, text_default = self.set_text_color(self.palette_color(title["text_color"]))
            results.append(f"글자색={text_ok}(default={text_default})")
        except Exception as exc:
            results.append(f"글자색 실패:{type(exc).__name__}")
            self.debug(f"[title-cell] 글자색 실패: {exc}")

        try:
            para_ok = self.apply_style_by_name(title["paragraph_style"])
            results.append(f"문단스타일={para_ok}({title['paragraph_style']})")
        except Exception as exc:
            results.append(f"문단스타일 실패:{type(exc).__name__}")
            self.debug(f"[title-cell] 문단스타일 실패: {exc}")

        char_style = title.get("character_style", "(적용 안 함)")
        if char_style and char_style != "(적용 안 함)":
            try:
                char_ok = self.apply_style_by_name(char_style)
                results.append(f"글자스타일={char_ok}({char_style})")
            except Exception as exc:
                results.append(f"글자스타일 실패:{type(exc).__name__}")
                self.debug(f"[title-cell] 글자스타일 실패: {exc}")

        self.log("제목셀 자동화: " + ", ".join(results))

    def check_hwp(self) -> None:
        try:
            self.hwp = connect_hwp()
            self.log(["한글 COM 연결 성공", *LAST_HWP_CONNECTION_LOG, *describe_hwp(self.hwp)])
        except Exception as exc:
            messagebox.showerror("한글 연결 실패", str(exc))
            self.log(f"한글 COM 연결 실패: {exc}")

    def on_global_undo(self, event=None):
        widget = getattr(event, "widget", None)
        if isinstance(widget, (tk.Text, tk.Entry, ttk.Entry, ttk.Combobox)):
            return None
        self.undo_hwp()
        return "break"

    def undo_hwp(self) -> None:
        if not self.ensure_hwp():
            return
        ok = self.run_hwp_command("Undo")
        self.log(f"한글 실행취소: result={ok}")
        if not ok:
            messagebox.showwarning("실행취소 실패", "한글 실행취소 명령이 실패했습니다.")

    def show_current_target(self) -> None:
        try:
            if self.hwp is None:
                self.hwp = connect_hwp()
            self.refresh_current_doc_style_map()
            self.log(["현재 한글 적용 대상", *LAST_HWP_CONNECTION_LOG, *describe_hwp(self.hwp)])
        except Exception as exc:
            self.log(f"현재 대상 확인 실패: {type(exc).__name__}: {exc}")
            messagebox.showerror("현재 대상 확인 실패", str(exc))

    def open_style_test_copy(self) -> None:
        try:
            if not STYLE_FILE.exists():
                raise FileNotFoundError(STYLE_FILE)
            TEST_OUTPUT_DIR.mkdir(exist_ok=True)
            stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
            target = TEST_OUTPUT_DIR / f"style-test-{stamp}.hwpx"
            shutil.copy2(STYLE_FILE, target)

            if self.hwp is None:
                self.debug("[hwp] COM 연결 시작")
                self.hwp = connect_hwp()
                self.debug(f"[hwp] COM 연결 성공 version={self.hwp.Version}")
                for line in LAST_HWP_CONNECTION_LOG:
                    self.debug(line)
            try:
                self.hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
            except Exception:
                pass
            try:
                self.hwp.RegisterModule("FilePathCheckerDLL", "FilePathCheckerModule")
            except Exception:
                pass
            ok = self.hwp.Open(str(target), "HWPX", "")
            self.log(f"서식 테스트 문서 열기: {target.name}, result={ok}")
            self.current_doc_style_path = None
            self.current_doc_style_map = {}
            self.current_doc_style_norm_map = {}
            self.refresh_current_doc_style_map()
            if ok:
                messagebox.showinfo(
                    "테스트 문서 열림",
                    "서식 파일 복사본을 열었습니다.\n한글 문서에서 문단을 클릭한 뒤 스타일 목록을 클릭해보세요.",
                )
        except Exception as exc:
            self.log(f"서식 테스트 문서 열기 실패: {type(exc).__name__}: {exc}")
            messagebox.showerror("서식 테스트 문서 열기 실패", str(exc))

    def on_style_selected(self, _event=None) -> None:
        if self._refreshing or not self.apply_on_select.get():
            return
        # Tk listbox selection event can fire before the new selection is fully settled.
        self.after_idle(lambda: self.apply_selected_style(reason="single-click"))

    def get_current_hwp_path(self) -> Path | None:
        if self.hwp is None:
            return None
        try:
            raw_path = safe_str(self.hwp.Path)
        except Exception:
            return None
        if not raw_path:
            return None
        return Path(raw_path)

    def refresh_current_doc_style_map(self) -> None:
        path = self.get_current_hwp_path()
        if path is None:
            self.current_doc_style_path = None
            self.current_doc_style_map = {}
            self.current_doc_style_norm_map = {}
            self.debug("[style-map] 현재 문서 경로를 확인하지 못함")
            return

        if path == self.current_doc_style_path and self.current_doc_style_map:
            return

        self.current_doc_style_path = path
        self.current_doc_style_map = {}
        self.current_doc_style_norm_map = {}
        if path.suffix.lower() != ".hwpx":
            self.debug(f"[style-map] 현재 문서가 hwpx가 아니어서 이름 매칭 불가: {path}")
            return
        if not path.exists():
            self.debug(f"[style-map] 현재 문서 파일을 디스크에서 찾지 못함: {path}")
            return

        records = read_style_records(path)
        self.current_doc_style_map = {record.name: record for record in records}
        for record in records:
            key = normalize_style_name(record.name)
            self.current_doc_style_norm_map.setdefault(key, record)
        matched = sum(1 for record in self.style_records if self.find_current_doc_style_record(record.name) is not None)
        self.debug(f"[style-map] 현재 문서 스타일 {len(records)}개 읽음: {path}")
        self.debug(f"[style-map] 기준 서식과 이름 일치 {matched}/{len(self.style_records)}개")

    def find_current_doc_style_record(self, style_name: str) -> StyleRecord | None:
        exact = self.current_doc_style_map.get(style_name)
        if exact is not None:
            return exact
        normalized = normalize_style_name(style_name)
        match = self.current_doc_style_norm_map.get(normalized)
        if match is not None:
            self.debug(f"[style-map] 정규화 이름 매칭: 요청={style_name!r}, 현재문서={match.name!r}")
            return match
        return None

    def apply_selected_style(self, reason: str = "manual") -> None:
        selection = self.style_list.curselection()
        if not selection:
            messagebox.showwarning("스타일 선택 필요", "적용할 스타일을 먼저 선택하세요.")
            return

        record = self.style_records[selection[0]]
        self.debug(f"[style] 요청 reason={reason}, id={record.style_id}, type={record.style_type}, name={record.name}")
        if record.style_type != "PARA":
            self.debug(f"[style] 글자 스타일 적용 시도: type={record.style_type}, name={record.name}")

        try:
            if self.hwp is None:
                self.debug("[hwp] COM 연결 시작")
                self.hwp = connect_hwp()
                self.debug("[hwp] COM 연결 성공")
                for line in LAST_HWP_CONNECTION_LOG:
                    self.debug(line)
                for line in describe_hwp(self.hwp):
                    self.debug(f"[hwp] {line}")
            self.refresh_current_doc_style_map()
            target_record = self.find_current_doc_style_record(record.name)
            if target_record is None:
                current_path = self.get_current_hwp_path()
                target_record = record
                self.debug(
                    f"[style] 현재 문서 스타일 이름 매칭 실패, 기준 스타일 ID fallback: "
                    f"name={record.name}, current_doc={current_path}, id={record.style_id}"
                )

            ok = self.execute_style_record(target_record)
            self.log(
                "스타일 적용 요청: "
                f"name={record.name}, 기준id={record.style_id}, 현재문서id={target_record.style_id}, "
                f"type={target_record.style_type}, result={ok}"
            )
            for line in describe_hwp(self.hwp):
                self.debug(f"[hwp-after] {line}")
            if not ok:
                messagebox.showwarning(
                    "스타일 적용 결과 확인 필요",
                    "한글이 스타일 적용 명령을 실패로 반환했습니다. 현재 문서에 같은 스타일 세트가 있는지 확인하세요.",
                )
        except Exception as exc:
            self.log(f"스타일 적용 실패: {type(exc).__name__}: {exc}")
            messagebox.showerror(
                "스타일 적용 실패",
                f"{exc}\n\n현재 문서에 같은 스타일 세트가 없거나, 선택 영역 상태가 맞지 않을 수 있습니다.",
            )

    def insert_file_at_cursor(self, path: Path) -> bool:
        option = self.hwp.HParameterSet.HInsertFile
        self.hwp.HAction.GetDefault("InsertFile", option.HSet)
        try:
            option.filename = str(path)
        except Exception:
            pass
        try:
            option.Filename = str(path)
        except Exception:
            pass
        option.KeepSection = 0
        option.KeepCharshape = 1
        option.KeepParashape = 1
        option.KeepStyle = 1
        return bool(self.hwp.HAction.Execute("InsertFile", option.HSet))

    def create_cover_from_selected_lines(self) -> None:
        if not self.ensure_hwp():
            return
        original_text = ""
        try:
            copy_ok = self.run_hwp_command("Copy")
            time.sleep(0.08)
            original_text = self.get_clipboard_text()
            if not copy_ok or not original_text:
                messagebox.showwarning("표지 입력 선택 필요", "한글에서 제목, 날짜, 부서명 3줄을 먼저 선택하세요.")
                self.log("표지 만들기: 선택 텍스트 없음")
                return

            title, date_text, department = parse_cover_input(original_text)
            cover_path = create_filled_cover_file(
                title,
                date_text,
                department,
                self.cover_confidential.get(),
            )

            delete_ok = self.run_hwp_command("Delete")
            if not delete_ok:
                messagebox.showwarning("표지 만들기 실패", "선택한 3줄을 삭제하지 못했습니다. 선택 영역을 확인하세요.")
                self.log("표지 만들기: 선택 영역 삭제 실패")
                return

            insert_ok = self.insert_file_at_cursor(cover_path)
            if not insert_ok:
                self.set_clipboard_text(original_text)
                restore_ok = self.run_hwp_command("Paste")
                messagebox.showwarning(
                    "표지 삽입 실패",
                    f"표지 파일은 만들었지만 현재 문서에 삽입하지 못했습니다.\n\n{cover_path}\n원문 복구: {restore_ok}",
                )
                self.log(f"표지 삽입 실패: file={cover_path}, restored={restore_ok}")
                return

            self.activate_hwp_window()
            self.log(
                "표지 만들기 완료: "
                f"제목={title}, 날짜={date_text}, 부서={department}, "
                f"대외비={self.cover_confidential.get()}, file={cover_path.name}"
            )
        except Exception as exc:
            self.log(f"표지 만들기 실패: {type(exc).__name__}: {exc}")
            messagebox.showerror("표지 만들기 실패", str(exc))

    def run_first_hwp_command(self, commands: tuple[str, ...]) -> tuple[str, bool]:
        last = commands[-1]
        for command in commands:
            last = command
            if self.run_hwp_command(command):
                return command, True
        return last, False

    def run_hwp_command(self, command: str) -> bool:
        for runner in (
            lambda: self.hwp.HAction.Run(command),
            lambda: self.hwp.Run(command),
        ):
            try:
                return bool(runner())
            except Exception:
                continue
        return False

    def activate_hwp_window(self) -> bool:
        if win32gui is None:
            return False

        path = self.get_current_hwp_path()
        names: list[str] = []
        if path is not None:
            names.extend([path.name, path.stem])

        candidates: list[tuple[int, str, int]] = []
        for hwnd, title in list_visible_window_titles():
            if "한글" not in title:
                continue
            score = 10
            for name in names:
                if name and name in title:
                    score += 100
                    break
            if title.endswith("- 한글"):
                score += 20
            candidates.append((hwnd, title, score))

        if not candidates:
            return False

        hwnd, title, _score = sorted(candidates, key=lambda item: item[2], reverse=True)[0]
        try:
            if win32con is not None:
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            else:
                win32gui.ShowWindow(hwnd, 9)
        except Exception:
            pass
        try:
            win32gui.SetForegroundWindow(hwnd)
            self.debug(f"[hwp-focus] 한글 창 활성화: {title}")
            return True
        except Exception as exc:
            self.debug(f"[hwp-focus] 한글 창 활성화 실패: {type(exc).__name__}: {exc}")
            return False

    def auto_confirm_dialog(self, title_part: str, timeout_sec: float = 2.0) -> None:
        if win32gui is None or win32con is None:
            return

        def worker() -> None:
            deadline = time.monotonic() + timeout_sec
            while time.monotonic() < deadline:
                try:
                    matches: list[int] = []

                    def collect(hwnd, _extra):
                        try:
                            title = win32gui.GetWindowText(hwnd) or ""
                            if title_part in title and win32gui.IsWindowVisible(hwnd):
                                matches.append(hwnd)
                        except Exception:
                            pass

                    win32gui.EnumWindows(collect, None)
                    if matches:
                        hwnd = matches[0]
                        clicked = False

                        def click_ok(child, _extra):
                            nonlocal clicked
                            try:
                                text = win32gui.GetWindowText(child) or ""
                                if "확인" in text:
                                    win32gui.PostMessage(child, win32con.BM_CLICK, 0, 0)
                                    clicked = True
                            except Exception:
                                pass

                        try:
                            win32gui.EnumChildWindows(hwnd, click_ok, None)
                        except Exception:
                            pass
                        if not clicked:
                            win32gui.PostMessage(hwnd, win32con.WM_COMMAND, 1, 0)
                        self.debug(f"[dialog-auto-confirm] {title_part} 확인 처리")
                        return
                except Exception as exc:
                    self.debug(f"[dialog-auto-confirm] {title_part} 확인 실패: {type(exc).__name__}: {exc}")
                    return
                time.sleep(0.05)

        threading.Thread(target=worker, daemon=True).start()

    def get_clipboard_text(self) -> str:
        if win32clipboard is None:
            raise RuntimeError("win32clipboard를 사용할 수 없습니다.")
        last_exc = None
        for _ in range(5):
            try:
                win32clipboard.OpenClipboard()
                try:
                    if not win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_UNICODETEXT):
                        return ""
                    return win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
                finally:
                    win32clipboard.CloseClipboard()
            except Exception as exc:
                last_exc = exc
                time.sleep(0.05)
        raise RuntimeError(f"클립보드 읽기 실패: {last_exc}")

    def set_clipboard_text(self, text: str) -> None:
        if win32clipboard is None:
            raise RuntimeError("win32clipboard를 사용할 수 없습니다.")
        last_exc = None
        for _ in range(5):
            try:
                win32clipboard.OpenClipboard()
                try:
                    win32clipboard.EmptyClipboard()
                    if hasattr(win32clipboard, "SetClipboardText"):
                        win32clipboard.SetClipboardText(text, win32clipboard.CF_UNICODETEXT)
                    else:
                        win32clipboard.SetClipboardData(win32clipboard.CF_UNICODETEXT, text)
                    return
                finally:
                    win32clipboard.CloseClipboard()
            except Exception as exc:
                last_exc = exc
                time.sleep(0.05)
        raise RuntimeError(f"클립보드 쓰기 실패: {last_exc}")

    def set_clipboard_html_table(self, html_table: str, fallback_text: str) -> None:
        if win32clipboard is None:
            raise RuntimeError("win32clipboard를 사용할 수 없습니다.")
        last_exc = None
        for _ in range(5):
            try:
                win32clipboard.OpenClipboard()
                try:
                    win32clipboard.EmptyClipboard()
                    html_format = win32clipboard.RegisterClipboardFormat("HTML Format")
                    win32clipboard.SetClipboardData(html_format, build_cf_html(html_table))
                    if hasattr(win32clipboard, "SetClipboardText"):
                        win32clipboard.SetClipboardText(fallback_text, win32clipboard.CF_UNICODETEXT)
                    else:
                        win32clipboard.SetClipboardData(win32clipboard.CF_UNICODETEXT, fallback_text)
                    return
                finally:
                    win32clipboard.CloseClipboard()
            except Exception as exc:
                last_exc = exc
                time.sleep(0.05)
        raise RuntimeError(f"HTML 클립보드 쓰기 실패: {last_exc}")

    def clear_hwp_selection(self) -> bool:
        for command in ("Cancel",):
            if self.run_hwp_command(command):
                time.sleep(0.03)
                return True
        return False

    def select_current_table_cell_for_replace(self) -> bool:
        if self.run_hwp_command("TableCellBlock"):
            return True
        self.clear_hwp_selection()
        time.sleep(0.03)
        return self.run_hwp_command("TableCellBlock")

    def move_table_cell(self, command: str, count: int = 1) -> bool:
        for _ in range(count):
            if not self.run_hwp_command(command):
                return False
            time.sleep(0.02)
        return True

    def set_hwp_pos(self, pos: tuple[int, int, int]) -> bool:
        try:
            return bool(self.hwp.SetPos(pos[0], pos[1], pos[2]))
        except Exception:
            return False

    def read_current_cell_text(self) -> str | None:
        if not self.select_current_table_cell_for_replace():
            return None
        if not self.run_hwp_command("Copy"):
            return None
        time.sleep(0.05)
        value = self.get_clipboard_text()
        self.clear_hwp_selection()
        return value.strip()

    def paste_text_into_selected_cell(self) -> bool:
        try:
            option = self.hwp.HParameterSet.HSelectionOpt
            self.hwp.HAction.GetDefault("Paste", option.HSet)
            option.Option = 5
            return bool(self.hwp.HAction.Execute("Paste", option.HSet))
        except Exception as exc:
            self.debug(f"[cell-paste] HSelectionOpt paste failed: {type(exc).__name__}: {exc}")
        return self.run_hwp_command("Paste")

    def paste_with_selection_option(self, option_value: int, log_prefix: str) -> bool:
        try:
            option = self.hwp.HParameterSet.HSelectionOpt
            default_ok = self.hwp.HAction.GetDefault("Paste", option.HSet)
            option.Option = option_value
            ok = bool(self.hwp.HAction.Execute("Paste", option.HSet))
            self.debug(f"[{log_prefix}] HSelectionOpt Option={option_value}, default={default_ok}, result={ok}")
            if ok:
                return True
        except Exception as exc:
            self.debug(f"[{log_prefix}] HSelectionOpt Option={option_value} 실패: {type(exc).__name__}: {exc}")
        fallback_ok = self.run_hwp_command("Paste")
        self.debug(f"[{log_prefix}] Paste fallback={fallback_ok}")
        return fallback_ok

    def paste_html_original_format(self) -> bool:
        try:
            html_set = self.hwp.HParameterSet.HPasteHtml
            html_default = self.hwp.HAction.GetDefault("PasteHtmlDialog", html_set.HSet)
            html_set.Format = 0
            self.auto_confirm_dialog("HTML 문서 붙이기")
            html_ok = bool(self.hwp.HAction.Execute("PasteHtmlDialog", html_set.HSet))
            self.debug(f"[markdown-table] PasteHtmlDialog Format=0, default={html_default}, result={html_ok}")
        except Exception as exc:
            self.debug(f"[markdown-table] PasteHtmlDialog 실패: {type(exc).__name__}: {exc}")
            html_ok = False

        try:
            option = self.hwp.HParameterSet.HSelectionOpt
            paste_default = self.hwp.HAction.GetDefault("Paste", option.HSet)
            paste_ok = bool(self.hwp.HAction.Execute("Paste", option.HSet))
            self.debug(f"[markdown-table] Paste after PasteHtmlDialog, default={paste_default}, result={paste_ok}")
            if paste_ok:
                return True
        except Exception as exc:
            self.debug(f"[markdown-table] Paste after PasteHtmlDialog 실패: {type(exc).__name__}: {exc}")

        fallback_ok = self.run_hwp_command("Paste")
        self.debug(f"[markdown-table] Paste fallback after html_ok={html_ok}: {fallback_ok}")
        return fallback_ok

    def move_from_endpoint_to_top_left(
        self,
        endpoint: tuple[int, int, int],
        rows: int,
        cols: int,
        endpoint_corner: str,
    ) -> bool:
        if not self.set_hwp_pos(endpoint):
            return False
        self.clear_hwp_selection()
        if endpoint_corner == "bottom_right":
            return self.move_table_cell("TableLeftCell", cols - 1) and self.move_table_cell("TableUpperCell", rows - 1)
        if endpoint_corner == "bottom_left":
            return self.move_table_cell("TableUpperCell", rows - 1)
        if endpoint_corner == "top_right":
            return self.move_table_cell("TableLeftCell", cols - 1)
        if endpoint_corner == "top_left":
            return True
        return False

    def read_cell_rect_values(
        self,
        endpoint: tuple[int, int, int],
        rows: int,
        cols: int,
        endpoint_corner: str,
    ) -> tuple[list[str] | None, str]:
        if not self.move_from_endpoint_to_top_left(endpoint, rows, cols, endpoint_corner):
            return None, ""
        try:
            start_label = safe_str(self.hwp.KeyIndicator()[-1])
        except Exception:
            start_label = ""
        values: list[str] = []
        for row_index in range(rows):
            if row_index > 0:
                if not self.move_table_cell("TableLowerCell"):
                    return None, start_label
                if cols > 1 and not self.move_table_cell("TableLeftCell", cols - 1):
                    return None, start_label
            for col_index in range(cols):
                value = self.read_current_cell_text()
                if value is None:
                    return None, start_label
                values.append(value)
                if col_index < cols - 1 and not self.move_table_cell("TableRightCell"):
                    return None, start_label
        return values, start_label

    def detect_selected_cell_rect(
        self,
        endpoint: tuple[int, int, int],
        selected_values: list[str],
    ) -> dict | None:
        total = len(selected_values)
        if total < 2:
            return None
        corners = ("bottom_right", "bottom_left", "top_right", "top_left")
        best: dict | None = None
        for rows in range(1, total + 1):
            if total % rows != 0:
                continue
            cols = total // rows
            for corner in corners:
                values, start_label = self.read_cell_rect_values(endpoint, rows, cols, corner)
                if values is None:
                    continue
                score = sum(1 for left, right in zip(values, selected_values) if left == right)
                if score == total:
                    return {
                        "rows": rows,
                        "cols": cols,
                        "corner": corner,
                        "start_label": start_label,
                    }
                if best is None or score > best["score"]:
                    best = {
                        "score": score,
                        "rows": rows,
                        "cols": cols,
                        "corner": corner,
                        "start_label": start_label,
                    }
        self.debug(f"[cell-detect] best_partial={best}, selected={selected_values}")
        return None

    def transform_detected_cell_rect(
        self,
        label: str,
        endpoint: tuple[int, int, int],
        rect: dict,
        transform,
        *,
        strip_wrapping_lines: bool = False,
    ) -> tuple[int, int] | None:
        rows = int(rect["rows"])
        cols = int(rect["cols"])
        if not self.move_from_endpoint_to_top_left(endpoint, rows, cols, str(rect["corner"])):
            return None
        visited = 0
        changed = 0
        for row_index in range(rows):
            if row_index > 0:
                if not self.move_table_cell("TableLowerCell"):
                    return None
                if cols > 1 and not self.move_table_cell("TableLeftCell", cols - 1):
                    return None
            for col_index in range(cols):
                if not self.select_current_table_cell_for_replace():
                    return None
                if not self.run_hwp_command("Copy"):
                    return None
                time.sleep(0.05)
                current_text = self.get_clipboard_text()
                source_text = strip_wrapping_blank_lines(current_text) if strip_wrapping_lines else current_text
                transformed = transform(source_text) if source_text else source_text
                visited += 1
                if source_text and transformed != source_text:
                    self.set_clipboard_text(transformed)
                    if not self.paste_text_into_selected_cell():
                        return None
                    changed += 1
                self.clear_hwp_selection()
                if col_index < cols - 1 and not self.move_table_cell("TableRightCell"):
                    return None
        return visited, changed

    def transform_selected_cells_by_iteration(
        self,
        label: str,
        text: str,
        transform,
        *,
        strip_wrapping_lines: bool = False,
        preserve_clipboard_cells: bool = False,
    ) -> bool:
        if preserve_clipboard_cells and "\t" not in text:
            return False

        selected_values = clipboard_cell_values(text, preserve_tabs=preserve_clipboard_cells)
        if len(selected_values) < 2:
            return False

        changed_in_selection_preview = 0
        for cell_text in selected_values:
            source_text = strip_wrapping_blank_lines(cell_text) if strip_wrapping_lines else cell_text
            transformed = transform(source_text) if source_text else source_text
            if transformed != source_text:
                changed_in_selection_preview += 1
        if changed_in_selection_preview == 0:
            return False

        try:
            endpoint = self.hwp.GetPos()
        except Exception:
            return False

        if not self.set_hwp_pos(endpoint):
            return False
        self.clear_hwp_selection()
        if not self.move_table_cell("TableUpperCell", len(selected_values) - 1):
            messagebox.showwarning(
                label,
                "한 열 범위를 위에서 아래로 선택한 경우만 처리할 수 있습니다.\n\n"
                "아래에서 위로 선택했거나, 선택한 셀 위쪽으로 충분히 이동할 수 없어 중단했습니다.",
            )
            self.log(f"{label}: 한 열 위->아래 선택 확인 실패, values={len(selected_values)}")
            return True

        start_pos = self.hwp.GetPos()
        actual_values: list[str] = []
        for index in range(len(selected_values)):
            if index > 0 and not self.move_table_cell("TableLowerCell"):
                actual_values = []
                break
            value = self.read_current_cell_text()
            if value is None:
                actual_values = []
                break
            actual_values.append(strip_wrapping_blank_lines(value) if strip_wrapping_lines else value)

        expected_values = [
            strip_wrapping_blank_lines(value) if strip_wrapping_lines else value
            for value in selected_values
        ]
        if actual_values != expected_values:
            self.set_hwp_pos(endpoint)
            messagebox.showwarning(
                label,
                "선택한 범위가 한 열 위→아래 선택으로 확인되지 않았습니다.\n\n"
                "한 열만 위에서 아래로 드래그해 선택한 뒤 다시 실행하세요.",
            )
            self.log(
                f"{label}: 한 열 선택 검증 실패, expected={expected_values}, actual={actual_values}"
            )
            return True

        self.debug(
            f"[cell-column-iterate] {label}: cells={len(selected_values)}, "
            f"preview_changed={changed_in_selection_preview}, start={start_pos}, endpoint={endpoint}"
        )
        if not self.set_hwp_pos(start_pos):
            return False

        visited = 0
        changed = 0
        for index in range(len(selected_values)):
            if index > 0 and not self.move_table_cell("TableLowerCell"):
                break
            if not self.select_current_table_cell_for_replace():
                break
            if not self.run_hwp_command("Copy"):
                break
            time.sleep(0.05)
            current_text = self.get_clipboard_text()
            source_text = strip_wrapping_blank_lines(current_text) if strip_wrapping_lines else current_text
            transformed = transform(source_text) if source_text else source_text
            visited += 1
            if source_text and transformed != source_text:
                self.set_clipboard_text(transformed)
                if not self.paste_text_into_selected_cell():
                    break
                changed += 1
            self.clear_hwp_selection()

        if visited != len(selected_values):
            messagebox.showwarning(
                label,
                "선택한 셀 범위를 순회하던 중 중단했습니다.\n\n"
                "일부 셀이 이미 바뀌었으면 한글에서 Ctrl+Z로 되돌린 뒤, 더 작은 범위로 다시 실행하세요.",
            )
            self.log(f"{label}: 한 열 순회 중단, visited={visited}, expected={len(selected_values)}")
            return True

        self.activate_hwp_window()
        self.log(f"{label}: 한 열 셀 순회 완료, visited={visited}, changed={changed}")
        return True

    def transform_selected_text(
        self,
        label: str,
        transform,
        *,
        block_multiline_number_block: bool = False,
        allow_cell_iteration: bool = False,
        preserve_clipboard_cells: bool = False,
        strip_wrapping_lines: bool = False,
        reselect_current_cell: bool = False,
    ) -> None:
        if not self.ensure_hwp():
            return
        try:
            copy_ok = self.run_hwp_command("Copy")
            time.sleep(0.08)
            text = self.get_clipboard_text()
            if not copy_ok or not text:
                messagebox.showwarning(label, "한글에서 변환할 글자 영역을 먼저 선택하세요.")
                self.log(f"{label}: 선택 텍스트 없음")
                return
            if allow_cell_iteration and self.transform_selected_cells_by_iteration(
                label,
                text,
                transform,
                strip_wrapping_lines=strip_wrapping_lines,
                preserve_clipboard_cells=preserve_clipboard_cells,
            ):
                return
            if preserve_clipboard_cells and "\t" in text:
                messagebox.showwarning(
                    label,
                    "여러 표 셀 선택은 셀별 순회로만 처리합니다.\n\n"
                    "셀별 순회 검증에 실패해 전체 붙여넣기를 중단했습니다.",
                )
                self.log(f"{label}: 셀 범위 전체 붙여넣기 중단")
                return
            if block_multiline_number_block and looks_like_multiline_number_block(text):
                self.log(f"{label}: 여러 셀 숫자 블록으로 보여 자동 붙여넣기 중단")
                messagebox.showwarning(
                    label,
                    "여러 표 셀을 한 번에 선택한 숫자 블록은 한글 붙여넣기에서 셀 내용이 흐트러질 수 있어 중단했습니다.\n\n"
                    "방금처럼 셀이 깨진 경우 한글에서 Ctrl+Z로 되돌린 뒤, 우선 한 셀 안의 텍스트만 선택해서 실행하세요.",
                )
                return
            source_text = strip_wrapping_blank_lines(text) if strip_wrapping_lines else text
            if not source_text:
                self.log(f"{label}: 선택 텍스트가 빈 줄만 포함")
                if reselect_current_cell:
                    self.run_hwp_command("TableCellBlock")
                    self.activate_hwp_window()
                return
            transformed = transform(source_text)
            if transformed == source_text:
                cell_selected = False
                if reselect_current_cell:
                    cell_selected = self.run_hwp_command("TableCellBlock")
                    self.activate_hwp_window()
                self.log(f"{label}: 변경 없음, CellReselect={cell_selected}")
                return
            self.set_clipboard_text(transformed)
            paste_ok = self.run_hwp_command("Paste")
            cell_selected = False
            if paste_ok and reselect_current_cell:
                time.sleep(0.03)
                cell_selected = self.run_hwp_command("TableCellBlock")
            if paste_ok:
                self.activate_hwp_window()
            self.log(
                f"{label}: Copy={copy_ok}, Paste={paste_ok}, "
                f"CellReselect={cell_selected}, {len(text)}자 -> {len(transformed)}자"
            )
            if not paste_ok:
                messagebox.showwarning(label, "변환 텍스트를 클립보드에 넣었지만 한글 붙여넣기가 실패했습니다.")
        except Exception as exc:
            self.log(f"{label} 실패: {type(exc).__name__}: {exc}")
            messagebox.showerror(f"{label} 실패", str(exc))

    def convert_selected_markdown_table(self) -> None:
        label = "Markdown 표 → 한글 표"
        if not self.ensure_hwp():
            return
        try:
            copy_ok = self.run_hwp_command("Copy")
            time.sleep(0.08)
            text = self.get_clipboard_text()
            if not copy_ok or not text:
                messagebox.showwarning(label, "변환할 markdown 표 텍스트를 먼저 선택하세요.")
                self.log(f"{label}: 선택 텍스트 없음")
                return
            rows = parse_markdown_table(text)
            if rows is None:
                messagebox.showwarning(
                    label,
                    "markdown 표를 찾지 못했습니다.\n\n"
                    "예: | 항목 | 값 | 형식과 그 아래 |---|---| 구분선까지 함께 선택하세요.",
                )
                self.log(f"{label}: markdown 표 인식 실패")
                return
            html_table = markdown_rows_to_html_table(rows)
            fallback = rows_to_tsv(rows)
            self.set_clipboard_html_table(html_table, fallback)
            paste_ok = self.paste_html_original_format()
            if paste_ok:
                self.activate_hwp_window()
            self.log(f"{label}: rows={len(rows)}, cols={len(rows[0])}, Paste={paste_ok}, PasteHtmlFormat=0")
            if not paste_ok:
                messagebox.showwarning(label, "HTML 표를 클립보드에 넣었지만 한글 붙여넣기가 실패했습니다.")
        except Exception as exc:
            self.log(f"{label} 실패: {type(exc).__name__}: {exc}")
            messagebox.showerror(f"{label} 실패", str(exc))

    def clean_selected_line_breaks(self) -> None:
        self.transform_selected_text("줄바꿈/띄어쓰기 정리", clean_manual_line_breaks)

    def clean_selected_outline_prefixes(self) -> None:
        self.transform_selected_text(
            "개요 기호 제거",
            remove_manual_outline_prefixes,
            allow_cell_iteration=True,
            preserve_clipboard_cells=True,
            strip_wrapping_lines=True,
            reselect_current_cell=True,
        )

    def add_commas_to_selection(self) -> None:
        self.transform_selected_text(
            "천원단위 쉼표 넣기",
            add_thousand_commas,
            block_multiline_number_block=True,
            allow_cell_iteration=True,
            strip_wrapping_lines=True,
            reselect_current_cell=True,
        )

    def remove_commas_from_selection(self) -> None:
        self.transform_selected_text(
            "쉼표 빼기",
            remove_number_commas,
            block_multiline_number_block=True,
            allow_cell_iteration=True,
            strip_wrapping_lines=True,
            reselect_current_cell=True,
        )

    def normalize_dates_to_korean(self) -> None:
        self.transform_selected_text(
            "날짜 정규화: 0000년 0월 0일",
            lambda text: normalize_dates(text, "korean"),
            allow_cell_iteration=True,
            strip_wrapping_lines=True,
            reselect_current_cell=True,
        )

    def normalize_dates_to_dot(self) -> None:
        self.transform_selected_text(
            "날짜 정규화: 0000. 0. 0.",
            lambda text: normalize_dates(text, "dot"),
            allow_cell_iteration=True,
            strip_wrapping_lines=True,
            reselect_current_cell=True,
        )

    def normalize_dates_to_dot_padded(self) -> None:
        self.transform_selected_text(
            "날짜 정규화: 0000. 00. 00.",
            lambda text: normalize_dates(text, "dot_padded"),
            allow_cell_iteration=True,
            strip_wrapping_lines=True,
            reselect_current_cell=True,
        )

    def add_weekdays_to_selection(self) -> None:
        self.transform_selected_text(
            "요일 추가",
            add_weekdays_to_dates,
            allow_cell_iteration=True,
            strip_wrapping_lines=True,
            reselect_current_cell=True,
        )

    def palette_color_or_rgb(self, name: str, rgb: tuple[int, int, int]) -> PaletteColor:
        for color in self.palette:
            if color.name == name or color.rgb == rgb:
                return color
        return PaletteColor(name=name, rgb=rgb, hex_value=rgb_to_hex(rgb), purpose="임시 색상")

    def clean_excel_table(self) -> None:
        if not self.ensure_hwp():
            return
        results: list[str] = []
        try:
            fill_action, fill_ok, fill_default = self.set_cell_fill_color(self.palette_color_or_rgb("흰색", (255, 255, 255)))
            results.append(f"배경={fill_ok}({fill_action}, default={fill_default})")
        except Exception as exc:
            results.append(f"배경 실패:{type(exc).__name__}")
            self.debug(f"[excel-table] 배경 실패: {exc}")

        try:
            text_ok, text_default = self.set_text_color(self.palette_color_or_rgb("검정", (0, 0, 0)))
            results.append(f"글자색={text_ok}(default={text_default})")
        except Exception as exc:
            results.append(f"글자색 실패:{type(exc).__name__}")
            self.debug(f"[excel-table] 글자색 실패: {exc}")

        try:
            border_action, border_ok = self.set_table_border_preset("thin_all")
            results.append(f"얇은전체={border_ok}({border_action})")
        except Exception as exc:
            results.append(f"테두리 실패:{type(exc).__name__}")
            self.debug(f"[excel-table] 테두리 실패: {exc}")

        try:
            para_ok = self.apply_style_by_name("표내용-중간")
            results.append(f"문단스타일={para_ok}(표내용-중간)")
        except Exception as exc:
            results.append(f"문단스타일 실패:{type(exc).__name__}")
            self.debug(f"[excel-table] 문단스타일 실패: {exc}")

        self.log("엑셀표 정리: " + ", ".join(results))

def smoke_test() -> int:
    print(f"ROOT={ROOT}")
    print(f"styles={len(read_style_records())}")
    print(f"cover={read_cover_placeholders()}")
    print(f"logos={len(list_logos())}")
    print(f"commas={add_thousand_commas('1234567 2026-07-15 051-123-4567')}")
    print(f"remove={remove_number_commas('1,234,567')}")
    try:
        hwp = connect_hwp()
        for line in LAST_HWP_CONNECTION_LOG:
            print(line)
        print(f"hwp_version={hwp.Version}")
        print(f"hwp_docs={hwp.XHwpDocuments.Count}")
    except Exception as exc:
        print(f"hwp_error={type(exc).__name__}: {exc}")
        return 2
    return 0


if __name__ == "__main__":
    if "--smoke" in sys.argv:
        raise SystemExit(smoke_test())
    app = MvpApp()
    app.mainloop()
