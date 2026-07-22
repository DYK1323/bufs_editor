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

import hashlib
import json
import re
import shutil
import sys
import threading
import time
import unicodedata
import xml.etree.ElementTree as ET
import zipfile
from collections import Counter
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
    style_index: int
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

DEFAULT_CAPTION_PATTERN = r"^\[\s*(?P<kind>표|그림)\s+(?P<prefix>[^\]]*-)\s*(?:\d+)?\s*\]\s*(?P<title>.*)$"

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
        "outside_left": "0.0",
        "outside_right": "0.0",
        "outside_top": "0.0",
        "outside_bottom": "0.0",
    },
    "caption_parser": {
        "pattern": DEFAULT_CAPTION_PATTERN,
        "prefix_template": "[{kind} {prefix}",
        "suffix_template": "] {title}",
        "sample": "[그림 2.4-1] 학사관리 체계 개선 실적",
    },
}

TABLE_STYLE_ICON_PRESETS = (
    ("얇은전체", "thin_all", "얇은 전체 테두리"),
    ("굵은외곽", "thick_outer", "굵은 외곽선"),
    ("얇은상하", "thin_top_bottom", "얇은 상하선"),
    ("굵은상하", "thick_top_bottom", "굵은 상하선"),
)

TABLE_BORDER_ICON_PRESETS = (
    ("이중아래", "double_bottom", "아래 이중선"),
    ("좌우없음", "no_left_right", "좌우선 없음"),
    ("점선가로", "dotted_inside_horz", "안쪽 가로 점선"),
    ("점선세로", "dotted_inside_vert", "안쪽 세로 점선"),
    ("실선가로", "thin_inside_horz", "안쪽 가로 실선"),
    ("실선세로", "thin_inside_vert", "안쪽 세로 실선"),
)

TABLE_STYLE_ICON_BUTTON_SIZE = 48
TABLE_BORDER_ICON_BUTTON_SIZE = 48
TABLE_ICON_BUTTON_GAP = 6
TABLE_ICON_BUTTON_BG = "#FFFFFF"


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
    for style_index, match in enumerate(pattern.finditer(header)):
        attrs = dict(re.findall(r'(\w+)="([^"]*)"', match.group(1)))
        if "id" not in attrs or "name" not in attrs:
            continue
        records.append(
            StyleRecord(
                style_id=int(attrs["id"]),
                style_index=style_index,
                style_type=attrs.get("type", ""),
                name=unescape(attrs["name"]),
            )
        )
    return sorted(records, key=lambda item: (item.style_type != "PARA", item.style_id, item.name))


def read_style_records_from_hwpml(style_xml: str) -> list[StyleRecord]:
    if not style_xml.strip():
        return []
    try:
        root = ET.fromstring(style_xml)
    except ET.ParseError:
        return []

    records: list[StyleRecord] = []
    for style_index, element in enumerate(root.findall(".//STYLE")):
        raw_id = element.get("Id")
        name = element.get("Name")
        if raw_id is None or not name:
            continue
        try:
            style_id = int(raw_id)
        except ValueError:
            continue
        records.append(
            StyleRecord(
                style_id=style_id,
                style_index=style_index,
                style_type="PARA",
                name=unescape(name),
            )
        )
    return records


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


def read_image_pixel_size(path: Path) -> tuple[int, int] | None:
    data = path.read_bytes()
    if data.startswith(b"\x89PNG\r\n\x1a\n") and len(data) >= 24:
        return int.from_bytes(data[16:20], "big"), int.from_bytes(data[20:24], "big")
    if data.startswith(b"BM") and len(data) >= 26:
        width = int.from_bytes(data[18:22], "little", signed=True)
        height = int.from_bytes(data[22:26], "little", signed=True)
        return abs(width), abs(height)
    if data.startswith(b"\xff\xd8"):
        index = 2
        while index + 9 < len(data):
            if data[index] != 0xFF:
                index += 1
                continue
            marker = data[index + 1]
            index += 2
            if marker in {0xD8, 0xD9}:
                continue
            if index + 2 > len(data):
                break
            length = int.from_bytes(data[index : index + 2], "big")
            if length < 2 or index + length > len(data):
                break
            if 0xC0 <= marker <= 0xCF and marker not in {0xC4, 0xC8, 0xCC}:
                height = int.from_bytes(data[index + 3 : index + 5], "big")
                width = int.from_bytes(data[index + 5 : index + 7], "big")
                return width, height
            index += length
    return None


def create_logo_thumbnail(path: Path, max_width: int = 96, max_height: int = 48) -> Path | None:
    if win32com is None:
        return None
    try:
        stat = path.stat()
        key = f"{path.resolve()}:{stat.st_mtime_ns}:{stat.st_size}:{max_width}:{max_height}"
        name = hashlib.sha1(key.encode("utf-8", errors="ignore")).hexdigest()[:16] + ".png"
        thumb_dir = TEST_OUTPUT_DIR / "logo-thumbs"
        thumb_dir.mkdir(parents=True, exist_ok=True)
        target = thumb_dir / name
        if target.exists():
            return target

        image = win32com.client.Dispatch("WIA.ImageFile")
        image.LoadFile(str(path.resolve()))
        process = win32com.client.Dispatch("WIA.ImageProcess")
        process.Filters.Add(process.FilterInfos("Scale").FilterID)
        process.Filters(1).Properties("MaximumWidth").Value = max_width
        process.Filters(1).Properties("MaximumHeight").Value = max_height
        process.Filters.Add(process.FilterInfos("Convert").FilterID)
        process.Filters(2).Properties("FormatID").Value = "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}"
        thumbnail = process.Apply(image)
        thumbnail.SaveFile(str(target))
        return target
    except Exception:
        return None


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
    r"(?:\s*년\s*|\s*\.\s*|\s*-\s*)"
    r"(?P<month>\d{1,2})"
    r"(?:\s*월\s*|\s*\.\s*|\s*-\s*)"
    r"(?P<day>\d{1,2})"
    r"(?:\s*일|\s*\.)?"
    r"(?!\d)"
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


def remove_weekdays_from_dates(text: str) -> str:
    return re.sub(r"\s*\([월화수목금토일]\)", "", text)


def is_wrapped_regulation_name(text: str) -> bool:
    return (text.startswith(("｢", "「")) and text.endswith(("｣", "」")))


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


def clipboard_cell_values(text: str, *, preserve_tabs: bool = False) -> list[str]:
    if preserve_tabs:
        rows = parse_cell_clipboard_matrix(text)
        return [cell.strip() for row in rows for cell in row if cell.strip()]
    normalized = text.replace("\r\n", "\n").replace("\r", "\n")
    return [value.strip() for value in normalized.split("\n") if value.strip()]


def parse_cell_address(address: str) -> tuple[int, int] | None:
    match = re.fullmatch(r"\$?([A-Za-z]+)\$?(\d+)", address.strip())
    if match is None:
        return None
    col = 0
    for char in match.group(1).upper():
        col = col * 26 + (ord(char) - ord("A") + 1)
    return col, int(match.group(2))


def parse_cell_addresses_from_formula_command(command: str) -> list[tuple[int, int]]:
    raw_addresses = re.findall(r"(?<![A-Za-z0-9_])\$?[A-Za-z]+\$?\d+(?![A-Za-z0-9_])", command)
    addresses: list[tuple[int, int]] = []
    seen: set[tuple[int, int]] = set()
    for raw_address in raw_addresses:
        address = parse_cell_address(raw_address)
        if address is not None and address not in seen:
            seen.add(address)
            addresses.append(address)
    return addresses


def parse_cell_address_from_keyindicator(value: str) -> tuple[int, int] | None:
    match = re.search(r"\(([A-Za-z]+\d+)\)", value)
    if match is None:
        return None
    return parse_cell_address(match.group(1))


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


@dataclass(frozen=True)
class CaptionParts:
    prefix: str
    suffix: str


def split_table_caption_parts(text: str, parser_settings: dict | None = None) -> CaptionParts:
    cleaned = strip_wrapping_blank_lines(text).strip()
    settings = merge_dict(DEFAULT_TABLE_SETTINGS["caption_parser"], parser_settings or {})
    pattern = str(settings.get("pattern") or DEFAULT_CAPTION_PATTERN)
    try:
        match = re.match(pattern, cleaned, flags=re.S)
    except re.error as exc:
        raise ValueError(f"캡션 제목 인식 정규식 오류: {exc}") from exc
    if match is None:
        raise ValueError("`[표 ...-] 제목` 또는 `[그림 ...-] 제목` 형태를 찾지 못했습니다.")
    groupdict = match.groupdict()
    try:
        number_prefix = groupdict.get("prefix", match.group(1)).rstrip()
        title = groupdict.get("title", match.group(2) if len(match.groups()) >= 2 else "")
    except IndexError as exc:
        raise ValueError("캡션 제목 인식 정규식에는 prefix/title 또는 1번/2번 그룹이 필요합니다.") from exc
    values = {
        "kind": groupdict.get("kind", "표"),
        "prefix": number_prefix,
        "title": title,
        "g1": match.group(1) if len(match.groups()) >= 1 else "",
        "g2": match.group(2) if len(match.groups()) >= 2 else "",
    }
    try:
        prefix = str(settings.get("prefix_template", "[표 {prefix}")).format_map(values)
        suffix = str(settings.get("suffix_template", "] {title}")).format_map(values)
    except KeyError as exc:
        raise ValueError(f"캡션 조립 템플릿에서 알 수 없는 값 이름을 사용했습니다: {exc}") from exc
    return CaptionParts(prefix=prefix, suffix=suffix)


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


def manual_outline_prefix_length(text: str) -> int:
    match = OUTLINE_PREFIX_RE.match(text)
    return 0 if match is None else match.end()


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


class ToolTip:
    def __init__(self, widget, text: str) -> None:
        self.widget = widget
        self.text = text
        self.window = None
        widget.bind("<Enter>", self.show, add="+")
        widget.bind("<Leave>", self.hide, add="+")
        widget.bind("<ButtonPress>", self.hide, add="+")

    def show(self, _event=None) -> None:
        if self.window is not None or not self.text:
            return
        x = self.widget.winfo_rootx() + 16
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 6
        self.window = tk.Toplevel(self.widget)
        self.window.wm_overrideredirect(True)
        self.window.wm_geometry(f"+{x}+{y}")
        label = ttk.Label(
            self.window,
            text=self.text,
            padding=(6, 3),
            relief="solid",
            borderwidth=1,
        )
        label.pack()

    def hide(self, _event=None) -> None:
        if self.window is not None:
            self.window.destroy()
            self.window = None


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
        self.logo_height_var = tk.StringVar(value="12")
        self.logo_chip_images: list[tk.PhotoImage] = []
        self.icon_images: dict[str, tk.PhotoImage] = {}
        self.pending_caption_title = ""
        self.pending_caption_parts = CaptionParts(prefix="", suffix="")
        self._refreshing = False
        self._applying_style = False

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=6, pady=(6, 4))

        log_frame = ttk.LabelFrame(self, text="디버그 로그", padding=5)
        log_frame.pack(fill="both", expand=False, padx=6, pady=(0, 6))
        self.debug_text = tk.Text(log_frame, height=7, wrap="word")
        self.debug_text.pack(side="left", fill="both", expand=True)
        log_scroll = ttk.Scrollbar(log_frame, orient="vertical", command=self.debug_text.yview)
        log_scroll.pack(side="right", fill="y")
        self.debug_text.configure(yscrollcommand=log_scroll.set)

        self._build_styles_tab()
        self._build_table_tab()
        self._build_cover_logo_tab()
        self._build_status_tab()
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
        ttk.Button(buttons, text="대상 확인", command=self.show_current_target).pack(side="left", padx=(8, 0))
        ttk.Button(buttons, text="테스트 문서 열기", command=self.open_style_test_copy).pack(side="left", padx=(8, 0))
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

        row3 = ttk.Frame(style_group)
        row3.pack(fill="x", pady=(5, 0))
        ttk.Button(row3, text="위로", command=lambda: self.move_selected_style(-1)).pack(side="left")
        ttk.Button(row3, text="아래로", command=lambda: self.move_selected_style(1)).pack(side="left", padx=(8, 0))
        ttk.Button(row3, text="기본순서", command=self.reset_style_order).pack(side="left", padx=(8, 0))

        self.build_cleanup_controls(frame)

    def _build_table_tab(self) -> None:
        frame = ttk.Frame(self.notebook, padding=8)
        self.table_frame = frame
        self.notebook.add(frame, text="표 편집")
        self.build_table_tab_content()

    def load_icon_image(self, icon_name: str) -> tk.PhotoImage | None:
        if icon_name in self.icon_images:
            return self.icon_images[icon_name]
        icon_path = ROOT / "icons" / "2x" / f"{icon_name}@2x.png"
        if not icon_path.exists():
            self.debug(f"[icon] 아이콘 없음: {icon_path}")
            return None
        image = tk.PhotoImage(file=str(icon_path))
        self.icon_images[icon_name] = image
        return image

    def create_icon_preset_grid(
        self,
        parent: ttk.Frame,
        presets: tuple[tuple[str, str, str], ...],
        columns: int,
        button_size: int | None = None,
    ) -> None:
        for index, (icon_name, preset_name, tooltip) in enumerate(presets):
            row, col = divmod(index, columns)
            image = self.load_icon_image(icon_name)
            command = lambda name=preset_name: self.apply_table_border_preset(name)
            size = button_size or TABLE_STYLE_ICON_BUTTON_SIZE
            button = tk.Button(
                parent,
                image=image,
                command=command,
                width=size,
                height=size,
                padx=0,
                pady=0,
                relief="raised",
                bd=1,
                bg=TABLE_ICON_BUTTON_BG,
                activebackground=TABLE_ICON_BUTTON_BG,
                highlightbackground=TABLE_ICON_BUTTON_BG,
                takefocus=True,
            )
            button.grid(
                row=row,
                column=col,
                padx=(0 if col == 0 else TABLE_ICON_BUTTON_GAP, 0),
                pady=(0 if row == 0 else TABLE_ICON_BUTTON_GAP, 0),
                sticky="w",
            )
            ToolTip(button, tooltip)

    def build_table_tab_content(self) -> None:
        frame = self.table_frame
        for child in frame.winfo_children():
            child.destroy()

        ttk.Button(frame, text="표 설정", command=self.open_table_settings_window).pack(fill="x", pady=(0, 6))
        ttk.Button(frame, text="markdown 표 → 한글 표", command=self.convert_selected_markdown_table).pack(
            fill="x", pady=(0, 12)
        )

        table_action_row = ttk.Frame(frame)
        table_action_row.pack(fill="x", pady=(0, 6))
        ttk.Button(table_action_row, text="셀 속성 복사/적용", command=self.apply_cell_shape_copy_paste).pack(
            side="left", fill="x", expand=True
        )
        ttk.Button(table_action_row, text="표 바깥 여백 적용", command=self.apply_table_outside_margins).pack(
            side="left", fill="x", expand=True, padx=(6, 0)
        )
        ttk.Button(table_action_row, text="쪽 너비 맞춤", command=self.apply_table_page_width).pack(
            side="left", fill="x", expand=True, padx=(6, 0)
        )

        caption_row = ttk.Frame(frame)
        caption_row.pack(fill="x", pady=(0, 12))
        ttk.Button(caption_row, text="캡션원클릭", command=self.cut_title_and_apply_to_next_table_caption).pack(
            side="left", fill="x", expand=True
        )
        ttk.Button(caption_row, text="표 제목 Cut", command=self.cut_selected_table_caption_title).pack(
            side="left", fill="x", expand=True, padx=(6, 0)
        )
        ttk.Button(caption_row, text="캡션 조립", command=self.apply_pending_table_caption_to_next_table).pack(
            side="left", fill="x", expand=True, padx=(6, 0)
        )

        ttk.Label(frame, text="셀 배경색").pack(anchor="w")
        bg_row = ttk.Frame(frame)
        bg_row.pack(fill="x", pady=(6, 12))
        self.create_cell_fill_chips(bg_row)

        ttk.Label(frame, text="글자색").pack(anchor="w")
        text_row = ttk.Frame(frame)
        text_row.pack(fill="x", pady=(6, 12))
        self.create_color_chips(text_row, self.apply_text_color)

        ttk.Label(frame, text="표 스타일").pack(anchor="w")
        table_style_icons = ttk.Frame(frame)
        table_style_icons.pack(fill="x", pady=(6, 12))
        self.create_icon_preset_grid(
            table_style_icons,
            TABLE_STYLE_ICON_PRESETS,
            columns=4,
            button_size=TABLE_STYLE_ICON_BUTTON_SIZE,
        )

        ttk.Label(frame, text="테두리").pack(anchor="w")
        border_icons = ttk.Frame(frame)
        border_icons.pack(fill="x", pady=(6, 12))
        self.create_icon_preset_grid(
            border_icons,
            TABLE_BORDER_ICON_PRESETS,
            columns=6,
            button_size=TABLE_BORDER_ICON_BUTTON_SIZE,
        )

        ttk.Button(frame, text="제목셀 자동화", command=self.apply_title_cell_preset).pack(fill="x", pady=(0, 12))

    def create_cell_fill_chips(self, parent: ttk.Frame) -> None:
        none_button = tk.Button(
            parent,
            text="색 없음",
            bg="#FFFFFF",
            fg="#000000",
            width=8,
            relief="raised",
            command=self.clear_cell_fill_color,
        )
        none_button.grid(row=0, column=0, padx=(0, 4), pady=(0, 6), sticky="ew")
        self.create_color_chips(parent, self.apply_cell_fill_color, start_index=1)

    def create_color_chips(self, parent: ttk.Frame, command, start_index: int = 0) -> None:
        columns = 6
        for offset, color in enumerate(self.palette):
            index = start_index + offset
            button = tk.Button(
                parent,
                text=color.name,
                bg=color.hex_value,
                fg=self.contrast_text_color(color.rgb),
                width=7,
                relief="raised",
                command=lambda item=color: command(item),
            )
            button.grid(row=index // columns, column=index % columns, padx=(0, 4), pady=(0, 6), sticky="ew")
        for col in range(columns):
            parent.columnconfigure(col, weight=1)

    def contrast_text_color(self, rgb: tuple[int, int, int]) -> str:
        red, green, blue = rgb
        brightness = (red * 299 + green * 587 + blue * 114) / 1000
        return "#000000" if brightness >= 140 else "#FFFFFF"

    def open_table_settings_window(self) -> None:
        window = tk.Toplevel(self)
        window.title("표 설정")
        window.geometry("460x760")
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
        body_window = canvas.create_window((0, 0), window=body, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        def resize_body(event) -> None:
            canvas.itemconfigure(body_window, width=event.width)

        def scroll_settings(event) -> str:
            event_num = getattr(event, "num", None)
            if event_num == 4:
                canvas.yview_scroll(-3, "units")
            elif event_num == 5:
                canvas.yview_scroll(3, "units")
            else:
                canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
            return "break"

        def bind_scroll(_event) -> None:
            canvas.bind_all("<MouseWheel>", scroll_settings)
            canvas.bind_all("<Button-4>", scroll_settings)
            canvas.bind_all("<Button-5>", scroll_settings)

        def unbind_scroll(_event) -> None:
            canvas.unbind_all("<MouseWheel>")
            canvas.unbind_all("<Button-4>")
            canvas.unbind_all("<Button-5>")

        canvas.bind("<Configure>", resize_body)
        canvas.bind("<Enter>", bind_scroll)
        canvas.bind("<Leave>", unbind_scroll)
        window.bind("<Destroy>", unbind_scroll, add="+")

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
        ttk.Label(body, text="표 바깥 여백(mm)").pack(anchor="w")
        margins = settings["table_margins"]
        margin_labels = [
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

        ttk.Separator(body).pack(fill="x", pady=12)
        ttk.Label(body, text="캡션 제목 인식").pack(anchor="w")
        caption_parser = settings["caption_parser"]
        caption_rows = [
            ("pattern", "정규식", 42),
            ("prefix_template", "앞 템플릿", 24),
            ("suffix_template", "뒤 템플릿", 24),
            ("sample", "테스트 제목", 42),
        ]
        for key, label, width in caption_rows:
            row = ttk.Frame(body)
            row.pack(fill="x", pady=(6, 0))
            ttk.Label(row, text=label, width=12).pack(side="left")
            entry(row, ("caption_parser", key), str(caption_parser.get(key, "")), width=width)

        caption_test_var = tk.StringVar(value="")

        def test_caption_parser() -> None:
            parser = {
                "pattern": vars_by_path[("caption_parser", "pattern")].get(),
                "prefix_template": vars_by_path[("caption_parser", "prefix_template")].get(),
                "suffix_template": vars_by_path[("caption_parser", "suffix_template")].get(),
            }
            sample = vars_by_path[("caption_parser", "sample")].get()
            try:
                parts = split_table_caption_parts(sample, parser)
            except ValueError as exc:
                caption_test_var.set(f"실패: {exc}")
                return
            caption_test_var.set(f"앞: {parts.prefix} / 뒤: {parts.suffix}")

        caption_test_row = ttk.Frame(body)
        caption_test_row.pack(fill="x", pady=(6, 0))
        ttk.Button(caption_test_row, text="인식 테스트", command=test_caption_parser).pack(side="left")
        ttk.Label(caption_test_row, textvariable=caption_test_var, wraplength=300).pack(
            side="left", fill="x", expand=True, padx=(8, 0)
        )

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
                        messagebox.showwarning("여백 값 확인", "표 바깥 여백은 mm 단위 숫자로 입력하세요.", parent=window)
                        return
                    if number < 0:
                        messagebox.showwarning("여백 값 확인", "표 바깥 여백은 0 이상으로 입력하세요.", parent=window)
                        return
                    value = f"{number:g}"
                if path[:1] == ("caption_parser",) and path[-1] == "pattern":
                    try:
                        re.compile(value)
                    except re.error as exc:
                        messagebox.showwarning("캡션 정규식 확인", f"정규식 오류: {exc}", parent=window)
                        return
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
        hwp_buttons2 = ttk.Frame(cleanup_group)
        hwp_buttons2.pack(fill="x", pady=(6, 0))
        ttk.Button(hwp_buttons2, text="｢규정명｣", command=self.wrap_regulation_names_in_selection).pack(
            side="left", fill="x", expand=True
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
        ttk.Button(date_buttons2, text="요일 제거", command=self.remove_weekdays_from_selection).pack(
            side="left", fill="x", expand=True, padx=(6, 0)
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

        logo_group = ttk.LabelFrame(frame, text="로고 삽입", padding=8)
        logo_group.pack(fill="both", expand=True)
        logo_size_row = ttk.Frame(logo_group)
        logo_size_row.pack(fill="x", pady=(0, 8))
        ttk.Label(logo_size_row, text="높이(mm)").pack(side="left")
        ttk.Entry(logo_size_row, textvariable=self.logo_height_var, width=8).pack(side="left", padx=(6, 0))
        ttk.Label(logo_size_row, text="칩을 누르면 커서 위치에 삽입됩니다.").pack(side="left", padx=(10, 0))

        self.logo_chip_frame = ttk.Frame(logo_group)
        self.logo_chip_frame.pack(fill="both", expand=True)

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

        logos = list_logos()
        self.populate_logo_chips(logos)
        self.log(f"로고 파일 수: {len(logos)}")
        self._refreshing = False

    def populate_logo_chips(self, logos: list[Path]) -> None:
        for child in self.logo_chip_frame.winfo_children():
            child.destroy()
        self.logo_chip_images = []
        if not logos:
            ttk.Label(self.logo_chip_frame, text="bufs/logos 폴더에 PNG/JPG/BMP 파일을 넣어주세요.").grid(
                row=0, column=0, sticky="w"
            )
            return

        row = 0
        col = 0
        max_columns = 6
        for logo in logos:
            rel = str(logo.relative_to(LOGO_DIR))
            pixel_size = read_image_pixel_size(logo)
            aspect = (pixel_size[0] / pixel_size[1]) if pixel_size and pixel_size[1] else 1.0
            is_wide = aspect > 1.35
            thumb_width, thumb_height = (220, 46) if is_wide else (46, 46)
            span = 4 if is_wide else 1
            if col + span > max_columns:
                row += 1
                col = 0

            image = None
            thumb_path = create_logo_thumbnail(logo, thumb_width, thumb_height)
            if thumb_path is not None:
                try:
                    image = tk.PhotoImage(file=str(thumb_path))
                    self.logo_chip_images.append(image)
                except Exception as thumb_exc:
                    self.debug(f"[logo-chip] 썸네일 로드 실패: {rel}: {type(thumb_exc).__name__}: {thumb_exc}")

            try:
                if image is None:
                    image = tk.PhotoImage(file=str(logo))
                    factor = max(1, (image.width() + thumb_width - 1) // thumb_width, (image.height() + thumb_height - 1) // thumb_height)
                    if factor > 1:
                        image = image.subsample(factor, factor)
                    self.logo_chip_images.append(image)
            except Exception as exc:
                self.debug(f"[logo-chip] 미리보기 실패: {rel}: {type(exc).__name__}: {exc}")

            button = ttk.Button(
                self.logo_chip_frame,
                text="" if image is not None else rel,
                image=image,
                command=lambda item=logo: self.insert_logo_at_cursor(item),
            )
            button.grid(row=row, column=col, columnspan=span, sticky="nsew", padx=(0, 6), pady=(0, 6))
            col += span
            if col >= max_columns:
                row += 1
                col = 0

        for column in range(max_columns):
            self.logo_chip_frame.columnconfigure(column, weight=1, minsize=50)

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

    def hwp_brush_type(self, name: str, fallback: int) -> int:
        try:
            return int(self.hwp.BrushType(name))
        except Exception:
            return fallback

    def hwp_hatch_style(self, name: str, fallback: int) -> int:
        try:
            return int(self.hwp.HatchStyle(name))
        except Exception:
            return fallback

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

    def clear_cell_fill_color(self) -> None:
        if not self.ensure_hwp():
            return
        try:
            action, ok, default_ok = self.set_cell_fill_none()
            self.log(f"셀 배경색 제거: action={action}, GetDefault={default_ok}, result={ok}")
            if not ok:
                messagebox.showwarning(
                    "셀 배경색 제거 확인",
                    "한글이 셀 배경색 제거를 실패로 반환했습니다.\n표 안의 셀을 선택한 상태인지 확인하세요.",
                )
        except Exception as exc:
            self.log(f"셀 배경색 제거 실패: {type(exc).__name__}: {exc}")
            messagebox.showerror("셀 배경색 제거 실패", str(exc))

    def set_cell_fill_color(self, color: PaletteColor) -> tuple[str, bool, bool]:
        value = self.hwp_rgb(color)
        pset = self.hwp.HParameterSet.HCellBorderFill
        default_ok = bool(self.hwp.HAction.GetDefault("CellBorderFill", pset.HSet))

        fill_targets = [pset.FillAttr]
        try:
            fill_targets.append(pset.SelCellsBorderFill.FillAttr)
        except Exception:
            pass
        try:
            fill_targets.append(pset.AllCellsBorderFill.FillAttr)
        except Exception:
            pass

        brush_type = self.hwp_brush_type("WinBrush", 1)
        no_hatch = self.hwp_hatch_style("None", -1)
        for fill in fill_targets:
            for attr, attr_value in (
                ("type", brush_type),
                ("WindowsBrush", 1),
                ("WinBrushFaceStyle", no_hatch),
                ("WinBrushFaceColor", value),
                ("WinBrushHatchColor", value),
                ("ToolbarColor", value),
            ):
                self.set_com_attr(fill, attr, attr_value)

        action, ok = self.execute_first_hwp_action(("CellBorderFill", "CellBorder"), pset.HSet)
        return action, ok, default_ok

    def set_cell_fill_none(self) -> tuple[str, bool, bool]:
        pset = self.hwp.HParameterSet.HCellBorderFill
        default_ok = bool(self.hwp.HAction.GetDefault("CellBorderFill", pset.HSet))

        fill_targets = [pset.FillAttr]
        try:
            fill_targets.append(pset.SelCellsBorderFill.FillAttr)
        except Exception:
            pass
        try:
            fill_targets.append(pset.AllCellsBorderFill.FillAttr)
        except Exception:
            pass

        white = self.hwp_rgb(PaletteColor("흰색", (255, 255, 255), "#FFFFFF", ""))
        no_hatch = self.hwp_hatch_style("None", -1)
        for fill in fill_targets:
            for attr, value in (
                ("type", 0),
                ("WindowsBrush", 0),
                ("WinBrushFaceStyle", no_hatch),
                ("WinBrushFaceColor", white),
                ("WinBrushHatchColor", white),
                ("ToolbarColor", white),
            ):
                self.set_com_attr(fill, attr, value)

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
        self.debug(
            f"[style] Execute target name={record.name}, id={record.style_id}, "
            f"index={record.style_index}, type={record.style_type}"
        )
        self.configure_style_apply_set(pset, record.style_id)
        return bool(self.hwp.HAction.Execute("Style", pset.HSet))

    def apply_style_by_name(self, style_name: str) -> bool:
        self.refresh_current_doc_style_map(force=True)
        record = self.find_current_doc_style_record(style_name)
        if record is None:
            self.debug(f"[title-cell] 현재 문서에서 스타일 이름 매칭 실패, 적용 중단: {style_name}")
            messagebox.showwarning(
                "스타일 이름 매칭 실패",
                f"현재 문서에서 같은 이름의 스타일을 찾지 못해 적용을 중단했습니다.\n\n"
                f"찾는 스타일: {style_name}\n\n"
                "스타일 번호로 대신 적용하면 다른 스타일이 적용될 수 있습니다.",
            )
            return False
        ok = self.execute_style_record(record)
        self.debug(
            f"[title-cell] 스타일 적용 {style_name}, id={record.style_id}, "
            f"index={record.style_index}, type={record.style_type}, "
            f"result={ok}"
        )
        return ok

    def apply_first_style_containing(self, text: str) -> bool:
        needle = normalize_style_name(text)
        self.refresh_current_doc_style_map(force=True)
        candidates = list(self.current_doc_style_map.values())
        if not candidates:
            self.debug(f"[style] 현재 문서 스타일 맵이 비어 있어 포함 이름 적용 중단: {text}")
            return False
        record = next(
            (
                item
                for item in candidates
                if item.style_type == "PARA" and needle in normalize_style_name(item.name)
            ),
            None,
        )
        if record is None:
            self.debug(f"[style] 포함 이름 없음: {text}")
            return False
        ok = self.execute_style_record(record)
        self.debug(f"[style] 포함 이름 적용: {record.name}, id={record.style_id}, index={record.style_index}, result={ok}")
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

    def get_parameter_child(self, target, child_name: str, log_prefix: str = "param"):
        for candidate in (target, getattr(target, "HSet", None)):
            if candidate is None:
                continue
            try:
                return getattr(candidate, child_name)
            except Exception:
                pass
            try:
                return candidate.Item(child_name)
            except Exception as exc:
                self.debug(f"[{log_prefix}] {child_name} Item 실패: {type(exc).__name__}: {exc}")
        return None

    def commit_cell_shape(self, cell_shape) -> bool:
        try:
            setattr(self.hwp, "CellShape", cell_shape)
            return True
        except Exception as exc:
            self.debug(f"[title-cell-header] CellShape 커밋 실패: {type(exc).__name__}: {exc}")
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

    def apply_cell_shape_copy_paste(self) -> None:
        if not self.ensure_hwp():
            return
        try:
            pset = self.hwp.HParameterSet.HShapeCopyPaste
            default_ok = self.hwp.HAction.GetDefault("ShapeCopyPaste", pset.HSet)
            self.debug(f"[shape-copy-paste] GetDefault ShapeCopyPaste={default_ok}")
            for attr, value in (
                ("CellAttr", 1),
                ("CellBorder", 0),
                ("CellFill", 0),
                ("TypeBodyAndCellOnly", 1),
            ):
                applied = self.set_com_attr(pset, attr, value)
                self.debug(f"[shape-copy-paste] {attr}={value}, set={applied}")
            action, ok = self.execute_first_hwp_action(("ShapeCopyPaste",), pset.HSet)
            self.log(
                "셀 속성 복사/적용: "
                f"action={action}, CellAttr=1, CellBorder=0, CellFill=0, TypeBodyAndCellOnly=1, result={ok}"
            )
            if not ok:
                messagebox.showwarning(
                    "셀 속성 복사/적용 확인",
                    "복사할 셀에 커서를 두거나, 적용할 셀 범위를 선택한 상태인지 확인하세요.",
                )
        except Exception as exc:
            self.log(f"셀 속성 복사/적용 실패: {type(exc).__name__}: {exc}")
            messagebox.showerror("셀 속성 복사/적용 실패", str(exc))

    def cut_selected_table_caption_title(self) -> None:
        self.cut_selected_table_caption_title_core("표 제목 잘라두기")

    def selected_block_contains_table(self) -> bool:
        try:
            block_xml = safe_str(self.hwp.GetTextFile("HWPML2X", "saveblock"))
        except Exception as exc:
            self.debug(f"[caption-cut] 선택 블록 HWPML 확인 실패: {type(exc).__name__}: {exc}")
            return False
        upper_xml = block_xml.upper()
        return "<TABLE" in upper_xml or "<HP:TBL" in upper_xml

    def cut_selected_table_caption_title_core(self, label: str) -> bool:
        if not self.ensure_hwp():
            return False
        try:
            if self.has_selected_object() or self.selected_block_contains_table():
                messagebox.showwarning(
                    label,
                    "표까지 선택된 상태에서는 제목을 자르지 않습니다.\n\n"
                    "표 바로 윗줄 제목 텍스트만 블록 선택한 뒤 다시 실행하세요.",
                )
                self.log(f"{label}: 표 포함 선택으로 Cut 중단")
                return False

            copy_ok = self.run_hwp_command("Copy")
            time.sleep(0.08)
            title_text = strip_wrapping_blank_lines(self.get_clipboard_text()).strip()
            if not copy_ok or not title_text:
                messagebox.showwarning(label, "표 바로 윗줄 제목만 블록 선택한 뒤 다시 실행하세요.")
                self.log(f"{label}: Copy={copy_ok}, 제목 텍스트 없음")
                return False

            try:
                parts = split_table_caption_parts(title_text, self.table_settings.get("caption_parser", {}))
            except ValueError as exc:
                preview = title_text.replace("\r", "\\r").replace("\n", "\\n")
                messagebox.showwarning(label, f"{exc}\n\n클립보드 텍스트:\n{preview[:200]}")
                self.log(f"{label}: 제목 형식 인식 실패, clipboard={preview!r}")
                return False

            cut_ok = self.run_hwp_command("Cut")
            time.sleep(0.08)
            if not cut_ok:
                messagebox.showwarning(label, "제목 형식은 확인했지만 선택 텍스트를 자르지 못했습니다.")
                self.log(f"{label}: Cut 실패")
                return False

            self.pending_caption_title = title_text
            self.pending_caption_parts = parts
            self.activate_hwp_window()
            self.log(
                f"{label}: Cut={cut_ok}, 제목 {len(title_text)}자, "
                f"앞 {len(parts.prefix)}자 / 뒤 {len(parts.suffix)}자"
            )
            return True
        except Exception as exc:
            self.log(f"{label} 실패: {type(exc).__name__}: {exc}")
            messagebox.showerror(f"{label} 실패", str(exc))
            return False

    def apply_pending_table_caption_to_next_table(self) -> None:
        label = "다음 표 캡션 조립"
        if not self.ensure_hwp():
            return
        if not self.pending_caption_title:
            messagebox.showwarning(label, "먼저 `표 제목 잘라두기`를 실행하세요.")
            return
        try:
            table_selected = self.has_selected_object()
            self.debug(f"[caption-build] selected_object={table_selected}")
            if not table_selected:
                messagebox.showwarning(label, "표 개체를 선택한 상태에서 실행하세요.")
                self.log(f"{label}: 표 개체 선택 상태 아님")
                return
            self.apply_pending_table_caption_to_selected_table_core(label)
        except Exception as exc:
            self.log(f"{label} 실패: {type(exc).__name__}: {exc}")
            messagebox.showerror(f"{label} 실패", str(exc))

    def apply_pending_table_caption_to_selected_table_core(self, label: str) -> bool:
        try:
            action, caption_ok = "ShapeObjCaption", self.run_hwp_command("ShapeObjCaption")
            self.debug(f"[caption-build] caption action={action}, result={caption_ok}")
            if not caption_ok:
                messagebox.showwarning(label, "표는 선택했지만 캡션을 만들지 못했습니다.")
                self.log(f"{label}: 캡션 액션 실패, action={action}")
                return False

            parts = self.pending_caption_parts
            remove_tail_space_ok = self.run_hwp_command("DeleteBack")
            suffix_ok = remove_tail_space_ok and (not parts.suffix or self.insert_hwp_text(parts.suffix))

            begin_action, begin_ok = "MoveLineBegin", self.run_hwp_command("MoveLineBegin")
            remove_label_ok = begin_ok and self.delete_hwp_chars(2)
            prefix_ok = remove_label_ok and (not parts.prefix or self.insert_hwp_text(parts.prefix))

            style_ok = False
            if prefix_ok and suffix_ok:
                style_ok = self.apply_first_style_containing("캡션")
                moved_text = self.pending_caption_title
                self.pending_caption_title = ""
                self.pending_caption_parts = CaptionParts(prefix="", suffix="")
                self.activate_hwp_window()
                self.log(
                    f"{label}: caption={action}:{caption_ok}, begin={begin_action}:{begin_ok}, "
                    f"remove_label={remove_label_ok}, prefix={prefix_ok}, "
                    f"remove_tail_space={remove_tail_space_ok}, "
                    f"suffix={suffix_ok}, caption_style={style_ok}, 제목 {len(moved_text)}자"
                )
                return True

            messagebox.showwarning(label, "캡션은 만들었지만 제목 앞/뒤 조립에 실패했습니다.")
            self.log(
                f"{label}: 조립 실패, begin={begin_action}:{begin_ok}, remove_label={remove_label_ok}, "
                f"prefix={prefix_ok}, remove_tail_space={remove_tail_space_ok}, suffix={suffix_ok}"
            )
            return False
        except Exception as exc:
            self.log(f"{label} 실패: {type(exc).__name__}: {exc}")
            messagebox.showerror(f"{label} 실패", str(exc))
            return False

    def cut_title_and_apply_to_next_table_caption(self) -> None:
        label = "선택 제목 → 다음 표 캡션"
        if not self.ensure_hwp():
            return
        if not self.cut_selected_table_caption_title_core(label):
            return
        self.run_hwp_command("Cancel")
        time.sleep(0.05)
        table_selected = self.select_next_table_object_below_cursor()
        self.debug(f"[caption-one-step] select_next_table={table_selected}")
        if not table_selected:
            messagebox.showwarning(label, "제목은 잘라냈지만 다음 표를 선택하지 못했습니다.")
            self.log(f"{label}: 다음 표 선택 실패")
            return
        self.apply_pending_table_caption_to_selected_table_core(label)

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

    def current_page_text_width(self) -> int:
        sec = self.hwp.HParameterSet.HSecDef
        default_ok = bool(self.hwp.HAction.GetDefault("PageSetup", sec.HSet))
        page = sec.PageDef
        paper_width = int(page.PaperWidth)
        left = int(page.LeftMargin)
        right = int(page.RightMargin)
        gutter = int(page.GutterLen)
        width = max(0, paper_width - left - right - gutter)
        self.debug(
            "[table-page-width] "
            f"PageSetup={default_ok}, paper={paper_width}, left={left}, right={right}, "
            f"gutter={gutter}, text_width={width}"
        )
        return width

    def set_table_page_width(self) -> tuple[str, bool, int]:
        view_option = None
        pset = self.get_table_property_set()
        try:
            table_ctrl = self.hwp.CurSelectedCtrl
        except Exception:
            table_ctrl = None
        target_width = self.current_page_text_width()
        try:
            target_width -= int(pset.OutsideMarginLeft) + int(pset.OutsideMarginRight)
        except Exception:
            pass
        if target_width <= 0:
            return "HWPML2X", False, target_width

        table_xml = self.hwp.GetTextFile("HWPML2X", "saveblock")
        if not table_xml:
            return "HWPML2X", False, target_width

        root = ET.fromstring(table_xml)
        table = root.find(".//TABLE")
        if table is None:
            return "HWPML2X", False, target_width

        rows = table.findall(".//ROW")
        first_row_cells = rows[0].findall(".//CELL") if rows else table.findall(".//CELL")
        current_width = sum(int(cell.get("Width") or 0) for cell in first_row_cells)
        if current_width <= 0:
            return "HWPML2X", False, target_width

        ratio = target_width / current_width
        shape_size = table.find(".//SIZE")
        if shape_size is not None and shape_size.get("Width"):
            shape_size.set("Width", str(target_width))
        changed = 0
        for cell in table.findall(".//CELL"):
            width = cell.get("Width")
            if not width:
                continue
            cell.set("Width", str(max(1, int(round(int(width) * ratio)))))
            changed += 1

        updated_xml = ET.tostring(root, encoding="UTF-16").decode("utf-16")
        try:
            view_option = self.hwp.ViewProperties.Item("OptionFlag")
            if view_option not in (2, 6):
                prop = self.hwp.ViewProperties
                prop.SetItem("OptionFlag", 6)
                self.hwp.ViewProperties = prop
        except Exception:
            view_option = None

        deleted = False
        restored = False
        if table_ctrl is not None:
            try:
                deleted = bool(self.hwp.DeleteCtrl(table_ctrl))
            except Exception as exc:
                self.debug(f"[table-page-width] DeleteCtrl 실패: {type(exc).__name__}: {exc}")

        ok = deleted and bool(self.hwp.SetTextFile(updated_xml, "HWPML2X", "insertfile"))
        if deleted and not ok:
            try:
                restored = bool(self.hwp.SetTextFile(table_xml, "HWPML2X", "insertfile"))
            except Exception as exc:
                self.debug(f"[table-page-width] 원본 복구 실패: {type(exc).__name__}: {exc}")

        if view_option is not None:
            try:
                prop = self.hwp.ViewProperties
                prop.SetItem("OptionFlag", view_option)
                self.hwp.ViewProperties = prop
            except Exception:
                pass
        self.debug(
            "[table-page-width] "
            f"target={target_width}, current={current_width}, ratio={ratio:.6f}, "
            f"cells={changed}, shape_size={shape_size is not None}, "
            f"deleted={deleted}, restored={restored}, result={ok}"
        )
        return "HWPML2X", ok, target_width

    def apply_table_page_width(self) -> None:
        if not self.ensure_hwp():
            return
        try:
            action, ok, width = self.set_table_page_width()
            self.log(f"표 쪽 너비 맞춤: width={width}, action={action}, result={ok}")
            if not ok:
                messagebox.showwarning("표 쪽 너비 맞춤 확인", "표 안에 커서를 두거나 표 개체를 선택한 상태인지 확인하세요.")
        except Exception as exc:
            self.log(f"표 쪽 너비 맞춤 실패: {type(exc).__name__}: {exc}")
            messagebox.showerror("표 쪽 너비 맞춤 실패", str(exc))

    def select_current_table_object(self) -> bool:
        if self.run_hwp_command("SelectCtrlReverse"):
            time.sleep(0.03)
            return True
        return False

    def select_next_table_object_below_cursor(self, max_down_moves: int = 8) -> bool:
        self.clear_hwp_selection()
        for move_count in range(1, max_down_moves + 1):
            moved = self.run_hwp_command("MoveDown")
            time.sleep(0.03)
            try:
                in_cell = bool(getattr(self.hwp, "CellShape"))
            except Exception:
                in_cell = False
            selected = self.select_current_table_object()
            self.debug(
                f"[caption-one-step] next-table search move={move_count}, "
                f"MoveDown={moved}, CellShape={in_cell}, SelectCtrlReverse={selected}"
            )
            if selected:
                return True
            if not moved:
                break
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

    def set_selected_cells_as_header(self) -> tuple[str, bool, bool]:
        set_ok = False
        commit_ok = False
        action_name = "CellShape"

        try:
            cell_shape = self.hwp.CellShape
        except Exception as exc:
            cell_shape = None
            self.debug(f"[title-cell-header] CellShape 읽기 실패: {type(exc).__name__}: {exc}")

        if cell_shape is not None:
            if self.set_parameter_item(cell_shape, "RepeatHeader", 1, "title-cell-header"):
                self.debug("[title-cell-header] CellShape.RepeatHeader=1")
                set_ok = True
            cell_param = self.get_parameter_child(cell_shape, "Cell", "title-cell-header")
            if cell_param is not None and self.set_parameter_item(cell_param, "Header", 1, "title-cell-header"):
                self.debug("[title-cell-header] CellShape.Cell.Header=1")
                set_ok = True
                self.set_parameter_item(cell_shape, "Cell", cell_param, "title-cell-header")
            if set_ok:
                commit_ok = self.commit_cell_shape(cell_shape)
                self.debug(f"[title-cell-header] CellShape 커밋={commit_ok}")

        if not commit_ok:
            try:
                pset = self.hwp.HParameterSet.HShapeObject
                default_ok = bool(self.hwp.HAction.GetDefault("TablePropertyDialog", pset.HSet))
                self.debug(f"[title-cell-header] TablePropertyDialog GetDefault={default_ok}")
                for item, value in (("ShapeType", 3), ("ShapeCellSize", 0)):
                    if self.set_parameter_item(pset, item, value, "title-cell-header"):
                        self.debug(f"[title-cell-header] HShapeObject.{item}={value}")
                shape_cell = self.get_parameter_child(pset, "ShapeTableCell", "title-cell-header")
                if shape_cell is not None and self.set_parameter_item(shape_cell, "Header", 1, "title-cell-header"):
                    self.debug("[title-cell-header] HShapeObject.ShapeTableCell.Header=1")
                    set_ok = True
                if default_ok:
                    action_name, commit_ok = self.execute_first_hwp_action(("TablePropertyDialog", "ShapeObjDialog"), pset.HSet)
            except Exception as exc:
                self.debug(f"[title-cell-header] TablePropertyDialog fallback 실패: {type(exc).__name__}: {exc}")
        return action_name, set_ok, commit_ok

    def apply_title_cell_preset(self) -> None:
        if not self.ensure_hwp():
            return
        title = self.table_settings["title_cell"]
        results: list[str] = []

        try:
            header_action, header_set_ok, header_action_ok = self.set_selected_cells_as_header()
            results.append(f"제목셀속성=set:{header_set_ok}, action:{header_action_ok}({header_action})")
        except Exception as exc:
            results.append(f"제목셀속성 실패:{type(exc).__name__}")
            self.debug(f"[title-cell] 제목셀속성 실패: {exc}")

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

        try:
            border_action, border_ok = self.set_title_cell_borders()
            results.append(f"테두리={border_ok}({border_action})")
        except Exception as exc:
            results.append(f"테두리 실패:{type(exc).__name__}")
            self.debug(f"[title-cell] 테두리 실패: {exc}")

        try:
            bg_action, bg_ok, bg_default = self.set_cell_fill_color(self.palette_color(title["background_color"]))
            results.append(f"배경={bg_ok}({bg_action}, default={bg_default})")
        except Exception as exc:
            results.append(f"배경 실패:{type(exc).__name__}")
            self.debug(f"[title-cell] 배경 실패: {exc}")

        try:
            text_ok, text_default = self.set_text_color(self.palette_color(title["text_color"]))
            results.append(f"글자색={text_ok}(default={text_default})")
        except Exception as exc:
            results.append(f"글자색 실패:{type(exc).__name__}")
            self.debug(f"[title-cell] 글자색 실패: {exc}")

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
        if self._refreshing or self._applying_style or not self.apply_on_select.get():
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

    def read_current_doc_style_records_from_memory(self) -> list[StyleRecord]:
        if self.hwp is None:
            return []
        for option in ("", "saveblock:true", "saveblock"):
            try:
                style_xml = safe_str(self.hwp.GetTextFile("HWPML2X", option))
            except Exception as exc:
                self.debug(f"[style-map] HWPML2X 스타일 읽기 실패 option={option!r}: {type(exc).__name__}: {exc}")
                continue
            records = read_style_records_from_hwpml(style_xml)
            if records:
                self.debug(f"[style-map] HWPML2X 메모리 스타일 {len(records)}개 읽음 option={option!r}")
                return records
            self.debug(f"[style-map] HWPML2X 스타일 없음 option={option!r}, bytes={len(style_xml)}")
        return []

    def refresh_current_doc_style_map(self, *, force: bool = False) -> None:
        path = self.get_current_hwp_path()
        if path is None:
            self.current_doc_style_path = None
            self.current_doc_style_map = {}
            self.current_doc_style_norm_map = {}
            self.debug("[style-map] 현재 문서 경로를 확인하지 못함")
            return

        if not force and path == self.current_doc_style_path and self.current_doc_style_map:
            return

        self.current_doc_style_path = path
        self.current_doc_style_map = {}
        self.current_doc_style_norm_map = {}

        records = self.read_current_doc_style_records_from_memory()
        if not records:
            if path.suffix.lower() != ".hwpx":
                self.debug(f"[style-map] 현재 문서가 hwpx가 아니고 메모리 스타일도 없어 이름 매칭 불가: {path}")
                return
            if not path.exists():
                self.debug(f"[style-map] 현재 문서 파일을 디스크에서 찾지 못함: {path}")
                return
            records = read_style_records(path)
            self.debug(f"[style-map] HWPX 파일 fallback 스타일 {len(records)}개 읽음: {path}")
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
        if self._applying_style:
            self.debug(f"[style] 적용 중 재진입 무시: reason={reason}")
            return
        selection = self.style_list.curselection()
        if not selection:
            messagebox.showwarning("스타일 선택 필요", "적용할 스타일을 먼저 선택하세요.")
            return

        record = self.style_records[selection[0]]
        self.debug(f"[style] 요청 reason={reason}, id={record.style_id}, type={record.style_type}, name={record.name}")
        if record.style_type != "PARA":
            self.debug(f"[style] 글자 스타일 적용 시도: type={record.style_type}, name={record.name}")

        self._applying_style = True
        try:
            if self.hwp is None:
                self.debug("[hwp] COM 연결 시작")
                self.hwp = connect_hwp()
                self.debug("[hwp] COM 연결 성공")
                for line in LAST_HWP_CONNECTION_LOG:
                    self.debug(line)
                for line in describe_hwp(self.hwp):
                    self.debug(f"[hwp] {line}")
            self.refresh_current_doc_style_map(force=True)
            target_record = self.find_current_doc_style_record(record.name)
            if target_record is None:
                current_path = self.get_current_hwp_path()
                self.debug(
                    f"[style] 현재 문서 스타일 이름 매칭 실패, 적용 중단: "
                    f"name={record.name}, current_doc={current_path}, 기준id={record.style_id}"
                )
                messagebox.showwarning(
                    "스타일 이름 매칭 실패",
                    "현재 문서에서 같은 이름의 스타일을 찾지 못해 적용을 중단했습니다.\n\n"
                    "스타일 번호로 대신 적용하면 다른 스타일이 적용될 수 있습니다.",
                )
                return

            ok = self.execute_style_record(target_record)
            self.log(
                "스타일 적용 요청: "
                f"name={record.name}, 기준id={record.style_id}, 현재문서id={target_record.style_id}, "
                f"현재문서index={target_record.style_index}, "
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
        finally:
            self.after(250, self._finish_style_apply)

    def _finish_style_apply(self) -> None:
        self._applying_style = False

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

    def insert_logo_at_cursor(self, path: Path) -> None:
        if not self.ensure_hwp():
            return
        try:
            height_mm = float(self.logo_height_var.get())
            if height_mm <= 0:
                raise ValueError
        except ValueError:
            messagebox.showwarning("로고 높이 확인", "로고 높이는 0보다 큰 mm 단위 숫자로 입력하세요.")
            return

        try:
            pixel_size = read_image_pixel_size(path)
            if pixel_size is None:
                width_mm = height_mm
                self.debug(f"[logo-insert] 이미지 크기 인식 실패, 정사각형으로 삽입: {path.name}")
            else:
                pixel_width, pixel_height = pixel_size
                width_mm = height_mm * pixel_width / pixel_height

            ok = self.insert_picture_at_cursor(path, width_mm, height_mm)
            self.log(
                f"로고 삽입: file={path.name}, width={width_mm:g}mm, height={height_mm:g}mm, result={ok}"
            )
            if not ok:
                messagebox.showwarning("로고 삽입 확인", "한글의 그림 삽입 가능한 위치에 커서를 둔 뒤 다시 실행하세요.")
            else:
                self.activate_hwp_window()
        except Exception as exc:
            self.log(f"로고 삽입 실패: {type(exc).__name__}: {exc}")
            messagebox.showerror("로고 삽입 실패", str(exc))

    def insert_picture_at_cursor(self, path: Path, width_mm: float, height_mm: float) -> bool:
        try:
            result = self.hwp.InsertPicture(
                str(path),
                Embedded=True,
                sizeoption=1,
                Width=width_mm,
                Height=height_mm,
            )
            return bool(result)
        except Exception as method_exc:
            self.debug(f"[logo-insert] InsertPicture method 실패: {type(method_exc).__name__}: {method_exc}")

        option = self.hwp.HParameterSet.HInsertPicture
        self.hwp.HAction.GetDefault("InsertPicture", option.HSet)
        for attr in ("FileName", "Filename", "filename"):
            self.set_com_attr(option, attr, str(path))
        for attr, value in {
            "Width": width_mm,
            "Height": height_mm,
            "SizeOption": 1,
            "sizeoption": 1,
            "Embedded": 1,
        }.items():
            self.set_com_attr(option, attr, value)
        action, ok = self.execute_first_hwp_action(("InsertPicture",), option.HSet)
        self.debug(f"[logo-insert] {action}={ok}")
        return bool(ok)

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

    def insert_hwp_text(self, text: str) -> bool:
        if not text:
            return True
        try:
            pset = self.hwp.HParameterSet.HInsertText
            self.hwp.HAction.GetDefault("InsertText", pset.HSet)
            pset.Text = text
            return bool(self.hwp.HAction.Execute("InsertText", pset.HSet))
        except Exception as exc:
            self.debug(f"[insert-text] 실패: {type(exc).__name__}: {exc}")
            return False

    def delete_hwp_chars(self, count: int) -> bool:
        ok = True
        for _ in range(max(0, count)):
            ok = self.run_hwp_command("Delete") and ok
        return ok

    def move_hwp_chars(self, command: str, count: int) -> bool:
        ok = True
        for _ in range(max(0, count)):
            ok = self.run_hwp_command(command) and ok
        return ok

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
            # SW_RESTORE breaks Windows snapped/split layouts even when the
            # window is already visible. Restore only minimized HWP windows.
            if win32gui.IsIconic(hwnd):
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
        clipboard_text = normalize_clipboard_newlines(text)
        last_exc = None
        for _ in range(5):
            try:
                win32clipboard.OpenClipboard()
                try:
                    win32clipboard.EmptyClipboard()
                    if hasattr(win32clipboard, "SetClipboardText"):
                        win32clipboard.SetClipboardText(clipboard_text, win32clipboard.CF_UNICODETEXT)
                    else:
                        win32clipboard.SetClipboardData(win32clipboard.CF_UNICODETEXT, clipboard_text)
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

    def get_current_cell_address(self) -> tuple[int, int] | None:
        try:
            indicator = self.hwp.KeyIndicator()
        except Exception as exc:
            self.debug(f"[cell-address] KeyIndicator 실패: {type(exc).__name__}: {exc}")
            return None
        for value in reversed(indicator if isinstance(indicator, tuple) else (indicator,)):
            address = parse_cell_address_from_keyindicator(safe_str(value))
            if address is not None:
                return address
        self.debug(f"[cell-address] KeyIndicator 주소 없음: {indicator!r}")
        return None

    def move_between_table_addresses(
        self,
        current: tuple[int, int],
        target: tuple[int, int],
    ) -> bool:
        current_col, current_row = current
        target_col, target_row = target
        if target_col > current_col and not self.move_table_cell("TableRightCell", target_col - current_col):
            return False
        if target_col < current_col and not self.move_table_cell("TableLeftCell", current_col - target_col):
            return False
        if target_row > current_row and not self.move_table_cell("TableLowerCell", target_row - current_row):
            return False
        if target_row < current_row and not self.move_table_cell("TableUpperCell", current_row - target_row):
            return False
        return True

    def set_hwp_pos(self, pos: tuple[int, int, int]) -> bool:
        try:
            return bool(self.hwp.SetPos(pos[0], pos[1], pos[2]))
        except Exception:
            return False

    def set_hwp_pos_by_set(self, pos_set) -> bool:
        try:
            return bool(self.hwp.SetPosBySet(pos_set))
        except Exception as exc:
            self.debug(f"[hwp-pos-set] SetPosBySet 실패: {type(exc).__name__}: {exc}")
            return False

    def get_hwp_pos_by_set(self):
        try:
            return self.hwp.GetPosBySet()
        except Exception as exc:
            self.debug(f"[hwp-pos-set] GetPosBySet 실패: {type(exc).__name__}: {exc}")
            return None

    def get_selected_cell_block_start_pos_by_scan(self):
        try:
            if not getattr(self.hwp, "CellShape"):
                return None
        except Exception:
            return None

        original_pos = self.get_hwp_pos_by_set()
        start_pos = None
        try:
            self.hwp.InitScan(1, 0xFF)
            try:
                self.hwp.GetText()
                self.hwp.MovePos(201)
                start_pos = self.get_hwp_pos_by_set()
            finally:
                try:
                    self.hwp.ReleaseScan()
                except Exception:
                    pass
        except Exception as exc:
            self.debug(f"[cell-scan-start] 선택 셀 첫 위치 확인 실패: {type(exc).__name__}: {exc}")
            start_pos = None
        finally:
            if original_pos is not None:
                self.set_hwp_pos_by_set(original_pos)
        return start_pos

    def get_selected_cell_range_by_formula(self) -> dict | None:
        try:
            pset = self.hwp.HParameterSet.HFieldCtrl
            self.hwp.HAction.GetDefault("TableFormula", pset.HSet)
            command = safe_str(pset.Command)
        except Exception as exc:
            self.debug(f"[cell-formula-range] TableFormula 범위 확인 실패: {type(exc).__name__}: {exc}")
            return None
        addresses = parse_cell_addresses_from_formula_command(command)
        if not addresses:
            self.debug(f"[cell-formula-range] 주소 없음: command={command!r}")
            return None
        cols = [address[0] for address in addresses]
        rows = [address[1] for address in addresses]
        first_col, last_col = min(cols), max(cols)
        first_row, last_row = min(rows), max(rows)
        return {
            "command": command,
            "addresses": addresses,
            "first_col": first_col,
            "last_col": last_col,
            "first_row": first_row,
            "last_row": last_row,
            "cols": last_col - first_col + 1,
            "rows": last_row - first_row + 1,
            "cells": (last_col - first_col + 1) * (last_row - first_row + 1),
        }

    def get_selection_mode_base(self) -> tuple[int | None, int | None]:
        try:
            raw = int(getattr(self.hwp, "SelectionMode"))
        except Exception as exc:
            self.debug(f"[selection-mode] SelectionMode 확인 실패: {type(exc).__name__}: {exc}")
            return None, None
        return raw, raw & 0x0F

    def is_selected_cell_block(self) -> bool:
        raw, base = self.get_selection_mode_base()
        if base == 3:
            self.debug(f"[selection-mode] 셀 선택 확인: raw={raw}, base={base}")
            return True
        self.debug(f"[selection-mode] 셀 선택 아님: raw={raw}, base={base}")
        return False

    def get_selected_text_positions(self) -> tuple[tuple[int, int, int], tuple[int, int, int]] | None:
        try:
            result = self.hwp.GetSelectedPos()
        except Exception as exc:
            self.debug(f"[selection-pos] GetSelectedPos 실패: {type(exc).__name__}: {exc}")
            return None
        if not isinstance(result, tuple) or len(result) < 7 or not result[0]:
            self.debug(f"[selection-pos] GetSelectedPos invalid: {result}")
            return None
        start = (int(result[1]), int(result[2]), int(result[3]))
        end = (int(result[4]), int(result[5]), int(result[6]))
        if start == end:
            return None
        ordered_start, ordered_end = sorted((start, end))
        return ordered_start, ordered_end

    def get_selected_pos_endpoints(self) -> tuple[tuple[int, int, int], tuple[int, int, int]] | None:
        try:
            result = self.hwp.GetSelectedPos()
        except Exception as exc:
            self.debug(f"[selection-pos] GetSelectedPos raw 실패: {type(exc).__name__}: {exc}")
            return None
        if not isinstance(result, tuple) or len(result) < 7 or not result[0]:
            self.debug(f"[selection-pos] GetSelectedPos raw invalid: {result}")
            return None
        start = (int(result[1]), int(result[2]), int(result[3]))
        end = (int(result[4]), int(result[5]), int(result[6]))
        return start, end

    def selected_paragraph_range(self) -> tuple[int, int, int] | None:
        positions = self.get_selected_text_positions()
        if positions is None:
            return None
        start, end = positions
        if start[0] != end[0]:
            return None
        first_para = start[1]
        last_para = end[1]
        if end[2] == 0 and last_para > first_para:
            last_para -= 1
        if last_para < first_para:
            return None
        return start[0], first_para, last_para

    def parse_hwp_get_text_result(self, result) -> tuple[int | None, str]:
        if isinstance(result, str):
            return None, result
        if isinstance(result, (tuple, list)):
            status = next((int(value) for value in result if isinstance(value, int)), None)
            text = next((value for value in result if isinstance(value, str)), "")
            return status, text
        if isinstance(result, int):
            return result, ""
        return None, ""

    def read_current_paragraph_text(self, list_id: int, para: int) -> str | None:
        original_pos = self.get_hwp_pos_by_set()
        parts: list[str] = []
        try:
            if not self.set_hwp_pos((list_id, para, 0)):
                return None
            try:
                init_ok = self.hwp.InitScan(0, 0x0010 | 0x0003, para, 0, para, 0)
            except TypeError:
                init_ok = self.hwp.InitScan(0, 0x0010 | 0x0003)
            if not init_ok:
                return None
            try:
                for _ in range(200):
                    status, text = self.parse_hwp_get_text_result(self.hwp.GetText())
                    if text:
                        parts.append(text)
                    if status in (0, 1, 101, 102):
                        break
                    if status == 3 and not text:
                        break
            finally:
                try:
                    self.hwp.ReleaseScan()
                except Exception:
                    pass
        except Exception as exc:
            self.debug(f"[paragraph-text] 문단 읽기 실패 para={para}: {type(exc).__name__}: {exc}")
            return None
        finally:
            if original_pos is not None:
                self.set_hwp_pos_by_set(original_pos)
        text = "".join(parts).replace("\r\n", "\n").replace("\r", "\n")
        return text.split("\n", 1)[0]

    def get_current_heading_string(self) -> str:
        try:
            return safe_str(self.hwp.GetHeadingString())
        except Exception as exc:
            self.debug(f"[paragraph-heading] GetHeadingString 실패: {type(exc).__name__}: {exc}")
            return ""

    def clear_current_paragraph_heading(self) -> bool:
        try:
            heading = self.get_current_heading_string()
            pset = self.hwp.HParameterSet.HParaShape
            default_ok = self.hwp.HAction.GetDefault("ParagraphShape", pset.HSet)
            current_heading_type = int(getattr(pset, "HeadingType", 0))
            if not heading and current_heading_type == 0:
                return False
            pset.HeadingType = 0
            ok = bool(self.hwp.HAction.Execute("ParagraphShape", pset.HSet))
            self.debug(
                f"[paragraph-heading] 제거 default={default_ok}, "
                f"heading={heading!r}, type={current_heading_type}, result={ok}"
            )
            return ok
        except Exception as exc:
            self.debug(f"[paragraph-heading] 제거 실패: {type(exc).__name__}: {exc}")
            return False

    def delete_hwp_text_range(self, list_id: int, para: int, start_pos: int, end_pos: int) -> bool:
        if end_pos <= start_pos:
            return False
        try:
            if not self.set_hwp_pos((list_id, para, start_pos)):
                return False
            if hasattr(self.hwp, "SelectText") and self.hwp.SelectText(para, start_pos, para, end_pos):
                return self.run_hwp_command("Delete")
        except Exception as exc:
            self.debug(f"[paragraph-delete] SelectText 실패 para={para}: {type(exc).__name__}: {exc}")
        if not self.set_hwp_pos((list_id, para, start_pos)):
            return False
        for _ in range(end_pos - start_pos):
            if not self.run_hwp_command("MoveSelRight"):
                return False
        return self.run_hwp_command("Delete")

    def reselect_hwp_text(self, start: tuple[int, int, int], text: str) -> bool:
        if not text or not self.set_hwp_pos(start):
            return False
        for _ in range(len(text)):
            if not self.run_hwp_command("MoveSelRight"):
                return False
        return True

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
        preferred_shape: tuple[int, int] | None = None,
        only_preferred_shape: bool = False,
    ) -> dict | None:
        total = len(selected_values)
        if total < 2:
            return None
        corners = ("bottom_right", "bottom_left", "top_right", "top_left")
        best: dict | None = None
        selected_counter = Counter(selected_values)
        shapes: list[tuple[int, int]] = []
        if preferred_shape is not None and preferred_shape[0] * preferred_shape[1] == total:
            shapes.append(preferred_shape)
        if not only_preferred_shape:
            for rows in range(1, total + 1):
                if total % rows != 0:
                    continue
                cols = total // rows
                shape = (rows, cols)
                if cols > 1 and shape not in shapes:
                    shapes.append(shape)
            for rows in range(1, total + 1):
                if total % rows != 0:
                    continue
                shape = (rows, total // rows)
                if shape not in shapes:
                    shapes.append(shape)
        for rows, cols in shapes:
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
                if cols > 1 and Counter(values) == selected_counter:
                    return {
                        "rows": rows,
                        "cols": cols,
                        "corner": corner,
                        "start_label": start_label,
                        "match": "unordered",
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

    def transform_formula_address_cells(
        self,
        label: str,
        addresses: list[tuple[int, int]],
        transform,
        *,
        strip_wrapping_lines: bool = False,
    ) -> tuple[int, int] | None:
        if len(addresses) < 2:
            return None
        address_positions = self.snapshot_formula_address_positions(addresses)
        if address_positions is None:
            return None

        visited = 0
        changed = 0
        for address, cell_pos in address_positions:
            if not self.set_hwp_pos_by_set(cell_pos):
                self.debug(f"[cell-address-iterate] 셀 위치 복원 실패: address={address}")
                return None
            current_text = self.read_current_cell_text()
            if current_text is None:
                return None
            if not self.set_hwp_pos_by_set(cell_pos):
                return None
            source_text = strip_wrapping_blank_lines(current_text) if strip_wrapping_lines else current_text
            transformed = transform(source_text) if source_text else source_text
            visited += 1
            if source_text and transformed != source_text:
                if not self.select_current_table_cell_for_replace():
                    return None
                self.set_clipboard_text(transformed)
                if not self.paste_text_into_selected_cell():
                    return None
                changed += 1
            self.clear_hwp_selection()
        self.log(f"{label}: TableFormula 주소 순회 완료, visited={visited}, changed={changed}")
        return visited, changed

    def snapshot_formula_address_positions(
        self,
        addresses: list[tuple[int, int]],
    ) -> list[tuple[tuple[int, int], object]] | None:
        self.clear_hwp_selection()
        time.sleep(0.03)
        current_address = self.get_current_cell_address()
        if current_address is None:
            return None
        if current_address != addresses[0]:
            if not self.move_between_table_addresses(current_address, addresses[0]):
                return None
            current_address = self.get_current_cell_address()
            if current_address != addresses[0]:
                self.debug(
                    f"[cell-address-snapshot] 첫 셀 이동 검증 실패: "
                    f"expected={addresses[0]}, actual={current_address}"
                )
                return None

        address_positions: list[tuple[tuple[int, int], object]] = []
        current_address = addresses[0]
        for address in addresses:
            if address != current_address:
                if not self.move_between_table_addresses(current_address, address):
                    return None
                actual_address = self.get_current_cell_address()
                if actual_address != address:
                    self.debug(
                        f"[cell-address-snapshot] 셀 이동 검증 실패: "
                        f"from={current_address}, expected={address}, actual={actual_address}"
                    )
                    return None
                current_address = actual_address
            cell_pos = self.get_hwp_pos_by_set()
            if cell_pos is None:
                self.debug(f"[cell-address-snapshot] 셀 위치 저장 실패: address={address}")
                return None
            address_positions.append((address, cell_pos))
            current_address = address
        self.debug(f"[cell-address-snapshot] positions={len(address_positions)}, addresses={addresses}")
        return address_positions

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

        expected_values = [
            strip_wrapping_blank_lines(value) if strip_wrapping_lines else value
            for value in selected_values
        ]

        range_info = self.get_selected_cell_range_by_formula()
        if range_info is not None:
            if range_info["cols"] != 1 or range_info["rows"] != len(selected_values):
                self.log(
                    f"{label}: TableFormula 참고값 불일치, "
                    f"range={range_info}, values={len(selected_values)}; 실제 셀 내용 검증으로 진행"
                )
            else:
                self.debug(f"[cell-column-iterate] TableFormula range={range_info}")

        def read_column_values_from_current() -> tuple[tuple[int, int, int] | None, list[str]]:
            try:
                start = self.hwp.GetPos()
            except Exception:
                return None, []
            values: list[str] = []
            for index in range(len(selected_values)):
                if index > 0 and not self.move_table_cell("TableLowerCell"):
                    return start, []
                value = self.read_current_cell_text()
                if value is None:
                    return start, []
                values.append(strip_wrapping_blank_lines(value) if strip_wrapping_lines else value)
            return start, values

        start_pos: tuple[int, int, int] | None = None
        actual_values: list[str] = []
        scan_start_pos = self.get_selected_cell_block_start_pos_by_scan()
        if scan_start_pos is not None and self.set_hwp_pos_by_set(scan_start_pos):
            start_pos, actual_values = read_column_values_from_current()
        if actual_values != expected_values and self.set_hwp_pos(endpoint) and self.move_table_cell("TableUpperCell", len(selected_values) - 1):
            start_pos, actual_values = read_column_values_from_current()
        if actual_values != expected_values:
            if self.set_hwp_pos(endpoint):
                self.clear_hwp_selection()
                top_start_pos, top_actual_values = read_column_values_from_current()
                if top_actual_values == expected_values:
                    start_pos = top_start_pos
                    actual_values = top_actual_values

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
        if start_pos is None or not self.set_hwp_pos(start_pos):
            return False
        self.clear_hwp_selection()

        visited = 0
        changed = 0
        stop_reason = ""
        for index in range(len(selected_values)):
            if index > 0 and not self.move_table_cell("TableLowerCell"):
                stop_reason = f"row-move-failed:index={index}"
                break
            if not self.select_current_table_cell_for_replace():
                stop_reason = f"cell-select-failed:index={index}"
                break
            if not self.run_hwp_command("Copy"):
                stop_reason = f"copy-failed:index={index}"
                break
            time.sleep(0.05)
            current_text = self.get_clipboard_text()
            source_text = strip_wrapping_blank_lines(current_text) if strip_wrapping_lines else current_text
            transformed = transform(source_text) if source_text else source_text
            visited += 1
            if source_text and transformed != source_text:
                self.set_clipboard_text(transformed)
                if not self.paste_text_into_selected_cell():
                    stop_reason = f"paste-failed:index={index}"
                    break
                changed += 1
            self.clear_hwp_selection()

        if visited != len(selected_values):
            if changed == 0:
                self.set_hwp_pos(start_pos)
            messagebox.showwarning(
                label,
                "선택한 셀 범위를 순회하던 중 중단했습니다.\n\n"
                "일부 셀이 이미 바뀌었으면 한글에서 Ctrl+Z로 되돌린 뒤, 더 작은 범위로 다시 실행하세요.",
            )
            self.log(
                f"{label}: 한 열 순회 중단, visited={visited}, changed={changed}, "
                f"expected={len(selected_values)}, reason={stop_reason}"
            )
            return True

        self.activate_hwp_window()
        self.log(f"{label}: 한 열 셀 순회 완료, visited={visited}, changed={changed}")
        return True

    def transform_selected_cell_matrix(
        self,
        label: str,
        text: str,
        transform,
        *,
        strip_wrapping_lines: bool = False,
        range_info: dict | None = None,
    ) -> bool:
        if not self.is_selected_cell_block():
            self.debug(f"[cell-matrix] {label}: SelectionMode가 셀 블록이 아니어서 일반 텍스트 경로 사용")
            return False
        if range_info is None:
            range_info = self.get_selected_cell_range_by_formula()

        matrix = parse_cell_clipboard_matrix(text)
        if matrix and len(matrix) == 1 and len(matrix[0]) == 1:
            return False

        formula_addresses = list(range_info.get("addresses", [])) if range_info is not None else []
        if len(formula_addresses) < 2 and looks_like_single_copied_cell(text):
            self.debug(f"[cell-matrix] {label}: 단일 셀 블록으로 보고 일반 텍스트 경로 사용")
            return False
        if len(formula_addresses) >= 2:
            result = self.transform_formula_address_cells(
                label,
                formula_addresses,
                transform,
                strip_wrapping_lines=strip_wrapping_lines,
            )
            if result is not None:
                visited, iter_changed = result
                self.activate_hwp_window()
                self.log(
                    f"{label}: TableFormula 주소 기반 셀 변환 완료, "
                    f"rows={range_info['rows'] if range_info else '?'}, "
                    f"cols={range_info['cols'] if range_info else '?'}, "
                    f"visited={visited}, changed={iter_changed}"
                )
                return True
            messagebox.showwarning(
                label,
                "선택한 셀 주소 범위를 순회하지 못해 중단했습니다.\n\n"
                "TableFormula가 선택 셀 주소를 반환했으므로 다른 셀을 추측해서 훑지 않습니다.",
            )
            self.log(
                f"{label}: TableFormula 주소 기반 순회 실패, 감지 스캔 없이 중단, "
                f"addresses={formula_addresses}"
            )
            return True

        messagebox.showwarning(
            label,
            "선택한 셀 주소 범위를 확인하지 못해 중단했습니다.\n\n"
            "테이블 전체를 추측해서 훑지 않습니다. 셀 범위를 다시 블록 선택한 뒤 실행하세요.",
        )
        self.log(f"{label}: 셀 주소 스냅샷 없음, 테이블 추측 순회 중단")
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
        reselect_after_paste: bool = False,
    ) -> None:
        if not self.ensure_hwp():
            return
        try:
            selected_positions = self.get_selected_text_positions() if reselect_after_paste else None
            pre_copy_cell_range = None
            pre_copy_cell_block = allow_cell_iteration and self.is_selected_cell_block()
            if pre_copy_cell_block:
                pre_copy_cell_range = self.get_selected_cell_range_by_formula()
                if pre_copy_cell_range is not None:
                    self.log(
                        f"{label}: Copy 전 셀 범위 스냅샷, "
                        f"rows={pre_copy_cell_range['rows']}, cols={pre_copy_cell_range['cols']}, "
                        f"addresses={pre_copy_cell_range.get('addresses')}"
                    )
            copy_ok = self.run_hwp_command("Copy")
            time.sleep(0.08)
            text = self.get_clipboard_text()
            if not copy_ok or not text:
                messagebox.showwarning(label, "한글에서 변환할 글자 영역을 먼저 선택하세요.")
                self.log(f"{label}: 선택 텍스트 없음")
                return
            if pre_copy_cell_block and self.transform_selected_cell_matrix(
                label,
                text,
                transform,
                strip_wrapping_lines=strip_wrapping_lines,
                range_info=pre_copy_cell_range,
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
            if block_multiline_number_block and self.is_selected_cell_block() and looks_like_multiline_number_block(text):
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
            text_reselected = False
            if paste_ok and reselect_current_cell:
                time.sleep(0.03)
                cell_selected = self.run_hwp_command("TableCellBlock")
            if paste_ok and not cell_selected and selected_positions is not None:
                time.sleep(0.03)
                text_reselected = self.reselect_hwp_text(selected_positions[0], transformed)
            if paste_ok:
                self.activate_hwp_window()
            self.log(
                f"{label}: Copy={copy_ok}, Paste={paste_ok}, "
                f"CellReselect={cell_selected}, TextReselect={text_reselected}, "
                f"{len(text)}자 -> {len(transformed)}자"
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
        self.transform_selected_text(
            "줄바꿈/띄어쓰기 정리",
            clean_manual_line_breaks,
            allow_cell_iteration=True,
            reselect_after_paste=True,
        )

    def clean_selected_outline_prefixes(self) -> None:
        if not self.ensure_hwp():
            return
        if not self.is_selected_cell_block():
            self.clean_selected_paragraph_outline_prefixes()
            return
        self.transform_selected_text(
            "개요 기호 제거",
            remove_manual_outline_prefixes,
            allow_cell_iteration=True,
            preserve_clipboard_cells=True,
            strip_wrapping_lines=True,
            reselect_current_cell=True,
            reselect_after_paste=True,
        )

    def clean_selected_paragraph_outline_prefixes(self) -> None:
        label = "개요 기호 제거"
        try:
            paragraph_range = self.selected_paragraph_range()
            if paragraph_range is None:
                messagebox.showwarning(label, "한글에서 개요 기호를 제거할 문단 영역을 먼저 선택하세요.")
                self.log(f"{label}: 문단 선택 범위 확인 실패")
                return

            list_id, first_para, last_para = paragraph_range
            original_pos = self.get_hwp_pos_by_set()
            self.clear_hwp_selection()
            changed = 0
            visited = 0
            for para in range(first_para, last_para + 1):
                if not self.set_hwp_pos((list_id, para, 0)):
                    self.log(f"{label}: 문단 이동 실패 para={para}")
                    continue
                visited += 1
                if self.clear_current_paragraph_heading():
                    changed += 1

                paragraph_text = self.read_current_paragraph_text(list_id, para)
                if paragraph_text is None:
                    self.log(f"{label}: 문단 텍스트 읽기 실패 para={para}")
                    continue
                prefix_length = manual_outline_prefix_length(paragraph_text)
                if prefix_length == 0:
                    continue
                if self.delete_hwp_text_range(list_id, para, 0, prefix_length):
                    changed += 1
                else:
                    self.log(f"{label}: 수동 기호 삭제 실패 para={para}, prefix={paragraph_text[:prefix_length]!r}")

            if original_pos is not None:
                self.set_hwp_pos_by_set(original_pos)
            self.activate_hwp_window()
            self.log(
                f"{label}: 문단 순회 완료, paras={visited}, changed={changed}, "
                f"range={first_para}-{last_para}"
            )
            if changed == 0:
                messagebox.showinfo(label, "선택한 문단에서 제거할 개요 기호를 찾지 못했습니다.")
        except Exception as exc:
            self.log(f"{label} 실패: {type(exc).__name__}: {exc}")
            messagebox.showerror(f"{label} 실패", str(exc))

    def wrap_regulation_names_in_selection(self) -> None:
        label = "｢규정명｣"
        if not self.ensure_hwp():
            return
        try:
            positions = self.get_selected_text_positions()
            copy_ok = self.run_hwp_command("Copy")
            time.sleep(0.08)
            text = self.get_clipboard_text()
            if not copy_ok or not text:
                messagebox.showwarning(label, "한글에서 감쌀 규정명을 먼저 선택하세요.")
                self.log(f"{label}: 선택 텍스트 없음")
                return
            if "\t" in text:
                messagebox.showwarning(
                    label,
                    "표 셀 여러 칸 선택은 문단/글자 스타일 보존이 불안정해서 중단했습니다.\n\n"
                    "한 셀 안의 규정명 텍스트만 선택한 뒤 다시 실행하세요.",
                )
                self.log(f"{label}: 표 셀 범위 선택으로 중단")
                return
            if is_wrapped_regulation_name(text.strip()):
                self.activate_hwp_window()
                self.log(f"{label}: 이미 감싸져 있어 변경 없음")
                return
            stripped_text = strip_wrapping_blank_lines(text)
            if not stripped_text or "\n" in stripped_text or "\r" in stripped_text:
                messagebox.showwarning(
                    label,
                    "규정명 한 줄만 선택한 뒤 다시 실행하세요.\n\n"
                    "여러 줄 선택은 개요번호/문단 경계 때문에 서식 보존이 불안정해서 중단했습니다.",
                )
                self.log(f"{label}: 여러 줄 또는 빈 선택으로 중단")
                return

            if positions is None:
                messagebox.showwarning(
                    label,
                    "선택한 글자 위치를 확인하지 못해 중단했습니다.\n\n"
                    "규정명 글자만 드래그해 선택한 뒤 다시 실행하세요.",
                )
                self.log(f"{label}: 선택 위치 확인 실패")
                return

            start, end = positions
            if not self.set_hwp_pos(end) or not self.insert_hwp_text("｣"):
                messagebox.showwarning(
                    label,
                    "닫는 괄호 삽입에 실패했습니다.",
                )
                self.log(f"{label}: 닫는 괄호 삽입 실패")
                return
            if not self.set_hwp_pos(start) or not self.insert_hwp_text("｢"):
                messagebox.showwarning(
                    label,
                    "여는 괄호 삽입에 실패했습니다.\n\n"
                    "문서에 닫는 괄호만 들어갔다면 한글에서 Ctrl+Z로 되돌리세요.",
                )
                self.log(f"{label}: 여는 괄호 삽입 실패")
                return

            self.activate_hwp_window()
            self.log(f"{label}: 선택 텍스트 스타일 유지 방식으로 괄호 삽입 완료")
        except Exception as exc:
            self.log(f"{label} 실패: {type(exc).__name__}: {exc}")
            messagebox.showerror(f"{label} 실패", str(exc))

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

    def remove_weekdays_from_selection(self) -> None:
        self.transform_selected_text(
            "요일 제거",
            remove_weekdays_from_dates,
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
