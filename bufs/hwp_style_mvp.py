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
import io
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import unicodedata
import urllib.request
import webbrowser
import xml.etree.ElementTree as ET
import zipfile
from collections import Counter
from dataclasses import dataclass, field
from decimal import Decimal, ROUND_DOWN, ROUND_FLOOR, ROUND_HALF_UP, ROUND_UP
from datetime import datetime
from html import escape
from html import unescape
from pathlib import Path
from typing import Callable, Iterable

from config_store import ensure_user_config_file
from config_store import read_json_file
from config_store import write_json_file
from template_assets import builtin_template_dir_or_fallback
from template_assets import sync_builtin_templates as sync_builtin_template_assets
from text_transforms import CURRENCY_UNIT_FACTORS
from text_transforms import SuspiciousDecimalNumberError
from text_transforms import UnsafeUnitConversionNumberError
from text_transforms import add_thousand_commas
from text_transforms import add_weekdays_to_dates
from text_transforms import cell_matrix_size
from text_transforms import cell_matrix_to_tsv
from text_transforms import convert_decimal_numbers_to_currency_unit
from text_transforms import ensure_no_suspicious_decimal_numbers
from text_transforms import ensure_safe_decimal_unit_conversion_text
from text_transforms import find_suspicious_decimal_numbers
from text_transforms import flatten_cell_matrix
from text_transforms import is_rectangular_cell_matrix
from text_transforms import is_standard_regulation_wrapper
from text_transforms import looks_like_cell_clipboard_matrix
from text_transforms import looks_like_multiline_number_block
from text_transforms import looks_like_single_copied_cell
from text_transforms import needs_regulation_wrapper_conversion
from text_transforms import normalize_clipboard_newlines
from text_transforms import normalize_dates
from text_transforms import normalize_decimal_numbers
from text_transforms import normalize_years
from text_transforms import parse_cell_clipboard_matrix
from text_transforms import remove_number_commas
from text_transforms import remove_weekdays_from_dates
from text_transforms import scale_decimal_numbers
from text_transforms import scale_decimal_numbers_for_unit_conversion
from text_transforms import strip_wrapping_blank_lines
from text_transforms import transform_cell_matrix



def app_root() -> Path:
    if getattr(sys, "frozen", False):
        bundle_root = Path(getattr(sys, "_MEIPASS", Path(sys.executable).resolve().parent))
        for candidate in (bundle_root / "bufs", Path(sys.executable).resolve().parent / "bufs", bundle_root):
            if (
                (candidate / "style-sets.json").exists()
                or (candidate / "templates" / "보고서 본문 서식.hwpx").exists()
                or (candidate / "보고서 본문 서식.hwpx").exists()
            ):
                return candidate
    return Path(__file__).resolve().parent


def user_config_root() -> Path:
    base = os.environ.get("APPDATA")
    if base:
        return Path(base) / "BUFS-HWP-Editor"
    return Path.home() / ".bufs-hwp-editor"


ROOT = app_root()
if Path.cwd().resolve() != ROOT:
    # 설치 폴더 루트에는 쉽다한글이 번들링한 python312.dll이 있어
    # Python 3.11 가상환경의 pywin32 로딩과 충돌할 수 있다.
    os.chdir(ROOT)

CONFIG_ROOT = user_config_root()
BUNDLED_TEMPLATE_DIR = ROOT / "templates"
USER_TEMPLATE_ROOT = CONFIG_ROOT / "templates"
BUILTIN_TEMPLATE_DIR = USER_TEMPLATE_ROOT / "builtin"
CUSTOM_TEMPLATE_DIR = USER_TEMPLATE_ROOT / "custom"


def sync_builtin_templates(
    bundled_dir: Path = BUNDLED_TEMPLATE_DIR,
    builtin_dir: Path = BUILTIN_TEMPLATE_DIR,
    custom_dir: Path = CUSTOM_TEMPLATE_DIR,
) -> list[str]:
    return sync_builtin_template_assets(bundled_dir, builtin_dir, custom_dir)


TEMPLATE_SYNC_WARNINGS = sync_builtin_templates()

try:
    import tkinter as tk
    import tkinter.font as tkfont
    from tkinter import ttk, messagebox, filedialog
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


INSTALL_ROOT = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parents[1]
TEMPLATE_DIR = builtin_template_dir_or_fallback(BUILTIN_TEMPLATE_DIR, BUNDLED_TEMPLATE_DIR)
STYLE_FILE = TEMPLATE_DIR / "보고서 본문 서식.hwpx"
COVER_FILE = TEMPLATE_DIR / "표지.hwpx"
GENERAL_REPORT_TEMPLATE_FILE = TEMPLATE_DIR / "일반보고_양식.hwpx"
TITLE_NUMBER_BOX_TEMPLATE_FILE = TEMPLATE_DIR / "대제목_번호박스.hwpx"
PROOF_TEMPLATE_FILE = TEMPLATE_DIR / "증빙_양식.hwpx"
LOGO_DIR = ROOT / "logos"
TEST_OUTPUT_DIR = ROOT / "test-output"
BUNDLED_STYLE_ORDER_FILE = ROOT / "style-order.json"
BUNDLED_STYLE_SETS_FILE = ROOT / "style-sets.json"
BUNDLED_TABLE_SETTINGS_FILE = ROOT / "table-settings.json"
BUNDLED_UPDATE_SETTINGS_FILE = ROOT / "update-settings.json"
STYLE_ORDER_FILE = CONFIG_ROOT / "style-order.json"
STYLE_SETS_FILE = CONFIG_ROOT / "style-sets.json"
TABLE_SETTINGS_FILE = CONFIG_ROOT / "table-settings.json"
UPDATE_SETTINGS_FILE = CONFIG_ROOT / "update-settings.json"
SPECIAL_CHARS_FILE = CONFIG_ROOT / "special-chars.json"
LAST_HWP_CONNECTION_LOG: list[str] = []
APP_VERSION = "1.0.3"
APP_NAME = "BUFS-HWP-Editor"
TITLE_NUMBER_BOX_MARKER = "{{bufs_title}}"
TITLE_NUMBER_BOX_MARKER_SEPARATOR = "::"
TITLE_NUMBER_BOX_NUMBER_PLACEHOLDER = "{{no}}"
TITLE_NUMBER_BOX_TITLE_PLACEHOLDER = "{{header}}"
PROOF_TITLE_MARKER = "{{증빙제목}}"
PROOF_IMAGE_MARKER = "{{증빙자료}}"
HWPUNIT_PER_MM = 7200 / 25.4
PROOF_IMAGE_CELL_WIDTH = 48190
PROOF_IMAGE_FIRST_CELL_HEIGHT = 63566
PROOF_IMAGE_FULL_CELL_HEIGHT = 70016
PROOF_IMAGE_CELL_MARGIN_X = 510
PROOF_IMAGE_CELL_MARGIN_Y = 141
PROOF_DEFAULT_DPI = 300
PROOF_JPEG_QUALITY = 92
DEFAULT_UPDATE_SETTINGS = {
    "enabled": True,
    "check_on_start": True,
    "provider": "github",
    "github_owner": "DYK1323",
    "github_repo": "bufs_editor",
    "asset_pattern": "BUFS-HWP-Editor-v*.zip",
    "expected_root_dir": "BUFS-HWP-Editor",
    "expected_exe_name": "BUFS-HWP-Editor.exe",
    "version_url": "",
    "download_url": "",
    "timeout_seconds": 5,
}
DEFAULT_TITLE_NUMBER_BOX_SETTINGS = {
    "enabled": True,
    "role": "title_number_box",
    "marker": TITLE_NUMBER_BOX_MARKER,
    "template_file": "templates/builtin/대제목_번호박스.hwpx",
    "placeholder_title": TITLE_NUMBER_BOX_TITLE_PLACEHOLDER,
    "placeholder_number": TITLE_NUMBER_BOX_NUMBER_PLACEHOLDER,
    "outside_margin": {
        "left": 0,
        "right": 0,
        "top": 0,
        "bottom": 850,
    },
    "table_style_preset": "thin_all",
    "cells": [
        {
            "width": 2487,
            "height": 2414,
            "paragraph_style": "바탕글 사본1",
            "fill_rgb": [217, 217, 217],
            "margins": {"left": 510, "right": 510, "top": 141, "bottom": 141},
        },
        {
            "width": 45211,
            "height": 2414,
            "paragraph_style": "대제목",
            "margins": {"left": 510, "right": 510, "top": 141, "bottom": 141},
        },
        {
            "width": 498,
            "height": 2414,
            "text_color_rgb": [255, 255, 255],
            "margins": {"left": 0, "right": 0, "top": 0, "bottom": 0},
        },
    ],
}
DEFAULT_SPECIAL_CHARS: list[dict[str, str]] = []
DEFAULT_SPECIAL_CHAR_GROUPS: tuple[tuple[str, tuple[str, ...]], ...] = (
    ("원숫자", tuple(chr(codepoint) for codepoint in range(0x2460, 0x2474))),
    ("로마숫자", tuple(chr(codepoint) for codepoint in range(0x2160, 0x216C))),
    (
        "화살표/기호",
        (
            "※",
            *tuple(chr(codepoint) for codepoint in range(0x2190, 0x2194)),
            chr(0x279C),
            *tuple(chr(codepoint) for codepoint in range(0xF0E7, 0xF0EB)),
        ),
    ),
    ("수학 연산자", ("±", "÷", "×")),
    (
        "도형",
        (
            chr(0x25B3),
            chr(0x25B7),
            chr(0x25BD),
            chr(0x25C1),
            chr(0x25B2),
            chr(0x25B6),
            chr(0x25BC),
            chr(0x25C0),
            chr(0x2606),
            chr(0x2605),
            chr(0x25CB),
            chr(0x25CF),
            chr(0x25C7),
            chr(0x25C6),
            chr(0x25A1),
            chr(0x25A0),
            chr(0x25A3),
            chr(0xF06D),
            chr(0xF071),
        ),
    ),
)
SPECIAL_CHAR_BUTTON_DRAWINGS = {
    chr(0xF0E7): "arrow-left",
    chr(0xF0E8): "arrow-right",
    chr(0xF0E9): "arrow-up",
    chr(0xF0EA): "arrow-down",
    chr(0xF06D): "shadow-circle",
    chr(0xF071): "shadow-square",
}
NUMBERING_GROUP_OPTIONS: tuple[tuple[str, str], ...] = (
    ("", ""),
    ("body", "본문"),
    ("table", "표"),
    ("note", "주석"),
    ("proof", "증빙"),
    ("appendix", "별첨"),
    ("custom1", "기타1"),
    ("custom2", "기타2"),
)
NUMBERING_GROUP_LABELS = tuple(label for _key, label in NUMBERING_GROUP_OPTIONS)
NUMBERING_GROUP_BY_LABEL = {label: key for key, label in NUMBERING_GROUP_OPTIONS}
NUMBERING_GROUP_LABEL_BY_KEY = {key: label for key, label in NUMBERING_GROUP_OPTIONS}
INLINE_RULE_TYPE_OPTIONS: tuple[tuple[str, str], ...] = (
    ("leading_parenthesized_after_identifier", "첫머리 괄호 키워드"),
    ("leading_text_before_colon_after_identifier", "첫머리 콜론 앞"),
    ("markdown_bold", "마크다운 굵게"),
)
INLINE_RULE_TYPE_LABELS = tuple(label for _key, label in INLINE_RULE_TYPE_OPTIONS)
INLINE_RULE_TYPE_BY_LABEL = {label: key for key, label in INLINE_RULE_TYPE_OPTIONS}
INLINE_RULE_TYPE_LABEL_BY_KEY = {key: label for key, label in INLINE_RULE_TYPE_OPTIONS}


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


@dataclass(frozen=True)
class StyleEntry:
    name: str
    table_style: bool = False
    caption_style: bool = False
    outline_markers: tuple[str, ...] = ()
    roles: tuple[str, ...] = ()
    numbering_group: str = ""
    numbering_level: int = 0
    restart_after_higher_level: bool = False
    template_file: str = ""


@dataclass(frozen=True)
class InlineRule:
    name: str
    rule_type: str
    target: str = "body"
    style_role: str = "body_bold"
    include_wrapper: bool = True
    remove_markers: bool = False
    enabled: bool = True


@dataclass(frozen=True)
class StyleSet:
    name: str
    paragraph_styles: list[StyleEntry]
    character_styles: list[StyleEntry]
    inline_rules: list[InlineRule] = field(default_factory=list)
    role_configs: dict[str, dict[str, object]] = field(default_factory=dict)


@dataclass
class NumberingRunState:
    seen_levels: dict[str, set[int]] = field(default_factory=dict)
    reset_levels: dict[str, set[int]] = field(default_factory=dict)
    continued_levels: dict[str, int] = field(default_factory=dict)
    last_levels: dict[str, int] = field(default_factory=dict)
    restarted: int = 0
    restart_failed: int = 0
    continued: int = 0
    continue_failed: int = 0


@dataclass
class HwpCandidate:
    display_name: str
    hwp: object
    score: int
    visible: bool
    path: str
    documents: str
    windows: str


@dataclass(frozen=True)
class UpdateInfo:
    latest_version: str
    download_url: str
    notes: str = ""
    html_url: str = ""
    asset_name: str = ""
    sha256: str = ""
    size: int = 0
    source: str = "custom"


@dataclass(frozen=True)
class ProofImagePage:
    path: Path
    width_px: int
    height_px: int
    rotated: bool = False


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

LEGACY_DEFAULT_CAPTION_PATTERN = r"^\[\s*(?P<kind>표|그림)\s+(?P<prefix>[^\]]*-)\s*(?:\d+)?\s*\]\s*(?P<title>.*)$"
LEGACY_DEFAULT_CAPTION_PREFIX_TEMPLATE = "[{kind} {prefix}"
NUMBERED_DEFAULT_CAPTION_PREFIX_TEMPLATE = "[{kind}{gap}{prefix}{number}"
DEFAULT_CAPTION_PATTERN = r"^\[\s*(?P<kind>표|그림)(?P<gap>\s*)(?P<prefix>(?:\d+(?:\.\d+)*\s*)?-)\s*(?P<number>\d+)?\s*\]\s*(?P<title>.*)$"
DEFAULT_CAPTION_PREFIX_TEMPLATE = "[{kind}{gap}{prefix}"
DEFAULT_CAPTION_SUFFIX_TEMPLATE = "] {title}"

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
        "cell_left": "2.0",
        "cell_right": "2.0",
        "cell_top": "2.0",
        "cell_bottom": "2.0",
        "outside_left": "0.0",
        "outside_right": "0.0",
        "outside_top": "0.0",
        "outside_bottom": "0.0",
    },
    "markdown_table": {
        "paragraph_style": "(적용 안 함)",
        "table_style_preset": "thin_top_bottom",
    },
    "caption_parser": {
        "pattern": DEFAULT_CAPTION_PATTERN,
        "prefix_template": DEFAULT_CAPTION_PREFIX_TEMPLATE,
        "suffix_template": DEFAULT_CAPTION_SUFFIX_TEMPLATE,
        "sample": "[그림 2.4-1] 학사관리 체계 개선 실적",
    },
    "title_number_box": DEFAULT_TITLE_NUMBER_BOX_SETTINGS,
}

TABLE_STYLE_ICON_PRESETS = (
    ("얇은전체", "thin_all", "얇은 전체 테두리"),
    ("굵은외곽", "thick_outer", "굵은 외곽선"),
    ("얇은상하", "thin_top_bottom", "얇은 상하선"),
    ("굵은상하", "thick_top_bottom", "굵은 상하선"),
    ("얇은상하이중선", "thin_top_bottom_header_double", "얇은 상하선 + 헤더 아래 이중선"),
    ("굵은상하이중선", "thick_top_bottom_header_double", "굵은 상하선 + 헤더 아래 이중선"),
)
TABLE_STYLE_PRESET_LABELS = {preset_name: icon_name for icon_name, preset_name, _tooltip in TABLE_STYLE_ICON_PRESETS}

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


def xml_local_name(name: str) -> str:
    return name.rsplit("}", 1)[-1] if "}" in name else name


def hwp_page_landscape_value(raw_value: str | None, width: int, height: int) -> int:
    if raw_value:
        normalized = raw_value.strip().upper()
        if normalized in {"1", "TRUE", "LANDSCAPE", "WIDE", "WIDELY"}:
            return 1 if width > height else 0
        if normalized in {"0", "FALSE", "PORTRAIT", "NARROWLY"}:
            return 0
    return 1 if width > height else 0


def hwp_gutter_type_value(raw_value: str | None) -> int:
    normalized = (raw_value or "").strip().upper()
    if normalized in {"BOTH_SIDES", "BOTH", "DOUBLE_SIDE", "1"}:
        return 1
    if normalized in {"TOP_ONLY", "TOP", "2"}:
        return 2
    return 0


def read_hwpx_page_settings(path: Path) -> PageSettings | None:
    if not path.exists():
        return None
    try:
        section_xml = read_zip_text(path, "Contents/section0.xml")
        root = ET.fromstring(section_xml)
    except Exception:
        return None

    page_pr = None
    for element in root.iter():
        if xml_local_name(element.tag) == "pagePr":
            page_pr = element
            break
    if page_pr is None:
        return None

    margin = next((child for child in page_pr if xml_local_name(child.tag) == "margin"), None)
    if margin is None:
        return None

    def int_attr(element, name: str, default: int = 0) -> int:
        try:
            return int(element.get(name) or default)
        except (TypeError, ValueError):
            return default

    width = int_attr(page_pr, "width")
    height = int_attr(page_pr, "height")
    return PageSettings(
        paper_width=width,
        paper_height=height,
        landscape=hwp_page_landscape_value(page_pr.get("landscape"), width, height),
        left_margin=int_attr(margin, "left"),
        right_margin=int_attr(margin, "right"),
        top_margin=int_attr(margin, "top"),
        bottom_margin=int_attr(margin, "bottom"),
        header_len=int_attr(margin, "header"),
        footer_len=int_attr(margin, "footer"),
        gutter_len=int_attr(margin, "gutter"),
        gutter_type=hwp_gutter_type_value(page_pr.get("gutterType")),
    )


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
        raw_type = element.get("Type") or element.get("type") or "PARA"
        style_type = str(raw_type).upper()
        if style_type not in {"PARA", "CHAR"}:
            style_type = "PARA"
        records.append(
            StyleRecord(
                style_id=style_id,
                style_index=style_index,
                style_type=style_type,
                name=unescape(name),
            )
        )
    return records


def register_hwp_file_path_checker(hwp) -> None:
    for module_name in ("FilePathCheckDLL", "FilePathCheckerDLL"):
        try:
            hwp.RegisterModule(module_name, "FilePathCheckerModule")
        except Exception:
            pass


def hwp_open_format(path: Path) -> str:
    suffix = path.suffix.lower()
    if suffix == ".hwpx":
        return "HWPX"
    if suffix == ".hwp":
        return "HWP"
    return ""


def close_import_hwp(hwp) -> None:
    try:
        hwp.HAction.Run("FileClose")
    except Exception:
        pass
    try:
        hwp.HAction.Run("FileQuit")
    except Exception:
        pass
    try:
        hwp.Quit()
    except Exception:
        pass


def read_style_records_from_hwp_file(path: Path) -> list[StyleRecord]:
    if not path.exists():
        raise FileNotFoundError(path)
    if path.suffix.lower() == ".hwpx":
        return read_style_records(path)
    if path.suffix.lower() != ".hwp":
        raise ValueError("HWPX 또는 HWP 파일만 불러올 수 있습니다.")
    if win32com is None or pythoncom is None:
        raise RuntimeError("HWP 파일을 불러오려면 pywin32와 한컴오피스 한글 COM이 필요합니다.")

    pythoncom.CoInitialize()
    hwp = win32com.client.Dispatch("HWPFrame.HwpObject")
    try:
        register_hwp_file_path_checker(hwp)
        ok = hwp.Open(str(path), hwp_open_format(path), "")
        if not ok:
            raise RuntimeError(f"한글에서 파일을 열지 못했습니다: {path}")
        for option in ("", "saveblock:true", "saveblock"):
            style_xml = safe_str(hwp.GetTextFile("HWPML2X", option))
            records = read_style_records_from_hwpml(style_xml)
            if records:
                return records
        return []
    finally:
        close_import_hwp(hwp)


def load_style_order() -> list[str]:
    ensure_user_config_file(STYLE_ORDER_FILE, BUNDLED_STYLE_ORDER_FILE)
    if not STYLE_ORDER_FILE.exists():
        return []
    try:
        data = read_json_file(STYLE_ORDER_FILE)
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
    write_json_file(STYLE_ORDER_FILE, data)


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


DEFAULT_INLINE_RULES = [
    InlineRule(
        name="첫머리 괄호 키워드",
        rule_type="leading_parenthesized_after_identifier",
        target="body",
        style_role="body_bold",
        include_wrapper=True,
        remove_markers=False,
        enabled=True,
    ),
    InlineRule(
        name="마크다운 굵게",
        rule_type="markdown_bold",
        target="body",
        style_role="body_bold",
        include_wrapper=True,
        remove_markers=True,
        enabled=True,
    ),
    InlineRule(
        name="첫머리 콜론 앞",
        rule_type="leading_text_before_colon_after_identifier",
        target="body",
        style_role="body_bold",
        include_wrapper=False,
        remove_markers=False,
        enabled=True,
    ),
    InlineRule(
        name="첫머리 괄호 키워드",
        rule_type="leading_parenthesized_after_identifier",
        target="table",
        style_role="table_bold",
        include_wrapper=True,
        remove_markers=False,
        enabled=True,
    ),
    InlineRule(
        name="마크다운 굵게",
        rule_type="markdown_bold",
        target="table",
        style_role="table_bold",
        include_wrapper=True,
        remove_markers=True,
        enabled=True,
    ),
    InlineRule(
        name="첫머리 콜론 앞",
        rule_type="leading_text_before_colon_after_identifier",
        target="table",
        style_role="table_bold",
        include_wrapper=False,
        remove_markers=False,
        enabled=True,
    ),
]


def unique_role_names(roles: Iterable[object]) -> list[str]:
    result: list[str] = []
    seen: set[str] = set()
    for role in roles:
        text = str(role).strip()
        if not text or text in seen:
            continue
        result.append(text)
        seen.add(text)
    return result


def parse_role_text(text: str) -> list[str]:
    return unique_role_names(part for part in text.split(","))


STYLE_ROLE_NONE_LABEL = "(없음)"
PARAGRAPH_STYLE_ROLE_OPTIONS: tuple[str, ...] = ("title_number_box",)
CHARACTER_STYLE_ROLE_OPTIONS: tuple[str, ...] = ("body_bold", "table_bold")


def allowed_style_roles(style_type: str) -> tuple[str, ...]:
    return PARAGRAPH_STYLE_ROLE_OPTIONS if style_type == "문단" else CHARACTER_STYLE_ROLE_OPTIONS


def normalize_style_role_value(value: object) -> str:
    text = str(value or "").strip()
    if not text or text == STYLE_ROLE_NONE_LABEL:
        return ""
    return text


def role_selection_values(style_type: str, current_role: object = "") -> list[str]:
    values = [STYLE_ROLE_NONE_LABEL, *allowed_style_roles(style_type)]
    current = normalize_style_role_value(current_role)
    if current and current not in values:
        values.append(current)
    return values


def style_roles_from_selection(value: object) -> tuple[str, ...]:
    role = normalize_style_role_value(value)
    return (role,) if role else ()


def primary_style_role(roles: Iterable[object]) -> str:
    names = unique_role_names(roles)
    return names[0] if names else ""


def normalize_config_path(value: object) -> str:
    text = str(value or "").strip()
    if not text:
        return ""
    path = Path(text)
    try:
        if path.is_absolute():
            return str(path.relative_to(CONFIG_ROOT)).replace("\\", "/")
    except ValueError:
        pass
    try:
        if path.is_absolute():
            return str(path.relative_to(ROOT)).replace("\\", "/")
    except ValueError:
        return str(path)
    return str(path).replace("\\", "/")


def resolve_config_path(value: object, base: Path = ROOT) -> Path:
    text = str(value or "").strip()
    if not text:
        return base
    path = Path(text)
    if path.is_absolute():
        return path
    parts = path.parts
    if parts and parts[0] == "templates":
        if len(parts) >= 2 and parts[1] in ("builtin", "custom"):
            return CONFIG_ROOT / path
        return BUILTIN_TEMPLATE_DIR.joinpath(*parts[1:])
    return base / path


def display_path(path: Path) -> str:
    for base in (CONFIG_ROOT, ROOT):
        try:
            return str(path.relative_to(base))
        except ValueError:
            continue
    return str(path)


def normalize_role_configs(item: object) -> dict[str, dict[str, object]]:
    if not isinstance(item, dict):
        return {}
    result: dict[str, dict[str, object]] = {}
    for raw_role, raw_config in item.items():
        role = str(raw_role or "").strip()
        if not role:
            continue
        config = raw_config if isinstance(raw_config, dict) else {}
        normalized: dict[str, object] = {}
        template_file = normalize_config_path(config.get("template_file", ""))
        if template_file:
            normalized["template_file"] = template_file
        if normalized:
            result[role] = normalized
    return result


def normalize_numbering_level(value: object) -> int:
    try:
        return max(0, int(str(value).strip() or "0"))
    except (TypeError, ValueError):
        return 0


def normalize_numbering_group(value: object) -> str:
    text = str(value or "").strip()
    if not text:
        return ""
    aliases = {
        "본문": "body",
        "body": "body",
        "표": "table",
        "table": "table",
        "주석": "note",
        "note": "note",
        "증빙": "proof",
        "proof": "proof",
        "별첨": "appendix",
        "appendix": "appendix",
        "기타1": "custom1",
        "custom1": "custom1",
        "기타2": "custom2",
        "custom2": "custom2",
    }
    return aliases.get(text, text)


def numbering_group_label(value: object) -> str:
    key = normalize_numbering_group(value)
    return NUMBERING_GROUP_LABEL_BY_KEY.get(key, key)


def normalize_inline_rule_type(value: object) -> str:
    text = str(value or "").strip()
    return INLINE_RULE_TYPE_BY_LABEL.get(text, text)


def inline_rule_type_label(value: object) -> str:
    key = normalize_inline_rule_type(value)
    return INLINE_RULE_TYPE_LABEL_BY_KEY.get(key, key)


def inferred_numbering_level(name: str) -> int:
    match = re.search(r"개요(?:번호)?[^\d]*(\d+)", normalize_style_name(name))
    return normalize_numbering_level(match.group(1)) if match else 0


def inferred_numbering_group(name: str, *, table_style: bool = False) -> str:
    level = inferred_numbering_level(name)
    if not level:
        return ""
    return "table" if table_style else "body"


def inferred_style_roles(name: str, *, table_style: bool = False) -> tuple[str, ...]:
    normalized = normalize_style_name(name)
    roles: list[str] = []
    if "굵게" in normalized:
        roles.append("table_bold" if table_style else "body_bold")
    return tuple(roles)


def normalize_style_entry(item: object) -> StyleEntry | None:
    if isinstance(item, StyleEntry):
        name = item.name.strip()
        if not name:
            return None
        roles = item.roles or inferred_style_roles(name, table_style=item.table_style)
        numbering_level = item.numbering_level or inferred_numbering_level(name)
        numbering_group = normalize_numbering_group(item.numbering_group) or inferred_numbering_group(name, table_style=item.table_style)
        return StyleEntry(
            name,
            table_style=item.table_style,
            caption_style=item.caption_style,
            outline_markers=tuple(unique_style_marker_names(item.outline_markers)),
            roles=tuple(unique_role_names(roles)),
            numbering_group=numbering_group,
            numbering_level=numbering_level,
            restart_after_higher_level=bool(item.restart_after_higher_level),
            template_file=normalize_config_path(item.template_file),
        )
    if isinstance(item, dict):
        name = str(item.get("name", "")).strip()
        if not name:
            return None
        table_style = bool(item.get("table_style", False))
        roles = unique_role_names(item.get("roles", inferred_style_roles(name, table_style=table_style)))
        numbering = item.get("numbering") if isinstance(item.get("numbering"), dict) else {}
        numbering_level = normalize_numbering_level(numbering.get("level", item.get("numbering_level", 0)))
        if not numbering_level:
            numbering_level = inferred_numbering_level(name)
        numbering_group = normalize_numbering_group(numbering.get("group", item.get("numbering_group", "")))
        if not numbering_group:
            numbering_group = inferred_numbering_group(name, table_style=table_style)
        restart_after_higher = numbering.get("restart_after_higher_level", item.get("restart_after_higher_level", None))
        if restart_after_higher is None:
            restart_after_higher = bool(numbering_level)
        return StyleEntry(
            name=name,
            table_style=table_style,
            caption_style=bool(item.get("caption_style", False)),
            outline_markers=tuple(unique_style_marker_names(item.get("outline_markers", []))),
            roles=tuple(roles),
            numbering_group=numbering_group,
            numbering_level=numbering_level,
            restart_after_higher_level=bool(restart_after_higher),
            template_file=normalize_config_path(item.get("template_file", "")),
        )
    name = str(item).strip()
    if not name:
        return None
    numbering_level = inferred_numbering_level(name)
    return StyleEntry(
        name=name,
        roles=inferred_style_roles(name),
        numbering_group=inferred_numbering_group(name),
        numbering_level=numbering_level,
        restart_after_higher_level=bool(numbering_level),
        template_file="",
    )


def normalize_inline_rule(item: object) -> InlineRule | None:
    if isinstance(item, InlineRule):
        name = item.name.strip()
        rule_type = item.rule_type.strip()
        style_role = item.style_role.strip()
        if not name or not rule_type or not style_role:
            return None
        return InlineRule(
            name=name,
            rule_type=rule_type,
            target=normalize_inline_rule_target(item.target),
            style_role=style_role,
            include_wrapper=bool(item.include_wrapper),
            remove_markers=bool(item.remove_markers),
            enabled=bool(item.enabled),
        )
    if not isinstance(item, dict):
        return None
    name = str(item.get("name", "")).strip()
    rule_type = str(item.get("type", item.get("rule_type", ""))).strip()
    style_role = str(item.get("style_role", "")).strip()
    if not name or not rule_type or not style_role:
        return None
    return InlineRule(
        name=name,
        rule_type=rule_type,
        target=normalize_inline_rule_target(str(item.get("target", "body"))),
        style_role=style_role,
        include_wrapper=bool(item.get("include_wrapper", True)),
        remove_markers=bool(item.get("remove_markers", False)),
        enabled=bool(item.get("enabled", True)),
    )


def normalize_inline_rule_target(target: str) -> str:
    value = target.strip().lower()
    if value in {"table", "표"}:
        return "table"
    if value in {"all", "both", "body_table", "본문+표", "전체"}:
        return "all"
    return "body"


def unique_inline_rules(items: Iterable[object]) -> list[InlineRule]:
    result: list[InlineRule] = []
    seen: set[tuple[str, str, str]] = set()
    for item in items:
        rule = normalize_inline_rule(item)
        if rule is None:
            continue
        key = (rule.name, rule.rule_type, rule.style_role)
        if key in seen:
            continue
        result.append(rule)
        seen.add(key)
    return result


def unique_style_marker_names(markers: Iterable[object]) -> list[str]:
    result: list[str] = []
    seen: set[str] = set()
    for marker in markers:
        text = str(marker).strip()
        if not text or text in seen:
            continue
        result.append(text)
        seen.add(text)
    return result


def parse_style_marker_text(text: str) -> list[str]:
    return unique_style_marker_names(part for part in text.split(","))


def unique_style_entries(items: Iterable[object]) -> list[StyleEntry]:
    result: list[StyleEntry] = []
    seen: set[str] = set()
    for item in items:
        entry = normalize_style_entry(item)
        if entry is None or entry.name in seen:
            continue
        result.append(entry)
        seen.add(entry.name)
    return result


def style_set_to_records(style_set: StyleSet) -> list[StyleRecord]:
    records: list[StyleRecord] = []
    for style_type, entries in (
        ("PARA", style_set.paragraph_styles),
        ("CHAR", style_set.character_styles),
    ):
        for index, entry in enumerate(entries):
            records.append(StyleRecord(style_id=index, style_index=index, style_type=style_type, name=entry.name))
    return records


def style_sets_from_records(records: list[StyleRecord]) -> list[StyleSet]:
    paragraph_styles = [
        style_entry_from_record(record)
        for record in records
        if record.style_type == "PARA"
    ]
    character_styles = [
        style_entry_from_record(record)
        for record in records
        if record.style_type == "CHAR"
    ]
    return [
        StyleSet("보고서용", unique_style_entries(paragraph_styles), unique_style_entries(character_styles), list(DEFAULT_INLINE_RULES), {}),
        StyleSet("일반보고용", [], []),
        StyleSet("공문용", [], []),
    ]


def style_entry_from_record(record: StyleRecord) -> StyleEntry:
    normalized = normalize_style_name(record.name)
    return StyleEntry(
        record.name,
        table_style=normalized.startswith(normalize_style_name("표내용-")),
        caption_style="캡션" in record.name,
        roles=inferred_style_roles(record.name, table_style=normalized.startswith(normalize_style_name("표내용-"))),
    )


def style_set_from_records(name: str, records: list[StyleRecord], source_path: Path | None = None) -> StyleSet:
    paragraph_styles = [style_entry_from_record(record) for record in records if record.style_type == "PARA"]
    character_styles = [style_entry_from_record(record) for record in records if record.style_type == "CHAR"]
    role_configs: dict[str, dict[str, object]] = {}
    normalized_source = normalize_config_path(source_path) if source_path is not None else ""
    has_title_role = any("title_number_box" in entry.roles for entry in paragraph_styles)
    has_title_style = any(normalize_style_name(entry.name) == normalize_style_name("대제목") for entry in paragraph_styles)
    if normalized_source and (has_title_role or has_title_style):
        role_configs["title_number_box"] = {"template_file": normalized_source}
        paragraph_styles = [
            StyleEntry(
                entry.name,
                table_style=entry.table_style,
                caption_style=entry.caption_style,
                outline_markers=entry.outline_markers,
                roles=entry.roles,
                numbering_group=entry.numbering_group,
                numbering_level=entry.numbering_level,
                restart_after_higher_level=entry.restart_after_higher_level,
                template_file=normalized_source if ("title_number_box" in entry.roles or normalize_style_name(entry.name) == normalize_style_name("대제목")) else entry.template_file,
            )
            for entry in paragraph_styles
        ]
    return StyleSet(
        name,
        unique_style_entries(paragraph_styles),
        unique_style_entries(character_styles),
        list(DEFAULT_INLINE_RULES),
        role_configs,
    )


def normalize_style_set_item(item: object) -> StyleSet | None:
    if not isinstance(item, dict):
        return None
    name = str(item.get("name", "")).strip()
    if not name:
        return None
    paragraph_styles = unique_style_entries(item.get("paragraph_styles", []))
    character_styles = unique_style_entries(item.get("character_styles", []))
    inline_rules = unique_inline_rules(item.get("inline_rules", DEFAULT_INLINE_RULES))
    role_configs = normalize_role_configs(item.get("role_configs", {}))
    return StyleSet(
        name=name,
        paragraph_styles=paragraph_styles,
        character_styles=character_styles,
        inline_rules=inline_rules,
        role_configs=role_configs,
    )


def migrate_title_number_box_roles(style_sets: list[StyleSet]) -> list[StyleSet]:
    migrated_sets: list[StyleSet] = []
    for style_set in style_sets:
        migrated_entries: list[StyleEntry] = []
        for entry in style_set.paragraph_styles:
            roles = tuple(entry.roles)
            if (
                entry.name == "대제목"
                and "대제목" in entry.outline_markers
                and "title_number_box" not in roles
            ):
                roles = (*roles, "title_number_box")
            migrated_entries.append(
                StyleEntry(
                    entry.name,
                    table_style=entry.table_style,
                    caption_style=entry.caption_style,
                    outline_markers=entry.outline_markers,
                    roles=roles,
                    numbering_group=entry.numbering_group,
                    numbering_level=entry.numbering_level,
                    restart_after_higher_level=entry.restart_after_higher_level,
                    template_file=entry.template_file,
                )
            )
        migrated_sets.append(
            StyleSet(
                style_set.name,
                migrated_entries,
                style_set.character_styles,
                style_set.inline_rules,
                style_set.role_configs,
            )
        )
    return migrated_sets


def migrate_title_number_box_role_configs(
    style_sets: list[StyleSet],
    template_file: str,
) -> tuple[list[StyleSet], bool]:
    migrated_sets: list[StyleSet] = []
    changed = False
    normalized_template = str(template_file or "").strip() or DEFAULT_TITLE_NUMBER_BOX_SETTINGS["template_file"]
    for style_set in style_sets:
        has_title_role = any("title_number_box" in entry.roles for entry in style_set.paragraph_styles)
        role_configs = normalize_role_configs(style_set.role_configs)
        if has_title_role:
            title_config = dict(role_configs.get("title_number_box", {}))
            if not str(title_config.get("template_file", "")).strip():
                title_config["template_file"] = normalized_template
                role_configs["title_number_box"] = title_config
                changed = True
        migrated_sets.append(
            StyleSet(
                style_set.name,
                style_set.paragraph_styles,
                style_set.character_styles,
                style_set.inline_rules,
                role_configs,
            )
        )
    return migrated_sets, changed


def default_style_sets() -> tuple[str, list[StyleSet]]:
    records = apply_saved_style_order(read_style_records())
    sets = style_sets_from_records(records)
    return sets[0].name, sets


def load_style_sets() -> tuple[str, list[StyleSet]]:
    default_active, default_sets = default_style_sets()
    ensure_user_config_file(STYLE_SETS_FILE, BUNDLED_STYLE_SETS_FILE)
    if not STYLE_SETS_FILE.exists():
        return default_active, default_sets
    try:
        data = read_json_file(STYLE_SETS_FILE)
    except Exception:
        return default_active, default_sets
    if not isinstance(data, dict):
        return default_active, default_sets
    raw_sets = data.get("style_sets", [])
    if not isinstance(raw_sets, list):
        return default_active, default_sets
    style_sets = [item for item in (normalize_style_set_item(raw) for raw in raw_sets) if item is not None]
    if not style_sets:
        return default_active, default_sets
    style_sets = migrate_title_number_box_roles(style_sets)
    active_name = str(data.get("active_set", "")).strip()
    if active_name not in {item.name for item in style_sets}:
        active_name = style_sets[0].name
    return active_name, style_sets


def save_style_sets(active_name: str, style_sets: list[StyleSet]) -> None:
    if not style_sets:
        _, style_sets = default_style_sets()
        active_name = style_sets[0].name
    if active_name not in {item.name for item in style_sets}:
        active_name = style_sets[0].name
    data = {
        "active_set": active_name,
        "style_sets": [
            {
                "name": item.name,
                "paragraph_styles": [
                    {
                        "name": entry.name,
                        "table_style": entry.table_style,
                        "caption_style": entry.caption_style,
                        "outline_markers": list(entry.outline_markers),
                        "roles": list(entry.roles),
                        "template_file": entry.template_file,
                        "numbering": {
                            "group": entry.numbering_group,
                            "level": entry.numbering_level,
                            "restart_after_higher_level": entry.restart_after_higher_level,
                        },
                    }
                    for entry in unique_style_entries(item.paragraph_styles)
                ],
                "character_styles": [
                    {
                        "name": entry.name,
                        "table_style": entry.table_style,
                        "caption_style": entry.caption_style,
                        "outline_markers": list(entry.outline_markers),
                        "roles": list(entry.roles),
                        "template_file": entry.template_file,
                        "numbering": {
                            "group": entry.numbering_group,
                            "level": entry.numbering_level,
                            "restart_after_higher_level": entry.restart_after_higher_level,
                        },
                    }
                    for entry in unique_style_entries(item.character_styles)
                ],
                "inline_rules": [
                    {
                        "name": rule.name,
                        "type": rule.rule_type,
                        "target": rule.target,
                        "style_role": rule.style_role,
                        "include_wrapper": rule.include_wrapper,
                        "remove_markers": rule.remove_markers,
                        "enabled": rule.enabled,
                    }
                    for rule in unique_inline_rules(item.inline_rules)
                ],
                "role_configs": normalize_role_configs(item.role_configs),
            }
            for item in style_sets
        ],
    }
    write_json_file(STYLE_SETS_FILE, data)


def merge_dict(default: dict, loaded: dict) -> dict:
    merged = dict(default)
    for key, value in loaded.items():
        if isinstance(value, dict) and isinstance(default.get(key), dict):
            merged[key] = merge_dict(default[key], value)
        else:
            merged[key] = value
    return merged


def load_update_settings() -> dict:
    ensure_user_config_file(UPDATE_SETTINGS_FILE, BUNDLED_UPDATE_SETTINGS_FILE)
    if not UPDATE_SETTINGS_FILE.exists():
        return merge_dict(DEFAULT_UPDATE_SETTINGS, {})
    try:
        data = read_json_file(UPDATE_SETTINGS_FILE)
    except Exception:
        return merge_dict(DEFAULT_UPDATE_SETTINGS, {})
    if not isinstance(data, dict):
        return merge_dict(DEFAULT_UPDATE_SETTINGS, {})
    return merge_dict(DEFAULT_UPDATE_SETTINGS, data)


def version_parts(version: str) -> tuple[int, ...]:
    text = str(version).strip().lstrip("vV")
    parts = re.findall(r"\d+", text)
    if not parts:
        return (0,)
    return tuple(int(part) for part in parts)


def is_newer_version(latest: str, current: str) -> bool:
    left = list(version_parts(latest))
    right = list(version_parts(current))
    length = max(len(left), len(right))
    left.extend([0] * (length - len(left)))
    right.extend([0] * (length - len(right)))
    return tuple(left) > tuple(right)


def google_drive_file_id(url: str) -> str:
    text = str(url).strip()
    patterns = (
        r"drive\.google\.com/file/d/([^/?#]+)",
        r"drive\.google\.com/open\?id=([^&#]+)",
        r"drive\.google\.com/uc\?[^#]*\bid=([^&#]+)",
        r"[?&]id=([^&#]+)",
    )
    for pattern in patterns:
        match = re.search(pattern, text)
        if match:
            return match.group(1)
    return ""


def google_drive_direct_download_url(url: str) -> str:
    file_id = google_drive_file_id(url)
    if not file_id:
        return str(url).strip()
    return f"https://drive.google.com/uc?export=download&id={file_id}"


def wildcard_match(pattern: str, text: str) -> bool:
    regex = "^" + re.escape(pattern).replace(r"\*", ".*").replace(r"\?", ".") + "$"
    return re.match(regex, text, flags=re.IGNORECASE) is not None


def github_api_latest_release_url(owner: str, repo: str) -> str:
    owner = str(owner).strip()
    repo = str(repo).strip()
    if not owner or not repo:
        raise ValueError("GitHub owner/repo 설정이 비어 있습니다.")
    return f"https://api.github.com/repos/{owner}/{repo}/releases/latest"


def github_asset_sha256(asset: dict) -> str:
    digest = str(asset.get("digest") or "").strip()
    if digest.lower().startswith("sha256:"):
        return digest.split(":", 1)[1].strip()
    return ""


def parse_github_release_info(data: dict, asset_pattern: str = "BUFS-HWP-Editor-v*.zip") -> UpdateInfo:
    latest_version = str(data.get("tag_name") or data.get("name") or "").strip()
    notes = str(data.get("body") or "").strip()
    html_url = str(data.get("html_url") or "").strip()
    assets = data.get("assets")
    if not isinstance(assets, list):
        assets = []
    zip_assets = [
        asset
        for asset in assets
        if isinstance(asset, dict)
        and str(asset.get("browser_download_url") or "").strip()
        and wildcard_match(asset_pattern, str(asset.get("name") or ""))
    ]
    if not zip_assets:
        raise ValueError(f"GitHub Release에서 업데이트 ZIP을 찾지 못했습니다: {asset_pattern}")
    asset = zip_assets[0]
    download_url = str(asset.get("browser_download_url") or "").strip()
    asset_name = str(asset.get("name") or "").strip()
    size = int(asset.get("size") or 0)
    if not latest_version:
        match = re.search(r"v?(\d+(?:\.\d+)+)", asset_name)
        latest_version = match.group(1) if match else ""
    if not latest_version:
        raise ValueError("GitHub Release 버전(tag_name)을 찾지 못했습니다.")
    return UpdateInfo(
        latest_version=latest_version,
        download_url=download_url,
        notes=notes,
        html_url=html_url,
        asset_name=asset_name,
        sha256=github_asset_sha256(asset),
        size=size,
        source="github",
    )


def parse_update_info(data: dict, fallback_download_url: str = "") -> UpdateInfo:
    latest_version = str(data.get("latest_version") or data.get("version") or "").strip()
    download_url = str(data.get("download_url") or fallback_download_url or "").strip()
    notes = str(data.get("notes") or data.get("release_notes") or "").strip()
    html_url = str(data.get("html_url") or data.get("release_url") or "").strip()
    sha256 = str(data.get("sha256") or "").strip()
    asset_name = str(data.get("asset_name") or Path(download_url).name or "").strip()
    if not latest_version:
        raise ValueError("latest_version 값이 없습니다.")
    if not download_url:
        raise ValueError("download_url 값이 없습니다.")
    return UpdateInfo(
        latest_version=latest_version,
        download_url=download_url,
        notes=notes,
        html_url=html_url,
        asset_name=asset_name,
        sha256=sha256,
        source="custom",
    )


def fetch_update_info(settings: dict) -> UpdateInfo:
    provider = str(settings.get("provider") or "").strip().lower()
    timeout = float(settings.get("timeout_seconds") or DEFAULT_UPDATE_SETTINGS["timeout_seconds"])
    if provider == "github":
        request_url = github_api_latest_release_url(settings.get("github_owner", ""), settings.get("github_repo", ""))
        request = urllib.request.Request(
            request_url,
            headers={
                "Accept": "application/vnd.github+json",
                "User-Agent": APP_NAME,
            },
        )
        with urllib.request.urlopen(request, timeout=timeout) as response:
            payload = response.read().decode("utf-8-sig")
        data = json.loads(payload)
        if not isinstance(data, dict):
            raise ValueError("GitHub Release 응답은 객체 형태여야 합니다.")
        return parse_github_release_info(data, str(settings.get("asset_pattern") or DEFAULT_UPDATE_SETTINGS["asset_pattern"]))

    version_url = str(settings.get("version_url") or "").strip()
    if not version_url:
        raise ValueError("update-settings.json에 version_url이 비어 있습니다.")
    request_url = google_drive_direct_download_url(version_url)
    with urllib.request.urlopen(request_url, timeout=timeout) as response:
        payload = response.read().decode("utf-8-sig")
    data = json.loads(payload)
    if not isinstance(data, dict):
        raise ValueError("버전 정보 JSON은 객체 형태여야 합니다.")
    return parse_update_info(data, str(settings.get("download_url") or ""))


def update_cache_dir() -> Path:
    root = CONFIG_ROOT / "updates"
    root.mkdir(parents=True, exist_ok=True)
    return root


def file_sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as file:
        for chunk in iter(lambda: file.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def safe_zip_member_path(name: str) -> bool:
    if not name or name.startswith(("/", "\\")):
        return False
    parts = Path(name.replace("\\", "/")).parts
    return all(part not in ("", ".", "..") for part in parts)


def validate_update_zip(zip_path: Path, expected_root_dir: str = "BUFS-HWP-Editor", expected_exe_name: str = "BUFS-HWP-Editor.exe") -> None:
    expected_root = str(expected_root_dir).strip().strip("/\\")
    expected_exe = str(expected_exe_name).strip()
    if not expected_root:
        raise ValueError("업데이트 ZIP 루트 폴더 설정이 비어 있습니다.")
    try:
        with zipfile.ZipFile(zip_path) as archive:
            names = archive.namelist()
    except zipfile.BadZipFile as exc:
        raise ValueError("업데이트 ZIP 파일이 손상되었습니다.") from exc
    if not names:
        raise ValueError("업데이트 ZIP 파일이 비어 있습니다.")
    normalized = [name.replace("\\", "/") for name in names]
    unsafe = [name for name in normalized if not safe_zip_member_path(name)]
    if unsafe:
        raise ValueError(f"업데이트 ZIP에 안전하지 않은 경로가 있습니다: {unsafe[0]}")
    prefix = expected_root + "/"
    outside = [name for name in normalized if name != expected_root and not name.startswith(prefix)]
    if outside:
        raise ValueError(f"업데이트 ZIP 최상위 폴더는 {expected_root} 하나여야 합니다.")
    required_exe = prefix + expected_exe
    if required_exe not in normalized:
        raise ValueError(f"업데이트 ZIP에서 실행 파일을 찾지 못했습니다: {required_exe}")
    if prefix + "_internal/" not in normalized and not any(name.startswith(prefix + "_internal/") for name in normalized):
        raise ValueError("업데이트 ZIP에서 _internal 폴더를 찾지 못했습니다.")


def update_zip_name(info: UpdateInfo) -> str:
    if info.asset_name.lower().endswith(".zip"):
        return info.asset_name
    version = str(info.latest_version).strip().lstrip("vV") or "latest"
    return f"{APP_NAME}-v{version}.zip"


def download_update_asset(info: UpdateInfo, timeout: float, progress: Callable[[int, int], None] | None = None) -> Path:
    destination = update_cache_dir() / update_zip_name(info)
    request_url = google_drive_direct_download_url(info.download_url) if "drive.google.com" in info.download_url else info.download_url
    request = urllib.request.Request(request_url, headers={"User-Agent": APP_NAME})
    with urllib.request.urlopen(request, timeout=timeout) as response, destination.open("wb") as file:
        total = int(response.headers.get("Content-Length") or info.size or 0)
        received = 0
        while True:
            chunk = response.read(1024 * 256)
            if not chunk:
                break
            file.write(chunk)
            received += len(chunk)
            if progress:
                progress(received, total)
    if info.sha256:
        actual = file_sha256(destination)
        if actual.lower() != info.sha256.lower():
            destination.unlink(missing_ok=True)
            raise ValueError("다운로드한 업데이트 파일의 SHA-256이 일치하지 않습니다.")
    return destination


def updater_source_path() -> Path:
    return INSTALL_ROOT / "BUFS-HWP-Updater.exe"


def prepare_updater_copy() -> Path:
    source = updater_source_path()
    if not source.exists():
        raise FileNotFoundError(f"업데이터 실행 파일을 찾지 못했습니다: {source}")
    target = update_cache_dir() / "BUFS-HWP-Updater.exe"
    shutil.copy2(source, target)
    return target


def launch_update_helper(zip_path: Path, info: UpdateInfo, settings: dict) -> None:
    updater = prepare_updater_copy()
    args = [
        str(updater),
        "--pid",
        str(os.getpid()),
        "--install-dir",
        str(INSTALL_ROOT),
        "--zip",
        str(zip_path),
        "--root-dir",
        str(settings.get("expected_root_dir") or DEFAULT_UPDATE_SETTINGS["expected_root_dir"]),
        "--exe-name",
        str(settings.get("expected_exe_name") or DEFAULT_UPDATE_SETTINGS["expected_exe_name"]),
        "--version",
        str(info.latest_version),
    ]
    subprocess.Popen(args, cwd=str(update_cache_dir()), close_fds=True)


def normalize_table_settings(settings: dict) -> dict:
    title = settings.get("title_cell")
    if isinstance(title, dict):
        for key in ("apply_hwp_header", "apply_cell_background", "apply_cell_borders"):
            title.pop(key, None)
    settings["line_presets"]["none"]["label"] = "선 없음"
    markdown_table = settings.get("markdown_table")
    if isinstance(markdown_table, dict):
        if markdown_table.get("table_style_preset") not in TABLE_STYLE_PRESET_LABELS:
            markdown_table["table_style_preset"] = DEFAULT_TABLE_SETTINGS["markdown_table"]["table_style_preset"]
    caption_parser = settings.get("caption_parser")
    if isinstance(caption_parser, dict):
        if str(caption_parser.get("pattern") or "") == LEGACY_DEFAULT_CAPTION_PATTERN:
            caption_parser["pattern"] = DEFAULT_CAPTION_PATTERN
        if str(caption_parser.get("prefix_template") or "") in {
            LEGACY_DEFAULT_CAPTION_PREFIX_TEMPLATE,
            NUMBERED_DEFAULT_CAPTION_PREFIX_TEMPLATE,
        }:
            caption_parser["prefix_template"] = DEFAULT_CAPTION_PREFIX_TEMPLATE
        if not caption_parser.get("suffix_template"):
            caption_parser["suffix_template"] = DEFAULT_CAPTION_SUFFIX_TEMPLATE
    return settings


def load_table_settings() -> dict:
    ensure_user_config_file(TABLE_SETTINGS_FILE, BUNDLED_TABLE_SETTINGS_FILE)
    if not TABLE_SETTINGS_FILE.exists():
        settings = merge_dict(DEFAULT_TABLE_SETTINGS, {})
        return normalize_table_settings(settings)
    try:
        data = read_json_file(TABLE_SETTINGS_FILE)
    except Exception:
        settings = merge_dict(DEFAULT_TABLE_SETTINGS, {})
        return normalize_table_settings(settings)
    if not isinstance(data, dict):
        settings = merge_dict(DEFAULT_TABLE_SETTINGS, {})
        return normalize_table_settings(settings)
    settings = merge_dict(DEFAULT_TABLE_SETTINGS, data)
    return normalize_table_settings(settings)


def save_table_settings(settings: dict) -> None:
    write_json_file(TABLE_SETTINGS_FILE, settings)


def normalize_special_chars(raw_entries) -> list[dict[str, str]]:
    entries: list[dict[str, str]] = []
    if not isinstance(raw_entries, list):
        return entries
    for raw_entry in raw_entries:
        if isinstance(raw_entry, str):
            text = raw_entry.strip()
            label = text
        elif isinstance(raw_entry, dict):
            text = str(raw_entry.get("text") or "").strip()
            label = str(raw_entry.get("label") or text).strip()
        else:
            continue
        if text:
            entries.append({"label": label or text, "text": text})
    return entries


def load_special_chars() -> list[dict[str, str]]:
    if not SPECIAL_CHARS_FILE.exists():
        return normalize_special_chars(DEFAULT_SPECIAL_CHARS)
    try:
        data = read_json_file(SPECIAL_CHARS_FILE)
    except (OSError, json.JSONDecodeError):
        return normalize_special_chars(DEFAULT_SPECIAL_CHARS)
    return normalize_special_chars(data)


def save_special_chars(entries: list[dict[str, str]]) -> None:
    write_json_file(SPECIAL_CHARS_FILE, normalize_special_chars(entries))


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


def today_ym_text(now: datetime | None = None) -> str:
    value = now or datetime.now()
    return f"{value.year}. {value.month}."


def resolve_cover_date(value: str, now: datetime | None = None) -> str:
    stripped = value.strip()
    return today_ym_text(now) if not stripped or stripped in {"{{오늘}}", "오늘"} else stripped


def split_cover_input_lines(value: str) -> list[str]:
    normalized = value.replace("\r\n", "\n").replace("\r", "\n").replace("\v", "\n").replace("\u2028", "\n")
    lines = [line.strip() for line in normalized.split("\n")]
    while lines and not lines[0]:
        lines.pop(0)
    while lines and not lines[-1]:
        lines.pop()
    return lines


def parse_labeled_cover_input(text: str, now: datetime | None = None) -> tuple[str, str, str] | None:
    fields: dict[str, str] = {}
    current_key: str | None = None
    key_map = {
        "제목": "title",
        "문서제목": "title",
        "부서": "department",
        "부서명": "department",
        "날짜": "date_text",
    }
    for raw_line in text.replace("\r\n", "\n").replace("\r", "\n").replace("\v", "\n").split("\n"):
        line = raw_line.strip()
        if not line:
            continue
        match = re.match(r"^([^:：]+)\s*[:：]\s*(.*)$", line)
        if match and match.group(1).strip() in key_map:
            current_key = key_map[match.group(1).strip()]
            fields[current_key] = match.group(2).strip()
        elif current_key:
            fields[current_key] = (fields[current_key] + "\n" + line).strip()
    if "title" not in fields and "department" not in fields and "date_text" not in fields:
        return None
    missing = [label for label, key in (("제목", "title"), ("부서", "department")) if not fields.get(key)]
    if missing:
        raise ValueError("표지 입력에서 다음 항목을 찾지 못했습니다: " + ", ".join(missing))
    return fields["title"], resolve_cover_date(fields.get("date_text", ""), now), fields["department"]


def parse_cover_input(text: str, now: datetime | None = None) -> tuple[str, str, str]:
    labeled = parse_labeled_cover_input(text, now)
    if labeled is not None:
        return labeled
    values = split_cover_input_lines(text)
    if len(values) < 2:
        raise ValueError("표지로 만들 제목, 부서명 또는 제목, 날짜, 부서명 줄을 선택하세요.")
    if len(values) == 2:
        return values[0], today_ym_text(now), values[1]
    date_text = resolve_cover_date(values[1], now)
    return values[0], date_text, "\n".join(values[2:])


def today_ymd_text(now: datetime | None = None) -> str:
    value = now or datetime.now()
    return f"{value.year}. {value.month}. {value.day}."


def resolve_report_template_date(value: str, now: datetime | None = None) -> str:
    stripped = value.strip()
    return today_ymd_text(now) if not stripped or stripped in {"{{오늘}}", "오늘"} else stripped


def parse_report_template_input(text: str, now: datetime | None = None) -> ReportTemplateFields:
    fields: dict[str, str] = {}
    current_key: str | None = None
    key_map = {
        "제목": "title",
        "문서제목": "title",
        "부서": "department",
        "부서명": "department",
        "날짜": "date_text",
        "주석": "note",
        "표제목": "table_title",
        "표 제목": "table_title",
        "표제목줄": "table_title_line",
        "표 제목줄": "table_title_line",
    }
    for raw_line in text.replace("\r\n", "\n").replace("\r", "\n").replace("\v", "\n").split("\n"):
        line = raw_line.strip()
        if not line:
            continue
        match = re.match(r"^([^:：]+)\s*[:：]\s*(.*)$", line)
        if match and match.group(1).strip() in key_map:
            current_key = key_map[match.group(1).strip()]
            fields[current_key] = match.group(2).strip()
        elif current_key:
            fields[current_key] = (fields[current_key] + "\n" + line).strip()

    missing = [
        label
        for label, key in (("제목", "title"), ("부서", "department"))
        if key not in fields or (key != "date_text" and not fields.get(key))
    ]
    if missing:
        raise ValueError("보고양식 입력에서 다음 항목을 찾지 못했습니다: " + ", ".join(missing))
    return ReportTemplateFields(
        title=fields["title"],
        department=fields["department"],
        date_text=resolve_report_template_date(fields.get("date_text", ""), now),
        note=fields.get("note", ""),
        table_title=fields.get("table_title"),
        table_title_line=fields.get("table_title_line"),
    )


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


def collapse_all_linesegarrays(section_xml: str) -> str:
    collapsed = (
        '<hp:linesegarray><hp:lineseg textpos="0" vertpos="0" vertsize="0" '
        'textheight="0" baseline="0" spacing="0" horzpos="0" horzsize="0" flags="393216"/></hp:linesegarray>'
    )
    return re.sub(r"<hp:linesegarray>.*?</hp:linesegarray>", collapsed, section_xml, flags=re.DOTALL)


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


def create_filled_general_report_template_file(
    fields: ReportTemplateFields,
    template: Path = GENERAL_REPORT_TEMPLATE_FILE,
) -> Path:
    if not template.exists():
        raise FileNotFoundError(f"보고양식 템플릿을 찾지 못했습니다: {template}")

    TEST_OUTPUT_DIR.mkdir(exist_ok=True)
    target = TEST_OUTPUT_DIR / f"general-report-{datetime.now().strftime('%Y%m%d-%H%M%S')}.hwpx"
    replacements = {
        "{{제목}}": fields.title,
        "{{문서제목}}": fields.title,
        "{{부서}}": fields.department,
        "{{부서명}}": fields.department,
        "{{날짜}}": fields.date_text,
        "{{주석}}": fields.note,
    }
    if fields.table_title is not None:
        replacements["{{표제목}}"] = fields.table_title
    if fields.table_title_line is not None:
        replacements["{{표제목줄}}"] = fields.table_title_line

    with zipfile.ZipFile(template, "r") as src, zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as dst:
        for info in src.infolist():
            data = src.read(info.filename)
            if info.filename.endswith(".xml"):
                text = data.decode("utf-8")
                for old, new in replacements.items():
                    text = text.replace(old, hwp_xml_text(new))
                data = text.encode("utf-8")
            elif info.filename == "Preview/PrvText.txt":
                text = data.decode("utf-8")
                for old, new in replacements.items():
                    text = text.replace(old, new)
                data = text.encode("utf-8")
            dst.writestr(info, data)

    return target


def normalize_title_box_text(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def strip_xml_encoding_declaration(text: str) -> str:
    return re.sub(r"^\s*<\?xml[^>]*\?>", "", text, count=1)


def element_text_content(element: ET.Element) -> str:
    return "".join(element.itertext()).strip()


def title_number_box_series_source(entry: StyleEntry | None = None, settings: dict | None = None) -> str:
    template_file = ""
    if entry is not None:
        template_file = normalize_config_path(entry.template_file)
    if not template_file and isinstance(settings, dict):
        template_file = normalize_config_path(settings.get("template_file", ""))
    if template_file:
        return template_file
    if entry is not None and entry.name.strip():
        return f"style:{entry.name.strip()}"
    return ""


def title_number_box_series_key(entry: StyleEntry | None = None, settings: dict | None = None) -> str:
    source = title_number_box_series_source(entry, settings)
    if not source:
        return ""
    return hashlib.sha1(source.encode("utf-8")).hexdigest()[:12]


def build_title_number_box_marker(series_key: str = "") -> str:
    key = str(series_key or "").strip()
    if not key:
        return TITLE_NUMBER_BOX_MARKER
    return f"{TITLE_NUMBER_BOX_MARKER}{TITLE_NUMBER_BOX_MARKER_SEPARATOR}{key}"


def title_number_box_numbers_from_hwpml(hwpml: str, marker: str | None = None) -> list[int]:
    try:
        root = ET.fromstring(strip_xml_encoding_declaration(hwpml))
    except Exception:
        return []

    numbers: list[int] = []
    target_marker = None if marker is None else str(marker).strip()
    for table in root.iter():
        table_tag = xml_local_name(table.tag).lower()
        if table_tag != "tbl" and table_tag != "table":
            continue
        table_text = element_text_content(table)
        if TITLE_NUMBER_BOX_MARKER not in table_text:
            continue
        if target_marker and target_marker not in table_text:
            continue
        cells = [child for child in table.iter() if xml_local_name(child.tag).lower() in {"tc", "cell"}]
        if len(cells) < 2:
            continue
        first_text = element_text_content(cells[0])
        if re.fullmatch(r"\d{1,3}", first_text):
            numbers.append(int(first_text))
    return numbers


def create_filled_title_number_box_file(
    title: str,
    number: int = 1,
    template: Path = TITLE_NUMBER_BOX_TEMPLATE_FILE,
    marker: str = TITLE_NUMBER_BOX_MARKER,
) -> Path:
    if not template.exists():
        raise FileNotFoundError(f"대제목 번호박스 템플릿을 찾지 못했습니다: {template}")

    title = normalize_title_box_text(title)
    number_text = str(max(1, int(number)))
    TEST_OUTPUT_DIR.mkdir(exist_ok=True)
    target = TEST_OUTPUT_DIR / f"title-number-box-{datetime.now().strftime('%Y%m%d-%H%M%S')}.hwpx"

    def replace_title_number_box_text(text: str, *, xml_mode: bool) -> str:
        title_value = hwp_xml_text(title) if xml_mode else title
        marker_value = hwp_xml_text(marker) if xml_mode else marker
        text = text.replace(TITLE_NUMBER_BOX_MARKER, marker_value)
        if TITLE_NUMBER_BOX_NUMBER_PLACEHOLDER in text:
            text = text.replace(TITLE_NUMBER_BOX_NUMBER_PLACEHOLDER, number_text)
        elif xml_mode:
            text = re.sub(r"(<[^>]*:t>)1(</[^>]*:t>)", rf"\g<1>{hwp_xml_text(number_text)}\2", text, count=1)
        else:
            text = re.sub(r"^\s*1\b", number_text, text, count=1)
        if title:
            if TITLE_NUMBER_BOX_TITLE_PLACEHOLDER in text:
                text = text.replace(TITLE_NUMBER_BOX_TITLE_PLACEHOLDER, title_value)
            elif xml_mode:
                text = text.replace(">대제목<", f">{title_value}<")
            else:
                text = text.replace("대제목", title)
        return text

    with zipfile.ZipFile(template, "r") as src, zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as dst:
        for info in src.infolist():
            data = src.read(info.filename)
            if info.filename.endswith(".xml"):
                text = replace_title_number_box_text(data.decode("utf-8"), xml_mode=True)
                if info.filename == "Contents/section0.xml":
                    text = collapse_all_linesegarrays(text)
                data = text.encode("utf-8")
            elif info.filename == "Preview/PrvText.txt":
                data = replace_title_number_box_text(data.decode("utf-8"), xml_mode=False).encode("utf-8")
            dst.writestr(info, data)

    return target


def create_title_number_box_hwpml(
    title: str,
    number: int = 1,
    template: Path = TITLE_NUMBER_BOX_TEMPLATE_FILE,
    marker: str = TITLE_NUMBER_BOX_MARKER,
) -> str:
    if not template.exists():
        raise FileNotFoundError(f"대제목 번호박스 템플릿을 찾지 못했습니다: {template}")

    title = normalize_title_box_text(title)
    number_text = str(max(1, int(number)))
    section_xml = read_zip_text(template, "Contents/section0.xml")
    section_xml = section_xml.replace(TITLE_NUMBER_BOX_MARKER, hwp_xml_text(marker))
    if TITLE_NUMBER_BOX_NUMBER_PLACEHOLDER in section_xml:
        section_xml = section_xml.replace(TITLE_NUMBER_BOX_NUMBER_PLACEHOLDER, number_text)
    else:
        section_xml = re.sub(
            r"(<[^>]*:t>)1(</[^>]*:t>)",
            rf"\g<1>{hwp_xml_text(number_text)}\2",
            section_xml,
            count=1,
        )
    if title:
        if TITLE_NUMBER_BOX_TITLE_PLACEHOLDER in section_xml:
            section_xml = section_xml.replace(TITLE_NUMBER_BOX_TITLE_PLACEHOLDER, hwp_xml_text(title))
        else:
            section_xml = section_xml.replace(">대제목<", f">{hwp_xml_text(title)}<")
    return collapse_all_linesegarrays(section_xml)


def style_entry_has_role(entry: StyleEntry, role: str) -> bool:
    return role in entry.roles


def normalize_proof_title_from_path(path: Path) -> str:
    title = path.stem.strip()
    title = re.sub(r"^\s*\d+(?:\.\d+)?[\s._-]+", "", title).strip()
    return title or path.stem


def proof_sort_key(path: Path) -> tuple[float, str]:
    match = re.match(r"^\s*(\d+(?:\.\d+)?)", path.stem)
    number = float(match.group(1)) if match else float("inf")
    return number, path.name.casefold()


def create_filled_proof_template_file(
    title: str,
    template: Path = PROOF_TEMPLATE_FILE,
) -> Path:
    if not template.exists():
        raise FileNotFoundError(f"증빙 양식 템플릿을 찾지 못했습니다: {template}")

    TEST_OUTPUT_DIR.mkdir(exist_ok=True)
    safe_digest = hashlib.sha1(title.encode("utf-8", errors="ignore")).hexdigest()[:10]
    target = TEST_OUTPUT_DIR / f"proof-template-{datetime.now().strftime('%Y%m%d-%H%M%S')}-{safe_digest}.hwpx"

    with zipfile.ZipFile(template, "r") as src, zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as dst:
        for info in src.infolist():
            data = src.read(info.filename)
            if info.filename.endswith(".xml"):
                text = data.decode("utf-8")
                text = text.replace(PROOF_TITLE_MARKER, hwp_xml_text(title))
                data = text.encode("utf-8")
            elif info.filename == "Preview/PrvText.txt":
                text = data.decode("utf-8")
                text = text.replace(PROOF_TITLE_MARKER, title)
                data = text.encode("utf-8")
            dst.writestr(info, data)

    return target


def hwpunit_to_mm(value: int | float) -> float:
    return float(value) / HWPUNIT_PER_MM


def fit_mm_to_cell(
    image_width_px: int,
    image_height_px: int,
    cell_width_hwp: int,
    cell_height_hwp: int,
) -> tuple[float, float]:
    return fit_mm_to_box(
        image_width_px,
        image_height_px,
        cell_width_hwp,
        cell_height_hwp,
        margin_x_hwp=PROOF_IMAGE_CELL_MARGIN_X,
        margin_y_hwp=PROOF_IMAGE_CELL_MARGIN_Y,
    )


def fit_mm_to_box(
    image_width_px: int,
    image_height_px: int,
    box_width_hwp: int,
    box_height_hwp: int,
    *,
    margin_x_hwp: int = 0,
    margin_y_hwp: int = 0,
) -> tuple[float, float]:
    available_width = max(1, box_width_hwp - (margin_x_hwp * 2))
    available_height = max(1, box_height_hwp - (margin_y_hwp * 2))
    width_mm = hwpunit_to_mm(available_width)
    height_mm = hwpunit_to_mm(available_height)
    scale = min(width_mm / max(1, image_width_px), height_mm / max(1, image_height_px))
    return image_width_px * scale, image_height_px * scale


def render_pdf_pages_to_jpegs(
    pdf_path: Path,
    output_dir: Path,
    *,
    dpi: int = PROOF_DEFAULT_DPI,
    quality: int = PROOF_JPEG_QUALITY,
    auto_rotate_landscape: bool = True,
) -> list[ProofImagePage]:
    try:
        import fitz  # type: ignore[import-not-found]
        from PIL import Image  # type: ignore[import-not-found]
    except Exception as exc:
        raise RuntimeError("증빙 PDF 렌더링에는 PyMuPDF와 Pillow가 필요합니다.") from exc

    output_dir.mkdir(parents=True, exist_ok=True)
    pages: list[ProofImagePage] = []
    zoom = dpi / 72
    matrix = fitz.Matrix(zoom, zoom)
    with fitz.open(str(pdf_path)) as doc:
        for index, page in enumerate(doc, start=1):
            pix = page.get_pixmap(matrix=matrix, alpha=False)
            image = Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")
            rotated = False
            if auto_rotate_landscape and image.width > image.height:
                image = image.rotate(90, expand=True)
                rotated = True
            target = output_dir / f"{pdf_path.stem}-p{index:03d}.jpg"
            image.save(target, "JPEG", quality=quality, optimize=True)
            pages.append(ProofImagePage(target, image.width, image.height, rotated))
    return pages


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


def find_markdown_table_block(paragraphs: list[str], start_index: int = 0) -> tuple[int, int, list[list[str]]] | None:
    if start_index < 0 or start_index >= len(paragraphs):
        return None
    candidate_lines: list[str] = []
    best_match: tuple[int, int, list[list[str]]] | None = None
    for index in range(start_index, len(paragraphs)):
        stripped = paragraphs[index].strip()
        if not stripped or not stripped.startswith("|"):
            break
        candidate_lines.append(stripped)
        rows = parse_markdown_table("\n".join(candidate_lines))
        if rows is not None:
            best_match = (start_index, index, rows)
    return best_match


def normalize_markdown_table_cell_text(text: str) -> str:
    normalized = re.sub(r"(?i)<br\s*/?>", "\n", text)
    normalized = normalized.replace(r"\n", "\n")
    return normalized.replace("\n", "\r\n")


def markdown_table_cell_html(text: str) -> str:
    parts = normalize_markdown_table_cell_text(text).replace("\r\n", "\n").split("\n")
    return "<br/>".join(escape(part) for part in parts)


def markdown_rows_to_html_table(rows: list[list[str]]) -> str:
    header, body = rows[0], rows[1:]
    parts = [
        '<table border="1" cellspacing="0" cellpadding="4" '
        'style="border-collapse:collapse;">',
        "<thead><tr>",
    ]
    parts.extend(f"<th>{markdown_table_cell_html(cell)}</th>" for cell in header)
    parts.append("</tr></thead><tbody>")
    for row in body:
        parts.append("<tr>")
        parts.extend(f"<td>{markdown_table_cell_html(cell)}</td>" for cell in row)
        parts.append("</tr>")
    parts.append("</tbody></table>")
    return "".join(parts)


def rows_to_tsv(rows: list[list[str]]) -> str:
    return "\r\n".join("\t".join(cell for cell in row) for row in rows)


@dataclass(frozen=True)
class CaptionParts:
    prefix: str
    suffix: str


@dataclass(frozen=True)
class ReportTemplateFields:
    title: str
    department: str
    date_text: str
    note: str = ""
    table_title: str | None = None
    table_title_line: str | None = None


@dataclass(frozen=True)
class PageSettings:
    paper_width: int
    paper_height: int
    landscape: int
    left_margin: int
    right_margin: int
    top_margin: int
    bottom_margin: int
    header_len: int
    footer_len: int
    gutter_len: int
    gutter_type: int


PAGE_SETTING_COMPARE_FIELDS = (
    "paper_width",
    "paper_height",
    "landscape",
    "left_margin",
    "right_margin",
    "top_margin",
    "bottom_margin",
    "header_len",
    "footer_len",
    "gutter_len",
    "gutter_type",
)


def format_page_settings(settings: PageSettings | None) -> str:
    if settings is None:
        return "none"
    return (
        f"paper={settings.paper_width}x{settings.paper_height}, "
        f"landscape={settings.landscape}, "
        f"margin=L{settings.left_margin} R{settings.right_margin} "
        f"T{settings.top_margin} B{settings.bottom_margin}, "
        f"header_footer=H{settings.header_len} F{settings.footer_len}, "
        f"gutter={settings.gutter_len} type={settings.gutter_type}"
    )


def page_settings_match(expected: PageSettings, actual: PageSettings | None) -> bool:
    if actual is None:
        return False
    return all(getattr(expected, field_name) == getattr(actual, field_name) for field_name in PAGE_SETTING_COMPARE_FIELDS)


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
    values = {name: (value or "") for name, value in groupdict.items()}
    values.update(
        {
            "kind": values.get("kind") or "표",
            "prefix": number_prefix,
            "title": title,
            "g1": match.group(1) if len(match.groups()) >= 1 else "",
            "g2": match.group(2) if len(match.groups()) >= 2 else "",
        }
    )
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


def configured_outline_prefix_length(text: str, marker: str) -> int:
    leading_match = re.match(rf"^{INVISIBLE_PREFIX}{HORIZONTAL_SPACE}*", text)
    start = 0 if leading_match is None else leading_match.end()
    if not text.startswith(marker, start):
        return 0
    end = start + len(marker)
    while end < len(text) and re.match(HORIZONTAL_SPACE, text[end]):
        end += 1
    return end


def find_outline_style_rule(
    text: str,
    style_set: StyleSet,
    *,
    in_table: bool | None = None,
) -> tuple[StyleEntry, int] | None:
    candidates: list[tuple[StyleEntry, str]] = []
    for entry in style_set.paragraph_styles:
        if in_table is True and not entry.table_style:
            continue
        if in_table is False and entry.table_style:
            continue
        for marker in entry.outline_markers:
            candidates.append((entry, marker))
    candidates.sort(key=lambda item: len(item[1]), reverse=True)
    for entry, marker in candidates:
        prefix_length = configured_outline_prefix_length(text, marker)
        if prefix_length:
            return entry, prefix_length
    return None


def consume_numbering_entry(state: NumberingRunState | None, entry: StyleEntry) -> bool:
    if state is None or not entry.numbering_group or entry.numbering_level <= 0:
        return False
    group = entry.numbering_group
    level = entry.numbering_level
    reset_levels = state.reset_levels.setdefault(group, set())
    seen_levels = state.seen_levels.setdefault(group, set())
    previous_level = state.last_levels.get(group, 0)
    restart = entry.restart_after_higher_level and (level in reset_levels or (previous_level > 0 and previous_level < level))
    reset_levels.discard(level)
    for seen_level in seen_levels:
        if seen_level > level:
            reset_levels.add(seen_level)
    seen_levels.add(level)
    state.last_levels[group] = level
    return restart


def find_inline_rule_ranges(text: str, rule: InlineRule) -> list[tuple[int, int, tuple[int, int] | None]]:
    if not rule.enabled:
        return []
    if rule.rule_type == "leading_parenthesized_after_identifier":
        leading_match = re.match(rf"^{INVISIBLE_PREFIX}{HORIZONTAL_SPACE}*", text)
        wrapper_start = 0 if leading_match is None else leading_match.end()
        if wrapper_start >= len(text) or text[wrapper_start] != "(":
            return []
        wrapper_end = text.find(")", wrapper_start + 1)
        if wrapper_end < 0:
            return []
        keyword = text[wrapper_start + 1:wrapper_end]
        if not keyword.strip():
            return []
        keyword_start = wrapper_start + 1 + len(keyword) - len(keyword.lstrip())
        keyword_end = wrapper_start + 1 + len(keyword.rstrip())
        if rule.include_wrapper:
            return [(wrapper_start, wrapper_end + 1, None)]
        return [(keyword_start, keyword_end, None)]
    if rule.rule_type == "leading_text_before_colon_after_identifier":
        match = re.match(rf"^{INVISIBLE_PREFIX}{HORIZONTAL_SPACE}*(?P<keyword>[^:：\n]*\S){HORIZONTAL_SPACE}*[:：]", text)
        if match is None:
            return []
        end = match.end() if rule.include_wrapper else match.start("keyword") + len(match.group("keyword").rstrip())
        return [(match.start("keyword"), end, None)]
    if rule.rule_type == "markdown_bold":
        ranges: list[tuple[int, int, tuple[int, int] | None]] = []
        for match in re.finditer(r"\*\*(?P<keyword>[^*]*\S[^*]*)\*\*", text):
            if rule.remove_markers:
                apply_range = (match.start("keyword"), match.end("keyword"))
            elif rule.include_wrapper:
                apply_range = (match.start(), match.end())
            else:
                apply_range = (match.start("keyword"), match.end("keyword"))
            marker_range = (match.start(), match.end()) if rule.remove_markers else None
            ranges.append((apply_range[0], apply_range[1], marker_range))
        return ranges
    return []


def inline_rule_search_offsets(text: str, rule: InlineRule, style_set: StyleSet) -> list[int]:
    if rule.rule_type not in {
        "leading_parenthesized_after_identifier",
        "leading_text_before_colon_after_identifier",
    }:
        return [0]
    offsets: list[int] = []
    for entry in style_set.paragraph_styles:
        for marker in entry.outline_markers:
            prefix_length = configured_outline_prefix_length(text, marker)
            if prefix_length and prefix_length not in offsets:
                offsets.append(prefix_length)
    manual_prefix_length = manual_outline_prefix_length(text)
    if manual_prefix_length and manual_prefix_length not in offsets:
        offsets.append(manual_prefix_length)
    offsets.append(0)
    return offsets


def find_inline_rule_ranges_for_style_set(
    text: str,
    rule: InlineRule,
    style_set: StyleSet,
) -> list[tuple[int, int, tuple[int, int] | None]]:
    for offset in inline_rule_search_offsets(text, rule, style_set):
        ranges = find_inline_rule_ranges(text[offset:], rule)
        if ranges:
            adjusted: list[tuple[int, int, tuple[int, int] | None]] = []
            for start, end, marker_range in ranges:
                adjusted_marker_range = None
                if marker_range is not None:
                    marker_start, marker_end = marker_range
                    adjusted_marker_range = (marker_start + offset, marker_end + offset)
                adjusted.append((start + offset, end + offset, adjusted_marker_range))
            return adjusted
    return []


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
        self.active_style_set_name, self.style_sets = load_style_sets()
        self.current_doc_style_path: Path | None = None
        self.current_doc_style_map: dict[str, StyleRecord] = {}
        self.current_doc_style_norm_map: dict[str, StyleRecord] = {}
        self.table_settings = load_table_settings()
        self.ensure_title_number_box_role_configs()
        self.update_settings = load_update_settings()
        self.special_chars = load_special_chars()
        self.palette = palette_from_settings(self.table_settings)
        self.style_set_var = tk.StringVar(value=self.active_style_set_name)
        self.apply_on_select = tk.BooleanVar(value=True)
        self.cover_confidential = tk.BooleanVar(value=False)
        self.logo_height_var = tk.StringVar(value="12")
        self.proof_dpi_var = tk.StringVar(value=str(PROOF_DEFAULT_DPI))
        self.proof_auto_rotate_var = tk.BooleanVar(value=True)
        self.decimal_places_var = tk.StringVar(value="0")
        self.decimal_mode_var = tk.StringVar(value="반올림")
        self.decimal_use_commas_var = tk.BooleanVar(value=True)
        self.unit_decimal_places_var = tk.StringVar(value="0")
        self.unit_decimal_mode_var = tk.StringVar(value="반올림")
        self.unit_decimal_use_commas_var = tk.BooleanVar(value=True)
        self.compose_number_var = tk.StringVar(value="")
        self.logo_chip_images: list[tk.PhotoImage] = []
        self.icon_images: dict[str, tk.PhotoImage] = {}
        self.pending_caption_title = ""
        self.pending_caption_parts = CaptionParts(prefix="", suffix="")
        self._refreshing = False
        self._applying_style = False

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=6, pady=(6, 4))

        self._build_styles_tab()
        self._build_table_tab()
        self._build_page_tab()
        self._build_special_chars_tab()
        self._build_cover_logo_tab()
        self._build_status_tab()
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.bind_all("<Control-z>", self.on_global_undo)
        self.bind_all("<Control-Z>", self.on_global_undo)
        self.refresh_all()
        self.after(700, self.warm_current_doc_style_cache)
        if self.update_settings.get("enabled") and self.update_settings.get("check_on_start"):
            self.after(1500, lambda: self.check_for_updates(silent=True))

    def create_scrollable_tab(self, title: str, *, padding: int = 8) -> ttk.Frame:
        outer = ttk.Frame(self.notebook)
        self.notebook.add(outer, text=title)
        outer.grid_rowconfigure(0, weight=1)
        outer.grid_columnconfigure(0, weight=1)

        canvas = tk.Canvas(outer, highlightthickness=0)
        scrollbar = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        content = ttk.Frame(canvas, padding=padding)
        content_window = canvas.create_window((0, 0), window=content, anchor="nw")

        def update_scroll_region(_event=None) -> None:
            canvas.configure(scrollregion=canvas.bbox("all"))
            needs_scroll = content.winfo_reqheight() > canvas.winfo_height()
            if needs_scroll:
                scrollbar.grid()
            else:
                scrollbar.grid_remove()
                canvas.yview_moveto(0)

        def fit_content_width(event) -> None:
            canvas.itemconfigure(content_window, width=event.width)
            self.after_idle(update_scroll_region)

        def scroll_content(event) -> str:
            delta = getattr(event, "delta", 0)
            event_num = getattr(event, "num", None)
            if event_num == 4:
                canvas.yview_scroll(-3, "units")
            elif event_num == 5:
                canvas.yview_scroll(3, "units")
            elif delta:
                canvas.yview_scroll(-1 * int(delta / 120), "units")
            return "break"

        def bind_scroll(_event) -> None:
            canvas.bind_all("<MouseWheel>", scroll_content)
            canvas.bind_all("<Button-4>", scroll_content)
            canvas.bind_all("<Button-5>", scroll_content)

        def unbind_scroll(_event) -> None:
            canvas.unbind_all("<MouseWheel>")
            canvas.unbind_all("<Button-4>")
            canvas.unbind_all("<Button-5>")

        content.bind("<Configure>", update_scroll_region)
        canvas.bind("<Configure>", fit_content_width)
        canvas.bind("<Enter>", bind_scroll)
        canvas.bind("<Leave>", unbind_scroll)
        outer.bind("<Destroy>", unbind_scroll, add="+")
        return content

    def _build_status_tab(self) -> None:
        frame = self.create_scrollable_tab("상태")

        buttons = ttk.Frame(frame)
        buttons.pack(fill="x", pady=(0, 8))
        for column in range(3):
            buttons.grid_columnconfigure(column, weight=1)
        status_buttons = (
            ("한글 다시 연결", self.check_hwp),
            ("대상 확인", self.show_current_target),
            ("테스트 문서 열기", self.open_style_test_copy),
            ("자산 다시 읽기", self.refresh_all),
            ("업데이트 확인", lambda: self.check_for_updates(silent=False)),
            ("실행취소", self.undo_hwp),
        )
        for index, (text, command) in enumerate(status_buttons):
            ttk.Button(buttons, text=text, command=command).grid(
                row=index // 3,
                column=index % 3,
                sticky="ew",
                padx=(0 if index % 3 == 0 else 6, 0),
                pady=(0 if index < 3 else 6, 0),
            )

        self.status_text = tk.Text(frame, height=30, wrap="word")
        self.status_text.pack(fill="both", expand=True)

    def _build_styles_tab(self) -> None:
        frame = self.create_scrollable_tab("스타일/정리")
        frame.grid_columnconfigure(0, weight=1)
        frame.grid_rowconfigure(0, weight=1)

        style_group = ttk.LabelFrame(frame, text="스타일 적용", padding=8)
        style_group.grid(row=0, column=0, sticky="nsew")

        set_row = ttk.Frame(style_group)
        set_row.pack(fill="x")
        ttk.Label(set_row, text="스타일 세트").pack(side="left")
        self.style_set_combo = ttk.Combobox(
            set_row,
            textvariable=self.style_set_var,
            values=[style_set.name for style_set in self.style_sets],
            state="readonly",
            width=18,
        )
        self.style_set_combo.pack(side="left", fill="x", expand=True, padx=(8, 0))
        self.style_set_combo.bind("<<ComboboxSelected>>", self.on_style_set_selected)
        ttk.Button(set_row, text="관리", command=self.open_style_sets_window).pack(side="left", padx=(8, 0))

        ttk.Label(style_group, text="선택한 세트의 스타일 이름").pack(anchor="w", pady=(8, 0))
        self.style_list = tk.Listbox(style_group, height=10)
        self.style_list.pack(fill="both", expand=True, pady=(8, 0))
        self.style_list.bind("<<ListboxSelect>>", self.on_style_selected)
        self.style_list.bind("<Double-Button-1>", lambda _event: self.apply_selected_style(reason="double-click"))

        ttk.Label(
            style_group,
            text="클릭하면 현재 문서의 같은 이름 스타일을 적용합니다.",
            wraplength=390,
        ).pack(anchor="w", pady=(8, 0))

        row3 = ttk.Frame(style_group)
        row3.pack(fill="x", pady=(5, 0))
        ttk.Button(row3, text="위로", command=lambda: self.move_selected_style(-1)).pack(side="left")
        ttk.Button(row3, text="아래로", command=lambda: self.move_selected_style(1)).pack(side="left", padx=(8, 0))
        ttk.Button(row3, text="프롬프트", command=self.copy_ai_prompt_for_active_style_set).pack(side="left", padx=(8, 0))
        ttk.Button(row3, text="글자스타일제거", command=self.clear_selected_character_style).pack(side="left", padx=(8, 0))

        paragraph_tab_frame = ttk.Frame(frame)
        paragraph_tab_frame.grid(row=1, column=0, sticky="ew")
        self.build_paragraph_tab_controls(paragraph_tab_frame)

        cleanup_frame = ttk.Frame(frame)
        cleanup_frame.grid(row=2, column=0, sticky="ew")
        self.build_cleanup_controls(cleanup_frame)

    def _build_table_tab(self) -> None:
        frame = ttk.Frame(self.notebook, padding=8)
        self.table_frame = frame
        self.notebook.add(frame, text="표 편집")
        self.build_table_tab_content()

    def _build_page_tab(self) -> None:
        frame = self.create_scrollable_tab("쪽 편집")

        hide_group = ttk.LabelFrame(frame, text="감추기", padding=8)
        hide_group.pack(fill="x")
        ttk.Button(
            hide_group,
            text="모두 감추기(머리말/꼬리말/쪽번호)",
            command=self.hide_current_page_header_footer_page_number,
        ).pack(fill="x")

        number_group = ttk.LabelFrame(frame, text="새 번호 개체", padding=8)
        number_group.pack(fill="x", pady=(12, 0))
        for column in range(3):
            number_group.grid_columnconfigure(column, weight=1)
        for column, (text, num_type) in enumerate((("새 쪽 번호", 0), ("새 표 번호", 4), ("새 그림 번호", 3))):
            ttk.Button(
                number_group,
                text=text,
                command=lambda value=num_type, label=text: self.insert_new_number_object(value, label),
            ).grid(row=0, column=column, sticky="ew", padx=(0, 6 if column < 2 else 0), pady=(0, 6))

        header_group = ttk.LabelFrame(frame, text="머리말 개체", padding=8)
        header_group.pack(fill="x", pady=(12, 0))
        for column in range(3):
            header_group.grid_columnconfigure(column, weight=1)
        for column, (text, header_type) in enumerate((("머리말 양쪽", 0), ("머리말 짝수쪽", 1), ("머리말 홀수쪽", 2))):
            ttk.Button(
                header_group,
                text=text,
                command=lambda value=header_type, label=text: self.insert_header_footer_object(0, value, label),
            ).grid(row=0, column=column, sticky="ew", padx=(0, 6 if column < 2 else 0), pady=(0, 6))

    def _build_special_chars_tab(self) -> None:
        frame = self.create_scrollable_tab("특수문자")
        frame.grid_columnconfigure(0, weight=1)

        compose_group = ttk.LabelFrame(frame, text="겹치기글자", padding=8)
        compose_group.grid(row=0, column=0, sticky="ew")
        compose_group.grid_columnconfigure(0, weight=1)

        compose_buttons_frame = ttk.Frame(compose_group)
        compose_buttons_frame.grid(row=0, column=0, sticky="ew")
        compose_buttons: list[tk.Canvas] = []
        for number in range(1, 11):
            compose_buttons.append(self.create_square_compose_button(
                compose_buttons_frame,
                chars=str(number),
                command=lambda value=number: self.insert_square_composed_number(value),
                row=0,
                column=number - 1,
                padx=(0 if number == 1 else 6, 0),
            ))

        self.bind_dynamic_square_button_grid(compose_buttons_frame, compose_buttons)

        custom_row = ttk.Frame(compose_group)
        custom_row.grid(row=1, column=0, sticky="ew", pady=(8, 0))
        custom_row.grid_columnconfigure(0, weight=1)
        custom_entry = ttk.Entry(custom_row, textvariable=self.compose_number_var, width=12)
        custom_entry.grid(row=0, column=0, sticky="ew")
        custom_entry.bind("<Return>", lambda _event: self.insert_custom_square_composed_number())
        ttk.Button(
            custom_row,
            text="넣기",
            command=self.insert_custom_square_composed_number,
        ).grid(row=0, column=1, sticky="ew", padx=(6, 0))

        checkbox_group = ttk.LabelFrame(frame, text="체크박스", padding=8)
        checkbox_group.grid(row=1, column=0, sticky="ew", pady=(12, 0))
        self.create_square_compose_button(
            checkbox_group,
            chars="☑",
            command=self.insert_checked_box_symbol,
            row=0,
            column=0,
            draw_box=False,
        )
        self.create_square_compose_button(
            checkbox_group,
            chars="✓",
            command=self.insert_square_composed_standard_check_mark,
            row=0,
            column=1,
            padx=(6, 0),
        )
        self.create_square_compose_button(
            checkbox_group,
            chars="✔",
            command=self.insert_square_composed_check_mark,
            row=0,
            column=2,
            padx=(6, 0),
        )

        next_row = 2
        for title, chars in DEFAULT_SPECIAL_CHAR_GROUPS:
            self.create_default_special_char_group(frame, title, chars, next_row)
            next_row += 1

        favorite_group = ttk.LabelFrame(frame, text="자주 쓰는 특수문자", padding=8)
        favorite_group.grid(row=next_row, column=0, sticky="ew", pady=(12, 0))
        favorite_group.grid_columnconfigure(0, weight=1)
        self.favorite_special_chars_frame = ttk.Frame(favorite_group)
        self.favorite_special_chars_frame.grid(row=0, column=0, sticky="ew")
        ttk.Button(
            favorite_group,
            text="관리",
            command=self.open_special_chars_window,
        ).grid(row=1, column=0, sticky="ew", pady=(8, 0))
        self.refresh_favorite_special_chars_buttons()

    def create_square_compose_button(
        self,
        parent: ttk.Frame,
        chars: str,
        command,
        row: int,
        column: int,
        padx=0,
        pady=0,
        draw_box: bool = True,
        font_family: str | None = None,
        font_weight: str | None = None,
        button_size: int = 24,
        font_size: int | None = None,
    ) -> tk.Canvas:
        canvas = tk.Canvas(
            parent,
            width=button_size,
            height=button_size,
            highlightthickness=1,
            highlightbackground="#9a9a9a",
            background="#ffffff",
            cursor="hand2",
        )
        canvas.grid(row=row, column=column, padx=padx, pady=pady)
        canvas.configure(takefocus=True)

        def draw(pressed: bool = False) -> None:
            canvas.delete("all")
            width = max(canvas.winfo_width(), button_size)
            height = max(canvas.winfo_height(), button_size)
            offset = 1 if pressed else 0
            if draw_box:
                box_size = min(width - 8, height - 8, 16)
                left = (width - box_size) / 2 + offset
                top = (height - box_size) / 2 + offset
                canvas.create_rectangle(
                    left,
                    top,
                    left + box_size,
                    top + box_size,
                    outline="#202020",
                    width=1,
                )
            effective_font_size = font_size if font_size is not None else (8 if len(chars) >= 2 else 12)
            text_options = {
                "text": chars,
                "fill": "#111111",
            }
            if font_family is not None:
                text_options["font"] = (font_family, effective_font_size, font_weight or "normal")
            canvas.create_text(width / 2 + offset, height / 2 + offset, **text_options)

        def release(event) -> None:
            draw(False)
            if 0 <= event.x <= canvas.winfo_width() and 0 <= event.y <= canvas.winfo_height():
                command()

        canvas.bind("<Configure>", lambda _event: draw(False))
        canvas.bind("<ButtonPress-1>", lambda _event: draw(True))
        canvas.bind("<ButtonRelease-1>", release)
        canvas.bind("<space>", lambda _event: command())
        canvas.bind("<Return>", lambda _event: command())
        draw(False)
        return canvas

    def bind_dynamic_square_button_grid(self, parent: ttk.Frame, buttons: list[tk.Canvas]) -> None:
        button_size = max((int(str(button.cget("width"))) for button in buttons), default=24)
        gap = 6
        right_safety_margin = 18

        def layout(event=None) -> None:
            width = event.width if event is not None else parent.winfo_width()
            available_width = max(button_size, width - right_safety_margin)
            columns = 1
            for candidate in range(1, len(buttons) + 1):
                required_width = candidate * button_size + (candidate - 1) * gap
                if required_width > available_width:
                    break
                columns = candidate
            for index, button in enumerate(buttons):
                row = index // columns
                column = index % columns
                button.grid_configure(
                    row=row,
                    column=column,
                    padx=(0 if column == 0 else gap, 0),
                    pady=(0 if row == 0 else gap, 0),
                )

        parent.bind("<Configure>", layout)
        parent.after_idle(layout)

    def create_default_special_char_group(
        self,
        parent: ttk.Frame,
        title: str,
        chars: tuple[str, ...],
        row: int,
    ) -> None:
        group = ttk.LabelFrame(parent, text=title, padding=8)
        group.grid(row=row, column=0, sticky="ew", pady=(12, 0))
        group.grid_columnconfigure(0, weight=1)
        buttons_frame = ttk.Frame(group)
        buttons_frame.grid(row=0, column=0, sticky="ew")
        buttons: list[tk.Canvas] = []
        for index, char in enumerate(chars):
            buttons.append(self.create_symbol_button(
                buttons_frame,
                char,
                row=0,
                column=index,
                padx=(0 if index == 0 else 6, 0),
            ))
        self.bind_dynamic_square_button_grid(buttons_frame, buttons)

    def create_symbol_button(
        self,
        parent: ttk.Frame,
        char: str,
        row: int,
        column: int,
        padx=0,
        pady=0,
    ) -> tk.Canvas:
        drawing = SPECIAL_CHAR_BUTTON_DRAWINGS.get(char)
        if drawing is not None:
            return self.create_drawn_special_char_button(parent, char, drawing, row, column, padx, pady)
        font_size = 9 if 0x246A <= ord(char) <= 0x2473 else None
        return self.create_square_compose_button(
            parent,
            chars=char,
            command=lambda value=char: self.insert_default_special_char(value),
            row=row,
            column=column,
            padx=padx,
            pady=pady,
            draw_box=False,
            button_size=28,
            font_size=font_size,
        )

    def create_drawn_special_char_button(
        self,
        parent: ttk.Frame,
        char: str,
        drawing: str,
        row: int,
        column: int,
        padx=0,
        pady=0,
    ) -> tk.Canvas:
        button_size = 28
        canvas = tk.Canvas(
            parent,
            width=button_size,
            height=button_size,
            highlightthickness=1,
            highlightbackground="#9a9a9a",
            background="#ffffff",
            cursor="hand2",
        )
        canvas.grid(row=row, column=column, padx=padx, pady=pady)
        canvas.configure(takefocus=True)

        def draw_arrow(width: int, height: int, offset: int) -> None:
            center_x = width / 2 + offset
            center_y = height / 2 + offset
            if drawing == "arrow-left":
                canvas.create_line(19 + offset, center_y, 9 + offset, center_y, fill="#111111", width=2)
                canvas.create_polygon(8 + offset, center_y, 14 + offset, center_y - 5, 14 + offset, center_y + 5, fill="#111111")
            elif drawing == "arrow-right":
                canvas.create_line(9 + offset, center_y, 19 + offset, center_y, fill="#111111", width=2)
                canvas.create_polygon(20 + offset, center_y, 14 + offset, center_y - 5, 14 + offset, center_y + 5, fill="#111111")
            elif drawing == "arrow-up":
                canvas.create_line(center_x, 19 + offset, center_x, 9 + offset, fill="#111111", width=2)
                canvas.create_polygon(center_x, 8 + offset, center_x - 5, 14 + offset, center_x + 5, 14 + offset, fill="#111111")
            elif drawing == "arrow-down":
                canvas.create_line(center_x, 9 + offset, center_x, 19 + offset, fill="#111111", width=2)
                canvas.create_polygon(center_x, 20 + offset, center_x - 5, 14 + offset, center_x + 5, 14 + offset, fill="#111111")

        def draw_shape(offset: int) -> None:
            if drawing == "shadow-circle":
                canvas.create_oval(12 + offset, 12 + offset, 21 + offset, 21 + offset, fill="#111111", outline="")
                canvas.create_oval(8 + offset, 8 + offset, 18 + offset, 18 + offset, fill="#ffffff", outline="#111111", width=1)
            elif drawing == "shadow-square":
                canvas.create_rectangle(12 + offset, 12 + offset, 21 + offset, 21 + offset, fill="#111111", outline="")
                canvas.create_rectangle(8 + offset, 8 + offset, 18 + offset, 18 + offset, fill="#ffffff", outline="#111111", width=1)

        def draw(pressed: bool = False) -> None:
            canvas.delete("all")
            width = max(canvas.winfo_width(), button_size)
            height = max(canvas.winfo_height(), button_size)
            offset = 1 if pressed else 0
            if drawing.startswith("arrow-"):
                draw_arrow(width, height, offset)
            else:
                draw_shape(offset)

        def release(event) -> None:
            draw(False)
            if 0 <= event.x <= canvas.winfo_width() and 0 <= event.y <= canvas.winfo_height():
                self.insert_default_special_char(char)

        canvas.bind("<Configure>", lambda _event: draw(False))
        canvas.bind("<ButtonPress-1>", lambda _event: draw(True))
        canvas.bind("<ButtonRelease-1>", release)
        canvas.bind("<space>", lambda _event: self.insert_default_special_char(char))
        canvas.bind("<Return>", lambda _event: self.insert_default_special_char(char))
        draw(False)
        return canvas

    def refresh_favorite_special_chars_buttons(self) -> None:
        frame = getattr(self, "favorite_special_chars_frame", None)
        if frame is None:
            return
        favorite_font = ("Malgun Gothic", 10)
        for child in frame.winfo_children():
            child.destroy()
        for index, entry in enumerate(self.special_chars):
            label = entry.get("label") or entry.get("text") or "?"
            text = entry.get("text") or label
            tk.Button(
                frame,
                text=label,
                font=favorite_font,
                bg="#ffffff",
                activebackground="#ffffff",
                relief="raised",
                bd=1,
                padx=8,
                command=lambda value=text: self.insert_favorite_special_char(value),
            ).grid(row=0, column=index, padx=(0 if index == 0 else 6, 0))

    def insert_default_special_char(self, text: str) -> None:
        self.insert_plain_special_char(text, f"기본 특수문자 {text}")

    def insert_favorite_special_char(self, text: str) -> None:
        self.insert_plain_special_char(text, f"특수문자 {text}")

    def insert_plain_special_char(self, text: str, label: str) -> None:
        if not self.ensure_hwp():
            return
        try:
            self.activate_hwp_window()
            executed = self.insert_hwp_text(text)
            self.log(f"{label}: result={executed}")
            if not executed:
                messagebox.showwarning(
                    label,
                    "한글이 특수문자 삽입을 실패로 반환했습니다. 입력 가능한 본문 위치에 커서를 둔 상태인지 확인하세요.",
                )
        except Exception as exc:
            self.log(f"{label} 실패: {type(exc).__name__}: {exc}")
            messagebox.showerror(f"{label} 실패", str(exc))

    def open_special_chars_window(self) -> None:
        window = tk.Toplevel(self)
        window.title("자주 쓰는 특수문자 관리")
        window.transient(self)
        window.grab_set()
        window.geometry("360x360")
        window.minsize(320, 320)

        entries = [dict(entry) for entry in normalize_special_chars(self.special_chars)]
        list_var = tk.StringVar()
        label_var = tk.StringVar()
        text_var = tk.StringVar()
        special_char_font = ("Malgun Gothic", 10)
        selected_entry_index: dict[str, int | None] = {"value": None}

        def display_text(entry: dict[str, str]) -> str:
            label = entry.get("label") or entry.get("text") or ""
            text = entry.get("text") or ""
            return label if label == text else f"{label}    {text}"

        def refresh_list(select_index: int | None = None) -> None:
            list_var.set(tuple(display_text(entry) for entry in entries))
            if select_index is not None and entries:
                index = max(0, min(select_index, len(entries) - 1))
                selected_entry_index["value"] = index
                listbox.selection_clear(0, "end")
                listbox.selection_set(index)
                listbox.see(index)
                load_selected()
            elif not entries:
                selected_entry_index["value"] = None

        def selected_index() -> int | None:
            selection = listbox.curselection()
            if selection:
                selected_entry_index["value"] = int(selection[0])
            return selected_entry_index["value"]

        def load_selected(_event=None) -> None:
            selection = listbox.curselection()
            if selection:
                selected_entry_index["value"] = int(selection[0])
            index = selected_entry_index["value"]
            if index is None or not (0 <= index < len(entries)):
                return
            label_var.set(entries[index].get("label", ""))
            text_var.set(entries[index].get("text", ""))

        def entry_from_form() -> dict[str, str] | None:
            text = text_var.get().strip()
            label = label_var.get().strip() or text
            if not text:
                messagebox.showwarning("자주 쓰는 특수문자", "삽입할 문자를 입력해 주세요.", parent=window)
                return None
            return {"label": label, "text": text}

        def start_new_entry() -> None:
            selected_entry_index["value"] = None
            listbox.selection_clear(0, "end")
            label_var.set("")
            text_var.set("")
            text_entry.focus_set()

        def register_entry() -> None:
            entry = entry_from_form()
            if entry is None:
                return
            entries.append(entry)
            refresh_list(len(entries) - 1)

        def update_entry() -> None:
            index = selected_index()
            if index is None:
                messagebox.showwarning("자주 쓰는 특수문자", "수정할 항목을 선택해 주세요.", parent=window)
                return
            entry = entry_from_form()
            if entry is None:
                return
            entries[index] = entry
            refresh_list(index)

        def delete_entry() -> None:
            index = selected_index()
            if index is None:
                return
            entries.pop(index)
            refresh_list(index)

        def move_entry(delta: int) -> None:
            index = selected_index()
            if index is None:
                return
            target = index + delta
            if not (0 <= target < len(entries)):
                return
            entries[index], entries[target] = entries[target], entries[index]
            refresh_list(target)

        def save_and_close() -> None:
            normalized = normalize_special_chars(entries)
            save_special_chars(normalized)
            self.special_chars = normalized
            self.refresh_favorite_special_chars_buttons()
            self.log(f"자주 쓰는 특수문자 저장: {len(normalized)}개, file={SPECIAL_CHARS_FILE}")
            window.destroy()

        body = ttk.Frame(window, padding=8)
        body.pack(fill="both", expand=True)
        body.grid_columnconfigure(0, weight=1)
        body.grid_rowconfigure(0, weight=1)

        listbox = tk.Listbox(
            body,
            listvariable=list_var,
            height=8,
            font=special_char_font,
            exportselection=False,
        )
        listbox.grid(row=0, column=0, columnspan=4, sticky="nsew")
        listbox.bind("<<ListboxSelect>>", load_selected)

        ttk.Label(body, text="버튼 이름").grid(row=1, column=0, sticky="w", pady=(8, 0))
        tk.Entry(body, textvariable=label_var, font=special_char_font).grid(
            row=1,
            column=1,
            columnspan=3,
            sticky="ew",
            pady=(8, 0),
        )
        ttk.Label(body, text="삽입 문자").grid(row=2, column=0, sticky="w", pady=(6, 0))
        text_entry = tk.Entry(body, textvariable=text_var, font=special_char_font)
        text_entry.grid(row=2, column=1, columnspan=3, sticky="ew", pady=(6, 0))
        text_entry.bind("<Return>", lambda _event: register_entry() if selected_entry_index["value"] is None else update_entry())

        ttk.Button(body, text="추가", command=start_new_entry).grid(row=3, column=0, sticky="ew", pady=(8, 0))
        ttk.Button(body, text="등록", command=register_entry).grid(row=3, column=1, sticky="ew", padx=(6, 0), pady=(8, 0))
        ttk.Button(body, text="수정", command=update_entry).grid(row=3, column=2, sticky="ew", padx=(6, 0), pady=(8, 0))
        ttk.Button(body, text="삭제", command=delete_entry).grid(row=3, column=3, sticky="ew", padx=(6, 0), pady=(8, 0))
        ttk.Button(body, text="위", command=lambda: move_entry(-1)).grid(row=4, column=0, sticky="ew", pady=(6, 0))
        ttk.Button(body, text="아래", command=lambda: move_entry(1)).grid(row=4, column=1, sticky="ew", padx=(6, 0), pady=(6, 0))

        footer = ttk.Frame(body)
        footer.grid(row=5, column=0, columnspan=4, sticky="ew", pady=(12, 0))
        footer.grid_columnconfigure(0, weight=1)
        ttk.Button(footer, text="저장", command=save_and_close).grid(row=0, column=1)
        ttk.Button(footer, text="취소", command=window.destroy).grid(row=0, column=2, padx=(6, 0))

        refresh_list(0 if entries else None)
        text_entry.focus_set()

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
        command_factory=None,
    ) -> None:
        for index, (icon_name, preset_name, tooltip) in enumerate(presets):
            row, col = divmod(index, columns)
            image = self.load_icon_image(icon_name)
            command = (
                command_factory(preset_name)
                if command_factory is not None
                else lambda name=preset_name: self.apply_table_border_preset(name)
            )
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

        default_font = tkfont.nametofont("TkDefaultFont")
        section_label_font = default_font.copy()
        section_label_font.configure(size=max(1, int(round(default_font.cget("size") * 1.5))))
        title_button_font = default_font.copy()
        title_button_font.configure(weight="bold")
        self.table_section_label_font = section_label_font
        self.table_title_button_font = title_button_font

        bottom_frame = ttk.Frame(frame)
        bottom_frame.pack(side="bottom", fill="x")
        ttk.Button(bottom_frame, text="markdown 표 → 한글 표", command=self.convert_selected_markdown_table).pack(
            fill="x", pady=(0, 6), ipady=8
        )
        ttk.Button(bottom_frame, text="표 설정", command=self.open_table_settings_window).pack(fill="x", ipady=8)

        scroll_host = ttk.Frame(frame)
        scroll_host.pack(side="top", fill="both", expand=True)

        canvas = tk.Canvas(scroll_host, highlightthickness=0)
        scrollbar = ttk.Scrollbar(scroll_host, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        content = ttk.Frame(canvas)
        content_window = canvas.create_window((0, 0), window=content, anchor="nw")

        def update_scroll_region(_event=None) -> None:
            canvas.configure(scrollregion=canvas.bbox("all"))
            needs_scroll = content.winfo_reqheight() > canvas.winfo_height()
            if needs_scroll:
                scrollbar.pack(side="right", fill="y")
            else:
                scrollbar.pack_forget()
                canvas.yview_moveto(0)

        def fit_content_width(event) -> None:
            canvas.itemconfigure(content_window, width=event.width)
            self.after_idle(update_scroll_region)

        def on_mousewheel(event) -> str:
            canvas.yview_scroll(-1 * int(event.delta / 120), "units")
            return "break"

        content.bind("<Configure>", update_scroll_region)
        canvas.bind("<Configure>", fit_content_width)
        canvas.bind("<Enter>", lambda _event: canvas.bind_all("<MouseWheel>", on_mousewheel))
        canvas.bind("<Leave>", lambda _event: canvas.unbind_all("<MouseWheel>"))

        def section_label(text: str) -> None:
            ttk.Label(content, text=text, font=section_label_font).pack(anchor="w")

        section_label("셀 배경색")
        bg_row = ttk.Frame(content)
        bg_row.pack(fill="x", pady=(6, 12))
        self.create_cell_fill_chips(bg_row)

        section_label("글자색")
        text_row = ttk.Frame(content)
        text_row.pack(fill="x", pady=(6, 12))
        self.create_color_chips(text_row, self.apply_text_color)

        section_label("여백/캡션")
        table_action_row = ttk.Frame(content)
        table_action_row.pack(fill="x", pady=(6, 12))
        ttk.Button(table_action_row, text="셀여백", command=self.apply_table_cell_margins).pack(
            side="left", fill="x", expand=True
        )
        ttk.Button(table_action_row, text="표여백", command=self.apply_table_outside_margins).pack(
            side="left", fill="x", expand=True, padx=(6, 0)
        )
        ttk.Button(table_action_row, text="너비맞추기", command=self.apply_table_page_width).pack(
            side="left", fill="x", expand=True, padx=(6, 0)
        )
        ttk.Button(table_action_row, text="캡션원클릭", command=self.cut_title_and_apply_to_next_table_caption).pack(
            side="left", fill="x", expand=True, padx=(6, 0)
        )
        for button in table_action_row.winfo_children():
            button.pack_configure(ipady=8)

        section_label("번호종류")
        numbering_row = ttk.Frame(content)
        numbering_row.pack(fill="x", pady=(6, 12))
        for text, numbering_type in (("없음", 0), ("표", 2), ("그림", 1)):
            ttk.Button(
                numbering_row,
                text=text,
                command=lambda value=numbering_type, label=text: self.apply_selected_object_numbering_type(value, label),
            ).pack(side="left", fill="x", expand=True, padx=(0 if numbering_type == 0 else 6, 0))

        section_label("표내용 스타일")
        table_style_buttons = ttk.Frame(content)
        table_style_buttons.pack(fill="x", pady=(6, 12))
        table_content_styles = self.table_content_style_records()
        if table_content_styles:
            for index, record in enumerate(table_content_styles):
                row, col = divmod(index, 4)
                ttk.Button(
                    table_style_buttons,
                    text=self.table_content_style_button_label(record),
                    command=lambda name=record.name: self.apply_table_content_style(name),
                ).grid(
                    row=row,
                    column=col,
                    padx=(0 if col == 0 else 6, 0),
                    pady=(0 if row == 0 else 6, 0),
                    sticky="ew",
                    ipady=8,
                )
            clear_index = len(table_content_styles)
            row, col = divmod(clear_index, 4)
            ttk.Button(
                table_style_buttons,
                text="글자스타일제거",
                command=self.clear_selected_character_style,
            ).grid(
                row=row,
                column=col,
                padx=(0 if col == 0 else 6, 0),
                pady=(0 if row == 0 else 6, 0),
                sticky="ew",
                ipady=8,
            )
            for col in range(4):
                table_style_buttons.columnconfigure(col, weight=1)
        else:
            ttk.Label(table_style_buttons, text="표 서식 체크 스타일 없음").pack(anchor="w")
            ttk.Button(
                table_style_buttons,
                text="글자스타일제거",
                command=self.clear_selected_character_style,
            ).pack(fill="x", pady=(6, 0), ipady=8)

        section_label("표 스타일")
        table_style_icons = ttk.Frame(content)
        table_style_icons.pack(fill="x", pady=(6, 12))
        self.create_icon_preset_grid(
            table_style_icons,
            TABLE_STYLE_ICON_PRESETS,
            columns=6,
            button_size=TABLE_STYLE_ICON_BUTTON_SIZE,
            command_factory=lambda name: lambda name=name: self.apply_table_style_preset(name),
        )

        section_label("테두리")
        border_icons = ttk.Frame(content)
        border_icons.pack(fill="x", pady=(6, 12))
        self.create_icon_preset_grid(
            border_icons,
            TABLE_BORDER_ICON_PRESETS,
            columns=6,
            button_size=TABLE_BORDER_ICON_BUTTON_SIZE,
        )

        title_bg = self.palette_color_or_rgb("셀배경", (220, 221, 221)).hex_value
        title_button = tk.Button(
            content,
            text="제목셀 자동화",
            command=self.apply_title_cell_preset,
            bg=title_bg,
            activebackground=title_bg,
            fg="#000000",
            font=title_button_font,
            relief="raised",
            bd=1,
        )
        title_button.pack(fill="x", pady=(0, 6), ipady=8)

    def table_content_style_records(self) -> list[StyleRecord]:
        prefix = normalize_style_name("표내용-")
        active_set = self.active_style_set()
        table_style_names = {
            entry.name
            for entry in active_set.paragraph_styles + active_set.character_styles
            if entry.table_style
        }
        return [
            record
            for record in self.style_records
            if record.style_type in {"PARA", "CHAR"}
            and (record.name in table_style_names or normalize_style_name(record.name).startswith(prefix))
        ]

    def table_content_style_button_label(self, record: StyleRecord) -> str:
        name = normalize_style_name(record.name)
        prefix = normalize_style_name("표내용-")
        if name.startswith(prefix):
            name = name[len(prefix) :]
        return name

    def caption_style_record(self) -> StyleRecord | None:
        active_set = self.active_style_set()
        caption_style_names = {
            entry.name
            for entry in active_set.paragraph_styles + active_set.character_styles
            if entry.caption_style
        }
        if caption_style_names:
            return next((record for record in self.style_records if record.name in caption_style_names), None)
        return None

    def apply_table_content_style(self, style_name: str) -> None:
        if not self.ensure_hwp():
            return
        ok = self.apply_style_by_name(style_name)
        self.log(f"표내용 스타일 적용: {style_name}, result={ok}")

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
        table_style_choices = [f"{preset}: {label}" for preset, label in TABLE_STYLE_PRESET_LABELS.items()]

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
        ttk.Label(body, text="Markdown 표").pack(anchor="w")
        markdown_table = settings["markdown_table"]
        row = ttk.Frame(body)
        row.pack(fill="x", pady=(6, 0))
        ttk.Label(row, text="기본 문단 스타일", width=14).pack(side="left")
        combo(
            row,
            ("markdown_table", "paragraph_style"),
            markdown_table.get("paragraph_style", "(적용 안 함)"),
            ["(적용 안 함)", *para_styles],
            width=22,
        )
        current_table_style = markdown_table.get("table_style_preset", "thin_top_bottom")
        row = ttk.Frame(body)
        row.pack(fill="x", pady=(6, 0))
        ttk.Label(row, text="표 스타일", width=14).pack(side="left")
        combo(
            row,
            ("markdown_table", "table_style_preset"),
            f"{current_table_style}: {TABLE_STYLE_PRESET_LABELS.get(current_table_style, current_table_style)}",
            table_style_choices,
            width=28,
        )

        ttk.Separator(body).pack(fill="x", pady=12)
        ttk.Label(body, text="셀 안 여백(mm)").pack(anchor="w")
        margins = settings["table_margins"]
        cell_margin_labels = [
            ("cell_left", "셀 안 좌"),
            ("cell_right", "셀 안 우"),
            ("cell_top", "셀 안 상"),
            ("cell_bottom", "셀 안 하"),
        ]
        for index in range(0, len(cell_margin_labels), 2):
            row = ttk.Frame(body)
            row.pack(fill="x", pady=(6, 0))
            for key, label in cell_margin_labels[index : index + 2]:
                ttk.Label(row, text=label, width=12).pack(side="left")
                entry(row, ("table_margins", key), str(margins.get(key, "0.0")), width=7)

        ttk.Label(body, text="표 바깥 여백(mm)").pack(anchor="w", pady=(10, 0))
        outside_margin_labels = [
            ("outside_left", "표 바깥 좌"),
            ("outside_right", "표 바깥 우"),
            ("outside_top", "표 바깥 상"),
            ("outside_bottom", "표 바깥 하"),
        ]
        for index in range(0, len(outside_margin_labels), 2):
            row = ttk.Frame(body)
            row.pack(fill="x", pady=(6, 0))
            for key, label in outside_margin_labels[index : index + 2]:
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
            if path == ("markdown_table", "table_style_preset"):
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
                        messagebox.showwarning("여백 값 확인", "셀/표 여백은 mm 단위 숫자로 입력하세요.", parent=window)
                        return
                    if number < 0:
                        messagebox.showwarning("여백 값 확인", "셀/표 여백은 0 이상으로 입력하세요.", parent=window)
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

    def build_paragraph_tab_controls(self, parent: ttk.Frame) -> None:
        tab_group = ttk.LabelFrame(parent, text="문단 탭", padding=8)
        tab_group.pack(fill="x", pady=(8, 0))
        tab_group.grid_columnconfigure(0, weight=1, uniform="paragraph-tab")
        tab_group.grid_columnconfigure(1, weight=1, uniform="paragraph-tab")
        ttk.Button(
            tab_group,
            text="오른쪽 자동탭",
            command=self.apply_auto_right_paragraph_tab,
        ).grid(row=0, column=0, sticky="ew", padx=(0, 6), ipady=8)
        ttk.Button(
            tab_group,
            text="점선 오른쪽탭",
            command=self.apply_dotted_right_paragraph_tab,
        ).grid(row=0, column=1, sticky="ew", ipady=8)

    def build_cleanup_controls(self, parent: ttk.Frame) -> None:
        cleanup_group = ttk.LabelFrame(parent, text="문장 정리", padding=8)
        cleanup_group.pack(fill="x", pady=(8, 0))

        sentence_buttons = ttk.Frame(cleanup_group)
        sentence_buttons.pack(fill="x")
        sentence_buttons.grid_columnconfigure(0, weight=1, uniform="sentence-half")
        sentence_buttons.grid_columnconfigure(1, weight=1, uniform="sentence-half")

        main_buttons = ttk.Frame(sentence_buttons)
        main_buttons.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        main_buttons.grid_columnconfigure(0, weight=1)
        for row in range(2):
            main_buttons.grid_rowconfigure(row, weight=1)

        bulk_font = tkfont.nametofont("TkDefaultFont").copy()
        bulk_font.configure(weight="bold")
        tk.Button(
            main_buttons,
            text="일괄처리",
            command=self.apply_configured_outline_styles_to_selection,
            bg="#4a4a4a",
            fg="#ffffff",
            activebackground="#333333",
            activeforeground="#ffffff",
            font=bulk_font,
            relief="raised",
            bd=1,
        ).grid(row=0, column=0, sticky="nsew", ipady=8)
        ttk.Button(main_buttons, text="「 」 씌우기", command=self.wrap_regulation_names_in_selection).grid(
            row=1,
            column=0,
            sticky="nsew",
            pady=(6, 0),
            ipady=8,
        )

        right_buttons = ttk.Frame(sentence_buttons)
        right_buttons.grid(row=0, column=1, sticky="nsew")
        for column in range(2):
            right_buttons.grid_columnconfigure(column, weight=1, uniform="sentence-right")
        for row in range(2):
            right_buttons.grid_rowconfigure(row, weight=1)

        normal_defs = (
            ("줄바꿈정리", self.clean_selected_line_breaks),
            ("기호제거", self.clean_selected_outline_prefixes),
            ("새 번호", self.put_new_paragraph_number_at_cursor),
            ("이어서", self.put_continued_paragraph_number_at_cursor),
        )
        for index, (button_text, command) in enumerate(normal_defs):
            ttk.Button(right_buttons, text=button_text, command=command).grid(
                row=index // 2,
                column=index % 2,
                sticky="nsew",
                padx=(0 if index % 2 == 0 else 6, 0),
                pady=(0 if index < 2 else 6, 0),
                ipady=8,
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

        decimal_group = ttk.LabelFrame(parent, text="소숫점 지정", padding=8)
        decimal_group.pack(fill="x", pady=(8, 0))
        decimal_row = ttk.Frame(decimal_group)
        decimal_row.pack(fill="x")
        ttk.Label(decimal_row, text="소숫점").pack(side="left")
        ttk.Entry(decimal_row, textvariable=self.decimal_places_var, width=5).pack(side="left", padx=(4, 6))
        ttk.Label(decimal_row, text="자리").pack(side="left")
        ttk.Combobox(
            decimal_row,
            textvariable=self.decimal_mode_var,
            values=("반올림", "올림", "버림"),
            state="readonly",
            width=7,
        ).pack(side="left", padx=(8, 6))
        ttk.Checkbutton(decimal_row, text="쉼표", variable=self.decimal_use_commas_var).pack(side="left")
        ttk.Button(decimal_row, text="소숫점 처리", command=self.apply_decimal_rounding_to_selection).pack(
            side="left", fill="x", expand=True, padx=(6, 0)
        )

        unit_group = ttk.LabelFrame(parent, text="단위 변환", padding=8)
        unit_group.pack(fill="x", pady=(8, 0))
        unit_decimal_row = ttk.Frame(unit_group)
        unit_decimal_row.pack(fill="x")
        ttk.Label(unit_decimal_row, text="변환 후 소숫점").pack(side="left")
        ttk.Entry(unit_decimal_row, textvariable=self.unit_decimal_places_var, width=5).pack(side="left", padx=(4, 6))
        ttk.Label(unit_decimal_row, text="자리").pack(side="left")
        ttk.Combobox(
            unit_decimal_row,
            textvariable=self.unit_decimal_mode_var,
            values=("반올림", "올림", "버림"),
            state="readonly",
            width=7,
        ).pack(side="left", padx=(8, 6))
        ttk.Checkbutton(unit_decimal_row, text="쉼표", variable=self.unit_decimal_use_commas_var).pack(side="left")
        unit_row = ttk.Frame(unit_group)
        unit_row.pack(fill="x", pady=(6, 0))
        for column in range(5):
            unit_row.grid_columnconfigure(column, weight=1, uniform="unit-buttons")
        unit_buttons = (
            ("× 천", "scale", Decimal("1000"), {"천": "원", "천원": "원", "백만": "천원", "백만원": "천원"}),
            ("÷ 천", "scale", Decimal("0.001"), {"원": "천원", "천": "백만원", "천원": "백만원"}),
            ("원", "target", "원"),
            ("천", "target", "천"),
            ("백만", "target", "백만"),
        )
        unit_tooltips = {
            "× 천": "숫자에 1,000을 곱합니다. 단위가 붙은 경우 천→원, 백만→천원으로 바꿉니다.",
            "÷ 천": "숫자를 1,000으로 나눕니다. 단위가 붙은 경우 원→천원, 천→백만원으로 바꿉니다.",
            "원": "선택한 숫자를 원 단위로 바꿉니다. 예: 1백만 → 1,000,000원, 1천 → 1,000원",
            "천": "선택한 숫자를 천원 단위로 바꿉니다. 예: 1,000원 → 1천원, 1백만 → 1,000천원",
            "백만": "선택한 숫자를 백만원 단위로 바꿉니다. 예: 1,000,000원 → 1백만원, 1천 → 0.001백만원",
        }
        for index, button_def in enumerate(unit_buttons):
            button_text = button_def[0]
            if button_def[1] == "scale":
                _text, _mode, multiplier, unit_transitions = button_def
                command = lambda label=button_text, factor=multiplier, transitions=unit_transitions: self.apply_decimal_unit_conversion_to_selection(
                    label,
                    factor,
                    transitions,
                )
            else:
                _text, _mode, target_unit = button_def
                command = lambda label=button_text, unit=target_unit: self.apply_currency_unit_conversion_to_selection(label, unit)
            button = ttk.Button(
                unit_row,
                text=button_text,
                command=command,
            )
            button.grid(row=0, column=index, sticky="ew", padx=(6 if index else 0, 0))
            ToolTip(button, unit_tooltips[button_text])

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

        year_group = ttk.LabelFrame(parent, text="연도 정규화", padding=8)
        year_group.pack(fill="x", pady=(8, 0))
        year_buttons = ttk.Frame(year_group)
        year_buttons.pack(fill="x")
        ttk.Button(year_buttons, text="0000년", command=lambda: self.normalize_years_to_selection("full", "0000년")).pack(
            side="left", fill="x", expand=True
        )
        ttk.Button(year_buttons, text="00년", command=lambda: self.normalize_years_to_selection("short", "00년")).pack(
            side="left", fill="x", expand=True, padx=(6, 0)
        )
        ttk.Button(year_buttons, text="'00년", command=lambda: self.normalize_years_to_selection("apostrophe", "'00년")).pack(
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
        frame = self.create_scrollable_tab("양식/그림")

        cover_group = ttk.LabelFrame(frame, text="표지 자동화", padding=8)
        cover_group.pack(fill="x", pady=(0, 10))
        ttk.Checkbutton(cover_group, text="대외비 표시", variable=self.cover_confidential).pack(anchor="w")
        ttk.Button(cover_group, text="선택 내용으로 표지 만들기", command=self.create_cover_from_selected_lines).pack(
            fill="x", pady=(6, 0)
        )
        ttk.Label(
            cover_group,
            text="한글에서 제목, 날짜, 부서명을 선택한 뒤 실행합니다. 부서명은 Shift+Enter 줄바꿈처럼 여러 줄이어도 됩니다.",
            wraplength=390,
        ).pack(anchor="w", pady=(6, 0))

        template_group = ttk.LabelFrame(frame, text="보고 양식", padding=8)
        template_group.pack(fill="x", pady=(0, 10))
        ttk.Button(
            template_group,
            text="보고양식 불러오기",
            command=self.insert_general_report_template,
        ).pack(fill="x")
        ttk.Button(
            template_group,
            text="대제목 번호박스 삽입",
            command=self.insert_title_number_box_template,
        ).pack(fill="x", pady=(6, 0))

        proof_group = ttk.LabelFrame(frame, text="증빙 PDF", padding=8)
        proof_group.pack(fill="x", pady=(0, 10))
        proof_option_row = ttk.Frame(proof_group)
        proof_option_row.pack(fill="x", pady=(0, 6))
        ttk.Label(proof_option_row, text="해상도(dpi)").pack(side="left")
        ttk.Entry(proof_option_row, textvariable=self.proof_dpi_var, width=7).pack(side="left", padx=(6, 10))
        ttk.Checkbutton(
            proof_option_row,
            text="가로 자동 회전",
            variable=self.proof_auto_rotate_var,
        ).pack(side="left")
        ttk.Button(
            proof_group,
            text="증빙 PDF 폴더 삽입",
            command=self.insert_proof_pdf_folder,
        ).pack(fill="x")
        ttk.Button(
            proof_group,
            text="증빙 PDF 1건 삽입",
            command=self.insert_single_proof_pdf,
        ).pack(fill="x", pady=(6, 0))

        picture_group = ttk.LabelFrame(frame, text="그림", padding=8)
        picture_group.pack(fill="x", pady=(0, 10))
        picture_tool_row = ttk.Frame(picture_group)
        picture_tool_row.pack(fill="x")
        ttk.Button(
            picture_tool_row,
            text="그림 왼쪽 90도",
            command=lambda: self.rotate_selected_picture(clockwise=False),
        ).grid(row=0, column=0, sticky="ew")
        ttk.Button(
            picture_tool_row,
            text="그림 오른쪽 90도",
            command=lambda: self.rotate_selected_picture(clockwise=True),
        ).grid(row=0, column=1, sticky="ew", padx=(6, 0))
        ttk.Button(
            picture_tool_row,
            text="선택 그림 셀 맞춤",
            command=self.fit_selected_picture_to_proof_cell,
        ).grid(row=0, column=2, sticky="ew", padx=(6, 0))
        for column in range(3):
            picture_tool_row.grid_columnconfigure(column, weight=1, uniform="picture-tools")

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
            status_text = getattr(self, "status_text", None)
            if status_text is not None:
                status_text.insert("end", line + "\n")
        status_text = getattr(self, "status_text", None)
        if status_text is not None:
            status_text.see("end")

    def debug(self, line: str) -> None:
        self.log(line)

    def check_for_updates(self, silent: bool = False) -> None:
        settings = load_update_settings()
        self.update_settings = settings
        if not settings.get("enabled"):
            if not silent:
                messagebox.showinfo("업데이트 확인", "업데이트 확인이 꺼져 있습니다.")
            return
        provider = str(settings.get("provider") or "").strip().lower()
        if provider != "github" and not str(settings.get("version_url") or "").strip():
            message = f"업데이트 확인 URL이 설정되지 않았습니다.\n\n설정 파일: {UPDATE_SETTINGS_FILE}"
            self.log("업데이트 확인: version_url 미설정")
            if not silent:
                messagebox.showinfo("업데이트 확인", message)
            return

        def worker() -> None:
            try:
                info = fetch_update_info(settings)
                self.after(0, lambda: self.handle_update_result(info, silent))
            except Exception as exc:
                self.after(0, lambda error=exc: self.handle_update_error(error, silent))

        self.log("업데이트 확인 중...")
        threading.Thread(target=worker, daemon=True).start()

    def handle_update_result(self, info: UpdateInfo, silent: bool) -> None:
        if not is_newer_version(info.latest_version, APP_VERSION):
            self.log(f"업데이트 확인: 최신 버전입니다. current={APP_VERSION}, latest={info.latest_version}")
            if not silent:
                messagebox.showinfo("업데이트 확인", f"현재 최신 버전입니다.\n\n현재 버전: {APP_VERSION}")
            return

        self.log(f"업데이트 확인: 새 버전 발견 current={APP_VERSION}, latest={info.latest_version}")
        notes = f"\n\n변경 내용:\n{info.notes}" if info.notes else ""
        install_now = messagebox.askyesno(
            "새 버전 있음",
            f"새 버전이 있습니다.\n\n현재 버전: {APP_VERSION}\n최신 버전: {info.latest_version}{notes}\n\n지금 다운로드하고 설치할까요?\n앱은 설치 과정에서 자동으로 종료됩니다.",
        )
        if install_now:
            self.download_and_install_update(info)
        elif info.html_url or info.download_url:
            open_page = messagebox.askyesno("업데이트", "릴리즈 페이지를 브라우저로 열까요?")
            if open_page:
                webbrowser.open(info.html_url or info.download_url)

    def download_and_install_update(self, info: UpdateInfo) -> None:
        settings = self.update_settings
        timeout = float(settings.get("timeout_seconds") or DEFAULT_UPDATE_SETTINGS["timeout_seconds"])
        expected_root = str(settings.get("expected_root_dir") or DEFAULT_UPDATE_SETTINGS["expected_root_dir"])
        expected_exe = str(settings.get("expected_exe_name") or DEFAULT_UPDATE_SETTINGS["expected_exe_name"])

        def progress(received: int, total: int) -> None:
            if total:
                percent = int(received * 100 / total)
                self.after(0, lambda: self.log(f"업데이트 다운로드 중... {percent}%"))

        def worker() -> None:
            try:
                self.after(0, lambda: self.log("업데이트 다운로드 시작"))
                zip_path = download_update_asset(info, timeout, progress)
                validate_update_zip(zip_path, expected_root, expected_exe)
                self.after(0, lambda: self.start_update_install(zip_path, info, settings))
            except Exception as exc:
                self.after(0, lambda error=exc: self.handle_update_install_error(error, info))

        threading.Thread(target=worker, daemon=True).start()

    def start_update_install(self, zip_path: Path, info: UpdateInfo, settings: dict) -> None:
        try:
            launch_update_helper(zip_path, info, settings)
        except Exception as exc:
            self.handle_update_install_error(exc, info)
            return
        self.log(f"업데이트 설치 시작: {info.latest_version}")
        self.destroy()

    def handle_update_install_error(self, exc: Exception, info: UpdateInfo) -> None:
        self.log(f"업데이트 설치 준비 실패: {type(exc).__name__}: {exc}")
        open_page = messagebox.askyesno(
            "업데이트 설치 실패",
            f"자동 설치를 준비하지 못했습니다.\n\n{exc}\n\n릴리즈 페이지를 브라우저로 열까요?",
        )
        if open_page:
            webbrowser.open(info.html_url or info.download_url)

    def handle_update_error(self, exc: Exception, silent: bool) -> None:
        self.log(f"업데이트 확인 실패: {type(exc).__name__}: {exc}")
        if not silent:
            messagebox.showwarning("업데이트 확인 실패", str(exc))

    def refresh_all(self) -> None:
        self._refreshing = True
        self.status_text.delete("1.0", "end")
        self.update_settings = load_update_settings()
        self.log(f"{APP_NAME} 버전: {APP_VERSION}")
        self.log(f"작업 폴더: {ROOT}")
        self.log(f"사용자 설정 폴더: {CONFIG_ROOT}")
        self.log(f"업데이트 설정 파일: {UPDATE_SETTINGS_FILE.exists()} / {UPDATE_SETTINGS_FILE}")
        self.log(f"스타일 세트 파일: {STYLE_SETS_FILE.exists()} / {STYLE_SETS_FILE}")
        self.log(f"특수문자 설정 파일: {SPECIAL_CHARS_FILE.exists()} / {SPECIAL_CHARS_FILE}")
        self.log(f"기본 스타일 세트 파일: {BUNDLED_STYLE_SETS_FILE.exists()} / {BUNDLED_STYLE_SETS_FILE.name}")
        self.log(f"템플릿 폴더: 기본={BUILTIN_TEMPLATE_DIR}, 사용자={CUSTOM_TEMPLATE_DIR}")
        self.log(f"표지 파일: {COVER_FILE.exists()} / {display_path(COVER_FILE)}")
        self.log(f"보고양식 파일: {GENERAL_REPORT_TEMPLATE_FILE.exists()} / {display_path(GENERAL_REPORT_TEMPLATE_FILE)}")
        self.log(f"대제목 번호박스 파일: {TITLE_NUMBER_BOX_TEMPLATE_FILE.exists()} / {display_path(TITLE_NUMBER_BOX_TEMPLATE_FILE)}")

        self.active_style_set_name, self.style_sets = load_style_sets()
        self.ensure_title_number_box_role_configs()
        self.style_set_var.set(self.active_style_set_name)
        self.refresh_style_set_combo()
        self.style_records = style_set_to_records(self.active_style_set())
        self.populate_style_list()
        self.build_table_tab_content()
        self.log(f"스타일 세트: {self.active_style_set_name}")
        self.log(f"스타일 이름 수: {len(self.style_records)}")
        if self.hwp is not None:
            self.refresh_current_doc_style_map(force=True)

        logos = list_logos()
        self.populate_logo_chips(logos)
        self.log(f"로고 파일 수: {len(logos)}")
        self._refreshing = False

    def active_style_set(self) -> StyleSet:
        if "style_sets" not in self.__dict__:
            records = getattr(self, "style_records", [])
            return StyleSet("기본", [
                StyleEntry(record.name, table_style=normalize_style_name(record.name).startswith(normalize_style_name("표내용-")))
                for record in records
                if record.style_type == "PARA"
            ], [
                StyleEntry(record.name, table_style=normalize_style_name(record.name).startswith(normalize_style_name("표내용-")))
                for record in records
                if record.style_type == "CHAR"
            ], role_configs={})
        for style_set in self.style_sets:
            if style_set.name == self.active_style_set_name:
                return style_set
        if not self.style_sets:
            _, self.style_sets = default_style_sets()
        self.active_style_set_name = self.style_sets[0].name
        return self.style_sets[0]

    def ensure_title_number_box_role_configs(self) -> None:
        raw = self.table_settings.get("title_number_box", {}) if isinstance(self.table_settings, dict) else {}
        template_file = str(raw.get("template_file", DEFAULT_TITLE_NUMBER_BOX_SETTINGS["template_file"])).strip()
        migrated, changed = migrate_title_number_box_role_configs(self.style_sets, template_file)
        if changed:
            self.style_sets = migrated
            save_style_sets(self.active_style_set_name, self.style_sets)
            self.debug(f"[style-sets] title_number_box role config migration applied: template={template_file}")

    def style_entries_for_ai_prompt(self, style_set: StyleSet | None = None) -> list[StyleEntry]:
        selected_set = style_set or self.active_style_set()
        entries: list[StyleEntry] = []
        for entry in selected_set.paragraph_styles:
            if entry.table_style or entry.caption_style or not entry.outline_markers:
                continue
            entries.append(entry)
        return entries

    def primary_outline_marker(self, entry: StyleEntry) -> str:
        return next((marker.strip() for marker in entry.outline_markers if str(marker).strip()), "")

    def ai_prompt_outline_entries(self, style_set: StyleSet | None = None) -> list[tuple[StyleEntry, str]]:
        outline_entries: list[tuple[StyleEntry, str]] = []
        for entry in self.style_entries_for_ai_prompt(style_set):
            marker = self.primary_outline_marker(entry)
            if marker:
                outline_entries.append((entry, marker))
        return outline_entries

    def format_ai_prompt_tokens(self, values: Iterable[str]) -> str:
        return ", ".join(f"`{value}`" for value in values)

    def ai_prompt_hierarchy_rule_lines(self, entries: list[tuple[StyleEntry, str]]) -> list[str]:
        names = [entry.name for entry, _marker in entries]
        if not names:
            return []

        lines = [f"* 문단 계층은 반드시 `{' → '.join(names)}` 순으로만 사용한다."]
        if len(names) == 1:
            lines.append(f"* {names[0]}만 사용한다.")
            return lines

        lines.append("* 상위 단계를 건너뛰고 하위 단계를 직접 사용할 수 없다.")
        for index, name in enumerate(names):
            if index == len(names) - 1:
                lines.append(f"* {name}은 최하위 단계이며 {name} 아래에는 다른 계층을 사용할 수 없다.")
                continue

            allowed_children = [names[index + 1]]
            if index < 2 and index + 2 < len(names):
                allowed_children.append(names[index + 2])
            if len(allowed_children) == 1:
                lines.append(f"* {name} 아래에는 반드시 {allowed_children[0]}만 올 수 있다.")
            else:
                lines.append(f"* {name} 아래에는 반드시 {' 또는 '.join(allowed_children)}만 올 수 있다.")
        lines.append("* 같은 상위 계층 안에서는 하위 계층을 사용한 후 다시 상위 계층으로 돌아갈 수 있다.")
        lines.append("* 계층 구조를 임의로 생략하거나 역순으로 작성해서는 안 된다.")
        return lines

    def ai_prompt_for_style_set(self, style_set: StyleSet | None = None) -> str:
        selected_set = style_set or self.active_style_set()
        outline_entries = self.ai_prompt_outline_entries(selected_set)
        if not outline_entries:
            raise ValueError(f"현재 스타일 세트 '{selected_set.name}'에 식별자가 설정된 문단 스타일이 없습니다.")

        identifier_lines: list[str] = []
        important_markers: list[str] = []
        style_names: list[str] = []
        for entry, marker in outline_entries:
            identifier_lines.append(f"* {entry.name} 문단은 반드시 `{marker} `으로 시작한다.")
            if marker not in important_markers:
                important_markers.append(marker)
            style_names.append(entry.name)

        if not identifier_lines:
            raise ValueError(f"현재 스타일 세트 '{selected_set.name}'에 사용할 수 있는 대표 식별자가 없습니다.")

        marker_text = self.format_ai_prompt_tokens(important_markers)
        first_style_name = outline_entries[0][0].name
        hierarchy_text = " → ".join(style_names)
        first_marker = outline_entries[0][1]
        second_marker = outline_entries[1][1] if len(outline_entries) > 1 else outline_entries[0][1]
        third_marker = outline_entries[2][1] if len(outline_entries) > 2 else second_marker
        return "\n".join(
            [
                "[최우선 규칙]",
                "1. 설명문, 안내문, 메모 없이 최종 본문만 출력한다.",
                "2. 보고서 전체 제목, 문서 제목, 보고서명은 쓰지 말고 본문부터 시작한다.",
                f"3. `{first_style_name}`은 한 보고서 안에서 여러 번 반복될 수 있는 본문 섹션 제목이다.",
                f"4. `{first_style_name}`은 배경, 현황, 운영 방안, 기대 효과처럼 본문을 여러 장으로 나누는 데 사용한다.",
                f"5. 한 문서에는 `{first_style_name}`이 여러 개 나올 수 있으며, 첫 `{first_style_name}`이 문서 전체 제목을 대신하지 않는다.",
                f"6. 모든 문단은 반드시 {marker_text} 중 하나로 시작한다.",
                "7. `가.`, `나.`, `1.`, `2.`, `①`, `-`, `•` 같은 실제 번호나 기호는 직접 쓰지 않는다.",
                "8. 코드블록, JSON, YAML, 마크다운 제목(#)은 쓰지 않는다. 단, 표만 `markdown` 코드블록 안에 작성한다.",
                "9. 문단은 줄바꿈으로만 구분하고 문단 앞 공백이나 들여쓰기를 넣지 않는다.",
                "",
                "[구조 예시]",
                f"* `{first_marker} 배경 및 목적`",
                f"* `{second_marker} 추진 배경`",
                f"* `{third_marker} 제도 도입 필요`",
                f"* `{first_marker} 운영 방안`",
                "",
                "[스타일 식별자 규칙]",
                *identifier_lines,
                "",
                "[계층 규칙]",
                f"* 문단 계층은 `{hierarchy_text}` 순서만 사용한다.",
                "* 상위 단계를 건너뛰고 하위 단계를 직접 쓰지 않는다.",
                "* 같은 상위 문단 아래에서는 하위 단계를 쓴 뒤 다시 상위 단계로 돌아갈 수 있다.",
                "",
                "[문체 규칙]",
                "- 전체 문장은 행정 보고서용 개조식 문체로 작성한다.",
                "- 각 문단은 하나의 내용만 포함하도록 짧고 간결하게 작성한다.",
                "- 종결 표현은 `~함`, `~필요`, `~예정`, `~추진`, `~검토`, `~기대`, `~강화`, `~지원` 등 개조식 표현을 우선 사용한다.",
                "- `~합니다`, `~했습니다`, `~이다`, `~입니다` 등의 완결형 서술문은 사용하지 않는다.",
                "- 중복 표현과 수식어는 최소화한다.",
                "",
                "[표 작성 규칙]",
                "1. 표는 반드시 Markdown pipe table 형식으로 작성한다.",
                "2. 모든 표는 반드시 `markdown` 코드블록 안에만 넣는다.",
                "3. 표 앞뒤 설명은 일반 문단으로 분리한다.",
                "4. 셀 안 줄바꿈이 필요하면 `<br>`를 사용한다.",
                "5. 병합셀과 복잡한 중첩 표는 쓰지 않는다.",
                "",
                "[출력 전 확인]",
                f"* 모든 문단이 {marker_text} 중 하나로 시작하는지 확인한다.",
                f"* 계층이 `{hierarchy_text}` 순서를 어기지 않았는지 확인한다.",
                "* 실제 번호나 글머리기호를 쓰지 않았는지 확인한다.",
                "* 표가 있다면 모두 `markdown` 코드블록 안에 있는지 확인한다.",
            ]
        )

    def copy_ai_prompt_for_active_style_set(self) -> None:
        label = "AI 프롬프트 복사"
        try:
            prompt_text = self.ai_prompt_for_style_set()
            self.set_clipboard_text(prompt_text)
            self.log(f"{label}: set={self.active_style_set().name}, chars={len(prompt_text)}")
            messagebox.showinfo(label, "현재 스타일 세트 기준 프롬프트를 클립보드에 복사했습니다.")
        except Exception as exc:
            self.log(f"{label} 실패: {type(exc).__name__}: {exc}")
            messagebox.showwarning(label, str(exc))

    def set_active_style_set(self, name: str) -> None:
        if name not in {style_set.name for style_set in self.style_sets}:
            return
        self.active_style_set_name = name
        self.style_set_var.set(name)
        self.style_records = style_set_to_records(self.active_style_set())
        self.save_active_style_set_preference()
        self.populate_style_list()
        self.build_table_tab_content()
        if self.hwp is not None:
            self.refresh_current_doc_style_map(force=True)
        self.debug(f"[style-sets] 활성 세트 변경: {name}")

    def save_active_style_set_preference(self) -> None:
        names = {style_set.name for style_set in self.style_sets}
        selected_name = self.style_set_var.get().strip()
        if selected_name in names:
            self.active_style_set_name = selected_name
        save_style_sets(self.active_style_set_name, self.style_sets)

    def on_close(self) -> None:
        try:
            self.save_active_style_set_preference()
        except Exception as exc:
            self.debug(f"[style-sets] 종료 전 활성 세트 저장 실패: {type(exc).__name__}: {exc}")
        self.destroy()

    def refresh_style_set_combo(self) -> None:
        if "style_set_combo" in self.__dict__:
            self.style_set_combo.configure(values=[style_set.name for style_set in self.style_sets])

    def on_style_set_selected(self, _event=None) -> None:
        self.set_active_style_set(self.style_set_var.get())

    def open_style_sets_window(self) -> None:
        window = tk.Toplevel(self)
        window.title("스타일 세트 관리")
        window.geometry("1240x620")
        window.minsize(1100, 560)
        window.transient(self)

        working_sets = [
            StyleSet(
                item.name,
                list(item.paragraph_styles),
                list(item.character_styles),
                list(item.inline_rules),
                normalize_role_configs(item.role_configs),
            )
            for item in self.style_sets
        ]
        selected_set_var = tk.StringVar(value=self.active_style_set_name)
        style_type_var = tk.StringVar(value="문단")
        style_name_var = tk.StringVar()
        table_style_var = tk.BooleanVar(value=False)
        caption_style_var = tk.BooleanVar(value=False)
        outline_markers_var = tk.StringVar()
        roles_var = tk.StringVar()
        title_box_template_var = tk.StringVar()
        style_template_var = tk.StringVar()
        numbering_group_var = tk.StringVar()
        numbering_level_var = tk.StringVar()
        numbering_restart_var = tk.BooleanVar(value=False)
        rule_name_var = tk.StringVar()
        rule_type_var = tk.StringVar(value=inline_rule_type_label("leading_parenthesized_after_identifier"))
        rule_target_var = tk.StringVar(value="body")
        rule_role_var = tk.StringVar(value="body_bold")
        rule_include_wrapper_var = tk.BooleanVar(value=True)
        rule_remove_markers_var = tk.BooleanVar(value=False)
        rule_enabled_var = tk.BooleanVar(value=True)
        adding_style_var = tk.BooleanVar(value=False)
        adding_rule_var = tk.BooleanVar(value=False)
        style_form_snapshot: list[tuple[object, ...] | None] = [None]
        rule_form_snapshot: list[tuple[object, ...] | None] = [None]
        style_selection_guard = [False]
        rule_selection_guard = [False]
        set_selection_guard = [False]
        tab_selection_guard = [False]
        last_style_selection: list[str | None] = [None]
        last_rule_selection: list[str | None] = [None]
        last_set_index: list[int | None] = [None]
        last_tab: list[str | None] = [None]
        rule_role_combo_ref: list[ttk.Combobox] = []
        style_role_combo_ref: list[ttk.Combobox] = []
        style_template_entry_ref: list[ttk.Entry] = []
        style_template_button_ref: list[ttk.Button] = []

        body = ttk.Frame(window, padding=10)
        body.pack(fill="both", expand=True)
        body.grid_columnconfigure(1, weight=1)
        body.grid_rowconfigure(1, weight=1)

        set_list = tk.Listbox(body, height=6, exportselection=False)
        set_list.grid(row=0, column=0, rowspan=2, sticky="nsew", padx=(0, 10))

        notebook = ttk.Notebook(body)
        notebook.grid(row=1, column=1, sticky="nsew")
        style_tab = ttk.Frame(notebook, padding=(0, 8, 0, 0))
        rule_tab = ttk.Frame(notebook, padding=(0, 8, 0, 0))
        style_tab.grid_columnconfigure(0, weight=1)
        style_tab.grid_rowconfigure(0, weight=1)
        rule_tab.grid_columnconfigure(0, weight=1)
        rule_tab.grid_rowconfigure(0, weight=1)
        notebook.add(style_tab, text="스타일")
        notebook.add(rule_tab, text="규칙")

        style_tree = ttk.Treeview(
            style_tab,
            columns=("type", "table", "caption", "outline", "roles", "num_group", "num_level", "num_restart"),
            show="tree headings",
            height=10,
        )
        style_tree.heading("#0", text="스타일 이름")
        style_tree.heading("type", text="구분")
        style_tree.heading("table", text="표")
        style_tree.heading("caption", text="캡션")
        style_tree.heading("outline", text="식별자")
        style_tree.heading("roles", text="역할")
        style_tree.heading("num_group", text="번호그룹")
        style_tree.heading("num_level", text="수준")
        style_tree.heading("num_restart", text="새번호")
        style_tree.column("#0", width=260)
        style_tree.column("type", width=52, stretch=False)
        style_tree.column("table", width=42, stretch=False, anchor="center")
        style_tree.column("caption", width=52, stretch=False, anchor="center")
        style_tree.column("outline", width=150)
        style_tree.column("roles", width=130)
        style_tree.column("num_group", width=110)
        style_tree.column("num_level", width=44, stretch=False, anchor="center")
        style_tree.column("num_restart", width=56, stretch=False, anchor="center")
        style_tree.grid(row=0, column=0, sticky="nsew")

        rule_tree = ttk.Treeview(rule_tab, columns=("type", "target", "role", "options"), show="tree headings", height=8)
        rule_tree.heading("#0", text="규칙 이름")
        rule_tree.heading("type", text="찾을 형식")
        rule_tree.heading("target", text="대상")
        rule_tree.heading("role", text="글자 역할")
        rule_tree.heading("options", text="옵션")
        rule_tree.column("#0", width=180)
        rule_tree.column("type", width=210)
        rule_tree.column("target", width=70, stretch=False)
        rule_tree.column("role", width=100)
        rule_tree.column("options", width=120)
        rule_tree.grid(row=0, column=0, sticky="nsew")

        def selected_set_index() -> int:
            selection = set_list.curselection()
            if selection:
                return selection[0]
            return max(0, next((index for index, item in enumerate(working_sets) if item.name == selected_set_var.get()), 0))

        def selected_set() -> StyleSet:
            return working_sets[selected_set_index()]

        def refresh_set_list(select_name: str | None = None) -> None:
            set_list.delete(0, "end")
            for item in working_sets:
                set_list.insert("end", item.name)
            name = select_name or selected_set_var.get()
            index = next((idx for idx, item in enumerate(working_sets) if item.name == name), 0)
            if working_sets:
                set_list.selection_set(index)
                set_list.activate(index)
                selected_set_var.set(working_sets[index].name)

        def refresh_style_tree() -> None:
            style_tree.delete(*style_tree.get_children())
            item = selected_set()
            for style_type, entries in (("문단", item.paragraph_styles), ("글자", item.character_styles)):
                for index, entry in enumerate(entries):
                    style_tree.insert(
                        "",
                        "end",
                        iid=f"{style_type}:{index}",
                        text=entry.name,
                        values=(
                            style_type,
                            "✓" if entry.table_style else "",
                            "✓" if entry.caption_style else "",
                            ", ".join(entry.outline_markers),
                            ", ".join(entry.roles),
                            numbering_group_label(entry.numbering_group),
                            entry.numbering_level if entry.numbering_level else "",
                            "✓" if entry.restart_after_higher_level else "",
                        ),
                    )

        def rule_target_label(target: str) -> str:
            return {"body": "본문", "table": "표", "all": "본문+표"}.get(target, target)

        def rule_type_label(rule_type: str) -> str:
            return inline_rule_type_label(rule_type)

        def refresh_rule_tree() -> None:
            rule_tree.delete(*rule_tree.get_children())
            for index, rule in enumerate(selected_set().inline_rules):
                options = []
                if rule.include_wrapper:
                    options.append("괄호포함")
                if rule.remove_markers:
                    options.append("표식삭제")
                if not rule.enabled:
                    options.append("꺼짐")
                rule_tree.insert(
                    "",
                    "end",
                    iid=f"rule:{index}",
                    text=rule.name,
                    values=(
                        rule_type_label(rule.rule_type),
                        rule_target_label(rule.target),
                        rule.style_role,
                        ", ".join(options),
                    ),
                )

        def character_role_options() -> list[str]:
            roles: list[str] = []
            seen: set[str] = set()
            for entry in selected_set().character_styles:
                for role in entry.roles:
                    if role and role not in seen:
                        roles.append(role)
                        seen.add(role)
            return roles

        def selected_set_title_box_template() -> str:
            role_configs = normalize_role_configs(selected_set().role_configs)
            config = role_configs.get("title_number_box", {})
            return str(config.get("template_file", "")).strip()

        def refresh_title_box_template_form() -> None:
            title_box_template_var.set(selected_set_title_box_template())

        def style_uses_title_box_role(style_type: object | None = None, role_value: object | None = None) -> bool:
            current_style_type = str(style_type if style_type is not None else style_type_var.get()).strip()
            current_role = normalize_style_role_value(roles_var.get() if role_value is None else role_value)
            return current_style_type == "문단" and current_role == "title_number_box"

        def refresh_style_template_controls() -> None:
            enabled = style_uses_title_box_role()
            state = "normal" if enabled else "disabled"
            for widget in style_template_entry_ref:
                widget.configure(state=state)
            for widget in style_template_button_ref:
                widget.configure(state=state)
            if not enabled:
                style_template_var.set("")

        def update_style_role_options(current_role: object | None = None) -> None:
            if not style_role_combo_ref:
                return
            explicit_role = normalize_style_role_value(current_role) if current_role is not None else ""
            values = role_selection_values(style_type_var.get(), explicit_role)
            style_role_combo_ref[0].configure(values=values)
            current_value = normalize_style_role_value(roles_var.get())
            if current_value and current_value in values:
                roles_var.set(current_value)
            elif explicit_role and explicit_role in values:
                roles_var.set(explicit_role)
            elif roles_var.get() not in values:
                roles_var.set(STYLE_ROLE_NONE_LABEL)
            refresh_style_template_controls()

        def refresh_rule_role_options() -> None:
            if not rule_role_combo_ref:
                return
            roles = character_role_options()
            rule_role_combo_ref[0].configure(values=roles)
            if roles and rule_role_var.get() not in roles:
                rule_role_var.set(roles[0])
            elif not roles:
                rule_role_var.set("")

        def validate_style_roles(style_type: str, roles: list[str], replace_index: int | None = None) -> bool:
            allowed = set(allowed_style_roles(style_type))
            invalid_roles = [role for role in roles if role not in allowed]
            if invalid_roles:
                allowed_text = ", ".join(allowed_style_roles(style_type)) or STYLE_ROLE_NONE_LABEL
                messagebox.showwarning(
                    "역할 제한",
                    f"{style_type} 스타일에는 다음 역할만 지정할 수 있습니다.\n\n{allowed_text}",
                    parent=window,
                )
                return False
            target_entries = selected_set().character_styles if style_type == "글자" else selected_set().paragraph_styles
            for role in roles:
                if style_type == "문단" and role == "title_number_box":
                    continue
                for index, entry in enumerate(target_entries):
                    if replace_index is not None and index == replace_index:
                        continue
                    if role in entry.roles:
                        messagebox.showwarning(
                            "역할 중복",
                            f"'{role}' 역할은 이미 '{entry.name}' {style_type} 스타일에 지정되어 있습니다.",
                            parent=window,
                        )
                        return False
            return True

        def style_form_values() -> tuple[object, ...]:
            return (
                style_type_var.get(),
                style_name_var.get(),
                table_style_var.get(),
                caption_style_var.get(),
                outline_markers_var.get(),
                roles_var.get(),
                style_template_var.get(),
                numbering_group_var.get(),
                numbering_level_var.get(),
                numbering_restart_var.get(),
                adding_style_var.get(),
            )

        def rule_form_values() -> tuple[object, ...]:
            return (
                rule_name_var.get(),
                rule_type_var.get(),
                rule_target_var.get(),
                rule_role_var.get(),
                rule_include_wrapper_var.get(),
                rule_remove_markers_var.get(),
                rule_enabled_var.get(),
                adding_rule_var.get(),
            )

        def mark_style_form_clean() -> None:
            style_form_snapshot[0] = style_form_values()

        def mark_rule_form_clean() -> None:
            rule_form_snapshot[0] = rule_form_values()

        def style_form_dirty() -> bool:
            return style_form_snapshot[0] is not None and style_form_values() != style_form_snapshot[0]

        def rule_form_dirty() -> bool:
            return rule_form_snapshot[0] is not None and rule_form_values() != rule_form_snapshot[0]

        def confirm_discard_pending(scope: str = "") -> bool:
            check_style = scope in {"", "스타일", "스타일/규칙"}
            check_rule = scope in {"", "규칙", "스타일/규칙"}
            dirty_scopes = []
            if check_style and style_form_dirty():
                dirty_scopes.append("스타일")
            if check_rule and rule_form_dirty():
                dirty_scopes.append("규칙")
            if not dirty_scopes:
                return True
            label = ", ".join(dirty_scopes)
            return messagebox.askyesno("저장 안 된 변경", f"{label}에 저장하지 않은 변경이 있습니다. 버리고 이동할까요?", parent=window)

        def persist_working_sets() -> None:
            active_name = selected_set().name
            self.style_sets = working_sets
            self.active_style_set_name = active_name
            self.style_set_var.set(active_name)
            self.refresh_style_set_combo()
            self.style_records = style_set_to_records(self.active_style_set())
            save_style_sets(active_name, self.style_sets)
            self.populate_style_list()
            self.build_table_tab_content()
            if self.hwp is not None:
                self.refresh_current_doc_style_map(force=True)
            refresh_rule_role_options()
            self.debug(f"[style-sets] 즉시 저장: {STYLE_SETS_FILE}")

        def replace_selected_set_role_configs(role_configs: dict[str, dict[str, object]]) -> None:
            item = selected_set()
            replace_selected_set(
                StyleSet(
                    item.name,
                    item.paragraph_styles,
                    item.character_styles,
                    item.inline_rules,
                    normalize_role_configs(role_configs),
                )
            )
            persist_working_sets()

        def browse_title_box_template() -> None:
            current = title_box_template_var.get().strip()
            initial_path = resolve_config_path(current) if current else TITLE_NUMBER_BOX_TEMPLATE_FILE
            raw_path = filedialog.askopenfilename(
                parent=window,
                title="대제목 번호박스 템플릿 선택",
                initialdir=str(initial_path.parent if initial_path.parent.exists() else TEMPLATE_DIR),
                filetypes=[
                    ("한글 문서", "*.hwpx *.hwp"),
                    ("HWPX", "*.hwpx"),
                    ("HWP", "*.hwp"),
                    ("모든 파일", "*.*"),
                ],
            )
            if not raw_path:
                return
            title_box_template_var.set(normalize_config_path(raw_path))

        def browse_style_template() -> None:
            current = style_template_var.get().strip() or title_box_template_var.get().strip()
            initial_path = resolve_config_path(current) if current else TITLE_NUMBER_BOX_TEMPLATE_FILE
            raw_path = filedialog.askopenfilename(
                parent=window,
                title="이 스타일의 대제목 번호박스 템플릿 선택",
                initialdir=str(initial_path.parent if initial_path.parent.exists() else TEMPLATE_DIR),
                filetypes=[
                    ("한글 문서", "*.hwpx *.hwp"),
                    ("HWPX", "*.hwpx"),
                    ("HWP", "*.hwp"),
                    ("모든 파일", "*.*"),
                ],
            )
            if not raw_path:
                return
            style_template_var.set(normalize_config_path(raw_path))

        def save_title_box_template() -> None:
            template_file = normalize_config_path(title_box_template_var.get())
            title_box_template_var.set(template_file)
            role_configs = normalize_role_configs(selected_set().role_configs)
            title_config = dict(role_configs.get("title_number_box", {}))
            if template_file:
                title_config["template_file"] = template_file
                role_configs["title_number_box"] = title_config
            else:
                role_configs.pop("title_number_box", None)
            replace_selected_set_role_configs(role_configs)
            refresh_title_box_template_form()

        def select_set(_event=None) -> None:
            if set_selection_guard[0]:
                return
            if not working_sets:
                return
            new_index = selected_set_index()
            if last_set_index[0] is not None and new_index != last_set_index[0] and not confirm_discard_pending("스타일/규칙"):
                set_selection_guard[0] = True
                set_list.selection_clear(0, "end")
                set_list.selection_set(last_set_index[0])
                set_list.activate(last_set_index[0])
                set_selection_guard[0] = False
                return
            selected_set_var.set(selected_set().name)
            refresh_style_tree()
            refresh_rule_tree()
            refresh_rule_role_options()
            refresh_title_box_template_form()
            clear_style_form()
            clear_rule_form()
            last_set_index[0] = selected_set_index()

        def replace_selected_set(style_set: StyleSet) -> None:
            working_sets[selected_set_index()] = style_set
            set_selection_guard[0] = True
            refresh_set_list(style_set.name)
            last_set_index[0] = selected_set_index()
            set_selection_guard[0] = False
            refresh_style_tree()
            refresh_rule_tree()
            refresh_rule_role_options()

        def add_set() -> None:
            if not confirm_discard_pending("스타일/규칙"):
                return
            base = "새 세트"
            existing = {item.name for item in working_sets}
            name = base
            suffix = 2
            while name in existing:
                name = f"{base} {suffix}"
                suffix += 1
            working_sets.append(StyleSet(name, [], [], list(DEFAULT_INLINE_RULES), {}))
            refresh_set_list(name)
            refresh_style_tree()
            refresh_rule_tree()
            clear_style_form()
            clear_rule_form()
            last_set_index[0] = selected_set_index()
            persist_working_sets()

        def unique_set_name(base: str) -> str:
            existing = {item.name for item in working_sets}
            base_name = base.strip() or "가져온 스타일"
            name = base_name
            suffix = 2
            while name in existing:
                name = f"{base_name} {suffix}"
                suffix += 1
            return name

        def import_style_set_from_file() -> None:
            if not confirm_discard_pending("스타일/규칙"):
                return
            raw_path = filedialog.askopenfilename(
                parent=window,
                title="스타일 세트로 불러올 한글 파일 선택",
                filetypes=[
                    ("한글 문서", "*.hwpx *.hwp"),
                    ("HWPX", "*.hwpx"),
                    ("HWP", "*.hwp"),
                    ("모든 파일", "*.*"),
                ],
            )
            if not raw_path:
                return
            path = Path(raw_path)
            try:
                records = read_style_records_from_hwp_file(path)
            except Exception as exc:
                self.log(f"스타일 세트 불러오기 실패: {type(exc).__name__}: {exc}")
                messagebox.showerror("스타일 세트 불러오기 실패", str(exc), parent=window)
                return
            if not records:
                messagebox.showwarning("스타일 없음", "선택한 파일에서 스타일 이름을 찾지 못했습니다.", parent=window)
                self.log(f"스타일 세트 불러오기: 스타일 없음, file={path}")
                return
            name = unique_set_name(path.stem)
            working_sets.append(style_set_from_records(name, records, path))
            refresh_set_list(name)
            refresh_style_tree()
            refresh_rule_tree()
            clear_style_form()
            clear_rule_form()
            last_set_index[0] = selected_set_index()
            persist_working_sets()
            self.log(f"스타일 세트 불러오기: {path.name}, styles={len(records)}, set={name}")

        def remove_set() -> None:
            if not confirm_discard_pending("스타일/규칙"):
                return
            if len(working_sets) <= 1:
                messagebox.showwarning("삭제 불가", "스타일 세트는 최소 1개가 필요합니다.", parent=window)
                return
            index = selected_set_index()
            del working_sets[index]
            refresh_set_list(working_sets[min(index, len(working_sets) - 1)].name)
            refresh_style_tree()
            refresh_rule_tree()
            clear_style_form()
            clear_rule_form()
            last_set_index[0] = selected_set_index()
            persist_working_sets()

        def rename_set() -> None:
            name = selected_set_var.get().strip()
            if not name:
                messagebox.showwarning("세트 이름 필요", "세트 이름을 입력하세요.", parent=window)
                return
            index = selected_set_index()
            if any(item.name == name for idx, item in enumerate(working_sets) if idx != index):
                messagebox.showwarning("세트 이름 중복", "이미 있는 세트 이름입니다.", parent=window)
                return
            item = selected_set()
            working_sets[index] = StyleSet(name, item.paragraph_styles, item.character_styles, item.inline_rules, item.role_configs)
            refresh_set_list(name)
            persist_working_sets()

        def style_entry_from_form(style_type: str, replace_index: int | None = None) -> StyleEntry | None:
            name = style_name_var.get().strip()
            if not name:
                messagebox.showwarning("스타일 이름 필요", "스타일 이름을 입력하세요.", parent=window)
                return None
            item = selected_set()
            target = item.character_styles if style_type == "글자" else item.paragraph_styles
            if any(index != replace_index and entry.name == name for index, entry in enumerate(target)):
                messagebox.showwarning("스타일 이름 중복", "같은 구분에 이미 있는 스타일 이름입니다.", parent=window)
                return None
            markers = parse_style_marker_text(outline_markers_var.get())
            roles = list(style_roles_from_selection(roles_var.get()))
            if not validate_style_roles(style_type, roles, replace_index):
                return None
            numbering_level = normalize_numbering_level(numbering_level_var.get())
            numbering_group = normalize_numbering_group(numbering_group_var.get())
            if numbering_level and not numbering_group:
                numbering_group = "table" if table_style_var.get() else "body"
            template_file = normalize_config_path(style_template_var.get()) if style_uses_title_box_role(style_type, roles_var.get()) else ""
            return StyleEntry(
                name,
                table_style_var.get(),
                caption_style_var.get(),
                tuple(markers),
                tuple(roles),
                numbering_group,
                numbering_level,
                numbering_restart_var.get(),
                template_file,
            )

        def clear_style_form() -> None:
            style_name_var.set("")
            table_style_var.set(False)
            caption_style_var.set(False)
            outline_markers_var.set("")
            roles_var.set(STYLE_ROLE_NONE_LABEL)
            update_style_role_options()
            style_template_var.set("")
            numbering_group_var.set("")
            numbering_level_var.set("")
            numbering_restart_var.set(False)
            mark_style_form_clean()

        def start_new_style() -> None:
            if not confirm_discard_pending("스타일"):
                return
            adding_style_var.set(True)
            style_selection_guard[0] = True
            style_tree.selection_remove(style_tree.selection())
            last_style_selection[0] = None
            style_selection_guard[0] = False
            clear_style_form()

        def save_style_form() -> None:
            style_type = style_type_var.get()
            ref = selected_style_ref()
            item = selected_set()
            paragraph_styles = list(item.paragraph_styles)
            character_styles = list(item.character_styles)
            is_new = adding_style_var.get() or ref is None
            replace_index = None
            if not is_new:
                selected_type, selected_index = ref
                style_type = selected_type
                replace_index = selected_index
            entry = style_entry_from_form(style_type, replace_index)
            if entry is None:
                return
            if style_type == "글자":
                if is_new:
                    character_styles.append(entry)
                elif replace_index is not None and replace_index < len(character_styles):
                    character_styles[replace_index] = entry
                else:
                    return
            else:
                if is_new:
                    paragraph_styles.append(entry)
                elif replace_index is not None and replace_index < len(paragraph_styles):
                    paragraph_styles[replace_index] = entry
                else:
                    return
            replace_selected_set(StyleSet(item.name, paragraph_styles, character_styles, item.inline_rules, item.role_configs))
            adding_style_var.set(False)
            clear_style_form()
            last_style_selection[0] = None
            persist_working_sets()

        def delete_selected_style() -> None:
            if adding_style_var.get():
                if confirm_discard_pending("스타일"):
                    adding_style_var.set(False)
                    clear_style_form()
                return
            update_selected_style(remove=True)
            clear_style_form()
            last_style_selection[0] = None

        def selected_style_ref() -> tuple[str, int] | None:
            selection = style_tree.selection()
            if not selection:
                return None
            style_type, raw_index = selection[0].split(":", 1)
            return style_type, int(raw_index)

        def update_selected_style(
            *,
            remove: bool = False,
            save_form: bool = False,
        ) -> None:
            ref = selected_style_ref()
            if ref is None:
                return
            style_type, index = ref
            item = selected_set()
            paragraph_styles = list(item.paragraph_styles)
            character_styles = list(item.character_styles)
            entries = character_styles if style_type == "글자" else paragraph_styles
            if index >= len(entries):
                return
            entry = entries[index]
            if remove:
                del entries[index]
            else:
                saved_roles = style_roles_from_selection(roles_var.get()) if save_form else entry.roles
                if save_form and not validate_style_roles(style_type, list(saved_roles), index):
                    return
                saved_numbering_level = normalize_numbering_level(numbering_level_var.get()) if save_form else entry.numbering_level
                saved_numbering_group = normalize_numbering_group(numbering_group_var.get()) if save_form else entry.numbering_group
                if save_form and saved_numbering_level and not saved_numbering_group:
                    saved_numbering_group = "table" if table_style_var.get() else "body"
                saved_template_file = (
                    normalize_config_path(style_template_var.get()) if save_form and style_uses_title_box_role(style_type, roles_var.get()) else ""
                )
                entries[index] = StyleEntry(
                    entry.name,
                    table_style=table_style_var.get() if save_form else entry.table_style,
                    caption_style=caption_style_var.get() if save_form else entry.caption_style,
                    outline_markers=tuple(parse_style_marker_text(outline_markers_var.get())) if save_form else entry.outline_markers,
                    roles=saved_roles,
                    numbering_group=saved_numbering_group,
                    numbering_level=saved_numbering_level,
                    restart_after_higher_level=numbering_restart_var.get() if save_form else entry.restart_after_higher_level,
                    template_file=saved_template_file if save_form else entry.template_file,
                )
            replace_selected_set(StyleSet(item.name, paragraph_styles, character_styles, item.inline_rules, item.role_configs))
            persist_working_sets()

        def fill_form_from_selected_style(_event=None) -> None:
            if style_selection_guard[0]:
                return
            selection = style_tree.selection()
            new_selection = selection[0] if selection else None
            if new_selection != last_style_selection[0] and not confirm_discard_pending("스타일"):
                style_selection_guard[0] = True
                style_tree.selection_remove(style_tree.selection())
                if last_style_selection[0]:
                    style_tree.selection_set(last_style_selection[0])
                    style_tree.focus(last_style_selection[0])
                style_selection_guard[0] = False
                return
            ref = selected_style_ref()
            if ref is None:
                return
            style_type, index = ref
            item = selected_set()
            entries = item.character_styles if style_type == "글자" else item.paragraph_styles
            if index >= len(entries):
                return
            entry = entries[index]
            style_type_var.set(style_type)
            style_name_var.set(entry.name)
            table_style_var.set(entry.table_style)
            caption_style_var.set(entry.caption_style)
            outline_markers_var.set(", ".join(entry.outline_markers))
            roles_var.set(primary_style_role(entry.roles) or STYLE_ROLE_NONE_LABEL)
            update_style_role_options(primary_style_role(entry.roles))
            style_template_var.set(entry.template_file)
            numbering_group_var.set(numbering_group_label(entry.numbering_group))
            numbering_level_var.set(str(entry.numbering_level) if entry.numbering_level else "")
            numbering_restart_var.set(entry.restart_after_higher_level)
            adding_style_var.set(False)
            last_style_selection[0] = new_selection
            mark_style_form_clean()

        def selected_rule_index() -> int | None:
            selection = rule_tree.selection()
            if not selection:
                return None
            _prefix, raw_index = selection[0].split(":", 1)
            return int(raw_index)

        def fill_form_from_selected_rule(_event=None) -> None:
            if rule_selection_guard[0]:
                return
            selection = rule_tree.selection()
            new_selection = selection[0] if selection else None
            if new_selection != last_rule_selection[0] and not confirm_discard_pending("규칙"):
                rule_selection_guard[0] = True
                rule_tree.selection_remove(rule_tree.selection())
                if last_rule_selection[0]:
                    rule_tree.selection_set(last_rule_selection[0])
                    rule_tree.focus(last_rule_selection[0])
                rule_selection_guard[0] = False
                return
            index = selected_rule_index()
            if index is None:
                return
            rules = selected_set().inline_rules
            if index >= len(rules):
                return
            rule = rules[index]
            rule_name_var.set(rule.name)
            rule_type_var.set(inline_rule_type_label(rule.rule_type))
            rule_target_var.set(rule.target)
            rule_role_var.set(rule.style_role)
            rule_include_wrapper_var.set(rule.include_wrapper)
            rule_remove_markers_var.set(rule.remove_markers)
            rule_enabled_var.set(rule.enabled)
            adding_rule_var.set(False)
            last_rule_selection[0] = new_selection
            mark_rule_form_clean()

        def rule_from_form() -> InlineRule | None:
            name = rule_name_var.get().strip()
            rule_type = normalize_inline_rule_type(rule_type_var.get())
            style_role = rule_role_var.get().strip()
            if not name:
                messagebox.showwarning("규칙 이름 필요", "규칙 이름을 입력하세요.", parent=window)
                return None
            if not style_role:
                messagebox.showwarning("글자 역할 필요", "적용할 글자 역할을 선택하세요.", parent=window)
                return None
            if style_role not in character_role_options():
                messagebox.showwarning("글자 역할 없음", "현재 스타일 세트의 글자 스타일에 등록된 역할만 선택할 수 있습니다.", parent=window)
                return None
            return InlineRule(
                name=name,
                rule_type=rule_type,
                target=normalize_inline_rule_target(rule_target_var.get()),
                style_role=style_role,
                include_wrapper=rule_include_wrapper_var.get(),
                remove_markers=rule_remove_markers_var.get(),
                enabled=rule_enabled_var.get(),
            )

        def replace_inline_rules(rules: list[InlineRule]) -> None:
            item = selected_set()
            replace_selected_set(StyleSet(item.name, item.paragraph_styles, item.character_styles, rules, item.role_configs))
            persist_working_sets()

        def clear_rule_form() -> None:
            rule_name_var.set("")
            rule_type_var.set(inline_rule_type_label("leading_parenthesized_after_identifier"))
            rule_target_var.set("body")
            roles = character_role_options()
            rule_role_var.set(roles[0] if roles else "")
            rule_include_wrapper_var.set(True)
            rule_remove_markers_var.set(False)
            rule_enabled_var.set(True)
            mark_rule_form_clean()

        def start_new_rule() -> None:
            if not confirm_discard_pending("규칙"):
                return
            adding_rule_var.set(True)
            rule_selection_guard[0] = True
            rule_tree.selection_remove(rule_tree.selection())
            last_rule_selection[0] = None
            rule_selection_guard[0] = False
            clear_rule_form()

        def save_selected_rule() -> None:
            rule = rule_from_form()
            if rule is None:
                return
            rules = list(selected_set().inline_rules)
            index = selected_rule_index()
            if adding_rule_var.get() or index is None:
                rules.append(rule)
            elif index < len(rules):
                rules[index] = rule
            else:
                return
            adding_rule_var.set(False)
            replace_inline_rules(rules)
            clear_rule_form()
            last_rule_selection[0] = None

        def remove_selected_rule() -> None:
            if adding_rule_var.get():
                if confirm_discard_pending("규칙"):
                    adding_rule_var.set(False)
                    clear_rule_form()
                return
            index = selected_rule_index()
            if index is None:
                return
            rules = list(selected_set().inline_rules)
            if index >= len(rules):
                return
            del rules[index]
            replace_inline_rules(rules)
            clear_rule_form()
            last_rule_selection[0] = None

        def on_tab_changed(_event=None) -> None:
            if tab_selection_guard[0]:
                return
            current_tab = notebook.select()
            if last_tab[0] is not None and current_tab != last_tab[0] and not confirm_discard_pending("스타일/규칙"):
                tab_selection_guard[0] = True
                notebook.select(last_tab[0])
                tab_selection_guard[0] = False
                return
            last_tab[0] = current_tab

        def close_window() -> None:
            if confirm_discard_pending("스타일/규칙"):
                window.destroy()

        set_form = ttk.Frame(body)
        set_form.grid(row=0, column=1, sticky="ew", pady=(0, 8))
        set_form.grid_columnconfigure(0, weight=1)
        set_toolbar = ttk.Frame(set_form)
        set_toolbar.grid(row=0, column=0, sticky="ew")
        set_toolbar.grid_columnconfigure(0, weight=1)
        ttk.Entry(set_toolbar, textvariable=selected_set_var, width=18).grid(row=0, column=0, sticky="ew")
        ttk.Button(set_toolbar, text="이름변경", command=rename_set).grid(row=0, column=1, padx=(6, 0))
        ttk.Button(set_toolbar, text="세트추가", command=add_set).grid(row=0, column=2, padx=(6, 0))
        ttk.Button(set_toolbar, text="파일에서 불러오기", command=import_style_set_from_file).grid(row=0, column=3, padx=(6, 0))
        ttk.Button(set_toolbar, text="세트삭제", command=remove_set).grid(row=0, column=4, padx=(6, 0))

        role_config_form = ttk.Frame(set_form)
        role_config_form.grid(row=1, column=0, sticky="ew", pady=(8, 0))
        role_config_form.grid_columnconfigure(1, weight=1)
        ttk.Label(role_config_form, text="기본 title_number_box 템플릿").grid(row=0, column=0, sticky="w")
        ttk.Entry(role_config_form, textvariable=title_box_template_var).grid(row=0, column=1, sticky="ew", padx=(6, 0))
        ttk.Button(role_config_form, text="찾아보기", command=browse_title_box_template).grid(row=0, column=2, padx=(6, 0))
        ttk.Button(role_config_form, text="적용", command=save_title_box_template).grid(row=0, column=3, padx=(6, 0))

        add_form = ttk.Frame(style_tab)
        add_form.grid(row=1, column=0, sticky="ew", pady=(8, 0))
        ttk.Combobox(add_form, textvariable=style_type_var, values=["문단", "글자"], state="readonly", width=6).pack(side="left")
        ttk.Entry(add_form, textvariable=style_name_var, width=20).pack(side="left", fill="x", expand=True, padx=(6, 0))
        ttk.Checkbutton(add_form, text="표", variable=table_style_var).pack(side="left", padx=(6, 0))
        ttk.Checkbutton(add_form, text="캡션", variable=caption_style_var).pack(side="left")

        marker_form = ttk.Frame(style_tab)
        marker_form.grid(row=2, column=0, sticky="ew", pady=(8, 0))
        ttk.Label(marker_form, text="식별자").pack(side="left")
        ttk.Entry(marker_form, textvariable=outline_markers_var, width=22).pack(side="left", fill="x", expand=True, padx=(6, 0))
        ttk.Label(marker_form, text="역할").pack(side="left", padx=(8, 0))
        style_role_combo = ttk.Combobox(
            marker_form,
            textvariable=roles_var,
            values=role_selection_values(style_type_var.get()),
            state="readonly",
            width=16,
        )
        style_role_combo.pack(side="left", padx=(6, 0))
        style_role_combo_ref.append(style_role_combo)
        ttk.Label(marker_form, text="번호그룹").pack(side="left", padx=(8, 0))
        ttk.Combobox(
            marker_form,
            textvariable=numbering_group_var,
            values=NUMBERING_GROUP_LABELS,
            state="readonly",
            width=8,
        ).pack(side="left", padx=(6, 0))
        ttk.Label(marker_form, text="수준").pack(side="left", padx=(8, 0))
        ttk.Entry(marker_form, textvariable=numbering_level_var, width=4).pack(side="left", padx=(6, 0))
        ttk.Checkbutton(marker_form, text="상위 개요 뒤 새 번호", variable=numbering_restart_var).pack(side="left", padx=(8, 0))

        style_template_form = ttk.Frame(style_tab)
        style_template_form.grid(row=3, column=0, sticky="ew", pady=(8, 0))
        ttk.Label(style_template_form, text="이 스타일 번호박스 템플릿").pack(side="left")
        style_template_entry = ttk.Entry(style_template_form, textvariable=style_template_var, width=32)
        style_template_entry.pack(side="left", fill="x", expand=True, padx=(6, 0))
        style_template_entry_ref.append(style_template_entry)
        style_template_button = ttk.Button(style_template_form, text="찾아보기", command=browse_style_template)
        style_template_button.pack(side="left", padx=(6, 0))
        style_template_button_ref.append(style_template_button)

        item_buttons = ttk.Frame(style_tab)
        item_buttons.grid(row=4, column=0, sticky="ew", pady=(8, 0))
        ttk.Button(item_buttons, text="추가", command=start_new_style).pack(side="left")
        ttk.Button(item_buttons, text="삭제", command=delete_selected_style).pack(side="left", padx=(6, 0))
        ttk.Button(item_buttons, text="저장", command=save_style_form).pack(side="right")

        rule_form = ttk.Frame(rule_tab)
        rule_form.grid(row=1, column=0, sticky="ew", pady=(8, 0))
        ttk.Entry(rule_form, textvariable=rule_name_var, width=18).pack(side="left", fill="x", expand=True)
        ttk.Combobox(
            rule_form,
            textvariable=rule_type_var,
            values=[
                *INLINE_RULE_TYPE_LABELS,
            ],
            state="readonly",
            width=18,
        ).pack(side="left", padx=(6, 0))
        ttk.Combobox(
            rule_form,
            textvariable=rule_target_var,
            values=["body", "table", "all"],
            state="readonly",
            width=8,
        ).pack(side="left", padx=(6, 0))
        rule_role_combo = ttk.Combobox(
            rule_form,
            textvariable=rule_role_var,
            values=character_role_options(),
            state="readonly",
            width=14,
        )
        rule_role_combo.pack(side="left", padx=(6, 0))
        rule_role_combo_ref.append(rule_role_combo)

        rule_options = ttk.Frame(rule_tab)
        rule_options.grid(row=2, column=0, sticky="ew", pady=(8, 0))
        ttk.Checkbutton(rule_options, text="경계 포함", variable=rule_include_wrapper_var).pack(side="left")
        ttk.Checkbutton(rule_options, text="표식 삭제", variable=rule_remove_markers_var).pack(side="left", padx=(6, 0))
        ttk.Checkbutton(rule_options, text="사용", variable=rule_enabled_var).pack(side="left", padx=(6, 0))

        rule_buttons = ttk.Frame(rule_tab)
        rule_buttons.grid(row=3, column=0, sticky="ew", pady=(8, 0))
        ttk.Button(rule_buttons, text="추가", command=start_new_rule).pack(side="left")
        ttk.Button(rule_buttons, text="삭제", command=remove_selected_rule).pack(side="left", padx=(6, 0))
        ttk.Button(rule_buttons, text="저장", command=save_selected_rule).pack(side="right")

        buttons = ttk.Frame(body)
        buttons.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(12, 0))
        ttk.Button(buttons, text="닫기", command=close_window).pack(side="left")

        set_list.bind("<<ListboxSelect>>", select_set)
        style_tree.bind("<<TreeviewSelect>>", fill_form_from_selected_style)
        rule_tree.bind("<<TreeviewSelect>>", fill_form_from_selected_rule)
        notebook.bind("<<NotebookTabChanged>>", on_tab_changed)
        style_type_var.trace_add("write", lambda *_args: update_style_role_options())
        roles_var.trace_add("write", lambda *_args: refresh_style_template_controls())
        window.protocol("WM_DELETE_WINDOW", close_window)
        refresh_set_list()
        refresh_style_tree()
        refresh_rule_tree()
        refresh_rule_role_options()
        refresh_title_box_template_form()
        last_set_index[0] = selected_set_index()
        last_tab[0] = notebook.select()
        clear_style_form()
        clear_rule_form()

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
            label = "문단" if record.style_type == "PARA" else "글자"
            self.style_list.insert("end", f"[{label}] {record.name}")
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
        self.replace_active_style_records(self.style_records)
        save_style_sets(self.active_style_set_name, self.style_sets)
        self._refreshing = True
        self.populate_style_list(new_index)
        self._refreshing = False
        self.debug(f"[style-sets] 순서 변경 저장: {self.active_style_set_name} / {self.style_records[new_index].name}")

    def reset_style_order(self) -> None:
        default_active, default_sets = default_style_sets()
        replacement = next((item for item in default_sets if item.name == default_active), default_sets[0])
        self.replace_style_set(self.active_style_set_name, replacement)
        self.style_records = style_set_to_records(self.active_style_set())
        save_style_sets(self.active_style_set_name, self.style_sets)
        self._refreshing = True
        self.populate_style_list(0)
        self._refreshing = False
        self.build_table_tab_content()
        self.debug(f"[style-sets] 기본순서로 복원: {self.active_style_set_name}")

    def clear_selected_character_style(self) -> None:
        label = "글자스타일제거"
        if not self.ensure_hwp():
            return
        try:
            ok = self.run_hwp_command("StyleClearCharStyle")
            self.log(f"{label}: StyleClearCharStyle result={ok}")
            if not ok:
                messagebox.showwarning(label, "한글이 글자 스타일 제거를 실패로 반환했습니다. 텍스트를 선택한 상태인지 확인하세요.")
        except Exception as exc:
            self.log(f"{label} 실패: {type(exc).__name__}: {exc}")
            messagebox.showerror(f"{label} 실패", str(exc))

    def replace_style_set(self, target_name: str, replacement: StyleSet) -> None:
        updated = StyleSet(
            name=target_name,
            paragraph_styles=replacement.paragraph_styles,
            character_styles=replacement.character_styles,
            inline_rules=replacement.inline_rules,
            role_configs=replacement.role_configs,
        )
        self.style_sets = [updated if item.name == target_name else item for item in self.style_sets]

    def replace_active_style_records(self, records: list[StyleRecord]) -> None:
        old_set = self.active_style_set()
        flags = {
            ("PARA", entry.name): entry
            for entry in old_set.paragraph_styles
        } | {
            ("CHAR", entry.name): entry
            for entry in old_set.character_styles
        }
        paragraph_styles: list[StyleEntry] = []
        character_styles: list[StyleEntry] = []
        for record in records:
            old = flags.get((record.style_type, record.name))
            entry = StyleEntry(
                record.name,
                table_style=old.table_style if old is not None else False,
                caption_style=old.caption_style if old is not None else False,
                roles=old.roles if old is not None else inferred_style_roles(
                    record.name,
                    table_style=normalize_style_name(record.name).startswith(normalize_style_name("표내용-")),
                ),
                template_file=old.template_file if old is not None else "",
            )
            if record.style_type == "CHAR":
                character_styles.append(entry)
            else:
                paragraph_styles.append(entry)
        self.replace_style_set(
            self.active_style_set_name,
            StyleSet(self.active_style_set_name, paragraph_styles, character_styles, old_set.inline_rules, old_set.role_configs),
        )

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
            self.current_doc_style_path = None
            self.current_doc_style_map = {}
            self.current_doc_style_norm_map = {}
            self.debug("[hwp] COM 연결 성공")
            for line in LAST_HWP_CONNECTION_LOG:
                self.debug(line)
            return True
        except Exception as exc:
            self.log(f"한글 COM 연결 실패: {type(exc).__name__}: {exc}")
            messagebox.showerror("한글 연결 실패", str(exc))
            return False

    def ensure_current_doc_style_cache(self) -> None:
        if self.__dict__.get("current_doc_style_map"):
            return
        self.refresh_current_doc_style_map(force=True)

    def current_doc_style_reconnect_hint(self) -> str:
        if not self.__dict__.get("current_doc_style_map"):
            return ""
        return "\n\n스타일 불러오기를 이미 했다면 상태 탭에서 [한글 다시 연결]을 누른 뒤 다시 시도하세요."

    def warm_current_doc_style_cache(self) -> None:
        if self.current_doc_style_map:
            return
        try:
            if self.hwp is None:
                self.debug("[style-map] 시작 캐시 예열: 한글 연결")
                self.hwp = connect_hwp()
                self.current_doc_style_path = None
                self.current_doc_style_map = {}
                self.current_doc_style_norm_map = {}
            self.refresh_current_doc_style_map(force=True)
            if self.current_doc_style_map:
                self.debug(f"[style-map] 시작 캐시 예열 완료: {len(self.current_doc_style_map)}개")
        except Exception as exc:
            self.debug(f"[style-map] 시작 캐시 예열 생략: {type(exc).__name__}: {exc}")

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
                self.activate_hwp_window()
                time.sleep(0.03)
                result = self.hwp.HAction.Execute(action, hset)
                ok = result is not False and result != 0
            except Exception as exc:
                self.debug(f"[hwp-action] {action} 실패: {type(exc).__name__}: {exc}")
                continue
            self.debug(f"[hwp-action] {action} result={result!r}, ok={ok}")
            if ok:
                return action, True
        return last_action, False

    def set_caption_position_top_best_effort(self) -> tuple[str, bool, bool]:
        default_ok = False
        try:
            pset = self.hwp.HParameterSet.HShapeObject
            default_ok = bool(self.hwp.HAction.GetDefault("CaptionPosTop", pset.HSet))
            if default_ok:
                action_name, ok = self.execute_first_hwp_action(("CaptionPosTop",), pset.HSet)
                if ok:
                    return action_name, True, default_ok
        except Exception as exc:
            self.debug(f"[caption-build] CaptionPosTop Execute 실패: {type(exc).__name__}: {exc}")
        run_ok = self.run_hwp_command("CaptionPosTop")
        return "CaptionPosTop", run_ok, default_ok

    def hwp_parameter_set(self, name: str):
        return getattr(self.hwp.HParameterSet, name)

    def execute_configured_hwp_action(self, action_name: str, parameter_set_name: str, values: dict[str, object]) -> tuple[bool, bool]:
        pset = self.hwp_parameter_set(parameter_set_name)
        default_ok = bool(self.hwp.HAction.GetDefault(action_name, pset.HSet))
        for item, value in values.items():
            self.set_parameter_item(pset, item, value, action_name)
            try:
                self.set_parameter_item(pset.HSet, item, value, action_name)
            except Exception:
                pass
        executed = bool(self.hwp.HAction.Execute(action_name, pset.HSet))
        return default_ok, executed

    def hide_current_page_header_footer_page_number(self) -> None:
        label = "쪽 모두 감추기"
        if not self.ensure_hwp():
            return
        fields = 0x01 | 0x02 | 0x20
        try:
            self.activate_hwp_window()
            default_ok, executed = self.execute_configured_hwp_action("PageHiding", "HPageHiding", {"Fields": fields})
            self.log(f"{label}: Fields={fields:#04x}, GetDefault={default_ok}, result={executed}")
            if not executed:
                messagebox.showwarning(
                    label,
                    "한글이 쪽 감추기 명령을 실패로 반환했습니다. 본문 쪽에 커서를 둔 상태인지 확인하세요.",
                )
        except Exception as exc:
            self.log(f"{label} 실패: {type(exc).__name__}: {exc}")
            messagebox.showerror(f"{label} 실패", str(exc))

    def insert_new_number_object(self, num_type: int, label: str) -> None:
        if not self.ensure_hwp():
            return
        try:
            self.activate_hwp_window()
            default_ok, executed = self.execute_configured_hwp_action(
                "NewNumber",
                "HAutoNum",
                {
                    "NumType": num_type,
                    "NewNumber": 1,
                },
            )
            self.log(f"{label}: NumType={num_type}, NewNumber=1, GetDefault={default_ok}, result={executed}")
            if not executed:
                messagebox.showwarning(
                    label,
                    "한글이 새 번호 개체 삽입을 실패로 반환했습니다. 번호를 넣을 위치에 커서를 둔 상태인지 확인하세요.",
                )
        except Exception as exc:
            self.log(f"{label} 실패: {type(exc).__name__}: {exc}")
            messagebox.showerror(f"{label} 실패", str(exc))

    def insert_header_footer_object(self, header_footer_ctrl_type: int, location_type: int, label: str) -> None:
        if not self.ensure_hwp():
            return
        last_exc: Exception | None = None
        for parameter_set_name in ("HHeaderFooter", "HeaderFooter"):
            try:
                self.activate_hwp_window()
                default_ok, executed = self.execute_configured_hwp_action(
                    "HeaderFooter",
                    parameter_set_name,
                    {
                        "HeaderFooterCtrlType": header_footer_ctrl_type,
                        "Type": location_type,
                    },
                )
                self.log(
                    f"{label}: parameter_set={parameter_set_name}, "
                    f"HeaderFooterCtrlType={header_footer_ctrl_type}, Type={location_type}, "
                    f"GetDefault={default_ok}, result={executed}"
                )
                if not executed:
                    messagebox.showwarning(
                        label,
                        "한글이 머리말 개체 삽입을 실패로 반환했습니다. 본문 쪽에 커서를 둔 상태인지 확인하세요.",
                    )
                return
            except AttributeError as exc:
                last_exc = exc
                self.debug(f"[header-footer] {parameter_set_name} 없음: {exc}")
                continue
            except Exception as exc:
                last_exc = exc
                self.log(f"{label} 실패: {type(exc).__name__}: {exc}")
                break
        if last_exc is not None:
            messagebox.showerror(f"{label} 실패", str(last_exc))

    def insert_custom_square_composed_number(self) -> None:
        raw_value = self.compose_number_var.get().strip()
        if not raw_value.isdecimal():
            messagebox.showwarning("겹치기글자", "숫자만 입력해 주세요.")
            return
        self.insert_square_composed_number(int(raw_value))

    def insert_square_composed_number(self, number: int) -> None:
        self.insert_square_composed_chars(str(number), f"겹치기글자 □{number}")

    def insert_checked_box_symbol(self) -> None:
        label = "체크박스 ☑"
        if not self.ensure_hwp():
            return
        try:
            self.activate_hwp_window()
            executed = self.insert_hwp_text("☑")
            self.log(f"{label}: Chars=☑, result={executed}")
            if not executed:
                messagebox.showwarning(
                    label,
                    "한글이 체크박스 문자 삽입을 실패로 반환했습니다. 입력 가능한 본문 위치에 커서를 둔 상태인지 확인하세요.",
                )
        except Exception as exc:
            self.log(f"{label} 실패: {type(exc).__name__}: {exc}")
            messagebox.showerror(f"{label} 실패", str(exc))

    def insert_square_composed_standard_check_mark(self) -> None:
        self.insert_square_composed_chars("✓", "체크박스 □✓")

    def insert_square_composed_check_mark(self) -> None:
        self.insert_square_composed_chars("✔", "체크박스 □✔")

    def insert_square_composed_chars(self, chars: str, label: str) -> None:
        if not self.ensure_hwp():
            return
        try:
            self.activate_hwp_window()
            pset = self.hwp_parameter_set("HChCompose")
            default_ok = bool(self.hwp.HAction.GetDefault("ComposeChars", pset.HSet))
            values = {
                "Chars": chars,
                "CircleType": 3,
                "CharSize": -3,
                "CheckCompose": 0,
                "TextDir": 0,
            }
            for item, value in values.items():
                self.set_parameter_item(pset, item, value, "ComposeChars")
                try:
                    self.set_parameter_item(pset.HSet, item, value, "ComposeChars")
                except Exception:
                    pass
            executed = bool(self.hwp.HAction.Execute("ComposeChars", pset.HSet))
            self.log(
                f"{label}: Chars={chars}, CircleType=3, CharSize=-3, "
                f"CheckCompose=0, TextDir=0, GetDefault={default_ok}, result={executed}"
            )
            if not executed:
                messagebox.showwarning(
                    label,
                    "한글이 사각형 겹치기글자 삽입을 실패로 반환했습니다. 입력 가능한 본문 위치에 커서를 둔 상태인지 확인하세요.",
                )
        except Exception as exc:
            self.log(f"{label} 실패: {type(exc).__name__}: {exc}")
            messagebox.showerror(f"{label} 실패", str(exc))

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
        result = self.hwp.HAction.Execute("CharShape", pset.HSet)
        ok = result is not False and result != 0
        return ok, default_ok

    def configure_style_apply_set(self, pset, style_id: int) -> None:
        pset.Apply = style_id

    def execute_style_record(self, record: StyleRecord) -> bool:
        pset = self.hwp.HParameterSet.HStyle
        self.hwp.HAction.GetDefault("StyleEx", pset.HSet)
        self.configure_style_apply_set(pset, record.style_id)
        action, ok = self.execute_first_hwp_action(("StyleEx", "Style"), pset.HSet)
        self.debug(f"[style] action={action}, apply={record.style_id}, result={ok}")
        return ok

    def apply_style_by_name(self, style_name: str) -> bool:
        self.ensure_current_doc_style_cache()
        record = self.find_current_doc_style_record(style_name)
        if record is None:
            self.debug(f"[title-cell] 현재 문서에서 스타일 이름 매칭 실패, 적용 중단: {style_name}")
            messagebox.showwarning(
                "스타일 이름 매칭 실패",
                f"현재 문서에서 같은 이름의 스타일을 찾지 못해 적용을 중단했습니다.\n\n"
                f"찾는 스타일: {style_name}\n\n"
                "스타일 번호로 대신 적용하면 다른 스타일이 적용될 수 있습니다."
                f"{self.current_doc_style_reconnect_hint()}",
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
        self.ensure_current_doc_style_cache()
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
        return ("CellZoneBorderFill", "CellZoneBorder", "CellBorderFill", "CellBorder")

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

    def apply_table_style_preset_steps(self, preset_name: str) -> tuple[list[str], bool, str | None]:
        results: list[str] = []
        base_preset_name = {
            "thin_top_bottom_header_double": "thin_top_bottom",
            "thick_top_bottom_header_double": "thick_top_bottom",
        }.get(preset_name, preset_name)

        all_cells_ok = self.select_current_table_all_cells()
        results.append(f"전체셀선택={all_cells_ok}")
        if not all_cells_ok:
            return results, False, "표 안에 커서를 둔 상태에서 실행하세요."

        border_action, border_ok = self.set_table_border_preset(base_preset_name)
        results.append(f"표스타일={border_ok}({border_action}:{base_preset_name})")
        if not border_ok:
            return results, False, "표 전체 스타일 적용에 실패했습니다."

        header_ok, header_steps = self.select_current_table_first_row()
        results.extend(header_steps)
        results.append(f"헤더선택={header_ok}")
        if not header_ok:
            return results, False, "표 스타일은 적용했지만 헤더 셀을 선택하지 못했습니다."

        title = self.table_settings["title_cell"]
        try:
            header_action, header_set_ok, header_action_ok = self.set_selected_cells_as_header()
            results.append(f"제목셀속성=set:{header_set_ok}, action:{header_action_ok}({header_action})")
        except Exception as exc:
            results.append(f"제목셀속성 실패:{type(exc).__name__}")
            self.debug(f"[table-style] 제목셀속성 실패: {exc}")

        try:
            para_style = title["paragraph_style"]
            para_ok = self.apply_style_by_name(para_style)
            results.append(f"문단스타일={para_ok}({para_style})")
        except Exception as exc:
            results.append(f"문단스타일 실패:{type(exc).__name__}")
            self.debug(f"[table-style] 문단스타일 실패: {exc}")

        char_style = title.get("character_style", "(적용 안 함)")
        if char_style and char_style != "(적용 안 함)":
            try:
                char_ok = self.apply_style_by_name(char_style)
                results.append(f"글자스타일={char_ok}({char_style})")
            except Exception as exc:
                results.append(f"글자스타일 실패:{type(exc).__name__}")
                self.debug(f"[table-style] 글자스타일 실패: {exc}")

        if preset_name in {"thin_top_bottom_header_double", "thick_top_bottom_header_double"}:
            header_border_action, header_border_ok = self.set_table_border_preset("double_bottom")
            results.append(f"헤더아래이중선={header_border_ok}({header_border_action})")
            if not header_border_ok:
                return results, False, "헤더 아래쪽 이중선 적용에 실패했습니다."

        color_name = title["background_color"]
        bg_action, bg_ok, bg_default = self.set_cell_fill_color(self.palette_color(color_name))
        results.append(f"헤더배경={bg_ok}({bg_action}:{color_name}, default={bg_default})")
        if not bg_ok:
            return results, False, "헤더 배경색 적용에 실패했습니다."
        return results, True, None

    def apply_table_style_preset(self, preset_name: str) -> None:
        label = "표 스타일 적용"
        if not self.ensure_hwp():
            return
        try:
            results, ok, warning = self.apply_table_style_preset_steps(preset_name)
            self.log(f"{label}: {preset_name}, " + ", ".join(results))
            if not ok and warning:
                messagebox.showwarning(label, warning)
        except Exception as exc:
            self.log(f"{label} 실패: {type(exc).__name__}: {exc}")
            messagebox.showerror(f"{label} 실패", str(exc))

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

            caption_pos_action, caption_pos_ok, caption_pos_default = self.set_caption_position_top_best_effort()
            self.debug(
                f"[caption-build] caption position top action={caption_pos_action}, "
                f"default={caption_pos_default}, result={caption_pos_ok}"
            )

            parts = self.pending_caption_parts
            remove_tail_space_ok = self.run_hwp_command("DeleteBack")
            suffix_ok = remove_tail_space_ok and (not parts.suffix or self.insert_hwp_text(parts.suffix))

            begin_action, begin_ok = "MoveLineBegin", self.run_hwp_command("MoveLineBegin")
            remove_label_ok = begin_ok and self.delete_hwp_chars(2)
            prefix_ok = remove_label_ok and (not parts.prefix or self.insert_hwp_text(parts.prefix))

            style_ok = False
            if prefix_ok and suffix_ok:
                caption_record = self.caption_style_record()
                if caption_record is not None:
                    style_ok = self.apply_style_by_name(caption_record.name)
                else:
                    style_ok = self.apply_first_style_containing("캡션")
                moved_text = self.pending_caption_title
                self.pending_caption_title = ""
                self.pending_caption_parts = CaptionParts(prefix="", suffix="")
                self.activate_hwp_window()
                self.log(
                    f"{label}: caption={action}:{caption_ok}, begin={begin_action}:{begin_ok}, "
                    f"caption_pos_top={caption_pos_action}:{caption_pos_ok}, "
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

    def apply_table_cell_margins(self) -> None:
        if not self.ensure_hwp():
            return
        try:
            action, ok = self.set_table_cell_margins()
            margins = self.table_settings["table_margins"]
            self.log(
                "셀 안 여백 적용: "
                f"좌{margins['cell_left']} 우{margins['cell_right']} "
                f"상{margins['cell_top']} 하{margins['cell_bottom']}mm, "
                f"action={action}, result={ok}"
            )
            if not ok:
                messagebox.showwarning("셀 안 여백 적용 확인", "표 셀 안에 커서를 두거나 셀 범위를 선택한 상태인지 확인하세요.")
        except Exception as exc:
            self.log(f"셀 안 여백 적용 실패: {type(exc).__name__}: {exc}")
            messagebox.showerror("셀 안 여백 적용 실패", str(exc))

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
        self.activate_hwp_window()
        time.sleep(0.03)
        if self.run_hwp_command("SelectCtrlReverse"):
            time.sleep(0.03)
            return True
        return False

    def select_selected_table_first_cell(self) -> bool:
        self.activate_hwp_window()
        time.sleep(0.03)
        if self.run_hwp_command("ShapeObjTableSelCell"):
            time.sleep(0.03)
            return True
        return False

    def is_in_table_cell(self) -> bool:
        return self.get_current_cell_address() is not None

    def move_to_nearby_table_cell(self) -> bool:
        if self.is_in_table_cell():
            return True
        self.activate_hwp_window()
        time.sleep(0.03)
        for command in ("MoveLeft", "MoveUp", "MoveRight"):
            for _ in range(3):
                if not self.run_hwp_command(command):
                    break
                time.sleep(0.03)
                if self.is_in_table_cell():
                    self.debug(f"[table-cell-select] nearby cell entered by {command}")
                    return True
        return False

    def current_selected_table_dimensions(self) -> tuple[int, int] | None:
        try:
            table_xml = safe_str(self.hwp.GetTextFile("HWPML2X", "saveblock"))
        except Exception as exc:
            self.debug(f"[table-dimensions] HWPML 읽기 실패: {type(exc).__name__}: {exc}")
            return None
        if not table_xml:
            return None
        try:
            root = ET.fromstring(table_xml)
        except ET.ParseError as exc:
            self.debug(f"[table-dimensions] HWPML 파싱 실패: {exc}")
            return None
        table = root.find(".//TABLE")
        if table is None:
            return None

        rows = table.findall(".//ROW")
        if not rows:
            cells = table.findall(".//CELL")
            return (1, len(cells)) if cells else None

        def span_value(cell, *names: str) -> int:
            for name in names:
                try:
                    value = int(cell.get(name) or 0)
                except (TypeError, ValueError):
                    value = 0
                if value > 0:
                    return value
            return 1

        col_count = 0
        for row in rows:
            cells = row.findall(".//CELL")
            col_count = max(col_count, sum(span_value(cell, "ColSpan", "colSpan", "colspan") for cell in cells))
        if col_count <= 0:
            return None
        return len(rows), col_count

    def select_current_table_all_cells(self, row_count: int | None = None, col_count: int | None = None) -> bool:
        self.activate_hwp_window()
        time.sleep(0.03)
        if not self.has_selected_object():
            self.select_current_table_object()
        self.select_selected_table_first_cell()
        if not self.is_in_table_cell():
            self.move_to_nearby_table_cell()

        if row_count is not None and col_count is not None and row_count > 0 and col_count > 0:
            self.activate_hwp_window()
            time.sleep(0.03)
            block_ok = self.run_hwp_command("TableCellBlock")
            time.sleep(0.03)
            if not block_ok and not self.is_selected_cell_block():
                self.debug("[table-cell-select] TableCellBlock 실패 및 셀 블록 상태 아님")
                return False
            if not self.run_hwp_command("TableCellBlockExtend"):
                return False
            time.sleep(0.03)
            for _ in range(max(0, row_count - 1)):
                if not self.run_hwp_command("TableLowerCell"):
                    return False
                time.sleep(0.03)
            for _ in range(max(0, col_count - 1)):
                if not self.run_hwp_command("TableRightCell"):
                    return False
                time.sleep(0.03)
            if not self.is_selected_cell_block():
                self.debug(
                    "[table-cell-select] SelectionMode 확인 실패, "
                    f"rows={row_count}, cols={col_count}, 명령 성공 기준으로 계속 진행"
                )
            return True

        self.activate_hwp_window()
        time.sleep(0.03)
        block_ok = self.run_hwp_command("TableCellBlock")
        time.sleep(0.03)
        if not block_ok and not self.is_selected_cell_block():
            self.debug("[table-cell-select] TableCellBlock fallback 실패 및 셀 블록 상태 아님")
            return False
        row_ok = self.run_hwp_command("TableCellBlockRow")
        time.sleep(0.03)
        col_ok = row_ok and self.run_hwp_command("TableCellBlockCol")
        time.sleep(0.03)
        extend_ok = col_ok and self.run_hwp_command("TableCellBlockExtend")
        time.sleep(0.03)
        if extend_ok:
            if not self.is_selected_cell_block():
                self.debug("[table-cell-select] SelectionMode 확인 실패, 셀블록/줄블록/칸블록/연장 명령 성공 기준으로 계속 진행")
            return True
        return False

    def select_current_table_all_cells_by_variant(
        self,
        *,
        row_count: int | None = None,
        col_count: int | None = None,
        horizontal_first: bool = True,
        extend_command: str = "TableCellBlockExtend",
    ) -> tuple[bool, list[str]]:
        steps: list[str] = []
        self.activate_hwp_window()
        time.sleep(0.03)
        if not self.has_selected_object():
            selected = self.select_current_table_object()
            steps.append(f"SelectCtrlReverse={selected}")
        if row_count is None or col_count is None:
            dimensions = self.current_selected_table_dimensions()
            steps.append(f"dimensions={dimensions}")
            if dimensions is not None:
                row_count, col_count = dimensions

        first_cell = self.select_selected_table_first_cell()
        steps.append(f"ShapeObjTableSelCell={first_cell}")
        if not first_cell and self.move_to_nearby_table_cell():
            first_cell = True
            steps.append("nearby_cell=True")
        if not first_cell:
            return False, steps
        if not row_count or not col_count or row_count <= 0 or col_count <= 0:
            steps.append(f"invalid_size={row_count}x{col_count}")
            return False, steps

        block = self.run_hwp_command("TableCellBlock")
        block_ready = bool(block or self.is_selected_cell_block())
        steps.append(f"TableCellBlock={block}, ready={block_ready}")
        if not block_ready:
            return False, steps

        extend = self.run_hwp_command(extend_command)
        steps.append(f"{extend_command}={extend}")
        if not extend:
            return False, steps

        commands = (
            (("TableRightCell", max(0, col_count - 1)), ("TableLowerCell", max(0, row_count - 1)))
            if horizontal_first
            else (("TableLowerCell", max(0, row_count - 1)), ("TableRightCell", max(0, col_count - 1)))
        )
        for command, count in commands:
            for index in range(count):
                ok = self.run_hwp_command(command)
                steps.append(f"{command}[{index + 1}/{count}]={ok}")
                if not ok:
                    return False, steps
                time.sleep(0.03)
        return True, steps

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

    def get_parameter_item_value(self, target, item: str):
        for candidate in (target, getattr(target, "HSet", None)):
            if candidate is None:
                continue
            try:
                return candidate.Item(item)
            except Exception:
                pass
            try:
                return getattr(candidate, item)
            except Exception:
                pass
        return None

    def set_selected_ctrl_property_item(self, ctrl, item: str, value, log_prefix: str) -> tuple[bool, object | None, object | None]:
        try:
            props = ctrl.Properties
        except Exception as exc:
            self.debug(f"[{log_prefix}] Ctrl.Properties 읽기 실패: {type(exc).__name__}: {exc}")
            return False, None, None

        before = self.get_parameter_item_value(props, item)
        set_ok = self.set_parameter_item(props, item, value, log_prefix)
        after = self.get_parameter_item_value(props, item)
        commit_ok = False
        if set_ok:
            try:
                ctrl.Properties = props
                commit_ok = True
            except Exception as exc:
                self.debug(f"[{log_prefix}] Ctrl.Properties 커밋 실패: {type(exc).__name__}: {exc}")
        return set_ok and commit_ok, before, after

    def nearby_caption_autonum_ctrls(self, base_ctrl, max_steps: int = 8) -> list:
        found = []
        seen: set[int] = set()
        for attr in ("Next", "Prev"):
            ctrl = base_ctrl
            for _ in range(max_steps):
                try:
                    ctrl = getattr(ctrl, attr)
                except Exception:
                    break
                if ctrl is None:
                    break
                identity = id(ctrl)
                if identity in seen:
                    break
                seen.add(identity)
                ctrl_id = safe_str(getattr(ctrl, "CtrlID", ""))
                if ctrl_id == "atno":
                    found.append(ctrl)
                    break
        return found

    def set_nearby_caption_autonum_type(self, base_ctrl, numbering_type: int) -> tuple[int, int, list[str]]:
        autonum_type_by_numbering_type = {
            1: 3,  # 그림 번호
            2: 4,  # 표 번호
        }
        autonum_type = autonum_type_by_numbering_type.get(numbering_type)
        if autonum_type is None:
            return 0, 0, ["skipped"]

        changed = 0
        scanned = 0
        details: list[str] = []
        for ctrl in self.nearby_caption_autonum_ctrls(base_ctrl):
            scanned += 1
            ctrl_id = safe_str(getattr(ctrl, "CtrlID", ""))
            user_desc = safe_str(getattr(ctrl, "UserDesc", ""))
            ok, before, after = self.set_selected_ctrl_property_item(
                ctrl,
                "NumType",
                autonum_type,
                "caption-autonum",
            )
            if ok:
                changed += 1
            details.append(f"{ctrl_id}:{user_desc}:before={before}:after={after}:ok={ok}")
        return changed, scanned, details

    def set_selected_object_numbering_type(
        self,
        numbering_type: int,
    ) -> tuple[str, bool, object | None, object | None, tuple[int, int, list[str]]]:
        self.activate_hwp_window()
        time.sleep(0.03)
        selected = self.has_selected_object()
        if not selected:
            selected = self.select_current_table_object()

        try:
            ctrl = self.hwp.CurSelectedCtrl
        except Exception:
            ctrl = None
        if ctrl is None:
            self.debug(f"[object-numbering] 선택 개체 없음, SelectCtrlReverse={selected}")
            return "CurSelectedCtrl.Properties", False, None, None, (0, 0, [])

        ctrl_id = safe_str(getattr(ctrl, "CtrlID", ""))
        object_ok, before, after = self.set_selected_ctrl_property_item(
            ctrl,
            "NumberingType",
            numbering_type,
            "object-numbering",
        )
        autonum_result = self.set_nearby_caption_autonum_type(ctrl, numbering_type)
        try:
            self.hwp.Refresh()
        except Exception:
            pass
        self.debug(
            "[object-numbering] "
            f"selected={selected}, CtrlID={ctrl_id!r}, before={before}, "
            f"target={numbering_type}, after={after}, object_ok={object_ok}, "
            f"caption_autonum={autonum_result}"
        )
        return "CurSelectedCtrl.Properties", object_ok, before, after, autonum_result

    def apply_selected_object_numbering_type(self, numbering_type: int, label: str) -> None:
        title = "번호종류"
        if not self.ensure_hwp():
            return
        try:
            action_name, ok, before, after, autonum_result = self.set_selected_object_numbering_type(numbering_type)
            autonum_changed, autonum_scanned, _autonum_details = autonum_result
            self.log(
                f"{title} 적용: {label}({numbering_type}), "
                f"before={before}, after={after}, action={action_name}, "
                f"caption_autonum={autonum_changed}/{autonum_scanned}, result={ok}"
            )
            if not ok:
                messagebox.showwarning(
                    f"{title} 적용 확인",
                    "한글이 번호 종류 변경을 실패로 반환했습니다.\n표나 그림 개체를 선택한 상태인지 확인하세요.",
                )
        except Exception as exc:
            self.log(f"{title} 적용 실패: {type(exc).__name__}: {exc}")
            messagebox.showerror(f"{title} 적용 실패", str(exc))

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

    def table_cell_margin_values(self) -> dict[str, int]:
        margins = self.table_settings["table_margins"]
        return {
            "MarginLeft": self.mm_to_hwpunit(margins["cell_left"]),
            "MarginRight": self.mm_to_hwpunit(margins["cell_right"]),
            "MarginTop": self.mm_to_hwpunit(margins["cell_top"]),
            "MarginBottom": self.mm_to_hwpunit(margins["cell_bottom"]),
        }

    def table_default_cell_margin_values(self) -> dict[str, int]:
        margins = self.table_settings["table_margins"]
        return {
            "CellMarginLeft": self.mm_to_hwpunit(margins["cell_left"]),
            "CellMarginRight": self.mm_to_hwpunit(margins["cell_right"]),
            "CellMarginTop": self.mm_to_hwpunit(margins["cell_top"]),
            "CellMarginBottom": self.mm_to_hwpunit(margins["cell_bottom"]),
        }

    def set_shape_table_cell_margins(self, shape_cell, log_prefix: str) -> bool:
        ok = self.set_parameter_item(shape_cell, "HasMargin", 1, log_prefix)
        for item, value in self.table_cell_margin_values().items():
            applied = self.set_parameter_item(shape_cell, item, value, log_prefix)
            self.debug(f"[{log_prefix}] ShapeTableCell.{item}={value}, set={applied}")
            ok = applied and ok
        return ok

    def set_table_default_cell_margins(self, target, log_prefix: str) -> bool:
        ok = True
        for item, value in self.table_default_cell_margin_values().items():
            applied = self.set_parameter_item(target, item, value, log_prefix)
            self.debug(f"[{log_prefix}] {item}={value}, set={applied}")
            ok = applied and ok
        return ok

    def set_table_cell_margins_by_cell_shape(self) -> tuple[str, bool]:
        try:
            cell_shape = self.hwp.CellShape
        except Exception as exc:
            self.debug(f"[cell-margin-cell-shape] CellShape 읽기 실패: {type(exc).__name__}: {exc}")
            return "CellShape", False

        targets = []
        cell = self.get_parameter_child(cell_shape, "Cell", "cell-margin-cell-shape")
        if cell is not None:
            targets.append(cell)

        set_ok = False
        for target in targets:
            set_ok = self.set_shape_table_cell_margins(target, "cell-margin-cell-shape") or set_ok
        set_ok = self.set_table_default_cell_margins(cell_shape, "cell-margin-cell-shape") or set_ok
        if cell is not None:
            self.set_parameter_item(cell_shape, "Cell", cell, "cell-margin-cell-shape")
        commit_ok = self.commit_cell_shape(cell_shape)
        self.debug(f"[cell-margin-cell-shape] set={set_ok}, commit={commit_ok}")
        return "CellShape", set_ok and commit_ok

    def set_table_cell_margins_by_create_action(self) -> tuple[str, bool]:
        action = self.hwp.CreateAction("TablePropertyDialog")
        pset = action.CreateSet()
        default_ok = bool(action.GetDefault(pset))
        self.set_parameter_item(pset, "ShapeType", 3, "cell-margin-create-action")
        self.set_parameter_item(pset, "ShapeCellSize", 0, "cell-margin-create-action")
        self.set_table_default_cell_margins(pset, "cell-margin-create-action")
        shape_cell = pset.Item("ShapeTableCell")
        set_ok = self.set_shape_table_cell_margins(shape_cell, "cell-margin-create-action")
        executed = bool(action.Execute(pset))
        self.debug(
            f"[cell-margin-create-action] GetDefault={default_ok}, "
            f"set={set_ok}, Execute={executed}"
        )
        return "CreateAction(TablePropertyDialog)", default_ok and set_ok and executed

    def set_table_cell_margins_by_parameter_set(self) -> tuple[str, bool]:
        pset = self.hwp.HParameterSet.HShapeObject
        default_ok = bool(self.hwp.HAction.GetDefault("TablePropertyDialog", pset.HSet))
        self.set_parameter_item(pset.HSet, "ShapeType", 3, "cell-margin-parameter-set")
        self.set_parameter_item(pset.HSet, "ShapeCellSize", 0, "cell-margin-parameter-set")
        self.set_table_default_cell_margins(pset, "cell-margin-parameter-set")
        shape_cell = self.get_parameter_child(pset, "ShapeTableCell", "cell-margin-parameter-set")
        if shape_cell is None:
            return "TablePropertyDialog", False
        set_ok = self.set_shape_table_cell_margins(shape_cell, "cell-margin-parameter-set")
        action_name, executed = self.execute_first_hwp_action(("TablePropertyDialog", "ShapeObjDialog"), pset.HSet)
        self.debug(
            f"[cell-margin-parameter-set] GetDefault={default_ok}, "
            f"set={set_ok}, action={action_name}, Execute={executed}"
        )
        return action_name, default_ok and set_ok and executed

    def set_table_cell_margins(self) -> tuple[str, bool]:
        results: list[tuple[str, bool]] = []
        action_name, ok = self.set_table_cell_margins_by_cell_shape()
        results.append((action_name, ok))
        try:
            action_name, ok = self.set_table_cell_margins_by_create_action()
            results.append((action_name, ok))
        except Exception as exc:
            self.debug(f"[cell-margin-create-action] 실패: {type(exc).__name__}: {exc}")
        try:
            action_name, ok = self.set_table_cell_margins_by_parameter_set()
            results.append((action_name, ok))
        except Exception as exc:
            self.debug(f"[cell-margin-parameter-set] 실패: {type(exc).__name__}: {exc}")
        self.debug(f"[cell-margin] results={results}")
        if any(ok for _action_name, ok in results):
            return "+".join(action_name for action_name, _ok in results), True
        return (results[-1][0] if results else "TablePropertyDialog"), False


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

    def select_current_table_first_row(self) -> tuple[bool, list[str]]:
        results: list[str] = []

        cancel_ok = self.run_hwp_command("Cancel")
        results.append(f"Cancel={cancel_ok}")
        time.sleep(0.03)

        object_ok = self.select_current_table_object()
        results.append(f"SelectCtrlReverse={object_ok}")

        first_cell_ok = object_ok and self.select_selected_table_first_cell()
        if not first_cell_ok and self.move_to_nearby_table_cell():
            first_cell_ok = True
            results.append("nearby_cell=True")
        else:
            results.append(f"ShapeObjTableSelCell={first_cell_ok}")

        block_ok = first_cell_ok and self.run_hwp_command("TableCellBlock")
        block_ready = bool(block_ok or self.is_selected_cell_block())
        results.append(f"TableCellBlock={block_ok}, ready={block_ready}")
        time.sleep(0.03)

        row_ok = block_ready and self.run_hwp_command("TableCellBlockRow")
        results.append(f"TableCellBlockRow={row_ok}")
        time.sleep(0.03)

        return row_ok, results

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
            self.current_doc_style_path = None
            self.current_doc_style_map = {}
            self.current_doc_style_norm_map = {}
            self.refresh_current_doc_style_map(force=True)
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
            self.refresh_current_doc_style_map(force=True)
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
            self.refresh_current_doc_style_map(force=True)
            if ok:
                messagebox.showinfo(
                    "테스트 문서 열림",
                    "서식 파일 복사본을 열었습니다.\n한글 문서에서 문단을 클릭한 뒤 스타일 목록을 클릭해보세요.",
                )
        except Exception as exc:
            self.log(f"서식 테스트 문서 열기 실패: {type(exc).__name__}: {exc}")
            messagebox.showerror("서식 테스트 문서 열기 실패", str(exc))

    def on_style_selected(self, _event=None) -> None:
        if self._refreshing or self._applying_style:
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
        if not force and path == self.current_doc_style_path and self.current_doc_style_map:
            return

        self.current_doc_style_path = path
        self.current_doc_style_map = {}
        self.current_doc_style_norm_map = {}

        records = self.read_current_doc_style_records_from_memory()
        if not records:
            if path is None:
                self.debug("[style-map] 현재 문서 경로가 없고 메모리 스타일도 없어 이름 매칭 불가")
                return
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
        self.debug(f"[style] 요청 reason={reason}, set={self.active_style_set_name}, type={record.style_type}, name={record.name}")
        if record.style_type != "PARA":
            self.debug(f"[style] 글자 스타일 적용 시도: type={record.style_type}, name={record.name}")

        self._applying_style = True
        try:
            if self.hwp is None:
                self.debug("[hwp] COM 연결 시작")
                self.hwp = connect_hwp()
                self.current_doc_style_path = None
                self.current_doc_style_map = {}
                self.current_doc_style_norm_map = {}
                self.debug("[hwp] COM 연결 성공")
                for line in LAST_HWP_CONNECTION_LOG:
                    self.debug(line)
            self.ensure_current_doc_style_cache()
            target_record = self.find_current_doc_style_record(record.name)
            if target_record is None:
                current_path = self.get_current_hwp_path()
                self.debug(
                    f"[style] 현재 문서 스타일 이름 매칭 실패, 적용 중단: "
                    f"name={record.name}, current_doc={current_path}, set={self.active_style_set_name}"
                )
                messagebox.showwarning(
                    "스타일 이름 매칭 실패",
                    "현재 문서에서 같은 이름의 스타일을 찾지 못해 적용을 중단했습니다.\n\n"
                    "스타일 번호로 대신 적용하면 다른 스타일이 적용될 수 있습니다."
                    f"{self.current_doc_style_reconnect_hint()}",
                )
                return

            ok = self.execute_style_record(target_record)
            self.log(
                "스타일 적용 요청: "
                f"name={record.name}, set={self.active_style_set_name}, 현재문서id={target_record.style_id}, "
                f"현재문서index={target_record.style_index}, "
                f"type={target_record.style_type}, result={ok}"
            )
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

    def insert_file_at_cursor(self, path: Path, *, keep_section: int = 0) -> bool:
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
        option.KeepSection = keep_section
        option.KeepCharshape = 1
        option.KeepParashape = 1
        option.KeepStyle = 1
        return bool(self.hwp.HAction.Execute("InsertFile", option.HSet))

    def table_creation_parameter_set(self):
        for name in ("HTableCreation", "TableCreation"):
            try:
                return getattr(self.hwp.HParameterSet, name)
            except Exception:
                continue
        raise AttributeError("HParameterSet.TableCreation")

    def create_table_at_cursor(self, rows: int, cols: int) -> tuple[str, bool]:
        pset = self.table_creation_parameter_set()
        default_ok = False
        try:
            default_ok = bool(self.hwp.HAction.GetDefault("TableCreate", pset.HSet))
        except Exception as exc:
            self.debug(f"[table-create] GetDefault 실패: {type(exc).__name__}: {exc}")
        for item, value in (("Rows", rows), ("Cols", cols)):
            self.set_parameter_item(pset, item, value, "table-create")
            self.set_parameter_item(pset.HSet, item, value, "table-create")
        action_name = "TableCreate"
        executed = False
        for action in ("TableCreate", "TableInsert"):
            action_name = action
            try:
                self.activate_hwp_window()
                time.sleep(0.03)
                result = self.hwp.HAction.Execute(action, pset.HSet)
                executed = result is not False and result != 0
                self.debug(f"[table-create] {action} result={result!r}, ok={executed}")
            except Exception as exc:
                self.debug(f"[table-create] {action} 실패: {type(exc).__name__}: {exc}")
                continue
            if executed:
                break
        self.debug(f"[table-create] rows={rows}, cols={cols}, GetDefault={default_ok}, action={action_name}, result={executed}")
        return action_name, executed

    def page_settings_values(self, settings: PageSettings, *, apply_to: int = 3) -> dict[str, int]:
        return {
            "PaperWidth": settings.paper_width,
            "PaperHeight": settings.paper_height,
            "Landscape": settings.landscape,
            "LeftMargin": settings.left_margin,
            "RightMargin": settings.right_margin,
            "TopMargin": settings.top_margin,
            "BottomMargin": settings.bottom_margin,
            "HeaderLen": settings.header_len,
            "FooterLen": settings.footer_len,
            "GutterLen": settings.gutter_len,
            "GutterType": settings.gutter_type,
            "ApplyTo": apply_to,
            "ApplyClass": 0x08 if apply_to == 3 else 0x04,
        }

    def configure_secdef_page_settings(
        self,
        sec,
        settings: PageSettings,
        *,
        apply_to: int = 3,
    ) -> dict[str, int]:
        page = sec.PageDef
        values = self.page_settings_values(settings, apply_to=apply_to)
        for item, value in values.items():
            try:
                setattr(page, item, value)
            except Exception as exc:
                self.debug(f"[page-settings] {item} 설정 실패: {type(exc).__name__}: {exc}")
        try:
            sec.PageDef = page
        except Exception as exc:
            self.debug(f"[page-settings] PageDef 반영 실패: {type(exc).__name__}: {exc}")
        return values

    def read_current_page_settings(self) -> tuple[bool, PageSettings | None]:
        try:
            sec = self.hwp.HParameterSet.HSecDef
            default_ok = bool(self.hwp.HAction.GetDefault("PageSetup", sec.HSet))
            page = sec.PageDef
            return default_ok, PageSettings(
                paper_width=int(page.PaperWidth),
                paper_height=int(page.PaperHeight),
                landscape=int(page.Landscape),
                left_margin=int(page.LeftMargin),
                right_margin=int(page.RightMargin),
                top_margin=int(page.TopMargin),
                bottom_margin=int(page.BottomMargin),
                header_len=int(page.HeaderLen),
                footer_len=int(page.FooterLen),
                gutter_len=int(page.GutterLen),
                gutter_type=int(page.GutterType),
            )
        except Exception as exc:
            self.debug(f"[page-settings] 현재 쪽 설정 읽기 실패: {type(exc).__name__}: {exc}")
            return False, None

    def current_page_settings_summary(self) -> str:
        default_ok, settings = self.read_current_page_settings()
        return f"GetDefault={default_ok}, {format_page_settings(settings)}"

    def apply_page_settings(self, settings: PageSettings, *, apply_to: int = 3) -> tuple[str, bool, str]:
        action_name = "PageMarginSetup"
        default_ok = False
        executed = False
        current_ok = False
        current_settings = None
        values = self.page_settings_values(settings, apply_to=apply_to)
        try:
            sec = self.hwp.HParameterSet.HSecDef
            default_ok = bool(self.hwp.HAction.GetDefault(action_name, sec.HSet))
            values = self.configure_secdef_page_settings(sec, settings, apply_to=apply_to)
            executed = bool(self.hwp.HAction.Execute(action_name, sec.HSet))
            current_ok, current_settings = self.read_current_page_settings()
        except Exception as exc:
            self.debug(f"[page-settings] {action_name} 실패: {type(exc).__name__}: {exc}")
        matched = page_settings_match(settings, current_settings)
        detail = (
            f"{action_name}:default={default_ok}:exec={executed}:read={current_ok}:"
            f"match={matched}:current={format_page_settings(current_settings)}"
        )
        self.debug(f"[page-settings] values={values}, {detail}")
        return action_name, matched, detail

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
            path = path.resolve()
            if not path.exists():
                messagebox.showwarning("로고 삽입 확인", f"로고 파일을 찾을 수 없습니다.\n\n{path}")
                return
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

    def insert_general_report_template(self) -> None:
        label = "보고양식 불러오기"
        if not self.ensure_hwp():
            return
        template_path = GENERAL_REPORT_TEMPLATE_FILE
        original_text = ""
        try:
            if not template_path.exists():
                messagebox.showwarning(label, f"보고양식 파일을 찾을 수 없습니다.\n\n{template_path}")
                self.log(f"{label}: 파일 없음, path={template_path}")
                return
            style_count = len(read_style_records(template_path))

            copy_ok = self.run_hwp_command("Copy")
            time.sleep(0.08)
            original_text = self.get_clipboard_text() if copy_ok else ""
            filled_fields = None
            insert_path = template_path
            self.log(
                f"{label}: 시작, template={template_path.name}, exists={template_path.exists()}, "
                f"styles={style_count}, copy={copy_ok}, input_len={len(original_text.strip())}"
            )
            if original_text.strip():
                try:
                    filled_fields = parse_report_template_input(original_text)
                except ValueError as exc:
                    messagebox.showwarning(label, str(exc))
                    self.log(f"{label}: 입력 파싱 실패, Copy={copy_ok}, error={exc}")
                    return
                insert_path = create_filled_general_report_template_file(filled_fields, template_path)
                delete_ok = self.run_hwp_command("Delete")
                if not delete_ok:
                    messagebox.showwarning(label, "입력값은 확인했지만 선택 텍스트를 삭제하지 못했습니다. 선택 영역을 확인하세요.")
                    self.log(f"{label}: 선택 영역 삭제 실패")
                    return

            insert_ok = self.insert_file_at_cursor(insert_path, keep_section=0)
            if not insert_ok:
                if filled_fields is not None:
                    self.set_clipboard_text(original_text)
                    restore_ok = self.run_hwp_command("Paste")
                else:
                    restore_ok = None
                messagebox.showwarning(label, "한글이 보고양식 삽입을 실패로 반환했습니다. 삽입할 위치에 커서를 둔 상태인지 확인하세요.")
                self.log(f"{label}: 삽입 실패, file={insert_path.name}, styles={style_count}, restored={restore_ok}")
                return
            page_settings = read_hwpx_page_settings(template_path)
            if page_settings is None:
                page_action = "read_failed"
                page_ok = False
                page_detail = "none"
                page_target = "none"
            else:
                page_target = format_page_settings(page_settings)
                page_action, page_ok, page_detail = self.apply_page_settings(page_settings, apply_to=3)
            self.refresh_current_doc_style_map(force=True)
            self.activate_hwp_window()
            if filled_fields is None:
                fill_summary = "fill=none"
            else:
                fill_summary = (
                    f"제목={filled_fields.title}, 부서={filled_fields.department}, "
                    f"날짜={filled_fields.date_text}"
                )
            log_lines = [
                (
                    f"{label} 완료: file={insert_path.name}, template_styles={style_count}, "
                    f"KeepStyle=1, KeepSection=0, page_setup={page_ok}({page_action}), {fill_summary}"
                ),
                f"보고양식 쪽설정 목표: {page_target}",
                f"보고양식 쪽설정 결과: {page_detail}",
            ]
            self.log(log_lines)
        except Exception as exc:
            self.log(f"{label} 실패: {type(exc).__name__}: {exc}")
            messagebox.showerror(f"{label} 실패", str(exc))

    def title_number_box_marker(self, entry: StyleEntry | None = None) -> str | None:
        if entry is None:
            return None
        settings = self.title_number_box_settings(entry)
        return build_title_number_box_marker(title_number_box_series_key(entry, settings))

    def current_title_number_box_numbers(self, entry: StyleEntry | None = None) -> list[int]:
        if self.hwp is None:
            return []
        try:
            hwpml = safe_str(self.hwp.GetTextFile("HWPML2X", ""))
            return title_number_box_numbers_from_hwpml(hwpml, self.title_number_box_marker(entry))
        except Exception as exc:
            self.debug(f"[title-number-box] 현재 문서 번호 읽기 실패: {type(exc).__name__}: {exc}")
            return []

    def title_number_box_numbers_before_paragraph(self, list_id: int, para: int, entry: StyleEntry | None = None) -> list[int] | None:
        if self.hwp is None:
            return None
        if para <= 0:
            return []
        original_pos = self.get_hwp_pos_by_set()
        try:
            self.clear_hwp_selection()
            if not self.select_hwp_text_block(list_id, 0, 0, para, 0):
                self.debug(f"[title-number-box] 앞쪽 선택 실패 list={list_id}, para={para}")
                return None
            hwpml = safe_str(self.hwp.GetTextFile("HWPML2X", "saveblock"))
            return title_number_box_numbers_from_hwpml(hwpml, self.title_number_box_marker(entry))
        except Exception as exc:
            self.debug(f"[title-number-box] 앞쪽 번호 읽기 실패 list={list_id}, para={para}: {type(exc).__name__}: {exc}")
            return None
        finally:
            self.clear_hwp_selection()
            if original_pos is not None:
                self.set_hwp_pos_by_set(original_pos)

    def next_title_number_box_number(
        self,
        *,
        list_id: int | None = None,
        para: int | None = None,
        entry: StyleEntry | None = None,
    ) -> tuple[int, list[int]]:
        numbers = self.current_title_number_box_numbers(entry) if entry is not None else self.current_title_number_box_numbers()
        if list_id is not None and para is not None:
            previous_numbers = (
                self.title_number_box_numbers_before_paragraph(list_id, para, entry)
                if entry is not None
                else self.title_number_box_numbers_before_paragraph(list_id, para)
            )
            if previous_numbers is not None:
                if previous_numbers:
                    return previous_numbers[-1] + 1, numbers
                return 1, numbers
        if not numbers:
            return 1, []
        return max(numbers) + 1, numbers

    def has_title_number_boxes_after_paragraph(self, list_id: int, para: int, entry: StyleEntry | None = None) -> bool:
        all_numbers = self.current_title_number_box_numbers(entry) if entry is not None else self.current_title_number_box_numbers()
        if not all_numbers:
            return False
        previous_numbers = (
            self.title_number_box_numbers_before_paragraph(list_id, para + 1, entry)
            if entry is not None
            else self.title_number_box_numbers_before_paragraph(list_id, para + 1)
        )
        if previous_numbers is None:
            return bool(all_numbers)
        return len(all_numbers) > len(previous_numbers)

    def title_number_box_settings(self, entry: StyleEntry | None = None) -> dict:
        table_settings = self.__dict__.get("table_settings", {})
        raw = table_settings.get("title_number_box", {}) if isinstance(table_settings, dict) else {}
        settings = merge_dict(DEFAULT_TITLE_NUMBER_BOX_SETTINGS, raw if isinstance(raw, dict) else {})
        template_override = normalize_config_path(entry.template_file) if entry is not None else ""
        if not template_override:
            style_set = self.active_style_set()
            role_config = style_set.role_configs.get("title_number_box", {}) if isinstance(style_set.role_configs, dict) else {}
            template_override = str(role_config.get("template_file", "")).strip() if isinstance(role_config, dict) else ""
        if template_override:
            settings["template_file"] = template_override
        template_file = str(settings.get("template_file") or DEFAULT_TITLE_NUMBER_BOX_SETTINGS["template_file"])
        settings["template_path"] = resolve_config_path(template_file)
        return settings

    def title_number_box_cell_specs(self) -> list[dict]:
        specs = self.title_number_box_settings().get("cells", [])
        return [spec for spec in specs if isinstance(spec, dict)]

    def title_number_box_outside_margin_values(self) -> dict[str, int]:
        raw = self.title_number_box_settings().get("outside_margin", {})
        if not isinstance(raw, dict):
            raw = {}
        return {
            "OutsideMarginLeft": int(raw.get("left", 0)),
            "OutsideMarginRight": int(raw.get("right", 0)),
            "OutsideMarginTop": int(raw.get("top", 0)),
            "OutsideMarginBottom": int(raw.get("bottom", 0)),
        }

    def apply_current_table_outside_margins(self, margin_values: dict[str, int]) -> bool:
        pset = self.get_table_property_set()
        set_ok = True
        for attr, value in margin_values.items():
            applied = self.set_com_attr(pset, attr, int(value))
            self.debug(f"[title-number-box] {attr}={value}, set={applied}")
            set_ok = applied and set_ok
        action_name, executed = self.execute_first_hwp_action(("TablePropertyDialog", "ShapeObjDialog"), pset.HSet)
        self.debug(f"[title-number-box] outside-margin action={action_name}, executed={executed}, set={set_ok}")
        return set_ok and executed

    def apply_current_cell_dimensions(
        self,
        width_hwp: int,
        height_hwp: int,
        margins: dict[str, int] | None = None,
    ) -> bool:
        self.run_hwp_command_best_effort("TableCellBlock")
        time.sleep(0.03)
        margins = margins or {}
        has_margin = any(int(margins.get(name, 0)) for name in ("left", "right", "top", "bottom"))
        results: list[tuple[str, bool]] = []
        try:
            action = self.hwp.CreateAction("TablePropertyDialog")
            pset = action.CreateSet()
            default_ok = bool(action.GetDefault(pset))
            size_ok = self.set_parameter_item(pset, "ShapeType", 3, "title-number-box-cell")
            size_ok = self.set_parameter_item(pset, "ShapeCellSize", 1, "title-number-box-cell") and size_ok
            shape_cell = self.get_parameter_child(pset, "ShapeTableCell", "title-number-box-cell")
            cell_ok = shape_cell is not None
            if shape_cell is not None:
                cell_ok = self.set_parameter_item(shape_cell, "Width", int(width_hwp), "title-number-box-cell") and cell_ok
                cell_ok = self.set_parameter_item(shape_cell, "Height", int(height_hwp), "title-number-box-cell") and cell_ok
                cell_ok = self.set_parameter_item(shape_cell, "HasMargin", 1 if has_margin else 0, "title-number-box-cell") and cell_ok
                for axis_name, item_name in (
                    ("left", "MarginLeft"),
                    ("right", "MarginRight"),
                    ("top", "MarginTop"),
                    ("bottom", "MarginBottom"),
                ):
                    cell_ok = self.set_parameter_item(
                        shape_cell,
                        item_name,
                        int(margins.get(axis_name, 0)),
                        "title-number-box-cell",
                    ) and cell_ok
            executed = action.Execute(pset)
            ok = bool(default_ok and size_ok and cell_ok and executed is not False and executed != 0)
            self.debug(
                f"[title-number-box-cell] width={width_hwp}, height={height_hwp}, margins={margins}, "
                f"default={default_ok}, size={size_ok}, cell={cell_ok}, execute={executed!r}, ok={ok}"
            )
            results.append(("CreateAction(TablePropertyDialog)", ok))
            if ok:
                self.clear_hwp_selection()
                return True
        except Exception as exc:
            self.debug(f"[title-number-box-cell] CreateAction 실패: {type(exc).__name__}: {exc}")

        try:
            pset = self.hwp.HParameterSet.HShapeObject
            default_ok = bool(self.hwp.HAction.GetDefault("TablePropertyDialog", pset.HSet))
            size_ok = self.set_parameter_item(pset, "ShapeType", 3, "title-number-box-cell-fallback")
            size_ok = self.set_parameter_item(pset, "ShapeCellSize", 1, "title-number-box-cell-fallback") and size_ok
            shape_cell = self.get_parameter_child(pset, "ShapeTableCell", "title-number-box-cell-fallback")
            cell_ok = shape_cell is not None
            if shape_cell is not None:
                cell_ok = self.set_parameter_item(shape_cell, "Width", int(width_hwp), "title-number-box-cell-fallback") and cell_ok
                cell_ok = self.set_parameter_item(shape_cell, "Height", int(height_hwp), "title-number-box-cell-fallback") and cell_ok
                cell_ok = self.set_parameter_item(shape_cell, "HasMargin", 1 if has_margin else 0, "title-number-box-cell-fallback") and cell_ok
                for axis_name, item_name in (
                    ("left", "MarginLeft"),
                    ("right", "MarginRight"),
                    ("top", "MarginTop"),
                    ("bottom", "MarginBottom"),
                ):
                    cell_ok = self.set_parameter_item(
                        shape_cell,
                        item_name,
                        int(margins.get(axis_name, 0)),
                        "title-number-box-cell-fallback",
                    ) and cell_ok
            action_name, executed = self.execute_first_hwp_action(("TablePropertyDialog", "ShapeObjDialog"), pset.HSet)
            ok = bool(default_ok and size_ok and cell_ok and executed)
            self.debug(
                f"[title-number-box-cell] fallback width={width_hwp}, height={height_hwp}, margins={margins}, "
                f"default={default_ok}, size={size_ok}, cell={cell_ok}, action={action_name}, ok={ok}"
            )
            results.append((action_name, ok))
            if ok:
                self.clear_hwp_selection()
                return True
        except Exception as exc:
            self.debug(f"[title-number-box-cell] HShapeObject fallback 실패: {type(exc).__name__}: {exc}")

        self.debug(f"[title-number-box-cell] 모든 경로 실패: {results}")
        self.clear_hwp_selection()
        return False

    def apply_title_number_box_cell(
        self,
        cell_spec: dict,
        text: str,
    ) -> bool:
        margins = cell_spec.get("margins", {})
        if not isinstance(margins, dict):
            margins = {}
        dimension_ok = self.apply_current_cell_dimensions(
            int(cell_spec.get("width", 0)),
            int(cell_spec.get("height", 0)),
            {name: int(margins.get(name, 0)) for name in ("left", "right", "top", "bottom")},
        )
        if not dimension_ok:
            return False

        fill_ok = True
        fill_rgb = cell_spec.get("fill_rgb")
        if isinstance(fill_rgb, (list, tuple)) and len(fill_rgb) == 3:
            fill_ok = self.set_cell_fill_color(
                self.palette_color_or_rgb("대제목번호박스배경", tuple(int(value) for value in fill_rgb))
            )[1]
        if not fill_ok:
            return False
        self.clear_hwp_selection()

        style_name = str(cell_spec.get("paragraph_style", "")).strip()
        style_ok = True
        if style_name:
            style_ok = self.apply_style_by_name(style_name)
        if not style_ok:
            return False

        text_color_ok = True
        text_color_rgb = cell_spec.get("text_color_rgb")
        if isinstance(text_color_rgb, (list, tuple)) and len(text_color_rgb) == 3:
            text_color_ok, _default_ok = self.set_text_color(
                self.palette_color_or_rgb("대제목번호박스글자", tuple(int(value) for value in text_color_rgb))
            )
        if not text_color_ok:
            return False

        return self.insert_hwp_text(text)

    def insert_title_number_box_file_at_cursor(self, title: str, number: int, entry: StyleEntry | None = None) -> bool:
        settings = self.title_number_box_settings(entry)
        template_path = settings["template_path"]
        marker = self.title_number_box_marker(entry) or str(settings.get("marker") or TITLE_NUMBER_BOX_MARKER)
        insert_path = create_filled_title_number_box_file(title, number, template_path, marker=marker)
        return self.insert_file_at_cursor(insert_path, keep_section=0)

    def is_title_number_box_entry(self, entry: StyleEntry) -> bool:
        settings = self.title_number_box_settings()
        if not settings.get("enabled", True):
            return False
        return style_entry_has_role(entry, str(settings.get("role") or "title_number_box"))

    def draw_title_number_box_at_cursor(self, title: str, number: int) -> bool:
        title = normalize_title_box_text(title)
        settings = self.title_number_box_settings()
        marker = str(settings.get("marker") or TITLE_NUMBER_BOX_MARKER)
        action_name, table_ok = self.create_table_at_cursor(1, 3)
        if not table_ok:
            self.debug(f"[title-number-box] 표 생성 실패: action={action_name}")
            return False
        texts = [str(number), title, marker]
        cell_specs = self.title_number_box_cell_specs()
        if len(cell_specs) < len(texts):
            self.debug(f"[title-number-box] 셀 설정 부족: specs={len(cell_specs)}, texts={len(texts)}")
            return False

        entered_first_cell = self.is_in_table_cell()
        if not entered_first_cell:
            entered_first_cell = self.select_current_table_object() and self.select_selected_table_first_cell()
            if entered_first_cell:
                self.clear_hwp_selection()

        steps_ok = bool(entered_first_cell)
        move_results: list[bool] = []
        apply_results: list[bool] = []
        for index, (cell_spec, text_value) in enumerate(zip(cell_specs, texts)):
            if index > 0:
                moved = self.move_table_cell("TableRightCell")
                move_results.append(moved)
                steps_ok = moved and steps_ok
                if not moved:
                    break
            applied = self.apply_title_number_box_cell(cell_spec, text_value)
            apply_results.append(applied)
            steps_ok = applied and steps_ok
            if not applied:
                break

        border_ok = False
        valign_ok = False
        outside_ok = False
        border_preset = str(settings.get("table_style_preset") or "").strip()
        if steps_ok and self.select_current_table_all_cells(1, 3):
            if border_preset:
                _border_action, border_ok = self.set_table_border_preset(border_preset)
            else:
                border_ok = True
            valign_ok = self.run_hwp_command_best_effort("TableVAlignCenter")
            self.clear_hwp_selection()
        if steps_ok and self.select_current_table_object():
            outside_ok = self.apply_current_table_outside_margins(self.title_number_box_outside_margin_values())
            self.clear_hwp_selection()

        ok = steps_ok and border_ok and valign_ok and outside_ok
        self.debug(
            "[title-number-box] draw "
            f"number={number}, table={table_ok}, entered={entered_first_cell}, "
            f"moves={move_results}, cells={apply_results}, border={border_ok}, "
            f"valign={valign_ok}, outside={outside_ok}, ok={ok}"
        )
        return ok

    def replace_paragraph_with_title_number_box(
        self,
        list_id: int,
        para: int,
        title: str,
        entry: StyleEntry | None = None,
        forced_number: int | None = None,
    ) -> tuple[bool, int, list[int]]:
        title = normalize_title_box_text(title)
        if not title:
            return False, 0, []
        next_number, existing_numbers = self.next_title_number_box_number(list_id=list_id, para=para, entry=entry)
        if forced_number is not None:
            next_number = max(1, int(forced_number))
        paragraph_text = self.read_current_paragraph_text(list_id, para)
        if paragraph_text is None:
            return False, next_number, existing_numbers
        delete_ok = self.delete_hwp_text_range(list_id, para, 0, len(paragraph_text)) if paragraph_text else True
        if not delete_ok:
            return False, next_number, existing_numbers
        if not self.set_hwp_pos((list_id, para, 0)):
            return False, next_number, existing_numbers
        insert_ok = self.insert_title_number_box_file_at_cursor(title, next_number, entry)
        return insert_ok, next_number, existing_numbers

    def insert_title_number_box_template(self) -> None:
        label = "대제목 번호박스 삽입"
        if not self.ensure_hwp():
            return
        template_path = TITLE_NUMBER_BOX_TEMPLATE_FILE
        try:
            if not template_path.exists():
                messagebox.showwarning(label, f"대제목 번호박스 파일을 찾을 수 없습니다.\n\n{template_path}")
                self.log(f"{label}: 파일 없음, path={template_path}")
                return

            copy_ok = self.run_hwp_command("Copy")
            time.sleep(0.08)
            selected_text = self.get_clipboard_text().strip() if copy_ok else ""
            delete_ok = None
            if selected_text:
                delete_ok = self.run_hwp_command("Delete")
                if not delete_ok:
                    messagebox.showwarning(label, "선택 텍스트를 삭제하지 못했습니다. 커서 위치만 선택하거나 다시 시도하세요.")
                    self.log(f"{label}: 선택 영역 삭제 실패, copy={copy_ok}")
                    return

            next_number, existing_numbers = self.next_title_number_box_number()
            insert_ok = self.insert_title_number_box_file_at_cursor(selected_text, next_number)
            if not insert_ok:
                insert_ok = self.draw_title_number_box_at_cursor(selected_text, next_number)
                if not insert_ok:
                    messagebox.showwarning(label, "한글이 번호박스 삽입을 실패로 반환했습니다. 삽입할 위치에 커서를 둔 상태인지 확인하세요.")
                    self.log(f"{label}: 삽입 실패, template={template_path.name}, selected_text={bool(selected_text)}, delete={delete_ok}")
                    return
                insert_mode = "table-action-fallback"
            else:
                insert_mode = "insertfile-template"

            self.refresh_current_doc_style_map(force=True)
            self.activate_hwp_window()
            self.log(
                f"{label} 완료: number={next_number}, existing={existing_numbers}, "
                f"mode={insert_mode}, selected_text={bool(selected_text)}, "
                f"delete={delete_ok}, text_preview={selected_text[:40]!r}"
            )
        except Exception as exc:
            self.log(f"{label} 실패: {type(exc).__name__}: {exc}")
            messagebox.showerror(f"{label} 실패", str(exc))

    def proof_render_options(self) -> tuple[int, bool]:
        try:
            dpi = int(self.proof_dpi_var.get().strip())
        except ValueError:
            raise ValueError("증빙 PDF 해상도는 정수 dpi로 입력하세요.")
        if dpi < 72 or dpi > 600:
            raise ValueError("증빙 PDF 해상도는 72~600 dpi 사이로 입력하세요.")
        return dpi, bool(self.proof_auto_rotate_var.get())

    def hwp_find_text(self, text: str, *, direction: int = 2) -> bool:
        pset = self.hwp.HParameterSet.HFindReplace
        self.hwp.HAction.GetDefault("RepeatFind", pset.HSet)
        values = {
            "FindString": text,
            "Direction": direction,
            "IgnoreMessage": 1,
            "FindType": 1,
            "MatchCase": 0,
            "AllWordForms": 0,
            "SeveralWords": 0,
            "UseWildCards": 0,
            "WholeWordOnly": 0,
            "ReplaceMode": 0,
        }
        for item, value in values.items():
            try:
                setattr(pset, item, value)
            except Exception:
                try:
                    pset.HSet.SetItem(item, value)
                except Exception:
                    pass
        return bool(self.hwp.HAction.Execute("RepeatFind", pset.HSet))

    def current_proof_full_cell_height(self) -> int:
        _default_ok, settings = self.read_current_page_settings()
        if settings is None:
            return PROOF_IMAGE_FULL_CELL_HEIGHT
        body_height = (
            settings.paper_height
            - settings.top_margin
            - settings.bottom_margin
            - settings.header_len
            - settings.footer_len
        )
        if body_height <= 0:
            return PROOF_IMAGE_FULL_CELL_HEIGHT
        self.debug(
            f"[proof-cell-height-current] body_height={body_height} "
            f"({hwpunit_to_mm(body_height):.2f}mm), settings={format_page_settings(settings)}"
        )
        return body_height

    def current_page_body_size(self) -> tuple[int, int, str]:
        default_ok, settings = self.read_current_page_settings()
        if settings is None:
            self.debug("[picture-fit-box] page settings unavailable; using proof defaults")
            return PROOF_IMAGE_CELL_WIDTH, PROOF_IMAGE_FULL_CELL_HEIGHT, "page-default"
        body_width = (
            settings.paper_width
            - settings.left_margin
            - settings.right_margin
            - settings.gutter_len
        )
        body_height = (
            settings.paper_height
            - settings.top_margin
            - settings.bottom_margin
            - settings.header_len
            - settings.footer_len
        )
        body_width = max(1, body_width)
        body_height = max(1, body_height)
        self.debug(
            "[picture-fit-box] "
            f"page GetDefault={default_ok}, body={body_width}x{body_height} "
            f"({hwpunit_to_mm(body_width):.2f}x{hwpunit_to_mm(body_height):.2f}mm), "
            f"settings={format_page_settings(settings)}"
        )
        return body_width, body_height, "page"

    def current_table_cell_size(self) -> tuple[int, int, str] | None:
        return self.read_current_table_cell_size("picture-fit-box")

    def read_current_table_cell_size(self, log_prefix: str) -> tuple[int, int, str] | None:
        try:
            cell_shape = self.hwp.CellShape
            candidates: list[tuple[str, object]] = [("CellShape", cell_shape)]
            cell = self.get_parameter_child(cell_shape, "Cell", f"{log_prefix}-cellshape")
            if cell is not None:
                candidates.insert(0, ("CellShape.Cell", cell))
            for source, candidate in candidates:
                width = self.get_parameter_item_value(candidate, "Width")
                height = self.get_parameter_item_value(candidate, "Height")
                self.debug(f"[{log_prefix}] {source} Width={width!r}, Height={height!r}")
                try:
                    width_int = int(width)
                    height_int = int(height)
                except Exception:
                    continue
                if width_int > 0 and height_int > 0:
                    return width_int, height_int, source
        except Exception as exc:
            self.debug(f"[{log_prefix}] CellShape read failed: {type(exc).__name__}: {exc}")

        try:
            pset = self.hwp.HParameterSet.HShapeObject
            default_ok = bool(self.hwp.HAction.GetDefault("TablePropertyDialog", pset.HSet))
            shape_cell = self.get_parameter_child(pset, "ShapeTableCell", f"{log_prefix}-table-dialog")
            if shape_cell is not None:
                width = self.get_parameter_item_value(shape_cell, "Width")
                height = self.get_parameter_item_value(shape_cell, "Height")
                self.debug(
                    f"[{log_prefix}] TablePropertyDialog GetDefault={default_ok}, "
                    f"ShapeTableCell Width={width!r}, Height={height!r}"
                )
                try:
                    width_int = int(width)
                    height_int = int(height)
                except Exception:
                    width_int = 0
                    height_int = 0
                if width_int > 0 and height_int > 0:
                    return width_int, height_int, "TablePropertyDialog.ShapeTableCell"
            else:
                self.debug(f"[{log_prefix}] TablePropertyDialog GetDefault={default_ok}, ShapeTableCell=None")
        except Exception as exc:
            self.debug(f"[{log_prefix}] TablePropertyDialog read failed: {type(exc).__name__}: {exc}")
        return None

    def picture_fit_box_from_ctrl_anchor(self, ctrl, log_prefix: str) -> tuple[int, int, str] | None:
        original_pos = self.get_hwp_pos_by_set()
        try:
            try:
                anchor_pos = ctrl.GetAnchorPos(0)
            except Exception as exc:
                self.debug(f"[{log_prefix}] GetAnchorPos(0) failed: {type(exc).__name__}: {exc}")
                return None
            if anchor_pos is None:
                self.debug(f"[{log_prefix}] GetAnchorPos(0)=None")
                return None
            set_ok = self.set_hwp_pos_by_set(anchor_pos)
            time.sleep(0.03)
            cell_size = self.read_current_table_cell_size(f"{log_prefix}-anchor-0")
            if cell_size is None:
                self.debug(f"[{log_prefix}] anchor[0] SetPosBySet={set_ok}, cell=None")
                return None
            width, height, source = cell_size
            self.debug(
                f"[{log_prefix}] anchor[0] SetPosBySet={set_ok}, "
                f"source={source}, size={width}x{height}, "
                f"mm={hwpunit_to_mm(width):.2f}x{hwpunit_to_mm(height):.2f}"
            )
            return width, height, f"anchor[0].{source}"
        finally:
            if original_pos is not None:
                restored = self.set_hwp_pos_by_set(original_pos)
                self.debug(f"[{log_prefix}] restore_pos={restored}")

    def current_picture_fit_box(self, ctrl=None) -> tuple[int, int, str]:
        if ctrl is not None:
            anchor_size = self.picture_fit_box_from_ctrl_anchor(ctrl, "picture-fit-box")
            if anchor_size is not None:
                return anchor_size
        cell_size = self.current_table_cell_size()
        if cell_size is not None:
            return cell_size
        return self.current_page_body_size()

    def set_current_cell_height(self, height_hwp: int, *, shape_cell_size: int = 1) -> bool:
        self.run_hwp_command_best_effort("TableCellBlock")
        time.sleep(0.03)
        results: list[tuple[str, bool]] = []
        try:
            pset = self.hwp.HParameterSet.HShapeObject
            default_ok = bool(self.hwp.HAction.GetDefault("TablePropertyDialog", pset.HSet))
            pset.HSet.SetItem("ShapeType", 3)
            pset.HSet.SetItem("ShapeCellSize", shape_cell_size)
            pset.ShapeTableCell.Height = int(height_hwp)
            executed = bool(self.hwp.HAction.Execute("TablePropertyDialog", pset.HSet))
            ok = bool(default_ok and executed)
            self.debug(
                f"[proof-cell-height-macro] height={height_hwp}, height_mm={hwpunit_to_mm(height_hwp):.2f}, "
                f"GetDefault={default_ok}, ShapeCellSize={shape_cell_size}, Execute={executed}, ok={ok}"
            )
            results.append(("macro-style TablePropertyDialog", ok))
            if ok:
                self.clear_hwp_selection()
                return True
        except Exception as exc:
            self.debug(f"[proof-cell-height-macro] 실패: {type(exc).__name__}: {exc}")

        try:
            action = self.hwp.CreateAction("TablePropertyDialog")
            pset = action.CreateSet()
            default_ok = bool(action.GetDefault(pset))
            size_ok = self.set_parameter_item(pset, "ShapeType", 3, "proof-cell-height-create")
            size_ok = self.set_parameter_item(pset, "ShapeCellSize", shape_cell_size, "proof-cell-height-create") and size_ok
            shape_cell = self.get_parameter_child(pset, "ShapeTableCell", "proof-cell-height-create")
            cell_ok = False
            if shape_cell is not None:
                try:
                    shape_cell.Height = int(height_hwp)
                    cell_ok = True
                except Exception as attr_exc:
                    self.debug(f"[proof-cell-height-create] ShapeTableCell.Height 직접 설정 실패: {type(attr_exc).__name__}: {attr_exc}")
                    cell_ok = self.set_parameter_item(shape_cell, "Height", int(height_hwp), "proof-cell-height-create")
            executed = bool(action.Execute(pset))
            ok = bool(default_ok and size_ok and cell_ok and executed)
            self.debug(
                f"[proof-cell-height-create] height={height_hwp}, GetDefault={default_ok}, "
                f"ShapeCellSize={shape_cell_size}, size={size_ok}, cell={cell_ok}, Execute={executed}, ok={ok}"
            )
            results.append(("CreateAction(TablePropertyDialog)", ok))
            if ok:
                self.clear_hwp_selection()
                return True
        except Exception as exc:
            self.debug(f"[proof-cell-height-create] 실패: {type(exc).__name__}: {exc}")

        try:
            cell_shape = self.hwp.CellShape
            cell = self.get_parameter_child(cell_shape, "Cell", "proof-cell-height-cellshape")
            cell_ok = False
            if cell is not None:
                cell_ok = self.set_parameter_item(cell, "Height", int(height_hwp), "proof-cell-height-cellshape")
                self.set_parameter_item(cell_shape, "Cell", cell, "proof-cell-height-cellshape")
            commit_ok = self.commit_cell_shape(cell_shape)
            ok = bool(cell_ok and commit_ok)
            self.debug(f"[proof-cell-height-cellshape] height={height_hwp}, cell={cell_ok}, commit={commit_ok}, ok={ok}")
            results.append(("CellShape.Cell.Height", ok))
            if ok:
                return True
        except Exception as exc:
            self.debug(f"[proof-cell-height-cellshape] 실패: {type(exc).__name__}: {exc}")

        try:
            pset = self.hwp.HParameterSet.HShapeObject
            default_ok = bool(self.hwp.HAction.GetDefault("TablePropertyDialog", pset.HSet))
            size_ok = self.set_parameter_item(pset, "ShapeType", 3, "proof-cell-height")
            size_ok = self.set_parameter_item(pset, "ShapeCellSize", shape_cell_size, "proof-cell-height") and size_ok
            shape_cell = self.get_parameter_child(pset, "ShapeTableCell", "proof-cell-height")
            cell_ok = False
            if shape_cell is not None:
                try:
                    shape_cell.Height = int(height_hwp)
                    cell_ok = True
                except Exception as attr_exc:
                    self.debug(f"[proof-cell-height] ShapeTableCell.Height 직접 설정 실패: {type(attr_exc).__name__}: {attr_exc}")
                    cell_ok = self.set_parameter_item(shape_cell, "Height", int(height_hwp), "proof-cell-height")
            action_name, executed = self.execute_first_hwp_action(("TablePropertyDialog", "ShapeObjDialog"), pset.HSet)
            ok = bool(default_ok and size_ok and cell_ok and executed)
            self.debug(
                f"[proof-cell-height] height={height_hwp}, GetDefault={default_ok}, "
                f"ShapeCellSize={shape_cell_size}, size={size_ok}, cell={cell_ok}, action={action_name}, Execute={executed}, ok={ok}"
            )
            results.append((action_name, ok))
            if ok:
                self.clear_hwp_selection()
                return True
        except Exception as exc:
            self.debug(f"[proof-cell-height] TablePropertyDialog fallback 실패: {type(exc).__name__}: {exc}")
        self.debug(f"[proof-cell-height] 모든 경로 실패: {results}")
        self.clear_hwp_selection()
        return False

    def prepare_proof_image_cell(self, height_hwp: int) -> None:
        self.run_hwp_command("ParagraphShapeAlignCenter")
        self.run_hwp_command("TableVAlignCenter")
        self.set_current_cell_height(height_hwp, shape_cell_size=1)

    def insert_proof_image_page(self, page: ProofImagePage, cell_height_hwp: int) -> bool:
        self.prepare_proof_image_cell(cell_height_hwp)
        width_mm, height_mm = fit_mm_to_cell(
            page.width_px,
            page.height_px,
            PROOF_IMAGE_CELL_WIDTH,
            cell_height_hwp,
        )
        return self.insert_picture_at_cursor(page.path, width_mm, height_mm)

    def append_proof_image_row(self) -> bool:
        self.clear_hwp_selection()
        time.sleep(0.05)
        before = self.get_current_cell_address()
        appended = self.run_hwp_command_best_effort("TableAppendRow")
        if not appended:
            appended = self.run_hwp_command_best_effort("TableInsertLowerRow")
        time.sleep(0.05)
        moved = self.run_hwp_command_best_effort("TableLowerCell")
        time.sleep(0.05)
        after = self.get_current_cell_address()
        moved_by_address = bool(before and after and after[1] > before[1])
        ok = bool(appended and moved_by_address)
        self.debug(f"[proof-row] before_col_row={before}, appended={appended}, moved={moved}, after_col_row={after}, ok={ok}")
        return ok

    def run_hwp_command_best_effort(self, command: str) -> bool:
        for runner in (
            lambda: self.hwp.HAction.Run(command),
            lambda: self.hwp.Run(command),
        ):
            try:
                result = runner()
                self.debug(f"[hwp-command] {command}: result={result!r}")
                return True
            except Exception as exc:
                self.debug(f"[hwp-command] {command} 실패: {type(exc).__name__}: {exc}")
        return False

    def move_after_current_proof_table(self, *, insert_page_break: bool = True) -> bool:
        self.clear_hwp_selection()
        time.sleep(0.05)
        closed = self.run_hwp_command_best_effort("CloseEx")
        if not insert_page_break:
            self.debug(f"[proof-exit] CloseEx={closed}, BreakPage=skipped")
            return bool(closed)
        time.sleep(0.08)
        page_break = self.run_hwp_command_best_effort("BreakPage")
        time.sleep(0.05)
        self.debug(f"[proof-exit] CloseEx={closed}, BreakPage={page_break}")
        return bool(closed and page_break)

    def insert_proof_pdf(self, pdf_path: Path, *, insert_page_break_after: bool = False) -> tuple[int, int]:
        title = normalize_proof_title_from_path(pdf_path)
        dpi, auto_rotate = self.proof_render_options()
        with tempfile.TemporaryDirectory(prefix="bufs-proof-") as temp_name:
            pages = render_pdf_pages_to_jpegs(
                pdf_path,
                Path(temp_name),
                dpi=dpi,
                quality=PROOF_JPEG_QUALITY,
                auto_rotate_landscape=auto_rotate,
            )
            if not pages:
                raise RuntimeError(f"PDF 페이지가 없습니다: {pdf_path.name}")

            template_path = create_filled_proof_template_file(title)
            insert_ok = self.insert_file_at_cursor(template_path, keep_section=0)
            if not insert_ok:
                raise RuntimeError(f"증빙 양식을 삽입하지 못했습니다: {template_path.name}")

            if not self.hwp_find_text(PROOF_IMAGE_MARKER, direction=1):
                if not self.hwp_find_text(PROOF_IMAGE_MARKER, direction=2):
                    raise RuntimeError(f"증빙자료 위치를 찾지 못했습니다: {pdf_path.name}")
            self.run_hwp_command("Delete")

            inserted = 0
            rotated = 0
            full_cell_height = self.current_proof_full_cell_height()
            for index, page in enumerate(pages):
                if index > 0 and not self.append_proof_image_row():
                    raise RuntimeError(f"증빙 이미지 행 추가에 실패했습니다: {pdf_path.name} p.{index + 1}")
                cell_height = PROOF_IMAGE_FIRST_CELL_HEIGHT if index == 0 else full_cell_height
                if not self.insert_proof_image_page(page, cell_height):
                    raise RuntimeError(f"증빙 이미지 삽입에 실패했습니다: {pdf_path.name} p.{index + 1}")
                inserted += 1
                rotated += 1 if page.rotated else 0
                self.clear_hwp_selection()
            self.move_after_current_proof_table(insert_page_break=insert_page_break_after)
        self.run_hwp_command("UpdateAll")
        self.log(f"증빙 삽입 완료: {pdf_path.name}, title={title}, pages={inserted}, rotated={rotated}, dpi={dpi}, jpg={PROOF_JPEG_QUALITY}")
        return inserted, rotated

    def insert_single_proof_pdf(self) -> None:
        label = "증빙 PDF 1건 삽입"
        if not self.ensure_hwp():
            return
        path = filedialog.askopenfilename(
            title=label,
            filetypes=(("PDF 파일", "*.pdf"), ("모든 파일", "*.*")),
        )
        if not path:
            return
        try:
            pages, rotated = self.insert_proof_pdf(Path(path))
            self.activate_hwp_window()
            messagebox.showinfo(label, f"삽입 완료\n\n페이지: {pages}개\n자동 회전: {rotated}개")
        except Exception as exc:
            self.log(f"{label} 실패: {type(exc).__name__}: {exc}")
            messagebox.showerror(f"{label} 실패", str(exc))

    def insert_proof_pdf_folder(self) -> None:
        label = "증빙 PDF 폴더 삽입"
        if not self.ensure_hwp():
            return
        folder = filedialog.askdirectory(title=label)
        if not folder:
            return
        try:
            pdfs = sorted(
                (path for path in Path(folder).iterdir() if path.is_file() and path.suffix.lower() == ".pdf"),
                key=proof_sort_key,
            )
            if not pdfs:
                messagebox.showwarning(label, "선택한 폴더에 PDF 파일이 없습니다.")
                return
            total_pages = 0
            total_rotated = 0
            failures: list[str] = []
            for pdf_index, pdf_path in enumerate(pdfs):
                try:
                    pages, rotated = self.insert_proof_pdf(
                        pdf_path,
                        insert_page_break_after=pdf_index < len(pdfs) - 1,
                    )
                    total_pages += pages
                    total_rotated += rotated
                except Exception as proof_exc:
                    failures.append(f"{pdf_path.name}: {proof_exc}")
                    self.log(f"{label} 파일 실패: {pdf_path.name}: {type(proof_exc).__name__}: {proof_exc}")
            self.activate_hwp_window()
            if failures:
                messagebox.showwarning(
                    label,
                    f"일부 파일을 건너뛰었습니다.\n\n성공 페이지: {total_pages}개\n자동 회전: {total_rotated}개\n실패: {len(failures)}건",
                )
            else:
                messagebox.showinfo(label, f"삽입 완료\n\n파일: {len(pdfs)}개\n페이지: {total_pages}개\n자동 회전: {total_rotated}개")
        except Exception as exc:
            self.log(f"{label} 실패: {type(exc).__name__}: {exc}")
            messagebox.showerror(f"{label} 실패", str(exc))

    def rotate_selected_picture(self, *, clockwise: bool) -> None:
        label = "그림 회전"
        if not self.ensure_hwp():
            return
        try:
            command = "ShapeObjRightAngleRotater" if clockwise else "ShapeObjRightAngleRotaterAnticlockwise"
            self.activate_hwp_window()
            time.sleep(0.12)
            unpin_ok = self.set_selected_shape_treat_as_char(False)
            time.sleep(0.05)
            rotate_result = self.run_hwp_command(command)
            rotate_ok = bool(rotate_result)
            rotate_detail = f"{command}:Run={rotate_result!r}"
            time.sleep(0.05)
            repin_ok = self.set_selected_shape_treat_as_char(True)
            time.sleep(0.05)
            fit_ok = False
            if rotate_ok:
                fit_ok = self.fit_selected_picture_to_proof_cell(show_messages=False)
            self.log(
                f"{label}: TreatAsChar0={unpin_ok}, command={command}, "
                f"rotate={rotate_ok}({rotate_detail}), "
                f"TreatAsChar1={repin_ok}, fit={fit_ok}"
            )
            if not rotate_ok:
                messagebox.showwarning(label, "회전 명령을 실행하지 못했습니다. 그림 개체를 선택한 뒤 다시 실행하세요.")
            elif not fit_ok:
                messagebox.showwarning(label, "회전은 실행했지만 셀 크기 자동 맞춤은 적용하지 못했습니다.")
        except Exception as exc:
            self.log(f"{label} 실패: {type(exc).__name__}: {exc}")
            messagebox.showerror(f"{label} 실패", str(exc))

    def set_selected_shape_treat_as_char(self, enabled: bool) -> bool:
        try:
            pset = self.hwp.HParameterSet.HShapeObject
            default_result = self.hwp.HAction.GetDefault("ShapeObjTreatAsChar", pset.HSet)
            pset.TreatAsChar = 0 if enabled else 1
            execute_result = self.hwp.HAction.Execute("ShapeObjTreatAsChar", pset.HSet)
            self.debug(
                f"[shape-treat-as-char] enabled={enabled}, "
                f"TreatAsChar={pset.TreatAsChar}, GetDefault={default_result}, Execute={execute_result}"
            )
            return True
        except Exception as exc:
            self.debug(f"[shape-treat-as-char] 실패: {type(exc).__name__}: {exc}")
            return False

    def fit_selected_picture_to_proof_cell(self, *, show_messages: bool = True) -> bool:
        label = "선택 그림 셀 맞춤"
        if not self.ensure_hwp():
            return False
        try:
            ctrl = self.hwp.CurSelectedCtrl
        except Exception:
            ctrl = None
        if ctrl is None:
            if show_messages:
                messagebox.showwarning(label, "셀 안의 그림 개체를 선택한 뒤 다시 실행하세요.")
            return False
        try:
            props = ctrl.Properties
            width = int(self.get_parameter_item_value(props, "Width") or 0)
            height = int(self.get_parameter_item_value(props, "Height") or 0)
            if width <= 0 or height <= 0:
                raise RuntimeError("선택 그림의 현재 크기를 읽지 못했습니다.")
            target_width, target_height, fit_source = self.current_picture_fit_box(ctrl)
            if fit_source.startswith("page"):
                new_width_mm, new_height_mm = fit_mm_to_box(width, height, target_width, target_height)
            else:
                new_width_mm, new_height_mm = fit_mm_to_cell(width, height, target_width, target_height)
            values = {
                "Width": int(round(new_width_mm * HWPUNIT_PER_MM)),
                "Height": int(round(new_height_mm * HWPUNIT_PER_MM)),
                "ProtectSize": 0,
            }
            if not fit_source.startswith("page"):
                values["WidthRelTo"] = 4
                values["HeightRelTo"] = 2
            ok = True
            for item, value in values.items():
                ok = self.set_parameter_item(props, item, value, "proof-picture-fit") and ok
            ctrl.Properties = props
            self.log(
                f"{label}: source={fit_source}, target={target_width}x{target_height}, "
                f"before={width}x{height}, after={values['Width']}x{values['Height']}, result={ok}"
            )
            if not ok and show_messages:
                messagebox.showwarning(label, "그림 크기 속성 일부를 적용하지 못했습니다.")
            return ok
        except Exception as exc:
            self.log(f"{label} 실패: {type(exc).__name__}: {exc}")
            if show_messages:
                messagebox.showerror(f"{label} 실패", str(exc))
            return False

    def insert_picture_at_cursor(self, path: Path, width_mm: float, height_mm: float) -> bool:
        file_name = str(path.resolve())

        method_attempts = (
            ("positional-mm", lambda: self.hwp.InsertPicture(file_name, True, 1, False, False, 0, width_mm, height_mm)),
            ("keyword-mm", lambda: self.hwp.InsertPicture(file_name, Embedded=True, sizeoption=1, Width=width_mm, Height=height_mm)),
        )
        for attempt_name, call in method_attempts:
            try:
                result = call()
                self.debug(f"[logo-insert] InsertPicture method {attempt_name} result={result}")
                if result:
                    return True
            except Exception as method_exc:
                self.debug(
                    f"[logo-insert] InsertPicture method {attempt_name} 실패: "
                    f"{type(method_exc).__name__}: {method_exc}"
                )

        self.debug("[logo-insert] InsertPicture method 모든 경로 실패")
        return False

    def create_cover_from_selected_lines(self) -> None:
        if not self.ensure_hwp():
            return
        original_text = ""
        try:
            copy_ok = self.run_hwp_command("Copy")
            time.sleep(0.08)
            original_text = self.get_clipboard_text() if copy_ok else ""
            has_input = bool(copy_ok and original_text.strip())
            if has_input:
                title, date_text, department = parse_cover_input(original_text)
            else:
                title = ""
                date_text = ""
                department = ""

            cover_path = create_filled_cover_file(
                title,
                date_text,
                department,
                self.cover_confidential.get(),
            )

            if has_input:
                delete_ok = self.run_hwp_command("Delete")
                if not delete_ok:
                    messagebox.showwarning("표지 만들기 실패", "선택한 3줄을 삭제하지 못했습니다. 선택 영역을 확인하세요.")
                    self.log("표지 만들기: 선택 영역 삭제 실패")
                    return

            insert_ok = self.insert_file_at_cursor(cover_path)
            if not insert_ok:
                if has_input:
                    self.set_clipboard_text(original_text)
                    restore_ok = self.run_hwp_command("Paste")
                else:
                    restore_ok = None
                messagebox.showwarning(
                    "표지 삽입 실패",
                    f"표지 파일은 만들었지만 현재 문서에 삽입하지 못했습니다.\n\n{cover_path}\n원문 복구: {restore_ok}",
                )
                self.log(f"표지 삽입 실패: file={cover_path}, restored={restore_ok}")
                return

            page_settings = read_hwpx_page_settings(COVER_FILE)
            if page_settings is None:
                page_action = "read_failed"
                page_ok = False
                page_detail = "none"
                page_target = "none"
            else:
                page_target = format_page_settings(page_settings)
                page_action, page_ok, page_detail = self.apply_page_settings(page_settings, apply_to=3)
            self.activate_hwp_window()
            self.log([
                (
                    "표지 만들기 완료: "
                    f"제목={title}, 날짜={date_text}, 부서={department}, "
                    f"대외비={self.cover_confidential.get()}, file={cover_path.name}, "
                    f"page_setup={page_ok}({page_action})"
                ),
                f"표지 쪽설정 목표: {page_target}",
                f"표지 쪽설정 결과: {page_detail}",
            ])
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
            result = self.hwp.HAction.Execute("InsertText", pset.HSet)
            return result is not False and result != 0
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
                result = runner()
                if result is not False and result != 0:
                    self.debug(f"[hwp-command] {command}: result={result!r}")
                    return True
            except Exception as exc:
                self.debug(f"[hwp-command] {command} 실패: {type(exc).__name__}: {exc}")
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

    def auto_confirm_dialog(
        self,
        title_part: str,
        timeout_sec: float = 2.0,
        preferred_texts: tuple[str, ...] = (),
    ) -> None:
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
                        preferred_clicked = False

                        def click_preferred(child, _extra):
                            nonlocal preferred_clicked
                            try:
                                text = win32gui.GetWindowText(child) or ""
                                if any(preferred in text for preferred in preferred_texts):
                                    win32gui.PostMessage(child, win32con.BM_CLICK, 0, 0)
                                    preferred_clicked = True
                            except Exception:
                                pass

                        if preferred_texts:
                            try:
                                win32gui.EnumChildWindows(hwnd, click_preferred, None)
                                time.sleep(0.05)
                            except Exception:
                                pass

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
                        self.debug(
                            f"[dialog-auto-confirm] {title_part} 확인 처리, "
                            f"preferred={preferred_texts}, preferred_clicked={preferred_clicked}"
                        )
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

    def init_hwp_scan(
        self,
        option: int,
        scan_range: int,
        spara: int = 0,
        spos: int = 0,
        epara: int = 0,
        epos: int = 0,
    ) -> bool:
        attempts = (
            (option, scan_range),
            (option, scan_range, spara, spos, epara, epos),
        )
        last_exc: Exception | None = None
        for args in attempts:
            try:
                return bool(self.hwp.InitScan(*args))
            except Exception as exc:
                last_exc = exc
        if last_exc is not None:
            self.debug(f"[hwp-scan] InitScan 실패 args={attempts[-1]}: {type(last_exc).__name__}: {last_exc}")
        return False

    def move_hwp_pos(self, move_id: int, para: int = 0, pos: int = 0) -> bool:
        attempts = (
            (move_id,),
            (move_id, para, pos),
        )
        last_exc: Exception | None = None
        for args in attempts:
            try:
                return bool(self.hwp.MovePos(*args))
            except Exception as exc:
                last_exc = exc
        if last_exc is not None:
            self.debug(f"[hwp-move] MovePos 실패 args={attempts[-1]}: {type(last_exc).__name__}: {last_exc}")
        return False

    def get_selected_cell_block_start_pos_by_scan(self):
        try:
            if not getattr(self.hwp, "CellShape"):
                return None
        except Exception:
            return None

        original_pos = self.get_hwp_pos_by_set()
        start_pos = None
        try:
            if not self.init_hwp_scan(1, 0xFF):
                return None
            try:
                self.hwp.GetText()
                self.move_hwp_pos(201)
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

    def current_paragraph_range(self) -> tuple[int, int, int] | None:
        try:
            pos = self.hwp.GetPos()
        except Exception as exc:
            self.debug(f"[paragraph-range] 현재 커서 위치 확인 실패: {type(exc).__name__}: {exc}")
            return None
        if not isinstance(pos, tuple) or len(pos) < 2:
            self.debug(f"[paragraph-range] 현재 커서 위치 invalid: {pos!r}")
            return None
        try:
            list_id = int(pos[0])
            para = int(pos[1])
        except Exception:
            return None
        return list_id, para, para

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
            init_ok = self.init_hwp_scan(0, 0x0010 | 0x0003, para, 0, para, 0)
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

    def read_int_parameter_item(self, target, item: str, default: int = 0) -> int:
        value = self.get_parameter_item_value(target, item)
        try:
            return int(value)
        except Exception:
            return default

    def current_paragraph_horizontal_margins(self, pset) -> int:
        left = self.read_int_parameter_item(pset, "LeftMargin")
        right = self.read_int_parameter_item(pset, "RightMargin")
        return max(0, left) + max(0, right)

    def current_cell_horizontal_margins(self) -> tuple[int, str]:
        candidates: list[tuple[str, object]] = []
        try:
            cell_shape = self.hwp.CellShape
            cell = self.get_parameter_child(cell_shape, "Cell", "paragraph-tab-cellshape")
            if cell is not None:
                candidates.append(("CellShape.Cell", cell))
            candidates.append(("CellShape", cell_shape))
        except Exception as exc:
            self.debug(f"[paragraph-tab] CellShape 여백 읽기 실패: {type(exc).__name__}: {exc}")

        try:
            pset = self.hwp.HParameterSet.HShapeObject
            default_ok = bool(self.hwp.HAction.GetDefault("TablePropertyDialog", pset.HSet))
            shape_cell = self.get_parameter_child(pset, "ShapeTableCell", "paragraph-tab-table-dialog")
            if shape_cell is not None:
                candidates.append((f"TablePropertyDialog.ShapeTableCell(default={default_ok})", shape_cell))
        except Exception as exc:
            self.debug(f"[paragraph-tab] TablePropertyDialog 여백 읽기 실패: {type(exc).__name__}: {exc}")

        for source, candidate in candidates:
            left = self.read_int_parameter_item(candidate, "MarginLeft")
            right = self.read_int_parameter_item(candidate, "MarginRight")
            if left or right:
                return max(0, left) + max(0, right), source
            default_left = self.read_int_parameter_item(candidate, "CellMarginLeft")
            default_right = self.read_int_parameter_item(candidate, "CellMarginRight")
            if default_left or default_right:
                return max(0, default_left) + max(0, default_right), source
        return 0, "none"

    def current_cell_inner_text_width(self) -> tuple[int, str] | None:
        if self.get_current_cell_address() is None:
            return None
        cell_size = self.read_current_table_cell_size("paragraph-tab")
        if cell_size is None:
            return None
        width, _height, width_source = cell_size
        margins, margin_source = self.current_cell_horizontal_margins()
        return max(0, width - margins), f"셀({width_source}, margin={margin_source})"

    def current_paragraph_tab_width(self, pset) -> tuple[int, str]:
        paragraph_margins = self.current_paragraph_horizontal_margins(pset)
        cell_width = self.current_cell_inner_text_width()
        if cell_width is not None:
            width, source = cell_width
            tab_width = width - paragraph_margins
            if tab_width > 0:
                return tab_width, f"{source}, para_margin={paragraph_margins}"
        page_width = self.current_page_text_width()
        return max(1, page_width - paragraph_margins), f"쪽, para_margin={paragraph_margins}"

    def set_tab_item_value(self, tab_items, index: int, value: int) -> bool:
        try:
            tab_items.SetItem(index, value)
            return True
        except Exception:
            pass
        try:
            tab_items.Item(index, value)
            return True
        except Exception as exc:
            self.debug(f"[paragraph-tab] TabItem[{index}]={value} 설정 실패: {type(exc).__name__}: {exc}")
            return False

    def set_paragraph_auto_right_tab(self) -> tuple[str, bool]:
        pset = self.hwp.HParameterSet.HParaShape
        default_ok = bool(self.hwp.HAction.GetDefault("ParagraphShape", pset.HSet))
        pset.TabDef.AutoTabRight = 1
        executed = bool(self.hwp.HAction.Execute("ParagraphShape", pset.HSet))
        self.debug(f"[paragraph-tab] auto-right default={default_ok}, result={executed}")
        return "ParagraphShape.AutoTabRight", executed

    def set_paragraph_dotted_right_tab(self) -> tuple[str, bool, int, str]:
        pset = self.hwp.HParameterSet.HParaShape
        default_ok = bool(self.hwp.HAction.GetDefault("ParagraphShape", pset.HSet))
        tab_position, source = self.current_paragraph_tab_width(pset)
        tab_def = pset.TabDef
        tab_def.CreateItemArray("TabItem", 3)
        tab_items = tab_def.TabItem
        item_ok = (
            self.set_tab_item_value(tab_items, 0, tab_position)
            and self.set_tab_item_value(tab_items, 1, 7)
            and self.set_tab_item_value(tab_items, 2, 1)
        )
        executed = item_ok and bool(self.hwp.HAction.Execute("ParagraphShape", pset.HSet))
        self.debug(
            f"[paragraph-tab] dotted-right default={default_ok}, source={source}, "
            f"tab={tab_position}, item_ok={item_ok}, result={executed}"
        )
        return "ParagraphShape.TabDef.TabItem", executed, tab_position, source

    def apply_auto_right_paragraph_tab(self) -> None:
        if not self.ensure_hwp():
            return
        try:
            self.activate_hwp_window()
            action, ok = self.set_paragraph_auto_right_tab()
            self.log(f"오른쪽 자동탭: action={action}, result={ok}")
            if not ok:
                messagebox.showwarning("오른쪽 자동탭", "현재 문단 또는 선택 영역에 문단 탭 설정을 적용하지 못했습니다.")
        except Exception as exc:
            self.log(f"오른쪽 자동탭 실패: {type(exc).__name__}: {exc}")
            messagebox.showerror("오른쪽 자동탭 실패", str(exc))

    def apply_dotted_right_paragraph_tab(self) -> None:
        if not self.ensure_hwp():
            return
        try:
            self.activate_hwp_window()
            action, ok, tab_position, source = self.set_paragraph_dotted_right_tab()
            self.log(f"점선 오른쪽탭: 기준={source}, 위치={tab_position}, action={action}, result={ok}")
            if not ok:
                messagebox.showwarning("점선 오른쪽탭", "현재 문단 또는 선택 영역에 문단 탭 설정을 적용하지 못했습니다.")
        except Exception as exc:
            self.log(f"점선 오른쪽탭 실패: {type(exc).__name__}: {exc}")
            messagebox.showerror("점선 오른쪽탭 실패", str(exc))

    def execute_put_new_para_number(self) -> tuple[bool, bool]:
        pset = self.hwp.HParameterSet.HParaShape
        default_ok = bool(self.hwp.HAction.GetDefault("PutNewParaNumber", pset.HSet))
        executed = bool(self.hwp.HAction.Execute("PutNewParaNumber", pset.HSet))
        return default_ok, executed

    def execute_put_para_number(self) -> tuple[bool, bool]:
        pset = self.hwp.HParameterSet.HParaShape
        default_ok = bool(self.hwp.HAction.GetDefault("PutParaNumber", pset.HSet))
        executed = bool(self.hwp.HAction.Execute("PutParaNumber", pset.HSet))
        return default_ok, executed

    def put_new_para_number_at_paragraph(self, list_id: int, para: int) -> bool:
        original_pos = self.get_hwp_pos_by_set()
        try:
            self.clear_hwp_selection()
            if not self.set_hwp_pos((list_id, para, 0)):
                return False
            default_ok, executed = self.execute_put_new_para_number()
            self.debug(f"[para-number] PutNewParaNumber para={para}, default={default_ok}, result={executed}")
            return executed
        except Exception as exc:
            self.debug(f"[para-number] PutNewParaNumber 실패 para={para}: {type(exc).__name__}: {exc}")
            return False
        finally:
            if original_pos is not None:
                self.set_hwp_pos_by_set(original_pos)

    def put_continued_para_number_at_paragraph(self, list_id: int, para: int, *, clear_heading: bool = False) -> bool:
        original_pos = self.get_hwp_pos_by_set()
        try:
            self.clear_hwp_selection()
            if not self.set_hwp_pos((list_id, para, 0)):
                return False
            if clear_heading and self.get_current_heading_string():
                self.clear_current_paragraph_heading()
            default_ok, executed = self.execute_put_para_number()
            self.debug(f"[para-number] PutParaNumber para={para}, default={default_ok}, result={executed}")
            return executed
        except Exception as exc:
            self.debug(f"[para-number] PutParaNumber 실패 para={para}: {type(exc).__name__}: {exc}")
            return False
        finally:
            if original_pos is not None:
                self.set_hwp_pos_by_set(original_pos)

    def put_new_paragraph_number_at_cursor(self) -> None:
        label = "문단 새 번호"
        if not self.ensure_hwp():
            return
        try:
            self.activate_hwp_window()
            default_ok, executed = self.execute_put_new_para_number()
            self.log(f"{label}: PutNewParaNumber, GetDefault={default_ok}, result={executed}")
            if not executed:
                messagebox.showwarning(label, "현재 문단에서 새 번호 시작 명령이 실패했습니다. 문단번호가 적용된 문단인지 확인하세요.")
        except Exception as exc:
            self.log(f"{label} 실패: {type(exc).__name__}: {exc}")
            messagebox.showerror(f"{label} 실패", str(exc))

    def put_continued_paragraph_number_at_cursor(self) -> None:
        label = "문단 번호 이어서"
        if not self.ensure_hwp():
            return
        try:
            self.activate_hwp_window()
            default_ok, executed = self.execute_put_para_number()
            self.log(f"{label}: PutParaNumber, GetDefault={default_ok}, result={executed}")
            if not executed:
                messagebox.showwarning(label, "현재 문단에서 번호 이어서 명령이 실패했습니다. 문단번호가 적용될 문단인지 확인하세요.")
        except Exception as exc:
            self.log(f"{label} 실패: {type(exc).__name__}: {exc}")
            messagebox.showerror(f"{label} 실패", str(exc))

    def delete_hwp_text_range(self, list_id: int, para: int, start_pos: int, end_pos: int) -> bool:
        if end_pos <= start_pos:
            return False
        actual_start, actual_end = self.actual_hwp_text_range(list_id, para, start_pos, end_pos)
        try:
            if not self.set_hwp_pos((list_id, para, actual_start)):
                return False
            if hasattr(self.hwp, "SelectText") and self.hwp.SelectText(para, actual_start, para, actual_end):
                return self.run_hwp_command("Delete")
        except Exception as exc:
            self.debug(f"[paragraph-delete] SelectText 실패 para={para}: {type(exc).__name__}: {exc}")
        if not self.set_hwp_pos((list_id, para, actual_start)):
            return False
        for _ in range(end_pos - start_pos):
            if not self.run_hwp_command("MoveSelRight"):
                return False
        return self.run_hwp_command("Delete")

    def select_hwp_text_range(self, list_id: int, para: int, start_pos: int, end_pos: int) -> bool:
        if end_pos <= start_pos:
            actual_start, _actual_end = self.actual_hwp_text_range(list_id, para, start_pos, end_pos)
            return self.set_hwp_pos((list_id, para, actual_start))
        self.clear_hwp_selection()
        actual_start, actual_end = self.actual_hwp_text_range(list_id, para, start_pos, end_pos)
        try:
            if not self.set_hwp_pos((list_id, para, actual_start)):
                return False
            if hasattr(self.hwp, "SelectText") and self.hwp.SelectText(para, actual_start, para, actual_end):
                matches = self.selected_hwp_text_range_matches(list_id, para, actual_start, actual_end)
                if matches is not False:
                    return True
                self.debug(
                    f"[paragraph-select] SelectText 범위 불일치, fallback 사용 "
                    f"para={para}, expected={actual_start}-{actual_end}, visible={start_pos}-{end_pos}"
                )
                self.clear_hwp_selection()
        except Exception as exc:
            self.debug(f"[paragraph-select] SelectText 실패 para={para}: {type(exc).__name__}: {exc}")
        if not self.set_hwp_pos((list_id, para, actual_start)):
            return False
        for _ in range(end_pos - start_pos):
            if not self.run_hwp_command("MoveSelRight"):
                return False
        return self.selected_hwp_text_range_matches(list_id, para, actual_start, actual_end) is not False

    def hwp_paragraph_visible_text_offset(self, list_id: int, para: int) -> int:
        if not self.set_hwp_pos((list_id, para, 0)):
            return 0
        try:
            pos = self.hwp.GetPos()
        except Exception as exc:
            self.debug(f"[paragraph-pos] GetPos 실패 para={para}: {type(exc).__name__}: {exc}")
            return 0
        if not isinstance(pos, tuple) or len(pos) < 3:
            return 0
        try:
            actual_list_id = int(pos[0])
            actual_para = int(pos[1])
            actual_pos = int(pos[2])
        except Exception:
            return 0
        if actual_list_id == list_id and actual_para == para and actual_pos > 0:
            return actual_pos
        return 0

    def actual_hwp_text_range(self, list_id: int, para: int, start_pos: int, end_pos: int) -> tuple[int, int]:
        offset = self.hwp_paragraph_visible_text_offset(list_id, para)
        return offset + start_pos, offset + end_pos

    def selected_hwp_text_range_matches(self, list_id: int, para: int, start_pos: int, end_pos: int) -> bool | None:
        try:
            result = self.hwp.GetSelectedPos()
        except Exception as exc:
            self.debug(f"[paragraph-select] GetSelectedPos 확인 생략: {type(exc).__name__}: {exc}")
            return None
        if not isinstance(result, tuple) or len(result) < 7 or not result[0]:
            self.debug(f"[paragraph-select] GetSelectedPos 확인 생략: invalid={result}")
            return None
        start = (int(result[1]), int(result[2]), int(result[3]))
        end = (int(result[4]), int(result[5]), int(result[6]))
        start, end = sorted((start, end))
        expected_start = (list_id, para, start_pos)
        expected_end = (list_id, para, end_pos)
        matches = start == expected_start and end == expected_end
        if not matches:
            self.debug(f"[paragraph-select] 선택 범위 확인 실패: actual={start}-{end}, expected={expected_start}-{expected_end}")
        return matches

    def apply_style_to_paragraph_text(self, list_id: int, para: int, record: StyleRecord) -> bool:
        paragraph_text = self.read_current_paragraph_text(list_id, para)
        if paragraph_text is None:
            return False
        if not self.select_hwp_text_range(list_id, para, 0, len(paragraph_text)):
            return False
        return self.execute_style_record(record)

    def is_hwp_paragraph_in_table(self, list_id: int, para: int) -> bool:
        original_pos = self.get_hwp_pos_by_set()
        try:
            if not self.set_hwp_pos((list_id, para, 0)):
                return False
            return self.get_current_cell_address() is not None
        finally:
            if original_pos is not None:
                self.set_hwp_pos_by_set(original_pos)

    def inline_rule_applies_to_context(self, rule: InlineRule, *, in_table: bool) -> bool:
        target = normalize_inline_rule_target(rule.target)
        if target == "all":
            return True
        if target == "table":
            return in_table
        return not in_table

    def character_style_record_for_role(self, style_set: StyleSet, role: str) -> StyleRecord | None:
        for entry in style_set.character_styles:
            if role in entry.roles:
                return self.find_current_doc_style_record(entry.name)
        return None

    def apply_inline_rules_to_paragraph(
        self,
        list_id: int,
        para: int,
        style_set: StyleSet,
        *,
        missing_roles: set[str],
        force_in_table: bool | None = None,
    ) -> int:
        enabled_rules = [rule for rule in style_set.inline_rules if rule.enabled]
        if not enabled_rules:
            return 0
        in_table = self.is_hwp_paragraph_in_table(list_id, para) if force_in_table is None else force_in_table
        applied = 0
        for rule in enabled_rules:
            if not self.inline_rule_applies_to_context(rule, in_table=in_table):
                continue
            record = self.character_style_record_for_role(style_set, rule.style_role)
            if record is None:
                missing_roles.add(rule.style_role)
                self.debug(f"[inline-style] 현재 문서 글자 스타일 역할 없음: role={rule.style_role}")
                continue
            paragraph_text = self.read_current_paragraph_text(list_id, para)
            if paragraph_text is None:
                continue
            ranges = find_inline_rule_ranges_for_style_set(paragraph_text, rule, style_set)
            for start_pos, end_pos, marker_range in reversed(ranges):
                apply_start = start_pos
                apply_end = end_pos
                if marker_range is not None and rule.rule_type == "markdown_bold":
                    marker_start, marker_end = marker_range
                    close_start = marker_end - 2
                    if not self.delete_hwp_text_range(list_id, para, close_start, marker_end):
                        continue
                    if not self.delete_hwp_text_range(list_id, para, marker_start, marker_start + 2):
                        continue
                    apply_start = max(marker_start, start_pos - 2)
                    apply_end = max(apply_start, end_pos - 2)
                if self.select_hwp_text_range(list_id, para, apply_start, apply_end) and self.execute_style_record(record):
                    applied += 1
                self.clear_hwp_selection()
        return applied

    def process_configured_style_paragraph(
        self,
        label: str,
        list_id: int,
        para: int,
        active_set: StyleSet,
        missing_styles: set[str],
        missing_roles: set[str],
        force_in_table: bool | None = None,
        numbering_state: NumberingRunState | None = None,
        title_box_followup_styles: set[str] | None = None,
        warn_on_existing_title_boxes: bool = False,
    ) -> tuple[bool, bool, bool, int]:
        paragraph_text = self.read_current_paragraph_text(list_id, para)
        if paragraph_text is None:
            self.log(f"{label}: 문단 텍스트 읽기 실패 para={para}")
            return False, False, False, 0

        in_table = self.is_hwp_paragraph_in_table(list_id, para) if force_in_table is None else force_in_table
        matched = False
        changed = False
        rule = find_outline_style_rule(paragraph_text, active_set, in_table=in_table)
        if rule is not None:
            entry, prefix_length = rule
            matched = True
            continuation_level = 0
            if numbering_state is not None and entry.numbering_group and entry.numbering_level > 0:
                continuation_level = numbering_state.continued_levels.get(entry.numbering_group, 0)
                if continuation_level and entry.numbering_level < continuation_level:
                    numbering_state.continued_levels.pop(entry.numbering_group, None)
                    continuation_level = 0
            if self.is_title_number_box_entry(entry):
                if in_table:
                    self.debug(f"[title-number-box-bulk] 표 내부 문단은 번호박스 치환 건너뜀 para={para}")
                else:
                    should_restart_numbering = consume_numbering_entry(numbering_state, entry)
                    title = paragraph_text[prefix_length:]
                    insert_ok, next_number, existing_numbers = self.replace_paragraph_with_title_number_box(
                        list_id,
                        para,
                        title,
                        entry,
                        1 if should_restart_numbering else None,
                    )
                    changed = insert_ok
                    if insert_ok and numbering_state is not None and should_restart_numbering and entry.numbering_group:
                        numbering_state.restarted += 1
                        numbering_state.continued_levels[entry.numbering_group] = entry.numbering_level
                    if (
                        insert_ok
                        and warn_on_existing_title_boxes
                        and title_box_followup_styles is not None
                        and self.has_title_number_boxes_after_paragraph(list_id, para, entry)
                    ):
                        title_box_followup_styles.add(entry.name)
                    self.debug(
                        f"[title-number-box-bulk] para={para}, number={next_number}, "
                        f"existing={existing_numbers}, title={normalize_title_box_text(title)!r}, "
                        f"restart={should_restart_numbering}, result={insert_ok}"
                    )
                    return True, matched, changed, 0

            should_restart_numbering = consume_numbering_entry(numbering_state, entry)
            current_record = self.find_current_doc_style_record(entry.name)
            if current_record is None:
                missing_styles.add(entry.name)
                self.debug(f"[outline-style] 현재 문서 스타일 없음: {entry.name}")
            else:
                delete_ok = True
                if prefix_length > 0:
                    delete_ok = self.delete_hwp_text_range(list_id, para, 0, prefix_length)
                style_ok = self.apply_style_to_paragraph_text(list_id, para, current_record)
                if delete_ok and style_ok:
                    changed = True
                    if should_restart_numbering and numbering_state is not None:
                        if self.put_new_para_number_at_paragraph(list_id, para):
                            numbering_state.restarted += 1
                            numbering_state.continued_levels[entry.numbering_group] = entry.numbering_level
                        else:
                            numbering_state.restart_failed += 1
                    elif (
                        numbering_state is not None
                        and continuation_level
                        and continuation_level == entry.numbering_level
                    ):
                        if self.put_continued_para_number_at_paragraph(list_id, para, clear_heading=True):
                            numbering_state.continued += 1
                        else:
                            numbering_state.continue_failed += 1
                self.debug(
                    f"[outline-style] para={para}, marker={paragraph_text[:prefix_length]!r}, "
                    f"style={entry.name}, delete={delete_ok}, style_ok={style_ok}, "
                    f"numbering_restart={should_restart_numbering}"
                )

        inline_applied = self.apply_inline_rules_to_paragraph(
            list_id,
            para,
            active_set,
            missing_roles=missing_roles,
            force_in_table=in_table,
        )
        self.clear_hwp_selection()
        return True, matched, changed, inline_applied

    def selected_current_cell_paragraph_range(self) -> tuple[int, int, int] | None:
        if not self.select_current_table_cell_for_replace():
            self.debug("[bulk-cell-style] 현재 셀 선택 실패")
            return None
        try:
            raw_mode, base_mode = self.get_selection_mode_base()
            try:
                raw_selected_pos = self.hwp.GetSelectedPos()
            except Exception as exc:
                raw_selected_pos = f"{type(exc).__name__}: {exc}"
            try:
                key_indicator = self.hwp.KeyIndicator()
            except Exception as exc:
                key_indicator = f"{type(exc).__name__}: {exc}"
            paragraph_range = self.selected_paragraph_range()
            self.debug(
                f"[bulk-cell-style] 셀 선택 진단: selection_mode={raw_mode}/{base_mode}, "
                f"key_indicator={key_indicator!r}, selected_pos={raw_selected_pos!r}, "
                f"paragraph_range={paragraph_range!r}"
            )
            return paragraph_range
        finally:
            self.clear_hwp_selection()

    def scan_current_cell_paragraph_positions(self, limit: int = 200) -> list[tuple[int, int]]:
        original_pos = self.get_hwp_pos_by_set()
        positions: list[tuple[int, int]] = []
        seen: set[tuple[int, int]] = set()
        try:
            init_ok = self.init_hwp_scan(0, 0x0050 | 0x0005)
            if not init_ok:
                self.debug("[bulk-cell-style] 셀 list scan InitScan 실패")
                return []
            try:
                for _ in range(limit):
                    status, text = self.parse_hwp_get_text_result(self.hwp.GetText())
                    try:
                        moved = self.move_hwp_pos(201)
                        raw_pos = self.hwp.GetPos()
                    except Exception as exc:
                        moved = False
                        raw_pos = None
                        self.debug(f"[bulk-cell-style] scan MovePos/GetPos 실패: {type(exc).__name__}: {exc}")
                    if (
                        text.strip()
                        and moved
                        and isinstance(raw_pos, tuple)
                        and len(raw_pos) >= 3
                    ):
                        item = (int(raw_pos[0]), int(raw_pos[1]))
                        if item not in seen:
                            positions.append(item)
                            seen.add(item)
                    if status in (0, 1, 101, 102):
                        break
            finally:
                try:
                    self.hwp.ReleaseScan()
                except Exception:
                    pass
        except Exception as exc:
            self.debug(f"[bulk-cell-style] 셀 list scan 실패: {type(exc).__name__}: {exc}")
            return []
        finally:
            if original_pos is not None:
                self.set_hwp_pos_by_set(original_pos)
        self.debug(f"[bulk-cell-style] scan paragraph positions={positions}")
        return positions

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
            option = self.hwp.HParameterSet.HSelectionOpt
            paste_default = self.hwp.HAction.GetDefault("Paste", option.HSet)
            for option_value in (5, 0):
                try:
                    option.Option = option_value
                except Exception:
                    pass
                self.auto_confirm_dialog("HTML 문서 붙이기", preferred_texts=("원본", "원본 형식"))
                paste_ok = bool(self.hwp.HAction.Execute("Paste", option.HSet))
                self.debug(
                    f"[markdown-table] Paste Option={option_value}, "
                    f"default={paste_default}, result={paste_ok}"
                )
                if paste_ok:
                    return True
        except Exception as exc:
            self.debug(f"[markdown-table] Paste option 우선 경로 실패: {type(exc).__name__}: {exc}")

        try:
            html_set = self.hwp.HParameterSet.HPasteHtml
            html_default = self.hwp.HAction.GetDefault("PasteHtmlDialog", html_set.HSet)
            html_set.Format = 0
            self.auto_confirm_dialog("HTML 문서 붙이기", preferred_texts=("원본", "원본 형식"))
            html_ok = bool(self.hwp.HAction.Execute("PasteHtmlDialog", html_set.HSet))
            self.debug(f"[markdown-table] PasteHtmlDialog Format=0, default={html_default}, result={html_ok}")
            if html_ok:
                return True
        except Exception as exc:
            self.debug(f"[markdown-table] PasteHtmlDialog 실패: {type(exc).__name__}: {exc}")
            html_ok = False

        self.auto_confirm_dialog("HTML 문서 붙이기", preferred_texts=("원본", "원본 형식"))
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
        preflight=None,
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
            if preflight is not None:
                try:
                    preflight(text)
                except SuspiciousDecimalNumberError as exc:
                    examples = "\n".join(f"- {value}" for value in exc.values[:8])
                    suffix = ""
                    if len(exc.values) > 8:
                        suffix = f"\n- ... 외 {len(exc.values) - 8}개"
                    messagebox.showwarning(
                        label,
                        "콤마 위치가 이상하거나 소숫점을 콤마로 쓴 것으로 의심되는 숫자가 있어 변환을 중단했습니다.\n\n"
                        f"{examples}{suffix}\n\n"
                        "해당 값을 먼저 확인한 뒤 다시 실행하세요.",
                    )
                    self.log(f"{label}: 의심 숫자 감지로 중단, values={exc.values}")
                    return
                except UnsafeUnitConversionNumberError as exc:
                    examples = "\n".join(f"- {value}" for value in exc.values[:8])
                    suffix = ""
                    if len(exc.values) > 8:
                        suffix = f"\n- ... 외 {len(exc.values) - 8}개"
                    messagebox.showwarning(
                        label,
                        "단위 변환 대상이 아닌 단위이거나 버튼의 출발 단위와 다른 값이 있어 변환을 중단했습니다.\n\n"
                        f"{examples}{suffix}\n\n"
                        "셀 안 단위가 맞는지 확인하거나, 해당 단위에 맞는 변환 버튼을 사용하세요.",
                    )
                    self.log(f"{label}: 단위 변환 위험 값 감지로 중단, values={exc.values}")
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
        self.draw_selected_markdown_table(label="Markdown 표 → 한글 표")

    def read_selected_markdown_table(self, label: str, *, warn_if_missing: bool = True) -> tuple[str, list[list[str]]] | None:
        selected_positions = self.get_selected_text_positions()
        copy_ok = self.run_hwp_command("Copy")
        time.sleep(0.08)
        text = self.get_clipboard_text()
        if selected_positions is None or not copy_ok or not text:
            if warn_if_missing:
                messagebox.showwarning(label, "변환할 markdown 표 텍스트를 먼저 선택하세요.")
                self.log(f"{label}: 선택 텍스트 없음")
            return None
        rows = parse_markdown_table(text)
        if rows is None:
            if warn_if_missing:
                messagebox.showwarning(
                    label,
                    "markdown 표를 찾지 못했습니다.\n\n"
                    "예: | 항목 | 값 | 형식과 그 아래 |---|---| 구분선까지 함께 선택하세요.",
                )
                self.log(f"{label}: markdown 표 인식 실패")
            return None
        return text, rows

    def draw_selected_markdown_table(self, *, label: str = "Markdown 표 → 표 그리기", warn_if_missing: bool = True) -> str:
        if not self.ensure_hwp():
            return "failed"
        try:
            payload = self.read_selected_markdown_table(label, warn_if_missing=warn_if_missing)
            if payload is None:
                return "not_markdown"
            text, rows = payload
            delete_ok = self.run_hwp_command("Delete")
            if not delete_ok:
                if warn_if_missing:
                    messagebox.showwarning(label, "선택한 markdown 표 텍스트를 삭제하지 못했습니다. 선택 영역을 다시 확인하세요.")
                self.log(f"{label}: 선택 영역 삭제 실패")
                return "failed"
            action_name, table_ok = self.create_table_at_cursor(len(rows), len(rows[0]))
            if not table_ok:
                self.set_clipboard_text(text)
                restore_ok = self.run_hwp_command("Paste")
                if warn_if_missing:
                    messagebox.showwarning(
                        label,
                        "표 만들기 명령이 실패했습니다.\n\n"
                        f"원문 복구: {restore_ok}",
                    )
                self.log(f"{label}: 표 생성 실패, action={action_name}, restored={restore_ok}")
                return "failed"
            fill_ok = self.fill_current_table_with_rows(rows)
            if fill_ok:
                format_results = self.apply_markdown_table_post_formatting(label, len(rows), len(rows[0]))
                self.activate_hwp_window()
            else:
                format_results = []
            self.log(
                f"{label}: rows={len(rows)}, cols={len(rows[0])}, "
                f"action={action_name}, fill={fill_ok}, post_format={format_results}"
            )
            if not fill_ok:
                if warn_if_missing:
                    messagebox.showwarning(
                        label,
                        "표 틀은 만들었지만 셀 채우기에 실패했습니다.\n\n"
                        "한글에서 Ctrl+Z로 되돌린 뒤 다시 시도하세요.",
                    )
                return "failed"
            elif format_results and not all(ok for _name, ok in format_results):
                failed = ", ".join(name for name, ok in format_results if not ok)
                if warn_if_missing:
                    messagebox.showwarning(
                        label,
                        "표는 만들었지만 일부 후처리 적용에 실패했습니다.\n\n"
                        f"실패 항목: {failed}",
                    )
            return "converted"
        except Exception as exc:
            self.log(f"{label} 실패: {type(exc).__name__}: {exc}")
            if warn_if_missing:
                messagebox.showerror(f"{label} 실패", str(exc))
            return "failed"

    def read_markdown_table_block(
        self,
        list_id: int,
        start_para: int,
        last_para: int,
    ) -> tuple[int, str, list[list[str]]] | None:
        paragraphs: list[str] = []
        para_numbers: list[int] = []
        for para in range(start_para, last_para + 1):
            paragraph_text = self.read_current_paragraph_text(list_id, para)
            if paragraph_text is None:
                break
            stripped = paragraph_text.strip()
            if not paragraphs and (not stripped or not stripped.startswith("|")):
                return None
            paragraphs.append(paragraph_text)
            para_numbers.append(para)
            if not stripped or not stripped.startswith("|"):
                break
        block = find_markdown_table_block(paragraphs)
        if block is None:
            return None
        _start_index, end_index, rows = block
        block_text = "\n".join(paragraphs[: end_index + 1])
        return para_numbers[end_index], block_text, rows

    def select_hwp_text_block(
        self,
        list_id: int,
        start_para: int,
        start_pos: int,
        end_para: int,
        end_pos: int,
    ) -> bool:
        if end_para == start_para:
            return self.select_hwp_text_range(list_id, start_para, start_pos, end_pos)
        self.clear_hwp_selection()
        actual_start, _ = self.actual_hwp_text_range(list_id, start_para, start_pos, start_pos)
        _actual_end_start, actual_end = self.actual_hwp_text_range(list_id, end_para, end_pos, end_pos)
        try:
            if not self.set_hwp_pos((list_id, start_para, actual_start)):
                return False
            if hasattr(self.hwp, "SelectText") and self.hwp.SelectText(start_para, actual_start, end_para, actual_end):
                return True
        except Exception as exc:
            self.debug(
                f"[markdown-table-block] SelectText 실패 start={start_para}:{actual_start}, "
                f"end={end_para}:{actual_end}, error={type(exc).__name__}: {exc}"
            )
        return False

    def replace_markdown_table_block(
        self,
        label: str,
        list_id: int,
        start_para: int,
        end_para: int,
        source_text: str,
        rows: list[list[str]],
    ) -> str:
        end_text = self.read_current_paragraph_text(list_id, end_para)
        if end_text is None:
            self.log(f"{label}: markdown 표 끝 문단 읽기 실패 para={end_para}")
            return "failed"
        if not self.select_hwp_text_block(list_id, start_para, 0, end_para, len(end_text)):
            self.log(f"{label}: markdown 표 블록 선택 실패 paras={start_para}-{end_para}")
            return "failed"
        if not self.run_hwp_command("Delete"):
            self.log(f"{label}: markdown 표 블록 삭제 실패 paras={start_para}-{end_para}")
            return "failed"
        action_name, table_ok = self.create_table_at_cursor(len(rows), len(rows[0]))
        if not table_ok:
            self.set_clipboard_text(source_text)
            restore_ok = self.run_hwp_command("Paste")
            self.log(
                f"{label}: markdown 표 생성 실패 paras={start_para}-{end_para}, "
                f"action={action_name}, restored={restore_ok}"
            )
            return "failed"
        fill_ok = self.fill_current_table_with_rows(rows)
        format_results = self.apply_markdown_table_post_formatting(label, len(rows), len(rows[0])) if fill_ok else []
        self.activate_hwp_window()
        self.log(
            f"{label}: markdown 표 블록 변환 paras={start_para}-{end_para}, "
            f"rows={len(rows)}, cols={len(rows[0])}, action={action_name}, "
            f"fill={fill_ok}, post_format={format_results}"
        )
        if not fill_ok:
            return "failed"
        if format_results and not all(ok for _name, ok in format_results):
            return "partial"
        return "converted"

    def apply_markdown_table_post_formatting(
        self,
        label: str,
        row_count: int | None = None,
        col_count: int | None = None,
    ) -> list[tuple[str, bool]]:
        results: list[tuple[str, bool]] = []

        cells_selected = self.select_current_table_all_cells(row_count, col_count)
        markdown_style_name = self.table_settings.get("markdown_table", {}).get("paragraph_style", "(적용 안 함)")
        if cells_selected:
            if markdown_style_name and markdown_style_name != "(적용 안 함)":
                style_ok = self.apply_style_by_name(markdown_style_name)
            else:
                style_ok = True
            cell_margin_action, cell_margin_ok = self.set_table_cell_margins()
            table_style_preset = self.table_settings.get("markdown_table", {}).get("table_style_preset", "thin_top_bottom")
            table_style_results, table_style_ok, _table_style_warning = self.apply_table_style_preset_steps(table_style_preset)
        else:
            style_ok = False if markdown_style_name and markdown_style_name != "(적용 안 함)" else True
            cell_margin_action, cell_margin_ok = "셀선택", False
            table_style_preset = self.table_settings.get("markdown_table", {}).get("table_style_preset", "thin_top_bottom")
            table_style_results, table_style_ok = ["셀선택=False"], False
        if markdown_style_name and markdown_style_name != "(적용 안 함)":
            results.append((f"문단스타일:{markdown_style_name}", style_ok))
        results.append(("셀여백", cell_margin_ok))
        table_style_label = TABLE_STYLE_PRESET_LABELS.get(table_style_preset, table_style_preset)
        results.append((f"표스타일:{table_style_label}", table_style_ok))

        self.clear_hwp_selection()
        self.select_current_table_object()
        outside_action, outside_ok = self.set_table_outside_margins()
        results.append(("표여백", outside_ok))

        width_action, width_ok, _width = self.set_table_page_width()
        results.append(("너비맞추기", width_ok))

        self.debug(
            f"[markdown-table-format] {label}: "
            f"style={markdown_style_name}:{style_ok}, "
            f"cell_margin={cell_margin_action}:{cell_margin_ok}, "
            f"table_style={table_style_preset}:{table_style_ok}:{table_style_results}, "
            f"outside={outside_action}:{outside_ok}, width={width_action}:{width_ok}"
        )
        return results

    def fill_current_table_with_rows(self, rows: list[list[str]]) -> bool:
        if not rows or not rows[0]:
            return False
        col_count = len(rows[0])
        entered_first_cell = self.is_in_table_cell()
        if not entered_first_cell:
            entered_first_cell = self.select_current_table_object() and self.select_selected_table_first_cell()
            if entered_first_cell:
                self.clear_hwp_selection()
        if not entered_first_cell:
            return False

        for row_index, row in enumerate(rows):
            if row_index > 0:
                if not self.move_table_cell("TableLowerCell"):
                    return False
                if col_count > 1 and not self.move_table_cell("TableLeftCell", col_count - 1):
                    return False
            for col_index, cell_text in enumerate(row):
                if col_index > 0 and not self.move_table_cell("TableRightCell"):
                    return False
                self.clear_hwp_selection()
                if not self.insert_hwp_text(normalize_markdown_table_cell_text(cell_text)):
                    return False
        return True

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

    def apply_configured_outline_styles_to_selection(self) -> None:
        label = "일괄처리"
        if not self.ensure_hwp():
            return
        try:
            active_set = self.active_style_set()
            rule_count = sum(1 for entry in active_set.paragraph_styles if entry.outline_markers)
            inline_rule_count = sum(1 for rule in active_set.inline_rules if rule.enabled)
            if rule_count == 0 and inline_rule_count == 0:
                messagebox.showinfo(label, "현재 스타일 세트에 일괄처리 규칙이 없습니다.")
                self.log(f"{label}: 규칙 없음, set={active_set.name}")
                return

            if self.is_selected_cell_block():
                self.apply_configured_outline_styles_to_selected_cells(label, active_set, rule_count, inline_rule_count)
                return

            paragraph_range = self.selected_paragraph_range()
            if paragraph_range is None:
                paragraph_range = self.current_paragraph_range()
                if paragraph_range is None:
                    messagebox.showwarning(label, "한글에서 변환할 문단 영역을 먼저 선택하세요.")
                    self.log(f"{label}: 문단 선택 범위 확인 실패")
                    return
                self.log(f"{label}: 선택 범위 확인 실패, 현재 커서 문단 1개로 진행 range={paragraph_range}")

            self.ensure_current_doc_style_cache()
            list_id, first_para, last_para = paragraph_range
            original_pos = self.get_hwp_pos_by_set()
            self.clear_hwp_selection()
            warn_on_existing_title_boxes = self.has_title_number_boxes_after_paragraph(list_id, last_para)
            visited = 0
            matched = 0
            changed = 0
            inline_applied = 0
            markdown_converted = 0
            missing_styles: set[str] = set()
            missing_roles: set[str] = set()
            title_box_followup_styles: set[str] = set()
            numbering_state = NumberingRunState()

            para = first_para
            while para <= last_para:
                self.clear_hwp_selection()
                visited += 1
                markdown_block = self.read_markdown_table_block(list_id, para, last_para)
                if markdown_block is not None:
                    markdown_end_para, markdown_text, markdown_rows = markdown_block
                    markdown_result = self.replace_markdown_table_block(
                        label,
                        list_id,
                        para,
                        markdown_end_para,
                        markdown_text,
                        markdown_rows,
                    )
                    consumed_paragraphs = markdown_end_para - para + 1
                    if markdown_result in {"converted", "partial"}:
                        markdown_converted += 1
                        changed += 1
                        last_para -= max(0, consumed_paragraphs - 1)
                        para += 1
                        continue
                    self.log(
                        f"{label}: markdown 표 변환 건너뜀 paras={para}-{markdown_end_para}, "
                        f"result={markdown_result}"
                    )
                    para = markdown_end_para + 1
                    continue

                processed, para_matched, para_changed, para_inline_applied = self.process_configured_style_paragraph(
                    label,
                    list_id,
                    para,
                    active_set,
                    missing_styles,
                    missing_roles,
                    numbering_state=numbering_state,
                    title_box_followup_styles=title_box_followup_styles,
                    warn_on_existing_title_boxes=warn_on_existing_title_boxes,
                )
                if not processed:
                    continue
                if para_matched:
                    matched += 1
                if para_changed:
                    changed += 1
                inline_applied += para_inline_applied
                para += 1

            if original_pos is not None:
                self.set_hwp_pos_by_set(original_pos)
            self.activate_hwp_window()
            self.log(
                f"{label}: 문단 순회 완료, set={active_set.name}, "
                f"outline_rules={rule_count}, inline_rules={inline_rule_count}, "
                f"visited={visited}, matched={matched}, changed={changed}, inline_applied={inline_applied}, "
                f"title_box_followup_styles={sorted(title_box_followup_styles)}, "
                f"markdown_converted={markdown_converted}, "
                f"numbering_restarted={numbering_state.restarted}, "
                f"numbering_restart_failed={numbering_state.restart_failed}"
            )
            if missing_styles or missing_roles or title_box_followup_styles:
                parts = []
                if missing_styles:
                    parts.append("현재 문서에 같은 이름의 문단 스타일이 없어 건너뜀:\n" + "\n".join(sorted(missing_styles)))
                if missing_roles:
                    parts.append("현재 세트/문서에서 글자 역할을 찾지 못해 건너뜀:\n" + "\n".join(sorted(missing_roles)))
                if title_box_followup_styles:
                    style_names = ", ".join(sorted(title_box_followup_styles))
                    parts.append(
                        f"기존 {style_names} 항목이 있는 문서에 새 {style_names} 항목을 넣었습니다.\n"
                        f"문서 중간에 삽입한 경우 뒤쪽 {style_names} 번호를 확인하고 필요하면 수정하세요."
                    )
                messagebox.showwarning(
                    label,
                    ("일부 규칙을 건너뛰었습니다.\n\n" if (missing_styles or missing_roles) else "")
                    + "\n\n".join(parts)
                    + self.current_doc_style_reconnect_hint(),
                )
            elif changed == 0 and inline_applied == 0:
                messagebox.showinfo(label, "선택한 문단에서 일괄처리 규칙과 일치하는 텍스트를 찾지 못했습니다.")
        except Exception as exc:
            self.log(f"{label} 실패: {type(exc).__name__}: {exc}")
            messagebox.showerror(f"{label} 실패", str(exc))

    def apply_configured_outline_styles_to_selected_cells(
        self,
        label: str,
        active_set: StyleSet,
        rule_count: int,
        inline_rule_count: int,
    ) -> None:
        cell_range = self.get_selected_cell_range_by_formula()
        addresses = [] if cell_range is None else list(cell_range.get("addresses") or [])
        if not addresses:
            messagebox.showwarning(label, "선택한 셀 주소를 확인하지 못해 일괄처리를 중단했습니다.")
            self.log(f"{label}: 셀 주소 확인 실패")
            return

        self.ensure_current_doc_style_cache()
        original_pos = self.get_hwp_pos_by_set()
        address_positions = self.snapshot_formula_address_positions(addresses)
        if address_positions is None:
            if original_pos is not None:
                self.set_hwp_pos_by_set(original_pos)
            messagebox.showwarning(label, "선택한 셀 위치를 확인하지 못해 일괄처리를 중단했습니다.")
            self.log(f"{label}: 셀 위치 스냅샷 실패")
            return
        self.debug(f"[bulk-cell-style] TableFormula range={cell_range}, positions={len(address_positions)}")

        visited = 0
        matched = 0
        changed = 0
        cells_visited = 0
        cells_restore_failed = 0
        cells_range_failed = 0
        missing_styles: set[str] = set()
        missing_roles: set[str] = set()
        numbering_state = NumberingRunState()
        warn_on_existing_title_boxes = bool(self.current_title_number_box_numbers())

        for address, cell_pos in address_positions:
            if not self.set_hwp_pos_by_set(cell_pos):
                cells_restore_failed += 1
                self.debug(f"[bulk-cell-style] 셀 위치 복원 실패: address={address}")
                continue
            paragraph_positions = self.scan_current_cell_paragraph_positions()
            if not paragraph_positions:
                paragraph_range = self.selected_current_cell_paragraph_range()
                if paragraph_range is not None:
                    list_id, first_para, last_para = paragraph_range
                    paragraph_positions = [(list_id, para) for para in range(first_para, last_para + 1)]
            if not paragraph_positions:
                cells_range_failed += 1
                self.debug(f"[bulk-cell-style] 셀 문단 범위 확인 실패: address={address}")
                continue
            cells_visited += 1
            for list_id, para in paragraph_positions:
                visited += 1
                processed, para_matched, para_changed, para_inline_applied = self.process_configured_style_paragraph(
                    label,
                    list_id,
                    para,
                    active_set,
                    missing_styles,
                    missing_roles,
                    force_in_table=True,
                    numbering_state=numbering_state,
                    warn_on_existing_title_boxes=warn_on_existing_title_boxes,
                )
                if not processed:
                    continue
                if para_matched:
                    matched += 1
                if para_changed or para_inline_applied > 0:
                    changed += 1

        if original_pos is not None:
            self.set_hwp_pos_by_set(original_pos)
        self.activate_hwp_window()
        self.log(
            f"{label}: 셀 순회 완료, set={active_set.name}, "
            f"outline_rules={rule_count}, inline_rules={inline_rule_count}, "
            f"cells={cells_visited}/{len(address_positions)}, visited={visited}, "
            f"matched={matched}, changed={changed}, "
            f"numbering_restarted={numbering_state.restarted}, "
            f"numbering_restart_failed={numbering_state.restart_failed}, "
            f"restore_failed={cells_restore_failed}, range_failed={cells_range_failed}"
        )
        if missing_styles or missing_roles:
            parts = []
            if missing_styles:
                parts.append("현재 문서에 같은 이름의 문단 스타일이 없어 건너뜀:\n" + "\n".join(sorted(missing_styles)))
            if missing_roles:
                parts.append("현재 세트/문서에서 글자 역할을 찾지 못해 건너뜀:\n" + "\n".join(sorted(missing_roles)))
            messagebox.showwarning(
                label,
                "일부 규칙을 건너뛰었습니다.\n\n"
                + "\n\n".join(parts)
                + self.current_doc_style_reconnect_hint(),
            )
        elif cells_range_failed and visited == 0:
            messagebox.showwarning(
                label,
                "선택한 셀 안의 문단 위치를 확인하지 못해 일괄처리를 적용하지 못했습니다.\n\n"
                "로그의 [bulk-cell-style] 셀 선택 진단 값을 확인해 주세요.",
            )
        elif cells_range_failed:
            messagebox.showwarning(
                label,
                f"일부 셀의 문단 위치를 확인하지 못해 건너뛰었습니다.\n\n"
                f"처리한 셀: {cells_visited}개 / 실패한 셀: {cells_range_failed}개",
            )
        elif changed == 0:
            messagebox.showinfo(label, "선택한 셀에서 일괄처리 규칙과 일치하는 텍스트를 찾지 못했습니다.")

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
            stripped_selection = text.strip()
            if is_standard_regulation_wrapper(stripped_selection):
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
            if needs_regulation_wrapper_conversion(stripped_selection):
                if start[0] != end[0] or start[1] != end[1]:
                    messagebox.showwarning(
                        label,
                        "괄호 변환은 한 문단 안에서 선택한 경우에만 가능합니다.",
                    )
                    self.log(f"{label}: 괄호 변환 범위가 한 문단을 벗어나 중단")
                    return
                list_id, para, start_pos = start
                leading_offset = len(text) - len(text.lstrip())
                close_pos = start_pos + len(text.rstrip()) - 1
                open_pos = start_pos + leading_offset
                if close_pos <= open_pos:
                    messagebox.showwarning(label, "괄호 위치를 확인하지 못했습니다.")
                    self.log(f"{label}: 괄호 위치 계산 실패")
                    return
                if not self.delete_hwp_text_range(list_id, para, close_pos, close_pos + 1) or not self.set_hwp_pos(
                    (list_id, para, close_pos)
                ) or not self.insert_hwp_text("｣"):
                    messagebox.showwarning(label, "닫는 괄호 변환에 실패했습니다.")
                    self.log(f"{label}: 닫는 괄호 변환 실패")
                    return
                if not self.delete_hwp_text_range(list_id, para, open_pos, open_pos + 1) or not self.set_hwp_pos(
                    (list_id, para, open_pos)
                ) or not self.insert_hwp_text("｢"):
                    messagebox.showwarning(label, "여는 괄호 변환에 실패했습니다.")
                    self.log(f"{label}: 여는 괄호 변환 실패")
                    return
                self.activate_hwp_window()
                self.log(f"{label}: 기존 괄호를 표준 괄호로 변환 완료")
                return

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

    def get_decimal_transform_options(
        self,
        places_var: tk.StringVar,
        mode_var: tk.StringVar,
        use_commas_var: tk.BooleanVar,
        warning_title: str,
    ) -> tuple[int, str, bool] | None:
        try:
            places = int(places_var.get().strip())
        except ValueError:
            messagebox.showwarning(warning_title, "소숫점 자릿수는 0 이상 정수로 입력하세요.")
            return None
        if places < 0:
            messagebox.showwarning(warning_title, "소숫점 자릿수는 0 이상 정수로 입력하세요.")
            return None

        return places, mode_var.get() or "반올림", bool(use_commas_var.get())

    def apply_decimal_rounding_to_selection(self) -> None:
        options = self.get_decimal_transform_options(
            self.decimal_places_var,
            self.decimal_mode_var,
            self.decimal_use_commas_var,
            "소숫점 자릿수 확인",
        )
        if options is None:
            return
        places, mode, use_commas = options
        self.transform_selected_text(
            f"소숫점 처리: {mode} {places}자리",
            lambda text: normalize_decimal_numbers(text, places, mode, use_commas),
            block_multiline_number_block=True,
            allow_cell_iteration=True,
            strip_wrapping_lines=True,
            reselect_current_cell=True,
            preflight=ensure_no_suspicious_decimal_numbers,
        )

    def apply_decimal_unit_conversion_to_selection(
        self,
        label: str,
        multiplier: Decimal,
        unit_transitions: dict[str, str],
    ) -> None:
        options = self.get_decimal_transform_options(
            self.unit_decimal_places_var,
            self.unit_decimal_mode_var,
            self.unit_decimal_use_commas_var,
            "단위 변환 소숫점 자릿수 확인",
        )
        if options is None:
            return
        places, mode, use_commas = options
        self.transform_selected_text(
            f"단위 변환: {label}, {mode} {places}자리",
            lambda text: scale_decimal_numbers_for_unit_conversion(
                text,
                multiplier,
                unit_transitions,
                places,
                mode,
                use_commas,
            ),
            block_multiline_number_block=True,
            allow_cell_iteration=True,
            strip_wrapping_lines=True,
            reselect_current_cell=True,
            preflight=lambda text: ensure_safe_decimal_unit_conversion_text(text, set(unit_transitions)),
        )

    def apply_currency_unit_conversion_to_selection(self, label: str, target_unit: str) -> None:
        options = self.get_decimal_transform_options(
            self.unit_decimal_places_var,
            self.unit_decimal_mode_var,
            self.unit_decimal_use_commas_var,
            "단위 변환 소숫점 자릿수 확인",
        )
        if options is None:
            return
        places, mode, use_commas = options
        allowed_units = {"원", "천", "천원", "백만", "백만원"}
        self.transform_selected_text(
            f"단위 변환: {label}, {mode} {places}자리",
            lambda text: convert_decimal_numbers_to_currency_unit(
                text,
                target_unit,
                places,
                mode,
                use_commas,
            ),
            block_multiline_number_block=True,
            allow_cell_iteration=True,
            strip_wrapping_lines=True,
            reselect_current_cell=True,
            preflight=lambda text: ensure_safe_decimal_unit_conversion_text(text, allowed_units),
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

    def normalize_years_to_selection(self, style: str, label: str) -> None:
        self.transform_selected_text(
            f"연도 정규화: {label}",
            lambda text: normalize_years(text, style),
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
    print(f"version={APP_VERSION}")
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
