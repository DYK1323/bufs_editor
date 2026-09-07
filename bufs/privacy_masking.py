"""Local masking rules for explicitly selected table cells."""

from dataclasses import dataclass
import re
import unicodedata


@dataclass(frozen=True)
class MaskEdit:
    start: int
    end: int
    replacement: str
    kind: str


@dataclass(frozen=True)
class MaskResult:
    text: str
    edits: tuple[MaskEdit, ...]
    skipped: int


EMAIL = re.compile(r"[A-Za-z0-9.!#$%&'*+/=?^_`{|}~-]+@[A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)+")
NUMBER = re.compile(r"(?<![0-9*])[0-9*]+(?:[-. \t][0-9*]+)*(?![0-9*])")
NAME = re.compile(r"(?:[^\W\d_]|[\u0300-\u036f*])+(?:[ \t'\u2019-]+(?:[^\W\d_]|[\u0300-\u036f*])+)*")
PHONE = re.compile(r"(?:01[016789]|02|0[3-6][1-5]|070)[- \t]?[0-9]{3,4}[- \t]?[0-9]{4}")
REGISTRATION = re.compile(r"[0-9]{6}[- \t]?[1-8][0-9]{6}")
CARD = re.compile(r"(?:[0-9]{16}|[0-9]{4}(?P<sep>[- \t])[0-9]{4}(?P=sep)[0-9]{4}(?P=sep)[0-9]{4})")


def _hide_digits(value: str, first: int, last: int) -> str:
    index = 0
    result = []
    for char in value:
        if char in '0123456789':
            result.append('*' if first <= index < last else char)
            index += 1
        else:
            result.append(char)
    return ''.join(result)


def _name(value: str) -> str:
    if '*' in value:
        return value
    clusters: list[tuple[int, int]] = []
    for index, char in enumerate(value):
        if unicodedata.combining(char):
            if not clusters or clusters[-1][1] != index:
                return value
            clusters[-1] = (clusters[-1][0], index + 1)
        elif char.isalpha():
            name = unicodedata.name(char, '')
            if 'LATIN' not in name and 'HANGUL' not in name:
                return value
            clusters.append((index, index + 1))
    hidden = clusters if len(clusters) == 1 else clusters[1:-1] if len(clusters) > 2 else clusters[1:]
    for start, end in reversed(hidden):
        value = value[:start] + '*' + value[end:]
    return value


def mask_personal_text(text: str) -> MaskResult:
    """Protect structured tokens before applying the permitted name fallback."""
    occupied = [False] * len(text)
    edits: list[MaskEdit] = []
    skipped = 0

    def record(start: int, end: int, replacement: str, kind: str) -> None:
        occupied[start:end] = [True] * (end - start)
        if text[start:end] == replacement:
            return
        # Equal-length masks edit only changed runs, preserving intervening styles.
        original = text[start:end]
        if len(original) != len(replacement):
            edits.append(MaskEdit(start, end, replacement, kind))
            return
        cursor = 0
        while cursor < len(original):
            if original[cursor] == replacement[cursor]:
                cursor += 1
                continue
            first = cursor
            while cursor < len(original) and original[cursor] != replacement[cursor]:
                cursor += 1
            edits.append(MaskEdit(start + first, start + cursor, replacement[first:cursor], kind))

    for match in EMAIL.finditer(text):
        value = match.group()
        local, domain = value.rsplit('@', 1)
        masked = local if '*' in local else (local[0] + '*' * (len(local) - 1) if len(local) > 1 else '*')
        record(match.start(), match.end(), masked + '@' + domain, 'email')

    # Adjacent masked labels and identifiers can share one run of asterisks.
    for match in re.finditer(r"[\w*]+(?:[-'\u2019][\w*]+)*", text):
        value = match.group()
        if '*' in value and any(char.isalpha() for char in value):
            occupied[match.start():match.end()] = [True] * len(value)

    for match in NUMBER.finditer(text):
        if any(occupied[match.start():match.end()]):
            continue
        value = match.group()
        digits = ''.join(char for char in value if char in '0123456789')
        if not digits:
            continue
        replacement, kind = value, 'number'
        if '*' in value:
            pass
        elif REGISTRATION.fullmatch(value):
            # Mask structurally plausible identifiers even with a mistyped date.
            replacement, kind = _hide_digits(value, 0, 13), 'registration'
        elif CARD.fullmatch(value):
            replacement, kind = _hide_digits(value, 4, 12), 'card'
        elif PHONE.fullmatch(value):
            prefix = 2 if digits.startswith('02') else 3
            replacement, kind = _hide_digits(value, prefix, len(digits) - 4), 'phone'
        elif re.fullmatch(r'[0-9]{8}', value):
            replacement, kind = _hide_digits(value, 0, 4), 'identifier'
        elif len(digits) >= 10:
            skipped += 1
        record(match.start(), match.end(), replacement, kind)

    for match in NAME.finditer(text):
        # Split at protected tokens rather than masking an email domain again.
        start = match.start()
        for end in range(start, match.end() + 1):
            if end == match.end() or occupied[end]:
                if start < end:
                    value = text[start:end]
                    record(start, end, _name(value), 'name')
                start = end + 1

    edits.sort(key=lambda edit: edit.start)
    result = text
    for edit in reversed(edits):
        result = result[:edit.start] + edit.replacement + result[edit.end:]
    return MaskResult(result, tuple(edits), skipped)
