# -*- coding: utf-8 -*-
"""Template asset path and synchronization helpers for BUFS HWP Editor."""

from __future__ import annotations

import filecmp
import shutil
from pathlib import Path


def sync_builtin_templates(
    bundled_dir: Path,
    builtin_dir: Path,
    custom_dir: Path,
) -> list[str]:
    """Copy bundled templates into the user builtin template directory.

    Returns a list of warning messages instead of raising so app startup can
    fall back to bundled read-only assets when the user config directory is not
    writable.
    """
    warnings: list[str] = []
    try:
        builtin_dir.mkdir(parents=True, exist_ok=True)
        custom_dir.mkdir(parents=True, exist_ok=True)
    except Exception as exc:
        return [f"템플릿 사용자 폴더를 만들지 못했습니다: {exc}"]

    if not bundled_dir.exists():
        return [f"번들 템플릿 폴더를 찾지 못했습니다: {bundled_dir}"]

    for source in bundled_dir.rglob("*"):
        if not source.is_file():
            continue
        relative = source.relative_to(bundled_dir)
        target = builtin_dir / relative
        try:
            target.parent.mkdir(parents=True, exist_ok=True)
            if target.exists() and filecmp.cmp(source, target, shallow=False):
                continue
            shutil.copy2(source, target)
        except Exception as exc:
            warnings.append(f"템플릿 복사 실패: {relative}: {exc}")
    return warnings


def builtin_template_dir_or_fallback(builtin_dir: Path, bundled_dir: Path) -> Path:
    """Prefer synced user builtin templates, falling back to bundled assets."""
    return builtin_dir if builtin_dir.exists() else bundled_dir
