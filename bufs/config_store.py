# -*- coding: utf-8 -*-
"""Small JSON configuration file helpers."""

from __future__ import annotations

import json
import shutil
from pathlib import Path
from typing import Any


def ensure_user_config_file(user_path: Path, bundled_path: Path) -> bool:
    """Copy a bundled config to the user config path when it is missing."""
    if user_path.exists() or not bundled_path.exists():
        return False
    try:
        user_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(bundled_path, user_path)
        return True
    except Exception:
        return False


def read_json_file(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def write_json_file(path: Path, data: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
