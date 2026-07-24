# -*- coding: utf-8 -*-
from __future__ import annotations

import argparse
import ctypes
import shutil
import subprocess
import sys
import tempfile
import time
import zipfile
from datetime import datetime
from pathlib import Path


def log(message: str) -> None:
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{timestamp}] {message}"
    try:
        log_path = Path(tempfile.gettempdir()) / "BUFS-HWP-Updater.log"
        with log_path.open("a", encoding="utf-8") as file:
            file.write(line + "\n")
    except Exception:
        pass


def safe_zip_member_path(name: str) -> bool:
    if not name or name.startswith(("/", "\\")):
        return False
    parts = Path(name.replace("\\", "/")).parts
    return all(part not in ("", ".", "..") for part in parts)


def validate_update_zip(zip_path: Path, root_dir: str, exe_name: str) -> None:
    expected_root = root_dir.strip().strip("/\\")
    if not expected_root:
        raise ValueError("Expected root directory is empty.")
    with zipfile.ZipFile(zip_path) as archive:
        names = [name.replace("\\", "/") for name in archive.namelist()]
    if not names:
        raise ValueError("Update zip is empty.")
    unsafe = [name for name in names if not safe_zip_member_path(name)]
    if unsafe:
        raise ValueError(f"Unsafe path in update zip: {unsafe[0]}")
    prefix = expected_root + "/"
    outside = [name for name in names if name != expected_root and not name.startswith(prefix)]
    if outside:
        raise ValueError(f"Update zip must contain one top-level folder: {expected_root}")
    if prefix + exe_name not in names:
        raise ValueError(f"Missing executable in update zip: {prefix}{exe_name}")
    if not any(name.startswith(prefix + "_internal/") for name in names):
        raise ValueError("Missing _internal folder in update zip.")


def wait_for_pid(pid: int, timeout_seconds: int = 60) -> None:
    if pid <= 0:
        return
    if sys.platform != "win32":
        deadline = time.time() + timeout_seconds
        while time.time() < deadline:
            try:
                import os

                os.kill(pid, 0)
            except OSError:
                return
            time.sleep(0.5)
        raise TimeoutError(f"Process did not exit: {pid}")

    synchronize = 0x00100000
    wait_timeout = 0x00000102
    handle = ctypes.windll.kernel32.OpenProcess(synchronize, False, pid)
    if not handle:
        return
    try:
        result = ctypes.windll.kernel32.WaitForSingleObject(handle, timeout_seconds * 1000)
        if result == wait_timeout:
            raise TimeoutError(f"Process did not exit: {pid}")
    finally:
        ctypes.windll.kernel32.CloseHandle(handle)


def unique_backup_dir(install_dir: Path, version: str) -> Path:
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    base = install_dir.with_name(f"{install_dir.name}.backup-{version}-{stamp}")
    candidate = base
    index = 1
    while candidate.exists():
        candidate = install_dir.with_name(f"{base.name}-{index}")
        index += 1
    return candidate


def extract_staged(zip_path: Path, root_dir: str) -> Path:
    stage_parent = Path(tempfile.mkdtemp(prefix="BUFS-HWP-Update-"))
    with zipfile.ZipFile(zip_path) as archive:
        archive.extractall(stage_parent)
    staged = stage_parent / root_dir
    if not staged.is_dir():
        raise ValueError(f"Extracted folder not found: {staged}")
    return staged


def replace_install_dir(install_dir: Path, staged_dir: Path, version: str) -> Path:
    install_dir = install_dir.resolve()
    parent = install_dir.parent
    if not parent.exists():
        raise FileNotFoundError(f"Install parent does not exist: {parent}")
    backup_dir = unique_backup_dir(install_dir, version)
    if install_dir.exists():
        shutil.move(str(install_dir), str(backup_dir))
    try:
        shutil.move(str(staged_dir), str(install_dir))
    except Exception:
        if install_dir.exists():
            shutil.rmtree(install_dir, ignore_errors=True)
        if backup_dir.exists():
            shutil.move(str(backup_dir), str(install_dir))
        raise
    return backup_dir


def main() -> int:
    parser = argparse.ArgumentParser(description="BUFS HWP Editor updater")
    parser.add_argument("--pid", type=int, required=True)
    parser.add_argument("--install-dir", required=True)
    parser.add_argument("--zip", required=True)
    parser.add_argument("--root-dir", default="BUFS-HWP-Editor")
    parser.add_argument("--exe-name", default="BUFS-HWP-Editor.exe")
    parser.add_argument("--version", default="latest")
    args = parser.parse_args()

    install_dir = Path(args.install_dir)
    zip_path = Path(args.zip)
    try:
        log(f"Starting update to {args.version}")
        validate_update_zip(zip_path, args.root_dir, args.exe_name)
        wait_for_pid(args.pid)
        staged_dir = extract_staged(zip_path, args.root_dir)
        backup_dir = replace_install_dir(install_dir, staged_dir, args.version)
        exe_path = install_dir / args.exe_name
        subprocess.Popen([str(exe_path)], cwd=str(install_dir), close_fds=True)
        time.sleep(2)
        if backup_dir.exists():
            shutil.rmtree(backup_dir, ignore_errors=True)
        log("Update completed")
        return 0
    except Exception as exc:
        log(f"Update failed: {type(exc).__name__}: {exc}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
