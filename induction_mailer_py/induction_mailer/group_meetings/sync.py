from __future__ import annotations

import subprocess
from pathlib import Path


def rclone_copy(
    *,
    rclone_exe: Path,
    source: str,
    dest_dir: Path,
    logger=None,
) -> bool:
    """Port of ..."""

    exe = Path(rclone_exe)
    dest_dir = Path(dest_dir)
    dest_dir.mkdir(parents=True, exist_ok=True)

    cmd = [str(exe), "copy", source, str(dest_dir)]
    if logger:
        try:
            logger.verbose(f"Running rclone command: {' '.join(cmd)}")
        except Exception:
            pass
    if not exe.exists():
        if logger:
            try:
                logger.error(f"rclone.exe not found at: {exe}")
            except Exception:
                pass
        return False


def rclone_copyto(
    *,
    rclone_exe: Path,
    source: str | Path,
    dest: str | Path,
    logger=None,
) -> bool:
    """Copy one file to one exact destination path with rclone copyto."""

    exe = Path(rclone_exe)
    cmd = [str(exe), "copyto", str(source), str(dest), "--ignore-times"]
    if logger:
        try:
            logger.verbose(f"Running rclone command: {' '.join(cmd)}")
        except Exception:
            pass
    if not exe.exists():
        if logger:
            try:
                logger.error(f"rclone.exe not found at: {exe}")
            except Exception:
                pass
        return False

    try:
        proc = subprocess.run(cmd, capture_output=True, text=True, check=False)
        ok = proc.returncode == 0
        if logger:
            try:
                logger.log(f"rclone exit code: {proc.returncode}")
                if proc.stdout.strip():
                    logger.verbose(f"rclone stdout: {proc.stdout.strip()}")
                if proc.stderr.strip():
                    logger.verbose(f"rclone stderr: {proc.stderr.strip()}")
            except Exception:
                pass
        return ok
    except Exception as e:
        if logger:
            try:
                logger.error("rclone invocation failed", e)
            except Exception:
                pass
        return False

    try:
        proc = subprocess.run(cmd, capture_output=True, text=True, check=False)
        ok = proc.returncode == 0
        if logger:
            try:
                logger.log(f"rclone exit code: {proc.returncode}")
                if proc.stdout.strip():
                    logger.verbose(f"rclone stdout: {proc.stdout.strip()}")
                if proc.stderr.strip():
                    logger.verbose(f"rclone stderr: {proc.stderr.strip()}")
            except Exception:
                pass
        return ok
    except Exception as e:
        if logger:
            try:
                logger.error("rclone invocation failed", e)
            except Exception:
                pass
        return False
