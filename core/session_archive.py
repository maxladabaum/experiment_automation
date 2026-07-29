"""Create a local session ZIP and publish it to a configured archive share."""

from __future__ import annotations

import os
import shutil
import uuid
import zipfile
from pathlib import Path


def archive_session(session_path: str | Path, destination_dir: str | Path) -> Path:
    """ZIP a session locally, then atomically publish it to the destination."""
    source = Path(session_path).resolve()
    if not source.is_dir():
        raise FileNotFoundError(f"Session folder not found: {source}")

    destination = Path(destination_dir)
    destination.mkdir(parents=True, exist_ok=True)
    staging_dir = source.parent / ".archive_staging"
    staging_dir.mkdir(parents=True, exist_ok=True)
    local_zip = staging_dir / f"{source.name}_{uuid.uuid4().hex[:8]}.zip"
    final_zip = destination / f"{source.name}.zip"
    remote_part = destination / f".{source.name}.{uuid.uuid4().hex[:8]}.zip.part"

    try:
        with zipfile.ZipFile(local_zip, "w", zipfile.ZIP_DEFLATED) as archive:
            for path in sorted(source.rglob("*")):
                if path.is_file():
                    archive.write(path, path.relative_to(source.parent))
        shutil.copy2(local_zip, remote_part)
        os.replace(remote_part, final_zip)
        return final_zip
    finally:
        local_zip.unlink(missing_ok=True)
        remote_part.unlink(missing_ok=True)
        try:
            staging_dir.rmdir()
        except OSError:
            pass
