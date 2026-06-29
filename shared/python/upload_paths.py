"""
Paths for dashboard uploads under trainings/<slug>/uploads/.
"""

from __future__ import annotations

import json
import re
import shutil
from datetime import datetime
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[2]
TRAININGS_DIR = ROOT / "trainings"

UPLOAD_TYPES = ("base", "tool")


def load_training_configs() -> list[dict[str, Any]]:
    configs: list[dict[str, Any]] = []
    if not TRAININGS_DIR.is_dir():
        return configs
    for cfg_path in sorted(TRAININGS_DIR.glob("*/config.json")):
        try:
            data = json.loads(cfg_path.read_text(encoding="utf-8-sig"))
        except (OSError, json.JSONDecodeError):
            continue
        if "slug" in data and "dashboardTab" in data:
            data["_config_path"] = str(cfg_path)
            configs.append(data)
    return configs


def section_to_slug() -> dict[str, str]:
    return {c["dashboardTab"]: c["slug"] for c in load_training_configs()}


def slug_to_section() -> dict[str, str]:
    return {c["slug"]: c["dashboardTab"] for c in load_training_configs()}


def safe_filename(name: str) -> str:
    name = Path(name).name
    name = re.sub(r'[<>:"/\\|?*]', "_", name)
    return name or "upload.xlsx"


def upload_type_dir(slug: str, file_type: str) -> Path:
    if file_type not in UPLOAD_TYPES:
        raise ValueError(f"file_type must be one of {UPLOAD_TYPES}")
    return TRAININGS_DIR / slug / "uploads" / file_type


def dated_upload_dir(slug: str, file_type: str, run_date: str | None = None) -> Path:
    day = run_date or datetime.now().strftime("%Y-%m-%d")
    return upload_type_dir(slug, file_type) / day


def ensure_upload_tree() -> list[Path]:
    """Create uploads/base and uploads/tool under every training module."""
    created: list[Path] = []
    for cfg in load_training_configs():
        slug = cfg["slug"]
        for file_type in UPLOAD_TYPES:
            d = upload_type_dir(slug, file_type)
            d.mkdir(parents=True, exist_ok=True)
            keep = d / ".gitkeep"
            if not keep.exists():
                keep.write_text("", encoding="utf-8")
            created.append(d)
    return created


def save_upload(
    section: str,
    file_type: str,
    source_path: Path,
    original_name: str,
) -> dict[str, Any]:
    mapping = section_to_slug()
    if section not in mapping:
        raise ValueError(f"Unknown dashboard section: {section}")
    if file_type not in UPLOAD_TYPES:
        raise ValueError(f"file_type must be one of {UPLOAD_TYPES}")

    slug = mapping[section]
    safe = safe_filename(original_name)
    stamp = datetime.now().strftime("%H%M%S")
    day_dir = dated_upload_dir(slug, file_type)
    day_dir.mkdir(parents=True, exist_ok=True)

    dest = day_dir / f"{stamp}_{safe}"
    shutil.copy2(source_path, dest)

    ext = Path(safe).suffix.lower()
    latest = upload_type_dir(slug, file_type) / (
        "_latest" + ext if ext in (".xlsx", ".xls") else "_latest.xlsx"
    )
    shutil.copy2(dest, latest)

    return {
        "ok": True,
        "section": section,
        "slug": slug,
        "type": file_type,
        "path": str(dest.relative_to(ROOT)).replace("\\", "/"),
        "latest": str(latest.relative_to(ROOT)).replace("\\", "/"),
        "filename": dest.name,
    }
