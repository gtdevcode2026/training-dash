"""Create trainings/<slug>/uploads/base and uploads/tool for every module."""

from __future__ import annotations

import sys
from pathlib import Path

_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(_ROOT / "shared" / "python"))

from upload_paths import TRAININGS_DIR, ensure_upload_tree, load_training_configs  # noqa: E402


def main() -> None:
    dirs = ensure_upload_tree()
    print(f"Upload folders ready under {TRAININGS_DIR}:")
    for cfg in load_training_configs():
        slug = cfg["slug"]
        print(f"  - trainings/{slug}/uploads/base/")
        print(f"  - trainings/{slug}/uploads/tool/")
    print(f"({len(dirs)} directories checked)")


if __name__ == "__main__":
    main()
