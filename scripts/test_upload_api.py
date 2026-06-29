"""Quick smoke test for upload server."""

from __future__ import annotations

import json
import sys
import urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
TEST = ROOT / "tmp_test_upload.xlsx"


def main() -> int:
    if not TEST.is_file():
        print("Missing", TEST)
        return 1

    boundary = "----WebKitFormBoundary7MA4YWxkTrZu0gW"
    parts: list[bytes] = []
    for key, val in [("section", "newJoiner"), ("type", "tool")]:
        parts.append(
            f"--{boundary}\r\nContent-Disposition: form-data; name=\"{key}\"\r\n\r\n{val}\r\n".encode()
        )
    parts.append(
        (
            f"--{boundary}\r\n"
            f'Content-Disposition: form-data; name="file"; filename="test.xlsx"\r\n'
            f"Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\r\n\r\n"
        ).encode()
    )
    parts.append(TEST.read_bytes())
    parts.append(f"\r\n--{boundary}--\r\n".encode())
    body = b"".join(parts)

    req = urllib.request.Request(
        "http://127.0.0.1:8765/api/upload",
        data=body,
        method="POST",
        headers={"Content-Type": f"multipart/form-data; boundary={boundary}"},
    )
    with urllib.request.urlopen(req, timeout=10) as resp:
        data = json.loads(resp.read().decode())
    print(json.dumps(data, indent=2))
    path = ROOT / data["path"]
    print("exists:", path.is_file())
    return 0 if data.get("ok") and path.is_file() else 1


if __name__ == "__main__":
    sys.exit(main())
