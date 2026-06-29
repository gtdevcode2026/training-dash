"""
Local server: serves the dashboard and saves uploads under trainings/<slug>/uploads/.
"""

from __future__ import annotations

import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SHARED_PY = ROOT / "shared" / "python"
if str(SHARED_PY) not in sys.path:
    sys.path.insert(0, str(SHARED_PY))

from flask import Flask, jsonify, redirect, request, send_from_directory  # noqa: E402

from upload_paths import (  # noqa: E402
    ROOT as PROJECT_ROOT,
    ensure_upload_tree,
    load_training_configs,
    save_upload,
    section_to_slug,
)

app = Flask(__name__)
PORT = 8765


@app.after_request
def add_cors(resp):
    resp.headers["Access-Control-Allow-Origin"] = "*"
    resp.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS"
    resp.headers["Access-Control-Allow-Headers"] = "Content-Type"
    return resp


@app.route("/")
def home():
    return redirect("/dashboard/index.html")


@app.route("/dashboard/")
@app.route("/dashboard/<path:filename>")
def dashboard_files(filename: str = "index.html"):
    return send_from_directory(PROJECT_ROOT / "dashboard", filename)


@app.route("/shared/<path:filename>")
def shared_files(filename: str):
    return send_from_directory(PROJECT_ROOT / "shared", filename)


@app.route("/dash2.html")
def legacy_dash():
    return send_from_directory(PROJECT_ROOT, "dash2.html")


@app.route("/api/health")
def health():
    ensure_upload_tree()
    return jsonify(
        {
            "ok": True,
            "trainings": [
                {"section": c["dashboardTab"], "slug": c["slug"], "label": c.get("label")}
                for c in load_training_configs()
            ],
        }
    )


@app.route("/api/upload", methods=["POST", "OPTIONS"])
def upload():
    if request.method == "OPTIONS":
        return "", 204

    section = (request.form.get("section") or "").strip()
    file_type = (request.form.get("type") or "").strip().lower()
    upload_file = request.files.get("file")

    if not section or not file_type or not upload_file:
        return jsonify({"ok": False, "error": "Missing section, type, or file"}), 400

    if section not in section_to_slug():
        return jsonify({"ok": False, "error": f"Unknown section: {section}"}), 400

    suffix = Path(upload_file.filename or "").suffix.lower()
    if suffix not in (".xlsx", ".xls"):
        return jsonify({"ok": False, "error": "Only .xlsx or .xls files are allowed"}), 400

    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        upload_file.save(tmp.name)
        tmp_path = Path(tmp.name)

    try:
        result = save_upload(section, file_type, tmp_path, upload_file.filename or "upload.xlsx")
    except ValueError as e:
        return jsonify({"ok": False, "error": str(e)}), 400
    finally:
        tmp_path.unlink(missing_ok=True)

    return jsonify(result)


def main() -> None:
    ensure_upload_tree()
    print(f"Training dashboard: http://127.0.0.1:{PORT}/dashboard/index.html")
    print(f"Upload API:         http://127.0.0.1:{PORT}/api/upload")
    print("Uploads saved under: trainings/<module>/uploads/<base|tool>/YYYY-MM-DD/")
    app.run(host="127.0.0.1", port=PORT, debug=False, threaded=True)


if __name__ == "__main__":
    main()
