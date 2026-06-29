"""Repair dashboard/index.html: zone CSS, remove broken download buttons."""

from __future__ import annotations

import re
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
HTML = ROOT / "dashboard" / "index.html"

OLD_CSS = """  .zone-status-grid {
    display: grid;
    grid-template-columns: repeat(auto-fill, minmax(160px, 1fr));
    gap: 10px;
  }
  .zone-status-card {
    background: var(--bg3);
    border: 1px solid var(--border);
    border-radius: 10px;
    padding: 12px 14px;
  }
  .zone-status-name {
    font-family: Arial, sans-serif;
    font-size: 13px;
    font-weight: 700;
    color: var(--accent2);
    margin-bottom: 8px;
  }
  .zone-status-stats {
    display: flex;
    flex-wrap: wrap;
    gap: 8px 12px;
    font-size: 11px;
    color: var(--text2);
    margin-bottom: 8px;
  }
  .zone-stat-pending { color: var(--red); }
  .zone-status-pct {
    font-size: 10px;
    color: var(--text3);
    margin-top: 4px;
  }"""

NEW_CSS = """  .zone-status-grid {
    display: grid;
    grid-template-columns: repeat(8, minmax(0, 1fr));
    gap: 6px;
    width: 100%;
  }
  .zone-status-card {
    background: var(--bg3);
    border: 1px solid var(--border);
    border-radius: 8px;
    padding: 8px 6px;
    min-width: 0;
  }
  .zone-status-name {
    font-family: Arial, sans-serif;
    font-size: 11px;
    font-weight: 700;
    color: var(--accent2);
    margin-bottom: 4px;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
  }
  .zone-status-stats {
    display: flex;
    flex-wrap: nowrap;
    justify-content: space-between;
    gap: 2px;
    font-size: 9px;
    color: var(--text2);
    margin-bottom: 4px;
  }
  .zone-stat {
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
  }
  .zone-stat-pending { color: var(--red); }
  .zone-status-pct {
    font-size: 9px;
    color: var(--text3);
    margin-top: 2px;
    white-space: nowrap;
  }
  .zone-status-card .progress-bar-wrap { margin-top: 4px; }
  .zone-dl-actions {
    display: flex;
    gap: 4px;
    margin-top: 6px;
  }
  .zone-dl-btn {
    flex: 1 1 0;
    min-width: 0;
    padding: 4px 2px;
    font-size: 8px;
    font-weight: 600;
    font-family: Arial, sans-serif;
    border-radius: 4px;
    border: 1px solid var(--border);
    background: var(--bg4);
    color: var(--text2);
    cursor: pointer;
    line-height: 1.2;
  }
  .zone-dl-btn:hover:not(:disabled) {
    border-color: var(--accent);
    color: var(--accent2);
  }
  .zone-dl-completed:hover:not(:disabled) {
    background: rgba(74, 184, 64, 0.12);
    border-color: var(--green);
    color: var(--green);
  }
  .zone-dl-pending:hover:not(:disabled) {
    background: rgba(212, 64, 64, 0.1);
    border-color: var(--red);
    color: var(--red);
  }
  .zone-dl-btn:disabled {
    opacity: 0.35;
    cursor: not-allowed;
  }
  @media (max-width: 900px) {
    .zone-status-grid {
      display: flex;
      flex-wrap: nowrap;
      overflow-x: auto;
      gap: 6px;
      padding-bottom: 4px;
      scrollbar-width: thin;
    }
    .zone-status-card {
      flex: 0 0 calc((100% - 42px) / 8);
      min-width: 88px;
    }
  }"""


def main() -> None:
    text = HTML.read_text(encoding="utf-8")
    lines = text.splitlines()
    text = "\n".join(
        line
        for line in lines
        if "btn-dl-zone-" not in line and "btn-dl-pending-zone-" not in line
    )
    if OLD_CSS not in text and "repeat(8, minmax(0, 1fr))" not in text:
        if OLD_CSS in text:
            text = text.replace(OLD_CSS, NEW_CSS)
        else:
            raise SystemExit("Expected zone CSS block not found")
    elif OLD_CSS in text:
        text = text.replace(OLD_CSS, NEW_CSS)
    if 'id="server-status"' not in text:
        text = text.replace(
            '<div class="header-right">',
            '<div class="header-right">'
            '<span id="server-status" class="header-badge" style="display:none"></span>',
            1,
        )
    HTML.write_text(text, encoding="utf-8", newline="\n")
    n = text.count("downloadByZoneSection")
    print(f"Wrote {HTML}; downloadByZoneSection refs: {n}")


if __name__ == "__main__":
    main()
