"""Patch dashboard/index.html for Phishing Normal metrics UI."""
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
p = ROOT / "dashboard" / "index.html"
text = p.read_text(encoding="utf-8")

old_css = """  .metrics-row {
    display: grid; grid-template-columns: repeat(4, 1fr);
    gap: 1px; background: var(--border);
    border-bottom: 1px solid var(--border);
  }"""
new_css = old_css + """
  .metrics-row.metrics-row-5 {
    grid-template-columns: repeat(5, minmax(0, 1fr));
  }"""
if ".metrics-row.metrics-row-5" not in text:
    text = text.replace(old_css, new_css, 1)

old_block = """            <div class="section-title">Phishing — Normal Users</div>
          </div>
        </div>
        <div class="section-controls">
          <span class="date-label">Assignment Date:</span>
          <input type="date" class="date-input" id="date-phishingNormal" value="">
        </div>
      </div>

      <div class="metrics-row">
        <div class="metric-box">
          <div class="metric-label">Total Users</div>
          <div class="metric-val kpi-accent" id="m-phishingNormal-total">—</div>
        </div>
        <div class="metric-box">
          <div class="metric-label">Completed</div>
          <div class="metric-val kpi-green" id="m-phishingNormal-completed">—</div>
          <div style="font-size:11px;color:var(--accent);font-family:Arial,sans-serif;margin-top:2px" id="m-phishingNormal-pct">0%</div>
          <div class="progress-bar-wrap">
            <div class="progress-bar-bg"><div class="progress-bar-fill" id="pb-phishingNormal" style="width:0%;background:var(--accent)"></div></div>
          </div>
        </div>
        <div class="metric-box">
          <div class="metric-label">Pending</div>
          <div class="metric-val kpi-red" id="m-phishingNormal-pending">—</div>
          <div style="font-size:11px;color:var(--accent);font-family:Arial,sans-serif;margin-top:2px" id="m-phishingNormal-pct-pending">0%</div>
          <div class="progress-bar-wrap">
            <div class="progress-bar-bg"><div class="progress-bar-fill" id="pb-pending-phishingNormal" style="width:0%;background:var(--accent)"></div></div>
          </div>
        </div>
        <div class="metric-box">
          <div class="metric-label">Last Calculated</div>
          <div class="metric-val" style="font-size:12px;padding-top:4px;font-family:Arial,sans-serif;color:var(--text3)" id="m-phishingNormal-calc">Never</div>
        </div>
      </div>"""

new_block = """            <div class="section-title">Phishing — Normal Users</div>
            <div class="section-note" style="margin-top:4px;font-size:12px;color:var(--text3);max-width:720px">
              Non-direct userbase (BSC = no, Band 4+ = no). Total = deduped rows by Emp ID + email.
              Tool match → Completed or Not completed; no tool row → Not found.
            </div>
          </div>
        </div>
        <div class="section-controls">
          <span class="date-label">Assignment Date:</span>
          <input type="date" class="date-input" id="date-phishingNormal" value="">
        </div>
      </div>

      <div class="metrics-row metrics-row-5">
        <div class="metric-box">
          <div class="metric-label">Total (userbase)</div>
          <div class="metric-val kpi-accent" id="m-phishingNormal-total">—</div>
        </div>
        <div class="metric-box">
          <div class="metric-label">Completed</div>
          <div class="metric-val kpi-green" id="m-phishingNormal-completed">—</div>
          <div style="font-size:11px;color:var(--accent);font-family:Arial,sans-serif;margin-top:2px" id="m-phishingNormal-pct">0%</div>
          <div class="progress-bar-wrap">
            <div class="progress-bar-bg"><div class="progress-bar-fill" id="pb-phishingNormal" style="width:0%;background:var(--accent)"></div></div>
          </div>
        </div>
        <div class="metric-box">
          <div class="metric-label">Not completed</div>
          <div class="metric-val kpi-red" id="m-phishingNormal-notCompleted">—</div>
          <div style="font-size:11px;color:var(--accent);font-family:Arial,sans-serif;margin-top:2px" id="m-phishingNormal-pct-notCompleted">0%</div>
          <div class="progress-bar-wrap">
            <div class="progress-bar-bg"><div class="progress-bar-fill" id="pb-notCompleted-phishingNormal" style="width:0%;background:var(--accent)"></div></div>
          </div>
        </div>
        <div class="metric-box">
          <div class="metric-label">Not found</div>
          <div class="metric-val" style="color:var(--text2)" id="m-phishingNormal-notFound">—</div>
          <div style="font-size:11px;color:var(--text3);font-family:Arial,sans-serif;margin-top:2px" id="m-phishingNormal-pct-notFound">0%</div>
          <div class="progress-bar-wrap">
            <div class="progress-bar-bg"><div class="progress-bar-fill" id="pb-notFound-phishingNormal" style="width:0%;background:var(--text3)"></div></div>
          </div>
        </div>
        <div class="metric-box">
          <div class="metric-label">Last Calculated</div>
          <div class="metric-val" style="font-size:12px;padding-top:4px;font-family:Arial,sans-serif;color:var(--text3)" id="m-phishingNormal-calc">Never</div>
        </div>
      </div>"""

if "m-phishingNormal-notCompleted" not in text:
    if old_block not in text:
        raise SystemExit("phishing metrics block not found")
    text = text.replace(old_block, new_block, 1)

p.write_text(text, encoding="utf-8")
print("Patched", p)
