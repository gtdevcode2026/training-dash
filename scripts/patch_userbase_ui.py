"""Update dashboard HTML for userbase-enriched sections (BSC, Phishing Normal, Band 4+)."""
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
p = ROOT / "dashboard" / "index.html"
text = p.read_text(encoding="utf-8")

replacements = [
    (
        "appends start date_extracted, completion date &amp; status (zone &amp; center sheets in download)",
        "appends training columns to the uploaded userbase; download is the updated userbase only",
    ),
    (
        'onclick="downloadSection(\'bsc\')" disabled>⬇ Download Enriched Userbase</button>',
        'onclick="downloadSection(\'bsc\')" disabled>⬇ Download Updated Userbase</button>',
    ),
    (
        'onclick="downloadSection(\'phishingNormal\')" disabled>⬇ Download Excel (3 sheets)</button>',
        'onclick="downloadSection(\'phishingNormal\')" disabled>⬇ Download Updated Userbase</button>',
    ),
    (
        'onclick="downloadSection(\'band4\')" disabled>⬇ Download Excel (3 sheets)</button>',
        'onclick="downloadSection(\'band4\')" disabled>⬇ Download Updated Userbase</button>',
    ),
]

for old, new in replacements:
    if old in text:
        text = text.replace(old, new, 1)

# Band 4+ metrics row (mirror phishing normal)
old_band4 = """            <div class="section-title">Band 4+ / Senior Management</div>
          </div>
        </div>
        <div class="section-controls">
          <span class="date-label">Assignment Date:</span>
          <input type="date" class="date-input" id="date-band4" value="">
        </div>
      </div>

      <div class="metrics-row">
        <div class="metric-box">
          <div class="metric-label">Total Users</div>
          <div class="metric-val kpi-accent" id="m-band4-total">—</div>
        </div>
        <div class="metric-box">
          <div class="metric-label">Completed</div>
          <div class="metric-val kpi-green" id="m-band4-completed">—</div>
          <div style="font-size:11px;color:var(--accent);font-family:Arial,sans-serif;margin-top:2px" id="m-band4-pct">0%</div>
          <div class="progress-bar-wrap">
            <div class="progress-bar-bg"><div class="progress-bar-fill" id="pb-band4" style="width:0%;background:var(--accent)"></div></div>
          </div>
        </div>
        <div class="metric-box">
          <div class="metric-label">Pending</div>
          <div class="metric-val kpi-red" id="m-band4-pending">—</div>
          <div style="font-size:11px;color:var(--accent);font-family:Arial,sans-serif;margin-top:2px" id="m-band4-pct-pending">0%</div>
          <div class="progress-bar-wrap">
            <div class="progress-bar-bg"><div class="progress-bar-fill" id="pb-pending-band4" style="width:0%;background:var(--accent)"></div></div>
          </div>
        </div>
        <div class="metric-box">
          <div class="metric-label">Last Calculated</div>
          <div class="metric-val" style="font-size:12px;padding-top:4px;font-family:Arial,sans-serif;color:var(--text3)" id="m-band4-calc">Never</div>
        </div>
      </div>"""

new_band4 = """            <div class="section-title">Band 4+ / Senior Management</div>
            <div class="section-note" style="margin-top:4px;font-size:12px;color:var(--text3);max-width:720px">
              Userbase is treated as already filtered. Total = deduped rows by Emp ID + email.
              Tool match → Completed or Not completed; no tool row → Not found. Download = updated userbase.
            </div>
          </div>
        </div>
        <div class="section-controls">
          <span class="date-label">Assignment Date:</span>
          <input type="date" class="date-input" id="date-band4" value="">
        </div>
      </div>

      <div class="metrics-row metrics-row-5">
        <div class="metric-box">
          <div class="metric-label">Total (userbase)</div>
          <div class="metric-val kpi-accent" id="m-band4-total">—</div>
        </div>
        <div class="metric-box">
          <div class="metric-label">Completed</div>
          <div class="metric-val kpi-green" id="m-band4-completed">—</div>
          <div style="font-size:11px;color:var(--accent);font-family:Arial,sans-serif;margin-top:2px" id="m-band4-pct">0%</div>
          <div class="progress-bar-wrap">
            <div class="progress-bar-bg"><div class="progress-bar-fill" id="pb-band4" style="width:0%;background:var(--accent)"></div></div>
          </div>
        </div>
        <div class="metric-box">
          <div class="metric-label">Not completed</div>
          <div class="metric-val kpi-red" id="m-band4-notCompleted">—</div>
          <div style="font-size:11px;color:var(--accent);font-family:Arial,sans-serif;margin-top:2px" id="m-band4-pct-notCompleted">0%</div>
          <div class="progress-bar-wrap">
            <div class="progress-bar-bg"><div class="progress-bar-fill" id="pb-notCompleted-band4" style="width:0%;background:var(--accent)"></div></div>
          </div>
        </div>
        <div class="metric-box">
          <div class="metric-label">Not found</div>
          <div class="metric-val" style="color:var(--text2)" id="m-band4-notFound">—</div>
          <div style="font-size:11px;color:var(--text3);font-family:Arial,sans-serif;margin-top:2px" id="m-band4-pct-notFound">0%</div>
          <div class="progress-bar-wrap">
            <div class="progress-bar-bg"><div class="progress-bar-fill" id="pb-notFound-band4" style="width:0%;background:var(--text3)"></div></div>
          </div>
        </div>
        <div class="metric-box">
          <div class="metric-label">Last Calculated</div>
          <div class="metric-val" style="font-size:12px;padding-top:4px;font-family:Arial,sans-serif;color:var(--text3)" id="m-band4-calc">Never</div>
        </div>
      </div>"""

if "m-band4-notCompleted" not in text and old_band4 in text:
    text = text.replace(old_band4, new_band4, 1)

phishing_note = "Tool match → Completed or Not completed; no tool row → Not found."
if phishing_note in text and "Download = updated userbase" not in text:
    text = text.replace(
        phishing_note,
        phishing_note + " Download = updated userbase (all original columns + appended fields).",
        1,
    )

p.write_text(text, encoding="utf-8")
print("Patched", p)
