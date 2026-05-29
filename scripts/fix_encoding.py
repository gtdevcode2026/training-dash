"""Fix mojibake from PowerShell UTF-8 mishandling."""
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]

REPLACEMENTS = [
    (chr(0xE2) + chr(0x20AC) + chr(0x201D), chr(0x2014)),
    (chr(0xE2) + chr(0x20AC) + chr(0x201C), chr(0x2014)),
    (chr(0xE2) + chr(0x20AC) + chr(0xA6), chr(0x2026)),
    (chr(0xE2) + chr(0x2013) + chr(0xB6), chr(0x25B6)),
    (chr(0xE2) + chr(0xAC) + chr(0x2021), chr(0x2B07)),
    (chr(0xE2) + chr(0x161) + chr(0xA0), chr(0x26A0)),
    (chr(0xE2) + chr(0x153) + chr(0x2022), chr(0x2715)),
    (chr(0xE2) + chr(0x161) + chr(0xA1), chr(0x26A1)),
    (chr(0xE2) + chr(0x9C) + chr(0x93), chr(0x2713)),
    (chr(0xE2) + chr(0x9C) + chr(0x95), chr(0x2715)),
    (chr(0xE2) + chr(0x84) + chr(0xB9), chr(0x2139)),
    (chr(0xE2) + chr(0x9A) + chr(0xA0), chr(0x26A0)),
    (chr(0xE2) + chr(0xAD) + chr(0x90), chr(0x2B50)),
    (chr(0xE2) + chr(0x153) + chr(0x201C), chr(0x2713)),
    (chr(0xE2) + chr(0x89) + chr(0xA5), chr(0x2265)),
    (chr(0xE2) + chr(0x86) + chr(0x92), chr(0x2192)),
]


def fix_text(text: str) -> str:
    for bad, good in REPLACEMENTS:
        text = text.replace(bad, good)
    return text


def main():
    targets = [
        ROOT / "dashboard" / "index.html",
        ROOT / "shared" / "javascript" / "app.js",
    ]
    for path in targets:
        if not path.exists():
            continue
        raw = path.read_text(encoding="utf-8")
        fixed = fix_text(raw)
        if fixed != raw:
            path.write_text(fixed, encoding="utf-8", newline="\n")
            print("fixed:", path.relative_to(ROOT))
        else:
            print("unchanged:", path.relative_to(ROOT))
        i = fixed.find("<title>")
        print("  title:", fixed[i : fixed.find("</title>", i) + 8])


if __name__ == "__main__":
    main()
