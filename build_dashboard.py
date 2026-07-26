"""
Builds the standalone dashboard.html from web/dashboard_source.html by
inlining the vendored JS libraries (web/vendor/*.js), so the final file
can be opened directly in a browser with no server and no internet access.

Re-run after editing web/dashboard_source.html:
    python build_dashboard.py
"""
import os
import re

HERE = os.path.dirname(os.path.abspath(__file__))
SOURCE = os.path.join(HERE, "web", "dashboard_source.html")
VENDOR_DIR = os.path.join(HERE, "web", "vendor")
OUTPUT = os.path.join(HERE, "dashboard.html")


def inline_script(html, src_name):
    path = os.path.join(VENDOR_DIR, src_name)
    with open(path, "r", encoding="utf-8") as f:
        content = f.read()
    pattern = re.compile(r'<script src="vendor/' + re.escape(src_name) + r'"></script>')
    replacement = "<script>\n" + content + "\n</script>"
    new_html, count = pattern.subn(lambda m: replacement, html)
    if count != 1:
        raise SystemExit(f"Expected exactly 1 occurrence of vendor/{src_name} tag, found {count}")
    return new_html


def main():
    with open(SOURCE, "r", encoding="utf-8") as f:
        html = f.read()
    html = inline_script(html, "xlsx.full.min.js")
    html = inline_script(html, "chart.umd.min.js")
    with open(OUTPUT, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"Wrote {OUTPUT} ({os.path.getsize(OUTPUT):,} bytes)")


if __name__ == "__main__":
    main()
