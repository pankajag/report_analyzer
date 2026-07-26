"""
Builds the standalone dashboard HTML files from their web/*_source.html
templates by inlining the vendored JS libraries (web/vendor/*.js), so the
final files can be opened directly in a browser with no server and no
internet access.

Re-run after editing any web/*_source.html file:
    python build_dashboard.py
"""
import os
import re

HERE = os.path.dirname(os.path.abspath(__file__))
VENDOR_DIR = os.path.join(HERE, "web", "vendor")

BUILDS = [
    ("web/dashboard_source.html", "dashboard.html"),
    ("web/service_usage_source.html", "service-usage.html"),
]


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


def build(source_rel, output_rel):
    source = os.path.join(HERE, source_rel)
    output = os.path.join(HERE, output_rel)
    with open(source, "r", encoding="utf-8") as f:
        html = f.read()
    html = inline_script(html, "xlsx.full.min.js")
    html = inline_script(html, "chart.umd.min.js")
    with open(output, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"Wrote {output} ({os.path.getsize(output):,} bytes)")


def main():
    for source_rel, output_rel in BUILDS:
        build(source_rel, output_rel)


if __name__ == "__main__":
    main()
