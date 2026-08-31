"""
build.py — assembles the standalone artifact.html from the source pieces.

Run from inside this folder:
    python3 build.py

It reads shell.html (layout + CSS), ui.js (app logic), app.js (small
shared helpers), the bundled xlsx.mini.min.js reader, and the two sample
CSVs, and inlines everything into a single self-contained artifact.html —
no external dependencies except Google Fonts.
"""

import pathlib

base = pathlib.Path(__file__).resolve().parent
shell = (base / "shell.html").read_text(encoding="utf-8")
xlsx_bundle = (base / "vendor/xlsx.mini.min.js").read_text(encoding="utf-8")
app_js = (base / "app.js").read_text(encoding="utf-8")
ui_js = (base / "ui.js").read_text(encoding="utf-8")
employees_csv = (base / "examples/employees.csv").read_text(encoding="utf-8")
projects_csv = (base / "examples/projects.csv").read_text(encoding="utf-8")


def js_template_literal(s: str) -> str:
    # safe because we've verified no backticks / ${ in the source CSVs
    return "`" + s.replace("\\", "\\\\") + "`"


out = shell
out = out.replace("__EMPLOYEES_CSV__", js_template_literal(employees_csv))
out = out.replace("__PROJECTS_CSV__", js_template_literal(projects_csv))
out = out.replace("__XLSX_BUNDLE__", xlsx_bundle)
out = out.replace("__APP_JS__", app_js)
out = out.replace("__UI_JS__", ui_js)

out_path = base / "artifact.html"
out_path.write_text(out, encoding="utf-8")
print("wrote", out_path, len(out), "bytes")
