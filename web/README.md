# Web demo — source

This is the source for the zero-install, client-side version of the Data
Quality Profiler — the same detectors and scoring rules as the Python
engine in `src/`, ported to dependency-free vanilla JavaScript so it can
run entirely in the browser with no server and no install.

**Live version:** https://omaralthunaiyan-cloud.github.io/data-quality-profiler/

## Files

- `shell.html` — page layout, all CSS (light/dark theme, the whole visual
  design), and placeholders (`__UI_JS__`, `__XLSX_BUNDLE__`, etc.) that
  `build.py` fills in.
- `ui.js` — all app logic: file/sheet loading, running the profiler,
  rendering the report, CSV/PDF export.
- `app.js` — small shared helpers.
- `vendor/xlsx.mini.min.js` — the [SheetJS](https://sheetjs.com) library,
  used only to parse `.xlsx`/`.xls` files client-side (MIT-licensed, see
  `vendor/LICENSE`).
- `examples/employees.csv`, `examples/projects.csv` — the two sample
  datasets, embedded directly into the built page so "try a sample" works
  with zero network requests.
- `artifact.html` — the pre-built, single-file output. This is what
  actually gets published as the live demo — open it directly in any
  browser and it works standalone.
- `build.py` — assembles all of the above into `artifact.html`.

## Rebuilding

```bash
python3 build.py
```

This overwrites `artifact.html` with a fresh build from the source files.
Nothing else is needed — no npm install, no build tooling.

## Notes

- Everything runs client-side: no file you open is ever uploaded anywhere.
- The only external network calls `artifact.html` makes are to Google
  Fonts (`fonts.googleapis.com` / `fonts.gstatic.com`) for the page's
  typefaces.
- This source isn't wired into the Python app's test suite or CI — it's a
  separate, self-contained implementation kept here for reference and so
  the two engines' scoring can be spot-checked against each other (see
  "Two implementations, one engine" in the main README).
