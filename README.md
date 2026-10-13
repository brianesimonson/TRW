# Report Writer Python engine — first implementation

Generates editable Word documents from an existing Word template, without
launching Word or Excel. The original VSTO application and uploaded archives have
not been modified. This is a working first engine, not a complete replacement
for every feature of the VB.NET application.

## Run

Python 3.11 or newer:

```sh
cd /workspace/TRW
python -m venv .venv
.venv/bin/python -m pip install -e .
.venv/bin/python -m trw.cli examples/hello_world/report.json --output output/reports
.venv/bin/python -m unittest discover -s tests -v
```

On Windows, use `.venv\Scripts\python.exe` in place of `.venv/bin/python`.
After installation, the equivalent CLI is `trw report.json --output reports`.
Use `--report A` for one report, or repeat `--report` for selected reports.
Existing output files are protected from overwriting; choose a new output folder
for another run. A batch can leave earlier completed reports if a later one fails.

## Existing template, new engine

The example contains the original supplied `Example Project_Template.docx` and
six CSV inputs. `report.json` explicitly expresses the example's mappings and
calculations; it is not an automatic conversion of the `.xlsm` workbook.

The Word document remains the layout authority. Existing sections, margins,
page breaks, styles, headers, footers, and unrelated content are retained.
Text tags can span formatting runs and occur in tables, headers, or footers.
Text replacements retain surrounding run formatting and use the tag's first
run formatting for the replacement. Table, bullet, and chart tags must occupy
their own paragraph. Inserted tables use explicit column widths and repeating
header rows. Charts are PNG images with data-derived alternative text.
Bullets use actual Word numbering rather than font-specific bullet characters.

The reusable engine is `trw.engine.render(template, destination, elements)`.
The manifest/CSV adapter is separate. Another adapter can accept resolved data
and formatting from an Excel ribbon without changing the document engine.
Manifests can also omit `sources` and supply resolved `value`, `rows`,
`labels`/`values`, and `items` directly for text, table, chart, and bullet elements.

The Company B example intentionally selects its own CSV inputs (3 employees).
The uploaded simple example's fixed Company A filenames produced 30 employees
even in its B report. The advanced uploaded workbook already selects inputs
by report ID. This behavior is explicit in the new manifest, not a change to
the uploaded workbook.

## Implemented and checked

- Existing `@#$_...` tags; repeated tags and tags split across runs.
- CSV input selection per report, numeric totals, text, tables, charts, bullets.
- Individual and batch DOCX generation; chart and table accessibility metadata.
- Failure on unknown/unresolved tags, malformed table rows, or output collisions.
- Automated checks of both supplied example datasets, document sections/page
  breaks, header replacements, run formatting, chart count/alt text, and ZIP integrity.

## Remaining work

This engine does **not** calculate Excel formulas, execute workbook VBA, import
workbook mappings automatically, reproduce Excel charts exactly, or implement
SmartSections, document merging, mass-table rules, caption renumbering, licensing,
or the ribbon integration. The example calculations were expressed directly in
the manifest adapter. Inserted tables/charts have new explicit styling; their
appearance has not been proven identical to the Interop output.

No PDF renderer or Word pagination engine is included. The template's TOC fields
are preserved and marked for refresh on opening in Word; their cached page
numbers remain uncalculated here. Growing content may change final pagination.
Accessibility metadata alone is not a claim of Section 508 compliance.

Both example DOCX files were also successfully opened and exported by headless
LibreOffice as a compatibility check. The four-page Company A output was visually
inspected. This is not validation against Microsoft Word's exact rendering; the
existing table-of-contents page cache still needs refresh. LibreOffice is not a
runtime dependency of the report engine.

Next: connect the Excel ribbon to resolved report data, or implement a workbook
adapter with a deliberate strategy for formulas and chart styling; then port
the remaining VB.NET feature branches and validate against the advanced demo.
Keep the existing VSTO application available until these are verified.
