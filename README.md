# Report Writer — Python engine 0.2

Report Writer now reads the original Excel configuration workbook and fills the
existing Word template without launching Excel or Word. The template still
controls sections, page breaks, styles and layout. The reporting engine is Python;
there is no C++ layer.

This version implements the main reporting paths, with explicit compatibility
limits. Start with the [conversion audit](docs/CONVERSION_AUDIT.md),
[function ledger](docs/FUNCTION_LEDGER.md), [workbook tabs](docs/WORKBOOK_TABS.md),
and [validation results](docs/VALIDATION.md). The original VSTO application remains
available while the replacement is being accepted.

## Try it on Windows

1. Download the **python-engine-prototype** branch using GitHub **Code → Local →
   Download ZIP**, then extract it. Open the folder containing `pyproject.toml`
   and `RunReportWriter.cmd`.
2. Double-click **RunReportWriter.cmd**. It uses your installed `python` command,
   creates the environment and installs dependencies, then opens a small desktop
   interface. Python 3.11+ is required. The launcher/UI still needs Windows testing.
3. Browse to `examples\legacy\advanced-corrected\Demo Project 1.xlsm` and choose
   an output folder. Leave the report ID blank to generate A and B.
4. Click **Generate reports**, then **Open output folder**. Every run gets a fresh
   folder. Open the generated Word files or choose PDF output.

Install [LibreOffice](https://www.libreoffice.org/download/download-libreoffice/)
for automatic table-of-contents/page-number calculation, PDF output or legacy
DOC output. It runs headlessly during export. DOCX assembly alone needs only
Python. **Draft only** skips a requested contents refresh and reports that page
numbers are unfinished; it does not make a final report.

If you prefer the PowerShell path, run these commands from the extracted folder
that contains `pyproject.toml`:

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -e .
.\.venv\Scripts\python.exe -m trw.cli "examples\legacy\advanced-corrected\Demo Project 1.xlsm" --output "output\my-test"
explorer .\output\my-test
```

Use a new output folder for a repeat run; existing files are protected. A failed
batch can leave earlier completed reports. The advanced workbook's VBA flag is
On: the results explicitly warn that boutique queries were bypassed, as requested.

## What is implemented

- Original workbook mappings/options, batch IDs, CSV/XLSX offsets and clearing,
  index substitution, fresh formulas and formatted values.
- SmartSections deletion/reasons, text, bar charts, fixed/variable table shells,
  native merged tables and mass-table modes 1/2/3/scalars.
- Additional DOCX insertion, caption numbering, five section headers/footers,
  missing-value and global replacements, chart/table accessibility metadata.
- DOCX, PDF and legacy DOC export. Heading-based contents tables are rebuilt from
  actual rendered headings and page labels until pagination stabilizes. Generated
  files do not request a manual field refresh on opening.
- Single/batch/import-only CLI, chart-name listing, progress events and a basic Tk
  desktop front end. Import-only saves a separate prepared workbook with caches.

All 641 formula cells in the original three workbooks evaluate without cached
Excel values. This is a defined formula interpreter, not all of Excel: unknown
functions, unsupported chart types/TOC switches and malformed mappings fail.
Charts use Matplotlib PNGs and are not pixel-identical Excel exports. Native DOCX
tables replace the old Excel HTML/clipboard branch; conditional formatting and
complex Excel HTML presentation remain gaps.

LibreOffice determines final pagination. Microsoft Word can paginate differently,
particularly with different fonts. Custom TOCs, arbitrary Word fields, complex
SmartSections spanning table cells, distinct inserted-document headers/footers,
and a retained Excel ribbon bridge remain partial. The Windows UI/launcher has
not been exercised here. See the ledger rather than treating a demo pass as full
application parity. Metadata is not a certification of Section 508 compliance.

Access identity/license/logging and associated encryption, killing stray Word
processes, and user VBA execution are excluded or deferred by user direction.
Future authentication is separate work. No Access databases, signing keys,
compiled add-in binaries or obfuscation maps are included in these examples.

## Examples

`examples/legacy/advanced` and `finished` contain unchanged reporting inputs from
the supplied package. `blank` contains its original workbook for formula audit.

The **advanced-corrected** copy changes the female narrative from “males” to
“females,” points Table03 at the merged Table3 data, and supplies a matching native
Word table placeholder. Its gender values link to the imported Figure2 values so
A/B reports stay current. TOC refresh is On in this copy to demonstrate automatic
finalization. Original files remain unchanged. See [example details](examples/legacy/README.md).

The older JSON/CSV example remains available at `examples/hello_world/report.json`.
It is draft assembly with explicit mappings, not the workbook adapter. New
validated sample reports are in [sample-output/v0.2](sample-output/v0.2).

## Development and CLI

```sh
cd /workspace/TRW
python3 -m venv .venv
.venv/bin/python -m pip install -e .
.venv/bin/python -m pip check
.venv/bin/python -m unittest discover -s tests -v
.venv/bin/python -m trw.cli 'examples/legacy/advanced-corrected/Demo Project 1.xlsm' --output output/new-run
```

Add `--report A` for one report (repeat for selected IDs), `--format pdf` or `doc`,
`--as-of 2026-10-13` for deterministic dates, `--project-dir PATH` to relocate a
project, or `--soffice PATH` for a nonstandard LibreOffice installation.
`--update-toc`/`--no-update-toc` override B25. `--import-only` writes prepared
workbooks; `--list-charts` lists actual Excel drawing names.

Tests requiring LibreOffice are reported as skipped when it is absent. No
persistent server is required. Regenerate documentation with
`python tools/write_conversion_docs.py` and `python tools/write_workbook_docs.py`;
reviewed coverage lives in `docs/conversion-updates.json` and
`docs/workbook-coverage.json`. The raw legacy inventory remains source provenance.
