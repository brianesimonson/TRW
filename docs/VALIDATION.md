# Validation evidence — engine 0.2.0

Executed in the cloud on 2026-10-13 with Python 3.12, pinned project dependencies,
and the preinstalled LibreOfficeDev 26.8.0.0.alpha0 CLI. This renderer build is
not a Windows/stable-LibreOffice acceptance test. The user previously ran the old
prototype on Python 3.11.9; the new desktop UI/launcher is still untested there.

```sh
MPLCONFIGDIR=/tmp/trw-matplotlib XDG_CACHE_HOME=/tmp/trw-cache .venv/bin/python -m unittest discover -s tests -v
.venv/bin/python -m pip check
```

**34 tests pass**, with zero skips/failures in this environment. The three renderer
integration tests explicitly skip if LibreOffice is absent, so a pass elsewhere
must distinguish those skipped checks. Editable installation and pip check pass.

| Checks | Evidence / practical limit |
|---|---|
| Original formulas | All 641 cells evaluated from formula text across the three original workbooks; no data-only/cache fallback. Fresh A/B import totals, named ranges and dynamic sections checked. This is not proof of arbitrary Excel semantics. |
| Imports | Quoted/comma/newline CSV, source/destination offsets, typed numbers, blank clearing, XLSX formula evaluation; A/B filename changes and original-file immutability. |
| CommonCore | Index substitution for A/B and copying configuration; deferred VBA warning plus optional Python extension. No VBA executed. |
| Mass tables | Fixed modes 1/3, variable mode 2 stops at first blank, scalar mode and column formats; parser handles format10.0% correctly. |
| Document assembly | Split/case-insensitive/prefix tags, retained section/page-break properties, tables, headers, unresolved tags and collision failures. |
| SmartSections | A keeps gender section, B removes it and writes reason; deletion across an intervening body table retains surrounding text. Complex cross-cell ranges remain unsupported. |
| Native tables | Horizontal/vertical merged cells, font/fill/header/padding XML; corrected Table3 A counts 17/13, B counts 0/3; fixed shell fails rather than truncating. |
| Bullets | All three legacy transforms, header/count behavior and empty fallbacks. Demo Wingdings-l glyph becomes Unicode bullet for portable rendering. |
| Headers/replacements | Five sections' headers/footers and alignment; caption/global source mappings and all four missing sentinels, sequential replacements and Off behavior. |
| Added documents | Advanced A insertion, missing-file message, inserted image relationships and bold text; master-header policy and body anchors remain limitations. |
| Chart layout | Original source charts are scaled proportionally to page/cell width when too wide. Regression checks width/aspect ratio. Exact Excel styling and other chart types are not covered. |
| TOC | Change heading text/add substantial content, recompute layout, compare cached TOC/PAGEREF labels against final PDF headings/pages; updateFields false. Unsupported switches fail before publishing. |
| Exports/files | Final PDF outline/page counts match TOC; legacy DOC compound-file signature verified; no overwrite, bounded missing-file wait, scoped resources. Legacy DOC re-imported pagination/Windows locks not certified. |
| Import-only | Saves separate XLSM with recomputed formula caches and unchanged vbaProject.bin bytes; no workbook VBA execution or broad extension-preservation claim. |

## End-to-end examples

The original advanced and finished workbooks were run for both A/B. Original
advanced B25 Off was explicitly overridden On to exercise finalization. Finished
B25 On was honored; fixed Company A inputs for both IDs remain as supplied.
The corrected copy was run using its B25 On setting:

```sh
.venv/bin/python -m trw.cli 'examples/legacy/advanced-corrected/Demo Project 1.xlsm' --output output/v02-layout-fixed --as-of 2026-10-13
```

Both corrected DOCX outputs were exported again to PDF. The resulting
[samples](../sample-output/v0.2) have:

| Report | Employees | DOCX tables / chart images | PDF pages | TOC conclusion label |
|---|---:|---:|---:|---:|
| A | 30 | 6 / 2 | 6 | 5 |
| B | 3 | 5 / 1 | 5 | 4 |

The table count includes the template cover table. B's gender narrative/chart
section is removed, while its gender heading outside the original markers remains
and appears in the TOC. Missing-value demo lines and the extra final template
page are preserved; legacy sequential substring replacement can yield N/A.0.
No unresolved report tags or section delimiters remain. Advanced VBA On emits a
warning and results report partial status because boutique queries are deferred.
The Word file itself is generated and its supported TOC is finalized.

Rendered contents, merged table and chart/body pages were visually inspected.
An overwide-chart clipping problem was found, corrected and checked in the final
PDF. Inspection is not a golden comparison against Microsoft Word/Excel.
[sample validation.json](../sample-output/v0.2/validation.json) records final
page labels, titles, table and image counts. Original fixture bytes and the raw
inventory were preserved; only the corrected fixture has intentional changes.

## Unrun acceptance work

- New Windows Tk/launcher installation, control interactions and locked-file behavior.
- Microsoft Word opening/printing and stable LibreOffice Windows font/layout checks.
- Pixel/layout comparisons to Interop for broader chart types, conditional formatting,
  complex shells, first/even-page headers and inserted-document section policies.
- Original Access/VBA/KillWord are excluded or deferred, not passing parity checks.
- Signed standalone executable packaging, retained ribbon bridge and arbitrary Word
  fields/custom TOCs remain future work; all are acknowledged in the ledger.
