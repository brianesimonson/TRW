"""Write a tab-by-tab catalog from audited workbook metadata and reviewed roles."""
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DOCS = ROOT / 'docs'
ROLES = {
    'Welcome': ('User guidance and descriptions of workbook tabs.', 'Guidance only; new application documentation must retain the workflow explanation.'),
    'Current Parameters': ('Template/output paths, current report ID, feature flags, export format, suffix.', 'Partial JSON equivalent; no direct workbook/options adapter.'),
    'FigureDataImport': ('Import file/range/destination specifications, row/column limits, bullet mode.', 'Not converted; CSV filename selection exists, but offsets, clearing, typed imports, workbook data are absent.'),
    'Report Loop List': ('Batch report IDs and current-ID reference.', 'Partial JSON report_ids equivalent; sheet read and per-ID recalculation absent.'),
    'Report Numbers': ('Text IDs/values/formats, chart IDs/names/ranges/alt override, table IDs/layout/format/metadata.', 'Partial generic elements; full schema reader and formatting branches absent.'),
    'ChartRanges': ('Compute active table/chart dimensions from sentinel rows; build range strings.', 'Not converted; no formula or range adapter.'),
    'Table 1': ('Example age table staging, calculated presentation cells, employee total.', 'Partial example recreated directly from CSV; worksheet formulas and styles not converted.'),
    'Table 2': ('Advanced example table data/presentation with dynamic formulas.', 'Not converted; required advanced-demo table acceptance case.'),
    'Table 3': ('Advanced example calculated table used by configured report mappings.', 'Not converted; required advanced-demo table acceptance case.'),
    'Figure 1': ('Age chart data, presentation formulas, named Excel chart, override alt-text narratives.', 'Partial new bar chart; original Excel chart definition/formulas/alt override not ported.'),
    'Figure 2': ('Gender chart data, presentation formulas, named Excel chart.', 'Partial new bar chart; Excel styling/source range/calculation semantics not ported.'),
    'Bullet 1': ('Worksheet-generated narrative and glyph columns, often inserted as a table.', 'Partial real Word bullets from explicit manifest; original table/glyph/formula behavior differs.'),
    'Header and Footer': ('Up to five section-specific header/footer texts and alignments.', 'Not converted; preserving existing headers is not implementing Changeheader.'),
    'x-Header': ('Hidden older header/date helper text.', 'Review candidate; no current direct VB sheet reference; retain until user requirements confirm retirement.'),
    'SmartSections': ('Include/remove decisions, reason statements, bullet-boundary flags, TableX/FigureX renumber mappings.', 'Not converted; dynamic formula dependencies must be included.'),
    'Summary Sheet': ('Placeholder for analysis statistics driving section choices; empty in supplied samples.', 'Not converted; intended configuration role recorded; sample emptiness is not retirement approval.'),
    'Add Docx': ('Additional document base path, section tags, filenames; paths may be formula-generated.', 'Not converted; missing-file fallback and style/relationship-safe insertion needed.'),
    'Date': ('TODAY/DAY/MONTH/YEAR calculations and concatenated date strings.', 'Not converted; needs defined clock and locale semantics.'),
    'CSVIMPORTS': ('Scratch sheet populated by QueryTables for import/mass-table workflows.', 'Not converted as workbook staging; replace with data model preserving parsing/offset semantics.'),
    'GlobalReplace': ('Arbitrary source/destination text replacement pairs.', 'Not converted; generic template-tag replacement is a different contract.'),
    'index': ('Mass-table input file paths, Excel ranges, tags, modes, formats, and extra columns.', 'Not converted; mass-table workflow remains required.'),
    'index_altered': ('Alternate index used with report-specific string substitution.', 'Not converted; must preserve substitution rules and untouched fields.'),
    'index_temp': ('Runtime expanded/copied mass-table index from CommonCore.', 'Not converted; need deterministic per-report in-memory equivalent.'),
    'FFS MR TopX': ('Hidden domain-specific header/bullet helper with IF formulas and imported values.', 'Review candidate; no direct fixed-name VB reference found, but formulas/configured sheet names can still refer to it.'),
    'Sheet1': ('Hidden A–Z lookup/development scratch text.', 'Review candidate; no direct fixed-name VB reference found; not declared unused.'),
    'index_test': ('Hidden test index placeholder in the simple workbooks.', 'Review candidate; advanced demo omits it; no retirement approved.'),
}


def main():
    data = json.loads((DOCS / 'legacy-inventory.json').read_text())
    actual = {t['name'] for w in data['workbooks'] for t in w['tabs']}
    if actual != ROLES.keys():
        raise ValueError('Tab role catalog mismatch: ' + str(actual.symmetric_difference(ROLES)))
    text = ['# Workbook tabs, options, and embedded VBA', '',
            'All three original XLSM files were opened read-only for inspection. Macros were',
            'extracted, not executed. Formula counts include shared formulas expanded by',
            'openpyxl. Reading a cached result is not formula evaluation or conversion.', '',
            'There are 65 sheet occurrences and 26 distinct sheet names. No hidden sheet is',
            'silently excluded. Generic Python report elements do not mean their source',
            'workbook tabs have been converted.', '', '## All tab roles', '',
            '| Tab | Original responsibility | Conversion coverage |', '|---|---|---|']
    for name, (role, coverage) in ROLES.items():
        text.append(f'| `{name}` | {role} | {coverage} |')
    text += ['', '## Per-workbook inventory', '',
             'Formula function names below describe observed formulas, not a supported Python',
             'formula language. Formulas also contain operators, literals, and cross-sheet',
             'references. Every formula expression, its cell address and number format, plus',
             'Excel chart types and series references, are recorded in the JSON.', '']
    for workbook in data['workbooks']:
        text += [f'### {workbook["name"]}', '',
                 f'{len(workbook["tabs"])} tabs; {sum(t["formula_cells"] for t in workbook["tabs"])} formula cells; '
                 f'{len(workbook["vba_modules"])} extracted VBA modules.', '',
                 '| Tab | Visibility | Formula cells | Formula functions observed | Excel charts |',
                 '|---|---|---:|---|---:|']
        for tab in workbook['tabs']:
            functions = ', '.join(tab['formula_functions']) or '—'
            text.append(f'| `{tab["name"]}` | {tab["visibility"]} | {tab["formula_cells"]} | {functions} | {tab["chart_count"]} |')
        text += ['', 'VBA modules (including sheet/document stubs with no routines):', '',
                 '| Module | Source lines | Declared routines |', '|---|---:|---:|']
        for module in workbook['vba_modules']:
            text.append(f'| `{module["file"]}` | {module["line_count"]} | {module["declaration_count"]} |')
        text.append('')
    text += ['## Current Parameters contract', '',
             'Coordinates are taken from the original VB.NET accesses. Unsupported options',
             'must be reported before generation; silently ignoring an On flag is not parity.', '',
             '| Cell | Purpose | Original consumer | Python status |', '|---|---|---|---|',
             '| A2 / A5 | Word template folder / filename | RunReport1 | Partial: JSON template path only |',
             '| A8 | Output destination | RunReport1, chart/HTML temp files | Partial: CLI output path only |',
             '| A11 | Current report ID | LoopReports, workbook formulas | Partial: CLI report ID only |',
             '| B15 | Smart Sections On/Off | Cutsections | Not converted |',
             '| B17 | Headers and Footers On/Off | Changeheader | Not converted |',
             '| B19 | Auto-Import Data Files On/Off | CommonCore → CallFigures | Not converted as workbook option |',
             '| B21 | 508 Compliant Charts On/Off | AssignAlt | Partial metadata; not full compliance or override parity |',
             '| B23 | DOCX / DOC / PDF format | RunReport1 | DOCX only; DOC/PDF held |',
             '| B25 | Table of Contents Update On/Off | RunReport1 → TablesOfContents(1).Update | Held H01; updateFields request is not equivalent |',
             '| B27 | Report title/file suffix | RunReport1 | Not converted |',
             '| B29 | Massive Table Maker On/Off | CommonCore, MassTabMake | Not converted |',
             '| B31 | Mass table substitution identifier | CommonCore | Not converted |',
             '| B33 | User Defined Macros On/Off | Excel.Application.Run("TRW.AddMacros") | Held H02 |',
             '| B35 | Demo Mode On/Off (Word visibility) | Runonestate / LoopReports | No UI equivalent; not a license bypass |',
             '| B37 | Global Replace On/Off | RunReport1 → GlobalReplace | Not converted |',
             '| A36 / A39 | Old kernel location mode / path | ProcessReports workbook VBA | Held H04; kernel executable absent |', '',
             'Report names are derived from template filename before `_Template`, report ID,',
             'suffix, and selected extension. The PDF branch also uses document title metadata',
             'in an intermediate filename. Prototype output_pattern is only a partial replacement.', '',
             '## FigureDataImport and Report Numbers schemas', '',
             '| Sheet / columns | Original contract |', '|---|---|',
             '| FigureDataImport A2 | Base input folder |',
             '| FigureDataImport E / F / G | Enabled-row marker/name, destination sheet, destination starting cell |',
             '| FigureDataImport H / I | Rows and columns to clear/copy, including blank source cells |',
             '| FigureDataImport J / K | Input filename, source starting cell |',
             '| FigureDataImport L | Legacy bullet transform selector 1/2/3; caller currently commented in CallFigures |',
             '| Report Numbers A / B / C | Text tag ID, description, value with cell number format |',
             '| Report Numbers G / H / I / J / K / L | Chart tag ID, description, worksheet, chart object name, data range, alternative-text override column |',
             '| Report Numbers N / O / P / Q | Table tag ID, description, worksheet, range |',
             '| Report Numbers R / S | Fit-to-window / centered alignment flags |',
             '| Report Numbers T / U | Existing Word table shell flag; Variable Row / Fixed Row / Table Convert mode |',
             '| Report Numbers V | Number of repeating header rows |',
             '| Report Numbers W / X | Accessibility title / description |',
             '| Report Numbers Y / Z | Padding in pixels / side selector letters (L,R,T in implementation) |', '',
             'ChartRanges A/B/C/D/F supply element ID, sheet, column count, row count, and',
             'range. Sentinel `xx` formulas exclude unused rows. SmartSections has separate',
             'include/reason fields (A:D,G), table renumbering (J:L), and figure renumbering',
             '(P:R). Add Docx uses D2 as base folder and A/B starting at row 5 as section',
             'tag/filename. The mass index uses A filename, B range, C tag ID, E mode, and',
             'F format string; CommonCore copies A:H, including fields not read by Main.', '',
             'Fixed scan limits in the source (imports rows 4–100, chart rows up to 1000,',
             'several range lookups up to 100/200, headers up to five sections) are part of',
             'the audit. Increasing limits is a deliberate improvement, not an untested parity claim.', '',
             '## Embedded macro routes and uncertainties', '',
             '- Blank Shell and Finished Example each contain ProcessJobOne, ProcessJobMany,',
             '  and ShellMe. They save a handoff workbook and launch the old external kernel.',
             '  They are separate from the VSTO ribbon handlers; neither route is assumed unused.',
             '- Demo Project 1 contains TRW.test, TRW.AddMacros, and TRW.chartidx.',
             '  AddMacros is empty apart from a commented test call; preserving the extension',
             '  contract is still required. chartidx is a workbook diagnostic in addition to',
             '  the VSTO Chart ID button.',
             '- The simple workbooks do not contain a TRW.AddMacros procedure. Their B33',
             '  option is Off; the advanced demo has it On. Enabling it in a workbook without',
             '  that macro is a configuration condition to detect, not silently bypass.',
             '- Old Changeheader2 references a `Header` sheet, absent from all supplied',
             '  workbooks. `x-Header` exists but is a different name. Keep the old function',
             '  flagged for compatibility review rather than pretending those names match.',
             '- QueryTables uses CSVIMPORTS / CSVImports (Excel names are case-insensitive).',
             '  A future Python adapter must normalize matching consistently.', '',
             '## Word templates and fields', '',
             '| Document in supplied package | Tags | Field instruction fragments | External relationships | Embedded media |',
             '|---|---:|---:|---:|---:|']
    for document in data.get('word_documents', []):
        field_count = sum(len(part['field_instructions']) for part in document['field_parts'])
        text.append(f'| `{document["archive_path"].split("/",1)[-1]}` | {len(document["tags"])} | {field_count} | {document["external_relationship_count"]} | {document["media_count"]} |')
    text += ['', 'The example Word template contains a TOC field and four PAGEREF fragments;',
             'it has no external relationships. The generic Word warning seen by the user',
             'therefore does not by itself establish an external dependency. The advanced',
             'template has four external relationships; their presence is recorded, but',
             'no external resources were fetched or field targets executed.', '',
             'Word field-instruction fragments and template tag lists are preserved in',
             'legacy-inventory.json. Layout lives in the templates, but replacement content',
             'can alter pagination. Actual field calculation remains H01.', '']
    (DOCS / 'WORKBOOK_TABS.md').write_text('\n'.join(text))
    print('Documented', len(actual), 'unique tab roles and', sum(len(w['tabs']) for w in data['workbooks']), 'tab occurrences')


if __name__ == '__main__':
    main()
