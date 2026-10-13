"""Build the reviewed per-routine conversion ledger from the legacy inventory.

Descriptions and statuses are deliberately explicit here, not inferred from a
passing demo. An unknown source routine fails generation and needs review.
"""
import ast
from collections import Counter
import csv
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DOCS = ROOT / 'docs'

DESCRIPTIONS = {
    'AccessX.vb': {
        'Test100': 'Empty commented-out identity test stub.',
        'CheckIdentityStrict': 'Read UserAccess in Access; distinguish permitted, denied, and unknown Windows users.',
        'CheckIdentity': 'Check UserAccess; automatically add and permit unknown users in non-strict mode.',
        'CheckStrictness': 'Read StrictAccess policy from the Access database.',
        'LogTransaction': 'Write user, workbook path, timestamp, and single/batch report count to UserLogs.',
        'RequestReport': 'Gate report execution on license and identity, then log the request; return status 1–4.',
        'VerifyLicense': 'Read encrypted license records from IHTNGU, decrypt using current month/year, and validate marker.',
        'Testrun': 'Manual test entry to request one licensed report.',
        'SetConnstr': 'Resolve TRWlogs.accdb beside the active workbook and return ACE connection string; both active candidate paths are identical, older alternate path is commented.',
    },
    'EncryptX.vb': {
        'Encrypt': 'Encrypt Unicode text with AES, derived key/IV, and Base64 output.',
        'Decrypt': 'Decrypt Base64 license text using date-dependent key material; decode Unicode text.',
    },
    'ModuleConnection.vb': {
        'openConnection': 'Connect to MySQL using embedded workstation connection settings; display connection result.',
        'TestConnection': 'Call openConnection and enumerate database schemas; diagnostic/test workflow.',
    },
    'ChangeChart.vb': {
        'CHrt': 'Read Report Numbers chart rows and reset named Excel chart source ranges; suppress individual errors.',
    },
    'GrabFigureData.vb': {
        'Grabdata': 'Alternate importer opens a separate Excel instance and attempts formatted-cell copy; falls back on selected errors.',
        'GrabdataExcelOld': 'Older separate-Excel importer using FigureDataImport offsets and copy dimensions.',
        'GrabdataExcel': 'Intermediate external-workbook importer with per-cell COM release and staging copy.',
        'GrabdataExcel4': 'Current importer: resolve filename, query into CSVIMPORTS, clear destination rectangle, copy offset rows/columns, clean staging.',
        'Makebullets': 'Select one of three domain-specific bullet transformations from FigureDataImport column L.',
        'CallFigures': 'Import configured rows 4–100, run COM cleanup, and refresh chart ranges via ChangeChart.CHrt.',
        'ColNum': 'Map A–Z to 1–26; return legacy sentinel 1000 for other inputs.',
        'ColLetter': 'Map 1–26 to A–Z; return AA for other inputs.',
        'HandleX': 'Diagnostic message-box callback; former commented library hook.',
        'JustImportMe': 'Old one-time-import entry; creates Excel instance, calls GrabdataExcelOld, releases COM objects.',
        'JustImportMe3': 'Current ribbon one-time-import entry: imports configured rows using GrabdataExcel4, cleans COM, shows completion.',
        'JustImportMe2': 'Intermediate one-time-import entry calling GrabdataExcel with COM cleanup.',
        'NAR': 'Repeatedly release a COM object; swallow cleanup errors.',
        'ImportCSV': 'Import CSV using Excel QueryTables and legacy parse settings; remove connections/query tables afterward.',
        'ImportCSVsmart': 'Clear staging, derive basename from full path, query CSV into CSVIMPORTS, remove connections/query tables.',
    },
    'MassTabMake.vb': {
        'Main': 'Iterate index_temp rows until blank column-C tag; invoke SearchAndReplace for each tag.',
        'SearchAndReplace': 'Import indexed CSV, parse per-column formats, fill scalar/fixed table modes 1/3 or expand rows in mode 2.',
        'Xformat': 'Extract formatN segments from a concatenated per-column format specification.',
    },
    'ReportCycles.vb': {
        'MyForm': 'Lazily construct and expose the topmost Processing progress form.',
        'New': 'Construct FileNotExist/MyErrorMessage exceptions with message and optional inner exception.',
        'RunonestateInLoop': 'Run CommonCore followed by the main report generator for the current batch item.',
        'Runonestate': 'Capture active sheet, import/prepare workbook, launch Word, generate one report, quit Word, restore Excel UI.',
        'LoopReports': 'Launch Word once; iterate Report Loop List IDs, change Current Parameters A11, prepare and render each; restore UI.',
        'CommonCore': 'Build substituted index_temp for mass tables, optionally auto-import, and optionally invoke TRW.AddMacros.',
        'CountLoopReports': 'Count contiguous report IDs starting at Report Loop List A2.',
        'SetThisWorkbook': 'Bind shared xlapp workbook from the VSTO host active workbook.',
        'KillWord': 'Ask to close all existing winword processes; kill them if the user agrees.',
    },
    'ReportWriter.vb': {
        'CloseClipboard': 'Win32 declaration closing the system clipboard.',
        'EmptyClipboard': 'Win32 declaration clearing the system clipboard.',
        'OpenClipboard': 'Win32 declaration opening the system clipboard.',
        'SearchAndReplace': 'Find every text tag, apply Report Numbers column C value and Excel number format, split long text writes.',
        'ChartFill': 'Find each chart tag and insert the exported chart image at Word selection.',
        'TableFill': 'Find a table tag; insert published Excel HTML via CopyX; retry table-copy errors.',
        'Waitsec': 'Wait five seconds through Excel Application.Wait.',
        'ClearClipboard01': 'Clear Excel clipboard via the Clipboard command bar.',
        'ClearClipboard02': 'Disable copying objects with Excel cells.',
        'ClearClipboardEl2000': 'Clear Excel clipboard with legacy command ID 3634.',
        'ClearClipboard': 'Open, clear, and close the Windows clipboard using Win32 APIs.',
        'ClearClipboard100': 'Alternate Windows clipboard-clear implementation.',
        'RunReport1': 'Open/copy template, cut sections, mass tables, populate elements, insert docs, renumber, change headers, normalize missing values, update TOC, export DOCX/DOC/PDF.',
        'Main_after_word_tool': 'Populate text/charts/tables; choose shell-table, direct-cell table, or Excel HTML branch; apply size, alignment, headers, alt metadata, spacing.',
        'CopyTableAgain': 'Clipboard copy of Excel range then retry TableFill.',
        'Changeheader': 'Write configured headers/footers and alignment for up to five Word sections.',
        'Changeheader2': 'Older header-only routine driven by Header tab; adjusts Word view and header selection.',
        'Cutsections': 'Use SmartSections flags to remove Xsection…S/E regions or remove delimiters; insert reason text and bullet-adjacent edits.',
        'RenameTables': 'Replace TableX placeholder labels with numbers from SmartSections J:L.',
        'RenameFigures': 'Replace FigureX placeholder labels with numbers from SmartSections P:R.',
        'DeleteSection': 'Delete content between start/end markers with a Word wildcard range search.',
        'DeleteTags': 'Delete section delimiter markers while retaining their content.',
        'DoubleQuotes': 'Legacy caret-to-quote text conversion; includes an unfinished triple-caret search loop.',
        'GlobalReplace': 'Apply contiguous GlobalReplace sheet source/destination rows.',
        'ReplaceMe': 'Repeated case-insensitive literal Word find/replace of arbitrary text.',
        'MissingToNA': 'Replace four specific formatted missing-value strings with N/A.',
        'AddSections': 'Resolve additional DOCX files from Add Docx rows; insert at tags or emit missing-file text.',
        'SectionFill': 'Replace section tag with inserted document content, or a missing-file message; retry on insertion errors.',
        'CopySectionAgain': 'Select/copy additional document content and retry SectionFill.',
        'AssignAlt': 'Build chart descriptions from override column or formatted 2/3/many-column ranges; assign to images lacking alt text.',
        'UltTableFill': 'Fill existing Word table shell from formatted Excel cells; support Variable Row insertion or fixed rows.',
        'UltTable': 'Find direct-table tag and invoke UltTableCopy.',
        'UltTableShell': 'Find existing-table-shell tag and invoke UltTableFill.',
        'UltTableCopy': 'Create Word table from Excel range; reproduce merged cells, borders, alignment, fonts, colors, formatted values, repeating headers, and metadata.',
        'CopyX': 'Insert Excel-exported HTML, adjust padding/cleanup artifacts, manipulate selection around inserted table.',
        'ExportPic': 'Export named Excel chart; capture chart width/height; uses GIF filter with a .jpg filename.',
        'InsertPic': 'Wait for image file then embed it with captured Excel chart dimensions.',
        'FindLines': 'Strip Excel conditional HTML markup and apply selected left/right/top padding substitutions.',
        'FindLines2': 'Strip Excel conditional HTML markup without padding substitutions.',
        'Delextrarows': 'Selection-based cleanup below an inserted table; delete selection on table-selection failure.',
        'DelextraTOProws': 'Selection-based cleanup above an inserted table; backspace recovery on failure.',
        'XLtoWRDHorAlign': 'Map Excel horizontal alignment constants to Word paragraph alignment.',
        'XLtoWRDVerAlign': 'Map Excel vertical alignment constants to Word cell alignment (includes legacy value requiring validation).',
        'XLtoWRDBrdrThick': 'Map Excel border weights to Word line-width constants.',
        'XLtoWRDLineStyle': 'Map Excel line-style constants to Word border-style constants.',
    },
    'RWRibbon.vb': {
        'Ribbon1_Load': 'Empty ribbon load event.',
        'Button2_Click': 'Run All Reports button: progress UI, license gate, Word-close check, batch generation, error dialogs, cleanup.',
        'Button1_Click_1': 'Run One Report button: progress UI, license gate, Word-close check, single generation, error dialogs, cleanup.',
        'Button3_Click': 'One Time Import button: progress UI, JustImportMe3, workbook/error handling.',
        'ChartID_Click': 'Chart ID button: read selected chart name and strip worksheet name; show diagnostic/error message.',
    },
    'ProcessingFrm.vb': {
        'Processing_Load': 'Empty processing-form load event.',
        'GroupBox1_Enter': 'Empty group-box event routine; no Handles binding in declaration.',
    },
    'ThisAddIn.vb': {
        'ThisAddIn_Startup': 'Empty VSTO startup event.',
        'ThisAddIn_Shutdown': 'Force two garbage-collection/finalizer cycles at VSTO shutdown.',
    },
    'WaitProcedure.vb': {
        'WaitForClose': 'Poll file lock with two-second sleeps; no timeout.',
        'WaitForOpen': 'Require existing file, then poll file lock every 200ms; no timeout.',
        'WaitForExist': 'Retry existence ten times with 200ms sleeps; throw if still missing.',
        'FileInUse': 'Try exclusive read/write file open; any exception is treated as in-use.',
        'FileExists': 'Return System.IO.File.Exists result.',
    },
    'bullets.vb': {
        'Bulletthis': 'Turn imported rows into header/bullet pairs, append parenthetical counts, apply fonts, and emit no-common-causes fallback.',
        'Bulletthis2': 'Alternate header/bullet transform without counts; emit HCBS Medical Review no-errors fallback.',
        'Bulletthis3': 'Alternate header/bullet transform without counts; emit HCBS Data Processing no-errors fallback.',
    },
}

HELPERS = {
    'Startcolumn1': 'Parse the starting column from a dollar-qualified Excel address.',
    'Startrow1': 'Parse starting row from an Excel range; legacy single-cell branch returns zero.',
    'Totalcolumnmoves': 'Return end-column minus start-column (a movement count, not inclusive width).',
    'Totalrowmoves': 'Return end-row minus start-row (a movement count, not inclusive height).',
    'Test': 'Manual Countcolumns("AZ") message-box test.',
    'Countcolumns': 'Convert one/two-letter column names using Firstchar plus Secondchar.',
    'Firstchar': 'Map first letter to a multiple-of-26 column offset; percent sentinel means zero.',
    'Secondchar': 'Map second letter to 1–26.',
    'CellToRange': 'Convert numeric column index to a one/two-letter address through Convertme.',
    'Convertme': 'Map 0–26 to empty/A–Z; return A outside that range.',
}

GENERATED = {
    'New': 'Construct generated Office host/ribbon/global support object.',
    'Dispose': 'Dispose generated form/ribbon components and call base cleanup.',
    'InitializeComponent': 'Construct designer-defined ribbon controls or progress form controls.',
    'Ribbon1': 'Expose generated ribbon collection accessor.',
    'Initialize': 'Initialize VSTO host application, controls, cached data, and host item.',
    'FinishInitialization': 'Complete VSTO initialization and invoke startup.',
    'InitializeDataBindings': 'Begin/end initialization for generated data bindings.',
    'InitializeCachedData': 'Ask VSTO data host to initialize cached data when needed.',
    'InitializeData': 'Generated empty application-data hook.',
    'BindToData': 'Generated empty data-binding hook.',
    'StartCaching': 'Delegate named-member caching to VSTO data host.',
    'StopCaching': 'Delegate named-member cache stop to VSTO data host.',
    'IsCached': 'Query whether a named member is cached by the VSTO data host.',
    'BeginInitialization': 'Begin generated host/control initialization.',
    'EndInitialization': 'End generated host/control initialization.',
    'InitializeControls': 'Generated empty controls-initialization hook.',
    'InitializeComponents': 'Generated empty components-initialization hook.',
    'NeedsFill': 'Query whether the VSTO cached member needs population.',
    'OnShutdown': 'Dispose VSTO custom task panes and call base shutdown.',
    'ThisAddIn': 'Expose generated shared add-in instance; prevent duplicate assignment.',
    'Factory': 'Expose shared Office application factory; prevent duplicate assignment.',
    'Ribbons': 'Lazily construct generated ribbon collection.',
    'ResourceManager': 'Lazily create generated embedded-resource manager.',
    'Culture': 'Get/set resource culture override.',
    'AutoSaveSettings': 'Persist generated application settings on shutdown when requested.',
    'Default': 'Expose thread-safe default generated settings instance.',
    'Settings': 'Expose generated My.Settings accessor.',
}

# Tuple: available replacement, exact remaining gap. These are all partial.
PARTIAL = {
    ('ReportWriter.vb', 'SearchAndReplace'): ('engine.replace_text / render', 'No workbook C-column adapter, Excel format semantics, or legacy case-insensitive matching; new tag grammar is narrower.'),
    ('ReportWriter.vb', 'ChartFill'): ('engine.render chart branch', 'Standalone paragraph tags only; Excel chart export/styles/dimensions not reproduced.'),
    ('ReportWriter.vb', 'TableFill'): ('engine.insert_table', 'Basic new table only; HTML-format fidelity, padding, inline positions, and retries not equivalent.'),
    ('ReportWriter.vb', 'RunReport1'): ('engine.render / cli.main', 'Most orchestration branches absent; automatic TOC/pagination and DOC/PDF export absent.'),
    ('ReportWriter.vb', 'Main_after_word_tool'): ('engine.render', 'Only generic elements; no Report Numbers reader or three-way legacy table dispatch.'),
    ('ReportWriter.vb', 'AssignAlt'): ('cli.resolve / engine.render', 'Data-derived alt text only; workbook override column, Excel number formats, and legacy multi-column logic absent.'),
    ('ReportWriter.vb', 'UltTableCopy'): ('engine.insert_table', 'No merged cells, per-cell Excel styles, arbitrary header-row count, or legacy fit/alignment settings.'),
    ('ReportWriter.vb', 'CopyX'): ('engine.insert_table', 'Direct DOCX table avoids HTML transport but does not reproduce Excel HTML formatting or padding.'),
    ('ReportWriter.vb', 'ExportPic'): ('engine.chart_image', 'Generic vertical bars only; no existing Excel chart types, styles, labels, source ranges, or dimensions.'),
    ('ReportWriter.vb', 'InsertPic'): ('engine.render chart branch', 'PNG embed and explicit width only; Excel dimensions, inline placement, and legacy file-readiness contract absent.'),
    ('ReportCycles.vb', 'RunonestateInLoop'): ('cli.main report loop', 'JSON/CSV preparation only; no CommonCore, workbook/VBA, or legacy postprocessing.'),
    ('ReportCycles.vb', 'Runonestate'): ('cli.main --report', 'Single manifest report works; active workbook, licensing, progress UI, and full pipeline absent.'),
    ('ReportCycles.vb', 'LoopReports'): ('cli.main report loop', 'JSON report IDs only; workbook loop list/recalculation, all feature branches, and UI restoration absent.'),
    ('GrabFigureData.vb', 'ImportCSV'): ('cli.resolve CSV reader', 'Reads dictionaries only; no staging sheet, offsets, typed Excel parsing, or QueryTables settings equivalence.'),
    ('GrabFigureData.vb', 'ImportCSVsmart'): ('cli.resolve CSV reader', 'Reads configured file only; no staging-clear/range/format workflow or mass-table integration.'),
}

def classify(row, generated):
    file, name = row['file'], row['name']
    if generated:
        if name not in GENERATED:
            raise ValueError(f'Undocumented generated routine: {file}:{name}')
        return GENERATED[name], 'Host-specific', '', 'Generated VSTO/.NET support: retain if ribbon is retained; otherwise require an explicit replacement/retirement decision.'
    description = DESCRIPTIONS.get(file, {}).get(name)
    if description is None and file in ('ReportWriter.vb', 'MassTabMake.vb'):
        description = HELPERS.get(name)
    if description is None:
        raise ValueError(f'Undocumented routine: {file}:{name}')
    if (file, name) in PARTIAL:
        replacement, gap = PARTIAL[(file, name)]
        return description, 'Partial', replacement, gap
    if file in ('AccessX.vb', 'EncryptX.vb') and name not in ('Test100', 'Testrun', 'Encrypt'):
        return description, 'Held', '', 'H03: legacy Access/license compatibility needs fixtures and policy decisions; original production gating is not present in Python.'
    if name in ('CommonCore', 'SetThisWorkbook'):
        return description, 'Held', '', 'H02: workbook calculation/VBA strategy not implemented; retain complete behavior in scope.'
    if row['kind'].startswith('Declare') or name.startswith('ClearClipboard') or name in ('KillWord', 'NAR', 'ThisAddIn_Shutdown', 'Waitsec'):
        return description, 'Replacement pending', '', 'Office/clipboard/process-specific mechanism; replace outcome with scoped resources and file handling, do not blindly reproduce desktop workarounds.'
    if row['empty_body'] or name in ('Test', 'Test100', 'Testrun', 'HandleX', 'Encrypt', 'openConnection', 'TestConnection', 'DoubleQuotes', 'Changeheader2', 'Delextrarows', 'DelextraTOProws', 'JustImportMe', 'JustImportMe2', 'Grabdata', 'GrabdataExcelOld', 'GrabdataExcel', 'Makebullets', 'ColLetter', 'CopyTableAgain', 'CopySectionAgain'):
        return description, 'Review candidate', '', 'Retained in inventory/scope. Static reference evidence and workbook hooks need review before any retirement; not declared unused.'
    return description, 'Not converted', '', 'No behavior-equivalent Python implementation or parity acceptance test exists yet.'


def main():
    inventory = json.loads((DOCS / 'legacy-inventory.json').read_text())
    generated = {f['file']: f['generated'] for f in inventory['vbnet_files']}
    ledger = []
    for row in inventory['vbnet_declarations']:
        description, status, replacement, gap = classify(row, generated[row['file']])
        ledger.append(dict(row, language='VB.NET', purpose=description, status=status,
                           python_replacement=replacement, gap=gap,
                           verification=('Related Python unit/example checks only; no complete legacy parity test.'
                                         if status == 'Partial' else 'Source inspection only; not converted or parity-verified.')))
    vba_desc = {
        'ProcessJobOne': 'Save workbook; create TRWHubToKernel.xlsx with hub path/name and RunOne; launch external kernel.',
        'ProcessJobMany': 'Save workbook; create TRWHubToKernel.xlsx with RunMany; launch external kernel.',
        'ShellMe': 'Launch Report Writer Kernel.exe from configured folder through Windows Shell.',
        'test': 'Display a Running Excel Macros diagnostic message.',
        'AddMacros': 'User-defined macro hook invoked by CommonCore; supplied demo body contains only a commented test call.',
        'chartidx': 'Display selected Excel chart name stripped of worksheet prefix.',
    }
    for workbook in inventory['workbooks']:
        for row in workbook['vba_declarations']:
            if row['name'] not in vba_desc:
                raise ValueError('Undocumented VBA routine: ' + row['id'])
            ledger.append(dict(row, file=workbook['name'] + '/' + row['file'], language='VBA',
                               purpose=vba_desc[row['name']], status='Held', python_replacement='',
                               gap=('H02: runtime VBA extension contract unimplemented.' if row['name'] == 'AddMacros'
                                    else 'H04: workbook macro/old kernel interface unimplemented; legacy kernel executable is absent from supplied package.'),
                               verification='Extracted VBA source only; no macro execution or parity test.'))
    (DOCS / 'conversion-ledger.json').write_text(json.dumps(ledger, indent=2) + '\n')
    columns = ['id', 'language', 'file', 'scope', 'name', 'kind', 'start_line', 'end_line',
               'signature', 'purpose', 'status', 'python_replacement', 'gap', 'verification']
    with (DOCS / 'conversion-ledger.csv').open('w', newline='', encoding='utf-8') as handle:
        writer = csv.DictWriter(handle, fieldnames=columns, extrasaction='ignore')
        writer.writeheader()
        writer.writerows(ledger)
    lines = ['# Function-by-function conversion ledger', '',
             'Baseline: original uploaded archives and Python prototype commit `26ab3a4`.', '',
             'Every compiled VB.NET declaration is listed, including constructors, properties,',
             'Win32 declarations, empty events, and generated designer support. VBA declarations',
             'are listed per workbook occurrence. Line numbers refer to the original archive,',
             'not to decompiled DLLs. No original routine is marked **Complete**.', '',
             'See [scope, status definitions, and holds](CONVERSION_AUDIT.md),',
             '[workbook tabs](WORKBOOK_TABS.md), and the sortable [CSV ledger](conversion-ledger.csv).', '',
             'Descriptions are reviewed source interpretations. Reference candidates in the JSON',
             'are lexical evidence, not proof that a routine is reachable or unused.', '']
    for file in dict.fromkeys(r['file'] for r in ledger):
        lines += [f'## {file}', '', '| Routine (original lines) | Responsibility | Status | Python coverage / remaining gap |',
                  '|---|---|---|---|']
        for r in ledger:
            if r['file'] != file:
                continue
            name = r['scope'] + '.' + r['name']
            coverage = (f'`{r["python_replacement"]}`. ' if r['python_replacement'] else '') + r['gap']
            lines.append(f'| `{name}` ({r["start_line"]}–{r["end_line"]}) | {r["purpose"]} | **{r["status"]}** | {coverage} |')
        lines.append('')
    lines += ['## Files with no routine declarations', '',
              '`My Project/AssemblyInfo.vb` contains assembly metadata (45 lines). It is compiled',
              'and inventoried, with zero callable declarations. Python project metadata replaces',
              'some packaging concerns; original signing/version/deployment behavior is not ported.', '']
    (DOCS / 'FUNCTION_LEDGER.md').write_text('\n'.join(lines))
    py_inventory = []
    for file in sorted((ROOT / 'trw').glob('*.py')):
        for node in ast.walk(ast.parse(file.read_text())):
            if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
                py_inventory.append({'file': str(file.relative_to(ROOT)), 'name': node.name,
                                     'start_line': node.lineno, 'end_line': node.end_lineno})
    (DOCS / 'python-inventory.json').write_text(json.dumps(py_inventory, indent=2) + '\n')
    print('Ledger rows:', len(ledger), 'statuses:', dict(Counter(r['status'] for r in ledger)))
    if {r['id'] for r in ledger if r['language'] == 'VB.NET'} != {r['id'] for r in inventory['vbnet_declarations']}:
        raise ValueError('VB.NET coverage mismatch')
    if len({r['id'] for r in ledger}) != len(ledger):
        raise ValueError('Duplicate ledger ID')


if __name__ == '__main__':
    main()
