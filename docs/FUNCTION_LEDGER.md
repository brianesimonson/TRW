# Function-by-function conversion ledger

Baseline: original uploaded archives and Python prototype commit `26ab3a4`.

Every compiled VB.NET declaration is listed, including constructors, properties,
Win32 declarations, empty events, and generated designer support. VBA declarations
are listed per workbook occurrence. Line numbers refer to the original archive,
not to decompiled DLLs. No original routine is marked **Complete**.

See [scope, status definitions, and holds](CONVERSION_AUDIT.md),
[workbook tabs](WORKBOOK_TABS.md), and the sortable [CSV ledger](conversion-ledger.csv).

Descriptions are reviewed source interpretations. Reference candidates in the JSON
are lexical evidence, not proof that a routine is reachable or unused.

## AccessX.vb

| Routine (original lines) | Responsibility | Status | Python coverage / remaining gap |
|---|---|---|---|
| `AccessX.Test100` (7–11) | Empty commented-out identity test stub. | **Review candidate** | Retained in inventory/scope. Static reference evidence and workbook hooks need review before any retirement; not declared unused. |
| `AccessX.CheckIdentityStrict` (13–53) | Read UserAccess in Access; distinguish permitted, denied, and unknown Windows users. | **Held** | H03: legacy Access/license compatibility needs fixtures and policy decisions; original production gating is not present in Python. |
| `AccessX.CheckIdentity` (54–114) | Check UserAccess; automatically add and permit unknown users in non-strict mode. | **Held** | H03: legacy Access/license compatibility needs fixtures and policy decisions; original production gating is not present in Python. |
| `AccessX.CheckStrictness` (115–138) | Read StrictAccess policy from the Access database. | **Held** | H03: legacy Access/license compatibility needs fixtures and policy decisions; original production gating is not present in Python. |
| `AccessX.LogTransaction` (139–183) | Write user, workbook path, timestamp, and single/batch report count to UserLogs. | **Held** | H03: legacy Access/license compatibility needs fixtures and policy decisions; original production gating is not present in Python. |
| `AccessX.RequestReport` (185–214) | Gate report execution on license and identity, then log the request; return status 1–4. | **Held** | H03: legacy Access/license compatibility needs fixtures and policy decisions; original production gating is not present in Python. |
| `AccessX.VerifyLicense` (215–246) | Read encrypted license records from IHTNGU, decrypt using current month/year, and validate marker. | **Held** | H03: legacy Access/license compatibility needs fixtures and policy decisions; original production gating is not present in Python. |
| `AccessX.Testrun` (247–251) | Manual test entry to request one licensed report. | **Review candidate** | Retained in inventory/scope. Static reference evidence and workbook hooks need review before any retirement; not declared unused. |
| `AccessX.SetConnstr` (252–276) | Resolve TRWlogs.accdb beside the active workbook and return ACE connection string; both active candidate paths are identical, older alternate path is commented. | **Held** | H03: legacy Access/license compatibility needs fixtures and policy decisions; original production gating is not present in Python. |

## ChangeChart.vb

| Routine (original lines) | Responsibility | Status | Python coverage / remaining gap |
|---|---|---|---|
| `ChangeChart.CHrt` (6–53) | Read Report Numbers chart rows and reset named Excel chart source ranges; suppress individual errors. | **Not converted** | No behavior-equivalent Python implementation or parity acceptance test exists yet. |

## EncryptX.vb

| Routine (original lines) | Responsibility | Status | Python coverage / remaining gap |
|---|---|---|---|
| `EncryptX.Encrypt` (11–34) | Encrypt Unicode text with AES, derived key/IV, and Base64 output. | **Review candidate** | Retained in inventory/scope. Static reference evidence and workbook hooks need review before any retirement; not declared unused. |
| `EncryptX.Decrypt` (36–63) | Decrypt Base64 license text using date-dependent key material; decode Unicode text. | **Held** | H03: legacy Access/license compatibility needs fixtures and policy decisions; original production gating is not present in Python. |

## GrabFigureData.vb

| Routine (original lines) | Responsibility | Status | Python coverage / remaining gap |
|---|---|---|---|
| `GrabFigureData.Grabdata` (23–147) | Alternate importer opens a separate Excel instance and attempts formatted-cell copy; falls back on selected errors. | **Review candidate** | Retained in inventory/scope. Static reference evidence and workbook hooks need review before any retirement; not declared unused. |
| `GrabFigureData.GrabdataExcelOld` (148–251) | Older separate-Excel importer using FigureDataImport offsets and copy dimensions. | **Review candidate** | Retained in inventory/scope. Static reference evidence and workbook hooks need review before any retirement; not declared unused. |
| `GrabFigureData.GrabdataExcel` (252–392) | Intermediate external-workbook importer with per-cell COM release and staging copy. | **Review candidate** | Retained in inventory/scope. Static reference evidence and workbook hooks need review before any retirement; not declared unused. |
| `GrabFigureData.GrabdataExcel4` (393–535) | Current importer: resolve filename, query into CSVIMPORTS, clear destination rectangle, copy offset rows/columns, clean staging. | **Not converted** | No behavior-equivalent Python implementation or parity acceptance test exists yet. |
| `GrabFigureData.Makebullets` (536–564) | Select one of three domain-specific bullet transformations from FigureDataImport column L. | **Review candidate** | Retained in inventory/scope. Static reference evidence and workbook hooks need review before any retirement; not declared unused. |
| `GrabFigureData.CallFigures` (568–605) | Import configured rows 4–100, run COM cleanup, and refresh chart ranges via ChangeChart.CHrt. | **Not converted** | No behavior-equivalent Python implementation or parity acceptance test exists yet. |
| `GrabFigureData.ColNum` (608–666) | Map A–Z to 1–26; return legacy sentinel 1000 for other inputs. | **Not converted** | No behavior-equivalent Python implementation or parity acceptance test exists yet. |
| `GrabFigureData.ColLetter` (667–727) | Map 1–26 to A–Z; return AA for other inputs. | **Review candidate** | Retained in inventory/scope. Static reference evidence and workbook hooks need review before any retirement; not declared unused. |
| `GrabFigureData.HandleX` (729–731) | Diagnostic message-box callback; former commented library hook. | **Review candidate** | Retained in inventory/scope. Static reference evidence and workbook hooks need review before any retirement; not declared unused. |
| `GrabFigureData.JustImportMe` (733–772) | Old one-time-import entry; creates Excel instance, calls GrabdataExcelOld, releases COM objects. | **Review candidate** | Retained in inventory/scope. Static reference evidence and workbook hooks need review before any retirement; not declared unused. |
| `GrabFigureData.JustImportMe3` (773–802) | Current ribbon one-time-import entry: imports configured rows using GrabdataExcel4, cleans COM, shows completion. | **Not converted** | No behavior-equivalent Python implementation or parity acceptance test exists yet. |
| `GrabFigureData.JustImportMe2` (803–844) | Intermediate one-time-import entry calling GrabdataExcel with COM cleanup. | **Review candidate** | Retained in inventory/scope. Static reference evidence and workbook hooks need review before any retirement; not declared unused. |
| `GrabFigureData.NAR` (845–853) | Repeatedly release a COM object; swallow cleanup errors. | **Replacement pending** | Office/clipboard/process-specific mechanism; replace outcome with scoped resources and file handling, do not blindly reproduce desktop workarounds. |
| `GrabFigureData.ImportCSV` (854–903) | Import CSV using Excel QueryTables and legacy parse settings; remove connections/query tables afterward. | **Partial** | `cli.resolve CSV reader`. Reads dictionaries only; no staging sheet, offsets, typed Excel parsing, or QueryTables settings equivalence. |
| `GrabFigureData.ImportCSVsmart` (904–955) | Clear staging, derive basename from full path, query CSV into CSVIMPORTS, remove connections/query tables. | **Partial** | `cli.resolve CSV reader`. Reads configured file only; no staging-clear/range/format workflow or mass-table integration. |

## MassTabMake.vb

| Routine (original lines) | Responsibility | Status | Python coverage / remaining gap |
|---|---|---|---|
| `MassTabMake.Main` (8–24) | Iterate index_temp rows until blank column-C tag; invoke SearchAndReplace for each tag. | **Not converted** | No behavior-equivalent Python implementation or parity acceptance test exists yet. |
| `MassTabMake.SearchAndReplace` (25–222) | Import indexed CSV, parse per-column formats, fill scalar/fixed table modes 1/3 or expand rows in mode 2. | **Not converted** | No behavior-equivalent Python implementation or parity acceptance test exists yet. |
| `MassTabMake.Xformat` (224–244) | Extract formatN segments from a concatenated per-column format specification. | **Not converted** | No behavior-equivalent Python implementation or parity acceptance test exists yet. |
| `MassTabMake.Startcolumn1` (246–263) | Parse the starting column from a dollar-qualified Excel address. | **Not converted** | No behavior-equivalent Python implementation or parity acceptance test exists yet. |
| `MassTabMake.Startrow1` (264–280) | Parse starting row from an Excel range; legacy single-cell branch returns zero. | **Not converted** | No behavior-equivalent Python implementation or parity acceptance test exists yet. |
| `MassTabMake.Totalcolumnmoves` (281–301) | Return end-column minus start-column (a movement count, not inclusive width). | **Not converted** | No behavior-equivalent Python implementation or parity acceptance test exists yet. |
| `MassTabMake.Totalrowmoves` (302–319) | Return end-row minus start-row (a movement count, not inclusive height). | **Not converted** | No behavior-equivalent Python implementation or parity acceptance test exists yet. |
| `MassTabMake.Test` (320–324) | Manual Countcolumns("AZ") message-box test. | **Review candidate** | Retained in inventory/scope. Static reference evidence and workbook hooks need review before any retirement; not declared unused. |
| `MassTabMake.Countcolumns` (325–331) | Convert one/two-letter column names using Firstchar plus Secondchar. | **Not converted** | No behavior-equivalent Python implementation or parity acceptance test exists yet. |
| `MassTabMake.Firstchar` (332–393) | Map first letter to a multiple-of-26 column offset; percent sentinel means zero. | **Not converted** | No behavior-equivalent Python implementation or parity acceptance test exists yet. |
| `MassTabMake.Secondchar` (394–454) | Map second letter to 1–26. | **Not converted** | No behavior-equivalent Python implementation or parity acceptance test exists yet. |
| `MassTabMake.CellToRange` (455–463) | Convert numeric column index to a one/two-letter address through Convertme. | **Not converted** | No behavior-equivalent Python implementation or parity acceptance test exists yet. |
| `MassTabMake.Convertme` (464–524) | Map 0–26 to empty/A–Z; return A outside that range. | **Not converted** | No behavior-equivalent Python implementation or parity acceptance test exists yet. |

## ModuleConnection.vb

| Routine (original lines) | Responsibility | Status | Python coverage / remaining gap |
|---|---|---|---|
| `ModuleConnection.openConnection` (5–22) | Connect to MySQL using embedded workstation connection settings; display connection result. | **Review candidate** | Retained in inventory/scope. Static reference evidence and workbook hooks need review before any retirement; not declared unused. |
| `ModuleConnection.TestConnection` (24–51) | Call openConnection and enumerate database schemas; diagnostic/test workflow. | **Review candidate** | Retained in inventory/scope. Static reference evidence and workbook hooks need review before any retirement; not declared unused. |

## My Project/Resources.Designer.vb

| Routine (original lines) | Responsibility | Status | Python coverage / remaining gap |
|---|---|---|---|
| `Resources.ResourceManager` (39–47) | Lazily create generated embedded-resource manager. | **Host-specific** | Generated VSTO/.NET support: retain if ribbon is retained; otherwise require an explicit replacement/retirement decision. |
| `Resources.Culture` (54–61) | Get/set resource culture override. | **Host-specific** | Generated VSTO/.NET support: retain if ribbon is retained; otherwise require an explicit replacement/retirement decision. |

## My Project/Settings.Designer.vb

| Routine (original lines) | Responsibility | Status | Python coverage / remaining gap |
|---|---|---|---|
| `MySettings.AutoSaveSettings` (32–36) | Persist generated application settings on shutdown when requested. | **Host-specific** | Generated VSTO/.NET support: retain if ribbon is retained; otherwise require an explicit replacement/retirement decision. |
| `MySettings.Default` (40–55) | Expose thread-safe default generated settings instance. | **Host-specific** | Generated VSTO/.NET support: retain if ribbon is retained; otherwise require an explicit replacement/retirement decision. |
| `MySettingsProperty.Settings` (67–71) | Expose generated My.Settings accessor. | **Host-specific** | Generated VSTO/.NET support: retain if ribbon is retained; otherwise require an explicit replacement/retirement decision. |

## ProcessingFrm.Designer.vb

| Routine (original lines) | Responsibility | Status | Python coverage / remaining gap |
|---|---|---|---|
| `Processing.Dispose` (9–17) | Dispose generated form/ribbon components and call base cleanup. | **Host-specific** | Generated VSTO/.NET support: retain if ribbon is retained; otherwise require an explicit replacement/retirement decision. |
| `Processing.InitializeComponent` (26–137) | Construct designer-defined ribbon controls or progress form controls. | **Host-specific** | Generated VSTO/.NET support: retain if ribbon is retained; otherwise require an explicit replacement/retirement decision. |

## ProcessingFrm.vb

| Routine (original lines) | Responsibility | Status | Python coverage / remaining gap |
|---|---|---|---|
| `Processing.Processing_Load` (3–5) | Empty processing-form load event. | **Review candidate** | Retained in inventory/scope. Static reference evidence and workbook hooks need review before any retirement; not declared unused. |
| `Processing.GroupBox1_Enter` (7–9) | Empty group-box event routine; no Handles binding in declaration. | **Review candidate** | Retained in inventory/scope. Static reference evidence and workbook hooks need review before any retirement; not declared unused. |

## RWRibbon.Designer.vb

| Routine (original lines) | Responsibility | Status | Python coverage / remaining gap |
|---|---|---|---|
| `RWRibbon.New` (5–13) | Construct generated Office host/ribbon/global support object. | **Host-specific** | Generated VSTO/.NET support: retain if ribbon is retained; otherwise require an explicit replacement/retirement decision. |
| `RWRibbon.New` (16–22) | Construct generated Office host/ribbon/global support object. | **Host-specific** | Generated VSTO/.NET support: retain if ribbon is retained; otherwise require an explicit replacement/retirement decision. |
| `RWRibbon.Dispose` (26–34) | Dispose generated form/ribbon components and call base cleanup. | **Host-specific** | Generated VSTO/.NET support: retain if ribbon is retained; otherwise require an explicit replacement/retirement decision. |
| `RWRibbon.InitializeComponent` (43–151) | Construct designer-defined ribbon controls or progress form controls. | **Host-specific** | Generated VSTO/.NET support: retain if ribbon is retained; otherwise require an explicit replacement/retirement decision. |
| `ThisRibbonCollection.Ribbon1` (168–172) | Expose generated ribbon collection accessor. | **Host-specific** | Generated VSTO/.NET support: retain if ribbon is retained; otherwise require an explicit replacement/retirement decision. |

## RWRibbon.vb

| Routine (original lines) | Responsibility | Status | Python coverage / remaining gap |
|---|---|---|---|
| `RWRibbon.Ribbon1_Load` (5–7) | Empty ribbon load event. | **Review candidate** | Retained in inventory/scope. Static reference evidence and workbook hooks need review before any retirement; not declared unused. |
| `RWRibbon.Button2_Click` (10–63) | Run All Reports button: progress UI, license gate, Word-close check, batch generation, error dialogs, cleanup. | **Not converted** | No behavior-equivalent Python implementation or parity acceptance test exists yet. |
| `RWRibbon.Button1_Click_1` (65–121) | Run One Report button: progress UI, license gate, Word-close check, single generation, error dialogs, cleanup. | **Not converted** | No behavior-equivalent Python implementation or parity acceptance test exists yet. |
| `RWRibbon.Button3_Click` (123–167) | One Time Import button: progress UI, JustImportMe3, workbook/error handling. | **Not converted** | No behavior-equivalent Python implementation or parity acceptance test exists yet. |
| `RWRibbon.ChartID_Click` (169–202) | Chart ID button: read selected chart name and strip worksheet name; show diagnostic/error message. | **Not converted** | No behavior-equivalent Python implementation or parity acceptance test exists yet. |

## ReportCycles.vb

| Routine (original lines) | Responsibility | Status | Python coverage / remaining gap |
|---|---|---|---|
| `ReportCycles.MyForm` (15–21) | Lazily construct and expose the topmost Processing progress form. | **Not converted** | No behavior-equivalent Python implementation or parity acceptance test exists yet. |
| `ReportCycles.FileNotExist.New` (26–28) | Construct FileNotExist/MyErrorMessage exceptions with message and optional inner exception. | **Not converted** | No behavior-equivalent Python implementation or parity acceptance test exists yet. |
| `ReportCycles.FileNotExist.New` (30–33) | Construct FileNotExist/MyErrorMessage exceptions with message and optional inner exception. | **Not converted** | No behavior-equivalent Python implementation or parity acceptance test exists yet. |
| `ReportCycles.MyErrorMessage.New` (42–44) | Construct FileNotExist/MyErrorMessage exceptions with message and optional inner exception. | **Not converted** | No behavior-equivalent Python implementation or parity acceptance test exists yet. |
| `ReportCycles.MyErrorMessage.New` (46–49) | Construct FileNotExist/MyErrorMessage exceptions with message and optional inner exception. | **Not converted** | No behavior-equivalent Python implementation or parity acceptance test exists yet. |
| `ReportCycles.RunonestateInLoop` (54–60) | Run CommonCore followed by the main report generator for the current batch item. | **Partial** | `cli.main report loop`. JSON/CSV preparation only; no CommonCore, workbook/VBA, or legacy postprocessing. |
| `ReportCycles.Runonestate` (62–96) | Capture active sheet, import/prepare workbook, launch Word, generate one report, quit Word, restore Excel UI. | **Partial** | `cli.main --report`. Single manifest report works; active workbook, licensing, progress UI, and full pipeline absent. |
| `ReportCycles.LoopReports` (97–141) | Launch Word once; iterate Report Loop List IDs, change Current Parameters A11, prepare and render each; restore UI. | **Partial** | `cli.main report loop`. JSON report IDs only; workbook loop list/recalculation, all feature branches, and UI restoration absent. |
| `ReportCycles.CommonCore` (142–203) | Build substituted index_temp for mass tables, optionally auto-import, and optionally invoke TRW.AddMacros. | **Held** | H02: workbook calculation/VBA strategy not implemented; retain complete behavior in scope. |
| `ReportCycles.CountLoopReports` (204–229) | Count contiguous report IDs starting at Report Loop List A2. | **Not converted** | No behavior-equivalent Python implementation or parity acceptance test exists yet. |
| `ReportCycles.SetThisWorkbook` (231–235) | Bind shared xlapp workbook from the VSTO host active workbook. | **Held** | H02: workbook calculation/VBA strategy not implemented; retain complete behavior in scope. |
| `ReportCycles.KillWord` (236–259) | Ask to close all existing winword processes; kill them if the user agrees. | **Replacement pending** | Office/clipboard/process-specific mechanism; replace outcome with scoped resources and file handling, do not blindly reproduce desktop workarounds. |

## ReportWriter.vb

| Routine (original lines) | Responsibility | Status | Python coverage / remaining gap |
|---|---|---|---|
| `ReportWriter.CloseClipboard` (6–6) | Win32 declaration closing the system clipboard. | **Replacement pending** | Office/clipboard/process-specific mechanism; replace outcome with scoped resources and file handling, do not blindly reproduce desktop workarounds. |
| `ReportWriter.EmptyClipboard` (7–7) | Win32 declaration clearing the system clipboard. | **Replacement pending** | Office/clipboard/process-specific mechanism; replace outcome with scoped resources and file handling, do not blindly reproduce desktop workarounds. |
| `ReportWriter.OpenClipboard` (8–8) | Win32 declaration opening the system clipboard. | **Replacement pending** | Office/clipboard/process-specific mechanism; replace outcome with scoped resources and file handling, do not blindly reproduce desktop workarounds. |
| `ReportWriter.SearchAndReplace` (41–100) | Find every text tag, apply Report Numbers column C value and Excel number format, split long text writes. | **Partial** | `engine.replace_text / render`. No workbook C-column adapter, Excel format semantics, or legacy case-insensitive matching; new tag grammar is narrower. |
| `ReportWriter.ChartFill` (101–148) | Find each chart tag and insert the exported chart image at Word selection. | **Partial** | `engine.render chart branch`. Standalone paragraph tags only; Excel chart export/styles/dimensions not reproduced. |
| `ReportWriter.TableFill` (149–225) | Find a table tag; insert published Excel HTML via CopyX; retry table-copy errors. | **Partial** | `engine.insert_table`. Basic new table only; HTML-format fidelity, padding, inline positions, and retries not equivalent. |
| `ReportWriter.Waitsec` (227–233) | Wait five seconds through Excel Application.Wait. | **Replacement pending** | Office/clipboard/process-specific mechanism; replace outcome with scoped resources and file handling, do not blindly reproduce desktop workarounds. |
| `ReportWriter.ClearClipboard01` (235–240) | Clear Excel clipboard via the Clipboard command bar. | **Replacement pending** | Office/clipboard/process-specific mechanism; replace outcome with scoped resources and file handling, do not blindly reproduce desktop workarounds. |
| `ReportWriter.ClearClipboard02` (241–244) | Disable copying objects with Excel cells. | **Replacement pending** | Office/clipboard/process-specific mechanism; replace outcome with scoped resources and file handling, do not blindly reproduce desktop workarounds. |
| `ReportWriter.ClearClipboardEl2000` (246–249) | Clear Excel clipboard with legacy command ID 3634. | **Replacement pending** | Office/clipboard/process-specific mechanism; replace outcome with scoped resources and file handling, do not blindly reproduce desktop workarounds. |
| `ReportWriter.ClearClipboard` (252–260) | Open, clear, and close the Windows clipboard using Win32 APIs. | **Replacement pending** | Office/clipboard/process-specific mechanism; replace outcome with scoped resources and file handling, do not blindly reproduce desktop workarounds. |
| `ReportWriter.ClearClipboard100` (265–271) | Alternate Windows clipboard-clear implementation. | **Replacement pending** | Office/clipboard/process-specific mechanism; replace outcome with scoped resources and file handling, do not blindly reproduce desktop workarounds. |
| `ReportWriter.RunReport1` (277–455) | Open/copy template, cut sections, mass tables, populate elements, insert docs, renumber, change headers, normalize missing values, update TOC, export DOCX/DOC/PDF. | **Partial** | `engine.render / cli.main`. Most orchestration branches absent; automatic TOC/pagination and DOC/PDF export absent. |
| `ReportWriter.Main_after_word_tool` (456–677) | Populate text/charts/tables; choose shell-table, direct-cell table, or Excel HTML branch; apply size, alignment, headers, alt metadata, spacing. | **Partial** | `engine.render`. Only generic elements; no Report Numbers reader or three-way legacy table dispatch. |
| `ReportWriter.CopyTableAgain` (678–689) | Clipboard copy of Excel range then retry TableFill. | **Review candidate** | Retained in inventory/scope. Static reference evidence and workbook hooks need review before any retirement; not declared unused. |
| `ReportWriter.Changeheader` (690–723) | Write configured headers/footers and alignment for up to five Word sections. | **Not converted** | No behavior-equivalent Python implementation or parity acceptance test exists yet. |
| `ReportWriter.Changeheader2` (724–754) | Older header-only routine driven by Header tab; adjusts Word view and header selection. | **Review candidate** | Retained in inventory/scope. Static reference evidence and workbook hooks need review before any retirement; not declared unused. |
| `ReportWriter.Cutsections` (755–823) | Use SmartSections flags to remove Xsection…S/E regions or remove delimiters; insert reason text and bullet-adjacent edits. | **Not converted** | No behavior-equivalent Python implementation or parity acceptance test exists yet. |
| `ReportWriter.RenameTables` (825–872) | Replace TableX placeholder labels with numbers from SmartSections J:L. | **Not converted** | No behavior-equivalent Python implementation or parity acceptance test exists yet. |
| `ReportWriter.RenameFigures` (873–925) | Replace FigureX placeholder labels with numbers from SmartSections P:R. | **Not converted** | No behavior-equivalent Python implementation or parity acceptance test exists yet. |
| `ReportWriter.DeleteSection` (926–935) | Delete content between start/end markers with a Word wildcard range search. | **Not converted** | No behavior-equivalent Python implementation or parity acceptance test exists yet. |
| `ReportWriter.DeleteTags` (936–948) | Delete section delimiter markers while retaining their content. | **Not converted** | No behavior-equivalent Python implementation or parity acceptance test exists yet. |
| `ReportWriter.DoubleQuotes` (950–1005) | Legacy caret-to-quote text conversion; includes an unfinished triple-caret search loop. | **Review candidate** | Retained in inventory/scope. Static reference evidence and workbook hooks need review before any retirement; not declared unused. |
| `ReportWriter.GlobalReplace` (1006–1023) | Apply contiguous GlobalReplace sheet source/destination rows. | **Not converted** | No behavior-equivalent Python implementation or parity acceptance test exists yet. |
| `ReportWriter.ReplaceMe` (1024–1052) | Repeated case-insensitive literal Word find/replace of arbitrary text. | **Not converted** | No behavior-equivalent Python implementation or parity acceptance test exists yet. |
| `ReportWriter.MissingToNA` (1053–1162) | Replace four specific formatted missing-value strings with N/A. | **Not converted** | No behavior-equivalent Python implementation or parity acceptance test exists yet. |
| `ReportWriter.AddSections` (1164–1217) | Resolve additional DOCX files from Add Docx rows; insert at tags or emit missing-file text. | **Not converted** | No behavior-equivalent Python implementation or parity acceptance test exists yet. |
| `ReportWriter.SectionFill` (1218–1273) | Replace section tag with inserted document content, or a missing-file message; retry on insertion errors. | **Not converted** | No behavior-equivalent Python implementation or parity acceptance test exists yet. |
| `ReportWriter.CopySectionAgain` (1275–1292) | Select/copy additional document content and retry SectionFill. | **Review candidate** | Retained in inventory/scope. Static reference evidence and workbook hooks need review before any retirement; not declared unused. |
| `ReportWriter.AssignAlt` (1293–1401) | Build chart descriptions from override column or formatted 2/3/many-column ranges; assign to images lacking alt text. | **Partial** | `cli.resolve / engine.render`. Data-derived alt text only; workbook override column, Excel number formats, and legacy multi-column logic absent. |
| `ReportWriter.UltTableFill` (1402–1463) | Fill existing Word table shell from formatted Excel cells; support Variable Row insertion or fixed rows. | **Not converted** | No behavior-equivalent Python implementation or parity acceptance test exists yet. |
| `ReportWriter.UltTable` (1464–1510) | Find direct-table tag and invoke UltTableCopy. | **Not converted** | No behavior-equivalent Python implementation or parity acceptance test exists yet. |
| `ReportWriter.UltTableShell` (1511–1557) | Find existing-table-shell tag and invoke UltTableFill. | **Not converted** | No behavior-equivalent Python implementation or parity acceptance test exists yet. |
| `ReportWriter.UltTableCopy` (1558–1789) | Create Word table from Excel range; reproduce merged cells, borders, alignment, fonts, colors, formatted values, repeating headers, and metadata. | **Partial** | `engine.insert_table`. No merged cells, per-cell Excel styles, arbitrary header-row count, or legacy fit/alignment settings. |
| `ReportWriter.CopyX` (1790–1837) | Insert Excel-exported HTML, adjust padding/cleanup artifacts, manipulate selection around inserted table. | **Partial** | `engine.insert_table`. Direct DOCX table avoids HTML transport but does not reproduce Excel HTML formatting or padding. |
| `ReportWriter.ExportPic` (1838–1854) | Export named Excel chart; capture chart width/height; uses GIF filter with a .jpg filename. | **Partial** | `engine.chart_image`. Generic vertical bars only; no existing Excel chart types, styles, labels, source ranges, or dimensions. |
| `ReportWriter.InsertPic` (1855–1864) | Wait for image file then embed it with captured Excel chart dimensions. | **Partial** | `engine.render chart branch`. PNG embed and explicit width only; Excel dimensions, inline placement, and legacy file-readiness contract absent. |
| `ReportWriter.FindLines` (1868–1926) | Strip Excel conditional HTML markup and apply selected left/right/top padding substitutions. | **Not converted** | No behavior-equivalent Python implementation or parity acceptance test exists yet. |
| `ReportWriter.FindLines2` (1927–1973) | Strip Excel conditional HTML markup without padding substitutions. | **Not converted** | No behavior-equivalent Python implementation or parity acceptance test exists yet. |
| `ReportWriter.Delextrarows` (1974–1999) | Selection-based cleanup below an inserted table; delete selection on table-selection failure. | **Review candidate** | Retained in inventory/scope. Static reference evidence and workbook hooks need review before any retirement; not declared unused. |
| `ReportWriter.DelextraTOProws` (2000–2025) | Selection-based cleanup above an inserted table; backspace recovery on failure. | **Review candidate** | Retained in inventory/scope. Static reference evidence and workbook hooks need review before any retirement; not declared unused. |
| `ReportWriter.Startcolumn1` (2027–2044) | Parse the starting column from a dollar-qualified Excel address. | **Not converted** | No behavior-equivalent Python implementation or parity acceptance test exists yet. |
| `ReportWriter.Startrow1` (2045–2061) | Parse starting row from an Excel range; legacy single-cell branch returns zero. | **Not converted** | No behavior-equivalent Python implementation or parity acceptance test exists yet. |
| `ReportWriter.Totalcolumnmoves` (2062–2082) | Return end-column minus start-column (a movement count, not inclusive width). | **Not converted** | No behavior-equivalent Python implementation or parity acceptance test exists yet. |
| `ReportWriter.Totalrowmoves` (2083–2100) | Return end-row minus start-row (a movement count, not inclusive height). | **Not converted** | No behavior-equivalent Python implementation or parity acceptance test exists yet. |
| `ReportWriter.Test` (2101–2105) | Manual Countcolumns("AZ") message-box test. | **Review candidate** | Retained in inventory/scope. Static reference evidence and workbook hooks need review before any retirement; not declared unused. |
| `ReportWriter.Countcolumns` (2106–2112) | Convert one/two-letter column names using Firstchar plus Secondchar. | **Not converted** | No behavior-equivalent Python implementation or parity acceptance test exists yet. |
| `ReportWriter.Firstchar` (2113–2174) | Map first letter to a multiple-of-26 column offset; percent sentinel means zero. | **Not converted** | No behavior-equivalent Python implementation or parity acceptance test exists yet. |
| `ReportWriter.Secondchar` (2175–2235) | Map second letter to 1–26. | **Not converted** | No behavior-equivalent Python implementation or parity acceptance test exists yet. |
| `ReportWriter.CellToRange` (2236–2244) | Convert numeric column index to a one/two-letter address through Convertme. | **Not converted** | No behavior-equivalent Python implementation or parity acceptance test exists yet. |
| `ReportWriter.Convertme` (2245–2306) | Map 0–26 to empty/A–Z; return A outside that range. | **Not converted** | No behavior-equivalent Python implementation or parity acceptance test exists yet. |
| `ReportWriter.XLtoWRDHorAlign` (2307–2315) | Map Excel horizontal alignment constants to Word paragraph alignment. | **Not converted** | No behavior-equivalent Python implementation or parity acceptance test exists yet. |
| `ReportWriter.XLtoWRDVerAlign` (2316–2324) | Map Excel vertical alignment constants to Word cell alignment (includes legacy value requiring validation). | **Not converted** | No behavior-equivalent Python implementation or parity acceptance test exists yet. |
| `ReportWriter.XLtoWRDBrdrThick` (2325–2333) | Map Excel border weights to Word line-width constants. | **Not converted** | No behavior-equivalent Python implementation or parity acceptance test exists yet. |
| `ReportWriter.XLtoWRDLineStyle` (2334–2351) | Map Excel line-style constants to Word border-style constants. | **Not converted** | No behavior-equivalent Python implementation or parity acceptance test exists yet. |

## ThisAddIn.Designer.vb

| Routine (original lines) | Responsibility | Status | Python coverage / remaining gap |
|---|---|---|---|
| `ThisAddIn.New` (32–35) | Construct generated Office host/ribbon/global support object. | **Host-specific** | Generated VSTO/.NET support: retain if ribbon is retained; otherwise require an explicit replacement/retirement decision. |
| `ThisAddIn.Initialize` (41–50) | Initialize VSTO host application, controls, cached data, and host item. | **Host-specific** | Generated VSTO/.NET support: retain if ribbon is retained; otherwise require an explicit replacement/retirement decision. |
| `ThisAddIn.FinishInitialization` (56–58) | Complete VSTO initialization and invoke startup. | **Host-specific** | Generated VSTO/.NET support: retain if ribbon is retained; otherwise require an explicit replacement/retirement decision. |
| `ThisAddIn.InitializeDataBindings` (64–68) | Begin/end initialization for generated data bindings. | **Host-specific** | Generated VSTO/.NET support: retain if ribbon is retained; otherwise require an explicit replacement/retirement decision. |
| `ThisAddIn.InitializeCachedData` (74–81) | Ask VSTO data host to initialize cached data when needed. | **Host-specific** | Generated VSTO/.NET support: retain if ribbon is retained; otherwise require an explicit replacement/retirement decision. |
| `ThisAddIn.InitializeData` (87–88) | Generated empty application-data hook. | **Host-specific** | Generated VSTO/.NET support: retain if ribbon is retained; otherwise require an explicit replacement/retirement decision. |
| `ThisAddIn.BindToData` (94–95) | Generated empty data-binding hook. | **Host-specific** | Generated VSTO/.NET support: retain if ribbon is retained; otherwise require an explicit replacement/retirement decision. |
| `ThisAddIn.StartCaching` (100–102) | Delegate named-member caching to VSTO data host. | **Host-specific** | Generated VSTO/.NET support: retain if ribbon is retained; otherwise require an explicit replacement/retirement decision. |
| `ThisAddIn.StopCaching` (107–109) | Delegate named-member cache stop to VSTO data host. | **Host-specific** | Generated VSTO/.NET support: retain if ribbon is retained; otherwise require an explicit replacement/retirement decision. |
| `ThisAddIn.IsCached` (114–116) | Query whether a named member is cached by the VSTO data host. | **Host-specific** | Generated VSTO/.NET support: retain if ribbon is retained; otherwise require an explicit replacement/retirement decision. |
| `ThisAddIn.BeginInitialization` (122–126) | Begin generated host/control initialization. | **Host-specific** | Generated VSTO/.NET support: retain if ribbon is retained; otherwise require an explicit replacement/retirement decision. |
| `ThisAddIn.EndInitialization` (132–136) | End generated host/control initialization. | **Host-specific** | Generated VSTO/.NET support: retain if ribbon is retained; otherwise require an explicit replacement/retirement decision. |
| `ThisAddIn.InitializeControls` (142–145) | Generated empty controls-initialization hook. | **Host-specific** | Generated VSTO/.NET support: retain if ribbon is retained; otherwise require an explicit replacement/retirement decision. |
| `ThisAddIn.InitializeComponents` (151–152) | Generated empty components-initialization hook. | **Host-specific** | Generated VSTO/.NET support: retain if ribbon is retained; otherwise require an explicit replacement/retirement decision. |
| `ThisAddIn.NeedsFill` (157–159) | Query whether the VSTO cached member needs population. | **Host-specific** | Generated VSTO/.NET support: retain if ribbon is retained; otherwise require an explicit replacement/retirement decision. |
| `ThisAddIn.OnShutdown` (165–169) | Dispose VSTO custom task panes and call base shutdown. | **Host-specific** | Generated VSTO/.NET support: retain if ribbon is retained; otherwise require an explicit replacement/retirement decision. |
| `Globals.New` (178–180) | Construct generated Office host/ribbon/global support object. | **Host-specific** | Generated VSTO/.NET support: retain if ribbon is retained; otherwise require an explicit replacement/retirement decision. |
| `Globals.ThisAddIn` (188–199) | Expose generated shared add-in instance; prevent duplicate assignment. | **Host-specific** | Generated VSTO/.NET support: retain if ribbon is retained; otherwise require an explicit replacement/retirement decision. |
| `Globals.Factory` (201–212) | Expose shared Office application factory; prevent duplicate assignment. | **Host-specific** | Generated VSTO/.NET support: retain if ribbon is retained; otherwise require an explicit replacement/retirement decision. |
| `Globals.Ribbons` (214–221) | Lazily construct generated ribbon collection. | **Host-specific** | Generated VSTO/.NET support: retain if ribbon is retained; otherwise require an explicit replacement/retirement decision. |
| `ThisRibbonCollection.New` (231–233) | Construct generated Office host/ribbon/global support object. | **Host-specific** | Generated VSTO/.NET support: retain if ribbon is retained; otherwise require an explicit replacement/retirement decision. |

## ThisAddIn.vb

| Routine (original lines) | Responsibility | Status | Python coverage / remaining gap |
|---|---|---|---|
| `ThisAddIn.ThisAddIn_Startup` (3–5) | Empty VSTO startup event. | **Review candidate** | Retained in inventory/scope. Static reference evidence and workbook hooks need review before any retirement; not declared unused. |
| `ThisAddIn.ThisAddIn_Shutdown` (7–12) | Force two garbage-collection/finalizer cycles at VSTO shutdown. | **Replacement pending** | Office/clipboard/process-specific mechanism; replace outcome with scoped resources and file handling, do not blindly reproduce desktop workarounds. |

## WaitProcedure.vb

| Routine (original lines) | Responsibility | Status | Python coverage / remaining gap |
|---|---|---|---|
| `WaitProcedure.WaitForClose` (6–22) | Poll file lock with two-second sleeps; no timeout. | **Not converted** | No behavior-equivalent Python implementation or parity acceptance test exists yet. |
| `WaitProcedure.WaitForOpen` (23–43) | Require existing file, then poll file lock every 200ms; no timeout. | **Not converted** | No behavior-equivalent Python implementation or parity acceptance test exists yet. |
| `WaitProcedure.WaitForExist` (44–66) | Retry existence ten times with 200ms sleeps; throw if still missing. | **Not converted** | No behavior-equivalent Python implementation or parity acceptance test exists yet. |
| `WaitProcedure.FileInUse` (69–80) | Try exclusive read/write file open; any exception is treated as in-use. | **Not converted** | No behavior-equivalent Python implementation or parity acceptance test exists yet. |
| `WaitProcedure.FileExists` (81–87) | Return System.IO.File.Exists result. | **Not converted** | No behavior-equivalent Python implementation or parity acceptance test exists yet. |

## bullets.vb

| Routine (original lines) | Responsibility | Status | Python coverage / remaining gap |
|---|---|---|---|
| `bullets.Bulletthis` (5–50) | Turn imported rows into header/bullet pairs, append parenthetical counts, apply fonts, and emit no-common-causes fallback. | **Not converted** | No behavior-equivalent Python implementation or parity acceptance test exists yet. |
| `bullets.Bulletthis2` (57–99) | Alternate header/bullet transform without counts; emit HCBS Medical Review no-errors fallback. | **Not converted** | No behavior-equivalent Python implementation or parity acceptance test exists yet. |
| `bullets.Bulletthis3` (102–143) | Alternate header/bullet transform without counts; emit HCBS Data Processing no-errors fallback. | **Not converted** | No behavior-equivalent Python implementation or parity acceptance test exists yet. |

## The Report Writer - Blank Shell.xlsm/ProcessReports.bas

| Routine (original lines) | Responsibility | Status | Python coverage / remaining gap |
|---|---|---|---|
| `.ProcessJobOne` (2–37) | Save workbook; create TRWHubToKernel.xlsx with hub path/name and RunOne; launch external kernel. | **Held** | H04: workbook macro/old kernel interface unimplemented; legacy kernel executable is absent from supplied package. |
| `.ProcessJobMany` (38–72) | Save workbook; create TRWHubToKernel.xlsx with RunMany; launch external kernel. | **Held** | H04: workbook macro/old kernel interface unimplemented; legacy kernel executable is absent from supplied package. |
| `.ShellMe` (73–94) | Launch Report Writer Kernel.exe from configured folder through Windows Shell. | **Held** | H04: workbook macro/old kernel interface unimplemented; legacy kernel executable is absent from supplied package. |

## The Report Writer - Finished Example.xlsm/ProcessReports.bas

| Routine (original lines) | Responsibility | Status | Python coverage / remaining gap |
|---|---|---|---|
| `.ProcessJobOne` (2–37) | Save workbook; create TRWHubToKernel.xlsx with hub path/name and RunOne; launch external kernel. | **Held** | H04: workbook macro/old kernel interface unimplemented; legacy kernel executable is absent from supplied package. |
| `.ProcessJobMany` (38–72) | Save workbook; create TRWHubToKernel.xlsx with RunMany; launch external kernel. | **Held** | H04: workbook macro/old kernel interface unimplemented; legacy kernel executable is absent from supplied package. |
| `.ShellMe` (73–94) | Launch Report Writer Kernel.exe from configured folder through Windows Shell. | **Held** | H04: workbook macro/old kernel interface unimplemented; legacy kernel executable is absent from supplied package. |

## Demo Project 1.xlsm/TRW.bas

| Routine (original lines) | Responsibility | Status | Python coverage / remaining gap |
|---|---|---|---|
| `.test` (2–11) | Display a Running Excel Macros diagnostic message. | **Held** | H04: workbook macro/old kernel interface unimplemented; legacy kernel executable is absent from supplied package. |
| `.AddMacros` (12–15) | User-defined macro hook invoked by CommonCore; supplied demo body contains only a commented test call. | **Held** | H02: runtime VBA extension contract unimplemented. |
| `.chartidx` (16–30) | Display selected Excel chart name stripped of worksheet prefix. | **Held** | H04: workbook macro/old kernel interface unimplemented; legacy kernel executable is absent from supplied package. |

## Files with no routine declarations

`My Project/AssemblyInfo.vb` contains assembly metadata (45 lines). It is compiled
and inventoried, with zero callable declarations. Python project metadata replaces
some packaging concerns; original signing/version/deployment behavior is not ported.
