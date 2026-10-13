# Workbook tabs, options, and embedded VBA

All three original XLSM files were opened read-only for inspection. Macros were
extracted, not executed. Formula counts include shared formulas expanded by
openpyxl. Reading a cached result is not formula evaluation or conversion.

There are 65 sheet occurrences and 26 distinct sheet names. No hidden sheet is
silently excluded. Current workbook coverage is recorded below; original inventory metadata does not change. Generic elements alone do not mean their source
workbook tabs have been converted.

## All tab roles

| Tab | Original responsibility | Conversion coverage |
|---|---|---|
| `Welcome` | User guidance and descriptions of workbook tabs. | Guidance retained in the workbook and new README; no runtime calculations. |
| `Current Parameters` | Template/output paths, current report ID, feature flags, export format, suffix. | Workbook adapter consumes paths/ID/options. See option table; auth/VBA explicitly excluded or deferred. |
| `FigureDataImport` | Import file/range/destination specifications, row/column limits, bullet mode. | CSV/XLSX offsets, clearing rectangle, typed values and per-ID files implemented; legacy locale inference/.xls inputs remain gaps. |
| `Report Loop List` | Batch report IDs and current-ID reference. | Contiguous A2 IDs until blank; per-ID imports and fresh formula evaluation tested. |
| `Report Numbers` | Text IDs/values/formats, chart IDs/names/ranges/alt override, table IDs/layout/format/metadata. | Text formatting, named dynamic bar charts, shell/native tables, dimensions, padding/header/alt metadata implemented; chart/HTML fidelity remains partial. |
| `ChartRanges` | Compute active table/chart dimensions from sentinel rows; build range strings. | Recorded formulas evaluated freshly; dynamic ranges recomputed for A/B inputs. |
| `Table 1` | Example age table staging, calculated presentation cells, employee total. | Formulas, formatted values and styles consumed by real-workbook pipeline. |
| `Table 2` | Advanced example table data/presentation with dynamic formulas. | Dynamic formulas and shell/native table values tested on advanced demo. |
| `Table 3` | Advanced example calculated table used by configured report mappings. | Merged native table supported/tested in corrected copy. Original mispointed mapping remains preserved. |
| `Figure 1` | Age chart data, presentation formulas, named Excel chart, override alt-text narratives. | Fresh data, named source chart, multiple bar series/dimensions and alt override supported; exact Excel styling not reproduced. |
| `Figure 2` | Gender chart data, presentation formulas, named Excel chart. | Fresh gender counts and bar-chart source dimensions supported; SmartSections removes B chart as configured. |
| `Bullet 1` | Worksheet-generated narrative and glyph columns, often inserted as a table. | Original worksheet narrative/formulas/table preserved; three legacy domain bullet transforms also available as opt-in API. |
| `Header and Footer` | Up to five section-specific header/footer texts and alignments. | Configured section 1–5 texts and alignment implemented; first/even-page variants remain gaps. |
| `x-Header` | Hidden older header/date helper text. | Hidden legacy helper remains readable; older Changeheader2 requires absent Header sheet and remains under review. |
| `SmartSections` | Include/remove decisions, reason statements, bullet-boundary flags, TableX/FigureX renumber mappings. | Include/delete/reason, caption renumbering and bullet modes implemented; complex cross-table marker ranges unsupported. |
| `Summary Sheet` | Placeholder for analysis statistics driving section choices; empty in supplied samples. | Available as ordinary worksheet/formula input; supplied sheet has no statistics to verify. No fixed business calculation invented. |
| `Add Docx` | Additional document base path, section tags, filenames; paths may be formula-generated. | Formula filenames/base path, relationship-aware body insertion and missing-file fallback implemented; inserted distinct headers/footers remain partial. |
| `Date` | TODAY/DAY/MONTH/YEAR calculations and concatenated date strings. | TODAY/DAY/MONTH/YEAR and concatenation computed freshly; deterministic --as-of accepted. |
| `CSVIMPORTS` | Scratch sheet populated by QueryTables for import/mass-table workflows. | Replaced by in-memory CSV/XLSX rows; imports do not retain an Excel QueryTable connection. |
| `GlobalReplace` | Arbitrary source/destination text replacement pairs. | On flag and contiguous case-insensitive literal replacement pairs implemented/tested. |
| `index` | Mass-table input file paths, Excel ranges, tags, modes, formats, and extra columns. | Mass-table input/range/tag/modes 1/2/3/column formats consumed. |
| `index_altered` | Alternate index used with report-specific string substitution. | Report-specific filename substitution, copying A:H, tested for A/B. |
| `index_temp` | Runtime expanded/copied mass-table index from CommonCore. | Fresh per-report in-memory index; scalar/fixed/variable modes tested. |
| `FFS MR TopX` | Hidden domain-specific header/bullet helper with IF formulas and imported values. | Hidden input/helper formulas evaluate with workbook model; no supplied report proves every domain query path. |
| `Sheet1` | Hidden A–Z lookup/development scratch text. | Hidden scratch sheet retained and addressable; no new reporting behavior invented. |
| `index_test` | Hidden test index placeholder in the simple workbooks. | Hidden test placeholder retained; no fixed production reader found, retirement unapproved. |

## Per-workbook inventory

Formula function names below describe observed formulas, not a supported Python
formula language. Formulas also contain operators, literals, and cross-sheet
references. Every formula expression, its cell address and number format, plus
Excel chart types and series references, are recorded in the JSON.

### The Report Writer - Blank Shell.xlsm

20 tabs; 186 formula cells; 23 extracted VBA modules.

| Tab | Visibility | Formula cells | Formula functions observed | Excel charts |
|---|---|---:|---|---:|
| `Welcome` | visible | 0 | — | 0 |
| `Current Parameters` | visible | 1 | — | 0 |
| `FigureDataImport` | visible | 1 | — | 0 |
| `Report Loop List` | visible | 1 | — | 0 |
| `Report Numbers` | visible | 0 | — | 0 |
| `ChartRanges` | visible | 4 | CONCATENATE, COUNTA, COUNTIF, LOOKUP | 0 |
| `Table 1` | visible | 19 | IF, SUM | 0 |
| `Figure 1` | visible | 27 | IF, TEXT | 1 |
| `FFS MR TopX` | hidden | 64 | IF | 0 |
| `Sheet1` | hidden | 0 | — | 0 |
| `index_test` | hidden | 0 | — | 0 |
| `Figure 2` | visible | 27 | IF, TEXT | 1 |
| `Bullet 1` | visible | 28 | IF | 0 |
| `Header and Footer` | visible | 0 | — | 0 |
| `x-Header` | hidden | 1 | — | 0 |
| `SmartSections` | visible | 0 | — | 0 |
| `Summary Sheet` | visible | 0 | — | 0 |
| `Add Docx` | hidden | 8 | — | 0 |
| `Date` | visible | 5 | DAY, MONTH, TODAY, YEAR | 0 |
| `CSVIMPORTS` | visible | 0 | — | 0 |

VBA modules (including sheet/document stubs with no routines):

| Module | Source lines | Declared routines |
|---|---:|---:|
| `ThisWorkbook.cls` | 8 | 0 |
| `Sheet11.cls` | 8 | 0 |
| `Sheet36.cls` | 8 | 0 |
| `Sheet1.cls` | 8 | 0 |
| `Sheet38.cls` | 8 | 0 |
| `Sheet23.cls` | 8 | 0 |
| `Sheet10.cls` | 8 | 0 |
| `Sheet7.cls` | 8 | 0 |
| `ProcessReports.bas` | 95 | 3 |
| `Sheet2.cls` | 8 | 0 |
| `Sheet4.cls` | 8 | 0 |
| `Sheet14.cls` | 8 | 0 |
| `Sheet35.cls` | 8 | 0 |
| `Sheet24.cls` | 8 | 0 |
| `Sheet6.cls` | 8 | 0 |
| `Sheet9.cls` | 8 | 0 |
| `Sheet5.cls` | 8 | 0 |
| `Sheet32.cls` | 8 | 0 |
| `Sheet33.cls` | 8 | 0 |
| `Sheet40.cls` | 8 | 0 |
| `Sheet3.cls` | 8 | 0 |
| `Module1.bas` | 2 | 0 |
| `Sheet8.cls` | 8 | 0 |

### The Report Writer - Finished Example.xlsm

20 tabs; 202 formula cells; 23 extracted VBA modules.

| Tab | Visibility | Formula cells | Formula functions observed | Excel charts |
|---|---|---:|---|---:|
| `Welcome` | visible | 0 | — | 0 |
| `Current Parameters` | visible | 1 | — | 0 |
| `FigureDataImport` | visible | 1 | — | 0 |
| `Report Loop List` | visible | 1 | — | 0 |
| `Report Numbers` | visible | 6 | — | 0 |
| `ChartRanges` | visible | 13 | CONCATENATE, COUNTA, COUNTIF, LOOKUP | 0 |
| `Table 1` | visible | 19 | IF, SUM | 0 |
| `Figure 1` | visible | 27 | IF, TEXT | 1 |
| `FFS MR TopX` | hidden | 64 | IF | 0 |
| `Sheet1` | hidden | 0 | — | 0 |
| `index_test` | hidden | 0 | — | 0 |
| `Figure 2` | visible | 27 | IF, TEXT | 1 |
| `Bullet 1` | visible | 28 | IF, LOWER | 0 |
| `Header and Footer` | visible | 1 | — | 0 |
| `x-Header` | hidden | 1 | — | 0 |
| `SmartSections` | visible | 0 | — | 0 |
| `Summary Sheet` | visible | 0 | — | 0 |
| `Add Docx` | hidden | 8 | — | 0 |
| `Date` | visible | 5 | DAY, MONTH, TODAY, YEAR | 0 |
| `CSVIMPORTS` | visible | 0 | — | 0 |

VBA modules (including sheet/document stubs with no routines):

| Module | Source lines | Declared routines |
|---|---:|---:|
| `ThisWorkbook.cls` | 8 | 0 |
| `Sheet11.cls` | 8 | 0 |
| `Sheet36.cls` | 8 | 0 |
| `Sheet1.cls` | 8 | 0 |
| `Sheet38.cls` | 8 | 0 |
| `Sheet23.cls` | 8 | 0 |
| `Sheet10.cls` | 8 | 0 |
| `Sheet7.cls` | 8 | 0 |
| `ProcessReports.bas` | 95 | 3 |
| `Sheet2.cls` | 8 | 0 |
| `Sheet4.cls` | 8 | 0 |
| `Sheet14.cls` | 8 | 0 |
| `Sheet35.cls` | 8 | 0 |
| `Sheet24.cls` | 8 | 0 |
| `Sheet6.cls` | 8 | 0 |
| `Sheet9.cls` | 8 | 0 |
| `Sheet5.cls` | 8 | 0 |
| `Sheet32.cls` | 8 | 0 |
| `Sheet33.cls` | 8 | 0 |
| `Sheet40.cls` | 8 | 0 |
| `Sheet3.cls` | 8 | 0 |
| `Module1.bas` | 2 | 0 |
| `Sheet8.cls` | 8 | 0 |

### Demo Project 1.xlsm

25 tabs; 253 formula cells; 26 extracted VBA modules.

| Tab | Visibility | Formula cells | Formula functions observed | Excel charts |
|---|---|---:|---|---:|
| `Welcome` | visible | 0 | — | 0 |
| `Current Parameters` | visible | 1 | — | 0 |
| `FigureDataImport` | visible | 5 | CONCATENATE | 0 |
| `Report Numbers` | visible | 11 | CONCATENATE, IF, SUM | 0 |
| `Report Loop List` | visible | 1 | — | 0 |
| `ChartRanges` | visible | 25 | CONCATENATE, COUNTA, COUNTIF, LOOKUP | 0 |
| `Table 1` | visible | 28 | CONCATENATE, IF, TEXT | 0 |
| `Table 2` | visible | 18 | — | 0 |
| `Table 3` | visible | 6 | IF | 0 |
| `Figure 1` | visible | 27 | IF, TEXT | 1 |
| `Figure 2` | visible | 18 | IF | 1 |
| `FFS MR TopX` | hidden | 64 | IF | 0 |
| `Sheet1` | hidden | 0 | — | 0 |
| `Bullet 1` | visible | 28 | IF | 0 |
| `Header and Footer` | visible | 1 | — | 0 |
| `x-Header` | hidden | 1 | — | 0 |
| `SmartSections` | visible | 12 | CONCATENATE, IF, OR, VLOOKUP | 0 |
| `Summary Sheet` | visible | 0 | — | 0 |
| `Add Docx` | visible | 2 | — | 0 |
| `Date` | visible | 5 | DAY, MONTH, TODAY, YEAR | 0 |
| `GlobalReplace` | visible | 0 | — | 0 |
| `index` | visible | 0 | — | 0 |
| `index_altered` | visible | 0 | — | 0 |
| `index_temp` | visible | 0 | — | 0 |
| `CSVIMPORTS` | visible | 0 | — | 0 |

VBA modules (including sheet/document stubs with no routines):

| Module | Source lines | Declared routines |
|---|---:|---:|
| `ThisWorkbook.cls` | 8 | 0 |
| `Sheet11.cls` | 8 | 0 |
| `Sheet36.cls` | 8 | 0 |
| `Sheet1.cls` | 8 | 0 |
| `Sheet37.cls` | 8 | 0 |
| `Sheet23.cls` | 8 | 0 |
| `Sheet10.cls` | 8 | 0 |
| `Sheet7.cls` | 8 | 0 |
| `Sheet8.cls` | 8 | 0 |
| `Sheet2.cls` | 8 | 0 |
| `Sheet3.cls` | 8 | 0 |
| `Sheet14.cls` | 8 | 0 |
| `Sheet35.cls` | 8 | 0 |
| `Sheet24.cls` | 8 | 0 |
| `Sheet6.cls` | 8 | 0 |
| `Sheet9.cls` | 8 | 0 |
| `Sheet5.cls` | 8 | 0 |
| `Sheet32.cls` | 8 | 0 |
| `Sheet33.cls` | 8 | 0 |
| `Sheet40.cls` | 8 | 0 |
| `Sheet4.cls` | 8 | 0 |
| `Sheet17.cls` | 8 | 0 |
| `Sheet12.cls` | 8 | 0 |
| `Sheet13.cls` | 8 | 0 |
| `TRW.bas` | 30 | 3 |
| `Sheet16.cls` | 8 | 0 |

## Current Parameters contract

Coordinates are taken from the original VB.NET accesses. Unsupported options
must be reported before generation; silently ignoring an On flag is not parity.

| Cell | Purpose | Original consumer | Python status |
|---|---|---|---|
| A2 / A5 | Word template folder / filename | RunReport1 | Implemented portable workbook folder + configured filename |
| A8 | Output destination | RunReport1, chart/HTML temp files | Implemented output override; unique nonoverwriting files |
| A11 | Current report ID | LoopReports, workbook formulas | Implemented per-ID workbook recalculation |
| B15 | Smart Sections On/Off | Cutsections | Implemented; cross-table marker limit remains |
| B17 | Headers and Footers On/Off | Changeheader | Implemented five section texts/alignments; variants partial |
| B19 | Auto-Import Data Files On/Off | CommonCore → CallFigures | Implemented CSV/XLSX imports with fresh formulas |
| B21 | 508 Compliant Charts On/Off | AssignAlt | Override/default formatted alt text implemented; no full 508 certification |
| B23 | DOCX / DOC / PDF format | RunReport1 | DOCX/PDF/legacy DOC implemented; LibreOffice needed for DOC/PDF |
| B25 | Table of Contents Update On/Off | RunReport1 → TablesOfContents(1).Update | Automatic heading TOC/PAGEREF cache from rendered pages; custom TOCs/other fields partial |
| B27 | Report title/file suffix | RunReport1 | Implemented |
| B29 | Massive Table Maker On/Off | CommonCore, MassTabMake | Implemented modes 1/2/3 |
| B31 | Mass table substitution identifier | CommonCore | Implemented report-ID filename substitution |
| B33 | User Defined Macros On/Off | Excel.Application.Run("TRW.AddMacros") | Deferred by user; On flag warns |
| B35 | Demo Mode On/Off (Word visibility) | Runonestate / LoopReports | Progress events; On flag warns that no live Word window is displayed |
| B37 | Global Replace On/Off | RunReport1 → GlobalReplace | Implemented contiguous literal pairs |
| A36 / A39 | Old kernel location mode / path | ProcessReports workbook VBA | Old handoff protocol pending; direct CLI/GUI available |

Report names are derived from template filename before `_Template`, report ID,
suffix, and selected extension. The PDF branch also uses document title metadata
in an intermediate filename. Workbook naming is implemented; PDFs use the same configured output name, avoiding a shared title-named temporary file.

## FigureDataImport and Report Numbers schemas

| Sheet / columns | Original contract |
|---|---|
| FigureDataImport A2 | Base input folder |
| FigureDataImport E / F / G | Enabled-row marker/name, destination sheet, destination starting cell |
| FigureDataImport H / I | Rows and columns to clear/copy, including blank source cells |
| FigureDataImport J / K | Input filename, source starting cell |
| FigureDataImport L | Legacy bullet transform selector 1/2/3; caller currently commented in CallFigures |
| Report Numbers A / B / C | Text tag ID, description, value with cell number format |
| Report Numbers G / H / I / J / K / L | Chart tag ID, description, worksheet, chart object name, data range, alternative-text override column |
| Report Numbers N / O / P / Q | Table tag ID, description, worksheet, range |
| Report Numbers R / S | Fit-to-window / centered alignment flags |
| Report Numbers T / U | Existing Word table shell flag; Variable Row / Fixed Row / Table Convert mode |
| Report Numbers V | Number of repeating header rows |
| Report Numbers W / X | Accessibility title / description |
| Report Numbers Y / Z | Padding in pixels / side selector letters (L,R,T in implementation) |

ChartRanges A/B/C/D/F supply element ID, sheet, column count, row count, and
range. Sentinel `xx` formulas exclude unused rows. SmartSections has separate
include/reason fields (A:D,G), table renumbering (J:L), and figure renumbering
(P:R). Add Docx uses D2 as base folder and A/B starting at row 5 as section
tag/filename. The mass index uses A filename, B range, C tag ID, E mode, and
F format string; CommonCore copies A:H, including fields not read by Main.

Fixed scan limits in the source (imports rows 4–100, chart rows up to 1000,
several range lookups up to 100/200, headers up to five sections) are part of
the audit. Increasing limits is a deliberate improvement, not an untested parity claim.

## Embedded macro routes and uncertainties

- Blank Shell and Finished Example each contain ProcessJobOne, ProcessJobMany,
  and ShellMe. They save a handoff workbook and launch the old external kernel.
  They are separate from the VSTO ribbon handlers; neither route is assumed unused.
- Demo Project 1 contains TRW.test, TRW.AddMacros, and TRW.chartidx.
  AddMacros is empty apart from a commented test call; preserving the extension
  execution is explicitly deferred by the user. chartidx is a workbook diagnostic in addition to
  the VSTO Chart ID button.
- The simple workbooks do not contain a TRW.AddMacros procedure. Their B33
  option is Off; the advanced demo has it On. VBA is bypassed with a warning. Enabling it in a workbook without
  that macro does not invoke Excel; no boutique query result is claimed.
- Old Changeheader2 references a `Header` sheet, absent from all supplied
  workbooks. `x-Header` exists but is a different name. Keep the old function
  flagged for compatibility review rather than pretending those names match.
- QueryTables uses CSVIMPORTS / CSVImports (Excel names are case-insensitive).
  The Python adapter normalizes sheet matching consistently.

## Word templates and fields

| Document in supplied package | Tags | Field instruction fragments | External relationships | Embedded media |
|---|---:|---:|---:|---:|
| `Report Writer Example Project/Client Supplied Template Document.docx` | 0 | 5 | 0 | 0 |
| `Report Writer Example Project (Finished)/Client Supplied Template Document.docx` | 0 | 5 | 0 | 0 |
| `Report Writer Example Project (Finished)/Example Project_Template.docx` | 6 | 5 | 0 | 0 |
| `Report Writer Example Project (Finished)/Final Reports/Example Project_A.docx` | 0 | 5 | 0 | 2 |
| `Report Writer Example Project (Finished)/Final Reports/Example Project_B.docx` | 0 | 5 | 0 | 2 |
| `Report Writer Template/Additional Docs/Add this A.docx` | 0 | 0 | 0 | 0 |
| `Report Writer Template/Demo Project 1_Template.docx` | 12 | 5 | 4 | 0 |
| `Report Writer Template/Final Reports/Demo Project 1_A.docx` | 0 | 5 | 4 | 2 |
| `Report Writer Template/Final Reports/Demo Project 1_B.docx` | 0 | 5 | 4 | 1 |

The example Word template contains a TOC field and four PAGEREF fragments;
it has no external relationships. The generic Word warning seen by the user
therefore does not by itself establish an external dependency. The advanced
template has four external relationships; their presence is recorded, but
no external resources were fetched or field targets executed.

Word field-instruction fragments and template tag lists are preserved in
legacy-inventory.json. Layout lives in the templates, but replacement content
can alter pagination. Heading TOCs are now rebuilt automatically with LibreOffice layout; custom fields remain limited.
