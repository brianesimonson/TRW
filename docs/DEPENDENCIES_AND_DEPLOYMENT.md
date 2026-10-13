# Dependencies, deployment, and explicit replacement decisions

This is a source/configuration audit, not a claim that Windows dependencies were
installed or the VSTO project rebuilt. The uploaded binary's correspondence to
the readable source has not been proven by a reproducible build.

| Dependency / artifact | Original role / evidence | Python state / disposition |
|---|---|---|
| VB.NET / .NET Framework 4.8 | Project targets v4.8 with late-bound code and Option Explicit/Strict disabled in several modules | New engine is Python 3.11+; syntax translation alone cannot replace Office's behavior |
| VSTO and Office Tools runtime | Excel host, ribbon, Globals.ThisAddIn, generated task panes/factory/cached-data hooks | Host-specific ledger; preserve if retaining ribbon, add explicit data bridge; no bridge currently |
| Excel Interop / Office COM | Workbook mappings, formulas, named charts, QueryTables, styles, VBA, active workbook | H02; no automatic workbook import/calculation adapter yet |
| Word Interop | Selection/range editing, template layout, tables, fields, DOC/DOCX/PDF export | Partial replacement by python-docx; finalization/rendering remains H01/H05 |
| ADODB + ACE OLEDB 12.0 | AccessX reads/writes TRWlogs.accdb: UserAccess, StrictAccess, UserLogs, IHTNGU | H03; schema inferred from code only; database data/crypto parity unverified |
| System.Security.Cryptography | EncryptX license AES/PBKDF derivation, Unicode/Base64, month/year dependence | H03; no compatible Python license layer; key material is not published |
| MySql.Data 8.0.11 | ModuleConnection test/schema-enumeration module | Review candidate; no inbound call to TestConnection found in compiled source; not established unused in all deployments |
| Google.Protobuf 3.5.1 | packages.config/project reference, included in deployment | No direct source use found; inspect MySql dependency relationship and retained features before removing |
| WinForms / System.Drawing | Ribbon message dialogs and progress form; assorted imports | CLI only today; UI/progress packaging remains H04 |
| Win32 user32 clipboard APIs | Clipboard-clear helper alternatives | Replacement pending; direct DOCX assembly avoids clipboard, but final paths/cleanup need acceptance |
| COM garbage collection/release | Excel/Word lifecycle and import cleanup | Replacement pending; do not transplant forced COM cleanup into pure Python |
| Scripting.FileSystemObject | HTML table rewriting in FindLines / FindLines2 | Replace transport with DOCX table formatting; HTML fidelity/padding not currently reproduced |
| Report Writer Kernel.exe | ProcessReports VBA writes TRWHubToKernel.xlsx then launches kernel | Absent from supplied package; old macro route held H04, separate from VSTO ribbon |
| Crypto Obfuscator build targets | Project imports workstation-specific LogicNP target path | Legacy build dependency; no Python equivalent needed for execution; deployment/IP strategy must be explicit |
| ClickOnce VSTO manifests / setup.exe | Original installed add-in delivery | No Python desktop installer yet; downloaded source needs venv/pip today |
| Setup .vdproj files / MSI artifacts | Additional legacy installer projects/artifacts in code archive | Inventory as packaging artifacts; not callable VB functions, not ported |
| Temporary signing PFX / obfuscation maps | Original publication/signing and debug artifacts | Preserved only in original archive; not uploaded to the prototype Git repository or documentation |

The deployed `Report Writer 4.0.dll.deploy` in the package has SHA-256
`35e07b6826a6a328b858c88bdee6cb03d1bac66d88f8c52a450f6aa4a9902655`.
It matches the main original project's bin/Debug DLL and ClickOnce publish copy.
Older alternative obfuscator outputs differ. This establishes byte identity,
not whether the deployed binary is obfuscated or reproducible from source.

Current engine dependencies are `python-docx==1.2.0` and `matplotlib==3.10.8`.
Their installation and four tests were checked. LibreOffice was used only for
a compatibility export; it is not wired into the engine. There is no C++, .NET,
Office Interop, or hidden renderer used by the Python report generation command.

## Explicit application differences to resolve

- Original asks to kill every Word process before generation. A new engine should
  own only its own resources. This is a proposed deliberate replacement, not a
  reason to kill user documents from Python.
- Original can delete/replace output files. Prototype refuses overwrite. Define
  an explicit overwrite policy and predictable behavior when a document is open.
- Original communicates status via progress form/dialogs. Prototype prints JSON
  and exits nonzero on selected errors. User-facing UI and progress remain missing.
- Original uses a fixed temporary Word filename, selection state, scratch images,
  HTML exports, and several wait/retry helpers. Prefer unique per-job resources,
  bounded waits, and deterministic edits, while preserving report output.
- Original uses Excel's implicit type coercion/number formats and culture-dependent
  VB formatting. Python float/string conversion is not a complete substitute.
- Original calls license/permission/logging before reports. Prototype is a demo
  engine without these gates; it must not be represented as a licensed production
  replacement until H03 is closed or an explicit authorized policy change occurs.

These differences must receive a documented compatibility or retirement decision
in the function ledger before full application acceptance.

`AccessX.SetConnstr` assigns the same workbook-local TRWlogs.accdb path to both
candidate variables. A workstation-specific alternative is commented out, not
an active parent-directory fallback. Source intent and actual behavior must be
distinguished when defining the replacement's database lookup rules.
