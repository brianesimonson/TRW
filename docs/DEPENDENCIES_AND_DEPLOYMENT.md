# Dependencies, deployment, and explicit replacement decisions

This is a source/configuration audit, not a claim that Windows dependencies were
installed or the VSTO project rebuilt. The uploaded binary's correspondence to
the readable source has not been proven by a reproducible build.

| Dependency / artifact | Original role / evidence | Python state / disposition |
|---|---|---|
| VB.NET / .NET Framework 4.8 | Project targets v4.8 with late-bound code and Option Explicit/Strict disabled in several modules | New engine is Python 3.11+; syntax translation alone cannot replace Office's behavior |
| VSTO and Office Tools runtime | Excel host, ribbon, Globals.ThisAddIn, generated task panes/factory/cached-data hooks | Host-specific ledger; preserve if retaining ribbon, add explicit data bridge; no bridge currently |
| Excel Interop / Office COM | Workbook mappings, formulas, named charts, QueryTables, styles, VBA, active workbook | openpyxl workbook model + explicit formula interpreter; 641 original formula cells checked; arbitrary Excel/VBA not covered |
| Word Interop | Selection/range editing, template layout, tables, fields, DOC/DOCX/PDF export | python-docx + docxcompose assembly, LibreOffice CLI layout/export + Python heading TOC; Word fidelity remains partial |
| ADODB + ACE OLEDB 12.0 | AccessX reads/writes TRWlogs.accdb: UserAccess, StrictAccess, UserLogs, IHTNGU | Excluded by user; future authentication to be designed separately |
| System.Security.Cryptography | EncryptX license AES/PBKDF derivation, Unicode/Base64, month/year dependence | Excluded with legacy Access licensing; key material not published |
| MySql.Data 8.0.11 | ModuleConnection test/schema-enumeration module | Review candidate; no inbound call to TestConnection found in compiled source; not established unused in all deployments |
| Google.Protobuf 3.5.1 | packages.config/project reference, included in deployment | No direct source use found; inspect MySql dependency relationship and retained features before removing |
| WinForms / System.Drawing | Ribbon message dialogs and progress form; assorted imports | CLI and Tk desktop/progress implementation; Windows UI/launcher/installer unverified |
| Win32 user32 clipboard APIs | Clipboard-clear helper alternatives | Replaced mechanism: direct DOCX operations and owned temporary files |
| COM garbage collection/release | Excel/Word lifecycle and import cleanup | Replaced by scoped workbook files/temporary resources; no COM in engine |
| Scripting.FileSystemObject | HTML table rewriting in FindLines / FindLines2 | Replace transport with DOCX table formatting; native style/merge/padding support implemented; exact HTML/conditional presentation remains partial |
| Report Writer Kernel.exe | ProcessReports VBA writes TRWHubToKernel.xlsx then launches kernel | Absent from supplied package; old handoff protocol unported; direct CLI/GUI single/batch routes exist |
| Crypto Obfuscator build targets | Project imports workstation-specific LogicNP target path | Legacy build dependency; no Python equivalent needed for execution; deployment/IP strategy must be explicit |
| ClickOnce VSTO manifests / setup.exe | Original installed add-in delivery | No Python desktop installer yet; RunReportWriter.cmd bootstraps venv/pip and Tk UI; Windows testing and standalone installer pending |
| Setup .vdproj files / MSI artifacts | Additional legacy installer projects/artifacts in code archive | Inventory as packaging artifacts; not callable VB functions, not ported |
| Temporary signing PFX / obfuscation maps | Original publication/signing and debug artifacts | Preserved only in original archive; not uploaded to the prototype Git repository or documentation |

The deployed `Report Writer 4.0.dll.deploy` in the package has SHA-256
`35e07b6826a6a328b858c88bdee6cb03d1bac66d88f8c52a450f6aa4a9902655`.
It matches the main original project's bin/Debug DLL and ClickOnce publish copy.
Older alternative obfuscator outputs differ. This establishes byte identity,
not whether the deployed binary is obfuscated or reproducible from source.

Current runtime dependencies are pinned in pyproject.toml: python-docx 1.2.0,
Matplotlib 3.10.8, openpyxl 3.1.5, docxcompose 2.2.0 and pypdf 6.10.0. Installation
and pip check pass. LibreOffice CLI is wired into finalization and renderer tests;
it is required for automatic TOC refresh and DOC/PDF output, not basic DOCX assembly.
No C++, .NET or Office Interop component is called by the Python engine. Linux
validation used Python 3.12; supported minimum is 3.11, which the user used for
the earlier prototype. New Windows behavior is not yet verified.

## Deliberate implementation differences

- No Access policy, legacy encryption or global Word-process killing, per user scope.
- Output files refuse overwrite. A locked output gets an error; completed earlier
  reports remain if a later batch item fails.
- Unique per-job resources and bounded renderer/file waits replace shared files,
  clipboard retries and unbounded process polling.
- Native DOCX tables replace Excel HTML transport. Chart PNGs are generated by
  Matplotlib. Their output fidelity limits are recorded, not hidden.
- Address helpers accept valid modern Excel columns and single-cell rows, fixing
  legacy sentinel/row-zero behavior instead of reproducing those defects.
- Python errors/progress/desktop messages replace Word/Excel selection and dialogs;
  Demo Mode does not open a live Word window and emits a warning when requested.
- LibreOffice determines pagination and export. Heading TOCs are calculated;
  arbitrary Word fields and pixel-identical Word rendering remain unverified.

No VSTO rebuild, signed Python installer or fresh-machine Windows installation has
been performed. Keep the current VSTO deployment available until user acceptance.
