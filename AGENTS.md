# Report Writer conversion requirements

The user's scope is the entire original Report Writer application, not just the
Hello World example. Preserve the Word-template-driven workflow and account for
every original routine, workbook tab/formula contract, and used feature.

Read `docs/CONVERSION_AUDIT.md`, `docs/FUNCTION_LEDGER.md`, and
`docs/WORKBOOK_TABS.md` before implementation. The current Python engine is a
workbook-driven replacement with explicit parity gaps. There is no C++ component. Keep source and binary provenance
separate; uploaded archives are preserved outside the Git checkout.

- Update the conversion ledger in the same change as feature work. Maintain
  reviewed descriptions/statuses in `docs/conversion-updates.json` and the generator, regenerate
  the Markdown/CSV/JSON, and update the scope summary and acceptance evidence.
- Do not label a routine Complete because the demo passes. Verify its original
  behavior contract and all relevant branches; record deliberate differences.
- Keep unsupported functionality explicit. Name holds and their dependencies;
  never silently ignore workbook flags, unsupported formulas, or VBA hooks.
- Do not declare code unused solely from sample coverage or lexical call scans.
  Candidate older/diagnostic/generated routines require documented retirement
  decisions; retain them in the inventory.
- Automatic TOC/page-number refresh is required when enabled. Setting updateFields
  or asking the user to refresh manually is not a complete replacement. Document
  the rendering mode and verify final field values/PDF behavior.
- Do not use cached Excel values as if they were freshly recalculated. Preserve
  template layout and workbook-driven formatting rules; compare rendered outputs
  where layout matters.
- Do not upload embedded credentials, license records/key material, signing PFX
  files, or raw obfuscation maps into the repository. Audit metadata can record
  hashes and source locations without exposing values.
- Each cloud task is isolated. Use this existing checkout; do not create Git
  worktrees unless the user explicitly requests one.

Use the `.venv` interpreter for runtime commands. No persistent service is needed
for the CLI. `python -m unittest discover -s tests -v` runs unit and renderer integration checks. LibreOffice is required for TOC/PDF/DOC validation; skipped renderer tests must not be reported as passing. The Tk UI/Windows launcher still needs platform validation.

User-approved exclusions: Access identity/license/logging/EncryptX and KillWord.
User VBA boutique queries (AddMacros) are deferred; emit a warning when requested.
Do not reintroduce them as conversion prerequisites. Originals remain unchanged;
only the separate advanced-corrected fixture contains agreed demo corrections.
