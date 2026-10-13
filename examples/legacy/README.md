# Original workbook fixtures and corrected example

These fixtures come from the user-supplied deployment package. They contain only
reporting workbooks, Word templates/additional documents and CSV inputs. Access
records, installer binaries, signing material and obfuscation maps are omitted.
VBA bytes remain embedded in XLSM fixtures for provenance, but the engine never
executes them.

- `advanced`: original Demo Project 1; all reporting fixtures are unchanged.
- `finished`: original simple finished example. It selects Company A filenames
  for both A/B IDs; the engine preserves that configuration and totals 30 in both.
- `blank`: original blank shell workbook used to audit its formulas; missing
  templates/inputs make it a setup shell rather than a ready report project.
- `advanced-corrected`: user-approved separate corrected example, preserving the
  originals. Female wording is corrected; Table03 targets the actual merged
  Table3 sheet/range, with two repeating header rows and native-table mode. The
  incompatible two-column Word shell becomes a body tag and its caption changes
  to gender counts. Table3 presentation values link to imported Figure2 values
  so Company B does not reuse stale Company A data. B25 is On for automatic TOC
  refresh. Other options, including the deferred VBA flag, remain as supplied.

The runtime sets project/output paths in memory to the selected folder and output
argument, so original machine paths do not need to exist. Index Results and
Additional Docs paths are relocated. Inputs outside this portable structure must
exist or be explicitly reconfigured. Workbooks and templates are not overwritten;
Import only writes a separate prepared workbook.

New reports and verification notes are in `sample-output/v0.2`. The original
advanced configuration produces duplicate age tables for Table03; the corrected
copy demonstrates the otherwise hidden merged-cell conversion branch.
