# Native XLSB Data Model fixture inventory

The bounded local inventory considered 35 XLSB paths under `3rdparty/` and `test-data/`. All 34 inspectable ZIP workbooks contained `xl/workbook.bin`; all 713 workbook records decoded completely, with zero occurrences of the Data Model record family. No native XLSB Data Model owner fixture was established in these inputs. The encrypted CFB input was not decoded, and no conclusion is made about its contents. Eight framing errors occurred in other binary entries (printer settings/VBA); no entries were skipped for resource bounds.

This is fixture-coverage evidence, not a native producer compatibility result for the XLSB model API. The separately bound native XLSX model fixture is a different format. Model API validation remains synthetic until an appropriate native XLSB fixture is available.

Version 3 replaces a hardcoded workbook hit count in the companion coverage summary with actual aggregation. Its positive control recognizes explicit bytes `C9 10 00` as record ID 2121 and checks seven additional framing cases. Root independently reran coverage and the positive control, compared nine resulting files, and rehashed all 35 source archives and every archived artifact. Inputs remain external; the bundle contains scanners, source manifests, reports, specification notes, and prior receipts retained for traceability. Use the latest `RECEIPT.md` and version-3 coverage outputs for conclusions.

The archive is deterministic and the receipt binds each member, the archive itself, source-manifest digest, and root verification results. Archive/entry/aggregate scan ceilings are 128 MiB/128 MiB/512 MiB. This batch changes no runtime code and makes no performance claim.
