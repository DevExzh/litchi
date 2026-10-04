# M0a step log (merged into PROGRESS.md when M0a lands)

## S2 — unified error (`7010055311`)
- Added `litchi_ops::error`: 39 codes (36 OfficeCLI + `unsupported_capability`, `limit_exceeded`, `cancelled`), a pointer-sized `Error`, a typed `Source`, and per-leaf classifiers for core/io/execution/opc/cfb/docx/pptx/xlsx/xlsb/doc/xls/ppt/rtf. 27 tests.
- 10 optional leaf edges registered (ADR 0024 amendment: 255 declarations / 244 edges). No iWork edges.
- Follow-ups:
  - (S6/S7) The leaf message is both the `Display` text and `source()`, so error-chain printers show it twice. Make `Display` facade-level only.
  - (M2) Add a diagnostics slot.
  - Add `Source::Crypto` with the `encryption` feature.
  - The design's code list has 36 entries, not the "37 + 1" that plan 0003 states.
