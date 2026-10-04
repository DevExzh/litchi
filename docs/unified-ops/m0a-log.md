# M0a step log (merged into PROGRESS.md when M0a lands)

## S2 — unified error (`7010055311`)
- Added `litchi_ops::error`: 39 codes (36 OfficeCLI + `unsupported_capability`, `limit_exceeded`, `cancelled`), a pointer-sized `Error`, a typed `Source`, and per-leaf classifiers for core/io/execution/opc/cfb/docx/pptx/xlsx/xlsb/doc/xls/ppt/rtf. 27 tests.
- 10 optional leaf edges registered (ADR 0024 amendment: 255 declarations / 244 edges). No iWork edges.
- Follow-ups:
  - (S6/S7) The leaf message is both the `Display` text and `source()`, so error-chain printers show it twice. Make `Display` facade-level only.
  - (M2) Add a diagnostics slot.
  - Add `Source::Crypto` with the `encryption` feature.
  - The design's code list has 36 entries, not the "37 + 1" that plan 0003 states.

## S3 — OpenOptions, Locked/Signed shells (`b9951ae15f`)
- `OpenOptions` (const `new()` = `Default`; common knobs plus feature-gated per-format limit sub-options; no credentials), `SourceMode`, `Family::of` (all 21 formats), and the pointer-sized `Locked`/`Signed` handles with redacted `Debug`. No unlock or unsign yet (S6/S6b).
- Merge with S2: the dependency and feature lists were combined. The only new edge is `litchi-ops -> litchi-odf-common`, which brings the inventory to 256 declarations and 245 canonical edges. The ADR 0024 amendments were reconciled. The boundary checker, the non-iWork gate and the litchi-ops clippy/tests pass on the merged tree.
