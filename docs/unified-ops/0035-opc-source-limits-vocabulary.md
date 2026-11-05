# 0035 — OPC and OOXML physical source-limit vocabulary

OPC now exports `SourceLimits` and `SourceLimitsBuilder`. The retired
`ReadLimits` and `ReadLimitsBuilder` names have no compatibility aliases.
DOCX, PPTX, XLSX and XLSB re-export the renamed physical profile; the unified
operations layer, facade, tests, examples and active performance tools use it.
The README and public API documentation show the new names.

This is the shared physical-profile rename required by ADR 0034 §5. The rename
preserves default values, checked builder setters, error behavior, ingress
plumbing, cache policy and execution admission. A separate no-op publication
regression found by broader testing is corrected below. Existing limit-boundary and ingress tests
exercise the renamed profile. Domain-local types named ReadLimits in XLSB,
ODF, Markdown and iWork are distinct types and remain for their own migrations.
Historical measurement source captures and numbered records retain their
original vocabulary so they still describe their measured revisions.

## Enforcement

Minimal and full rustdoc profiles now check OPC, PPTX, XLSX and XLSB, alongside
the existing core/DOCX profiles. DOCX also gains both retired-name bans.

A ban may set `recursive: false` to match direct children of its scope. XLSB
uses this for its package-level re-export, preserving unrelated nested semantic
ReadLimits. Recursive prefix matching remains the default. Policy parsing rejects
non-boolean values and unknown fields. Real-rustdoc tests distinguish nested
names from direct exports and enum variants, including an identical root/nested
ReadLimits spelling. The new XLSX/XLSB profiles exposed an external glob that
rustdoc cannot inspect; their shape facades now explicitly re-export the same
40 items instead. The checker continues to refuse unresolved external globs.

## Verification

- OPC/DOCX/PPTX/XLSX/XLSB/operations tests: 8,013 distinct tests passed,
  four existing ignored. Counts: DOCX 2,019; OPC 1,095; operations 388;
  PPTX 1,318; XLSB 1,031; XLSX 2,162. OPC and XLSX were rerun after the
  empty-plan fix. The already-passed slow XLSX worksheet corpus oracle was
  filtered from that rerun; the other 2,161 XLSX tests passed together.
- Workspace all-feature library compilation, including facade and Python:
  passed.
- Warning-denied OPC/DOCX/PPTX/XLSX/XLSB/operations library Clippy: passed;
  XLSX/XLSB checked again after replacing their external shape globs.
- Warning-denied OPC/DOCX/PPTX/XLSX/XLSB doctests: 115 passed,
  44 existing ignored.
- Glossary: 11 feature profiles, 42 banned-name rules, zero complete
  transaction families; passed. Checker unit tests: 21 passed.
- Crate boundaries: 66 packages, 267 internal dependency declarations,
  11 existing explicit debt items; passed.
- Both active performance-tool packages compile, including perf-baseline bins.
- Changed OPC/DOCX fuzz targets compile. The unchanged Numbers table-merges
  facade fuzz target also compiles. These are compilation checks, not fuzz runs.
- Changed Rust-file formatting and diff whitespace checks: passed.
- Independent read-only review confirmed the migration scope and all 40
  explicit shape exports; no concrete regression found.
- Rewritten base remains 32ad64e320 and is an ancestor of the branch.

Compiling the OPC fuzz targets also exposed a pre-existing metadata/payload
API mismatch: the harness called blob() on PartMetadata. It now explicitly
requests the payload with fallible get_part() before observing its length.
A successful catalog admission is not treated as proof that lazy payload
materialization succeeds.

The exploratory workspace all-features/all-targets check found an existing
unrelated error in `crates/litchi-iwa/examples/create_numbers_donut_chart.rs:68`:
`donut_data()` declares `litchi_iwa::Result<ChartData>` but directly returns
`Result<ChartData, DataError>`. The file and the relevant iWork code are unchanged
by this migration. The same broad check reports existing iWork deprecation and
unused-import warnings. This is not evidence of a green workspace all-targets
check; library and affected-target validation are reported separately.

## Empty-plan publication budget regression

The full XLSX run exposed a regression from retained topology introduction:
`write_topology_to_stream()` allocated and admitted a PreparedTopology owner
before publishing even an empty plan. The existing managed multi-sheet exact
no-op test supplies selected payload bytes plus the 64 KiB exact-copy window;
the extra 144-byte owner made that call exceed its budget (66,679 observed
versus 66,535 allowed).

The one-shot empty-plan route now disables speculative read-ahead and calls the
existing checked exact-source writer directly. It retains source, cancellation,
output accounting and partial-output checks, with no needless retained-candidate
allocation. Changed plans still use the retained planner. Explicit callers of
prepare_topology continue to pay for the owner they retain. The regression test
keeps its original budget and assertions; no limit was raised to make it pass.
It also passes with OPC signing enabled. A separate read-only review confirmed
that source/context, signed-source, partial-output and final-source checks
remain in the exact-copy path.

## Remaining work

The physical limit kind and error names (`ReadResource`, `ReadLimit`,
`InvalidReadLimit`), dimension vocabulary, OpenOptions envelopes and semantic
limits still need migration. ODF/legacy/iWork type renames, package/projection
and transaction shapes, other unified milestones and full acceptance remain
open. This record claims neither complete limits-policy conformance nor a
performance improvement.

## Cleanup

Removed workspace build output, generated root/fuzz Cargo.lock files, temporary check logs
and the borrowed SVG test fixtures. No worktree was created. Base and iWork
sources and pre-existing worktrees were preserved.

## Borrowed verification fixtures

Copied from the base checkout to the identical relative path for testing, then
removed. Reproduce with `cp /home/zhuhe/code/litchi/<path> <path>` after creating
the destination parent. None is committed.

| Path | Bytes | SHA-256 |
|---|---:|---|
| `3rdparty/Open-XML-SDK/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestFiles/svg.docx` | 74104 | `1ce06ef88f89c5e14d1d22cd8e4b3a5f70951e7bd53e19053d4942873333d21d` |
| `3rdparty/libreoffice-core/sw/qa/extras/ooxmlexport/data/tdf164835_nonDummyLineHeight.docx` | 302210 | `569572dc2ec5dfee9334504ae331482ebb8188b80a0d3c93d44b8884cda29283` |
| `3rdparty/libreoffice-core/sc/qa/unit/data/xlsx/tdf169496_hidden_graphic.xlsx` | 12470 | `0b647da300a085f39914fdfae961463ae9e54ffe772b2e0eb9860a841ab93f72` |
| `3rdparty/libreoffice-core/sd/qa/unit/data/pptx/tdf163852.pptx` | 9579 | `2525665a1e430a9ee543faf456424bd7b994cce0a2d76470a5a07453e3e2ab39` |
| `3rdparty/libreoffice-core/sd/qa/unit/data/pptx/tdf164622.pptx` | 18696 | `33b86388804554e35cb3b07e81cf5db5cb6bb02e76716c71f233b57325c2d199` |
| `3rdparty/libreoffice-core/sd/qa/unit/data/pptx/tdf169496_hidden_graphic.pptx` | 35202 | `f717ca4657257b538466c772eca2a4d4886cb5b919c2ad5f044377fcee124fc8` |
