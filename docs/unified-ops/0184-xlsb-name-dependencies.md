# 0184 — XLSB name dependencies

## Contract

Changed XLSB sheet names now update modeled textual dependencies alongside
BundleSh metadata. A shared OOXML lexical scanner, extracted from the existing
XLSX implementation, rewrites recognized formula/reference carriers and
extended-property sheet titles. Each format retains its own name-equivalence
policy. Unrelated bytes, external workbook references and formula string
literals remain unchanged. Unknown matching carriers and markup-compatibility
alternatives receive typed refusal before publication.

The scanner uses namespace identities, bounded depth and captured text, and
ordered byte-range replacements. It does not serialize whole chart models.
Native XLSB stages dependency replacements on a disposable candidate before
retained-policy validation and publication. Exact no-ops skip dependency
rewrites; failed edits preserve the caller's source snapshot.

## Durable compatibility

New name operations use tag 7 with the existing catalog-position/name payload.
Historical tags 1 (worksheet ordinal) and 6 (catalog position) keep their
BundleSh-only replay behavior. Tag 1 still normalizes its target against the
before image and re-encodes as tag 6. Merge selection identity is shared across
all encodings, while replay retains the operation's dependency policy.
Historical patches are tested with actual metadata-only after images, not
just by changing a new operation's tag.

## Verification

The combined shared-OOXML/native-XLSB suites pass 1,609 tests/doctests,
the full XLSX suite passes 2,398, and the combined DOCX/XLSX/PPTX/XLSB/encryption
facade suite passes 1,237. The final focused native suite passes 20 tests,
including quoted/escaped name dependencies, exact unrelated payloads,
historical metadata-only after images, and durable replay/inverse. Facade
regressions exercise source/eager owners, typed source errors, unchanged-name
no-ops and atomic/best-effort refusal behavior.

The boundary audit passes for 66 packages and 268 internal declarations with
11 existing debt items. Generated capability goldens and the OfficeCLI delta
include the two specific unsupported XML-carrier cases. Repository lockfiles
remain unchanged; verification uses an external copied lockfile with cached
cc 1.6.0. One missing local XLSX test fixture was copied temporarily from the
base checkout and removed afterward. No base-only commits were pending.

Delegated work stopped on an agent usage limit; the coordinator completed and
verified the saved patches. All four temporary worktrees were removed after
integration. Python/CLI/parser-fuzz acceptance is not claimed by this record.
The full all-format/Python/CLI goal, structural verbs and general name-command
coalescing remain open.


Warnings-denied full and minimal all-target Clippy, affected-crate rustdoc,
the shared helper's runnable example, four final XLSX scanner tests, four final
facade dependency/error tests, scoped formatting and generated parity checks
pass. XML scan limits now retain their typed source through both leaf adapters
and map to `limit_exceeded` in the facade. The compatibility rules for ordinary
XLSX renames and blocked carriers remain unchanged.

This closes modeled XML rename dependencies, not opaque binary or macro
reference editing. External-link parts are excluded. General structural CRUD,
name-command batching and full acceptance remain separate work.


## Paired measurements

The unchanged single-name harness ran 1, 10, 100 and 1,000 stored cells on one
worksheet, eager/source owners, nine ABBA rounds, CPU 0, with no concurrent
build/test/boundary audit. Source promotion is timed; opening, prepared inputs
and final serialization are excluded. Every case has byte-exact native/facade
output equivalence and an empty failure matrix. Receipts retain source-file,
binary and build provenance plus individual samples:
[before](results/0184-xlsb-sheet-name-before.json),
[after](results/0184-xlsb-sheet-name-after.json),
[repeat](results/0184-xlsb-sheet-name-repeat.json).

| Owner / cells | Before native / facade µs | After native / facade µs | Repeat native / facade µs |
|---|---:|---:|---:|
| eager / 1 | 292.632 / 305.221 | 372.201 / 391.951 | 376.122 / 385.341 |
| source / 1 | 421.992 / 434.462 | 514.483 / 526.983 | 524.402 / 532.553 |
| eager / 10 | 292.911 / 301.861 | 376.212 / 393.332 | 378.992 / 396.571 |
| source / 10 | 438.102 / 446.242 | 528.363 / 539.283 | 534.122 / 546.932 |
| eager / 100 | 355.102 / 379.312 | 443.033 / 458.432 | 443.412 / 463.792 |
| source / 100 | 513.323 / 524.602 | 612.882 / 623.722 | 614.743 / 622.183 |
| eager / 1000 | 1117.155 / 1135.645 | 1202.146 / 1229.455 | 1209.616 / 1240.576 |
| source / 1000 | 1409.666 / 1424.486 | 1512.747 / 1525.967 | 1525.807 / 1553.997 |

After/repeat facade overhead is 0.9–5.4%, within the scoped 10% single-edit bar.
The native route itself gains roughly 80–120 µs of necessary XML dependency
work (about 20–29% on the smallest packages). The before route left app titles
stale; this is a correctness cost, not a speed improvement. No batch scaling,
XLSX performance improvement or full-goal acceptance is claimed. Temporary
build and measurement artifacts were removed after retaining these receipts.
