# 0183 — XLSB sheet-name selectors

## Contract

`Package::edit().set_sheet_name(selector, String)` replaces `rename_sheet`.
Selectors use the immutable complete tab catalog, matching `set_sheet_state`:
missing returns `None`; resolved edits return `Some(changed)` relative to the
current staged name. Repeated assignments replace the retained operation.
Returning to the original spelling produces an exact no-op. Changed names
are admitted against the effective staged namespace; a duplicate refuses
without modifying earlier staged operations. No compatibility alias remains.

Worksheet, chart, dialog, macro and unbound tab names share the workbook
BundleSh record. The unified writer now uses the complete tab position for
name admission and editing, including combined name/visibility assignments.
The obsolete non-worksheet-name capability gap is removed from the registry,
golden snapshot and generated OfficeCLI delta. Existing source fences, quotas,
transaction barriers, signature policy and name coercion remain in place.

Durable operation tag 6 identifies complete catalog positions. Legacy tag 1
continues to mean worksheet ordinal; decoding resolves that ordinal against
the before image and normalizes to tag 6 before replay or merge. Thus old and
new patches share conflict identity even with a chart before the worksheet.
Re-encoding an old patch emits the canonical operation. Existing operation
tags and the envelope remain unchanged.

## Limits and next work

This is a selector and metadata API migration. It retains the existing
BundleSh-only rename behavior: ancillary XML is preserved, including chart
textual formulas and cached app.xml TitlesOfParts metadata. Updating those
name-dependent values remains a correctness follow-up; this record does not
claim complete rename dependency closure. Formula token indexes, tab identities,
relationships, unrelated records and exact inverse images are preserved.

Native name admission still scans the catalog for each changed name; general
name-command coalescing and large name-batch scaling are not qualified here.
The structural audit identifies the next prerequisites: identity-preserving
catalog permutation, BookView and internal ExternSheet index remapping,
sheet-scoped defined names, then graph-aware add/remove and unified structural
verbs. Indexed opaque dependencies require explicit handling or refusal.
The full all-format/Python/CLI goal remains active.

## Verification

The signing-enabled native suite passes 1,154 tests/doctests; the combined
DOCX/XLSX/PPTX/XLSB/encryption facade suite passes 1,234. Regressions cover
complete-catalog names, staged cancellation and duplicates, all tab kinds,
source/eager combined name+visibility, atomic and best-effort rollback,
legacy chart-first ordinal decoding, merge conflict identity, and exact
forward/inverse publication. Existing malformed-source and signature tests
remain green.

Warnings-denied full-feature all-target Clippy and affected-crate rustdoc
pass. Both XLSB glossary profiles pass with 32 banned names each, including
the removed rename method. Scoped formatting and generated OfficeCLI parity
checks pass. Repository lockfiles are unchanged; verification used an external
copy with cached cc 1.6.0. No base-only commits were pending. Temporary
implementation worktrees were removed after integration. Python/CLI/parser-fuzz
acceptance and full goal completion are not claimed.


The minimal-feature all-target Clippy profile also passes. The boundary audit
passes for 66 packages / 268 internal declarations with 11 existing debt items.

## Paired measurements

The single-name harness uses 1, 10, 100 and 1,000 stored cells on one worksheet,
eager/source owners, nine ABBA rounds, and CPU 0. Opening, corpus preparation
and final serialization are outside timing; source promotion is inside the
source leg. No builds/tests/heavy audits ran during timing. Route names and
historical JSON labels are retained while the native call is migrated.
All cases retain exact ZIP equivalence, one native commit and empty failure
matrices. Source/build/binary hashes and raw samples are retained in
[before](results/0183-xlsb-sheet-name-before.json),
[after](results/0183-xlsb-sheet-name-after.json) and
[repeat](results/0183-xlsb-sheet-name-repeat.json).

| Owner / cells | Before native / facade µs | After native / facade µs | Repeat native / facade µs |
|---|---:|---:|---:|
| eager / 1 | 290.722 / 303.742 | 291.822 / 306.771 | 294.281 / 311.921 |
| source / 1 | 422.542 / 434.782 | 433.432 / 441.512 | 432.012 / 444.102 |
| eager / 10 | 297.241 / 310.431 | 291.781 / 310.592 | 294.521 / 307.771 |
| source / 10 | 438.132 / 454.822 | 437.672 / 449.942 | 441.442 / 453.232 |
| eager / 100 | 358.661 / 376.062 | 359.052 / 377.012 | 362.742 / 381.982 |
| source / 100 | 520.292 / 532.532 | 513.682 / 530.032 | 524.773 / 533.362 |
| eager / 1000 | 1120.235 / 1145.646 | 1108.775 / 1131.595 | 1116.705 / 1131.095 |
| source / 1000 | 1412.847 / 1437.087 | 1397.056 / 1418.007 | 1395.456 / 1434.537 |

After/repeat facade overhead is 1.3–6.5%, inside the scoped 10% single-edit bar.
Native and facade medians remain within 3% of baseline. These measurements
qualify a single worksheet name edit, not mixed-kind rename dependency closure,
name-command batches or full-goal performance acceptance. Temporary build and
measurement artifacts were removed after retaining these small receipts.
