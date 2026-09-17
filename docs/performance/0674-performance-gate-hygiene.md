# 0674: close the remaining harness gates and reproducibility hygiene

Status: retained, harness and gate work. **No file under `crates/` was
modified.** `performance_claim: none` — the selectors and checks below are
evidence infrastructure, and this record registers no performance result.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred
until that goal completes; iWork is excluded.

Change [0651](0651-queue-refresh-after-the-second-wave.md) left small selector
and gate follow-ups in rows 13, 17, 19 and 20. Change [0664](0664-perf-harness-marker-bearing-corpora-and-save-allocations.md)
had already landed marker-bearing corpora, ordinary-save allocation regions and
the DOCX text-sink selectors. This change finishes the remaining harness
selector, standing feature gate, allocator isolation and evidence hygiene from
that queue, on the 0664 head `5fa92d7ce`.

## What changed

### A selector for the opened PPTX transaction phases

`pptx_semantic_opened_transaction_phases` applies one text edit to the tiny
semantic PPTX corpus and records six public stages separately:

1. `Package::opened_presentation`;
2. `Snapshot::edit`;
3. `set_shape_text`;
4. `Transaction::commit`;
5. `apply_opened_presentation_commit`; and
6. `Package::to_bytes`.

`total_ns` encloses exactly those stages. Package construction, reopening,
semantic oracles, output digests and all correctness checks stay outside the
clocks. Each retained phase vector is reordered with the enclosing total's
sample order, every output digest must agree, and the reopened package must
pass the same semantic oracle as the existing one-edit selector. The selector
is attribution evidence; it makes no speedup, allocation, RSS, physical-I/O,
cold-cache or producer claim. It is opt-in, so `Case::DEFAULT` and its catalog
hash remain unchanged. The PPTX selector raised the count from 527 to 528;
the XLSB selector below raises it to 529.

The XLSB semantic matrix also has the opt-in
`xlsb_semantic_workbook_structure_edit_save` selector. It measures detached
`edit_workbook_structure` planning and commit, `apply_workbook_structure`, and
`Workbook::save` as separate phases. Each sample reopens the saved package and
checks the renamed worksheet and original typed numeric cell, with
deterministic output and phase-sum gates.

### Standing gates and hygiene

The non-iWork release gate now has a named `facade-polyglot-tests` mode that
compiles the facade with `--no-default-features --features docx,odt`, covering
the three DOCX/ODT polyglot detector tests. The harness gate keeps its fast
default suite and adds an explicit `allocator-metrics` run of
`docx_bounded_tail_append_compare` with `--test-threads=1`; those tests inspect
a process-global allocation counter and must not rely on Cargo's default test
parallelism.

`tools/native-resave/Cargo.lock` was regenerated with the current offline
index, and `cargo metadata --locked` plus `cargo check --locked --offline`
now agree with it. The repository's broad `*.log` rule remains in place, with
a final negation for `docs/performance/results/**/*.log`. This makes result
packet logs visible to Git while callers still choose which historical or new
logs to add explicitly; the parent rollup does not sweep unrelated untracked
logs. The harness README records that the root release profile's
`lto = true` means a rebuild need not have the same bytes as the executable
whose digest appears in a report.

The marker census's basis-point division now uses checked division after its
nonzero guard. This removes the existing `clippy -D warnings` diagnostic
without weakening an overflow check or suppressing a lint.

## Authority and scope

This is queue rows 13, 17, 19 and 20 of change 0651, read with the standing
correctness and safety decisions in [0652](0652-owner-decisions-for-the-third-wave.md).
No production crate, public API, malformed-input defense or output contract
changed. The harness remains under `tools/`, and its library continues to
forbid `unsafe_code`.

## Verification

The focused PPTX phase unit test and JSON smoke run pass. The allocator binary
passes all five tests serially. The DOCX/ODT facade gate passes all 95 facade
library tests and its integration targets. The native-resave manifest and
locked offline check pass, with four existing warnings in its stub binary.
The Python gate planner, source policy, corpus manifest, CRUD coverage and
corpus-binding test modules pass (59, 26, 12, 41 and 9 tests respectively),
and `cargo clippy --lib --tests -- -D warnings` is clean for the harness.

The full pre-fix harness sweep reached 526 passing tests, one failure in the
hard-coded selectable-case count, and one ignored test; every integration
target passed. The assertion was updated from 528 to 529, and the affected
count and XLSB lifecycle tests pass. The sweep failure was therefore a stale
harness expectation caused by this selector, not a semantic or integration
failure.

`check_perf_claims.py --mode structural` now returns status 0 and validates all
ten claims without opening retained evidence. The strict command with
`--evidence-root .` remains the named evidence gate and also returns status 0;
it independently verifies all ten claims and their retained evidence. Strict
mode still requires an evidence root when called directly, preserving the
required-evidence contract.

## Evidence

The packet is [results/change-0674](results/change-0674/). Its `gates.log`
retains the command outcomes, `decision.json` records the disposition,
`cleanup.json` names the isolated target and scratch paths, and
`log-sections.md` supplies the four paragraphs for the coordinator's rollup.
