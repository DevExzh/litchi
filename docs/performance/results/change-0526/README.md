# 0526: XLSX scanner layout investigation

This batch advances the next owner selected by accepted change 0525. It audits
retained commit profiles and the current scanner/writer source, and prepares a
reviewable candidate for fresh measurement. It makes no new production speedup
claim. OLE2/OOXML remains the active priority; ODF is deferred and iWork is
excluded. The complete performance goal remains open.

The previous goal turn made progress: commit `67028ab60` retained source-bound
unchanged-cell readback reuse after native, allocation, profile, correctness,
and eager-guard checks. The current source and all 30 accepted ADR/index hashes
are revalidated in `source-binding.json`.

The investigation keeps complete XML validation, independent semantic readback,
source/version and execution fences, byte preservation, resource limits, and
formula/dependency behavior. It excludes the rejected 0514/0516 fusion and
0522 combined cell-reference/tag scanner mechanisms.

## Candidate boundary

The current scanner stores a separate `Vec<Span>` in each pending cell, then
converts it to `Box<[Span]>` at cell close. The writer consumes the list only
when replacing cell payload. A row-owned span array can retain the same spans
in the same order while cell slots retain ranges into that array. The two
writer paths must resolve ranges through their owning row. This removes a
per-cell storage representation, without omitting scan events, payload spans,
or validation.

The draft must preserve empty cells, empty rows, multiple and repeated primary
payloads, unknown/interleaved markup, namespace aliases, formula metadata,
implicit coordinates, and exact output. A single inline span is insufficient
unless arbitrary multi-span cells remain representable. A whole-sheet arena
would increase the growth window; row ownership keeps that window bounded by
one row. Neither representation establishes a measured memory improvement.

## Required next experiment

Before production retention, capture a fresh baseline from the unchanged
accepted source; apply the reviewed candidate and focused tests; run the
applicable quality checks; then compare normal native and separate allocator
runs using the same harness, corpus, machine, and binary custody protocol.
Use both medium and dense-sparse source-backed edit/save shapes and ordinary,
managed, tiny, sparse, formula, and noncompact guards. Include multiple primary
payloads and style-only edits in correctness tests. Measure commit and total
edit/save independently, with reopened semantics outside the measured total
as defined by the existing harness.

Freeze practical admission thresholds before capturing the first baseline.
Use matched independent repeats, retain uncertainty and every individual >5%
adverse timing/RSS result, and reject the production patch if native results
do not justify its complexity. Lower allocation-call counts alone do not
prove useful end-to-end improvement; rejected change 0522 demonstrates why.
Do not multiply Callgrind shares by elapsed shares and present the product as
a wall-time bound. Retained Callgrind call metadata also includes
collection-off work and does not establish operation-local allocation counts.

## Evidence and reproduction

- [Source binding](source-binding.json) and [prior seal check](prior-seal-verification.json)
- [Raw profile decomposition](profile-analysis.json) and [profile review](profile-review.md)
- [Scanner design review](scan-design-review.md) and [root review](root-review.md)
- [Draft patch](row-primary-arena.patch), [candidate hashes](candidate-source-binding.json), and [candidate review](candidate-review.md)
- [Focused test plan](test-plan.md)
- [Verification](verification.json) and [negative analyzer controls](analyzer-tests.json)

```sh
python3 -B docs/performance/results/change-0526/analyze.py
python3 -B docs/performance/results/change-0526/analyze_test.py
python3 -B docs/performance/results/change-0526/verify.py --sealed
```

The decomposition checks raw edges and self/direct equations for five nested
owners across all four retained commit profiles. The negative controls reject
altered edge and self costs. Verification requires exact decomposition replay,
reconstructs the draft in a private Git index, checks its hashes against the
reviewed candidate, and proves source/ADR bindings and scratch cleanup. The
root working `Cargo.lock` is ignored by Git; its literal content is retained as
a working-only observation rather than falsely bound to a Git blob. The
standalone harness lock remains separately Git-bound.

The three-file candidate passed patch applicability and formatter checks only.
It has not been built or run. The owned candidate directory was removed, and
all verifier scratch directories are automatically removed. No prior evidence
or repository Rust source was modified by this batch.
