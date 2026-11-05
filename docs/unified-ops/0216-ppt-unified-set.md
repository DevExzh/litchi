# 0216 — Unified legacy PowerPoint setters

The PPT read adapter now feeds native package transactions for unified `set`.
Selection remains positional or semantic, and native source-order targets are
resolved from the retained shape array rather than inferred from the physical
`/shape[N]` ordinal. The facade performs property admission and bounded target
expansion; the native transaction owns preservation checks and publication.
Supported properties are shape/placeholder `text`, `x`, `y`, `width`, `height`
and slide `hidden`. Geometry rounds to the nearest 1/576-inch native unit (ties
away from zero); unchanged native bounds preserve exact Small/Full encoding.
Zero extents remain legal. Text follows binary PPT rules, including soft line
breaks, without an XML character filter.

Visibility uses the selected slide's timing projection, including after source
release. Writing it requires one native-admitted timing atom; insertion when
that record is absent is a named capability gap. Nested shapes and rich text
remain explicit native/write coverage gaps.

`File::set`, explicit edits and batches retain immutable source handles. Unified
history keeps the native patch and exact before/after facade snapshots. Failure
at admission, staging or publication discards the candidate. Consecutive setters
currently publish separate native transactions; coalescing remains a required
performance follow-up. Text-only query fanout uses the native batch replacement
method in chunks bounded by its 256-replacement limit, within one native edit.
Empty property sets preserve exact bytes and report unchanged history after
target validation. Native operation counts are preflighted before edit capture, and the
before-image snapshot lower bound is checked before the native edit copies it;
final before/after retention is checked after publication.

## Verification

The all-feature facade suite passes 1,585 tests/doctests. Final PPT-only read,
write and capability tests pass 26 cases, including empty-property no-ops and 257-target text fanout.
Twelve unchanged native primary-edit contract tests pass. Regressions cover
both eager and positional owners, picture-before-text source-order mapping,
scoped fanout, text/geometry/visibility readback, exact no-ops, soft breaks,
Small/Full anchors, quantization, zero extents, native missing-record refusal,
retention/signature/staleness, publication/reopen, snapshot-bound patch replay,
inverse history, transaction rollback and atomic/best-effort batches.

All-feature/all-target Clippy and rustdoc pass with warnings denied. The
no-feature build, crate-boundary check (66 packages / 268 declarations / 11
existing debt items), OfficeCLI delta freshness and its 15 unit tests pass.

## Paired write probe

Implementation: `8f25264cbf`. [Samples and build identity](0216-ppt-write-results.json)
record seven alternating native/unified pairs of ten immutable text writes,
with setup and exact output-byte comparisons outside timing. Release LTO is
disabled. This local probe is not frozen-harness qualification.

| Shapes on slide | Memory median paired ratio | Positional median paired ratio |
| --- | ---: | ---: |
| 1 | 1.064 | 1.106 |
| 100 | 1.325 | 1.340 |
| 1,000 | 1.392 | 1.388 |

The single-write 10% budget is missed except for the one-shape memory case.
Follow-up: profile target projection, native-owner adoption and source fences;
remove unnecessary work while retaining admission and snapshot checks. Add
frozen harness routes and qualify coalesced multi-command batches separately.

The base remains `32ad64e3207655b4ba558fc2793d9a7396655cdf`, already an ancestor.
Repository lockfiles remain unchanged; Cargo uses a scratch offline lock with
cached cc 1.6.0 instead of unavailable cc 1.7.0.

## Remaining scope

Complete PPT structural CRUD, nested shape writes, rich text, coalesced batch
performance, the frozen harness gates and the all-format/Python/CLI goal remain
open. This record does not claim the full objective is complete.
