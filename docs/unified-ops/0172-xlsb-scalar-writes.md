# XLSB scalar cell writes

The unified XLSB adapter accepts cell `value` assignments (normalized numbers or positive zero,
exactly representable integers, Booleans and text), null-to-blank assignments,
and `clear: true`. Negative zero, subnormal and non-finite numeric payloads
are refused before staging. Direct cells, ranges and cell queries share the native
package-owned cell editor. Empty properties are an exact no-op. Formula
expression authoring remains a visible property gap.

Native `cell_values::Scalar` and `Edit::set_scalar` provide the missing semantic
assignment operation. Unlike family-specific setters, this operation can change
storage families while preserving style and phonetic flags. Missing nonblank
cells are inserted with the default style; clearing an absent cell does not
insert a record. A same-value direct assignment preserves the original bytes.
Replacing an ordinary formula materializes a direct value even when that value
matches its cache. Shared, array and data-table formula ranges, including
unstored followers, are guarded. Formula replacement with an existing
calculation-chain relationship remains a typed refusal.

Consecutive direct writes to distinct cells in one worksheet share one native
edit. Repeated coordinates, worksheet changes, reads, queries and other write
families introduce commit barriers. Multi-worksheet selectors retain a sequential
chain of native package patches. The original immutable handle survives failures;
forward/inverse replay uses exact package images. Source freshness, explicit
signature consumption and atomic publication retain the existing XLSB policy.

Retained-byte accounting includes package images, worksheet images and owned
semantic cell payloads. Multi-worksheet writes decrement the remaining quota
before starting the next native edit. Coalesced writes enforce per-item and
aggregate staged-byte limits and report exact no-ops per batch item.

The scalar editor builds coordinate and formula-group metadata only for an edit;
ordinary cell-value reads do not acquire the new indexes. The facade performs
coercion and limit admission, while binary editing and shared-string consistency
remain native responsibilities.

## Shared strings

Changing shared-string cells updates the workbook-wide reference count through
the actual shared-string relationship target. Unused entries are preserved when
valid; necessary compaction remaps remaining worksheet references while retaining
raw record bytes. The count constraints come from
[MS-XLSB BrtBeginSst](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xlsb/5301dcf1-f4cc-4707-b97a-105e788cd880).
Opaque cell families that prevent a complete reference proof are refused.

## Verification and measurement

The final native `sign,vba-inspection` suite passed 1,128 tests/doctests;
the final XLSB-only facade suite passed 521. The combined
DOCX/XLSX/PPTX/XLSB/encryption facade suite passed 1,190 before the final
shared-string source-count and strict calculation-chain guards; final native
and XLSB-only reruns cover those guards. The twenty new facade tests cover
scalar coercion, blank/no-op semantics, ranges and multi-sheet queries,
coalescing and read/name/repeated-cell barriers, per-item changed flags,
atomic/best-effort batches, quotas, replay/inverse, capability metadata and
save/reopen. Native tests cover storage-family preservation, ordinary/grouped
formulas, relationship-targeted shared strings, required index compaction,
opaque FRT refusal, inconsistent source counts and retained patch quotas.

Native and combined-facade all-target Clippy passed with warnings denied.
Warning-denied native rustdoc JSON and facade documentation passed, as did the
full XLSB glossary profile (26 banned names), minimal facade compilation,
capability golden regeneration and deterministic OfficeCLI delta regeneration.
All 14 parity-generator tests passed. Crate boundaries passed for 66 packages
and 268 internal dependency declarations. Scoped formatting and diff checks
passed. Independent reviews checked package endpoints, dependency preservation,
coalesced attribution and retained-byte accounting; reported correctness issues
were fixed before acceptance.

Reproduce the suites with `cargo test -p litchi-xlsb --features
sign,vba-inspection`, `cargo test -p litchi-ops --no-default-features --features
xlsb`, and the latter command with `docx,xlsx,pptx,xlsb,encryption`. Builds used
external targets, offline resolution, zero dev/test debug information and no
incremental compilation. Root builds used an external lockfile copy selecting
cached `cc` 1.6.0; both inherited repository lockfiles remain unchanged.

### Paired measurements

The standalone harness is `tools/perf-baseline/unified-xlsb-cell-value`.
[Initial results](results/0172-xlsb-cell-values-initial.json) and
[final-build results](results/0172-xlsb-cell-values.json) retain all samples,
corpus/output hashes, source-file manifests and binary/build provenance. Each
run uses three ABBA rounds, six samples per leg, CPU 0, and no concurrent
validation job from this turn. Owners match eager with eager and source with
source. Source materialization occurs inside both clocks. Opening, target and
property preparation, output serialization and equivalence checks are outside
the clocks. Both legs use the same native scalar assignment owner.

All twelve cases in each run have exact warm-up ZIP-byte equality and empty
failure matrices. Every coalesced case reports one native commit. Final ratios:

| Operation | Cells | Eager unified/native | Source unified/native |
|---|---:|---:|---:|
| One-shot set | 1 | 1.186 | 1.046 |
| Transaction | 1 | 0.989 | 1.074 |
| Transaction | 10 | 1.194 | 1.080 |
| Transaction | 100 | 1.403 | 1.239 |
| Transaction | 1,000 | 1.634 | 1.577 |
| Transaction | 10,000 | 1.689 | 1.743 |

The 1k and 10k batches meet the 2x review threshold. Unified 10k/1k growth is
10.115x eager and 9.359x source in the final run, consistent with approximately
linear scaling in this generated numeric workload. Formula-group-heavy and
shared-string-heavy batches are not measured by this harness.

**The eager one-shot case does not close the 10% performance requirement.**
Its final native/unified medians are 109.961/130.381 microseconds, versus
112.640/123.000 microseconds in the initial run (ratios 1.186 and 1.092).
Raw small-case samples vary materially; this is evidence of a review-trigger
miss, not proof that the difference is only noise. Follow-up: profile and
reduce direct-cell target resolution/admission overhead, then repeat paired
small-edit measurements. The one-operation transaction is also close to that
threshold in the initial run (1.102). No broader performance acceptance or
allocation/RSS claim is made.

## Remaining scope

XLSB formula authoring, formatting properties, structural facade CRUD, creation,
re-encryption and broader corpus acceptance remain open. Package cell-value
patches still use process-local exact images; durable encoding, history/merge
and scoped error vocabulary remain follow-ups. Source edits still materialize
an eager workbook. The all-format Rust/Python goal remains active.


Temporary implementation worktrees and branches were removed after integration.
Scratch builds, copied locks and raw logs were removed after recording the
validation and reduced measurement evidence. No compiled binary or large
evidence artifact is committed. The rewritten base was already contained in
this branch at the start of the turn.
