# 0209 — XLS publication dependency scans and patch ownership

Profiling the indexed XLS batch benchmark found an avoidable scalar-publication
scan. `patch_selection_records` walked the complete raw-record list for every
structural cell operation to find selection records, then immediately returned
from selection patching because the operation was not an axis shift. This
multiplied scalar operation count by worksheet record count, even when no
selection needed changing.

Selection rewriting now checks for a row or column shift before traversing
the records. It still parses raw record framing first, and the surrounding
coordinate-dependency path still performs its existing framing and dependency
validation. Scalar operations preserve selection bytes; axis shifts retain the
same selection, formula, merge, hyperlink and drawing dependency handling.
No validation or native staging limit is relaxed.

Canonical `Package::Edit::commit` also consumes the native commit's snapshot
and patch after reading its diagnostics. Previously it cloned the entire
reversible patch, including both operation directions and their nested
precondition/value maps, before wrapping it. Moving the patch retains the same
no-op handling, diagnostics, exact inverse, policies and freshness fences.

## Evidence and verification

The pre-change profile used the existing complete batch harness, with three
samples, warmups and correctness verification. It collected 1,101 user-cycle
samples without loss. `patch_coordinate_dependencies` accounted for 14.17% of
self samples and semantic-map cloning for 4.35%. Those are whole-harness
sample shares, including setup and verification, not isolated commit timings;
source inspection established the redundant scan and deep clone. Paired
unprofiled timing and final verification are recorded with this change.

All 1,794 native XLS and 1,548 all-feature facade tests/doctests pass. New
regressions check exact selection bytes for scalar operations, malformed BIFF
frame refusal and mixed scalar/row-shift handling. Existing owner, inverse,
replay, source-freshness, policy, batch barrier and bounded-rollover tests pass.

Warning-denied all-target Clippy and rustdoc pass for native XLS and the facade.
Minimal and XLS-only feature checks and the XLS glossary audit pass (24 banned
names and one transaction family). The crate-boundary audit passes for 66
packages, 268 internal dependencies and 11 existing debt items. The OfficeCLI
generated table remains current and all 15 generator tests pass. Repository
lockfiles are unchanged; checks use an external offline lock selecting cached
cc 1.6.0 and an external build target.

Reproduce using `cargo test -p litchi-xls --all-features`,
`cargo test -p litchi-ops --all-features`, and
`cargo clippy -p litchi-xls -p litchi-ops --all-features --all-targets -- -D warnings`.
No API/capability claims or Python/CLI acceptance status changed.

## Remaining scope

The 10,000-operation workload still requires three native publications because
of the 4,096-operation bound. Each validates its complete candidate; the 1,000
case uses one publication. Their growth cannot by itself distinguish per-edit
scaling from repeated whole-file work. Native and facade general linear
scaling, source-freshness overhead, arbitrary insertion order, text/resource
workloads and peak-memory acceptance remain open.

Worksheet/axis CRUD, formula/style editing, source query streaming, the other
format tiers, Python/CLI and complete goal acceptance remain required. The
base remains an ancestor at `32ad64e3207655b4ba558fc2793d9a7396655cdf`; no new
rebase was needed.
