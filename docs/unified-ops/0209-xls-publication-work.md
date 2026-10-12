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

## Paired measurements

The [nine-sample baseline](results/0209-xls-batch-before.json) and
[nine-sample updated run](results/0209-xls-batch-after.json) used the same
ascending numeric corpus and release build settings. They ran sequentially
after compilation and repository audits finished. All 20 cases pass exact
paired output bytes, complete coordinate/value checks, original-source
immutability and equal native/unified commit counts. Setup, opening, payload
construction, serialization and verification are excluded from these timings.
Medians in milliseconds:

| Owner | Operation | Calls | Native before | Native after | Batch before | Batch after | After batch/native |
|---|---|---:|---:|---:|---:|---:|---:|
| Eager | set | 1,000 | 4.974 | 4.065 | 5.803 | 4.848 | 1.1926× |
| Eager | set | 10,000 | 102.376 | 67.840 | 113.678 | 94.760 | 1.3968× |
| Eager | remove | 1,000 | 4.050 | 3.127 | 4.735 | 3.837 | 1.2270× |
| Eager | remove | 10,000 | 89.327 | 44.124 | 97.545 | 62.867 | 1.4248× |
| Filesystem | set | 1,000 | 4.693 | 3.812 | 8.164 | 7.192 | 1.8864× |
| Filesystem | set | 10,000 | 87.213 | 57.197 | 112.529 | 93.711 | 1.6384× |
| Filesystem | remove | 1,000 | 3.988 | 3.081 | 7.321 | 6.449 | 2.0929× |
| Filesystem | remove | 10,000 | 82.818 | 39.360 | 103.507 | 69.157 | 1.7570× |

At 10,000 operations, atomic sets improve about 17% and removals about 33–36%.
Native whole-edit times improve about 34% for sets and 51–52% for removals.
The combined patch contains both changes, so these paired measurements do not
isolate their individual contributions. At 1,000 filesystem removals, the
batch drops from 7.32 to 6.45 ms, but native improves more and the ratio rises
to 2.093×. This exceeds the 2× review threshold and remains a facade-overhead
follow-up; absolute speedup does not establish relative acceptance. Broader
text/resource, arbitrary-order and peak-memory qualification is not claimed.

A [reduced profile record](results/0209-xls-publication-profile.json) retains
both source commits, binary hashes, sampling commands, capture sizes and
SHA-256 values. The updated capture has 795 samples and no lost samples.
Neither the coordinate-dependency scan nor semantic-map cloning has observed
self samples in that capture. This supports the source-level removal of the
redundant work, but zero observed samples does not mean zero cost. Raw captures
were temporary and removed after reduction; the recorded commands regenerate
them. No executable is committed.

Baseline source: `90a797358b100a9dcee3182e7428ba8c1fda204f`; binary SHA-256:
`8c759476a3d9c2c6b45758755c05543228b6be8bb13c5e530675e4f6d696cb3b`.
Updated source: `2241f4a7f3d79660efd6f34b4091c9afefb95d7b`; binary SHA-256:
`13e78288b32a80793dd3821d3c69a26b2f9bd833802ba4309d7f40f96097b157`.
Build with `cargo build --release -p litchi-ops --features xls --example
xls_batch_bench`; run `target/release/examples/xls_batch_bench` with
`LITCHI_XLS_BENCH_SAMPLES=9`.

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
