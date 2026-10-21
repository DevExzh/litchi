# 0359d — Graph package publication

- Date: 2026-10-21.
- Status: canonical Graph owner slice accepted; integrated gate and remaining semantic audit are required.
- Base: `1d7516988f` (0359c Graph typed budgets).
- Contract: 0351 §2.1 and §2.3; continuation of the semantic owner audit in 0359c.

The standalone Graph edit now starts on `Package` and commits
`litchi_core::edit::Commit<Package, Patch>`. The separate public `Snapshot`,
consuming `Package::snapshot`, and patch-side `apply`/`revert` are removed.
`Package::apply(&Patch)` returns the same commit triad, and undo uses
`Package::apply(&patch.inverse())`. Primary patch no-op inspection is
`is_noop()`. Every consumer migrates in this change.

Package bytes are shared without copying the input Vec. A private backing
retains the Vec directly until a borrow or clone needs sharing. OnceLock
serializes that promotion, and a private Mutex permits moving the Vec header
through `&self`; no byte content is copied or mutated. A package consumed
directly as an attachment needs no reference-counted allocation. After
promotion, a unique package still recovers the original allocation, while a
shared package copies when consumed as a Vec. Concurrent-borrow tests verify
one immutable allocation and recovery after other owners are dropped.
Edit creation only shares the source and initializes empty staging. The first
staged operation or commit performs the existing semantic chart admission, so creating and
dropping an edit is independent of package size. No-op commit still validates
chart semantics and preserves the exact source artifact.

Replay compares recorded base content, with a pointer fast path for the same
allocation. An independently opened equal package is accepted. The replacement
is validated under the recipient's retained limits, then its immutable bytes
are shared; replay does not copy a complete package. Publication releases the
writer buffer's spare capacity before retaining it in history. CFB reader
caches and the temporary Workbook copy are released before the writer's output
buffers overlap them. The existing changed-commit CFB writer, opaque-stream preservation, and
inverse artifact fidelity remain unchanged in this slice. Sharing untouched CFB streams during
changed publication still needs a storage-level review; this change does not
claim that replacing a return type alone proves that rule.

The root Graph family and immutable Package, PackageRef, Workbook, WorkbookRef
and Payload surfaces are registered with the glossary gate. A root Snapshot
ban and package/patch method bans prevent the former publication path from
returning. Nested raw chart models and shared OLE publication remain in the
semantic audit; they are not accepted exceptions.

## Validation

All 78 Graph tests, 19 PPT chart tests and 39 unified error tests pass (136
total, none failed or ignored). Graph/PPT all-target clippy, locked performance
harness compilation, Graph glossary checks (17 bans, one family, five
immutable owners), formatting, 36 tooling tests and the benchmark matrix
check pass. Commands and exact source/log hashes are retained in
[the validation receipt](results/0359d/graph-validation.json).

The paired probe in `results/0359d/graph-owner-probe.{py,rs}` builds the baseline
and candidate against identical dependencies outside the repository. It uses
an identically generated standalone Graph file, including opaque BIFF records,
CFB streams and either zero or 2 MiB of opaque CompObj data. It checks unchanged
output, cross-copy forward replay and exact inverse before timing ingress,
attachment output, edit creation, no-op/changed commit, replay, inverse and
128 retained patches. The adapter changes only the removed API calls. The
attachment route uses the existing ownership-transfer API on both revisions
and has independent setup, so the old Snapshot copy cannot change its
allocation warmup. Its preflight checks exact bytes, pointer and capacity.

Three process ABBA blocks per route ran on CPU 24 after all three workers
acknowledged holds and a process audit found no heavy competing work. Ratios
are candidate over baseline; ranges include every block:

| Route | Opaque bytes | Median | Range |
| --- | ---: | ---: | ---: |
| open_owner | 0 | 0.983 | 0.982–0.985 |
| open_payload | 0 | 1.004 | 0.982–1.012 |
| begin_edit | 0 | 0.018 | 0.017–0.018 |
| noop_commit | 0 | 0.989 | 0.986–1.029 |
| edit_commit | 0 | 0.991 | 0.987–1.006 |
| apply | 0 | 0.968 | 0.959–0.994 |
| inverse | 0 | 0.976 | 0.971–0.977 |
| open_owner | 2,097,152 | 0.077 | 0.077–0.080 |
| open_payload | 2,097,152 | 1.001 | 0.996–1.003 |
| begin_edit | 2,097,152 | 0.003 | 0.003–0.003 |
| noop_commit | 2,097,152 | 1.000 | 0.964–1.019 |
| edit_commit | 2,097,152 | 0.766 | 0.765–0.767 |
| apply | 2,097,152 | 0.081 | 0.079–0.084 |
| inverse | 2,097,152 | 0.081 | 0.081–0.083 |
| history128 | 0 | 0.948 | 0.944–0.954 |

The first shared-owner implementation increased peak history heap by 1.07 KB.
Releasing the CFB reader and temporary Workbook before output allocation
removed that overlap. Final whole-process peak heap is 437.88 KB versus
439.69 KB baseline, with 135,150 versus 135,154 allocation calls. This includes
fixture setup, preflight and five independent histories, each retaining all
128 patches and the current owner until observation; it is not a live-byte
measurement at a history marker. See [the heap receipt](results/0359d/graph-owner-heap.json).

Earlier attachment candidates were consistently 2–3.5% slower and were sent
back. Retaining the Vec until sharing is needed avoids a refcount allocation
for direct attachment transfer. Inlining the options boundary preserves
constant folding of the caller's default policy; structural validation stays
in the existing codec. The final attachment range overlaps baseline for both
sizes, while changed commit, replay and history improve. The final large
attachment instruction count is 111,728,371 against 111,704,452 baseline
(+0.021%, for the complete callgrind process including setup). Full samples,
source/binary hashes, rejected candidates, and a separately marked contaminated
run remain in `results/0359d/graph-*.json`; the contaminated run is not used for
acceptance. Binaries and raw profiling captures remain outside the checkout.

These native measurements cover the listed routes and fixture sizes. They do
not establish bounded-memory sharing of unchanged CFB streams in a changed
commit; that storage-level work remains required alongside the nested raw
chart and shared OLE audit.

The integrated workspace, corpus and performance gates remain required after
W4/W4b/W8 merge. Nothing is pushed by this record.
