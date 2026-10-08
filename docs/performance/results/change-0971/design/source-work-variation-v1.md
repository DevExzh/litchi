# 0971 — source work variation review

**Status: read-only source investigation.** This note does not identify a
root cause, change runtime code, build a binary, launch a workload, or make a
performance or adoption claim. It investigates why the 0970 historical
baseline could execute different amounts of work in fresh processes while
publishing the same bytes.

## Revalidated inputs and limits

The checkout was at `29f8c61d51b1c81f4f2afe6b9f61984e6347a6a6`
(`perf(doc): record baseline affinity diagnostic`). The pre-existing working
files were preserved. `docs/GOAL.md` was read in full (35,839 bytes,
SHA-256 `bed4058bb76330daab8ce9d4bceff639ab3fbd7ea06634158bef41b133c4d1f1`)
and `docs/adr/README.md` was read in full (5,766 bytes, SHA-256
`2af0e5814378c2765c12213e495d02c2895a27f142bc618ca1f3037824bbe028`). All 32
accepted ADR files were read and rehashed: 3,147,627 bytes in total. The
stable manifest of `docs/adr/0001` through `0032` (relative path plus SHA-256)
is SHA-256 `378670abf4699385a3501a661becf2feb89e0204acf4ef96f9ab255cab0c766`.
The ADR README has no outstanding proposed record.

The governing constraints matter here: correctness and lossless publication
precede speed; resource limits, typed errors, validation, source provenance,
and deterministic output must remain intact; optimization requires a mapped
path and a measured hypothesis. The observations below are therefore
mechanisms to discriminate, not a proposed production change.

## What 0970 actually established

The sealed 0970 packet used the same retained 0967 historical ordinary
baseline executable in every arm (SHA-256
`0db34357a1e725dc397a4fd895532433db1214c5862b7bc96bf23f21c0396ecb`). It had
four cells, 12 pinned/unpinned pairs per cell, 24 serial children per cell,
five warmups, and 80 samples per child. The independent scalar audit is
`docs/performance/results/change-0970/review/raw-independent-scalars-0970-v1.json`
(SHA-256
`c9bca1f01c575678f20632e7ecdd7ded0075e7baf7094523b22914f11abfcd70`).

For `docfloat/e1/default` (the `FloatingPictures` fixture), whole-process
`perf` instructions and the operation-local child p50 have these ranges:

| arm | instruction count, min–max | instruction CV | child p50, min–max | child p50 CV | Pearson(p50, instructions) |
| --- | ---: | ---: | ---: | ---: | ---: |
| pinned | 6,987,088,875–7,494,203,611 | 2.73687% | 978,289–1,270,130.5 ns | 10.3561% | 0.9552515504 |
| unpinned | 6,988,814,257–7,498,294,364 | 2.30574% | 979,605–1,267,456 ns | 9.52625% | 0.8918102135 |

These are associations over 12 children in each arm. `perf` covers the
launched GNU `time` command and inherited launcher/baseline process tree,
including wrapper startup; the harness `Instant` p50 covers only the native
operation region. The audit explicitly says the scopes are separate and the
association is descriptive. The 0970 report also records zero migrations in
all 48 pinned packets, but that does not identify the source of the remaining
variation.

A read-only census of the 96 retained packet reports found one output hash,
output size, output capacity, retention-route shape, and all-true correctness
gates for every child within each cell. For `docfloat/e1/default` the output
was 342,528 bytes with SHA-256
`6ba1bc972d66715e76d882f47c587b22000100f47abaa88428990619fc59fa20`; the
three retained-render route records were identical in all 24 children. Stable
bytes and route observations rule out an observed semantic or fallback-route
split, but do not imply identical internal work.

## Actual measured path

The current source files below match the retained historical baseline source
for these paths. Their current SHA-256 values are recorded so the anchors can
be checked against the revalidated HEAD:

| source | SHA-256 |
| --- | --- |
| `tools/perf-cfb-profile/src/lib.rs` | `b4893321ac9a9321a9055823a7f05ecb35e59319531913affc49342b8ccd73dc` |
| `crates/litchi-doc/src/body_text.rs` | `0e8cb6351be3327322882ae4035230cd241aed0b6208cb8589e7d76692d3d786` |
| `crates/litchi-doc/src/tracked_revision/package.rs` | `a04da643b093ec9724c9ac57ed8fbfdf2c2fedfc16eade105c91c32df7c9a446` |
| `crates/litchi-ole-common/src/object/editor.rs` | `ee6079567bf85a0ff68a335fb43999d780fab335a09a9cec8b845c16b7b5b319` |
| `crates/litchi-ole-common/src/object/codec.rs` | `99dc73a358cbe5a76c6fc242e42af6524ba6d0954c2ae508cf55f84cb6a5cafd` |
| `crates/litchi-cfb/src/writer/core.rs` | `a5ead3dded13baa8bdaa350742bfe79185e39cc2f16f86e433701e939fb2b19c` |
| `crates/litchi-cfb/src/writer/layout.rs` | `480b30fb1c6c018c3182fc5e0ab1f53bb55f61169396e2f2d8bce257c7d1460c` |
| `crates/litchi-cfb/src/writer/directory.rs` | `c810c4ae2e0fdeb89a2c3fe0c207422b65849ad516cd5343dcf211d424ca0b58` |

The ordinary probe's lifecycle is explicit in
`tools/perf-cfb-profile/src/lib.rs:290-329`: each `run_case` copies the
fixture into `Snapshot::open_bounded`, creates an edit, repeats paragraph-zero
replacements, commits, and copies the output `Vec`. `:353-358` starts and
stops the operation-local `Instant`; `:686-790` prepares cross-route and
semantic gates before warmups/samples; `:937-1018` performs warmups, samples,
and post-sample reporting. Thus the child performs more than one CFB
publication: preparation, five warmups, and each sampled edit lifecycle are
all in the process, while only the `measure_one` lifecycle is timed.
The fixed retention policy does change the deterministic route:
`crates/litchi-doc/src/tracked_revision/package.rs:2654-2659` takes a retained
render when one is held and otherwise renders again. That is route-dependent
work between the configured zero/default/release controls; the retention policy
is fixed within each 0970 child and does not explain a cross-child route flip.

`crates/litchi-doc/src/body_text.rs:652-697` performs source-length
preflight, strict `RevisionEditor::open_with_ole_file_with_policy`, complete
public-reader validation, and exact source retention. A changed commit
reopens the rendered candidate through the same owner at `:2086-2107`.
`crates/litchi-doc/src/tracked_revision/package.rs:968-1038` opens the
object owner and parses the Word streams and revision tables.

The hot CFB candidate publication is
`crates/litchi-ole-common/src/object/editor.rs:977-1000`:

1. check the candidate package;
2. try `render_copy_through` under the default reuse policy;
3. otherwise call `Package::render_with_layout`;
4. reopen the rendered bytes with `OleFile::open` and `codec::open`;
5. recapture the complete package with `Package::capture`; and
6. run `reuse_stream_allocations` and target discovery before publishing the
   candidate.

`crates/litchi-ole-common/src/object/codec.rs:639-722` constructs an
`OleWriter`, adopts the source layout, sorts the package's storage `Vec`,
registers storages/CLSID/metadata and streams, then calls `write_to`.
`render_copy_through` has deterministic topology, metadata, and stream-length
preconditions at `:739-883`; a process hash seed cannot make those equality
checks choose a different logical result. The reopened capture is sequential:
`codec.rs:1157-1221` walks CFB directory entries and charges the same bounded
storage/stream/byte budgets. `reuse_stream_allocations` at `:214-232`
builds a temporary path `HashMap`, then performs content equality checks and
shared-allocation replacement.

## Plausible process-specific work variation

The strongest source-level candidate is the CFB writer's use of seeded hash
tables as intermediate work structures. This is a cost hypothesis; no causal
claim is made.

* `crates/litchi-cfb/src/writer/core.rs:312-350` stores explicit storages in
  `HashSet<Vec<String>>` and storage CLSIDs and metadata in `HashMap`s.
  `:527-540` iterates the storage `HashSet` and inserts each path and ancestor
  into a `BTreeSet`. The final set is content-determined, but the insertion
  order is the process-specific hash iteration order. B-tree node comparisons,
  splits, branch decisions, and allocation layout can consequently differ
  while the in-order result remains the same.
* `:542-580` copies `storage_clsids` from a `HashMap` into a `BTreeMap` before
  planning. The final key order is canonical, but construction order can
  change tree insertion work and allocation layout. Metadata tables are then
  used through content lookups by the reuse planner (`writer/layout.rs:
  1444-1525`), so a seed can change probe cost without changing a gate.
* `:1631-1659` tries the source-layout plan first. If it declines, the
  from-scratch writer at `:1743-1777` explicitly converts the seeded
  `storages` and `storage_clsids` tables into temporary vectors and sorts them.
  The comment there documents that Rust seeds those table iterations per
  process and that sorting is required for deterministic SIDs and bytes. The
  sorted result is stable, but `sort_unstable` comparison/branch work depends
  on the input permutation. `:1779-1785` also iterates storage metadata
  directly when that table is populated.
* `crates/litchi-cfb/src/writer/directory.rs:301-312` holds the directory
  path and child indexes in `HashMap`s. `:403-462`, `:467-570`, and
  `:624-671` perform repeated path/child lookups while building and linking
  the directory. Seed-dependent bucket placement changes probe and branch
  work even though directory entries are emitted in SID order.
* On the reopen side, `object/codec.rs:214-232` uses a temporary path
  `HashMap` for allocation reuse. The capture order and equal-byte decision
  are vector/content driven, but seeded bucket placement can change lookup
  probes. `object/directory/validation.rs:127-160` similarly uses a bounded
  `HashSet` for SID membership; it is inserted and queried sequentially, so
  it can vary probe work but cannot reorder the validated directory.

This mechanism is more plausible for `FloatingPictures` than
`NoHeadFoot`: the former has nested `ObjectPool` directory structure and more
storage/path metadata work, while the latter has a small six-entry directory.
The 0970 comparison is consistent with that prioritization (`docfloat/e1`
had roughly 10% child-p50 CV; `docnohf/e1` had roughly 0.8%), but the
comparison is not proof. The 0970 packets do not expose hash order, B-tree
comparisons, hash probes, allocator addresses, or writer phase counts.

The candidate does **not** currently explain a route change. CFB path choices
are content/limit based: `render_copy_through` checks exact stream lengths and
identity; `plan_reuse` receives sorted/B-tree projections and validates source
paths, sizes, metadata, and budgets; capture traverses the parsed directory
rather than an unordered map. The retained packets show one **retention-route report shape** and one output
identity per cell; they do not expose the internal CFB layout-plan outcome. A
hash seed can plausibly change the cost of these paths, including temporary
allocation behavior, while leaving route, errors, semantic records, and bytes
unchanged. If a future probe observes a
route flip, that is an invariant/diagnostic issue to investigate before any
optimization.

The perf association also has an important alternative explanation: its event
counts include the external GNU `time`/Python launcher and the baseline
process tree, while the p50 is only the operation-local `Instant`. Startup,
allocator/cache state, scheduler effects, or wrapper work can contribute to
whole-process instruction spread. The source candidate remains unproven until
its internal counters are observed.

## Small discriminating probe for a future authorized run

No probe is implemented or run in this task. The existing feature-gated
content-free observer seams (`litchi-ole-common/src/object/editor.rs:38-77,
154-183`; `litchi-doc/src/body_text.rs:700-798,2111-2206`) provide a narrow
place for a diagnostic-only extension. A future diagnostic binary should
retain the ordinary uninstrumented binary as the performance reference and
use a separate instrumented binary only to classify work; instrumented timing
must not be presented as the baseline result.

The smallest useful per-publication record is:

* whether `render_copy_through` was accepted, declined, or never eligible;
* whether `plan_reuse` was accepted, declined, or fell through to the
  from-scratch writer;
* storage and CLSID entry counts plus a content-free digest of the observed
  hash-table iteration order before B-tree/vector construction;
* the `effective_storages` and metadata table counts; and
* capture/reuse counts, including the number of equal-byte stream allocations
  reused.

For the fallback writer only, a local comparison-counter around each existing
`sort_unstable` closure would distinguish permutation-dependent sort work.
The digest and counters must be bounded, content-free (for example, path
identity/length classes rather than names), and emitted only by the diagnostic
observer. They must not influence ordering, limits, error handling, output, or
source ownership. A diagnostic run should preserve the 0970 fixture hashes,
all existing correctness/oracle gates, and the exact no-retry packet policy.

A bounded follow-up would repeat the 0970 `FloatingPictures/e1/default`,
`FloatingPictures/e32/default`, and `NoHeadFoot/e1/default` cells with fresh
children under the same source, fixture, affinity arms, 80-sample lifecycle,
and wrapper commands. It would then compare, within the instrumented binary,
iteration/order fingerprints and phase counts against whole-process
instructions and the operation-local p50. This is a mechanism test, not a
qualification run.

Interpretation is deliberately narrow:

* varying order fingerprints with matching sort/B-tree phase variation and
  stable route/output would support randomized container work as a contributor;
* stable order fingerprints with instruction variation would leave hash probe,
  allocator/cache, and wrapper/host effects unresolved;
* a changed route, error, or semantic/output gate would invalidate the
  performance experiment and require correctness review; and
* stable source counters with differing whole-process counts would prioritize
  the wrapper/process-tree and host alternatives.

No current source change, build, workload, or formal performance decision
follows from this review.
