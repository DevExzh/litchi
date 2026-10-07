# 0943 — DOC Data snapshot sharing rejected

**Decision: rejected.** The private Data-sharing candidate saves 4.974452% of
gross requested bytes for FloatingPictures at 32 edits, below the frozen 5%
improvement gate. FloatingPictures at two edits also fails the paired latency
confidence gate. All 288 formal children pass correctness, paired capacities
match, and allocator safety passes. Neither failed gate is waived. Production
code and its tests remain unchanged; the candidate and its tests are retained
as a reproducible patch under `results/change-0943/`.

## Candidate and scope

The DOC revision editor currently stores the admitted `Data` stream as a
`Vec<u8>`. Clone-first revisions therefore copy that stream even when the
operation changes only text. The opportunity is material for the selected
fixtures: `NoHeadFoot.doc` has a 4,096-byte `Data` stream and
`FloatingPictures.doc` has a 255,308-byte stream. These sizes motivate the
candidate; they are not performance results.

The candidate changes the private editor representation to `Arc<Vec<u8>>`.
Admission wraps the existing owned stream. Text edits continue to share the
unchanged stream. The two Data mutation points — picture-block append and
picture-tail removal — call `Arc::make_mut`, so a clone-first candidate
detaches before mutation. Commit keeps the existing publication shape by
cloning the Data contents into the replacement `Arc<[u8]>` or added stream.
The source editor, limits, typed errors, package checks, and publication
boundaries are unchanged.

The six production hunks in
`crates/litchi-doc/src/tracked_revision/package.rs` are:

1. Change the private `RevisionEditor.data` field from `Vec<u8>` to
   `Arc<Vec<u8>>`.
2. Wrap the admitted `Data` stream in `Arc`.
3. Use `Arc::make_mut` before picture-data append.
4. Use `Arc::make_mut` before picture-data truncation.
5. Clone `self.data.as_ref()` for an existing-stream replacement.
6. Clone `self.data.as_ref()` when adding a missing stream.

Focused tests in that same file cover text-edit sharing, picture install and
restore detachment, invalid-graph refusal, a real 4,096-byte Data binding, and
late typed stream-limit refusal after a valid Data append. The tests assert
source bytes, source text, Word/Table buffers, and held shared Data remain
unchanged after candidate mutation or refusal.

The measured scenario maps to the [CRUD taxonomy](../CRUD_Scenario_Checklist.md)
opened-document field updates and measured-performance evidence: bounded owned
open, an edit transaction, 1/2/8/32 replacements of paragraph zero, commit, and
owned output copying. Preparation validates all three retention routes, semantic
readback, public-reader reopen, exact no-op sharing, and patch behavior outside
the operation timer. This does not qualify unrelated create/delete/concurrency
scenarios or all Office formats.

## Source and build provenance

The baseline is repository revision
`13562481317d3d219c6fcd93babd9377e84ea834`, exported at
`/home/zhuhe/litchi-performance-0943/baseline-source-v1`. Its package source
file is 146,828 bytes with SHA-256
`a04da643b093ec9724c9ac57ed8fbfdf2c2fedfc16eade105c91c32df7c9a446`.

The authoritative candidate is the v4 export at
`/home/zhuhe/litchi-performance-0943/candidate-source-v4`. Its source receipt
is `candidate-source-v4.json` (981 bytes,
`0b190a2e398c1c1f0d1084ba00e520640334c8b88bd2ca835b765b19cfe96381`). Only
the package source file differs from the baseline tree. That file is 154,854
bytes with SHA-256
`ce3db3822fcfb7ed88120ee6b761445f1b06a1f3ce25950c055295e1ee38b863`.
The committed patch omits whitespace-only context markers; applying it
reconstructs the exact candidate file hash above. The archived exported patch
retains its original formatting and descriptor below.
The v4 source listing is 1,722,211 bytes with SHA-256
`2c86ab8e0b16f659f2ecf84a2f7495ea367cd32ee298720b4496abafdcbb19a7`.
The baseline build source receipt is `baseline-source-v2.json` (682 bytes,
`3d4cc656f40da02f7736f1ed7744149dd98bd5ad46a30c16ee1bc84e8e84e804`), and
its bound source listing is 1,722,211 bytes with SHA-256
`bd485793d1495013faff2ee8ab3053796c7df867a0d08c2eb32d187f13d0e693`.
The exported v4 patch is 11,574 bytes with SHA-256
`a3bc5a4002c11f67314c97c43dae126904884c8034c004ea879dd00c9b1e15c3`.

The normal-release probe builds used Rust/Cargo 1.95.0,
`--offline --locked -j2 --release`, `CARGO_INCREMENTAL=0`, and
`CARGO_PROFILE_RELEASE_DEBUG=1`, with the unchanged probe lock
`abd9f47a5aaa973a760cd61392de79c2b6143026c300dc1329831652aa54c06c`.
The probe manifest is 541 bytes with SHA-256
`0e6f34cc57d56b2725034d4ed067e06a9a28d571f59348d0412c3c6966ea157d`.
Native and allocator builds used the same locked release command and changed
only the binary name (`doc_retention_probe` versus
`doc_retention_probe_alloc`); their start, terminal, and binary receipts are
under `builds/{baseline,candidate}-{native,allocator}-*/`.
The audited binaries are:

| lane | baseline | candidate |
| --- | --- | --- |
| native `doc_retention_probe` | 50,601,376 bytes, `7763ab3f323e3b7de8b4fe8f45dcb363dfcf483aa60ddb8428c9e31f5fef06dc` | 50,633,720 bytes, `db6a09d6128a43cdf0d1e01dba2b091cb8a4dc1cbccd34e60bce7a7d030fc154` |
| allocator `doc_retention_probe_alloc` | 50,605,064 bytes, `ce19bd6a2f923b57271ded2ad02b0aa635295a9432d1e4a324d133d92a148f91` | 50,637,400 bytes, `7b8994bdb1d6964d8f966a8df42cdaf3dc99d80bfae85c0cf4b0df214d2a7113` |

The build audit at
`/home/zhuhe/litchi-performance-0943/receipts/build-audit-v1.json` passed
source, lock, manifest, command, environment, and binary custody checks. It
makes no timing claim.

## Preparation history

The discarded preparations are retained in the external workspace and are
reported here so their failures are not confused with comparative results:

- v1 stopped before tests with an `unused_mut` compile error.
- v2 ran 1,123 tests successfully but failed two picture tests because the
  `NoHeadFoot` receiver used by the helper had a non-empty picture graph.
- v3 ran 1,122 tests successfully but failed three picture tests because the
  generated Writer receiver already contained a `Data` stream and the helper
  attempted `add_stream`.
- v4 corrected the helper to replace that stream with `put_stream` and is the
  authoritative source used for the final quality checks.

These were test-fixture preparation failures. They do not constitute workload
measurements or candidate rejection evidence.

## Scoped quality

The authoritative quality summary is
`/home/zhuhe/litchi-performance-0943/checks/quality-summary-v1.json`.
Candidate v4 passed:

- default `litchi-doc` tests: 1,365 passed, 13 ignored, 0 failed across 48
  suites (1,125 unit tests plus 240 additional integration/doc tests);
- `--all-features`: 1,383 passed, 14 ignored, 0 failed across 48 suites;
- warning-denied all-targets Clippy: exit 0;
- changed-file Rustfmt check: exit 0.

Both public and private rustdoc checks exit 101 with exactly their respective
baseline errors and locations. The public comparison records the four
unchanged files `parts/bookmarks.rs`, `parts/comments.rs`, `parts/rsids.rs`,
and `parts/sections/codec.rs`; the private comparison likewise matches its
baseline set, including additional unchanged-file links. Rustdoc is therefore
unqualified as a pre-existing workspace condition, not a candidate-file
regression. The final review reports the candidate production diff approved
for prospective comparison.

The quality receipt is the pre-measurement snapshot; its false measurement and
adoption fields describe that earlier stage, not the completed comparison below.

## Frozen prospective comparison

The proposed gates are frozen in
`/home/zhuhe/litchi-performance-0943/tooling/comparison-protocol.json`.
The protocol uses the two bound fixtures `docnohf` (26,112 bytes,
SHA-256 `45e5df073f34314da6f39d2dad119fb2ef23470878fd2df67f632864cd92ea48`)
and `docfloat` (335,360 bytes,
SHA-256 `41cc0cd8f2d7390266f844f5bddbcd29d08ac5338abc33832d9c9a6ed5d0f85d`),
with 1, 2, 8, and 32 edits and `zero`, `default`, and `release` retention.
The native lane is 192 serial children (12 paired rounds, CPU 14, five
warmups and 50 samples per child); the allocator lane is 96 serial children
(two reversed repetitions, no warmups and one sample per child). The combined
schedule has 288 children.

Adoption requires all of the following:

- every one of the eight native fixture/edit cells has a 10,000-draw,
  seed-943 paired bootstrap lower bound of at least
  `0.9523809523809523` (`1/1.05`);
- at least one fixture at 32 edits with default retention has candidate
  `allocated_bytes <= baseline * 0.95`;
- every allocator pair stays at or below `baseline * 1.01` allocated bytes and
  `baseline * 1.05` peak-live bytes;
- paired output digests and bytes, explicit capacities and retained bytes, and
  all semantic/correctness oracle fields match baseline;
- p50, p95, p99, and RSS increases above five percent remain explicit review
  flags. There is no automatic waiver.

## Measured result

The fixed run used the AMD EPYC 9R45 x86-64 host (32 logical CPUs), CPU 14,
the source and release binaries above, and the unchanged 0917 end-to-end
DOC probe. Native results measure that probe's operation boundary; process RSS
also includes preparation and validation. Allocator timing is not interpreted.

| Fixture | Edits | Median paired baseline/candidate p50 ratio | Paired bootstrap 95% interval | Latency gate |
| --- | ---: | ---: | --- | --- |
| docnohf | 1 | 0.999244 | [0.992864, 1.006044] | pass |
| docnohf | 2 | 0.996873 | [0.993920, 1.012209] | pass |
| docnohf | 8 | 0.993148 | [0.987841, 1.005572] | pass |
| docnohf | 32 | 0.998626 | [0.993963, 1.000137] | pass |
| docfloat | 1 | 1.010944 | [0.972274, 1.090326] | pass |
| docfloat | 2 | 0.961578 | [0.895700, 1.103816] | **fail** |
| docfloat | 8 | 1.010868 | [0.989218, 1.070720] | pass |
| docfloat | 32 | 1.015643 | [1.005130, 1.026529] | pass |

Ratios above one favor the candidate. These are medians of paired-round
ratios, not ratios of pooled medians. The required lower bound is 0.952381.

| 32-edit/default fixture | Baseline requested bytes | Candidate requested bytes | Reduction |
| --- | ---: | ---: | ---: |
| docnohf | 16,781,690 | 16,650,738 | 0.780327% |
| docfloat | 164,233,879 | 156,064,143 | 4.974452% |

Both allocator repetitions agree exactly on those counts. Gross requested
bytes describe allocator requests, not physical-copy traffic or RSS savings.
All 48 allocator pairs pass the +1% gross-request and +5% peak-live safety
limits. At 32/default, candidate peak live bytes increase by 80 bytes on each
fixture; retained bytes are unchanged. Every paired explicit capacity and
semantic field matches after independently validated input-path normalization.

All adverse individual flags remain in the committed reduction: 10 native
p50, 11 p95, 13 p99, and one native RSS flag exceed +5%; no allocator RSS
flag does. Worst p50 is +23.6812% (FloatingPictures, two edits, round 9);
worst p95/p99 are +22.1812%/+22.8397% (FloatingPictures, one edit, round 7).
The RSS flag is +5.3770% (FloatingPictures, eight edits, round 2). These
observations are retained without attributing them to a proven mechanism or
waiving them. The failed primary allocation and latency gates decide rejection.

## Qualification and evidence custody

Before any formal measurement, 64 actual qualification children passed the
0917 validator. The qualification process then exited 2 because pair comparison
incorrectly required the two isolated exports' absolute `input` paths to be
identical. The original reports and terminal remain intact. An offline replay
revalidated all 64 reports and all 32 pairs, removing only that already-validated
provenance path from comparison. Source digests, sizes, output, semantic fields,
and capacities remain exact. No workload was retried. The formal runner applies
this same normalization; the numerical gates were not changed.

Runner preparation also caught and corrected serial timestamp checks, the
gate conjunction, packet relocation, qualification-directory binding, terminal
state validation, and duplicate-key parsing. Synthetic attempts are retained
externally and are not performance evidence. The final synthetic matrix covers
288 children; moved-packet replay hides the original packet, and controls reject
changed terminal state, argv, environment, source, binary, and duplicate JSON
keys. Positive and negative gate-conjunction controls pass.

The formal outer process exits 0 before reduction, with 192 native and 96
allocator receipts. A moved replay reproduced every reduction field with the
original packet unavailable: 1,730 packet files, 46,323,941 bytes. Its temporary
copy was removed after verification. An independent formal audit validated all 288 children and recomputed every
paired gate and review flag; its reduction is byte-identical to the root
reduction (SHA-256 `645b8a13a696c41430ed2f92b81c273f3b8d9b0a545531bab0b0520d0f4a0f17`).
The independent build audit checks all
10,798 files in each bound source tree, exact build argv/environments/manifests/
locks, successful terminals, logs, and all four measured binary hashes.

The baseline source was restored before its first build to include ten
`target-*` named fuzz corpus files previously omitted by the 0942 archive
selector. The authoritative baseline v2 manifest binds that complete source.
0943 archive selection excludes only named top-level build-cache directories
and preserves every nested `target-*` source name.


## ADR invariant mapping

- **ADR 0003 (snapshots, edits, and patches):** immutable snapshots remain
  shareable; `Arc::make_mut` preserves clone-first isolation, and both success
  and late refusal leave the source snapshot untouched.
- **ADR 0005 (I/O, memory, and measured performance):** the candidate removes
  a private duplicate on unchanged paths without changing finite limits,
  charging, typed resource errors, or the requirement for representative
  measured evidence.
- **ADR 0006 (validation, security, and compatibility):** publication and
  preservation checks remain in place; failed validation and limit refusal are
  tested as non-mutating operations with typed errors.
- **ADR 0008 (migration and verification):** the source remains continuously
  buildable under the scoped test, Clippy, and formatting checks, while the
  existing rustdoc baseline is recorded rather than silently waived.
- **ADRs 0010, 0011, and 0024 (owner boundaries and current topology):** this
  is a private `litchi-doc` editor change in the existing package owner. No
  facade edge, physical package owner, dependency layer, or public ownership
  contract moves.

The candidate is rejected and no architectural or public API change is made.
Future work should investigate other current operation-local costs identified
by 0942, or the post-0905 XLSX publication path; this cohort does not justify
resubmitting unchanged Data sharing, bundling an unmeasured change, or lowering
the thresholds.

## Reproduction and retained artifacts

The committed [formal reduction](results/change-0943/formal-reduction.json)
contains every paired metric, confidence interval, and review flag. The
[manifest](results/change-0943/formal-manifest.json) binds the runner, validator,
source manifests, build receipts, binaries, and qualification replay.

Large source listings, raw packets, binaries, failed preparation attempts, and
logs stay outside the repository. See the archive receipt appended below.
To reconstruct the rejected candidate in a separate baseline checkout, apply
`git apply -p5 docs/performance/results/change-0943/candidate.patch`; a dry-run
`--check` passes against the unchanged production file. Keep this isolated from
production because the measured candidate was rejected.

Rebuild with `run-build-v1.py` and the recorded source receipt/target/binary
arguments; each exact Cargo invocation is in the build start and terminal
receipts. Replay the retained formal packet with:

```sh
python3 -B /home/zhuhe/litchi-performance-0943/tooling/comparison-runner.py \
  --reduce --manifest /home/zhuhe/litchi-performance-0943/formal-authorized-v1.json \
  --lane all --output /home/zhuhe/litchi-performance-0943/reductions/replay-new.json
```

Fresh captures require a new evidence destination and source/binary-bound
manifest; the runner refuses to replace existing children. Repeating a run
cannot replace this rejected cohort.

The verified external archive is
`/home/zhuhe/litchi-performance-0943-archive/litchi-performance-0943-v1.tar.zst`:
377,836,889 bytes, SHA-256
`072b45ce21ce9fa859960ee0cd00904edc38c1d75d5b0903421a95fe370a65e3`.
It preserves 78,847 selected files (2,277,926,535 logical bytes), including raw
qualification/formal/synthetic attempts, source exports and six bound binaries.
The [archive summary](results/change-0943/archive-summary.json) binds its full
file listing and archive, verify, cleanup, and source-closure receipts, with
exact regeneration commands. Source listings exceed 1 MiB and remain external.

Archive extraction and every source manifest were verified before cleanup.
Cleanup removed 8,435 build-cache paths, the tooling Python cache, and the
verified extraction. A second check confirmed all 10,798 files in each of the
five complete source exports remained intact. Bound binaries remain outside
the repository. No binary or evidence file of 1 MiB or more is committed.
