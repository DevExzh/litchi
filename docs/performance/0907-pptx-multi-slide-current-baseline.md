# 0907 — current PPTX multi-slide guarded workflow baseline

A fresh same-binary comparison retains a large cost separation between the
existing eager and source-backed PPTX multi-slide workflows: pooled p50 is
**263.429593 ms eager** versus **7.841286 ms source-backed**, an observed
ratio of **33.5952**. This is a descriptive comparison of two existing API
paths, not a before/after optimization or a release speedup. No production or
harness implementation changes are adopted in this batch.

## Current source and capture

The build uses frozen commit `3d3650ac03be2cd8ba923a5a6e940e84b14f5eab`,
with the two retained dependency lockfile overlays. The
[source receipt](results/change-0907/source.json) binds all 9,350 snapshot files, including the lockfiles and ownership marker;
Git archive reconstruction and raw-report replay pass. This is new timing for
that complete snapshot, including its ZIP changes, rather than a reuse of the
0905 executable or historical PPTX timings. The unrelated shared iWork
cherry-pick is outside this isolated source tree.

Rust 1.95.0 builds offline and locked, release opt-level 3, fat LTO, one codegen
unit, debug level 1, and two Cargo jobs. Binary SHA-256:
`e0a7707dcd849c92e01c8dc9115d61414fba10584bb8415e442f4c225828af84` (264,596,744 bytes).
The initial build omitted two compile-time fixtures and failed. Both fixtures
were then restored from the same commit; the retained
[v2 build receipt](results/change-0907/native-build-v2-run.json) passes.
The initial source inventory and failure log remain available.

Four fresh child processes run serially on CPU 12 in eager/source/source/eager
order, with `--workers 1`, 30 warmups and 200 retained samples each. Capture
ran from 2026-10-04 06:59:52.908831 UTC through 07:03:07.842130 UTC.
The host reports AMD EPYC 9R45, 32 CPUs and Linux 7.0.0-1012-aws.
The machine is shared; CPU pinning is not exclusive-host evidence. The
[environment](results/change-0907/environment.json),
[capture plan](results/change-0907/capture-plan.json), per-child commands,
exit codes, start/end times and raw samples are retained.

## Timing and guard boundary

Both modes edit all eight text boxes on slide positions
`[0, 28, 57, 85, 114, 142, 171, 199]` in a 200-slide generated presentation.
The source contains eight deterministic 2 MiB media Parts, 229 logical OPC
Parts and 445 ZIP members. Archive bytes: 17,017,139; SHA-256:
`61b2b99083ca27ebd37955db600955e3f41289b93dba71951983164239eff757`.
Every leg emits 17,017,145 bytes with SHA-256:
`23d6d7b8dd433ff453307ee56485efadc34d6b18b455a6858d9f565c4b1b6cd9`.
These identities match the historical fixture; only identity is compared with
history, never latency.

The clock surrounds `publish_pptx_multi_slide_batch_edit`. It includes open,
staging, commit, publication and harness patch checks. Source-backed mode
checks forward/inverse application before publication, then inverse acceptance
and stale forward refusal afterward. Eager mode applies inverse and forward
successfully before serialization. These guard sequences differ and are not
subtracted. Expected-output construction, source/sink allocation and reserve,
full semantic reopen, digesting and preservation verification are outside the
clock. The [protocol review](results/change-0907/protocol-review.md) gives exact
source boundaries and limitations.

For every iteration, built-in gates check mode-specific exact expected bytes,
selected and unselected text, slide order, topology, Part content types and
relationships, untouched logical Part payloads and all media payloads. Every
retained sample has nonzero ordinary reads, the required materialization count,
and sink writes at most 64 KiB. The source snapshot and binary hashes remain
unchanged. This harness does not independently census every raw ZIP header,
comment or member-order field, so this packet establishes its semantic/OPC
preservation contract, not a new raw-container fidelity claim.

## Results

All times below are milliseconds. Each row contains 200 retained samples.

| Leg | p50 | p95 | p99 | max |
| --- | ---: | ---: | ---: | ---: |
| eager-a | 263.289679 | 264.759446 | 278.711065 | 313.334012 |
| source-a | 7.712175 | 8.022402 | 8.174263 | 8.323143 |
| source-b | 7.927396 | 9.035897 | 10.655026 | 15.529992 |
| eager-b | 263.581606 | 265.864489 | 266.926522 | 359.348308 |

Pooled eager p95/p99 are 265.490578/270.110281 ms; pooled source-backed
p95/p99 are 8.550684/9.838051 ms. Pair p50 eager/source ratios are 34.1395
and 33.2495. Source-B has a higher tail than source-A; all samples are kept.
The [analysis](results/change-0907/analysis.json) reports per-leg and pooled
statistics. Two pairs support this descriptive baseline, not a confidence
interval for an optimization claim. The harness's raw mean intervals remain
in the reports; they are not intervals for the cross-path ratio.

| Per-sample logical counter | Eager | Source-backed |
| --- | ---: | ---: |
| Reported ordinary Part materializations | 229 | 9 |
| All source read calls | 1,864 | 3,372 |
| All source read bytes | 17,030,811 | 17,229,633 |
| Ordinary payload read calls | 741 | 486 |
| Ordinary payload read bytes | 16,907,750 | 16,909,038 |
| Sink write calls | 260 | 2,053 |
| Sink accepted bytes | 17,017,145 | 17,017,145 |
| Largest sink write | 65,536 | 65,536 |

These counters are constant within each path. The source-backed cache load
count is sampled after commit and pre-publication replay, before publication;
the eager count is the fully materialized OPC Part count. Neither counts the
untimed verification reopen. Fewer reported materializations do not
mean fewer total reads, fewer output writes, fewer allocated bytes, or less
physical I/O. The source is in-memory `ReadAt`; no device, cold-cache, remote,
allocator, copy-volume or RSS measurement is made. Source vectors are in
iteration order while elapsed values are sorted; no positional correlation
is inferred without applying `sample_order`.

## Verification and next action

The newly built executable lists all 534 registered selectors in `--help`,
providing an actual binary check for the prior help fix. Offline analyzer
checks accept the real report and reject seven targeted corruptions. Initial
mutation-test diagnostic-substring mismatches are documented; the analyzer
already rejected both malformed inputs correctly. Analysis replay and frozen
source reconstruction pass. No full Cargo test suite was run for this
measurement-only batch.

The [source review](results/change-0907/source-opportunities.md) identifies
repeated source comparisons, overlay planning scans and distinct XML validation
passes as hypotheses. Existing phase diagnostics do not cover this multi-slide
runner. The next step is matched route-specific phase attribution for eager
open/materialization/serialization and source-backed editing/publication,
with harness replay separately timed and all current correctness gates
retained. This is needed before choosing a production candidate; the total
ratio and materialization counts alone cannot identify a causal hotspot.
The DOCX existing-section baseline and legacy real-file CRUD qualifications
remain outstanding.

[0906 errata](results/change-0907/0906-clarifications.json) correct the audit's
`--output` spelling to `--json` and clarify stale forward refusal. The original
0906 packet remains unchanged. The broader OLE2/OOXML goal remains active;
iWork is excluded and ODF deferred.

## Reproduction and cleanup

Offline verification does not require the removed source/build directories:

```sh
python3 -B docs/performance/results/change-0907/replay.py --output /tmp/0907-replay-new.json
```

The output must not exist. For a fresh native capture, `restore.py` recreates
the exact source and lockfile inputs in absent owned directories. Use the
retained v2 build command/environment, then the capture plan's four commands
with new output paths; never replace accepted raw reports. `capture.py` and
`command.py` are create-only packet runners, so rerunning them requires a fresh
packet location with the input manifests/helpers copied there.

[Cleanup](results/change-0907/cleanup.json),
[post-cleanup replay](results/change-0907/replay-after-cleanup.json) and
[independent result review](results/change-0907/result-review.md) are retained.
A [root clarification](results/change-0907/review-clarifications.json) corrects
the review counter-sampling phrase; the original review is also retained. The compiled executable is retained locally under the packet's ignored
`binaries/` directory for later diagnostics; it is not part of the Git commit.
No worktree was created. Target-branch integration remains pending while the
shared checkout has unrelated cherry-pick conflicts.
