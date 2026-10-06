# 0920 — exact CFB output reservation rejected by latency gate

The isolated candidate reduces requested allocation bytes, but is **not
adopted**. All four FloatingPictures edit-count cells fail the frozen latency
confidence-bound gate. Production code remains unchanged; no selective
recapture, omitted adverse row, or relaxed gate was used.

## Mechanism and source

[0919](0919-doc-allocation-attribution.md) identified growing CFB output buffers
as a large requested-byte owner. This experiment reserves the checked output
length from the already validated reuse plan once, before emission. CFB owns
the reservation and typed failure; the common OLE renderer delegates to that
owned-output path. Generic sinks, declined-plan fallback, final reopen,
complete stream readback and semantic publication checks remain intact.

The [exact patch](results/change-0920/comparison-receipts/candidate.patch) changes CFB core/layout,
the common OLE codec, and an owned-output test. It is retained as experimental
evidence, not applied to production. Both versions use base revision
`bb19ca74bd52044fbd2c1c7cc9875e975bb334c0`; the candidate is identified by that
base plus patch SHA-256
`0441e3d739155ab4bbe05cf2dde02c837d10c48b3162352f6b3327055ab929e3`.
The source manifest and [final source review](results/change-0920/source-review-final.md)
resolve the earlier prototype admission questions without reconstructing
identical source. No ADR semantic exception was needed.

## Frozen comparison

The [protocol](results/change-0920/comparison-protocol.json),
[readiness receipt](results/change-0920/capture-readiness.json), and
[manifest](results/change-0920/comparison-manifest.json) bind source, builds,
fixtures, runner, validator, controls and review. Root rechecked 938 baseline
and 942 candidate files, both build artifact arrays, 157 compiler/profile/
feature fingerprints, and matching normalized dependency graphs. Final
qualification replay validates 96 retained reports and all 48 version pairs.

Machine: AMD EPYC 9R45, 32 logical CPUs, 129,447,072 KiB reported memory,
Linux 7.0.0-1013-aws x86_64, glibc 2.43, Rust 1.95.0. Every benchmark child is
pinned to CPU 14. Release settings, binary identities, allocator lane and
historical environment-recording limits are bound in the packet. Source is
in memory; input file I/O and correctness oracles are outside the timer.
This is not cold-cache, range-provider or concurrent evidence.

Native timing uses 12 paired rounds, two fixtures, 1/2/8/32 edits, default
retention, five warmups and 50 samples per child: **192 children and 9,600
measured lifecycle samples**. Version and cell order reverse prospectively.
The separate allocator lane uses two repetitions over zero/default/release
retention: **96 children**. All 288 children exit successfully and pass the
unchanged 0917 correctness validator.

## Results and rejection

Ratios below are the median of twelve paired baseline-p50/candidate-p50
ratios. Values above one favor the candidate. Intervals use 10,000 bootstrap
draws, seed 920, and the matched round as the resampling unit. Every lower
bound must be at least `1/1.05 = 0.952381`; all four FloatingPictures cells
fail. This includes the eight-edit cell, whose median improves but whose
uncertainty does not clear the guardrail.

| Fixture | Edits | Paired p50 ratio | 95% interval | Latency gate |
| --- | ---: | ---: | --- | --- |
| NoHeadFoot | 1 | 1.0087 | 1.0043–1.0126 | pass |
| NoHeadFoot | 2 | 1.0127 | 1.0083–1.0152 | pass |
| NoHeadFoot | 8 | 1.0162 | 1.0085–1.0172 | pass |
| NoHeadFoot | 32 | 1.0258 | 1.0207–1.0272 | pass |
| FloatingPictures | 1 | 0.9294 | 0.9072–0.9824 | fail |
| FloatingPictures | 2 | 0.9711 | 0.9365–0.9978 | fail |
| FloatingPictures | 8 | 1.0169 | 0.9485–1.0425 | fail |
| FloatingPictures | 32 | 0.9418 | 0.9264–0.9587 | fail |

At 32 edits with default retention, both allocator repetitions give identical
values:

| Fixture | Requested bytes before → candidate | Reduction | Peak-live bytes before → candidate |
| --- | --- | ---: | --- |
| NoHeadFoot | 16,781,690 → 14,251,898 | 15.07% | 642,194 → 631,954 |
| FloatingPictures | 164,233,879 → 139,873,431 | 14.83% | 7,635,361 → 7,146,401 |

All allocator cells stay within the 1% requested-byte and 5% peak-live increase
limits, and every explicitly admitted capacity field is nonincreasing.
Those successes do not override the latency gate. Lower requested allocation
bytes do not prove lower allocator time, fewer physical copies, or less RSS.

The [machine-readable reduction](results/change-0920/comparison-reduced.json)
retains every round, p50/p95/p99, dispersion, RSS, capacity comparison and
allocator cell. Of 96 native pairs, 23 p50, 22 p95, 20 p99 and one RSS comparison
exceed their 5% adverse thresholds; these overlap and must not be added.
No allocator RSS row is flagged. No geometric mean is used to hide a failure.

## Verification and limitations

[Shared-owner checks](results/change-0920/new-quality.json) against the exact
candidate overlay pass 554 CFB, 285 OLE common, 1,361 DOC, 1,705 XLS and 1,314
PPT tests: **5,219 passes**. Scoped warning-denied Clippy, consumer rustdoc,
compilation and candidate formatting pass. Manifest inventory and dependency
direction checks pass. The full source-topology checker was stopped during
excluded iWork audits; no full-checker pass is claimed.

Two unrelated quality failures remain explicit: existing DOC public rustdoc
links to private items, and workspace formatting drift outside the candidate.
The first XLS attempt included an accidental nested fixture symlink; removing
that setup artifact restored the intended 126-file XLS corpus, and the exact
command then passed. This setup correction did not affect benchmark fixtures
or permit a performance recapture.

Both final replay paths verify all 288 reports and reproduce the rejection;
[their receipt](results/change-0920/final-replay.json) binds the final artifacts.
The final reduction is 555,046 bytes, below the evidence limit. Earlier
prototype reviews/control receipts are historical: final controls are
[capacity-controls-final.json](results/change-0920/capacity-controls-final.json)
and final qualification is
[qualification-replay-final.json](results/change-0920/qualification-replay-final.json).
The earlier synthetic tooling replay is separate from the real comparison.

## Reproduction and cleanup

From the repository root, replay the retained comparison with:

```sh
python3 -B docs/performance/results/change-0920/comparison.py replay
python3 -B docs/performance/results/change-0920/comparison-replay.py --packet docs/performance/results/change-0920
```

For a fresh capture, rebuild the source-bound baseline and candidate using the
commands and explicit environment in the retained build receipts; the
candidate is the exact base plus retained patch. Use a fresh packet/output
destination and freeze those binaries before running both lanes and reducing.
Never overwrite this frozen comparison. Child receipts retain exact commands,
fixture hashes, process metrics, timestamps and output identities. Executables,
large source archives and raw captures stay external under the recorded
`/tmp/litchi-performance-0920*` paths; compressed replay reports are retained
in the packet. Each large artifact has its size/hash and regeneration command.

[Cleanup](results/change-0920/cleanup.json) removes disposable dry-run packets,
Python caches and candidate build intermediates, releasing 857,043,378 logical
bytes while rechecking all retained candidate build artifacts. The tester also
removed its disposable build target. Source trees, measured binaries,
fingerprints and evidence remain available for replay; unrelated user work is
untouched.

## Next evidence

Do not adopt this reservation path or tune allocator settings to rescue this
capture. A separate operation-local CPU/syscall attribution would be needed to
explain why smaller allocations regress FloatingPictures; allocation placement
and page-fault effects are hypotheses, not findings. The independently measured
source-bound recapture comparison from 0919 remains another candidate, but must
retain full CFB validation, directory capture and required semantic readback.
Broader CRUD, producer, cold/range, streaming and concurrency requirements in
`docs/GOAL.md` remain open. iWork remains excluded.
