# 0921 — source-bound recapture misses the latency confidence gate

The candidate is **not adopted**. It reduces allocation bytes and improves
every paired median timing ratio, but the FloatingPictures two-edit cell
narrowly misses the frozen confidence-bound threshold. Production remains
unchanged; no selective recapture or relaxed gate was used.

## Mechanism

[0919](0919-doc-allocation-attribution.md) identified recapture as a large
allocation owner. The existing publication path reopened the rendered CFB,
read every stream into a new vector, copied it into an Arc, and only then
replaced equal payloads with staged Arcs. This candidate compares each stream
against its staged payload directly in the immutable rendered source. Equal
streams retain the staged allocation; changed or absent streams use ordinary
readback.

CFB owns the comparison and binds it to the exact `Cursor<&[u8]>` parsed by
`OleFile`. One comparison session spans the directory traversal, sharing the
root mini-stream validation. Full chain, tail, readability and zero-fill
checks still run, including after a byte mismatch. Directory metadata is
captured afresh. Protection checks, resource limits, final reopen, discovery,
public semantics and output capacities remain in place.

The candidate is independent of the rejected
[0920 output reservation](0920-cfb-output-reservation-rejected.md); that patch
is not combined with this one. The unchanged 0917 public DOC probe measures
the same two fixtures, four edit counts and retention routes.

## Diagnosis of the earlier rejection

The separate [profile packet](results/change-0921/profile-results.json) uses
retained **0920** binaries. Scoped Callgrind instruction counts fall in every
profiled cell. They provide no evidence of increased total operation
instruction work, but do not explain the rejected candidate's native latency.
Whole-process perf samples include
copying, allocation, parsing and hashing, but their call stacks lack the
operation frames needed for reliable filtering. Page faults vary widely and
syscall counts do not show a consistent increase. These observations do not
identify a causal explanation or reverse the 0920 rejection.

Raw profiles and every executable remain outside the repository. The profile
packet records their sizes, hashes and reproduction commands, including
failed setup attempts and instrumentation limitations.

## Frozen comparison

The [protocol](results/change-0921/comparison-protocol.json),
[readiness review](results/change-0921/capture-readiness.json), and
[final manifest](results/change-0921/comparison-manifest.json) bind the final
six-file patch, fresh builds, unchanged probe, fixtures and comparison tools.
Both builds use Rust 1.95.0 and matching release settings. Root verified all
938 source files per version, all build artifacts, 157 compiler fingerprints,
and the matching 144-package dependency graphs. The fresh baseline binaries
reproduce their previous SHA-256 values exactly.

Final qualification replays 96 reports across both fixtures, four edit counts,
three retention routes and two lanes. All 48 version pairs have exactly equal
capacities. Eight negative comparator controls are rejected; the exact-match
positive control passes. A complete synthetic reduction for the planned
population stays below the 1 MiB evidence limit.

Native timing uses 12 paired rounds, five warmups and 50 samples per child:
192 children and 9,600 measured lifecycle samples, pinned to CPU 14 on the
same AMD EPYC 9R45 host. Version and cell order reverse prospectively. A
separate System allocator lane uses 96 children across zero/default/release
retention. Source I/O, correctness oracles and report generation are outside
the native timer. This is in-memory lifecycle evidence, not cold-cache,
range-provider or concurrent coverage.

Adoption requires at least 5% fewer requested allocation bytes in one
32-edit default fixture, no allocator cell exceeding the 1% requested-byte
or 5% peak-live increase limits, and every native cell's paired-bootstrap
95% lower ratio bound at least `1/1.05`. Bootstrap seed 921 and 10,000 draws
were fixed before capture. Exact output, semantic, retention and capacity
identity are separate mandatory gates.

## Results and rejection

All 288 children exit successfully and pass the unchanged correctness
validator. Every explicit capacity field remains equal. Allocation and
peak-live guardrails pass in every cell and repetition.

| Fixture | Edits | Paired p50 ratio | 95% interval | Latency gate |
| --- | ---: | ---: | --- | --- |
| NoHeadFoot | 1 | 1.009550 | 0.999496–1.022178 | pass |
| NoHeadFoot | 2 | 1.013729 | 1.008941–1.019148 | pass |
| NoHeadFoot | 8 | 1.039794 | 1.037308–1.043071 | pass |
| NoHeadFoot | 32 | 1.060335 | 1.051338–1.064176 | pass |
| FloatingPictures | 1 | 1.021267 | 0.986590–1.064541 | pass |
| FloatingPictures | 2 | 1.016481 | 0.952076–1.155981 | **fail** |
| FloatingPictures | 8 | 1.069966 | 1.043387–1.086826 | pass |
| FloatingPictures | 32 | 1.094489 | 1.090337–1.105749 | pass |

Ratios are baseline/candidate, so values above one favor the candidate.
FloatingPictures at two edits has lower bound `0.9520756186`, below the
required `0.9523809524`. Its improved median does not satisfy the prospective
uncertainty gate. This is not evidence that every workload regresses, and it
does not authorize dropping the failing cell.

At 32 edits with default retention, both allocator repetitions agree:

| Fixture | Requested bytes before → candidate | Reduction | Peak-live bytes before → candidate |
| --- | --- | ---: | --- |
| NoHeadFoot | 16,781,690 → 14,114,634 | 15.89% | 642,194 → 642,194 |
| FloatingPictures | 164,233,879 → 131,948,263 | 19.66% | 7,635,361 → 7,635,361 |

The [complete reduction](results/change-0921/comparison-reduced.json) retains
all rounds, p50/p95/p99, dispersion, RSS, allocation and capacity comparisons.
Among 96 native pairs, five p50, five p95, nine p99 and one RSS row exceed the
5% adverse review threshold. Three allocator RSS rows are flagged. These
categories overlap; no row was omitted. Requested-byte reductions do not
establish lower peak memory or RSS.

The candidate remains useful evidence for a future prospectively designed
experiment, but this batch makes no production speedup claim. The exact
[patch](results/change-0921/comparison-receipts/candidate.patch) is retained
against base `bb19ca74bd52044fbd2c1c7cc9875e975bb334c0`, with SHA-256
`7705c8ec7f12a5a5083b2ff29408451b0ee2afbbf8a25bb827c01b885aece67d`.

## Quality and source review

The [independent review](results/change-0921/shared-owner-quality.md) found
no semantic defect. Scoped functional suites passed 5,059 tests: CFB 479,
OLE common 199 plus 30 object tests, DOC 1,347, XLS 1,704 and PPT 1,300.
Those broad runs preceded a final lint-only change to lifetime/dereference
syntax and a test initializer. Final focused comparer tests, warning-denied
Clippy, compilation, feature-write check and targeted formatting pass.
The public source adapter matches ordinary readback across 215 OLE2 files
and 9,139 expectations.

Private comparer tests cover malformed chains, short tails and root failures;
a direct public-adapter synthetic corruption test remains a nonblocking
coverage gap. The candidate keeps the same underlying comparer. Manifest
inventories and dependency-direction checks pass. Existing workspace format
drift outside the candidate remains; no full-workspace or iWork verification
is claimed.

The initial candidate build receipt queried the compiler version from the
launcher directory. Target compiler metadata and fingerprints established
the actual Rust 1.95.0 build; the corrected receipt and original are retained.
Final lint cleanup was separately rebuilt and qualified. Preliminary
qualification and the scratch freeze validation are explicitly superseded;
only the manifest created after root readiness governs this comparison.

Both [offline replay paths](results/change-0921/final-replay.json) verify all
288 reports and reproduce the rejection. The canonical report-set SHA-256 is
`4f867b4a6f7aafe0119403065ef7b3f7db8201c247f1f5e358aa936ba55cef0d`.
The [independent audit](results/change-0921/final-audit.md) verifies all
288 child receipts, 864 packet artifacts and 864 external artifacts, and
recomputes the 560,529-byte reduction exactly. A copied packet also replays
successfully. [Cleanup](results/change-0921/cleanup.json)
removed 27,362,174,784 logical bytes of generated intermediates and scratch
files while retaining measured artifacts.

## Reproduction

Replay the retained comparison without starting benchmark children:

```sh
python3 -B docs/performance/results/change-0921/comparison.py replay
python3 -B docs/performance/results/change-0921/comparison-replay.py \
  --packet docs/performance/results/change-0921
```

The packet binds the exact base revision plus patch, source archive recipe,
build commands and environment, binaries, fixtures and every child command.
Raw artifacts and compiled executables remain outside Git; every committed
evidence file is smaller than 1 MiB. Cleanup removes generated intermediates
while retaining measured sources, binaries, fingerprints and evidence.
