# 0922 — borrowed FIB lookup for saved-selection updates

## Measured opportunity

The [independent selection review](results/change-0922/candidate-selection.md)
identifies three allocation traces in the retained
[0919 profile](0919-doc-allocation-attribution.md). Two saved-selection
preflight traces request 4,069,239 bytes each; the apply trace requests
4,256,256 bytes. Together they account for **12,394,734 requested bytes**
across 32 FloatingPictures edits, about 7.55% of lifecycle allocation bytes.
A similarly sized commit-handoff trace is unrelated and excluded. These are
allocation attributions, not native latency estimates.

`FileInformationBlock::parse_at` copies the complete WordDocument suffix into
owned storage. The selected saved-selection paths need only a validated FIB
table pointer. Earlier [0596 work](0596-doc-eager-open-terms.md) shared the
stream on ordinary document opening but deliberately preserved the public
owned parser and `raw_data()` suffix contract.

## Candidate mechanism and boundaries

The candidate factors the existing base validation, pointer layout and pointer
read into shared private helpers in the FIB owner. A crate-private lookup reads
one pointer from borrowed WordDocument bytes. Saved-selection parsing and
range checks share their existing pointer-based validation with the owned
route. Only splice preparation and application in the revision editor use
the new borrowed route.

The public owned FIB parser, offset handling and complete `raw_data()` suffix
remain unchanged. The lookup retains base-size and magic checks, the legacy
74-pair layout, counted-array bounds and the requirement that the complete
declared array be present. Missing/zero-length selections remain distinct
from unavailable or malformed edit ranges. Selsf bytes, CP remapping,
ambiguity refusals, clone-first atomicity, resource limits, protection,
package rendering, reopen, recapture and final publication stay intact.

This candidate does not incorporate the rejected 0918, 0920 or 0921 changes.
The baseline source remains `bb19ca74bd52044fbd2c1c7cc9875e975bb334c0`.
[Baseline reuse](results/change-0922/baseline-reuse.json) verifies that current
production/build paths are unchanged and rechecks the exact recently rebuilt
baseline binaries.

## Frozen measurement

The comparison uses the unchanged 0917 public lifecycle probe on NoHeadFoot
and FloatingPictures at 1/2/8/32 edits. Native timing has 12 paired rounds,
five warmups and 50 samples per child: 192 children and 9,600 samples. The
separate allocator lane has 96 children across zero/default/release retention.
All children are pinned to CPU 14; version and cell order reverse according
to the frozen schedule.

The gates remain exact output/semantic/retention/capacity identity; at least
5% fewer requested bytes for one 32-edit default fixture; no allocator cell
exceeding 1% requested-byte or 5% peak-live growth; and every native cell's
paired-bootstrap 95% lower ratio bound at least `1/1.05`. Bootstrap seed 922
and 10,000 draws are fixed prospectively. No failed cell may be removed or
selectively recaptured.

## Result: reject adoption

All 192 native and 96 allocator children passed the unchanged correctness
validator, with exact output, semantic, retention and capacity identity.
Allocation and peak-live gates passed. The candidate nevertheless fails the
mandatory native latency confidence gate for **FloatingPictures at one and
two edits**. The lower bounds are `0.9395640088249162` and
`0.9002944110515873`, below `1/1.05 = 0.9523809523809523`.
No production source is applied and no comparative cell is rerun or removed.

The paired baseline/candidate p50 ratios are shown below; larger ratios favor
the candidate. The intervals resample matched rounds, not individual samples.

| Fixture | Edits | Median ratio | 95% interval | Gate |
| --- | ---: | ---: | --- | --- |
| docnohf | 1 | 1.008729 | 1.000597–1.012620 | pass |
| docnohf | 2 | 1.014050 | 1.006416–1.023314 | pass |
| docnohf | 8 | 1.007554 | 1.002919–1.015111 | pass |
| docnohf | 32 | 1.024226 | 1.018722–1.031158 | pass |
| docfloat | 1 | 1.003758 | 0.939564–1.021444 | **fail** |
| docfloat | 2 | 0.966379 | 0.900294–1.003337 | **fail** |
| docfloat | 8 | 1.007475 | 0.966441–1.029702 | pass |
| docfloat | 32 | 1.011190 | 0.999994–1.019343 | pass |

At 32 edits with default retention, FloatingPictures requests
**164,233,879 → 151,835,305 bytes (−7.55%)** and NoHeadFoot requests
**16,781,690 → 14,828,154 bytes (−11.64%)**. The Float reduction of
12,398,574 bytes is consistent with the three attributed suffix-copy traces
plus ownership overhead; attribution alone does not establish a latency cause.
Peak-live bytes and all explicit capacities remain unchanged. None of the
48 allocator pairs exceeds allocation or peak-live limits.

All adverse observations remain in the reduction: **11 p50, 14 p95, 21 p99**
native pair flags, **two native RSS** flags, and no allocator RSS flags.
The complete [reduction](results/change-0922/comparison-reduced.json)
is 560,459 bytes, SHA-256
`cfe89f3da7b8b989697cb623932398b100b38ccf04e41a6514916a66d88ea846`.
It reports `adoption: false`; allocation savings do not override failed latency
gates. This batch closes the experiment, not the broader GOAL requirements
or the unproven 10× objective.

## Source review and quality

The final isolated patch changes only `parts/fib.rs`, `parts/saved_selection.rs`
and `tracked_revision/package.rs` in `litchi-doc`. Its 19,215-byte patch has
SHA-256 `dd7f0bafb81c23a3762e74246eccf991b56ff1f7e62816087e27441db636bd84`;
the 938-file measured source manifest is
`e2a8893bc67999ea20aa5320d8bef3484b41ba50a4cfafa26a72916efa2ff2ea`.
Root independently checked both source closures and all six artifacts in each
build receipt. The unchanged standalone probe lock is separate from the
regenerated external quality-workspace lock; no lockfile is part of this patch.

[Independent owner review](results/change-0922/shared-owner-quality.md)
approved the final patch. All **1,349 DOC library/integration tests passed**,
with two ignored. Differential tests cover malformed selection records,
missing and zero-length pointers, bounds, and nonzero FIB offsets. Scoped
formatting, Clippy with warnings denied, and manifest/dependency-direction
checks pass. Workspace formatting still reports unrelated existing files;
strict rustdoc still rejects four unchanged private links. Documentation
builds with only that existing lint allowed. These are scoped results, not
a whole-workspace quality claim.

## Evidence custody and reproduction

[Capture readiness](results/change-0922/capture-readiness.json) binds the
review, qualification, controls, source verification and build parity before
any comparative child. The final manifest is
`f502b69f8b6514ec694fc3c1f296135193510eb486274a2be1cbd8fd3f9b5ff1`.
All 96 qualification reports and 48 exact-capacity pairs passed offline replay.
A single earlier successful qualification child and an envelope-label setup
error are retained externally and excluded from the final qualification
matrix. They are not comparative captures. The tooling freeze dry-run also
stayed entirely outside the repository.

The build receipts record pinned Rust/Cargo 1.95.0, the unchanged 0917 probe
lock and exact release overrides. Root compared 157 compiler/profile/feature
fingerprints and all 144 dependency packages against the actual baseline.
Sorting metadata's target and feature lists resolves ordering differences;
no package/version/feature changes are admitted. A post-capture
[ELF section check](results/change-0922/binary-sections.json) shows the
candidate executable is 7,008,344 bytes smaller, chiefly from 6,956,750 fewer
DWARF bytes; `.text` grows by 1,280 bytes. File size supplies no latency
explanation. The quality setup's synthetic
local Git commit is not a measured source identity: custody uses the canonical
`bb19` base plus the exact three-file patch and source manifest.

Raw comparison artifacts and executables stay under
`/tmp/litchi-performance-0922-comparison` and
`/tmp/litchi-performance-0922/candidate-target`, outside Git. Build commands,
binary sizes and SHA-256 values are in the copied build receipts. The packet
contains the unchanged runner/replay inputs and compact child reports.

[Independent audit](results/change-0922/independent-audit.md) verified all
288 receipts and validator reports, 864 packet artifacts and 864 external
artifacts, including decompression round trips and copied-packet replay.

Root's [two replay checks](results/change-0922/root-replay.json) reproduce
`adoption: false`. The 288 canonical reports have SHA-256
`956b45d732360db78ff666d170d84a299bb11ec92ac1fdd3f1f56a5d4cf7128e`.
Replay the retained packet without rerunning benchmarks:

```sh
python3 -B docs/performance/results/change-0922/comparison.py replay
python3 -B docs/performance/results/change-0922/comparison-replay.py --packet docs/performance/results/change-0922
```

The copied-packet replay uses retained compact reports and fixture hashes;
external original fixture paths are validation aliases, not new measurements.

## Cleanup

[Cleanup](results/change-0922/cleanup.json) removed **16,904,507,237 logical
bytes** of unused archived evidence and completed build/quality intermediates.
The measured source closures, final executables, release fingerprints,
dependency files, logs and all raw comparison attempts remain available.
No executable or evidence file at or above 1 MiB is committed.
