# 0917 — real DOC repeated-edit retention scaling

This work measures the existing DOC validated-render handoff over longer edit
sequences. It makes no production-library change. The source is pinned to
`bb19ca74bd52044fbd2c1c7cc9875e975bb334c0`; the real `NoHeadFoot.doc` and
`FloatingPictures.doc` fixture bytes are bound by the frozen protocol.

## Current-state correction

The 0906 legacy coverage audit nominated a DOC handoff pilot from 0729's next
step. That recommendation overlooked the already-adopted
[0730 handoff](0730-doc-bounded-render-handoff.md). Current source retains the
bounded, observable, explicitly releasable render. The historical pilot measured
one and two edits, including a two-edit peak-live-memory increase. Its binaries
were removed and its source differs from current code, so its timings are not
current controls. The [source review](results/change-0917/design-audit.md) and
[evidence audit](results/change-0917/evidence-audit.md) establish the current
path and the unresolved longer-sequence question.

## Workload and ownership

Each fresh public transaction performs 1, 2, 8, or 32 successive replacements
of paragraph zero, alternating the existing 45- and 46-UTF-16-unit probe texts.
Every stage must change its current text. The final text follows edit-count
parity; the other paragraphs must remain unchanged. This deliberately measures
repeated edits to one object, not a distinct-paragraph bulk edit. The smaller
fixture cannot provide 32 distinct safe ordinary text targets.

All routes have identical operation and replacement limits. They vary only the
render retention policy: zero capacity, the current default 8 MiB ceiling, or
the same ceiling with explicit release after every successful stage. A retained
render can save the final duplicate container render. Every staged mutation
still performs its required validation, and commit still performs strict-owner
and independent public-reader checks. A clone-first mutation can temporarily
overlap the prior retained render with a new candidate; the ceiling is not a
bound on process memory or transient peak live bytes.

## Protocol

The [premeasurement protocol](results/change-0917/native-protocol.json) fixes
six balanced route-order rounds for every fixture/count combination: 144 native
children, three warmups and 30 samples per child. Two separate allocator rounds
contain 48 children with one measured lifecycle each. Source, probe, build,
binary, fixture and runner identities were frozen before measured children.

Native time covers the complete in-memory public lifecycle: input clone/open,
edit construction, all stages and optional releases, commit, output copy and
local owner destruction. File I/O, preparation, correctness checks and JSON
serialization are outside the timer. The allocator binary uses the existing
System wrapper; its region ends with only the output allocation retained.
Stage observations use fixed stack storage. Requested/live allocation bytes
exclude stack and allocator internals; instrumented latency is not reported as
native performance.

Results must retain all children and report each fixture/count/route separately.
The bootstrap resamples matched process rounds, not individual samples. Changes
above 5% in latency, allocator peak live or whole-child RSS are visible review
flags, not sample-exclusion rules or a production-adoption gate. Whole-child
RSS includes untimed oracles and setup and does not isolate the selected route.
Other shared-host work remains uncontrolled even when this batch's CPU work is
idle.

## Qualification and limits

The probe qualifies cross-route exact output equality, public semantic
readback, untouched stream and directory preservation, exact no-ops, reversible
patches, source conflicts and retention transitions. Existing staging,
authorization, policy, protection and final validation remain mandatory.

This is a two-fixture warm-memory study of existing policy choices. It does not
establish new producer support, native Office round trips, cold or range-source
behavior, distinct-object bulk editing, concurrency or a general DOC speedup.
The broader non-iWork performance goal remains open.

## Results

All 144 native children (4,320 measured samples and 432 warmups) and all 48
allocator children passed. No child or sample was excluded or rerun. The
[reduction](results/change-0917/reduced.json) retains every cell, round, tail,
confidence interval, allocation repetition and review flag. This compares
existing policies within one build; no production change is adopted.

Native pooled p50 is in milliseconds. Each ratio is the median of six matched
round ratios, zero-policy p50 divided by route p50; values above one favor that
route. Brackets are the frozen process-round bootstrap's 95% interval.

| Fixture | Edits | Zero p50 | Default p50 | Release p50 | Default ratio [95% CI] | Release ratio [95% CI] |
| --- | ---: | ---: | ---: | ---: | --- | --- |
| NoHeadFoot | 1 | 0.132621 | 0.117851 | 0.133130 | 1.1256 [1.1197, 1.1362] | 0.9951 [0.9880, 1.0004] |
| NoHeadFoot | 2 | 0.160461 | 0.145680 | 0.160611 | 1.1007 [1.0919, 1.1117] | 1.0025 [0.9880, 1.0127] |
| NoHeadFoot | 8 | 0.316861 | 0.300211 | 0.316471 | 1.0511 [1.0450, 1.0621] | 0.9940 [0.9761, 1.0167] |
| NoHeadFoot | 32 | 1.029085 | 1.017054 | 1.030925 | 1.0123 [1.0079, 1.0146] | 0.9977 [0.9933, 1.0053] |
| FloatingPictures | 1 | 1.091235 | 0.988355 | 1.097825 | 1.1032 [0.9955, 1.2109] | 0.9778 [0.8784, 1.1291] |
| FloatingPictures | 2 | 1.292796 | 1.341897 | 1.287455 | 0.9981 [0.8933, 1.0812] | 0.9979 [0.9151, 1.0617] |
| FloatingPictures | 8 | 2.703452 | 2.561052 | 2.636432 | 1.0695 [1.0088, 1.1945] | 1.0053 [0.9873, 1.0855] |
| FloatingPictures | 32 | 8.449399 | 8.131827 | 8.328318 | 1.0503 [1.0052, 1.0540] | 1.0032 [0.9674, 1.0503] |

NoHeadFoot's default-policy advantage narrows from a paired ratio of 1.1256
at one edit to 1.0123 at 32 edits. FloatingPictures is noisier: its one- and
two-edit intervals include one, and its two-edit default pooled p50 is 3.80%
higher than zero. The 8- and 32-edit default intervals are above one, but these
two-fixture results do not establish a general DOC speedup. Every explicit-
release interval includes one.

The flagged native tails are all FloatingPictures at two edits: default
p95/p99 rise 5.86%/6.48%, and release p95/p99 rise 8.81%/13.23% versus zero.
These adverse results remain in the packet. No pooled native RSS comparison
crosses 5%; per-round RSS observations remain visible. Allocator repetition 2
flags default FloatingPictures/two-edit whole-child RSS at 14,228 versus
13,236 KiB (+7.49%). That process includes untimed preparation and all-route
oracles, so this is not a selected-route peak-live measurement.

### Allocation and scaling

Both allocator repetitions are identical for all 24 cells. Peak live and
end-retained output capacity are exactly equal across all three routes at
each fixture/count. Zero and explicit release also have identical requested
bytes and allocation-call counts. Default retention reduces requested bytes
and calls, while the whole-lifecycle peak is unchanged in this build.

| Fixture | Edits | Peak live, all routes (bytes) | Output retained, all routes (bytes) | Zero/release requested (bytes) | Default requested (bytes) | Zero/release calls | Default calls |
| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| NoHeadFoot | 1 | 303,110 | 27,648 | 1,128,699 | 1,008,777 | 2,056 | 1,768 |
| NoHeadFoot | 2 | 312,324 | 28,672 | 1,403,356 | 1,282,207 | 2,558 | 2,269 |
| NoHeadFoot | 8 | 370,820 | 35,840 | 3,310,243 | 3,181,346 | 5,666 | 5,368 |
| NoHeadFoot | 32 | 642,194 | 63,488 | 17,017,321 | 16,781,690 | 19,353 | 19,024 |
| FloatingPictures | 1 | 4,284,552 | 342,528 | 15,722,762 | 14,304,400 | 16,666 | 15,498 |
| FloatingPictures | 2 | 4,370,204 | 353,280 | 18,799,147 | 17,369,691 | 19,183 | 18,014 |
| FloatingPictures | 8 | 4,891,053 | 418,304 | 39,236,818 | 37,737,565 | 34,365 | 33,187 |
| FloatingPictures | 32 | 7,635,361 | 686,592 | 167,215,584 | 164,233,879 | 96,264 | 95,057 |

The single retained render does not make the entire transaction constant-space.
From one to 32 edits, observed peak live grows from 303,110 to 642,194 bytes
for NoHeadFoot and from 4,284,552 to 7,635,361 bytes for FloatingPictures.
The materialized output itself also grows, even though the final selected text
is determined by parity. The next useful diagnosis is the many-edit lifecycle's
allocation and staging costs, preserving its required validation; these totals
do not attribute a particular internal function or authorize removing checks.

## Qualification, custody and replay

The [design review](results/change-0917/design-audit.md) approved the timed
ownership boundary and DOC-only extraction of the 0730 preservation oracle.
The final exported build passed format, three unit tests, warning-denied
Clippy and warning-denied rustdoc; [quality receipts](results/change-0917/quality-receipt.json)
cover the standalone harness, not a fresh whole-workspace certification.
All 48 premeasurement combinations passed with zero warmups and one sample,
as recorded in the [qualification receipt](results/change-0917/qualification-receipt.json).
These development/qualification runs are excluded from performance statistics.

The [source receipt](results/change-0917/receipts/source-receipt.json) binds the
164,352,000-byte external source archive, the eight exact probe files and the
1,108,086-byte source manifests by SHA-256. All 9,268 exported files remained
unchanged through build, quality and [capture](results/change-0917/post-capture-source.json).
The [build receipt](results/change-0917/receipts/build-receipt.json) records the
commands, environment, logs and both external executable hashes. Large evidence
and compiled executables remain outside Git. Preliminary unlogged development
smokes made no measurement claim; their three temporary JSON files were removed.

The initial postcapture reduction failed on one tuple-unpacking error. Both
failed commands are preserved in [reduction-failure.json](results/change-0917/reduction-failure.json).
The frozen runner, replay script, protocol, manifest and raw children remain
unchanged. [analysis.py](results/change-0917/analysis.py) verifies the frozen
runner hash and applies exactly the documented one-line correction in memory;
[reduction-repair.md](results/change-0917/reduction-repair.md) explains the repair.
Reduction and replay then passed. No measurements were replaced. The
[independent final review](results/change-0917/review.md) recomputed every
statistic and allocation comparison and rejected copied-packet corruptions.
[Cleanup](results/change-0917/cleanup.json) removed 6,497,677,312 bytes of
temporary build/export data while retaining both final executables and evidence.

Replay the retained packet with:

```sh
LITCHI_PERF_0917_ROOT=/home/zhuhe/litchi-performance-0917/capture \
  python3 -B docs/performance/results/change-0917/analysis.py replay
```

To rebuild, regenerate and extract the pinned archive using `base-source.json`,
copy the eight probe files including the ignored-but-committed `Cargo.lock` into
the same relative export path, then run the exact locked/offline release command
in the build receipt. Each raw child receipt records its command and environment.
New measurement runs require a fresh packet and empty external capture directory;
the runner intentionally refuses to overwrite existing evidence.

The exact manifest regeneration below runs from the restored export root and
writes outside it. The four source manifests have the same SHA-256
`6be84856c37c4729c79968c3616c7a07aa040a136971dba26301d6589e1b3c49`.

```sh
python3 -B - <<'PYMANIFEST' > ../source-tree-regenerated.sha256
from pathlib import Path
import hashlib
root = Path('.')
files = sorted((p for p in root.rglob('*') if p.is_file()),
               key=lambda p: p.relative_to(root).as_posix())
for path in files:
    print(hashlib.sha256(path.read_bytes()).hexdigest(),
          path.relative_to(root).as_posix(), sep='  ')
PYMANIFEST
```

This batch adds two-fixture repeated-object evidence to the existing DOC edit
scenario. It adds no selector, producer, concurrency, cold/range-source or native
Office certification. Shared-host activity was uncontrolled; the recorded
pre-capture check establishes only that this batch's coordinated CPU work was
idle. The non-iWork program goal remains open.
