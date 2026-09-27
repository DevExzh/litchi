# 0781 — fresh PPT text ownership and the repaired CI matrix

Status: **the borrowed `Cow<str>` candidate is rejected for production; the
PPT production source is restored byte-for-byte; the CI matrix repair is
retained.** This packet is descriptive performance and custody evidence. It
does not establish a retained production speedup, promote a CRUD row, or
complete the broader format goal.

The candidate addressed a current fresh PPT writer conversion boundary. Plain
text was copied into private `UserShapeData` immediately before OfficeArt
serialization. The candidate changed that private field to
`Option<Cow<'text, str>>`, borrowed ordinary plain text during synchronous
serialization, and kept rich paragraphs and mutation-sensitive fallbacks
owned. Lifetimes were carried through the crate-private group and codec types;
the public writer API, existing-document policy, and output encoding route
were unchanged. The candidate covered twelve source paths, including focused
tests, in [`candidate/changed-files.json`](results/change-0781/candidate/changed-files.json).

The active priority remains OLE2 and OOXML. ODF optimization remains deferred
under the 0758 owner decision, and iWork is excluded from this batch.

## Protocol and qualification

The frozen plan in [`plan.json`](results/change-0781/plan.json) measured ten
cases: five shapes (`tiny`, `many`, `payload`, `unicode`, and `rich`) in both
public `write` and full `lifecycle` modes. The native lane used six alternating
blocks, 30 measured samples after three warmups, CPU 12, and one process per
case and leg in each block: 120 processes and 3,600 timed samples. The
separate allocation lane used two blocks, three samples, no warmup, and 40
processes. A ten-process, one-sample before-only qualification ran outside
those paired measurements.

Fixture generation, reopening, source/output digest checks, and both text
oracles were outside the measured region. Write mode times only the public
`Writer::write_to` call after authoring; lifecycle also includes writer
construction and authoring from the prepared fixture. The writer and published
output stayed alive through the measured endpoint; their destruction and the
Cursor-to-Vec ownership conversion happen afterward. The
comparison is warm in-memory fresh writing; it makes no cold-cache, device,
concurrency, or physical-copy claim.

The first probe qualification stopped at Unicode write. Its fixture contains
one authored trailing space in each of 64 text boxes, while the public reader
trims each decoded text atom. The resulting 64-byte semantic difference was
an oracle mismatch, not a production or fixture change. The failed run, six
successful initial reports, initial binaries, and relocation receipts remain
under [`qualification-initial/`](results/change-0781/qualification-initial/),
[`initial-probe/`](results/change-0781/initial-probe/), and
[`qualification-correction.json`](results/change-0781/qualification-correction.json).

The corrected probe keeps the authored fixture digest and checks two separate
properties. Its public-reader oracle applies the documented per-atom trim
rule. Its strict raw oracle traverses bounded OfficeArt `ClientTextbox`
records, decodes UTF-16 and ASCII text atoms with record and payload bounds,
and compares the untrimmed authored strings, including trailing spaces and
rich paragraph separators. It is a direct OfficeArt text-atom check, not an
independent CFB parser. All ten corrected baseline cases pass. The six cases
that had completed before the correction retain identical source and output
identities in [`qualification-parity.json`](results/change-0781/qualification-parity.json).
The final raw oracle covers all 170 primary reports and 3,730 samples.

The first all-features probe test attempt also remains visible under
[`probe-tests-initial/`](results/change-0781/probe-tests-initial/). Its
synthetic global-counter tests assumed only explicit counter calls, but the
installed global allocator counted barriers, threads, and test-harness
allocations. The first peak assertion failed and poisoned the test mutex,
leaving 29 passes and seven failures. The repaired driver runs the 31
default-feature counter and oracle tests, then ten all-feature allocator and
oracle tests with the 26 synthetic counter tests filtered from that second
run. Both one-thread runs pass; no production or frozen fixture source was
changed.

## Native timing and RSS

The table shows the median of six process p50 values in microseconds. The
paired percentage is the median of six after/before block ratios. Its interval
is the percentile bootstrap 95% interval over those six ratios (10,000
resamples, seed 781078). RSS is the whole-process peak from `/usr/bin/time`;
the RSS percentage uses the same paired-block method and includes setup,
reopening, and verification.

| Case | Before p50 (µs) | Candidate p50 (µs) | Paired p50 change [95% CI] | Median peak RSS before → candidate (KiB) | Paired RSS change [95% CI] |
| --- | ---: | ---: | ---: | ---: | ---: |
| `many/lifecycle` | 317.917 | 351.892 | +10.551% [+9.756%, +12.641%] | 6,492 → 6,646 | +1.847% [+0.702%, +6.718%] |
| `many/write` | 273.926 | 307.917 | +12.121% [+11.575%, +12.931%] | 6,338 → 6,658 | +5.085% [+1.747%, +6.264%] |
| `payload/lifecycle` | 441.282 | 359.182 | -19.186% [-20.779%, -16.082%] | 38,916 → 36,812 | -5.325% [-5.553%, -5.147%] |
| `payload/write` | 371.522 | 268.026 | -25.555% [-34.269%, -23.203%] | 38,974 → 36,962 | -5.216% [-5.634%, -4.904%] |
| `rich/lifecycle` | 16.810 | 16.715 | -0.476% [-1.800%, +0.567%] | 3,486 → 3,628 | +2.411% [-2.939%, +4.373%] |
| `rich/write` | 16.425 | 16.495 | +0.333% [-0.809%, +1.037%] | 3,534 → 3,634 | +2.462% [-1.900%, +6.684%] |
| `tiny/lifecycle` | 13.050 | 13.165 | +0.881% [-0.038%, +2.089%] | 3,674 → 3,628 | -1.250% [-4.832%, +2.214%] |
| `tiny/write` | 12.725 | 12.870 | +0.748% [+0.545%, +1.815%] | 3,498 → 3,612 | +1.544% [-3.420%, +6.891%] |
| `unicode/lifecycle` | 2,419.841 | 2,228.785 | -8.073% [-8.195%, -7.549%] | 53,070 → 48,496 | -8.638% [-8.675%, -8.570%] |
| `unicode/write` | 2,279.011 | 2,090.204 | -8.314% [-8.634%, -7.958%] | 53,094 → 48,470 | -8.723% [-8.961%, -8.620%] |

The analyzer retains 20 positive paired metric-series flags above 5% and 42
within-group process-spread flags above 5%. Allocation has zero paired or
spread flags. A paired series is flagged when any block exceeds +5%, even
when its median does not. Every process and sample remains in the packet; no result-driven
rerun or deletion was performed. The full material is in
[`native-processes.csv`](results/change-0781/native-processes.csv),
[`native-pairs.csv`](results/change-0781/native-pairs.csv), and
[`all-flags.csv`](results/change-0781/all-flags.csv).

## Allocation evidence

The values below are repeated p50s from both allocation blocks. Allocated
bytes and allocation calls fall on the large plain and Unicode fixtures, but
`peak above entry` and `net live` are unchanged. These are allocator-region
counters; they are not RSS, physical bytes copied, or additive phase costs.

| Case | Allocated bytes before → candidate | Allocation calls before → candidate | Peak above entry before → candidate | Net live before → candidate |
| --- | ---: | ---: | ---: | ---: |
| `many/lifecycle` | 4,227,653 → 4,187,653 | 18,157 → 17,157 | 1,862,395 → 1,862,395 | 1,416,261 → 1,416,261 |
| `many/write` | 2,549,024 → 2,509,024 | 16,849 → 15,849 | 893,110 → 893,110 | 446,976 → 446,976 |
| `payload/lifecycle` | 15,774,881 → 13,214,881 | 1,947 → 1,883 | 10,426,495 → 10,426,495 | 7,761,029 → 7,761,029 |
| `payload/write` | 13,166,780 → 10,606,780 | 1,862 → 1,798 | 7,823,866 → 7,823,866 | 5,158,400 → 5,158,400 |
| `rich/lifecycle` | 92,175 → 92,175 | 624 → 624 | 41,133 → 41,133 | 21,221 → 21,221 |
| `rich/write` | 83,754 → 83,754 | 599 → 599 | 32,712 → 32,712 | 12,800 → 12,800 |
| `tiny/lifecycle` | 83,627 → 83,423 | 502 → 496 | 39,609 → 39,609 | 19,697 → 19,697 |
| `tiny/write` | 76,730 → 76,526 | 491 → 485 | 32,712 → 32,712 | 12,800 → 12,800 |
| `unicode/lifecycle` | 30,843,265 → 27,873,665 | 1,988 → 1,924 | 19,426,527 → 19,426,527 | 13,290,629 → 13,290,629 |
| `unicode/write` | 27,825,564 → 24,855,964 | 1,903 → 1,839 | 16,414,298 → 16,414,298 | 10,278,400 → 10,278,400 |

All 13 allocator fields, including deallocation and region-peak fields, are
retained in [`allocation-summary.csv`](results/change-0781/allocation-summary.csv)
and [`analysis.json`](results/change-0781/analysis.json). The independent
Heaptrack lane finds 192 conversion allocation events and 7,680,000 requested
bytes under `convert_shape_to_escher_with_sound_mapping` before the candidate,
and zero such events after it. Whole-process totals change from 15,665 to
15,473 allocation events and from 362,287,790 to 354,607,785 requested bytes.
The trace, histogram, and `heaptrack_print -f trace.zst -H histogram -n 15 -s 3`
summary cross-check. This is allocation ancestry evidence only: it carries no
timing, RSS, physical-copy, or causal-cost claim.

The separate `perf stat` observer has 12 whole-process receipts across three
blocks for payload write and lifecycle. Instructions move by a median −0.066%
in each mode; cycles move by a median −1.728% for write and −1.592% for
lifecycle. Setup, publication, reopening, and semantic verification are
included, so these counters cannot identify the source of the many-shape
latency regression.

## Disposition

The candidate passes static review, output and semantic checks, existing
refusal tests, and all six candidate quality gates, but it is rejected for
general adoption. The
`many` fixture is 100 slides × 10 short-text boxes. Its p50 regresses in all
six paired blocks by +11.49% to +13.15% for write and +9.33% to +14.04% for
lifecycle. The payload fixture improves by 25.555% and 19.186%, and Unicode
improves by 8.314% and 8.073%, but those targeted gains do not justify the
consistent short-text end-to-end regression. The unchanged live and peak
allocation gauges provide no retained-memory benefit that would offset it.

No cause is inferred for the latency regression. A narrower `Option<&str>`
representation is recorded as a future hypothesis in
[`candidate/next-candidate-notes.md`](results/change-0781/candidate/next-candidate-notes.md);
it has no implementation or result in this batch. The candidate's applied
archive, patch, reviews, measurements, and restored-source receipt remain
available for that separately gated experiment. The live PPT source matches
the before-build manifest recorded by
[`restored-source.json`](results/change-0781/restored-source.json), and
[`disposition.json`](results/change-0781/disposition.json) records
`production_change_retained: false`.

## CI matrix repair

The independent CI audit found that the workflow still asserted 37 default
cases and 201 rows while the current `Case::DEFAULT` has 41 cases and the
authoritative manifest has 213 full rows. The workflow also used the historical
CRUD coverage index v1. The retained repair makes default-matrix validation
manifest-backed, selects the active v2 index throughout the workflow, and adds
executable policy and matrix tests. It is independent of the rejected PPT
candidate and remains retained.

The captured baseline harness produced 41 rows in smoke mode (two samples,
tiny/compressible selectors) and 213 rows in full mode (15 samples, 43 corpus
entries and 213 bindings). The six offline CI gates pass with 75 Python unit
tests:

1. the validator, workflow-policy, and CRUD-index unit suite;
2. smoke report matrix validation;
3. full report matrix validation;
4. active v2 CRUD index contract validation;
5. full-report corpus binding;
6. active v2 coverage validation against the full report.

The retained CI review notes two bounded follow-ups. The manifest and existing
validators still contain a provenance literal naming `main.rs:Case::DEFAULT`,
while the declaration is currently in `lib.rs`; correcting that shared stale
locator belongs to a later provenance repair. The workflow policy test is
static, while the helper validates the retained real reports. Neither caveat
blocks the repaired matrix contract or promotes CRUD completeness.

## Quality, scope, and reproduction

The applied candidate's six PPT gates pass: format check, all-feature and
all-target check, all-feature tests, warning-denied Clippy, warning-denied
rustdoc, and the crate-boundary check. The bound test log reports 1,286 passed,
zero failed, 11 ignored, across 34 suites. The retained probe tests report 31
default-feature passes and ten all-feature passes. These numbers describe the
candidate validation archive; the production source was then restored and
custodied byte-for-byte.

The capture-time roots were:

```text
base commit:  fdca3e63037ac83ff27d4bf64a0f157b971b45f8
worktree:     /home/zhuhe/code/litchi-worktrees/0781-ppt-borrowed-text
packet:       docs/performance/results/change-0781
target:       /home/zhuhe/code/litchi-target-0781
```

Replay the retained packet from the repository root with:

```bash
python3 -B docs/performance/results/change-0781/validate.py
python3 -B docs/performance/results/change-0781/tables.py --check
```

A new measurement requires a checkout of the recorded base, a fresh output
packet initialized with the retained scripts, plans, probe and candidate
archives, and a new owned target. Do not overwrite the retained capture
folders. Restore the workspace/probe lockfiles, update the owned path in
`custody.py`, and render the probe manifest from its template. Run these steps
serially with script paths relative to the fresh packet:

1. Build `before` with `build.py before`; run `probe_tests.py` and
   `capture.py qualification`.
2. Run `observers.py heaptrack-before` and `heap_decode.py before` while the
   baseline source is still live.
3. Apply `candidate/applied-model.patch`, run `quality.py`, and build `after`
   with `build.py after`.
4. Run `capture.py native`, `capture.py allocation`,
   `observers.py heaptrack-after`, `heap_decode.py after`, and
   `observers.py perf` in that order and with the frozen plans.
5. Independently rebuild the unchanged baseline CI harness using the exact
   command, lock, environment and source records in `ci-baseline-build/`;
   capture the two modes with `ci_capture.py` and run `ci_quality.py`.
6. Produce retained observer and CI analyses through their pure `analyze`
   APIs, record the quality test summary and decision/source custody, then run
   `analyze.py` and `tables.py`. Seal the complete packet and run `validate.py`.

The reproduction has its own identities and receipts; it must not reuse this
packet's timestamps or binary custody claims. The archived initial oracle and
test-setup failures explain this run's history, rather than mandatory failures
for a new corrected run. The full current packet, including that history, is
independently replayable using the commands above.

The packet-specific validator also binds the 0781 correction history. A new
packet must adapt that provenance validation to its own qualification history;
the historical receipts must not be relabeled as evidence from a new run.

## Integration and cleanup

All seven executable identities were checked before removing the owned target
(3,248,857,328 file bytes). Offline replay and all five table checks pass with
cleanup witnesses after removal. The final seal covers 766 payload files;
all 766 files plus the seal were checked against staged Git blobs before
commit `6b817fc663`, which was fast-forwarded into
`feat/office-format-completeness`. Sealed replay and all five table checks
pass from the main checkout both before and after removal of the original
worktree and executables.

The owned worktree `/home/zhuhe/code/litchi-worktrees/0781-ppt-borrowed-text`
and branch `perf/0781-ppt-borrowed-text` were removed after verifying and
removing the copied workspace lockfile and three exact reference symlinks.
All pre-existing worktree records and the three unrelated main-worktree file
hashes remain unchanged. No production crate diff is retained. The broader
non-iWork goal remains active.
