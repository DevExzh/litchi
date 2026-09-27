# 0782 — optional borrowed PPT text slice rejected

Status: rejected. The private `Option<&str>` candidate improves large ASCII payload writing but violates the frozen latency guard in four public cases. Production source is restored byte-for-byte to `f1df64d15a8f2711e554f58acb413b06d56ef1ca`. No production speedup or CRUD coverage promotion is retained.

The experiment follows 0781’s rejected `Cow<str>` representation. Current construction sites only need a borrowed plain string during synchronous serialization, so the narrower field removed the unused owned alternative. Rich paragraphs, smart-tag mutation and centered fallback retained their existing owned paths. Both public serialization routes and private grouped-shape lifetimes were covered. The public writer still owns its strings; no public API, dependency, cache, unsafe code or existing-document policy changed. The applied twelve-file candidate remains archived.

The narrower field passes its current-target layout test, but that does not measure the whole shape layout or establish a cause for either experiment’s latency. This run compares its own freshly built baseline and candidate; historical timing values are neither pooled nor used as a third leg. The corrected 0781 probe source, template and lock are reused byte-for-byte, including the old tool identifier and authored fixture identities. Replay binds those files and all ten qualification source/output identities to the actual sealed historical artifacts.

## Protocol and native results

The ten cases combine tiny (2 slides × 3 short boxes), many (100 × 10 short boxes), payload (16 × 4 boxes of 40,000 ASCII bytes), Unicode (16 × 4 boxes of 40,000 UTF-16 units), and rich (2 × 2 formatted boxes) with write and lifecycle modes. Write times `Writer::write_to` after authoring; lifecycle also creates and populates the writer from a prepared fixture. Fixture generation, reopen, source/output hashing and semantic/raw checks occur outside the timer. Writer and output remain alive through timing/allocation endpoints; their destruction is outside.

CPU 12, six alternating native blocks, 30 samples after three warmups, two separate allocation blocks with three samples and no warmup, and ten before-only qualification processes were fixed before capture. There are 120 native processes (3,600 samples), 40 allocator processes (120 samples), and ten qualification samples. No result-driven retries or resampling occurred.

The table uses the median of six nearest-rank process p50s in microseconds. Paired changes use the median after/before ratio across the six blocks; intervals use 10,000 percentile-bootstrap resamples, seed 782078, 95% confidence. Peak RSS includes whole-process setup and readback.

| Case | Before p50 (µs) | Candidate p50 (µs) | Paired p50 change | 95% interval | Paired peak RSS change |
| --- | ---: | ---: | ---: | ---: | ---: |
| `many/lifecycle` | 316.851 | 352.407 | +11.089% | [+10.850%, +12.332%] | +3.036% |
| `many/write` | 274.151 | 310.047 | +12.897% | [+11.016%, +13.592%] | +3.465% |
| `payload/lifecycle` | 438.552 | 352.467 | -21.642% | [-22.810%, -18.167%] | -0.046% |
| `payload/write` | 354.667 | 272.877 | -24.328% | [-26.402%, -20.472%] | -0.026% |
| `rich/lifecycle` | 16.856 | 16.700 | -0.865% | [-1.569%, -0.033%] | +1.124% |
| `rich/write` | 16.480 | 16.465 | +0.121% | [-0.993%, +1.977%] | -1.231% |
| `tiny/lifecycle` | 13.050 | 12.535 | -3.677% | [-5.425%, -2.656%] | +0.172% |
| `tiny/write` | 12.775 | 12.320 | -3.562% | [-4.256%, -1.294%] | +3.822% |
| `unicode/lifecycle` | 1,873.703 | 2,232.775 | +19.069% | [+18.976%, +19.254%] | -8.850% |
| `unicode/write` | 1,734.943 | 2,091.539 | +20.650% | [+20.221%, +20.948%] | -8.913% |

The frozen [adoption policy](results/change-0782/adoption-policy.json) counts a median p50 increase above 5% with its 95% interval wholly above zero against adoption in any case. Many and Unicode write/lifecycle all violate that guard. Tiny improves modestly and rich is approximately neutral. Large ASCII payload improvements do not offset the four consistent regressions. No causal attribution is made.

All 42 native process-spread flags and 24 paired metric-series flags remain in [all-flags.csv](results/change-0782/all-flags.csv). A paired series is flagged if any block exceeds +5%, even when its median does not. Tails, individual blocks, RSS distributions and confidence intervals remain in the complete analysis and five generated tables. There are no allocation spread or regression flags.

## Allocation and profiler observations

| Case | Allocation calls before → candidate | Requested bytes before → candidate |
| --- | ---: | ---: |
| `many/lifecycle` | 18,157 → 17,157 | 4,227,653 → 4,179,653 |
| `many/write` | 16,849 → 15,849 | 2,549,024 → 2,501,024 |
| `payload/lifecycle` | 1,947 → 1,883 | 15,774,881 → 13,214,369 |
| `payload/write` | 1,862 → 1,798 | 13,166,780 → 10,606,268 |
| `rich/lifecycle` | 624 → 624 | 92,175 → 92,143 |
| `rich/write` | 599 → 599 | 83,754 → 83,722 |
| `tiny/lifecycle` | 502 → 496 | 83,627 → 83,375 |
| `tiny/write` | 491 → 485 | 76,730 → 76,478 |
| `unicode/lifecycle` | 1,988 → 1,924 | 30,843,265 → 27,873,153 |
| `unicode/write` | 1,903 → 1,839 | 27,825,564 → 24,855,452 |

Values are stable repeated p50s in both allocation blocks. Net live bytes retained at region exit and peak live bytes above region entry are unchanged in every case. The eight-byte-per-shape reduction in intermediate request volume is consistent with the narrower field, but it is not a measured whole-struct layout, RSS, physical-copy or latency attribution. Absolute live gauges also include a one-byte executable-name difference; those are not production savings.

Separate Heaptrack payload/write processes contain three operations per leg. Complete ancestry under `convert_shape_to_escher_with_sound_mapping` falls from 192 allocation events / 7,680,000 requested bytes to zero. Whole-process totals are 15,665 → 15,473 events and 362,287,791 → 354,606,250 requested bytes. Exact trace counts/sums cross-check the histogram and print summary. These observations confirm removal of transient allocation requests, not a timed phase share.

Twelve separate `perf stat` processes cover payload write/lifecycle across three blocks. Median paired whole-process instructions change −0.046% in each mode, while cycles increase 13.857% / 14.219%. The process includes fixture setup and expensive post-clock verification. These counters neither contradict the scoped operation timer nor identify the cause of the many/Unicode regressions; they must not be reported as operation-local counters or pooled with native timings.

## Correctness, custody and decision

Every primary report checks deterministic source/output identities, slide counts, reader text with its documented per-atom trimming, and direct untrimmed OfficeArt text atoms including Unicode trailing spaces. The raw oracle uses public record parsers with record/depth/UTF-16/ASCII checks; it is not an independent CFB parser or native Office interoperability result. All 170 primary reports and 3,730 samples pass. Existing text goldens, rich smart-tag source immutability, centered fallback, grouped borrowed text, and both public output paths pass.

All six candidate quality gates pass: formatting, all-feature/all-target check, all-feature PPT tests, warning-denied library Clippy, warning-denied rustdoc, and crate boundaries. The test log records 1,287 passed, zero failed, 11 ignored across 34 suites. The inherited probe tests pass 31 default-feature tests and ten all-feature tests; synthetic counter tests run without installing the real global allocator. No failed build or capture is hidden.

The selected draft, applied source, original source and restored census are bound in the [packet](results/change-0782/README.md). The [static review](results/change-0782/candidate-review.md) found no correctness blocker. The [decision audit](results/change-0782/decision-audit.json) independently recomputes raw process p50 ratios and bootstrap intervals and confirms the four policy violations. An [independent results review](results/change-0782/results-review.md) also recommends rejection. The analysis agent reached a service usage limit after writing its modules; the coordinator inspected, strengthened and replayed the completed validator locally. This is disclosed process history, not evidence of a missing measurement.

The production candidate is rejected under [disposition.json](results/change-0782/disposition.json). Another representation-only trial is not selected. A subsequent performance investigation should first attribute the many-short and Unicode costs in the actual before/after binaries or inspect generated code under a separately frozen diagnostic plan; it must not assume string ownership explains the observations. The 0780 large-lifecycle phase question, comprehensive CRUD metrics, cold/range/concurrency/scaling and broader budget work remain open. OLE2/OOXML remain active, ODF deferred, and iWork excluded. All 35 architecture input hashes remain unchanged.

## Replay, reproduction and integration

From the repository root:

```bash
python3 -B docs/performance/results/change-0782/validate.py
python3 -B docs/performance/results/change-0782/tables.py --check
python3 -B docs/performance/results/change-0782/decision_audit.py
```

New measurements require a fresh checkout of the recorded base, fresh packet/output directories and target, retained workspace/probe locks, identical probe/template/plans, and new source/binary receipts. Render the manifest for the new path. Run `build.py before`, `probe_tests.py`, `capture.py qualification`, `observers.py heaptrack-before`, and `heap_decode.py before` before applying `candidate/applied-model.patch`. Then run `quality.py`, `build.py after`, `capture.py native`, `capture.py allocation`, `observers.py heaptrack-after`, `heap_decode.py after`, and `observers.py perf` serially. Do not overwrite retained artifacts or relabel old receipts. Generate observer/main analyses, test summary and decision/source custody, then tables and the seal. The packet-specific decision audit records this run’s four violations; a new run needs its own reviewed decision, not copied results.

Commit `8307452c01` was fast-forwarded into `feat/office-format-completeness`.
The seal covers 693 payload files; all 694 packet Git blobs, including the seal,
were checked before commit. All four executable identities were verified before
removing the owned build target (1,981,285,708 file bytes). The owned worktree
and `perf/0782-ppt-borrowed-slice` branch are also removed. The preexisting
worktree inventory and hashes of the three unrelated main-checkout files remain
unchanged.

Complete sealed replay, all five tables and the independent decision audit pass
from the main checkout after target and worktree removal. The audit confirms
four latency guard violations. No production candidate changes are retained;
the broader performance goal remains active.
