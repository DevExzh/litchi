# 0821 — real-file save durability attribution

Fresh measurements separate synchronization policy from the ordinary-save
runtime on three small real files. Default and explicit-full controls agree
within their paired 95% intervals. File-only/default p50 ratios range from
0.62841 to 0.75003; no-sync/default ratios range from 0.01676 to 0.28031.
The weaker settings trade away durability guarantees and are not adopted as
optimizations. Production code and the full-durability default are unchanged.

The next compatible investigation is profiling the ordinary PPTX edit/publication
CPU path while retaining full durability. Its no-sync lifecycle remains about
2.078 ms, versus 0.242/0.534 ms for DOCX/XLSX on these distinct inputs. Those
cross-format values describe different workloads; they do not prove a shared
bottleneck or justify a code change without a fresh profile.

## Scope and controls

This batch compares default save, explicit full durability, file-only durability,
and no-sync publication for the admitted DOCX, XLSX, and PPTX real files. Each
format has separately prepared lifecycle and atomic-publication cases. All
four routes retain the same output-content and semantic admission requirements.
Default and explicit full have the same durability contract; file-only omits
parent-directory synchronization, and no-sync also omits temporary-file sync.
These are configuration comparisons, not production optimizations. The default
remains full durability.

The fresh base is the allocator-test repair in `8312aaa29b`. Production and
benchmark runtime are unchanged. The preceding repair's six quality gates are
reused only after checking exact source, lock, normative-input, and committed
receipt identity. Fresh release binaries, semantic and ZIP admission, and all
measured reports belong to this packet. No earlier timing samples are pooled.

## Measurement contract

The host is an AMD EPYC 9R45 running Linux 7.0.0-1012-aws, with 32 visible
logical CPUs and capture affinity 12–19. The source and destination filesystem
is ext4 on `/dev/nvme0n1p1`; the cgroup records no CPU or memory cap. Builds use
Rust 1.95.0, opt-level 3, thin LTO, one codegen unit, debug level 1, two build
jobs, and no additional Rust flags. These host/storage details bound the
interpretation of synchronization latency.

The inputs are `documentProperties.docx` (23,503 bytes), `dateAutofilter.xlsx`
(8,435 bytes), and `shapes.pptx` (68,822 bytes). Exact repository paths and
SHA-256 values are in [the corpus manifest](results/change-0821/corpus-inputs.json).

The frozen matrix has 24 cases: three inputs × two phases × four policies.
Six native process blocks have 30 samples and three warmups per case. The six
format/phase groups follow forward/reverse orders while policy order rotates
and reverses across blocks. Ratios compare policy p50 to default p50 within the
same block and case; six default self-controls and 18 other comparisons are
retained. Native and allocator/procfs observer latency are never pooled.

Nearest-rank p50/p95/p99 are computed within each process; tables report the
median across six blocks. Absolute p50 and matched-block policy/default ratios
use 10,000 bootstrap resamples, seed 821821, with sorted endpoints 250 and 9749.
The harness's raw midpoint p50 remains in the packet alongside the recomputed
nearest-rank statistic. All block spread and tail flags remain descriptive.

Each measured destination starts absent and is removed outside timing. The
filesystem/provider cache is warm; this does not demonstrate physical cold I/O,
overwrite permission handling, or recovery from a crash. The policy-specific
`atomic_publication_steps` and `save_durability` fields identify synchronization;
the generic `timing_scope` prose is not evidence of full synchronization under
weaker policies. Output equality is not proof of crash durability.

## Results

Native p50 in milliseconds; each value is the median of six block p50s.

| Format / phase | Default | Full | File-only | No-sync |
| --- | ---: | ---: | ---: | ---: |
| DOCX / lifecycle | 5.222937 | 5.205017 | 3.377918 | 0.242312 |
| DOCX / path save | 5.024506 | 5.026841 | 3.195921 | 0.083996 |
| XLSX / lifecycle | 5.410838 | 5.415628 | 3.593988 | 0.534127 |
| XLSX / path save | 4.948680 | 4.937820 | 3.104771 | 0.099605 |
| PPTX / lifecycle | 7.420598 | 7.403842 | 5.569784 | 2.078195 |
| PPTX / path save | 5.618328 | 5.608739 | 3.780639 | 0.296617 |

Matched-block policy/default p50 ratios with 95% bootstrap intervals:

| Format / phase | Full / default | File-only / default | No-sync / default |
| --- | ---: | ---: | ---: |
| DOCX / lifecycle | 0.99802 [0.99446, 1.00161] | 0.64702 [0.64229, 0.65166] | 0.04632 [0.04569, 0.04693] |
| DOCX / path save | 1.00087 [0.99801, 1.00368] | 0.63572 [0.63161, 0.63944] | 0.01676 [0.01654, 0.01691] |
| XLSX / lifecycle | 1.00087 [0.99352, 1.00727] | 0.66372 [0.66088, 0.66743] | 0.09902 [0.09805, 0.09961] |
| XLSX / path save | 0.99857 [0.99330, 1.00418] | 0.62841 [0.62400, 0.63089] | 0.02009 [0.01998, 0.02026] |
| PPTX / lifecycle | 0.99839 [0.99348, 1.00070] | 0.75003 [0.74772, 0.75491] | 0.28031 [0.27862, 0.28474] |
| PPTX / path save | 0.99791 [0.99373, 1.00147] | 0.67326 [0.67065, 0.67344] | 0.05274 [0.05221, 0.05354] |

All six explicit-full intervals include 1.0; all six default self-controls are
exactly 1.0. The twelve weaker-policy intervals lie below 1.0. This supports
synchronization policy as a material contributor to the observed path-save
latency on this filesystem, without pricing individual syscalls or establishing
a portable synchronization cost.

Seventeen cases retain 27 block-spread flags across p95, p99, or mean; none has
a p50 block-spread flag. Six cases have p99/p50 above 1.05: DOCX no-sync in both
phases, XLSX file-only lifecycle, XLSX no-sync in both phases, and PPTX no-sync
path save. All samples, p95/p99/mean values, spread ratios, and RSS observations
remain in [the full analysis](results/change-0821/analysis.json).

## Resource observations and limits

The six observer samples per format/phase/policy have the following medians.
All four policies have identical allocation-call, allocated-byte, and region-peak-above-entry medians within each row.

| Format / phase | Allocation calls | Allocated bytes | Region peak above entry, bytes |
| --- | ---: | ---: | ---: |
| DOCX / lifecycle | 2,617 | 1,258,172 | 634,214 |
| DOCX / path save | 380 | 590,203 | 568,921 |
| XLSX / lifecycle | 3,542 | 4,561,313 | 1,103,104 |
| XLSX / path save | 126 | 983,559 | 561,522 |
| PPTX / lifecycle | 11,402 | 5,814,055 | 939,897 |
| PPTX / path save | 674 | 641,339 | 597,823 |

Native whole-process peak-RSS medians range from 135,980 to 136,120 KiB.
These include process startup and verification; they are not operation-retained
heap measurements. Region peak above entry is computed per observer sample as
`region_peak_live_bytes - live_bytes_before`, then summarized across six samples.
Allocation/procfs values include observer overhead and remain separate from native latency.

Phase preparation differs, so phase medians are not additive and are not
subtracted to claim synchronization cost. Observer timings are diagnostic only;
empty procfs controls are retained without subtraction. Three small files on
one host do not establish broad producer, corpus-size, storage-device, or
platform performance. Weaker-policy ratios cannot authorize weakening defaults.

## Verification and replay

Three fresh release builds and all capture children completed successfully.
The committed 0820 repair provides six passing quality gates: 641 tests passed,
zero failed, one ignored, plus the five-test focused allocator suite. The fresh
adapter verifies source-byte equality, both locks, all 35 normative inputs,
result/log hashes, and the prior committed seal before reuse; these are reused
quality results, not a new Cargo test invocation.

Fresh artifact admission covers six corpora and thirty outputs. Semantic/XML
checks, decoded/compressed untouched-member bytes, ZIP metadata, member order,
and archive comments pass. All 24 qualification selectors are bound to admitted
source and policy-output hashes before native or observer capture.

Root initially invoked admission before producing `zip-preservation.json`.
The artifact auditor passed, but the parent invocation failed on that missing
prerequisite. The failed invocation and `admission-0` audit remain retained;
preservation then passed and `admission-1` accepted the same exports. No driver,
frozen input, output, or timing sample was replaced to resolve the error.

Offline analysis and replay accept all 216 reports and 4,488 samples. The root's
independent raw audit reproduces native quantiles, spreads, tails, and all 24
absolute and paired bootstrap intervals. Independent result review is retained
in [results-review.md](results/change-0821/results-review.md).

The first analysis replay rejected a seal count/map mismatch; the aggregate
validator subsequently rejected an extra nesting assumption for admission
attempts. Both failures are retained, and bounded reader-only corrections match
the frozen evidence. Raw samples and frozen drivers were not modified.

After independent review, cleanup verified all three executable hashes and
removed 1,808 owned build/scratch files totaling 3,767,474,281 logical bytes.
Post-cleanup replay passes using those retained removal witnesses. The final
seal binds the packet and six performance documents; unrelated changes remain
outside the commit.

The packet is at [change-0821](results/change-0821/README.md). Prior packets and
unrelated workspace changes remain untouched; no iWork work is included.

```sh
python3 -B docs/performance/results/change-0821/analyze.py --check
python3 -B docs/performance/results/change-0821/validate.py --final
python3 -B docs/performance/results/change-0821/seal.py --check-head
```

The first evidence commit exposed a post-commit reader defect: it required
HEAD to equal the measured base. The follow-up requires base ancestry while
retaining exact source-hash checks. The original seal and failed HEAD replay
are retained; the final seal verifies the aggregate batch changes since the
measured base across both normal commits. No measurement was rerun or changed.
