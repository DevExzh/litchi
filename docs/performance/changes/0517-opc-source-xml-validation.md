# 0517: validate immutable source XML once per topology publication

A source-proven XML replacement or addition was fully validated twice within
one `SourceBackedPackage::write_topology_to_stream` call. The first check
already binds immutable candidate bytes to immutable destination limits and
content type. The later transfer-registration loop repeated the complete XML
scan. This batch replaces that second scan with `check_source_state()`, keeping
the late source freshness and cancellation fence.

Both initial validation paths remain intact. Replacements still check source
lineage/version, exact original destination bytes and Part identity; additions
still enforce destination content type and limits. Candidate reparse/readback,
signature and encryption refusals, graph validation, preservation planning,
source guards on bounded writes/flushes, and partial-output error handling
remain in their original owners. There is no public API or dependency change.

The source proof and ADR matrix are in the
[publication map](../results/change-0517/publication-map.md), and the
[independent semantic review](../results/change-0517/semantic-review.md)
checks both paths. New tests inject source changes and cancellation after the
initial proof, assert no output on refusal, check managed release, and require
one payload-sized validation work charge. Existing destination depth/event
limit tests now require their specific typed errors.

The unchanged DOCX paragraph harness exercises 24 shapes/routes/providers in
two fresh campaigns per binary, with 30 samples, three warmups and two internal
repeats per child. This is a same-route production comparison at base
`201f65094ff0ca6186fbf1e2ffa6b3641290253b`, distinct from batch-versus-repeated
API choice. The six owned-source profile arms each have two fresh Callgrind
captures per binary. Collection is restricted to the public publication method;
its instruction count excludes caller destruction of the returned snapshot,
which the native publication timer includes.

All raw CSVs, source/build/command receipts, profile data, analysis helpers,
review records and reproduction instructions are retained in
[change-0517](../results/change-0517/README.md). Whole-child hardware counters
are a separate baseline supplement. Managed budget counters are not allocation
measurements; GNU time RSS is not publication-local memory. Historical 0499/0500
flags and broad cold/range/native-producer/scaling coverage remain open.

OLE2/OOXML optimization remains active. ODF is deferred until that goal is
complete, and iWork remains outside this workstream.

## Measured result

All 48 same-route comparisons improve lifecycle p50 (1.87–14.09%), p95
(2.02–14.88%) and p99 (1.02–15.70%). Publication p50 improves 7.79–30.85%;
the spread includes the phase variability visible in the baseline. Whole-child
RSS changes range from −3.51% to +4.23%, with no >5% adverse RSS flag.
The six scoped profile arms reduce publication Ir by approximately 17.9–21.1%
in both repetitions. These are synthetic DOCX results, not measured speedups
for every shared-OPC caller.

Each cell below gives r1 / r2 candidate-versus-baseline percent changes.
Negative values mean less time or RSS. Individual phase statistics, confidence
intervals and all adverse short-phase flags remain in the
[native comparison](../results/change-0517/candidate-comparison.md) and its JSON.
The bootstrap intervals describe within-child samples; they do not establish
host-general confidence from only two fresh children per binary/case.

| Workload and route | Lifecycle p50 | Publication p50 | Whole-child RSS |
| --- | ---: | ---: | ---: |
| p128-k1-file-batch | -11.36% / -11.13% | -19.25% / -18.18% | -0.12% / -0.79% |
| p128-k1-file-repeated | -10.67% / -10.61% | -19.04% / -19.43% | +3.21% / +0.68% |
| p128-k1-owned-batch | -12.52% / -4.96% | -21.87% / -8.72% | +1.21% / -2.69% |
| p128-k1-owned-repeated | -4.91% / -12.52% | -7.79% / -22.57% | -2.96% / -0.72% |
| p128-k32-file-batch | -10.15% / -9.32% | -17.09% / -17.36% | +0.06% / -3.51% |
| p128-k32-file-repeated | -3.57% / -1.87% | -16.95% / -16.36% | -0.95% / -1.25% |
| p128-k32-owned-batch | -9.66% / -10.09% | -19.11% / -19.50% | +1.01% / -0.56% |
| p128-k32-owned-repeated | -3.86% / -3.70% | -10.61% / -20.16% | -2.30% / -0.11% |
| p128-k8-file-batch | -10.62% / -10.46% | -19.16% / -18.72% | -1.34% / -1.34% |
| p128-k8-file-repeated | -8.49% / -6.64% | -26.55% / -17.19% | +3.04% / -0.42% |
| p128-k8-owned-batch | -11.73% / -11.50% | -21.64% / -21.82% | +0.00% / -0.05% |
| p128-k8-owned-repeated | -9.58% / -5.51% | -30.85% / -10.01% | +1.83% / -0.28% |
| p512-k1-file-batch | -12.74% / -11.66% | -21.95% / -20.96% | +0.00% / +1.68% |
| p512-k1-file-repeated | -11.99% / -14.09% | -20.96% / -21.43% | +0.72% / +4.23% |
| p512-k1-owned-batch | -12.76% / -12.30% | -22.27% / -21.87% | +2.10% / +1.60% |
| p512-k1-owned-repeated | -12.39% / -12.85% | -22.29% / -22.24% | +0.17% / -0.06% |
| p512-k32-file-batch | -11.27% / -13.41% | -21.76% / -23.21% | +0.86% / -1.25% |
| p512-k32-file-repeated | -2.29% / -3.23% | -20.50% / -21.58% | -0.69% / +0.23% |
| p512-k32-owned-batch | -11.44% / -11.16% | -22.33% / -21.93% | -2.79% / +2.72% |
| p512-k32-owned-repeated | -3.11% / -3.00% | -21.76% / -21.71% | +2.27% / -0.96% |
| p512-k8-file-batch | -13.70% / -13.41% | -24.32% / -24.62% | -0.35% / +0.58% |
| p512-k8-file-repeated | -6.34% / -6.57% | -20.56% / -20.23% | +0.48% / -0.24% |
| p512-k8-owned-batch | -12.22% / -10.51% | -21.94% / -21.04% | +0.82% / +1.87% |
| p512-k8-owned-repeated | -7.12% / -7.34% | -21.16% / -21.64% | -1.57% / +0.61% |

No adverse lifecycle, publication or RSS threshold is triggered. Other phase
flags are retained: open, commit and drop show variable tails, and several
short-phase medians rise despite lower total lifecycle time. There are 93 phase flags across the two comparisons. For example, p512/K32
owned batch open p50 rises 49.08% and 71.69%; the latter is 13,880 to 23,830 ns.
Their cause remains unproven, and these phase-level review flags carry forward.
No claim is made that every individual phase improves. The performance retention decision is
based on repeatable end-to-end gains and the independent instruction mechanism,
not on averaging away these flags.

## Verification and retention

The all-features OPC/DOCX/XLSX/PPTX/XLSB suites pass 4,984 tests with no
failures or filtered tests. The one optional independent ZIP64 integration
test passes separately against a freshly generated Python ZIP64 source:
4,175,373 physical bytes, a 4 GiB declared large member, data descriptors,
and exact untouched-record preservation. The 44 ignored documentation examples
remain explicitly unexecuted. Total executed passing tests: 4,985.

Formatting, all-features workspace checking, warning-denied OPC/DOCX Clippy,
warning-denied rustdoc and crate boundaries pass. Boundaries retain the existing
64-package / 241-declaration topology with 11 explicit migration debt items.
The existing strict claim-registry gate is a separate check; it does not replace
this batch's native/profile/counter evidence validation. No new allocator,
parallel-scaling, cold-cache or native Office producer result is claimed.

Retain the optimization: all measured same-route end-to-end rows improve and
the instruction profiles identify the removed work. Source I/O, exact output,
and retained/released owner gauges match in all 3,168 paired rows, including
warmups. The lower Work charge follows the eliminated scan. The next target
is current-source snapshot construction, which remains roughly 54–66% of
publication instructions; any reuse must preserve source authorization and
required format validation/readback. The full OLE2/OOXML goal remains open.
