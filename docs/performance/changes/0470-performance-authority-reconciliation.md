# Change 0470: reconcile current performance evidence and open scope

`performance_claim: none; evidence-authority and scope record`

`claim_authorized: false`

This record indexes performance evidence available on 2026-09-10. It keeps
historical controls, resource captures and the accepted optimization separate. The program remains open;
the records below are useful because their scopes do not overlap enough to be
combined into a current end-to-end latency or production-wide speedup claim.

## Source epochs and accepted claims

| record | measured source | what it establishes | boundary |
|---|---|---|---|
| [Full default allocator publication](../results/full-default-allocator-20260910/README.md) | `fccbe6595f3a29561a8bfc6c192d8fa5e874ad05`, with the frozen instrumentation patch; integration `bfdc911c0577c6a3304b4a70b8b70e2661721355` and publication commit `cf4f383c4c203381880d4c2e8445d2616a53d283` | The canonical default selection has 201 operation-scoped rows covering 37 cases and 31 corpora, 15 samples and three warmups, and 11 aligned allocation vectors per row. The portable verifier accepts digest `f0fd76293959e72211e06e51b0a2b41f371423fda34c55144077193d262b1670`. | The elapsed field is explicitly `allocator_instrumented_elapsed_not_latency_claim`. The source is historical relative to the publication integration, and the artifact does not prove current normal latency, throughput, RSS, I/O, or whole-program completion. |
| [Historical full baseline](../results/full-baseline-20260910/README.md) | `1b3f2c2d0c059e8e59272775a97d0567e14e67f2`, published from `cf98bee37455e8e6d0e73002ff2e94299ef546cb` | One reproducible normal report with 201 rows, 37 cases, 31 deterministic corpora, 15 samples and three warmups. Raw and derived hashes are retained by the publication. | This is a historical control, not the current feature HEAD. It has zero normal filesystem rows; `warm` and `cold-requested` are requested harness states rather than physical cache eviction. It makes no latency-improvement, scaling, native-producer, or production-optimization claim. |
| [Latest normal 201-row control publication](../results/normal-201-control-v2-20260910/README.md) | measured source `995bdaf09352297bde17e6ed7d986360f8c10134`, tree `7f02507971d036c174273bff0ad643c8338d6227`; packaging candidate base `fdd2b80dfa9efc246a5983a3fec64de25e5c4f63`, wrapper commit `be698d900369f8a6b6d2908f49e2f50ac15889dc` | Three fresh-process repeats retain 201 rows over 31 corpora, 15 samples and three warmups per row. The sealed verifier recomputes all 603 raw row sample sets and their p50/p95/p99 values; independent review and root verification passed. | The packaging candidate was prepared against `fdd2b80d`; the measured code predates later feature changes. This is one normal control without ABBA ordering or a paired optimization comparison. The raw matrix has zero filesystem rows, no common logical-bytes denominator for throughput, and GNU-time whole-process RSS rather than operation-scoped RSS. `warm` and `cold-requested` are harness selectors, not physical cache evidence; no current feature-head latency claim is authorized. |
| [Accepted plain worksheet-cell tag optimization](../results/xlsx-plain-cell-tag-20260910/README.md) | measured base `995bdaf09352297bde17e6ed7d986360f8c10134`, accepted production commit `556c61ef355bac77a1dcef0870209c68bf3290e9` | One dense-wide one-percent commit/save ABBA reports 5.568% lower pooled normal p50, 20.559% fewer allocation calls, and 7.079% fewer allocated bytes. Source-span retention preserves plain-cell edits/readback and untouched neighboring XML; root gates and independent review passed. | The region peak live-byte value increased by four bytes. This is one historical scenario on a shared host, not a suite-wide gain or current feature-head measurement. The corpus ZIP is reproduced by the recorded generator and hash rather than embedded in the publication. |
| [Committed `cf98bee` scaling publication](../results/paired-scaling-cf98-20260910/README.md) | `cf98bee37455e8e6d0e73002ff2e94299ef546cb`, tree `46c3a5c1c10c0178a4e31db4bf2f2d838ef2beab`, binary `727c6e718cf5d64aeb40edaadb2fa3e5c70b005d8c1e827bb296c12341a398b7`; publication commit `fdd2b80dfa9efc246a5983a3fec64de25e5c4f63` | The committed 104-file archive contains ten reports, ten catalogs, exact runners, source/build/host receipts, analyses and a portable verifier. Root verification and independent `opc_capture_review` both passed; all 16 scaling case/corpus/width combinations have matching work across captures. | The paired capture is a repeatability observation on a shared host, not a before/after comparison. Superlinear rows are model-invalid; process-wide lock-wait and utilization counters are unavailable. The range lane is deterministic simulation, not remote or physical-storage evidence. No Amdahl fraction, causal speedup, current normal latency, or production-wide scaling law is authorized. |

These artifacts retain distinct scopes. Allocator elapsed samples are excluded
from latency claims. The older normal report measures `1b3f2c2d`; the latest
normal control and accepted XLSX optimization measure `995bdaf09`. These code
snapshots predate later production changes. The scaling captures use a separate
synthetic workload family. The 201-row count is now reconciled with the checked comparator policy
and default contract; older 36-case/198-record prose must not be used for a new
comparison.

## Scaling observations retained for context

The paired publication reports raw p50 elapsed nanoseconds for fixed generated
workloads. The paired worker-1-to-worker-8 ratios are 3.036x for CFB
`few-large`, 1.555x for CFB `many-small`, 5.287x for OPC `few-large`, and
1.060x for OPC `many-small`. These are observations, not a common speedup
claim. CFB `many-small` is a warm MiniFAT-cache reread; OPC `few-large` has
four typed payload tasks and two serial metadata tasks. The captures use
workers as harness widths, not a proof that every logical task maps to a
process thread.

The first-capture simulated range records use one worker, 2,000 microseconds
fixed latency, 250 microseconds request overhead, 16 MiB/s bandwidth, and a
65,536-byte maximum physical range. They retain request counts and selected
payload reads, but they are not network, remote, physical-disk, or cold-cache
measurements.

## Latest normal-control publication and current gap

The [normal 201-row control publication](../results/normal-201-control-v2-20260910/README.md)
retains three fresh-process repeats from source
`995bdaf09352297bde17e6ed7d986360f8c10134` (tree
`7f02507971d036c174273bff0ad643c8338d6227`). Its packaging candidate was prepared against
`fdd2b80dfa9efc246a5983a3fec64de25e5c4f63`, and the measured production code predates later feature changes. The sealed verifier recomputes 603 raw row sample sets and the
derived p50/p95/p99 values; it is therefore a reproducible historical control,
not a current feature-head latency baseline. The raw matrix has zero filesystem
rows and no common logical-bytes denominator for aggregate throughput; RSS is
GNU-time whole-process maximum RSS. The current-normal requirement remains open
for a timed feature-head control with the required process, I/O, resource,
uncertainty, and warm/cold evidence.

## Exact dated audit snapshot

The 2026-09-10 row-level audit evaluated 264 requirements at the `cf4f383c4`
audit epoch: 50 checklist rows and 214 `GOAL.md`/program rows. Its status
counts are:

| status | rows |
|---|---:|
| complete | 10 |
| incomplete | 99 |
| weak | 116 |
| missing | 39 |
| **total** | **264** |

Those counts are a historical audit snapshot, not a current completion score;
the committed scaling and normal-control publications do not change them. The
portable [audit snapshot](../results/performance-requirements-audit-20260910/README.md)
contains the exact report, machine-readable matrix, and compressed original
`docs/GOAL.md` input. The audited
`docs/GOAL.md` input was the supplied untracked document with SHA-256
`bed4058bb76330daab8ce9d4bceff639ab3fbd7ea06634158bef41b133c4d1f1`; the
exact deterministic gzip is [`GOAL.md.gz`](../results/performance-requirements-audit-20260910/GOAL.md.gz)
and its compressed hash is recorded in that directory's `SHA256SUMS`. The
checked CRUD checklist has SHA-256
`d3a63f3aa001ad6e3dda7627b6294230ae87b3b5aaa595c45c3327e026dbf02a`. The
full goal text is not duplicated in this authority record; the portable audit
input is available through the retained gzip.

The remaining rows fall into these eight execution classes, ordered by the
audit's priority. This table preserves the distinction between a narrow
artifact that is complete and a program requirement that is not.

| priority | requirement class | current disposition | evidence still needed |
|---:|---|---|---|
| 1 | Evidence authority | This record reconciles the source epochs and 201-row policy; it does not create a current normal baseline. | Keep current source, publication, binary, and gate identities bound in one release-readable record. |
| 2 | Current normal baseline | Reviewed historical observations include the `995bdaf09` control and the later `cab92a6bb` default/provider capture published in `c44a1ac0c`; subsequent production changes are outside those measured source identities. | Process-isolated current feature-HEAD ABBA or repeated rows with p50/p95/p99, throughput, RSS, source/binary/corpus identity, uncertainty, and explicit warm/cold semantics. |
| 3 | Dense XLSX hotspot | The `556c61ef` plain-cell tag change is accepted for one historical dense-wide one-percent commit/save scenario; its 5.568% pooled-p50 and allocation reductions do not authorize an end-to-end or suite-wide improvement claim. | Matched current-feature-head latency, resource, copy, and validation measurements that retain no-op, source, unknown-content, and adverse-tail behavior. |
| 4 | Provider and sink axes | Four opt-in provider/sink axes are implemented in `cab92a6bb`. The [verified historical capture](../results/provider-sinks-capture-v3-20260910/README.md) published in `c44a1ac0c` contains 18 provider/sink and 603 default rows, with 15 samples per row and no gain claim. Corpus reproduction is identity-only, cold requests are advisory, and range reads are simulated; later feature-source evidence remains open. | Owned, borrowed, positional filesystem/`ReadAt`, simulated range, non-seek sink, and atomic-save rows with physical/logical I/O and copy/decompression accounting. |
| 5 | Scaling evidence | The paired publication is committed and independently reviewable for its descriptive fixed-work rows; program-level scaling remains open because lock-wait/utilization evidence and a valid common model are absent. | Retained raw reports and hashes, CPU utilization, task overhead, lock wait, efficiency, uncertainty, and model-valid serial-fraction analysis. |
| 6 | Native, security, and corpus breadth | Missing for the default performance contract. The 201-row matrix is primarily deterministic synthetic coverage. | Content-addressed producer/version corpora, inert encrypted/signed/protected/macro/external fixtures, malformed/adversarial bounds, very-large shapes, and corresponding resource/timing rows. |
| 7 | CRUD and streaming | Incomplete across conversion, large creation/append, bulk/delete/sanitize/copy/merge/patch/validation/repair and the four distinct append semantics. | Separate streaming-create, logical append, new-Part, and arbitrary-repackage scenarios with preservation, correctness, latency, and bounded-memory evidence. |
| 8 | Release and safety gates | Missing for the current whole-workspace certified scope. The allocator publication has scoped harness tests/Clippy and portable verification only. | Pinned workspace check, format, Clippy, rustdoc, boundary, unit/integration/doc/compile, fuzz, native, security, and cross-platform receipts. |

The row-level audit and its generated machine-readable matrix remain the
authoritative detail for individual requirements. This record does not alter
raw reports, derived baseline files, scaling captures, or their hashes.

## ADR disposition

ADR-0005's measurement contract requires representative corpus axes, latency,
resource, I/O, contention, and scaling evidence with profiles and statistical
support. ADR-0008's evidence levels separately require native/readback and
single- versus multi-thread evidence where applicable. The allocator artifact
meets its own operation-scoped resource contract, the accepted XLSX change
meets its named dense-wide ABBA contract, and the scaling publication meets its
own descriptive capture contract; none satisfies those ADRs for the full
program. Any later optimization record must preserve that boundary and keep
exact no-op, lossless, bounded, security, and explicit-execution contracts
intact.

The [provider/sink implementation](../results/provider-sinks-validation-v3-20260910/README.md) was committed in `cab92a6bbc2f2eb09081c69ce90ba7bb29509155`. It adds correctness-checked opt-in axes without claiming measurements. Release evidence binds production source and binary identities; documentation-only commits do not require repeating an unchanged code measurement.
