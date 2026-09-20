# 0716 — DOCX baseline process variation

Status: diagnostic evidence; no production or Rust benchmark changes. A large
between-process timing shift occurs on the unchanged baseline, even with 100
warmups. The rejected 0715 candidate is not needed for such a shift to occur.
This does not identify the cause of either experiment or reverse 0715's rejection.

## Scope and frozen protocol

The previous batch was progress: its frozen pilot rejected a section-collection
fusion after one NumberedList pair regressed about 42%. All production changes
were restored. This batch tests the restored baseline's repeatability before
another optimization decision.

One new release binary is built from revision `9c16221fae`. The source census
is identical to the restored 0715 and retained 0713/0714 source; the executable
has its own recorded identity. The corpora remain the deterministic
200-paragraph generated document and the 55,519-byte NumberedList fixture,
SHA-256 `ebb078d791b6deb4d0a2dae69a15a5274baf2ade815f07b336b394bb43af4c88`.
The host is an AMD EPYC 9R45 virtual machine with Rust 1.95.0; exact host,
compiler, storage and binary records accompany the packet.

Eight fixed blocks each contain both corpora at warmup counts 10 and 100.
Four cyclic orderings repeat twice, putting each treatment twice at every
position. All 32 fresh native processes use CPU 12 and 200 in-process measured
counting-publication samples. All 6,400 samples remain retained. Cargo and
native captures run serially in the coordinator lane. No candidate binary,
replacement child, sample exclusion, or stop-on-good-result rule is used.
The capture plan and driver were frozen after independent review and before
the first child.

This is the serialization suboperation of ordinary opened-document edit/save:
opening and editing occur outside its timer, and publication targets a bounded
sequential counting sink. It is not an end-to-end lifecycle, atomic-save,
cold-cache, throughput or scaling result. Phase measurements are not additive.
No ADR or preservation contract changes; the existing validation and publication
paths run unchanged. ADRs 0005/0006 and the CRUD scenario taxonomy remain binding.

## Between-process observations

Each range below spans eight process summaries. Spread is `(max−min)/min`.
Values are nanoseconds; every spread above the frozen 5% descriptive threshold
is a review flag, not an optimization gate or a formal stationarity test.

| Corpus | Warmups | p50 range | p50 spread | Mean range | Mean spread | p95 spread | p99 spread |
|---|---:|---:|---:|---:|---:|---:|---:|
| generated | 10 | 319266–354627 | 11.08% | 323028.30–351920.49 | 8.94% | 12.90% | 2.63% |
| generated | 100 | 318326–362222 | 13.79% | 320857.40–364012.18 | 13.45% | 13.72% | 13.72% |
| numbered-list | 10 | 121095–123571 | 2.04% | 122980.05–125770.35 | 2.27% | 4.57% | 33.76% |
| numbered-list | 100 | 120031–176706 | 47.22% | 121682.99–179384.86 | 47.42% | 48.19% | 48.01% |

All 12 over-threshold group metrics remain visible above. NumberedList
with ten warmups has a p99 flag despite relatively close medians and means.
Increasing warmup to 100 does not establish repeatability for either corpus.

The standout child is `native-b6-numbered-list-w100`: its p50 is 176,706 ns,
versus 120,031–122,105 ns for the seven other children in that group. Its minimum
is 171,751 ns, above every other child's p99 (the largest is 145,791 ns).
Its four successive 50-sample means are 179,673.32, 178,769.28, 179,396.24 and
179,700.58 ns. This is a sustained process shift, not an isolated upper-tail
sample. The source, binary, fixture and normalized published output are identical.

## Within-process observations

Execution order is reconstructed from the retained `sample_order` permutation
and checked against the operation-metric sample vectors. The following table
lists every child flagged by either its four 50-sample means or two half means.
The JSON retains those means and all raw samples for all 32 children.

| Child | Four-block mean spread | Two-half mean spread |
|---|---:|---:|
| native-b1-generated-w10 | 8.42% | 3.77% |
| native-b2-generated-w100 | 10.07% | 5.10% |
| native-b2-numbered-list-w10 | 5.84% | 3.57% |
| native-b2-generated-w10 | 7.10% | 3.00% |
| native-b3-generated-w10 | 10.29% | 6.71% |
| native-b4-generated-w10 | 10.99% | 9.08% |
| native-b5-generated-w10 | 10.78% | 7.19% |
| native-b5-numbered-list-w10 | 5.05% | 3.37% |
| native-b6-numbered-list-w10 | 6.36% | 2.41% |
| native-b6-generated-w10 | 12.34% | 11.52% |
| native-b7-numbered-list-w10 | 6.91% | 3.57% |
| native-b7-generated-w10 | 9.97% | 5.09% |

Twelve children have a four-block flag; six also have a half-mean flag. The slow
B6 NumberedList child has neither, illustrating why a flat within-process series
does not establish agreement between fresh processes. The harness retains raw
standard deviations and mean confidence intervals, but these do not justify
assuming independent samples or a stable between-process distribution.

## Warmup comparisons

These are descriptive pairs from each fixed block, expressed as
`(warmup100/warmup10−1)*100`. Position and global acquisition order are retained
for each member. More warmup also changes prior work and process state; the
comparison is not an isolated cache intervention.

| Block | Corpus | p50 delta | Mean delta |
|---|---|---:|---:|
| 1 | generated | -10.11% | -8.83% |
| 1 | numbered-list | -0.28% | -0.04% |
| 2 | generated | +0.64% | +0.54% |
| 2 | numbered-list | -2.86% | -2.81% |
| 3 | generated | +0.88% | +2.77% |
| 3 | numbered-list | -0.72% | -0.24% |
| 4 | generated | -0.17% | +2.91% |
| 4 | numbered-list | -1.11% | -1.11% |
| 5 | generated | -9.69% | -7.11% |
| 5 | numbered-list | +0.60% | +0.22% |
| 6 | generated | +2.19% | +5.44% |
| 6 | numbered-list | +44.13% | +42.63% |
| 7 | generated | +2.59% | +4.19% |
| 7 | numbered-list | -0.81% | -0.43% |
| 8 | generated | +11.32% | +11.15% |
| 8 | numbered-list | -1.55% | -1.79% |

## Process context and limits

The parent uses `wait4` after each child and records Linux snapshots before and
after execution. These cover corpus construction, warmup, untimed opens and
edits, output verification and report writing as well as serialization. CPU12
counter deltas and system pressure also include unrelated activity. No counter
is divided by the sampled serialization times to claim a utilization or causal
fraction.

The slow B6 NumberedList/w100 child records 47,209 whole-process minor faults;
the seven peers record 16,987–18,890. B6 has 85 involuntary context switches;
peers have 81–88. All eight have zero major faults. This association does not
locate faults within the timed owner or establish an allocator, scheduler,
frequency or hardware cause. CPU frequency/governor files are unavailable in
the captured sysfs records. No hardware performance counters were collected.

Rust hash seeds and allocator/ASLR layout remain uncontrolled. Perl seed
variables do not control Rust hashing. Only the listed environment fields are
constrained or recorded. Nothing here establishes a general memory improvement
or peak-RSS regression; retained RSS counters describe whole benchmark processes.

The generic report booleans `filesystem_fresh_child_per_sample` and
`filesystem_process_isolated` are emitted unconditionally by the harness. They
are not applicable evidence for the ordinary-save loop. Dedicated
`filesystem_evidence` is absent for these children; each has 200 samples in one
process. The read-only metadata review identifies this scope ambiguity for a
separate harness correction. This packet does not silently change frozen source.

## Verification and next decision

All 32 receipts pass exact source, executable, fixture, command, environment
and artifact checks. Output parity matches the retained 0715 baseline, including
decoded target identities, member manifests, edit outcome and publication bytes.
The analyzer recomputes raw statistics, sample order and every diagnostic row.
Five in-memory evidence corruptions are rejected without modifying retained
inputs. Post-cleanup replay and the artifact census pass; all three owned
scratch roots are removed with an exact executable witness.

No format correctness suite is represented as newly run. Exact-source 0713
verification is reused: 4,995 tests, 92 passed/46 ignored doctests, scoped Clippy,
warning-denied rustdoc and six repository evidence gates. Seven preexisting
PPTX/XLSB test Clippy exceptions remain explicit. The release harness build and
final report-classification check are fresh.

The next useful diagnostic is to localize process-state effects to the timed
owner with a separately frozen, opt-in resource/syscall observation protocol.
Whole-process fault counts motivate that probe but do not prove its mechanism.
A larger warmup alone is not a demonstrated remedy. Future candidate comparisons
must retain process replication and individual variation instead of treating a
flat sample sequence as a stable baseline. The 0715 rejection remains unchanged;
no optimization or speedup is retained in this batch. The non-iWork goal remains
active.

[Evidence and replay instructions](results/change-0716/README.md).
