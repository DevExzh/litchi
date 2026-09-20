# 0717 — DOCX phase process-counter diagnostic

Status: retained opt-in benchmark enabler and diagnostic evidence. No production
optimization or native speedup claim. The rejected 0715 candidate remains rejected.

## Why this measurement

The previous batch was progress: 0716 reproduced a sustained timing shift on the
unchanged baseline and found more whole-child minor faults in the slow process.
Those counters included setup, warmup, untimed opening and verification, so they
could not locate faults in the serialization interval. This batch brackets that
interval with an explicitly instrumented, opt-in procfs probe.

`ordinary-save-process-metrics` is a standalone harness feature. The default
build performs no new procfs reads and omits probe metadata. The feature records
same-process counter deltas around each ordinary-save phase, after untimed
preparation and before owner drop, digest or readback verification. It works
with the existing allocator feature and carries a distinct instrumentation
identity in both combinations. The latency ABBA checker rejects both identities.
No production crate, dependency, preservation contract or ADR changes.

Each feature-enabled child also retains 32 adjacent empty snapshot pairs before
warmup. These record counter changes, including probe overhead; they do not
measure durations or quantify latency overhead. Controls are never subtracted.
Raw sample deltas retain acquisition order, and the analyzer checks all sixteen
fields against elapsed-sorted operation vectors through the sample permutation.

## Fixed protocol

The final native and procfs release executables have separate source-bound build
records. The host is an AMD EPYC 9R45 virtual machine with Rust 1.95.0. The corpus
pair is unchanged: a deterministic 200-paragraph generated document and the
55,519-byte NumberedList fixture, SHA-256
`ebb078d791b6deb4d0a2dae69a15a5274baf2ade815f07b336b394bb43af4c88`.

Sixteen blocks each contain all four corpus/lane treatments in cyclic Latin
order, repeated four times. Each treatment occupies each position four times.
All 64 children use CPU 12, 100 warmups and 200 measured in-process samples.
All 12,800 samples are retained. Cargo, repository gates and captures execute
serially in the coordinator lane. No replacement child, exclusion, acceptance
rerun or stop-on-good-result rule is used.

The measured operation writes an edited, opened DOCX to a bounded sequential
counting sink. Opening, editing, sink reservation and output verification are
outside its timer. This is serialization, not the ordinary end-to-end lifecycle
or atomic filesystem publication. Instrumented elapsed is diagnostic only;
there is no native/procfs speedup ratio. Within each lane, all process-summary
spreads and within-child four-block/half-mean spreads above 5% are descriptive
review flags, not formal stationarity tests or optimization gates.

## Observations

The following ranges span sixteen process summaries each. Values are ns;
spreads are `(max−min)/min`. Procfs rows describe only the instrumented lane.

| Corpus | Lane | p50 range | p50 spread | Mean range | Mean spread | p95 spread | p99 spread |
|---|---|---:|---:|---:|---:|---:|---:|
| generated | native | 317151–371337 | 17.09% | 320129.56–373356.85 | 16.63% | 16.85% | 16.44% |
| generated | procfs | 323707–368487 | 13.83% | 326834.23–369570.57 | 13.08% | 13.81% | 13.65% |
| numbered-list | native | 118261–122220 | 3.35% | 120179.59–124053.13 | 3.22% | 4.76% | 15.90% |
| numbered-list | procfs | 127645–133641 | 4.70% | 129280.82–135573.34 | 4.87% | 11.46% | 14.90% |

Eleven group metrics exceed 5%. Generated native p50 varies by 17.09% across
fresh processes. The NumberedList native medians in this matrix are much closer;
0716's sustained 176,706 ns median is not observed here. Its absence does not
establish stability or invalidate the previous observation.

All 6,400 instrumented samples have measured process counters and zero major
faults. Of 3,200 generated samples, 2,463 record exactly 78 minor faults, 662
record zero and 75 have other counts. Eleven generated children have at least
186 of their 200 samples at exactly 78 faults. B5 instead has 195 zero-fault
samples and no 78-fault samples. B8, B9, B11 and B16 mix the two regimes, as does
B2. The complete distributions and acquisition order remain in `analysis.json`.

The five mixed children give the following selected 0/78-fault buckets. These are
associations within the instrumented process, not randomized interventions or
counterfactual fault costs; small buckets do not establish a stable distribution.

| Generated procfs block | Zero-fault count | Zero-fault p50 | 78-fault count | 78-fault p50 |
|---|---:|---:|---:|---:|
| 2 | 13 | 324661 | 186 | 365486 |
| 8 | 167 | 325022 | 16 | 366777 |
| 9 | 28 | 325486 | 167 | 364782 |
| 11 | 105 | 326402 | 80 | 363437 |
| 16 | 154 | 324087 | 36 | 360957 |

This localizes a recurring fault/timing association to the probed phase interval
for the generated corpus. It does not establish that faults caused the elapsed
difference or identify which allocation or write incurred them. NumberedList
has 3,129 zero-fault samples out of 3,200; its remaining counts range from 1 to
13. The current data therefore do not explain 0716's slow NumberedList process.

Across the 1,024 empty controls, 994 have zero minor faults and 30 have one;
all have zero major faults. Controls are collected before warmup and are not a
matched subtraction baseline for each later operation. Their small counter
values do not quantify elapsed perturbation.

Every within-child flag is listed below. Eleven children exceed 5% across their
four sequential 50-sample means; two also exceed 5% across their half means.
All other children, sample statistics and context remain in the machine-readable
analysis. A flat within-child series does not imply cross-process repeatability.

| Child | Four-block mean spread | Half-mean spread |
|---|---:|---:|
| native-b1-generated | 5.72% | 4.38% |
| procfs-b4-numbered-list | 6.38% | 3.55% |
| native-b6-generated | 6.39% | 2.73% |
| native-b7-generated | 8.41% | 3.63% |
| native-b8-generated | 9.25% | 4.57% |
| procfs-b8-generated | 7.01% | 3.35% |
| procfs-b9-generated | 7.05% | 3.41% |
| native-b10-generated | 6.78% | 2.80% |
| procfs-b11-generated | 11.76% | 8.77% |
| native-b14-generated | 11.27% | 5.35% |
| procfs-b16-generated | 9.53% | 4.74% |

## Counter scope and limits

Snapshots read `/proc/self/io`, `/proc/self/stat` and `/proc/self/status`
sequentially. Their fields have different read instants and kernel scopes;
they are same-process evidence, not owner-exclusive atomic snapshots. Probe
activity and other activity reflected by each field can contribute. Clock ticks
are too coarse for fine-grained CPU utilization in these operations. RSS delta
is a nonnegative endpoint change, while peak RSS is process-lifetime high water,
not an operation peak. No allocator, scheduler, hardware or hash-seed cause is
inferred, and no controls or whole-child counters are divided by native elapsed
to claim a causal fraction.

The parent retains `wait4` and system snapshots separately. These cover the whole
child, including setup, warmup, opening, verification and reporting. CPU and
pressure context can include unrelated activity. Rust hash seeds, allocator and
ASLR layout, and unlisted environment fields remain uncontrolled. Perl seed
variables do not control Rust hashing. The generic filesystem isolation flags
remain inapplicable to this ordinary-save loop: each child has 200 in-process
samples and no dedicated filesystem evidence. This feature does not resolve
that separately recorded metadata ambiguity.

## Verification and next decision

Fresh harness verification passes: 531 default tests and 530 feature-enabled
tests, with one opt-in real-producer security corpus test ignored in each mode;
warning-denied all-target Clippy with both instrumentation features; warning-denied
rustdoc; formatting; and all 89 latency ABBA checker tests. Existing phase tests
also verify feature metadata and sample identity. The unused initial builds are
archived with their source, logs and removal witnesses; captures use only final
builds after the scope clarification and Clippy cleanup.

Production verification is reused, not represented as newly run: the production
source is byte-identical to 0713/0716. The retained results cover 4,995 tests,
92 passed/46 ignored doctests, scoped Clippy and rustdoc. Seven baseline-proven
PPTX/XLSB test Clippy exceptions remain explicit. Fresh repository gates and the
final report gate are recorded separately.

All 64 receipts pass executable, source, fixture, command, environment and
artifact checks. Normalized publication bytes, decoded targets, edit outcomes
and member identities match the retained 0716 baseline. The initial analyzer
attempt is archived: its broad filename glob included `source-baseline.json`
as a child artifact. The post-cleanup audit also caught an analyzer output field that changed from
`live-binary` to `cleanup-witness` despite identical verified executable identity.
Both analyzer attempts are retained; the output now records identity validation
consistently through either route. No captures were replaced or statistics changed.
Nine in-memory corruptions are rejected, including stale statistics, sample
permutation, CPU command, decoded size, negative counters, instrumentation,
raw-delta alignment, control count and scope. Exact replay, final audit and
artifact sealing pass after removal of the three owned scratch roots, with
both final executable identities retained as cleanup witnesses.

The feature is retained as a measured diagnostic enabler. The next bounded
investigation is phase-matched allocation-size and memory-mapping attribution
to identify where recurring faults arise in generated serialization, while
preserving allocation initialization and publication semantics. The current
association does not justify changing the allocator, suppressing page faults,
or accepting the rejected section-collection candidate. Native replication
remains necessary; no broader stability or performance conclusion follows.
The non-iWork goal remains active.

[Evidence and replay instructions](results/change-0717/README.md).
