# 0715 — DOCX section publication collection

Status: rejected pilot; all three production files and their tests are restored
to the exact baseline. The candidate passes 46 of 48 hard gates, but the first
NumberedList serialization pair fails both median and mean nonregression gates.
No production optimization or speedup is retained.

The private ordinary DOCX writer previously collected owned section
header/footer parts and explicit relationship IDs in two separate body
traversals. Each traversal scanned every preserved paragraph for section
properties. The rejected candidate collects both outputs in one validated traversal.
Its archived patch changes only `package/codec.rs`, `writer/doc/model.rs` and
`writer/doc/package.rs` within `crates/litchi-docx/src`.

The body-final section remains first, followed by paragraph sections in body
order. The same section validator runs unconditionally; the paragraph-range
helper is unchanged. Both outputs are gathered before the existing part-key
conflict/dedup pass, then relationship tuples are sorted and deduplicated.
Tables, preserved body-final section XML, alt chunks and opaque nodes remain
outside this collector. There is no new cache, public API, dependency, unsafe
code or parallel execution. ADRs 0005/0006 preservation and validation rules
remain binding; no architectural change or new ADR is needed.

## Measured motivation

Four fresh profiles on baseline revision
`653a87ffef` precede the candidate. Each selects the exact `write_plain`
owner by its positive raw incoming edge from `ordinary_save::run_case`.
Generated profiles retain five setup dumps plus the measured dump; NumberedList
retains four setup dumps plus the measured dump. Zero-Ir termination dumps
remain retained. The self/direct-child partition reconstructs each owner.

| Corpus | Repeat | Owner Ir | Parts collector Ir | Relationship collector Ir |
|---|---:|---:|---:|---:|
| generated | 1 | 7,262,870 | 1,710,383 | 1,715,263 |
| generated | 2 | 7,306,422 | 1,737,094 | 1,743,130 |
| numbered-list | 1 | 3,043,920 | 27,800 | 27,612 |
| numbered-list | 2 | 3,046,291 | 29,086 | 28,537 |

These are guest-instruction observations, not native latency fractions or
hardware instructions. The generated corpus is the primary optimization
workload because the duplicated collectors are substantial there. NumberedList
is a real-file nonregression guard. The change retains the existing parse
inside `paragraph_section_range` and its caller's subsequent section parse;
it removes the second traversal without redesigning that helper.

## Frozen comparison

The pair consists of the deterministic 200-paragraph generated-medium DOCX
and the admitted 55,519-byte `NumberedList.docx` at
`test-data/libreoffice-core/sw/qa/extras/ooxmlexport/data/NumberedList.docx`,
SHA-256 `ebb078d791b6deb4d0a2dae69a15a5274baf2ade815f07b336b394bb43af4c88`.
Host/toolchain evidence and exact binary identities are retained in the packet.

A1 runs on baseline source before implementation. B1/B2/A2 run with candidate
source checked out and use the appropriate frozen executable. Each stage has
native and allocator children for both corpora and counting-publication /
lifecycle phases: 32 children total. Native uses 100 samples and ten warmups;
allocator instrumentation uses three samples and no warmup. CPU 12 is pinned,
and reverse stages reverse corpus/phase order. Cargo, native timing and initial profiles were serialized in the coordinator
lane. The post-acceptance RSS lane was deferred after rejection.

The frozen gates require generated counting-publication p50 and mean to improve
by at least 10% in both pairs. NumberedList counting publication and both
lifecycle workloads must regress by no more than 3% in p50/mean. Allocation
request counts and requested bytes must regress by no more than 3%. Tail and
repeat spreads above 5% are review flags, retained individually. The phases
are independently sampled and cannot be added or subtracted; allocator elapsed
values are not native latency.

## Pilot decision and native observations

The candidate is rejected under the frozen rules. Generated serialization
improves in both pairs, but the first NumberedList counting-publication p50
regresses 42.30% and mean regresses 41.32%, exceeding the 3% limits. The second
NumberedList pair is approximately flat; it does not erase the first failure.
No capture was rerun, dropped or averaged away to obtain acceptance.

Values are nanoseconds; deltas are candidate relative to its paired baseline.
Raw sample vectors, standard deviations and 95% mean intervals are retained
in each report and recomputed by the analyzer.

| Pair | Corpus | Phase | Baseline p50 | Candidate p50 | p50 delta | Mean delta |
|---|---|---|---:|---:|---:|---:|
| pair-1 | generated | counting_publish | 319611 | 243521 | -23.81% | -21.87% |
| pair-1 | generated | lifecycle | 5765883 | 5583648 | -3.16% | -4.06% |
| pair-1 | numbered-list | counting_publish | 123345 | 175516 | +42.30% | +41.32% |
| pair-1 | numbered-list | lifecycle | 5761508 | 5731844 | -0.51% | +0.09% |
| pair-2 | generated | counting_publish | 374196 | 248131 | -33.69% | -31.60% |
| pair-2 | generated | lifecycle | 5696768 | 5597852 | -1.74% | -1.71% |
| pair-2 | numbered-list | counting_publish | 123345 | 122975 | -0.30% | -0.53% |
| pair-2 | numbered-list | lifecycle | 5729348 | 5737818 | +0.15% | -0.25% |

Three native tail regression flags remain explicit:

| Pair | Corpus | Phase | Metric | Regression |
|---|---|---|---|---:|
| pair-1 | numbered-list | counting_publish | p95 | 41.23% |
| pair-1 | numbered-list | counting_publish | p99 | 36.80% |
| pair-1 | numbered-list | lifecycle | p99 | 17.23% |

All 16 repeat-spread flags follow (spread is `(max−min)/min`). Eleven belong
to native timing and five to instrumented allocator timing. The latter do not
become native latency evidence or allocation-count variability claims.

| Lane | Source | Corpus | Phase | Metric | Repeat spread |
|---|---|---|---|---|---:|
| native | baseline | generated | counting_publish | p50 | 17.08% |
| native | baseline | generated | counting_publish | mean | 16.59% |
| native | baseline | generated | counting_publish | p95 | 16.70% |
| native | baseline | generated | counting_publish | p99 | 15.51% |
| native | baseline | generated | lifecycle | p95 | 7.35% |
| native | baseline | generated | lifecycle | p99 | 17.65% |
| native | candidate | numbered-list | counting_publish | p50 | 42.72% |
| native | candidate | numbered-list | counting_publish | mean | 41.68% |
| native | candidate | numbered-list | counting_publish | p95 | 38.18% |
| native | candidate | numbered-list | counting_publish | p99 | 33.92% |
| native | candidate | numbered-list | lifecycle | p99 | 17.25% |
| allocator | baseline | numbered-list | counting_publish | p50 | 10.62% |
| allocator | baseline | numbered-list | counting_publish | p95 | 6.72% |
| allocator | baseline | numbered-list | counting_publish | p99 | 6.72% |
| allocator | candidate | numbered-list | counting_publish | p50 | 28.04% |
| allocator | candidate | numbered-list | counting_publish | mean | 8.51% |

The slow B1 NumberedList child is not an isolated tail: all 100 samples exceed
A1's median, 99 exceed its p95, and 98 exceed its maximum. Reconstructing
execution order gives four sequential 25-sample means of 175,346.9,
178,529.8, 175,981.3 and 178,192.4 ns. This is a sustained child-level shift.
[The sample-order diagnostic](results/change-0715/child-shift.json) binds all
four raw reports; it does not identify the cause.

The baseline generated counting median also drifts by 17.08%, while candidate
NumberedList counting medians differ by 42.72%. This establishes repeat
instability in this packet. It does not identify a hardware, allocator,
frequency, scheduling or code-level cause, and is not an exception to the
rejection rule. A follow-up must diagnose the timing shift with a newly frozen
protocol and retain this failed evidence; rerunning until the same gate passes
would not resolve it.

## Allocation observations and deferred diagnostics

Allocation p50 values below agree across both repeats within each source.
These measurements describe the rejected candidate only.

| Corpus | Phase | Baseline calls | Candidate calls | Baseline requested bytes | Candidate requested bytes | Net live bytes, both | Peak above start bytes, both |
|---|---|---:|---:|---:|---:|---:|---:|
| generated | counting_publish | 6522 | 3722 | 801740 | 738540 | 21575 | 544147 |
| generated | lifecycle | 15616 | 12816 | 3199408 | 3136208 | 461654 | 1049866 |
| numbered-list | counting_publish | 752 | 712 | 1113016 | 1111866 | 6487 | 529073 |
| numbered-list | lifecycle | 4674 | 4634 | 2661485 | 2660335 | 122437 | 710663 |

All 32 allocation hard gates pass. Exact normalized output parity passes
across all 32 native/allocator children, including decoded target identities,
member manifests and publication digests. Reduced allocation work and passing
correctness checks do not override the native nonregression failures.

The post-acceptance plan for four candidate profiles and 16 RSS children is
**deferred**, because its admission condition failed. The capture plan/scripts
remain retained; no candidate instruction-reduction or RSS claim is made.


## Correctness, custody and limits

Eight focused tests compare the fused result with a test-only copy of the
original two-pass collection. They check first-seen part order, exact lexical
XML and role conflicts, relationship tuple deduplication, ignored node kinds,
malformed placement/namespaces, unconditional validation, and the precedence
of a later malformed section over an earlier cross-section conflict.

For the rejected candidate, all 1,461 affected DOCX all-feature/all-target tests pass. Warning-denied
all-target Clippy, 75 doctests and warning-denied rustdoc pass; 31 doctests are
ignored. The first quality attempt caught a denied unused-qualification lint
after the new type import; it is retained under `quality-attempts/01`. The
qualification was corrected before the final quality checks and candidate
captures. No lint allowance was added. Other format suites are not claimed as
new runs in this DOCX-local batch.

Source snapshots, the exact patch, full source census, commands, fixtures,
raw reports, setup profiles and per-child receipts remain retained. All six
repository evidence gates passed on the candidate. Three profile-corruption
checks and four pilot-corruption checks reject altered evidence; exact
post-cleanup replay and the artifact manifest bind the final result. All three owned scratch roots are
removed after verification, with exact executable identities preserved.

The restored source is byte-identical to 0713/0714. Its prior exact-source
quality and evidence records are explicitly reused, including 4,995 tests,
92 passed/46 ignored doctests and seven baseline-proven PPTX/XLSB test Clippy
violations. Full all-target Clippy for those two crates is not claimed clean.
The candidate-only checks above do not replace that final-source scope.

The experiment preserved full paragraph/section validation and the durable
atomic publication contract established in 0714. It does not establish a hardware
cause, physical latency floor, cold-cache result, scaling result, throughput
claim or native Office compatibility result. The non-iWork goal remains active.
[Evidence and replay instructions](results/change-0715/README.md).
