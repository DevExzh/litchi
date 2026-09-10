# 0498: bounded source-backed Part batches

`SourceBackedPackage::read_parts_ordered` adds an explicit ordered multi-Part
read operation. A caller-supplied execution context can admit local worker
waves bounded by worker count, task count, declared bytes, minimum work, and
memory reservations. Calls without an execution context remain serial. This
is a low-level OPC capability; ordinary format CRUD methods do not silently
start parallel work.

The owning `PartBatch` keeps its collection memory and object reservations
until drop. Each `PartData` retains its existing payload reservation. Duplicate
requests preserve order and may share cached payloads; occurrences count toward
the `max_parts` request cap. Managed metadata is prepared once, while unmanaged
reads retain the ordinary one-Part path. Package total-byte limits retain their
catalog meaning and are not reapplied to repeated handles.

Each wave checks source freshness and cancellation before admission and joins
all admitted workers. Errors preserve their original typed fields, including
resource-limit scope; the lowest input ordinal wins among observed member
failures, subject to final source and cancellation fences. No partial batch is
returned. Actual `Work` and `InputBytes` consumption remains monotonic, and
successful cache admissions can survive a later failure. Explicit two MiB
worker stacks are included in scheduler memory reservations. Deflated-member
stress coverage supports the exercised decoder path, not a universal stack
bound for arbitrary caller provider code.

## Measured behavior

The [final evidence](../results/change-0498/README.md) uses one frozen executable
for ordinary serial `PartView::data` and the production batch API, with optional
operation accounting disabled. Thirty children contain 1,800 measured samples
and 180 warmups: two corpora, three sources, serial and batch widths 1/2/4/8,
three warmups and thirty measurements in each of two repeats. Every child
passed byte-digest, logical-length, cold-load-count, released-budget, and
scratch-cleanup checks.

The table gives pooled median microseconds across both repeats. It is scoped
to this synthetic corpus and shared host; per-repeat tails and adverse flags
are retained in the [analysis](../results/change-0498/final-analysis.md).
The aggregate comparisons retain 50 adverse flags: 29 latency, seven
throughput, and 14 whole-child RSS. The 48 per-repeat comparisons retain
71 latency/throughput flags; RSS is not counted again per repeat.

| Corpus / source | Serial | Batch 1 | Batch 4 | Batch 8 |
| --- | ---: | ---: | ---: | ---: |
| Four 1 MiB Parts / owned | 575.9 | 581.7 | 82.4 | 82.5 |
| Four 1 MiB Parts / warm file | 578.6 | 577.2 | 91.0 | 91.7 |
| Four 1 MiB Parts / short-read delay | 82,119.6 | 82,474.8 | 20,028.2 | 20,016.4 |
| Sixty-four 16 KiB Parts / owned | 94.8 | 93.9 | 560.0 | 480.6 |
| Sixty-four 16 KiB Parts / warm file | 182.3 | 201.8 | 577.2 | 512.8 |
| Sixty-four 16 KiB Parts / short-read delay | 29,572.8 | 29,600.1 | 8,016.4 | 4,190.6 |

The many-small memory/file regressions are material. Per-wave thread creation
is a plausible contributor, but the measurements do not isolate its cost.
Callers must choose execution policy for their workload; these results do not
justify enabling parallel work by default. The delayed many-small case shows
about 7.06x lower median latency with eight workers than ordinary serial reads.
The four-Part corpus cannot use more than four independent Part tasks.

The few-large owned/file medians appear superlinear at four workers. Source
calls, bytes, cold loads, and cumulative work match the serial route, but
allocator/cache and shared-host effects are not isolated. These observations
must not be fitted to a negative or clamped Amdahl serial fraction or generalized
as a scheduler-only speedup. Even the delayed four-Part ratio slightly exceeds
four and is outside the simple model at that width.

Historical controls were frozen before production edits, but their serial
route enabled optional accounting. They are retained as descriptive before/
after evidence with that instrumentation difference. The same-executable
controls above are the primary scaling comparison. Timing excludes fixture
construction, package open, and digest verification. Whole-child RSS and six
perf counter captures include those phases and cannot supply operation-local
CPU or memory attribution. File runs are warm; synthetic provider delay is not
a production remote-service measurement. See the
[measurement scope](../results/change-0498/measurement-scope.md).

## Validation and remaining program scope

Fourteen focused integration tests cover order and duplicate allocation sharing,
short reads, independent task/byte admission, serial fallbacks, deflated payloads,
typed limits and resource scope, source changes, cancellation, ordinal errors,
no later wave after failure, and structural/payload reservation release.
The default suite passed 619 tests and the all-feature suite passed 641,
with the same external-corpus ZIP64 test ignored in both. Five documentation
tests, warning-denied lint and rustdoc, formatting, crate-boundary checks, and
downstream DOCX/XLSX/PPTX/XLSB checks passed.
The reviewed source has no deliberate operation panic and adds no unsafe code,
global pool, runtime dependency, or facade-to-archive dependency.

The full non-iWork goal remains open. This batch supplies measured low-level
parallel Part reads; it does not establish broad end-to-end CRUD improvements,
controlled cold-cache behavior, native-producer validation, allocation-count
attribution, arbitrary history/composition, or a program-wide tenfold result.
