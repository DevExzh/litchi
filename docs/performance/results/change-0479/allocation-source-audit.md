# DOCX paragraph-tail allocation source audit

This is a source attribution for the 0479 pilot allocator reports. It does not
change the frozen DOCX implementation and it does not claim that allocator
requested bytes are copied-byte traffic. The report fields used here are the
single-sample 131,072-paragraph records in
`pilot-allocator-total.report.json` and `pilot-allocator-phases.report.json`.

## Observation

The allocator-only total pilot reports the following lifecycle interval at
131,072 paragraphs:

| interval | allocation calls | reallocations | allocated bytes | deallocated bytes | live before/after | region peak |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| total | 3,653,986 | 507,973 | 549,272,923,585 | 549,272,923,585 | 699,964 / 699,964 | 42,493,834 |

The phase pilot is a separate execution. Its allocation totals are:

| phase | allocation calls | reallocations | allocated bytes | deallocated bytes | live before → after | region peak |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| open | 151 | 14 | 306,689 | 303,281 | 706,181 → 709,589 | 839,158 |
| snapshot | 913,457 | 126,991 | 137,315,271,396 | 137,306,751,128 | 709,589 → 9,229,857 | 11,327,556 |
| stage | 913,439 | 126,985 | 137,317,221,514 | 137,308,701,562 | 9,229,857 → 17,749,809 | 19,847,484 |
| commit | 0 | 0 | 0 | 0 | 17,749,809 → 17,749,809 | 17,749,809 |
| publish | 1,826,939 | 253,983 | 274,640,123,986 | 274,629,510,494 | 17,749,809 → 28,363,301 | 42,500,051 |
| drop | 0 | 0 | 0 | 0 | 28,363,301 → 706,181 | 28,363,301 |

The phase allocated-byte values sum to the phase-pilot total, but the phase
and total pilots are distinct executions. Their high-water marks and endpoint
values therefore are not expected to be identical.

## Primary source attribution: paragraph-range growth

The dominant source is `scan_document` in
`crates/litchi-docx/src/source_backed/paragraph_copy.rs`:

* It preallocates the paragraph-range vector to
  `limits.max_paragraphs.min(4_096)` at lines 1568–1580.
* For each non-empty paragraph, the `Event::End` path records one `Range` and
  calls `reserve_one` immediately before pushing it (lines 1653–1679). The
  generated corpus uses only non-empty `w:p` elements, so this is the path
  exercised by the pilot.
* `reserve_one` calls `try_reserve_exact(1)` whenever the vector is full
  (lines 1853–1859). Once the initial 4,096 entries are used, the operation
  consequently requests a one-element capacity increase for each remaining
  paragraph. This is the source-level explanation for the near-linear
  reallocation count.

For the 64-bit pilot target, `Range { start: usize, end: usize }` occupies 16
bytes (the definition is at lines 218–222). If each exact growth requests the
new vector size, the paragraph-vector portion is:

* source scan, 131,072 paragraphs: `16 × sum(4,097 ..= 131,072)` =
  **137,305,751,552** requested bytes;
* candidate scan, 131,073 paragraphs: `16 × sum(4,097 ..= 131,073)` =
  **137,307,848,720** requested bytes.

Those estimates explain the phase totals to within the non-range allocations:

| scan | observed allocated bytes | paragraph-vector estimate | residual |
| --- | ---: | ---: | ---: |
| snapshot/source | 137,315,271,396 | 137,305,751,552 | 9,519,844 |
| stage/candidate | 137,317,221,514 | 137,307,848,720 | 9,372,794 |
| publish/source + candidate | 274,640,123,986 | 274,613,600,272 | 26,523,714 |

The estimated paragraph-vector requests account for 99.9931% of snapshot,
99.9932% of stage, and 99.9903% of publish allocated bytes. The match is also
visible in the call counts. A source scan has six local-name clones per
paragraph in this corpus: three `Start` events and three `End` events. Those
clones use `checked_clone` (the event branches are at lines 1599–1608 and
1653–1659; the helper is at lines 1844–1850), so the expected direct allocation
count is approximately `6 × 131,072 = 786,432`. The observed direct count,
`allocation_calls - reallocation_calls`, is 786,466 for snapshot. For the
131,073-paragraph candidate, the expected count is 786,438 and the observed
stage direct count is 786,454. The small residuals include parser and package
setup allocations. Publish contains one source and one candidate scan, giving
an expected six-name-clone count of 1,572,870 against an observed direct count
of 1,572,956.

The same relationship appears in reallocations. The paragraph vector alone
requires 131,072 − 4,096 = 126,976 exact growth requests for the source scan
and 131,073 − 4,096 = 126,977 for the candidate scan. Snapshot reports 126,991
reallocations, stage 126,985, and publish 253,983 for the two scans together.
The small differences are consistent with other vectors or parser internals
also reallocating; the allocator counters do not carry call-site labels.

## Lifecycle mapping

The phase ownership explains which source operation produces each scan:

* `plain_paragraph_copy_snapshot_with_limits` obtains the main payload and
  scans it once at `paragraph_copy.rs:948–977`. This accounts for the
  source-sized scan in snapshot.
* `copy_fragment` reserves an exact complete output vector and copies the
  prefix, paragraph fragment, and suffix at `paragraph_copy.rs:1348–1368`.
  It then calls `source.with_xml(output)`, whose implementation scans the new
  candidate at `paragraph_copy.rs:279–287`. This accounts for the candidate
  scan in stage.
* `publish_plain_paragraph_copy_patch_to_stream` captures a new source
  snapshot and then applies the patch at `paragraph_copy.rs:1014–1021`.
  `Patch::apply` clones the complete `after` XML and calls `with_xml`, at
  `paragraph_copy.rs:413–424`; therefore publish repeats one source scan and
  one candidate scan before the OPC overlay begins.
* Publication then clones the target XML for the replacement payload at
  `paragraph_copy.rs:1028–1037`. The OPC implementation wraps that caller-owned
  vector in shared storage and validates/reads the original part before the
  changed-overlay writer at `crates/litchi-opc/src/source_backed.rs:7764–7827`.
  These steps, ZIP writer buffers, payload decoding/cache activity, source
  fingerprinting, and temporary parser values are plausible contributors to
  the publish residual, but the current counters cannot assign their exact
  portions.

The open phase is small and count-independent in this corpus. DOCX forwards
`from_read_at` to the source-backed OPC open at
`crates/litchi-docx/src/source_backed.rs:308–314`; the OPC open indexes the ZIP
catalog and builds the retained part catalog at
`crates/litchi-opc/src/source_backed.rs:5440–5624`. The measured 306,689
allocated bytes therefore describe package/catalog setup, while ordinary part
payloads remain deferred until the snapshot asks for `main.data()`.

## Requested allocation sizes versus copy traffic

`allocation_metrics::record_reallocation` adds the requested `new_size` to
`allocated_bytes` and the old allocation size to `deallocated_bytes`; it also
counts a realloc as an allocation call. The counter snapshot-to-sample
projection is in `tools/perf-baseline/src/allocation_metrics.rs:215–265`, and
the realloc callback accounting is in
`tools/perf-baseline/src/allocation_metrics.rs:539–565`. Thus the hundreds of
gigabytes are cumulative allocator-request sizes from repeated vector growth,
not hundreds of gigabytes written to the sink or read from the source.

The actual XML output in the 131,072-paragraph corpus is only 6,422,679 bytes
for the source and 6,422,728 bytes for the candidate. `copy_fragment` uses
`extend_from_slice` to populate the output, and `MeasureSource::read_at` uses
`copy_from_slice`; those memory-copy operations have no separate byte counter
in this report. Likewise, a system allocator realloc may copy the old vector
contents internally, but the harness records the allocator’s requested old
and new sizes rather than the bytes physically copied. `source_reads` measures
the public `ReadAt` requests and the sink record measures bytes accepted by the
hashing writer; neither is an allocation-copy metric.

The total region is intentionally whole-lifecycle scoped: it owns source
adapter/package creation, snapshot, stage, commit, publication, and drops in
`tools/perf-baseline/src/docx_plain_paragraph_tail_append.rs:908–938`. The
phase run uses one non-nested region per phase and retains the phase state until
the explicit drop phase at lines 978–1077. The live-byte increases between
snapshot, stage, and publish are therefore retained owners crossing phase
boundaries; only the complete total lifecycle’s equal live-byte endpoints
demonstrate zero net retention.

## Attribution hypothesis and proof limits

The primary optimization hypothesis is to eliminate the exact-one-element
growth loop in `scan_document` while retaining the finite paragraph limit and
the same refusal/readback contract. A bounded capacity or a single bounded
collection strategy would directly target the measured dominant source. The
per-event `checked_clone` allocations and the repeated full scans/copies in
patch application and publication are secondary candidates. Any optimization
must be measured against the same source identity, finite limits, exact XML,
reversible patch, raw ZIP preservation, and output-sink checks.

This audit does not prove a complete call-site accounting:

* the global allocator observer is operation-scoped and records requested
  allocator sizes; it has no source-location tags and does not expose physical
  copy traffic or allocator-internal realloc overlap;
* the residuals cannot be uniquely divided among XML decompression, package
  cache, ZIP preservation, hashing, fingerprints, parser buffers, and temporary
  vectors without call-site instrumentation or a profiler;
* total and phase pilots are separate executions, and phase peaks must not be
  summed or substituted for the total high-water mark;
* the public path does not expose an internal decompressed-byte or
  reallocated-byte counter, so those quantities remain explicitly unavailable;
* the exact capacity behavior is evidenced by the observed counts and the
  `try_reserve_exact(1)` source path on this frozen toolchain/target. A future
  allocator or toolchain may change implementation details while preserving
  the API contract.

The evidence supports a strong source attribution for the dominant cumulative
requested bytes, while leaving the smaller residual and physical memory-copy
traffic open for the later CPU/profile review.
