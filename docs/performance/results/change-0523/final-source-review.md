# 0523 final-source review

`status: source and measurement-boundary review complete`

`scope: frozen 0523 harness diff; no production optimization admission`

`plan revision: 0e6379307c5174babfc832edc6113a16c8e9233a`

I read the accepted ADR set recorded by `adr-manifest.json`, with particular
attention to ADR 0005's measurement contract, ADR 0006's validation and
compatibility boundary, ADR 0008's custody gates, and ADR 0026's OLE directory
ownership. OLE2/OOXML is the active priority; ODF remains deferred and iWork is
excluded.

## Frozen source custody

The hashes below were computed programmatically from the frozen worktree and
evidence files:

| Artifact | SHA-256 |
| --- | --- |
| [`tools/perf-baseline/src/lib.rs`](../../../../tools/perf-baseline/src/lib.rs) | `1d4f9ed81ffed0a65d3ff49572c50f2e8bf60867d4befeffc77d2174526fbafd` |
| [`baseline/source.patch`](baseline/source.patch) | `a37e4adbd02f44f2a8e0a0df6008f7262803689fc46d88a9a998c62aa1fb314d` |
| [`baseline/source-manifest.json`](baseline/source-manifest.json) | `de633055a8f19ef6d839a01ef716614287ebe910cb6cee2d440aa1668746bda1` |
| [`plan.json`](plan.json) | `cd85d0969f3f7d6b9127049147c302aa671949d3b01f4ed49735a3af0d4a7688` |

The current `lib.rs` hash equals its manifest entry. The 8,584-entry manifest
was checked for the same path. `source.patch` is byte-for-byte equal to
`git diff --binary 0e6379307c5174babfc832edc6113a16c8e9233a -- crates
tools/perf-baseline`; it is 4,380 bytes and changes exactly
`tools/perf-baseline/src/lib.rs` (63 insertions and one deletion). The current
production tree under `crates/` has no diff. No production file was edited in
this review; the frozen harness file contains the intended focused test.

## Preflight custody

The final focused preflight receipt is the one that binds the final source:

- `cargo test --release --locked --features allocator-metrics --lib
  cfb_open_operation_metrics_align_and_report_unavailable_allocator`, with
  `CARGO_BUILD_JOBS=2`, `CARGO_INCREMENTAL=0`, and the owned 0523 target;
- exit code `0`, with `1 passed; 0 failed` in the retained stdout; and
- `source_unchanged: true`, with the final `lib.rs` hash shown above.

The initial receipt remains under
[`preflight-initial/`](preflight-initial/). It also exited `0`, but its custody
guard correctly recorded `source_unchanged: false` and the earlier
`lib.rs` hash `75e04ccf63a5d8070cab08735aefcfa92f887bc5a6fd7107fe8073a5181df61d`.
The retained adjustment explains that the final two source-`NotApplicable`
assertions were added after the initial build began. It records
`production_changed: false`. The focused test was rerun against the final
unchanged source before freezing or capture. This review did not run a build,
test, benchmark, profile, or capture; the normal build and campaign receipts
remain root-owned.

## Review of the frozen CFB operation bracket

The only code change is additive instrumentation in `run_cfb_open`:

```text
allocation_metrics::begin()
Instant::now()
OleFile::open(Cursor::new(corpus.archive.as_slice()))
Instant::elapsed()
allocation_region.finish() or explicit unavailable sample
file_size() oracle
black_box(&ole)
retain an InProcessObservation for measured iterations
record_elapsed()
drop(ole) at the end of the iteration
```

The constructor clock itself is unchanged. The allocation region begins before
the clock and ends after the elapsed value is taken, so allocator metrics
include the timer calls while the timer does not include instrumentation work.
`finish` is before `file_size`, `black_box`, and the implicit `OleFile` drop;
those operations cannot enter the constructor allocation region. Report
construction occurs after the loop. Existing `expected_file_size`, input
slice, `record_elapsed`, and checked `elapsed_ns` behavior remain in place.

The constructor's `?` still returns an `OleFile::open` error at the same point.
When an active allocator region exists, its `Drop` path releases the region
token during that error path; the new observation conversion cannot mask or
reorder the constructor error. The file-size mismatch still occurs after the
constructor and before black-box retention, with the same error text. No
fallible reservation or production resource boundary was moved.

## Normal-mode status and alignment proof

The ordinary executable does not install the counting allocator wrapper, so
`allocation_metrics::begin().finish()` is disabled and the runner maps it to
`unavailable_sample()`. Every retained CFB observation therefore carries an
explicit `Unavailable` allocation status and the operation envelope omits all
allocation, byte, live-byte, and region-peak vectors. It cannot be confused
with a measured zero. The allocator target enables the existing wrapper and
uses the same bracket to publish checked samples; no global enable call was
added to the ordinary test path.

`from_in_process_observations_without_sink` receives exactly one observation
for each retained iteration. Warmups are excluded from both the elapsed
statistics and observations. Its existing `(elapsed_ns, original sample_index)`
sort produces `sample_indices` aligned with `Statistics.sample_order`; it also
rejects asymmetric or status-changing allocation observations rather than
combining partial vectors.

The focused test
`cfb_open_operation_metrics_align_and_report_unavailable_allocator` exercises
the tiny and few-large generated CFB shapes with one warmup and three retained
samples. It checks the three-sample count, exact sample-order equality,
`Unavailable` allocator status and scope, representative absent numeric
vectors, and raw-CFB source `NotApplicable` status with absent logical-read
vectors. It then independently reopens each archive and checks file size,
stream count, and target payload. The shared allocator test lock prevents
overlap with allocator-state tests. The test is harness-only and does not
change CFB behavior.

## Timer, oracle, provider, and profile boundaries

The measured owner remains `OleFile::open(Cursor<&[u8]>)` over already
materialized corpus bytes. Fixture generation and its setup validation open
remain outside the timed loop. The raw cursor has no logical `ReadAt` observer,
so source metrics are explicitly `NotApplicable`; no provider or source
counter was fabricated. The existing `CfbSharedOpen` source-observed path is a
different case and was not changed.

The exact profile owner remains `litchi_cfb::file::OleFile<R>::open` in the
frozen plan. The driver, setup dump separation, positive incoming constructor
dumps, timer exclusion, file-size oracle, report construction, and final drop
boundaries are unchanged. Whole-child process and PMU metrics remain separate
diagnostics and are not relabeled as constructor-local values.

The retained 0511 source-open **control** and **final** profiles are historical
comparison records: the control precedes FAT batching, while the final profile
is after FAT batching. Neither is a current 0523 profile. In particular, the
old control `load_fat` share of `9.30%` is not reused as a current attribution
or optimization decision. Fresh 0523 profiles and allocator captures must
establish the current ranking before the deferred chain candidate is
considered.

## Disposition

The frozen harness diff passes this bounded source-custody and measurement-
boundary review. It supplies the missing CFB operation allocation envelope
without changing production CFB/XLS code, providers, selectors, timer scope,
or correctness/drop boundaries. It supports the root-owned release and
allocation/profile campaign, subject to those receipts and their existing
quality gates. It makes no speedup, allocation improvement, RSS, physical-I/O,
native Office, cold-cache, provider, or broad CRUD claim, and it does not admit
the next production visited-bit candidate.
