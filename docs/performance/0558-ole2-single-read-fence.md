# 0558: one source-version fence per shared CFB read

Status: retained, with one documented layout regression that change
[0559](0559-cfb-ascii-simple-uppercase.md) removes.
`performance_claim: none` — this record carries paired and deterministic
measurements, not a registry claim.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred
until that goal completes; iWork is excluded.

## What was removed

`SharedOleFile::read_stream_range` and `SharedOleStreamCursor::read_exact` each
observed `ReadAt::version()` twice per call: once immediately before the payload
read and once after it. Only the leading observation was removed. Every read
still fences, and no read path became unfenced.

On a `litchi_core::FileSource` each observation is one `statx` syscall, so the
baseline issued two `statx` calls for every positional read the shared CFB
reader performed.

## Source-level argument

Change [0325](changes/0325-cfb-frame-transaction-rejected.md) requires a
source-level argument before any freshness boundary moves. This is that
argument, including the costs.

`expected_version` is captured once while opening, and the open path fences
before and after parsing. Every checked operation ends with a fence
(`finish_stream_range` / `finish_cursor_operation`), which compares against
`expected_version` on success and, on the error path, prefers
`OleError::SourceChanged` over the original error. A read is therefore still
bracketed by two successful observations of `expected_version`: the most recent
earlier one taken by the shared reader — for the first read, the capture taken
while opening it — and its own trailing fence.

What is preserved:

- ADR 0005's rule that mutation during a read returns `SourceChanged`. The
  retained fence is exactly the one that observes a mutation spanning the read.
- Error precedence when a read or chain error races a mutation: the trailing
  fence still promotes `SourceChanged` over the original error.
- Operation boundaries, cancellation granularity, and the number and size of
  physical reads. Nothing is coalesced and nothing is read ahead.
- All bounds and stream-lookup errors, which already preceded the removed
  observation in `read_stream_range`.

What changed, stated plainly:

- **Reverted-mutation sampling.** `FileSource` latches its revision only when
  `version()` is called, so halving the observations halves the sampling of
  that latch. A mutation that is made and fully reverted between the previous
  fence and this read's trailing fence is no longer observed. A mutation that
  persists is still observed by the trailing fence, and by every later one.
  `litchi_core::FileVersionPolicy` already documents that a transition fully
  reverted between observations is not visible and that callers who do not
  trust other writers must validate the bytes they consume.
- **Error identity when the fence itself fails.** If `version()` cannot observe
  the source *and* the payload read fails, the payload error is now reported
  where the leading fence previously reported the observation failure. When the
  read succeeds, the observation failure is still reported instead of the bytes.
  Both behaviors are now covered by
  `read_exact_reports_the_payload_error_when_the_fence_also_fails` and
  `read_exact_reports_a_fence_failure_after_a_successful_read`.
- **Payload I/O on an already-changed source.** The read is attempted before the
  change is reported, so the destination may receive bytes that the caller must
  discard. The existing contract already required discarding the destination on
  any error. Chain state comes from the validated in-memory index, never from
  the mutated source, so no size, allocation, or sector count can be inflated.

This is not the design of change
[0279](changes/0279-cfb-operation-freshness-session-rejected.md), which allowed
a successful read to complete with no fence at all, nor of change 0325, which
proposed cross-operation coalescing and overread. Both of those changed which
operation receives an error; this one does not.

The `read_exact`, `stream_cursor_at` and `skip_forward` rustdoc, and the
ADR-compliance matrix, are updated to describe the fence discipline that the
code now has. The 0277 matrix row's "before/after" wording is superseded for
these two entry points.

## Measured effect

Host, corpus, build and command identities are in
[`environment.json`](results/change-0558/environment.json),
[`plan.json`](results/change-0558/plan.json) (frozen before capture) and
[`binary-identity.json`](results/change-0558/binary-identity.json). The plan
pinned CPU 17 because `mpstat` sampled the program's usual CPU 2 at 93.97% busy;
host-wide quiescence is not established.

### Deterministic counters

Corpus `test-data/ole/xls/ConditionalFormattingSamples.xls`, 1,402,368 bytes.

| Mode | Operation | `version()` calls | `statx` (1 warmup + 5 samples) | `read_at` calls | read bytes |
| --- | --- | --- | --- | --- | --- |
| file-source | open | 1,266 → 644 (−49.13%) | 7,630 → 3,898 (−48.91%) | 655 → 655 | unchanged |
| file-source | list | 1,266 → 644 (−49.13%) | 7,630 → 3,898 (−48.91%) | 655 → 655 | unchanged |
| file-source | one-cell | 1,813 → 925 (−48.98%) | 10,912 → 5,584 (−48.83%) | 921 → 921 | unchanged |
| owned-readat | one-cell | 1,813 → 925 (−48.98%) | 4 → 4 | 921 → 921 | unchanged |

`pread64` counts are byte-for-byte identical in every cell, and every semantic
projection is identical per cell. The owned control confirms the syscalls came
from the file adapter, not from the reader's logical work.

### Paired latency

Two repeats of A1/B1/B2/A2, one fresh child per cell, 20 warmups and 100 samples
each: 48 children. **All 48 case/statistic comparisons improve in both ABBA
directions** across p50, mean, p95 and p99.

| Mode | Operation | p50 first direction | p50 second direction |
| --- | --- | ---: | ---: |
| file-source | open | −26.88% / −28.21% | −27.69% / −28.98% |
| file-source | list | −27.87% / −27.82% | −27.74% / −27.90% |
| file-source | one-cell | −28.64% / −28.53% | −28.44% / −28.03% |
| owned-readat | open | −14.75% / −14.71% | −14.52% / −14.85% |
| owned-readat | list | −14.45% / −14.35% | −14.54% / −14.45% |
| owned-readat | one-cell | −15.98% / −16.39% | −16.41% / −16.16% |

Each cell shows repeat 1 / repeat 2. Maximum same-binary drift across the paired
legs is 5.847% and maximum cross-repeat drift is 5.668%; both occur only on p99
cells, well below the observed effect. The owned lane counts `version()` calls
through an instrumented adapter, so part of its improvement is instrumentation
that production code does not pay.

### Guardrail matrix and the regression

Fourteen CFB/DOC/PPT/XLS/OLE-common selectors ran in all four ABBA legs: 56
case/corpus rows and 224 statistic comparisons. Thirteen comparisons are adverse
in both directions by more than 5%, concentrated in **one case**:

| Case | Corpus | p50 |
| --- | --- | ---: |
| `ole_common_one_edit_save` | many-small-incompressible | +26.22% / +27.04% |
| `ole_common_one_edit_save` | wide-root-incompressible | +41.01% / +40.77% |
| `ole_common_one_edit_save` | many-small-compressible | +27.54% / +12.08% |
| `doc_semantic_full_text` | doc-tiny (p99 only) | +75.52% / +34.49% |

With ASLR disabled the `ole_common_one_edit_save` many-small regression
reproduces deterministically at +26.0%/+26.5%/+26.7% over three repeats.

It is not extra work. One Callgrind child over the identical corpus measures:

| Counter | baseline | candidate | change |
| --- | ---: | ---: | ---: |
| Ir | 3,338,034,198 | 3,337,980,740 | −0.0016% |
| conditional branches | 37,030,781 | 37,030,751 | −0.0001% |
| conditional branch misses | 660,113 | 732,651 | **+10.99%** |

Per-function branch attribution assigns 70,571 of the 72,538 additional
simulated misses to `<core::char::ToUppercase as Iterator>::next` — a Unicode
case-mapping iterator used by CFB directory-name handling, whose own instruction
count is unchanged — so the cause is code placement, not the fence. The baseline
is itself bimodal on this case: the same binary and corpus produce 37.0 ms or
51.4 ms medians across children. Change 0559 removes that hotspot; after it the
case runs at 1,496–1,527 us against the pre-0558 baseline's 1,784–1,807 us on
many-small, and 34.2–34.6 ms against 36.9–51.3 ms on wide-root, without the
bimodality.

After 0559, a 21-selector A1/C1/C2/A2 guardrail against the pre-0558 baseline
records 222 of 376 comparisons improving in both directions and only two adverse
in both directions by more than 5%, both p99-only. See the
[0559 record](0559-cfb-ascii-simple-uppercase.md) for that table; it combines
both changes and is not attributed to either alone.

## Correctness evidence

4,904 tests pass with zero failures across `litchi-cfb`, `litchi-xls`,
`litchi-doc`, `litchi-ppt`, `litchi-ole-common` and `litchi-opc`. An independent
adversarial review traced the error-precedence and single-flight paths and
located every test and document that referenced the old discipline; its required
corrections are applied.

Four tests were added or renamed in `crates/litchi-cfb/src/shared.rs`:

- `checked_reads_observe_the_source_version_exactly_once` pins the new discipline
  as a counted invariant — one observation per cursor read, one per range read,
  one for an empty range read, none for a skip.
- `first_cursor_read_after_open_reports_a_change_from_its_trailing_fence` proves
  a mutation landing between open and the first read is still refused, and that
  the failed read does not commit the cursor.
- `read_exact_reports_a_fence_failure_after_a_successful_read` and
  `read_exact_reports_the_payload_error_when_the_fence_also_fails` pin the error
  identities described above.
- `stream_cursor_fences_before_during_and_after_hostile_reads` is renamed to
  `stream_cursor_reports_source_changes_around_hostile_reads`; it asserts
  outcomes, and its old name asserted a mechanism the reader no longer uses.

All ten final quality gates pass: 4,238 tests in the OLE2 crates and 4,325 in
the OOXML crates, with zero failures, plus warning-denied Clippy and rustdoc,
formatting, the all-feature workspace check, minimal-feature checks, crate
boundaries, and the strict claim registry. See
[`quality.md`](results/change-0558/quality.md).

## Limitations

No cold-cache, physical-device, remote/range-source, peak-RSS, allocation,
concurrency-scaling, real-producer, or cross-platform result is claimed. The
measured corpus is one real-producer XLS file plus the generated CFB/OLE2
selectors. The latency figures are host-specific and were captured on a shared
machine without proven quiescence; the deterministic counters are not.

Replay:

```sh
python3 -B docs/performance/results/change-0558/analyze.py
python3 -B docs/performance/results/change-0558/analyze_guardrail.py
```
