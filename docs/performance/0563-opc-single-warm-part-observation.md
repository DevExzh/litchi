# 0563: one fewer source observation per warm OPC part read

Status: retained. `performance_claim: none` — this record carries counted
invariants and syscall counts, and explicitly claims no latency improvement.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

## What was removed

A warm-cache part read through `SourceBackedPackage` took **three** source
observations with no `read_at` between them: one in `part()`, one at the head of
the cache loop in `read_part_with_observer_and_capture`, and one in the
cache-hit branch. On a file-backed source each is a `statx` syscall.

The loop-head observation is removed. It proved a strict subset of what the
hit-branch observation proves: only `enter_with_observer` runs between them — a
mutex, a clock bump and an `Arc` clone, no source read — and the hit-branch
observation is strictly later in time *and* carries a side effect the loop-head
one does not, invalidating the stale cache entry before returning
`SourceChanged`.

Two relocations keep the removed observation where it was load-bearing rather
than deleting it outright:

- **On the cache-entry error path.** `enter_with_observer` can fail on
  cancellation or on a memory, object or work reservation. Change
  [0317](changes/0317-opc-source-read-error-precedence.md) fixes the observable
  precedence as source-version failure, then execution failure, then the mapped
  ZIP error. Without an observation here, a stale source with an exhausted
  budget would report the execution error instead of `SourceChanged`.
- **At the head of the cold-load closure.** This is the opening bracket for
  every branch that consumes bytes — loader, bypass, capture, the batch paths,
  and any loop re-entry. It sits inside the closure so an early `SourceChanged`
  return still unwinds through the flight-completion paths rather than leaking a
  load flight and hanging waiters, and before the catalog lookup so a changed
  source still outranks `PartNotFound`.

`part()` also moved its observation after the retained-catalog lookup, which is
count-neutral. Its missing-Part branch observes before reporting, because
`litchi-pptx`'s cross-copy collision probe and `litchi-docx`'s validation walk
both map `PartNotFound` to a non-error — on names that are absent by
construction, so that branch is the dominant input shape, and discarding the
staleness signal there would discard it on essentially every call.
`crates/litchi-xlsx/src/workbook/source.rs:1007-1011` documents the cross-crate
contract that `part()` observes the version even when the semantic store is
already retained, so the observation could be moved but not deleted.

## Measured effect

### Counted invariants

This is the primary evidence, and it is deterministic. Three tests are added to
`crates/litchi-opc/src/source_backed.rs`, counting observations through the
crate's instrumented `CountingSource`:

| Path | observations before | after | source reads |
| --- | ---: | ---: | ---: |
| warm-cache part read | 3 | **2** | 0, unchanged |
| cold part read | 5 | **5** | unchanged |

`a_warm_part_read_observes_the_source_twice` also asserts the read was a cache
hit, so it cannot silently start measuring a cold read.
`a_cold_part_read_keeps_its_complete_observation_bracket` is the guard that the
cold path stayed bracketed. Both were confirmed to **fail** with the loop-head
observation restored.

`a_changed_source_outranks_a_missing_part` pins the `SourceChanged` over
`PartNotFound` ordering. No test asserted it before this change, although two
downstream crates depend on it.

### Syscalls

One perf-harness child per case, `--warmup 0 --samples 1`, whole-child counts:

| Case | `statx` | `pread64` |
| --- | ---: | ---: |
| `docx_file_source_open` | 624 → 600 (−3.85%) | 224 → 225 |
| `docx_file_source_full_text` | 660 → 636 (−3.64%) | 236 → 236 |
| `pptx_file_source_open` | 9,730 → 9,724 (−0.06%) | 10,216 → 10,216 |
| `pptx_file_source_selected_slide` | 9,758 → 9,752 (−0.06%) | 10,228 → 10,228 |

These paths are **cold-dominated**: the harness opens each package fresh per
sample and each part is read once, so almost no read reaches the warm branch
this change improves. The PPTX figures are near zero for that reason, not
because the removal did not happen. The harness has no file-backed
repeated-part-read selector, so the warm path's end-to-end effect is unmeasured;
that is a coverage gap, recorded here rather than worked around.

### Non-regression check

An A1/H1/H2/A2 matrix over four OOXML file-source selectors with ASLR disabled,
3 warmups and 30 samples per child: 16 children, 4 case/corpus cells, and **zero
cells adverse in both directions by more than 5%**.

| Selector | p50, first direction | p50, second direction |
| --- | ---: | ---: |
| `docx_file_source_full_text` | −4.76% | −4.58% |
| `docx_file_source_open` | −1.43% | −0.95% |
| `opc_file_source_open` | −0.40% | +1.93% |
| `xlsx_file_open` | +2.04% | −1.32% |

The two DOCX cells improve in both directions. That is **reported, not
attributed**: 24 removed `statx` calls at roughly 0.2 microseconds each is about
5 microseconds against a 300-microsecond operation, or 1.6%, so the removal
explains at most a third of the observed movement and the remainder is
unexplained. This matrix exists to show the change does no harm, not to claim a
gain.

### No latency claim

None is made, and none should be expected. Change
[0493](changes/0493-managed-opc-source-read-ahead.md) measured a far larger
version of this — collapsing 19 physical reads to 3 per DOCX lifecycle — and
found +0.98% and −2.01% p50 on a zero-delay local source, concluding in its own
words that this "is not evidence of a useful local-source speedup", against
−84.87% and −82.17% on a 1 ms-service range source. Removing syscalls from a
warm local read is worth approximately nothing; the value of this change is the
work removed on the warm path and the three invariants it makes enforceable.

## Correctness evidence

`litchi-opc` passes 426 library tests plus every integration suite with zero
failures. The `CancelOnHitVersionSource` fixture was ordinal-locked to the old
observation count through a shared `arm_after_cache_enter` helper; that helper
is removed and each of its three call sites now states its own count with the
reason, so a future change to the discipline fails loudly at the call site
instead of silently re-targeting a cancellation.

One accepted behavioural consequence: a cold read against an already-changed
source now creates and fails a load flight, so `cold_loads` and `failed_loads`
each increment where they previously did not. No test pins that; the tests that
combine mutation with load diagnostics all mutate during a gated load.

## Limitations

No cold-cache, physical-device, remote/range-source, peak-RSS, allocation,
concurrency-scaling, real-producer or cross-platform result is claimed. The
warm-path improvement is demonstrated by counted invariants, not by an
end-to-end measurement, because no file-backed repeated-part-read selector
exists.
