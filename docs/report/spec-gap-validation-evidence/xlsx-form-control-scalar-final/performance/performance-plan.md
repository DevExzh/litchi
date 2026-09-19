# Final-source measurement plan

## Scope

The target is the public XLSX form-control scalar lifecycle added by the
worksheet owner and the source-backed editor. The measurement covers the
small dependency closure that a scalar edit is required to preserve:

```text
source ReadAt / owned bytes
    -> worksheet owner projection
    -> typed scalar edit
    -> ctrlProp + VML paired patch
    -> commit/publication
    -> save bytes
    -> reopen and typed readback
```

The package's unrelated members are retained by the writer but are outside
the scalar semantic operation. Their bytes are checked by the correctness
probe and the final output hash; no claim is made that the package is
recompressed at a particular ratio.

This is evidence for the final implementation, not a synthetic microbenchmark
or an owner-read speedup claim. It does not measure Excel acceptance, rendering,
cold page-cache latency, concurrency scaling, or unsupported ActiveX behavior.

## Workloads

The harness emits the following lanes for each selected fixture and repetition.
Setup that would otherwise dominate a lane is performed before the timed
closure; the lane's name states the boundary explicitly.

| Lane | Timed public operation | Source reads counted |
| --- | --- | ---: |
| `eager_read` | `Workbook::sheet(0).form_controls()` on an already-open workbook | no |
| `source_read` | `SourceBackedFormControlEditor::snapshot("Sheet1")` on an already-open source editor | yes |
| `eager_noop_save_reopen` | ordinary `Workbook::edit`, exact scalar set, `commit`, serialization, reopen, typed read | no |
| `source_noop_save_reopen` | source editor exact scalar set, commit, raw-copy publication, reopen, typed read | yes |
| `eager_scalar_save_reopen` | ordinary edit of one scalar, commit, serialization, reopen, typed read | no |
| `source_scalar_save_reopen` | source scalar commit, paired publication, reopen, typed read | yes |
| `source_forward_apply` | `FormControlPatch::apply` to an exact `OpcPackage` source | no |
| `source_inverse_apply` | inverse patch application to the accepted forward package | no |

`eager_read` and `source_read` isolate selector/projection cost. The four
save/reopen lanes are end-to-end lifecycle measurements. The forward/inverse
lanes isolate the reversible source-checked patch boundary and are not added
to save/reopen timings.

The first retained control determines the scalar case from this ordered list:
`Checked`, `LockText`, `NoThreeD`, `JustLastX`, `NoThreeD2`, `Colored`. The
current authored value is the no-op value. A Boolean is inverted for the
changed case; `Checked::Checked`/`Unchecked` is toggled and `Mixed` becomes
`Checked`. If no supported authored scalar is available, the harness refuses
the fixture instead of guessing a field or claiming a scalar edit.

The initial representative matrix is:

| Fixture | Purpose |
| --- | --- |
| `singlecontrol.xlsx` | one-control checkbox; exercises `Checked` and paired mirror publication |
| `tdf120301_xmlSpaceParsing.xlsx` | two-control package; uses the authored `NoThreeD` scalar whose VML mirror is present |
| `tdf134769.xlsx` | checkbox package carrying the lexical `#REF!` formula token; uses the paired `NoThreeD` scalar |

`button-form-control.xlsx` is intentionally excluded from the scalar matrix:
its authored `LockText` has no VML `LockText` occurrence, and the bounded
source editor refuses insertion/removal of a missing mirror. It remains a
valid owner-read fixture, but cannot provide a paired source scalar edit under
the current facade contract.

The exact fixture SHA-256 values and each ZIP member hash are recorded by the
capture script. A final run must use the fixture files from the frozen source
root or a separately hashed immutable fixture root.

## Sampling contract

The default is three warmups and fifteen measured repetitions per lane. Each
repetition is one process-local sample and writes a raw JSON record before the
next sample. A capture may increase the count, but it must record the actual
warmup and sample counts in `run-manifest.json`.

Each raw sample records:

* monotonic elapsed nanoseconds;
* global allocator alloc/dealloc/realloc calls and requested bytes;
* cumulative requested event bytes, end live-byte delta, and peak live-byte delta
  relative to the phase baseline;
* Linux `VmRSS` and `VmHWM` before and after the operation, or explicit `null`
  on a platform without procfs;
* source `ReadAt` call count, requested bytes, returned bytes, and maximum
  request for source-backed lanes;
* fixture label, lane, iteration, selected field, no-op value, changed value,
  and a readback checksum.

Allocator counters are process-global observations. Live bytes are maintained
across phase boundaries by successful allocator events; they are not a heap
implementation proof. RSS/HWM are process observations and are reported as
raw values, not inferred per-object memory use. Percentiles in the generated
`stats.json`/`report.md` must be computed from the retained raw samples, with
p50/p95/p99 and min/max shown for elapsed time and the relevant memory
counters. The runner uses linear interpolation over each fixture/lane's sorted
samples and records that method beside the summaries.

## Final-source gate

Before authoritative capture, the root agent must:

1. freeze the production source and record the exact commit or immutable dirty
   source manifest;
2. finish and retain `Cargo.lock`, `rustc -Vv`, `cargo -V`, active toolchain,
   target triple, optimization flags, allocator, CPU, core count, memory,
   kernel, and storage information;
3. verify every fixture and ZIP member hash against the capture manifest;
4. build the release harness with `--locked --offline` into a temporary target;
5. hash the binary before and after collection and retain both hashes;
6. run the correctness probes before timing, including exact no-op bytes,
   changed paired members, forward/inverse exact restoration, save/reopen
   semantic readback, and unrelated-member preservation;
7. retain raw samples and the generated receipt index, then remove all temporary
   targets, staging projects, and generated packages.

The harness intentionally has no authoritative result files yet. Running it
against the mutable worktree can only be a build or smoke check and must not be
reported as final evidence.
