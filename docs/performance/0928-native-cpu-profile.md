# 0928 — native whole-process CPU profile and attribution limits

The 0927 phase counters locate provider reads, but do not measure CPU work.
This batch records one process-scoped native profile of the retained 0925
release binary. It makes no production change, speedup claim or timing
qualification, and does not replace the failed 0925 timing cohort.

The fixed run uses the generic ReadAt section-layout case, CPU 12, zero
warmups and 200 samples. `perf record` samples `cycles:u` at 997 Hz with
16 KiB DWARF stack capture. The original process handle was polled to exit
zero; there was one document child and no restart. Source/output oracles and
all reported correctness flags pass. The timed arrays remain in the raw
harness report but are not used to compare performance.

## Source and collection

The binary SHA-256 is
`7138c6c458dc13f1677deba3f413120ffc1af155d07848a243a43ebcf335feb7`.
Its source/build receipts are the retained 0925 v4 receipts. Independent
comparison of that source listing against the current checkout finds 6,275
of 6,276 files identical; the sole difference is the 0927 test-module
addition in the harness. The non-test harness prefix is byte-identical.
The release binary has symbols and unwind tables, but no debug-line section.

The capture records 3,267 samples, zero reported lost samples, and an
approximate event count of 12,961,608,407. Frequency-mode samples have varying
period weights, so raw sample fractions must not be substituted for weighted
cycle fractions. Parent-inclusive call-tree totals must not be added together.

## What the profile can establish

The leaf-only export identifies these sampled **self** costs, summed by
symbol name across recorded commands and rounded to two decimals:

| Sampled symbol | Whole-process self share |
| --- | ---: |
| `sha2::sha256::x86_sha::compress` | 41.50% |
| `litchi_perf_baseline::payload_bytes` | 20.03% |
| `__memmove_avx512_unaligned_erms` | 5.27% |
| `__memcmp_evex_movbe` | 4.10% |
| Unresolved sampled instruction pointer | 3.81% |
| `quick_xml::reader::ns_reader::NsReader<R>::process_event` | 2.52% |

The denominator is all 12,961,608,407 recorded period units. Perf inherits
metadata subprocesses: 3,254 records belong to the harness, 12 to `rustc` and
one to `uname`. The latter two contribute 39,540,346 period units. There was
one document-workload launch, not one operating-system process in the profile.
The standard report separates command rows; for example its harness-only
`memcmp` row is 4.05%, while the merged symbol total above is 4.10%.

The sampled SHA routine contains `sha256rnds2` instructions in the exact
binary. This is native hardware-assisted compression evidence, unlike the
older Valgrind guest-instruction profile. It does not identify which hash
invocation produced each sample. Source fingerprints, emitted-artifact
hashing, fixture creation and verification all occur in this process.

The large `payload_bytes` share is a warning about the observation boundary:
the process includes deterministic corpus construction and untimed
preservation checks. It is not evidence that document editing spends that
share generating payloads. The profile cannot be converted to a query or
publication speedup bound by applying whole-process percentages to the
benchmark's timed intervals.

## Unwind and analysis limitations

Successful collection and zero lost samples do not establish correct call
chains. The DWARF reader emits `address range overlaps an existing module`
errors, and many stacks contain unresolved or invalid-looking parent
addresses. Phase attribution needs independently verified caller addresses
and usable ancestors; a leaf symbol alone is insufficient.

The period-bearing stack export contains 1,937 samples with no printed frames,
covering 59.48% of the total period weight. Only three samples retain a frame
for the named harness routine. Summing just the printed stacks would discard
most sampled weight and bias the result. Leaf-cost analysis must use sampled
instruction pointers independently of this incomplete stack export.

The first offline `perf script --show-ip` command failed because the installed
perf does not support that option. A supported field list decoded the same
`perf.data` successfully, including sample periods and symbol offsets. The
failed reader attempt and verbose unwind diagnostics remain retained. There
was no recapture or discarded document child.

## Next experiment

The static binary map identifies the timed query call at `0x1057b79` and
publication call at `0x1058059`. The three retained harness frames instead
point to post-publication output verification. They do not supply a usable
timed-phase denominator. The map is retained with the binary inspection;
addresses apply only to this exact binary.

Use a separate diagnostic build with `-C force-frame-pointers=yes` and
frame-pointer sampling to avoid dependence on this failing DWARF decoder.
If phase identification still needs distinct symbols, add benchmark-only
non-inlined wrappers that execute the query and publication work in the sample
loop; empty markers placed before and after work would not establish ancestry.
Keep preflight and verification outside those wrappers. Inspect the resulting
assembly and prove callchain coverage before using phase shares. Such a build
is profiling evidence, not a drop-in timing comparison against this binary.
No provider cache, SIMD rewrite or production optimization is adopted here.

## Evidence and reproduction

`results/change-0928/receipts/perf-record-diagnostic-v1.json` binds the exact
command, original terminal status, source/build identities and artifact
sizes/hashes. Large artifacts remain outside the checkout:

- `perf.data`: 54,767,504 bytes, SHA-256
  `af3121cf9499534ee0d0581e5e9bf0e8a9d5f02855d4936b8738912c27796ff6`.
- `perf-script-symbol-ip.stderr`: 1,214,687 bytes, SHA-256
  `be421ec8e1f2da34d7e36571294154ff023c69efc5fa894c3536c3e795036364`.

Both are under `/tmp/litchi-performance-0928/`; their regenerating commands
are recorded in the receipt. The small binary-inspection packet confirms the
SHA instructions and that this perf build uses libdw for DWARF unwinding.
Small raw reports and the period-bearing stack export are also copied into
`results/change-0928/raw/` for portable inspection.

The committed leaf-only export avoids the missing-frame bias. Reproduce the
exclusive descriptive symbol-family grouping without running a document:

```sh
python3 -B docs/performance/results/change-0928/classify_leaf_period_v1.py \
  docs/performance/results/change-0928/raw/perf-script-leaf-period.txt \
  --output /tmp/0928-leaf-replay.json
```

The family labels are explicit symbol-name rules, not phase labels or proofs
of high-level operation cost. The retained classification and root replay
agree on every bucket, all 3,267 samples and all period units. The independent
source/report review and root stack-quality counts are retained under
`results/change-0928/quality/`. No runtime code changed, so Cargo test suites
were not rerun for this evidence-only batch.
