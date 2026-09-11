# 0510 helper guardrail: XML 1.0 CRLF length accounting

This is an isolated helper benchmark for the ODT `normalized_xml10_decoded_len`
hypothesis. It does not build or execute Cargo targets and it does not measure
the complete ODT export path. The probe compares the current validated byte
loop with two iterator forms and the bounded first-CR hybrid:

1. `scalar`: UTF-8 validation followed by the current byte loop.
2. `memchr_iter`: UTF-8 validation followed by `memchr_iter('\r')` and a
   next-byte LF check.
3. `memmem_iter`: UTF-8 validation followed by `memmem::find_iter("\r\n")`.
4. `first_cr_then_scalar`: UTF-8 validation, one `memchr('\r')`, immediate
   return on no CR, otherwise the current loop beginning at the first CR.

Each variant is checked for the same decoded length before timing. Inputs are
49, 1,024, and 65,536 bytes with ASCII/no-CR, sparse CRLF, dense CRLF, all CR,
and UTF-8 text patterns. The process is single-threaded and pinned to CPU 31;
the forward and reversed variant orders each retain 50 samples. Iterations per
sample are 4,096, 512, and 16 respectively. The complete raw matrix is in
[`raw-results.csv`](raw-results.csv), and the exact source is in
[`helper_guardrail.rs`](helper_guardrail.rs).

## Result

The CR iterators are unsuitable as an unconditional replacement. Across the
two order medians, their worst p50 regressions versus the scalar helper are:

| Variant | Worst p50 case | Delta | Worst p95 case | Delta |
| --- | --- | ---: | --- | ---: |
| `memchr_iter` | 65,536-byte all-CR | +672.3% | 65,536-byte all-CR | +667.3% |
| `memmem_iter` | 65,536-byte dense CRLF | +754.3% | 1,024-byte dense CRLF | +726.7% |
| `first_cr_then_scalar` | 49-byte dense CRLF | +33.3% | 49-byte dense CRLF | +46.7% |

The first-CR hybrid is the only candidate without a material long-input
CR-bearing regression in this probe. Its 65,536-byte p50 is −0.0% for all CR,
−1.2% for dense CRLF, and +8.4% for sparse CRLF. On no-CR inputs it is
−73.5% at 49 bytes, −95.3% at 1,024 bytes, and −95.8% at 65,536 bytes. The
same fast path is also favorable for the Unicode/no-CR pattern (about −43.8%
at 65,536 bytes). The 8.4% sparse-CRLF p50 increase is the remaining guardrail
to carry into the end-to-end profile.

The simplest production candidate is therefore a no-CR fast path: retain the
existing UTF-8 validation, use `memchr` only to detect the first CR, return the
raw length when none exists, and preserve the existing scalar CRLF walk when a
CR exists. Do not use `memchr_iter` or `memmem::find_iter` over every CR/CRLF
without a separate workload policy. This recommendation is diagnostic only;
admission still requires unchanged semantic output/error tests and the matched
0509/0510 end-to-end ODT profile.

## Custody

The probe was compiled with Rust 1.95.0 and the active target's
`libmemchr-291e0ba635faa0f2.rlib` (SHA-256
`73a74d96e4c307bec5f78cb6db20113198958de4c96107fb8d94b6872bb03815`). No
Cargo command was run for this helper. The run took 2.45 seconds with a
1,940 KiB maximum resident set. Commands and hashes are retained in
[`commands.txt`](commands.txt).

Source SHA-256: `3b13366fddf7c6842ed64cf0304c111bdded74e578ffdef39f2f9e5aeba09733`.
Raw-result SHA-256: `8d4047f1ce92ce7402b0b9b87ec620b4619e036bf509bec71dc1b7fd4f21c15e`.

## Evidence limit

This is a diagnostic probe run during compilation, not a formal native export
comparison. `raw-results.csv` retains per-row min/p50/p95/max aggregates, not
the fifty individual sample durations; those percentiles cannot be recomputed
from that CSV alone. Per-call durations are integer divisions of batched
elapsed time. The 49-byte sparse fixture contains no CRLF at its generator's
spacing and duplicates its no-CR control. The separate public export guardrail
uses an actual mid-string CRLF at that size and retains every sample. The
accepted candidate must be judged by that guardrail and the unchanged baseline,
not by these helper aggregates or a cross-probe percentage comparison.
