# 0924 retained Heaptrack stack attribution

Status: descriptive read-only diagnosis. No benchmark, build, recapture, restart,
qualification, or production adoption was performed.

The audit consumed the 0923 independent memory receipt
`docs/performance/results/change-0923/independent-memory-disposition.json` and
its 25 exact raw bindings: one retained v3 first process plus the 24-process
v4 rerun. Every raw path's recorded size and SHA-256 matched before decoding.
The v3/v4 qualification remains invalid as recorded by the receipt; these
profiles are used only as descriptive evidence.

## Method and custody

The external decoder is
`/tmp/litchi-performance-0924/stack-attribution-v2.py` (28,606 bytes,
SHA-256 `e4bef0c1fa0cc835f808d2afa77e002a25bbb7b1df40f4ed513f436f02d49294`).
It validates Heaptrack format `10500 3`, nonnegative table references, terminal
`strings`/`ips` summaries, allocation-index liveness and deallocation
underflow. Each `+` event contributes its allocation-info size exactly once to
one complete leaf-to-root stack. Parent stack rows are never added to child
rows.

The full exact stack rows for all 25 profiles are external:

- `stack-attribution-v2.json`: 135,378 bytes, SHA-256
  `fae07a13e700d00b1bf03c296daf8a3fd61cfe844e5ae89f93a3d0e226744661`
- `stack-rows-v2.json`: 150,263,792 bytes, SHA-256
  `7db5a787e43d1b7c700c93d609d66d298f4c5cbd09884c58fca1e001c1a6340b`

The summary binds the full-row descriptor and the exact regeneration argv. The
prior 0923 parser-integrity receipt also passed all 10 malformed/negative,
range, truncation, underflow and double-free controls; it is parser-only evidence
and makes no operation claim.

Source custody records 9,355 files at revision
`4e2532a0ce958f5928496b0ad9b356d8636916e1`; the candidate source closure list
also contains 9,355 entries. The retained candidate archive is 170,065,920
bytes with SHA-256
`a6880fb4f638fbdcff23afb728382ea0c807f3af201648f579844309492b10a9`; its
source listing is 1,437,118 bytes with SHA-256
`5eb6840368d0e601e2db4756ae2dbc4c7fdbc4d0b8bfd810934652a95090a83a`.
The baseline receipt records the same archive size/SHA, but its historical
`base-head.tar` path is absent after external cleanup; that absence is retained
in the JSON rather than treated as a recovered file. Baseline listing custody is
recorded as 1,156,468 bytes, SHA-256
`f9714cda62133005e6eb0046f79189bd8914946ae3f1a16baade631a8b4f12fc`.
Both frozen binaries were present and matched their recorded size/SHA.

## All retained profile totals

| Route / variant | Profiles | Allocation events | Requested bytes | `read_limited` bytes |
| --- | ---: | ---: | ---: | ---: |
| Generic ReadAt / baseline | 5 | 1,563,585 | 759,020,180 | 0 |
| Generic ReadAt / candidate | 4 | 1,250,868 | 607,216,228 | 0 |
| Owned vector / baseline | 4 | 1,250,820 | 607,212,080 | 0 |
| Owned vector / candidate | 4 | 1,250,824 | 606,950,356 | 0 |
| Owned reader / baseline | 4 | 1,251,856 | 1,693,554,156 | 1,086,341,836 |
| Owned reader / candidate | 4 | 1,251,860 | 1,693,292,432 | 1,086,341,836 |

The per-profile retained totals match the 0923 memory receipt. The owned-reader
and owned-vector raw reports use the same 2,101,427-byte generated source; the
reader's extra volume is cumulative allocation requests, not physical copied
bytes or resident heap.

## Disjoint gap accounting

The four baseline profiles and four candidate profiles give the same gap:

| Variant group | Reader total | Vector total | Difference | `read_limited` | Common-stack delta | Net route-equivalent setup |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Baseline, 4 + 4 profiles | 1,693,554,156 | 607,212,080 | 1,086,342,076 | 1,086,341,836 | 240 | 0 |
| Candidate, 4 + 4 profiles | 1,693,292,432 | 606,950,356 | 1,086,342,076 | 1,086,341,836 | 240 | 0 |

Per process, the difference is **271,585,519 bytes**. The primary stack is the
same in every owned-reader profile and contributes **271,585,459 bytes from
257 allocation events**:

```text
RawVecInner::finish_grow
  <- litchi_opc::phys_pkg::read_limited
  <- litchi_docx::source_backed::Package::from_reader
  <- litchi_perf_baseline::run_docx_source_backed_existing_section_layout_edit_save
```

The route-only source setup rows other than that primary stack total 285,317
bytes per reader process and 285,317 bytes per vector process. They have
different route wrapper names but equal gross totals, so their net contribution
is zero and they are reported once as an equality check. They are not summed
with the primary stack.

The remaining 60 bytes per process are common control/setup rows: 48 bytes in
`box_new_uninit`, 3 bytes in an environment lookup, and 9 bytes across three
`args_os` rows. Their aggregate is 240 bytes per four-profile group. The full
rows, including allocation counts and exact optimized symbols, are in the
external stack file.

## Interpretation boundary

The source confirms that `from_reader_with_limits` calls `read_limited`, whose
loop performs `try_reserve_exact(read)` after each read
(`crates/litchi-opc/src/package.rs:1045-1047` and
`crates/litchi-opc/src/phys_pkg.rs:1789-1834` in the frozen source). The harness
invokes the owned-reader constructor in its timed `open` phase
(`tools/perf-baseline/src/lib.rs:50573-50598`). The retained Heaptrack lane is
still whole-process: expected-package construction, route preparation,
correctness/reopen controls and output checks are in the same process. No trace
contains the 0919 `measure_one` boundary. Therefore the stack attribution
identifies the reader-ingress growth path and its source ownership, but it does
not promote these bytes to an operation-scoped allocation total.

This is consistent with the earlier 0779 finding that repeated exact capacity
requests can be large while physical copying and resident memory remain
unmeasured by requested-byte sums. It is a concrete hypothesis for future
bounded reader-ingress work, subject to a new protocol covering input limits,
short reads, retries, allocation failure and memory retention. It supplies no
new latency or adoption claim.
