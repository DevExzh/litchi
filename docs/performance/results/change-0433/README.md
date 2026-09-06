# 0433 bounded ODS scalar creation

This bundle measures the new sequential ODS scalar-row writer against the
existing buffered builder. It is a fresh one-sheet creation experiment and a
bounded-memory capability enabler. It does not add or measure append/edit
semantics for an existing document. The derived summary has `claims: []`.

## Capture and provenance

The frozen protocol uses rows of 64, 8,192, and 32,768, four scalar cells per
row, normal and allocator modes, two repeats, three warmups, and 30 retained
samples per report on CPU 2 with one worker. The three roles are:

- `before-buffered`: existing builder at `be82ef53dc191a64bc3c3365502bf44599528577`;
- `after-buffered`: the existing builder at the candidate revision, which
  controls for revision and harness drift;
- `after-streaming`: the new writer at
  `f5bf1696192f0007db56554aa5b16719cdf1950b`, with a fixed 4,096-byte row
  authoring window.

There are 36 formal reports (three roles × three shapes × two modes × two
repeats) and 1,080 retained samples. The report verifier and derived
[`summary.json`](summary.json) both report PASS. Build and capture custody is
retained in [`before/build.json`](before/build.json),
[`after/build.json`](after/build.json), and the before/after capture indexes.

The timed operation creates deterministic scalar rows, authors content,
validates publication, compresses and finalizes the package, and writes to a
hashing discard sink. Artifact creation and reopen, procfs probes, sink
construction, and digest extraction are outside the timer. Normal and
allocator latency are separate; the allocator policy reports operation-scoped
request vectors and does not authorize a normal-versus-allocator timing
comparison.

## Normal latency

The table gives p50 latency in milliseconds as `before-buffered → candidate`,
with the percentage from the summary in parentheses. `after-buffered` is the
same existing builder at the candidate revision; `after-streaming` is the new
writer.

| Rows | Before p50 R1 / R2 | After-buffered p50 R1 / R2 (delta) | After-streaming p50 R1 / R2 (delta) |
| ---: | ---: | ---: | ---: |
| 64 | 0.459567 / 0.458422 | 0.463737 / 0.466202 (+0.907% / +1.697%) | 0.185696 / 0.187876 (-59.593% / -59.017%) |
| 8,192 | 52.685611 / 52.547425 | 52.925857 / 52.972396 (+0.456% / +0.809%) | 20.807350 / 20.542384 (-60.507% / -60.907%) |
| 32,768 | 217.310707 / 215.050883 | 215.548614 / 216.286653 (-0.811% / +0.575%) | 82.279385 / 82.448135 (-62.137% / -61.661%) |

For the streaming role, p95 deltas R1/R2 are −58.851%/−58.264% (64 rows),
−60.425%/−61.133% (8,192), and −61.297%/−61.548% (32,768). The corresponding
p99 deltas are −58.856%/−58.260%, −60.389%/−61.125%, and
−61.115%/−61.396%.

All repeat flags are clear. The only comparison regression flag is the
after-buffered control at 64 rows, R1, normal p99: +7.037%. It is retained as
a control observation; no after-streaming comparison is flagged.

## Allocator request observations

These are p50 operation-scoped vectors. `regional peak−entry` is the derived
serialized-region peak above the operation entry value. R1 and R2 values are
identical for each role and shape in the retained summary. The middle role
matches the before-buffered allocator values; the table shows the before
buffered control against after-streaming.

| Rows | Requested bytes p50 | Requested calls p50 | Regional peak−entry p50 |
| ---: | ---: | ---: | ---: |
| 64 | 1,366,338 → 882,097 (−35.441%) | 3,583 → 1,059 (−70.444%) | 461,977 → 419,347 (−9.228%) |
| 8,192 | 66,578,958 → 6,230,321 (−90.642%) | 426,260 → 122,979 (−71.149%) | 17,758,827 → 419,347 (−97.639%) |
| 32,768 | 263,900,025 → 22,401,329 (−91.511%) | 1,704,218 → 491,619 (−71.153%) | 71,050,076 → 419,347 (−99.410%) |

The byte and call values describe allocator requests, and the regional value
describes the instrumented region's derived peak. They do not measure physical
copy volume, allocator-internal overhead, total process RSS, or all retained
objects. The large reductions in these dimensions do not support a 10×
end-to-end latency, throughput, or memory claim.

## Semantic and memory boundaries

Each role has a role-local deterministic archive hash. The buffered and
streaming writers may produce different XML/package bytes, so those hashes
are required to be stable within a role rather than equal across roles. The
cross-role oracle matches the semantic row/cell catalog, scalar columns,
`Sheet1`, and semantic hash. The Rust oracle also checks the generated package
structure and reopen result. Generated archive bytes are not retained in this
Python bundle; Python verification binds the recorded hashes and Rust oracle
flags but does not independently parse the output archives.

The streaming role's fixed 4 KiB window is the reusable row authoring buffer.
It is not a total heap or RSS bound: ZIP framing, compression, XML auditing,
metadata, parser/oracle work, process setup, and allocator state have separate
lifetimes and observations. No causal RSS or general bounded-memory claim is
made from the window alone.

The corpus is generated fresh scalar ODS content. No LibreOffice, Microsoft
Office, or other external native application is part of this capture, and no
existing source package is modified.

## Profile status and remaining scope

All six profile jobs and their portable report verification pass. The separate
large-corpus normal processes include setup, corpus construction, semantic
reopen, warmups, 30 operations, and report/binary hashing. They are not the
operation timer or allocator region.

| Role | Whole-process cycles | Whole-process instructions | Highest sampled self symbol |
| --- | ---: | ---: | --- |
| Before buffered | 36,846,570,842 | 147,377,295,571 | `__memmove_avx512_unaligned_erms`, 12.84% |
| After buffered | 36,566,751,601 | 147,370,772,071 | `__memmove_avx512_unaligned_erms`, 11.72% |
| After streaming | 16,223,390,023 | 60,094,311,825 | `ExecutionContext::consume`, 24.41% |

The raw `stat/perf-stat.csv` and `record/perf-report.txt` files under each
role's `before/profiles/` or `after/profiles/` directory retain the full
counters and stacks. The streaming profile also samples the untimed semantic
hash oracle at 8.99%, XML auditing at 7.80%, compression at 7.07%, and fragment
shape validation at 6.38%. These proportions motivate further scoped profiling;
they do not attribute operation latency or copied bytes. All record reports
show zero lost samples. The reported L1 load-miss counter is zero for every
role; it is not accepted as proof of zero cache misses. LLC counters were
unsupported in the retained probe. No operation IPC, physical-copy, RSS,
scaling, or hardware-limit claim is derived from these process profiles.

The global non-iWork goal remains open. This bundle establishes a measured
fresh-creation path with explicit semantic and allocator boundaries; it does
not establish existing-document append, native-application compatibility,
total-RSS reduction, physical-copy reduction, or an overall speedup.

## Portable replay and cleanup

Copied-bundle replay and mutation checks passed before and after removing the
five task temporary directories, including all four captured binary copies
(1.81 GB). Repository build targets and the user goal file were preserved.
See [`precleanup-portable`](checks/precleanup-portable.json),
[`task-cleanup`](checks/task-cleanup.json),
[`cleanup-inventory`](checks/cleanup-inventory.json), and
[`aftercleanup-portable`](checks/aftercleanup-portable.json).

From any copied bundle directory, verify the final retained artifacts with:

```sh
python3 -B verify.py --portable-check --require-inventory
```

This does not require the original checkout or removed binaries. The verifier
checks the complete `SHA256SUMS`, lossless compression, raw report vectors,
source/binary identities, command receipts, strict-diagnostic comparisons, and
profile artifacts. Mutation checks reject corrupted elapsed/output/semantic
values, allocator vectors, row-window declarations, role bindings, and verifier
source. Failed development commands and replay attempts remain retained.
