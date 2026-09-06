# 0432: XLSX streaming operation memory

The existing XLSX streaming selector exposed its 4 KiB row scratch but omitted
the operation allocation envelope already available for RTF. This tool-only
change adds that envelope and strengthens the untimed oracle to reject extra
stored cells and package members. Production authoring and the corpus generator
are unchanged. Implementation revision: `4c9cfd1fc9dd9df1a0f53a086bce274a8cf687ae`.

The [bundle](../results/change-0432/README.md) records normal and allocator
executables across 64, 8,192 and 131,072 rows, with two reversed repeats,
three warmups and thirty samples per fresh process on CPU 2. All 360 formal
samples pass report/output validation.

| Rows | Output bytes | Normal p50 ms R1 / R2 | Incremental allocator peak |
|---:|---:|---:|---:|
| 64 | 3,451 | 0.1160 / 0.1161 | 420,110 bytes |
| 8,192 | 167,418 | 13.4558 / 12.5297 | 420,110 bytes |
| 131,072 | 2,563,433 | 188.6571 / 188.1209 | 420,110 bytes |

All 180 instrumented samples have that peak above their own entry live count,
zero live-byte change at exit and no failed allocation calls. Cumulative
requested bytes grow from 2,491,428 to 13,234,084. This supports a stable
observed incremental requested-heap peak for the tested scalar writer; it is
not a universal heap, physical-copy or allocator-internal bound. The 4,096-byte
row buffer remains a distinct model. Full-process peak RSS rises from about
80.7 to 279.9 MiB; its scope includes setup and the materializing oracle, so
no causal writer RSS claim follows.

Normal medium p50 repeat drift is 6.88%, with corresponding mean/tail and
throughput flags. Its timing remains descriptive, not an accepted stable
latency baseline. No normal-versus-allocator timing comparison or production
speedup is claimed. Full distributions and uncertainty remain retained.

Validation: 310 release harness library/allocator-target tests pass with one ignored;
31 writer and two focused debug tests pass. Scoped formatting and warning-denied
rustdoc pass. Strict Clippy has the same 29 preexisting findings and no changed
streaming/oracle-region findings; five unchanged files retain formatting debt.
The initial interrupted debug suite, test import error, format finding and
analysis import error remain recorded. Portable replay includes source/driver
custody, result/lint rederivation and report/verifier mutation rejection.

This is fresh one-sheet inline-scalar XLSX creation. Logical append, package
Part addition and repackaging remain separate scenarios. The source audit
recommends bounded ODS scalar-row creation as the next coverage workstream;
its performance impact still needs measurement. Native PPTX cross-copy also
has a concrete optional-name/identity gate and fixture-provenance gap. The
overall non-iWork goal remains active.
