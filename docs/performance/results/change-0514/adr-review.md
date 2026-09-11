# 0514 ADR disposition

All 29 accepted ADRs and `docs/adr/README.md` were previously read and are
unchanged under the fresh 30-file [manifest](adr-manifest.json), bound to
control revision `33a21e0f087b4ead80ca63187c7cf30d0584b1f6`.

The tested prototype changes ten private XLSX production paths and four test
paths. It is rejected and restored; the retained deliverable is its exact
patch, measurements, reviews and an independent public-API guard. No accepted
ADR is amended, excepted or superseded.

| ADRs | Disposition |
| --- | --- |
| 0001, 0005, 0008 | Correctness and bounded practical resource behavior precede instruction savings. The prototype passes the tested owner suite but fails cold no-op latency and allocation admission. Rejection is retained explicitly; no accepted speedup, readiness or full-goal claim follows. |
| 0003, 0006 | The shared reader preserves semantic/materialization/style/projection errors before deferred snapshot errors, discards snapshot errors for ineffective edits, validates before publication, and preserves exact no-op bytes and reversible patches. The 17 new tests cover these paths and cache publication. |
| 0002, 0024 | No production crate, dependency edge or public API is added. The standalone probe is a separate workspace using the existing XLSX API and canonical allocator observer. |
| 0004, 0010, 0011 | Worksheet parsing and layout grammar stay inside XLSX; OPC still owns physical packages and sequential publication. The layout rewrite seam is restricted to the worksheet owner and its wrapper checks exact original source identity. |
| 0017, 0018 | Producer-template and calculation-chain boundaries remain unchanged. Changed output validation, compaction, style checks and the existing 4,096-cell/1 MiB Store handoff are not removed. |
| 0007, 0012–0016, 0019–0022, 0025 | Their semantic/format ownership is outside this private worksheet change. No policy or implementation change is retained in those owners. |
| 0026, 0027 | OLE directory metadata and XLS sheet-anchor ownership are unchanged. Completed 0511 CFB work remains part of the OLE2 evidence, not a reason to weaken its residual validation invariants. |
| 0009, 0023 | ODF ownership is unchanged; optimization remains deferred until the full OLE2/OOXML goal is complete. |
| 0028, 0029 | iWork/IWA is excluded from this workstream and untouched. |

The [design review](design-review.md), [implementation review](implementation-review.md)
and [admission review](admission-review.md) separate structural correctness,
tested behavior and empirical retention. No unsafe code, hidden executor,
ambient production I/O, persistent Layout cache or weakened limit is introduced.
The probe reuses the existing isolated counting-allocator wrapper rather than
adding an unsafe production implementation.

Memory reports distinguish absolute region peaks, entry live allocations and
per-sample incremental live demand. Instrumented timing/RSS and Callgrind
instruction references are not native speedup evidence. The full candidate
native protocol is deliberately not admitted after the pilot failure. Exact
base restoration ensures the rejected speculation cannot change user behavior.
