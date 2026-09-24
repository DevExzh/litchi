# Rejected selected-worksheet scanner V6

The historical V6 candidate regressed normal median latency by 37.37%, allocation calls by 119.41%, and allocated bytes by 8.51%. Its production changes were rejected. These results compare the historical `cf4f383c` source and do not include the separately accepted plain-cell-tag optimization.

Extract `evidence.tar.zst` while preserving recorded regular-file permissions, then run `python3 tools/verify_rejected_v6.py` inside the extracted directory. The standalone verifier recomputes raw results and validates source and receipt bindings without running a workload.

The archive retains reports, source patch, review and build evidence, depfiles, and exact binary hashes. Four duplicate executable byte streams are omitted, so it does not support a self-contained binary rerun. Original binaries and the full archive are retained outside this publication. Allocator elapsed time is instrumentation-only. Historical pre-capture plans and the superseded diagnostic V1 runner are preserved as evidence, not as current execution instructions.
