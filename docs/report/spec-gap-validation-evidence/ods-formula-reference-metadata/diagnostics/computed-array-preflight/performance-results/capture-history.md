# Reference-metadata performance capture history

The profile has one authorized frozen-source capture and no diagnostic retry. The raw baseline and candidate receipts are retained in `baseline-049c09cdde3978593149079c4257df047a3fa419/` and `candidate-final/`.

| session | disposition | baseline rows | candidate rows |
| --- | --- | ---: | ---: |
| `2878` | complete authorized frozen-source pair | 840 | 3,810 |

The built-in verifier passed the exact case matrix, 15 samples per phase/group, typed preflight and read bounds, source/profile manifests, allocator balance, raw receipts, and cleanup receipts. The independent interpretation is in `capture-analysis.md` and `capture-analysis.json`; `performance-report.md` and `.json` contain the p50 tables.

The candidate freeze file hash is `e9d8dbac4eb9d9fe964693ccb801e74a031f7c223b3595e5003de03e36f02288`. Host load, CPU affinity, and the process snapshot are retained in `capture-context-before.json`; the host was not isolated for timing.
