# 0733 PPT finish instruction partition

[Report](../../0733-ppt-finish-instruction-partition.md). This packet replays
sealed 0731 profiles and qualifies the current source relationship through the
0732 source census and ordinary-method archive. It does not modify either
ancestor packet or production Rust.

`analyze.py` verifies both complete ancestor seals, all 34 constraints, the
7,206-file current source census, the two diagnostic-only differences, all
retained function/edge costs, and seven uniquely attributable nested nodes.
`analysis.json` retains every direct edge of those nodes, with explicit
parent denominators. Shared descendant costs are never globally assigned.

The standalone `collection-witness.rs` demonstrates why Callgrind `calls=`
metadata cannot be treated as the number of invocations inside the collection
window. `run-witness.py` records its plan before execution, checks formatting,
compiles with warnings denied, runs one native output control and three fixed
Callgrind witnesses. It refuses existing owned output directories. The witness
is not a PPT latency benchmark. Tool versions and all raw results are retained.

Offline replay from the repository root:

```sh
PYTHONDONTWRITEBYTECODE=1 python3 docs/performance/results/change-0733/analyze.py
PYTHONDONTWRITEBYTECODE=1 python3 docs/performance/results/change-0733/checks.py
PYTHONDONTWRITEBYTECODE=1 python3 docs/performance/results/change-0733/artifact-seal.py --check
```

`source-review.md` maps the work and the ownership candidate. `review.md`
provides independent review. Cleanup preserves exact witness binary identity;
terminal receipts record successful replay after removal. The packet seal
covers every local artifact; the ancestry receipt binds the full external
profiles and native context without duplicating or rewriting them.
