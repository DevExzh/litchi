# 0732 public PPT native phase attribution

[Report](../../0732-ppt-native-phase-attribution.md): opt-in diagnostics only;
ordinary commit is byte-identical. No optimization or speedup is claimed.

The prospective matrix contains 36 native processes: three cycles × three
rounds × four routes, each with 50 samples and three warmups on CPU 12.
Ordinary opaque/split and profiled empty/clock run in the same feature-enabled
binary. The fixed semantic control pairs are compared in every round even
when execution order rotates/reverses. Every output matches the sealed 0728
PPT oracle. All samples, tails, observer flags and individual phase fractions
are retained in `analysis.json`; raw clock traces carry 20 events per owner.

`quality.py` records each attempt separately, including owned source/probe
snapshots. Attempt 0 passed PPT gates, then stopped on two probe string-type
compile errors. Attempt 1 includes those fixes, corrected phase documentation,
and an additional missing-record error parity test, then stopped on a partial
move in untimed probe report construction. Attempt 2 computes replacement
summaries before moving inventories and groups the commit-window arguments
into a tuple; it passed nine probe tests, then Clippy rejected a nested `if`.
Attempt 3 applies that lint fix. The final build receipt
binds the successful attempt, full workspace source census and probe sources.

Root runs all Cargo and native commands serially. The online sequence is
`quality.py`, `run.py qualify`, `run.py freeze`, `preflight.py`, then
`run.py capture`. The preflight duplicates the four qualification schemas into
a temporary full matrix; its data are discarded and never used as timing
evidence. Capture refuses a changed freeze or failed/mismatched preflight.
Scripts refuse to overwrite native capture roots. Reproduction requires an
isolated checkout with the recorded fixture, toolchain, source and probe, and
fresh owned output directories; preserve this sealed packet.

Offline replay from the repository root:

```sh
PYTHONDONTWRITEBYTECODE=1 python3 docs/performance/results/change-0732/source-guard.py
PYTHONDONTWRITEBYTECODE=1 python3 docs/performance/results/change-0732/analyze.py
PYTHONDONTWRITEBYTECODE=1 python3 docs/performance/results/change-0732/audit.py
PYTHONDONTWRITEBYTECODE=1 python3 docs/performance/results/change-0732/negative-checks.py
PYTHONDONTWRITEBYTECODE=1 python3 docs/performance/results/change-0732/artifact-seal.py --check
```

The independent audit recomputes statistics without importing the analyzer.
Corruption checks invoke the actual analyzer on isolated copies. Cleanup
retains binary identity so offline replay works after removal of this batch's
build products. `source-review.json`, the before/after source archives and
`source-guard.py` prove the ordinary method is unchanged. `review.md` records
source findings and their resolution; empirical review is recorded separately.
No iWork files or previously sealed packets are changed.
