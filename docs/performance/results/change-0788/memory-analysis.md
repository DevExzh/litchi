# 0788 memory phase analysis

This is an offline replay of acknowledged `/proc` snapshots. It records
point-in-time smaps classifications and paired deltas for the diagnostic
processes. It does not revise the 0787 decision, pool native timings, or
call a phase value an operation peak.

- Receipts: 64 (32 off, 32 on)
- Report samples: 992
- On-mode snapshots: 432
- Receipts with counter disagreements: 0

## Counter checks

Per-mapping sums are compared with `smaps_rollup`. RSS is required to
match exactly at the retained snapshot; PSS allows the documented
per-mapping kernel rounding tolerance. `status` VmRSS is retained as a
separate asynchronous counter.

No smaps-versus-rollup counter disagreement exceeded the recorded tolerance.

## GNU time versus acknowledged RSS

The diagnostic children retain GNU `/usr/bin/time` `ru_maxrss` separately
from the largest acknowledged `smaps_rollup` RSS and status VmHWM. The
time high-water value was below the retained smaps RSS in 32 of 32 on-mode children and below VmHWM in 32 children.
A high-water counter below a same-child observed RSS is a counter
disagreement requiring investigation. It does not provide causal
attribution or reinterpret the rejected 0787 result.

The complete per-process raw differences and phase deltas are retained
in `high_water_summary` in `memory-analysis.json`.

## Scope limits

Anonymous mappings are classified by their actual mapping labels. The
parser never infers heap or TLS from generic anonymous bytes. The enabled
phase handshake is an observational perturbation, so on/off differences
are retained as controls. Two repeats are shown as observations; no
bootstrap interval is calculated.

The complete per-mapping classification, phase values, operation/drop
deltas, source-metrics checks, output checks and counter disagreements are
in `memory-analysis.json`.
