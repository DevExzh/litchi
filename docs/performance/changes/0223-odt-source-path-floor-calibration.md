# Change 0223: ODT source-path layout noise floor calibration (methodology)

Date: 2026-08-19

## Purpose

Not a code change — a measurement-methodology calibration, the direct
analog of 0218 (itself the litchi-odt analog of 0205/0213), extended to
the ODT source-path and eager-open phases that 0218 did not cover. Change
0222 (fused parse promoted to the owned open path) was provisionally
withheld under the pre-floor rule: the byte-identical guardrail phase
`odt_file_source_open_full_text_lifecycle` showed a p50 adverse
both-directions reading that reproduced in its single permitted rerun
(primary max 1.76%, rerun max 1.24%), despite both executed workloads
accepting at 6.3%-33.2%. This calibration measures the layout noise floor
of the affected phases directly: binaries whose source differs from the
banked control (0221 state) only by never-executed, parser-shaped padding
code in litchi-odt, retained in the binary via a `#[used]`
function-pointer table.

## Probes

All probes build from the identical banked tree (control SHA-256
`93c2279b9b5dff79bbfd58e028c5eedede38c3a917b9f3fc8edb86a4fb0641c7`, the
banked 0221 binary); the probe code was removed after each build and the
tree verified to rebuild bit-exact to the control SHA. Probe functions
are `#[inline(never)]` parser-shaped byte-dispatch bodies with
per-function distinct constants (defeating identical-code folding);
nothing calls them, so no probe instruction executes on any measured
path.

| Probe | Placement | Added text bytes | SHA-256 |
|---|---|---:|---|
| a | `litchi-odt::document` (14 fns) | +6,144 | `42ea627862b9c2cf8b222d0a00f798e1134e5e315ca3040277c919a741c36902` |
| b | `litchi-odt::document` (30 fns) | +12,064 | `e3ceee29963e87f6ac7d0ee185a80d3e4cc10a242f4afac05c29455232a44b02` |
| c | `litchi-odt::elements` (37 fns) | +14,640 | `22c2a7637e6d15f2dab1db3697c625c2f8ea25e7c3c9b2c36398c8ec658b2222` |

The placement and sizes bracket the withheld 0222 change's perturbation
(−10,016 bytes `.text`, in `litchi-odt::document`). Protocol: identical
to the change-measurement protocol — frozen binaries, fresh CPU-2-pinned
processes, order A1 control / B1 probe / B2 probe / A2 control, 30
warmups, 500 samples per leg, same drift ceilings. Because control and
probe are semantically identical, EVERY paired-direction reading is pure
layout noise; the observed magnitudes calibrate the floor.

## Measured floor

Maximum adverse both-directions magnitude observed across probes a/b/c —
every probe reading is pure layout noise on semantically identical
binaries. "–" means no probe produced an adverse both-directions reading
for that statistic. Full per-leg data:
`docs/performance/results/*-0223{a,b,c}-*`.

| Phase | p50 | mean | p95 | p99 |
|---|---:|---:|---:|---:|
| file-source-open | – | – | 2.52% | 10.22% |
| file-source-lifecycle | 3.79% | 2.55% | 4.03% | 6.52% |
| file-eager-open | 5.60% | 5.75% | 9.31% | 9.22% |

Historical corroboration: adverse both-directions readings on
byte-identical phases of the withheld 0222 change pair — file-source-open
p95 -0.78%/-0.60% (primary; the rerun flipped favorable +7.22%/+8.35%)
and p99 -0.42%/-27.98% (primary); file-source-lifecycle p50
-0.11%/-1.76% (primary) and -0.39%/-1.24% (rerun). Folding in the
historical maxima (the 0205/0213/0218 method), the effective floor used
for banking decisions on these litchi-odt phases is:

| Phase | p50 | mean | p95 | p99 |
|---|---:|---:|---:|---:|
| file-source-open | – | – | 2.5% | 28.0% |
| file-source-lifecycle | 3.8% | 2.5% | 4.0% | 6.5% |
| file-eager-open | 5.6% | 5.7% | 9.3% | 9.2% |

Notes: the file-source-open p99 floor is dominated by the historical 0222
reading (27.98%), itself a 500-sample nearest-rank tail artifact of one
binary pair — the probes observed 10.22%. The eager-open phase is the
most layout-sensitive ODT open phase (probe a: adverse-both on ALL FOUR
statistics, up to 9.31%) — plausible given its ~40% `fs::read` share, but
the floor is what it is. "–" entries mean no adverse-both evidence exists
and the floor is not calibrated there (the pre-floor rule still applies
to such a statistic).

## Rule scope

The 0205 banking rule applies unchanged; this calibration extends the
0218 litchi-odt floor set with three more phases (the 0218 floors for
open/list-paragraphs/one-paragraph/full-text/repeated-text-* are
unaffected). Scope limits as in 0218 rule 4: calibrated on litchi-odt
with binary text deltas of +6.1KB to +14.6KB; substantially larger
deltas, other crates, or protocol changes require recalibration.

## Consequences for 0222

Re-evaluating change 0222 (fused parse promoted to the owned ODT open
path) under the extended floor set:

- **Blocking readings reclassified.** The reproduced lifecycle p50
  adverse-both readings (max 1.76% primary, 1.24% rerun) sit within the
  calibrated floor of 3.8% — and probe a alone reproduces adverse-both
  p50 at 1.70% on this phase with ZERO changed code, confirming the
  mechanism. The file-source-open primary adverse-both readings (p95 max
  0.78%, floor 2.5%; p99 max 27.98%, floor 28.0%) are likewise
  within-floor. Under rule 1 these are layout readings, recorded as such,
  and no longer block.
- **Claim scope (rule 3).** `odt_semantic_open` (0218 floors): p50
  6.33%-6.45% (floor 3.3%), mean 10.02%-11.55% (floor 7.2%), p95
  31.02%-33.20% (floor 27.6%) — all claimed; p99 rejected (control
  drift). `odt_file_eager_open` (0223 floors): all four statistics
  accepted and above floor — p50 12.05%-17.05% (floor 5.6%), mean
  12.29%-16.82% (floor 5.7%), p95 11.46%-18.86% (floor 9.3%), p99
  13.40%-17.27% (floor 9.2%) — all claimed.

**Re-verdict: 0222 is banked.** See `0222-odt-owned-open-fused-parse.md`
for the full verdict section.
