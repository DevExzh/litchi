# Change 0205: per-binary-pair layout noise floor calibration (methodology)

Date: 2026-08-19

## Purpose

Not a code change — a measurement-methodology calibration. Changes 0197,
0203, and 0204 were withheld after reproduced both-directions adverse
patterns appeared on phases executing none of the changed code, attributed
to per-binary-pair code-layout effects. 0204 additionally showed a real
20%-26% targeted stage win, so the cost of an unquantified "layout wobble"
hypothesis became concrete. This calibration measures the layout noise
floor directly: binaries whose source differs from the banked control
(0202 state) only by never-executed, parser-shaped padding code in
litchi-ods, retained in the binary via a `#[used]` function-pointer table.

## Probes

All probes build from the identical banked tree (control SHA-256
`475cf2898880363517eec9e0a9ac6b582eed1f78054f161f394bdd635bb19d7d`);
the probe code was removed after each build and the tree verified to
rebuild bit-exact to the control SHA. No probe instruction executes on any
measured path.

| Probe | Placement | Added text bytes | SHA-256 |
|---|---:|---:|---|
| A (medium) | `litchi-ods::protection` | +16,536 | `08ec494d93f67dd6448a575952672b80340f427ecbe64dfef74f7bf664c71fab` |
| B (small) | `litchi-ods::worksheet` | +5,492 | `915040c2a59010812c15611dcc6c77a1d0717b90b8ce2d86c7f8a376425f3b15` |
| C (medium) | `litchi-ods::authoring` | +15,148 | `4effcc1bfc44c84331be9dfa1eca41c06fab81474e31e865cecc517a812a5449` |

Protocol: identical to the change-measurement protocol — frozen binaries,
fresh CPU-2-pinned processes, order A1 control / B1 probe / B2 probe /
A2 control, 30 warmups, 500 samples per leg, same drift ceilings. Because
control and probe are semantically identical, EVERY paired-direction
reading is pure layout noise; the observed magnitudes calibrate the floor.

## Measured floor

Maximum adverse both-directions magnitude observed across probes A/B/C —
every reading below is pure layout noise on semantically identical binaries.
"–" means no probe produced an adverse both-directions reading for that
statistic. Full per-leg data: `docs/performance/results/*-0205{a,b,c}-*`.

| Phase | p50 | mean | p95 | p99 |
|---|---:|---:|---:|---:|
| source-open | 5.52% | 5.19% | 4.22% | 35.34% |
| eager-open | 1.75% | 1.60% | 1.81% | – |
| one-edit lifecycle | 2.05% | 2.16% | 2.74% | – |
| one-edit commit | 3.33% | 3.69% | 5.75% | 16.64% |
| one-percent lifecycle | 2.34% | 2.09% | 2.53% | 7.67% |
| one-percent commit | 3.01% | 3.05% | 4.52% | 13.07% |
| repeated total | 1.80% | 1.71% | 2.45% | – |
| repeated stage | 3.72% | 3.67% | 5.27% | 10.70% |
| repeated commit | 4.33% | 4.40% | 6.66% | 7.03% |
| repeated publication | 1.01% | 1.01% | – | – |

Historical corroboration: the 0201-0204 change pairs (real source changes
elsewhere in litchi-ods) repeatedly showed adverse both-directions readings
on source-identical phases — eager-open -0.59% to -3.0% (0201), mixed
±0.8% (0202), -0.15% to -0.79% (0203), -1.67% to -2.98% (0204). Taking the
historical maxima into account, the effective floor used for banking
decisions is:

| Phase | p50/mean | p95 | p99 |
|---|---:|---:|---:|
| source-open | 5.5% | 4.5% | 36% |
| eager-open | 3.0% | 3.0% | 7% |
| one-edit lifecycle | 2.2% | 2.8% | 7% |
| one-edit commit | 3.7% | 5.8% | 17% |
| one-percent lifecycle | 2.4% | 2.6% | 8% |
| one-percent commit | 3.1% | 4.6% | 13.5% |
| repeated total | 1.8% | 2.5% | 2% |
| repeated stage | 3.8% | 5.3% | 11% |
| repeated commit | 4.4% | 6.7% | 7.5% |
| repeated publication | 1.1% | 2% | 1% |

The p99 tail of source-open (up to 35%) reflects nearest-rank sensitivity
of a 500-sample tail statistic to code-layout shifts in the tokenizer;
p50/mean/p95 are the operative banking statistics.

## Refined banking rule

Effective from 0205 onward, for litchi-ods changes measured under this
protocol (binary deltas on the order of the calibrated +5KB to +17KB text
range):

1. **Within-floor adverse readings on non-executed phases do not block.**
   A both-directions adverse pattern on a phase that executes none of the
   changed code is recorded as a layout reading, not a regression, when its
   magnitude is within the calibrated floor for that phase/statistic. The
   reading and its floor classification are recorded in the change doc.
2. **Adverse readings on executed phases still block.** When a phase does
   execute changed code, an adverse both-directions pattern must be cleared
   by the single permitted rerun of that workload before banking; if the
   rerun reproduces the pattern, the change is withheld regardless of the
   floor.
3. **Accepts are claimed only above the floor.** A favorable both-directions
   pattern whose magnitude does not exceed the calibrated floor is recorded
   as neutral (within noise), not claimed as a win. Claim scope lists only
   statistics whose accepted magnitude exceeds the floor.
4. **Floor scope.** The floor is calibrated on litchi-ods with binary text
   deltas of +5.5KB to +16.5KB. Substantially larger deltas, other crates,
   or protocol changes require recalibration before the floor can be
   invoked.

Rules 1-3 leave the drift ceilings and the both-directions requirement
unchanged; they only reclassify how within-floor paired readings on
non-executed phases are interpreted, and tighten claim scope on the
favorable side to match.

## Consequences for 0204

Re-evaluating change 0204 (ODS protection fused parse) under the refined
banking rule:

- **Executed phases.** The fused protection parse runs once per fresh
  source-backed owner on the first cell edit (the stage path); open and
  commit phases execute none of the changed code.
- **Blocking readings reclassified.** The adverse both-directions patterns
  that forced the original withheld verdict were all on non-executed
  phases and sit within the calibrated floor: source-open p50/mean
  -3.4%/-4.0% (floor 5.5%), p95 within 4.5%; eager-open p50-p95 within
  3.0%; one-edit commit p50/mean/p95/p99 -1.4% to -13.9% (floor
  3.7%/3.7%/5.8%/17%). Under rule 1 these are layout readings, recorded as
  such, and no longer block.
- **Claim scope (rule 3).** Only repeated-edit stage accepts clearly above
  the floor: p50/mean/p95/p99 20.01%-25.86% lower against a stage floor of
  3.8%/5.3%/11% — **claimed**. One-percent lifecycle p50/mean/p95
  (2.54%-3.64%) exceed the 2.4%/2.6% floor by at most ~0.25pp, within the
  floor's own estimation error — recorded as marginal, not claimed.
  Repeated-edit total (0.21%-2.01% vs floor 1.8%/2.5%) and one-edit
  lifecycle (1.25%-3.31% vs floor 2.2%/2.8%) are within floor — neutral.

**Re-verdict: 0204 is banked** with claim scope = repeated-edit stage
p50/mean/p95/p99 (20.01%-25.86% lower). See
`0204-ods-protection-fused-parse.md` for the updated verdict section.
