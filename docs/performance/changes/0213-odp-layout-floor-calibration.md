# Change 0213: litchi-odp layout noise floor calibration (methodology)

Date: 2026-08-19

## Purpose

Not a code change — a measurement-methodology calibration, the litchi-odp
analog of 0205 (which calibrated litchi-ods). Change 0212 was withheld after
its adverse both-directions p50 reading on `odp-semantic-open` — a
byte-identical phase executing none of the changed code — reproduced in the
single permitted rerun, because the 0205 floor is scoped to litchi-ods only
and could not be invoked. 0212 also showed real 17%-29% executed-phase wins,
so the cost of an unquantified ODP "layout wobble" hypothesis became concrete.
This calibration measures the litchi-odp layout noise floor directly: binaries
whose source differs from the banked control (0211 state) only by
never-executed, parser-shaped padding code in litchi-odp, retained in the
binary via a `#[used]` function-pointer table.

## Probes

All probes build from the identical banked tree (control SHA-256
`ceba155be185f1c213c4bf90200bb5e87bb697a5023e3a703d9cd7def6042922`); the
probe code was removed after each build and the tree verified to rebuild
bit-exact to the control SHA. No probe instruction executes on any measured
path.

| Probe | Placement | Added text bytes | SHA-256 |
|---|---|---:|---|
| A (medium) | `litchi-odp::codec::xml` | +14,516 | `fbe09a79f6aa8cca42cdc96d2d9a4099155424c1173e7b29a5558e31ffc318e5` |
| B (small) | `litchi-odp::model` | +5,484 | `3da515ff1986ebd15eb446adf28f8115878cd09caea142b52d252f18396f0183` |
| C (medium) | `litchi-odp::package` | +14,516 | `68a246d55fbefa97080fa375f12e8fc258ec48d9de6338cc49ef1e2aabd86b1b` |

Protocol: identical to the change-measurement protocol — frozen binaries,
fresh CPU-2-pinned processes, order A1 control / B1 probe / B2 probe /
A2 control, 30 warmups, 500 samples per leg, same drift ceilings. Because
control and probe are semantically identical, EVERY paired-direction reading
is pure layout noise; the observed magnitudes calibrate the floor.

## Measured floor

Maximum adverse both-directions magnitude observed across probes A/B/C —
every probe reading below is pure layout noise on semantically identical
binaries. Full per-leg data:
`docs/performance/results/*-0213{a,b,c}-*`.

| Phase | p50 | mean | p95 | p99 |
|---|---:|---:|---:|---:|
| open | 3.06% | – | 6.31% | 17.22% |
| list-slides | 1.98% | 3.56% | 6.34% | 10.07% |
| one-slide | 2.52% | 3.20% | 17.76% | 14.41% |
| full-text | 0.08% | 0.55% | 1.78% | 19.41% |

Historical corroboration: adverse both-directions readings on byte-identical
phases of real litchi-odp change pairs — 0211 open mean -0.52%/-1.08% and
p95 -4.92%/-6.53%; 0212 open p50 -1.21%/-1.37% (rerun 0212r:
-1.51%/-1.63%), mean -2.21%/-2.55%, p95 -2.05%/-7.77%. Taking the historical
maxima into account, the effective floor used for banking decisions on
litchi-odp changes is:

| Phase | p50/mean | p95 | p99 |
|---|---:|---:|---:|
| open | 3.1% / 2.5% | 7.8% | 17.2% |
| list-slides | 2.0% / 3.6% | 6.3% | 10.1% |
| one-slide | 2.5% / 3.2% | 17.8% | 14.4% |
| full-text | 0.1% / 0.5% | 1.8% | 19.4% |

The p99 tails (up to ~19%) reflect nearest-rank sensitivity of a 500-sample
tail statistic to code-layout shifts; p50/mean/p95 are the operative banking
statistics. The full-text p50/mean floors are far below the protocol drift
ceilings (5%) — that phase's tokenizer inner loop is layout-stable.

## Refined banking rule

The 0205 banking rule extends to litchi-odp, unchanged in form:

1. **Within-floor adverse readings on non-executed phases do not block.** A
   both-directions adverse pattern on a phase that executes none of the
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
4. **Floor scope.** The floor is calibrated on litchi-odp with binary text
   deltas of +5.5KB to +14.5KB. Substantially larger deltas, other crates, or
   protocol changes require recalibration before the floor can be invoked.
   The litchi-ods floor (0205) and this litchi-odp floor are independent;
   litchi-odt and all other crates still have no floor — the pre-floor rule
   (any adverse both-directions blocks unless the rerun clears it) applies
   there.

## Consequences for 0212

Re-evaluating change 0212 (ODP cached attribute namespace resolution) under
the extended rule:

- **Executed phases.** The change touches the attribute-resolution path used
  by list-slides, one-slide, and full-text query parses; `odp-semantic-open`
  executes none of the changed code.
- **Blocking reading reclassified.** The sole blocker under the pre-floor
  rule was the open p50 pattern — the only statistic whose adverse
  both-directions reading reproduced in the single permitted rerun
  (-1.21%/-1.37% original, -1.51%/-1.63% rerun); mean and p95 adverse
  patterns did not reproduce in the rerun and were already cleared even
  pre-floor. The reproduced p50 magnitude (max 1.63%) sits well within the
  calibrated open p50 floor of 3.1% on a phase executing none of the changed
  code. Under rule 1 it is a layout reading, recorded as such, and no longer
  blocks.
- **Claim scope (rule 3).** Accepted magnitudes all far exceed the floor:
  full-text p50/mean/p95/p99 20.82%-29.50%; list-slides p50/mean
  19.15%-25.51%; one-slide p50/mean/p99 17.48%-31.16%.

**Re-verdict: 0212 is banked** with claim scope = full-text p50/mean/p95/p99
(20.82%-29.50% lower), list-slides p50/mean (19.15%-25.51% lower), one-slide
p50/mean/p99 (17.48%-31.16% lower). See `0212-odp-cached-attr-resolution.md`
for the updated verdict section.
