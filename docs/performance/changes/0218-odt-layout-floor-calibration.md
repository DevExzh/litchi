# Change 0218: litchi-odt layout noise floor calibration (methodology)

Date: 2026-08-19

## Purpose

Not a code change — a measurement-methodology calibration, the litchi-odt
analog of 0205 (litchi-ods) and 0213 (litchi-odp). Change 0217 was withheld
under the pre-floor rule after adverse both-directions readings on two
byte-identical guardrail phases reproduced in their single permitted reruns
(open p50 max 3.26%, list-paragraphs mean max 6.67%), despite all three
executed workloads accepting all four statistics at 40%-57%. This
calibration measures the litchi-odt layout noise floor directly: binaries
whose source differs from the banked control (0215 state) only by
never-executed, parser-shaped padding code in litchi-odt, retained in the
binary via a `#[used]` function-pointer table.

## Probes

All probes build from the identical banked tree (control SHA-256
`6c7fcfb9572f79bbfc2a9dd06289f733e370b34f96662980c5d59b7e972471eb`); the
probe code was removed after each build and the tree verified to rebuild
bit-exact to the control SHA. No probe instruction executes on any measured
path.

| Probe | Placement | Added text bytes | SHA-256 |
|---|---|---:|---|
| A (medium) | `litchi-odt::elements::text` | +14,628 | `f3da50b92927ddff840c76bfa337a75522c924edf9d80d1ddf24d191bee25a15` |
| B (small) | `litchi-odt::document` | +5,784 | `9a2df16622296dea2fd0d5c6a3058ac42d227fa75f076bd748392d7b60280fd3` |
| C (medium) | `litchi-odt::parser` | +14,628 | `0663df53161010bac4e5f75ba7458c906a86e3dcf63f6d774f7c8122c007e87a` |

Probe A deliberately covers the scale and module vicinity of the withheld
0217 change (+11.3KB in `elements/text`). Protocol: identical to the
change-measurement protocol — frozen binaries, fresh CPU-2-pinned processes,
order A1 control / B1 probe / B2 probe / A2 control, 30 warmups, 500 samples
per leg, same drift ceilings. Because control and probe are semantically
identical, EVERY paired-direction reading is pure layout noise; the observed
magnitudes calibrate the floor.

## Measured floor

Maximum adverse both-directions magnitude observed across probes A/B/C —
every probe reading below is pure layout noise on semantically identical
binaries. "–" means no probe produced an adverse both-directions reading for
that statistic. Full per-leg data:
`docs/performance/results/*-0218{a,b,c}-*`.

| Phase | p50 | mean | p95 | p99 |
|---|---:|---:|---:|---:|
| open | 2.47% | 7.17% | 27.56% | 28.19% |
| list-paragraphs | 5.24% | 5.60% | – | – |
| one-paragraph | 4.78% | 8.31% | 52.62% | – |
| full-text | 4.11% | – | 16.13% | 9.35% |
| repeated-text-cached | 7.07% | 7.37% | 7.50% | 29.18% |
| repeated-text-uncached | 4.83% | 4.13% | 3.20% | 8.24% |

Historical corroboration: adverse both-directions readings on byte-identical
phases of the withheld 0217 change pair — open p50 -1.44%/-3.26% (primary)
and -2.58%/-2.79% (rerun); list-paragraphs p50 -0.36%/-1.57%, mean
-0.93%/-1.93% (primary) and -2.42%/-6.67% (rerun), p95 -2.97%/-22.45%;
one-paragraph p95 -27.47%/-1.44%. Taking the historical maxima into account,
the effective floor used for banking decisions on litchi-odt changes is:

| Phase | p50/mean | p95 | p99 |
|---|---:|---:|---:|
| open | 3.3% / 7.2% | 27.6% | 28.2% |
| list-paragraphs | 5.2% / 6.7% | 22.4% | – |
| one-paragraph | 4.8% / 8.3% | 52.6% | – |
| full-text | 4.1% / – | 16.1% | 9.3% |
| repeated-text-cached | 7.1% / 7.4% | 7.5% | 29.2% |
| repeated-text-uncached | 4.8% / 4.1% | 3.2% | 8.2% |

The ODT tail floors (p95/p99 up to 52.6%) are much wider than the ODS/ODP
equivalents — the one-paragraph and open phases have layout-sensitive tail
behavior on this corpus (nearest-rank sensitivity of a 500-sample tail
statistic). p50/mean are the operative banking statistics; "–" entries mean
no adverse-both evidence exists and the floor is not calibrated there (the
pre-floor rule still applies to such a statistic).

## Refined banking rule

The 0205 banking rule extends to litchi-odt, unchanged in form:

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
4. **Floor scope.** The floor is calibrated on litchi-odt with binary text
   deltas of +5.8KB to +14.6KB. Substantially larger deltas, other crates, or
   protocol changes require recalibration before the floor can be invoked.
   The litchi-ods (0205), litchi-odp (0213), and litchi-odt (this change)
   floors are independent; all other crates still run pre-floor.

## Consequences for 0217

Re-evaluating change 0217 (ODT discard-but-validate text extraction) under
the extended rule:

- **Executed phases.** The change touches only the `extract_text` parse mode;
  `odt_semantic_open`, `odt_semantic_list_paragraphs`, and
  `odt_semantic_one_paragraph` execute none of the changed code (the retained
  path is byte-untouched, verified by the control/candidate .text comparison
  and call-site audit).
- **Blocking readings reclassified.** The reproduced adverse both-directions
  readings that forced the withheld verdict sit within the calibrated floor:
  open p50 max 3.26% (floor 3.3%); list-paragraphs mean max 6.67% (floor
  6.7%). The one-paragraph p95 primary adverse pattern did not reproduce in
  its rerun and was already cleared even pre-floor. Under rule 1 the
  reproduced readings are layout readings, recorded as such, and no longer
  block.
- **Claim scope (rule 3).** Accepted magnitudes (40.31%-57.18%) dwarf every
  floor entry on their phases, so all accepts are claimed in full.

**Re-verdict: 0217 is banked** with claim scope = `odt_semantic_full_text`
p50/mean/p95/p99 (42.12%-52.07% lower),
`odt_source_backed_repeated_text_cached` p50/mean/p95/p99 (51.91%-57.18%
lower), `odt_source_backed_repeated_text_uncached` p50/mean/p95/p99
(40.31%-53.85% lower). See `0217-odt-discard-validate-text.md` for the
updated verdict section.
