# Change 0204: ODS protection parse fuses its two content.xml passes

Date: 2026-08-19

## Decision

**Banked** (re-verdict under the 0205 calibrated layout-noise floor;
originally withheld and reverted, then restored byte-exact). The fusion's
targeted win is real: repeated-edit stage p50/mean/p95/p99 20.01%-25.86%
lower, accepted and above the calibrated stage layout-noise floor — this
is the claim scope. The originally blocking adverse readings (source-open
2.87%-11.53% slower both directions, reproduced in the single rerun;
source-identical eager-open and one-edit commit showing the same pattern)
all sit within the per-binary-pair layout noise floor measured by change
0205 for phases executing no changed code, and are recorded as layout
readings under the refined banking rule. See the Verdict section for the
full re-verdict reasoning.

## Mechanism and invariants

`protection::codec::parse` — the commit-side entry behind
`Snapshot::parse` and the source-backed `edit_protection` facts — tokenized
`content.xml` twice per call: `Location::parse` (source locator) and then
`model::protection::parse_protection` (semantic wire metadata), followed by
`CellStyleRegistry::parse` over styles.xml and the automatic-styles
fragment (separate documents, out of scope). Profiling of the stage phase
attributes ~14% self to `protection::codec::parse` plus ~2.2% to
`parse_protection` plus their share of the NsReader machinery; the parse
runs once per fresh source-backed owner on the first cell edit, so it sits
on the measured edit-workload stage path.

This change fuses the two content.xml passes into one shared `NsReader`
tokenization, following the 0200-0202 handler/driver pattern. Each loop
body becomes a handler (`LocationHandler`, `ProtectionHandler`) with
checks, limits, and error messages transcribed verbatim; the driver
dispatches every event to the location handler first, then the protection
handler. Error selection preserves the historical interleave exactly:
`Location::parse`'s 64 MiB size check stays up front (`parse_protection`
has no size check); read failures map to the identical
`"XML parsing error: {error}"` string in both passes and are recorded
against the first still-active handler in pass order; mid-stream handler
errors are recorded per pass and the final selection is locator mid-stream
error, locator finish checks, protection mid-stream error, protection
finish checks — reproducing the sequential precedence where the locator
ran to completion before the protection pass started. The
`"ODS protection sheet parser and source locator disagree"` check stays
in `codec::parse` in its exact position. The standalone
`Location::parse` and `parse_protection` shells keep their original
inline loops byte-identical as oracles for the equivalence tests.

Verification: the full `litchi-ods` suite (369 tests, +13 new) passes.
The new equivalence module compares the fused driver against the
standalone shells over the .ods corpus (content.xml + styles.xml) and
synthetic documents covering custom table/loext prefixes, mixed and
nested sheets, entity-decoded names, location-beats-protection ordering
at the same event, protection-only errors, duplicate table-protection
mapping preference, unterminated/missing-spreadsheet finish precedence,
the shared malformed-XML read mapping, and the 64 MiB size limit. fmt,
clippy (`-D warnings`), rustdoc (`-D warnings`), and
`tools/check_crate_boundaries.py` all pass.

## Matched release timing

Two frozen release binaries differ only in the protection-parse fusion;
both carry changes 0192-0196 and 0198-0202 as baseline and the identical
selector matrix. Control SHA-256
`475cf2898880363517eec9e0a9ac6b582eed1f78054f161f394bdd635bb19d7d` (the
banked 0202 binary), candidate SHA-256
`5f0dab648cdf8f693ec01c171bfb81b2776e9e5abfe117b5be85c42a5bd89f66`.
Fresh CPU-2-pinned processes ran `A1 control, B1 candidate, B2 candidate,
A2 control`, 30 warmups and 500 retained samples per leg. The predeclared
p50/mean/p95/p99 drift ceilings are 5%/5%/10%/15%; a statistic is accepted
only when both paired directions are lower and both drifts pass its
ceiling. The `ods_file_source_open` selector was rerun once (the single
permitted rerun) after its primary run read adverse in both directions
with clean drifts on a phase the change does not execute.

### Repeated-edit selector (`ods_source_backed_repeated_edit`)

| Statistic | A1 -> B1 reduction | A2 -> B2 reduction | Control drift | Candidate drift | Decision |
|---|---:|---:|---:|---:|---|
| stage p50 | 25.81% | 25.79% | 0.26% | 0.29% | accept |
| stage mean | 25.44% | 25.86% | 0.86% | 0.29% | accept |
| stage p95 | 24.07% | 24.82% | 1.72% | 0.72% | accept |
| stage p99 | 20.01% | 23.78% | 5.77% | 0.77% | accept |
| total p50 | 1.24% | 0.93% | -0.15% | 0.16% | accept |
| total mean | 1.07% | 1.01% | -0.02% | 0.04% | accept |
| total p95 | 1.05% | 1.71% | 0.85% | 0.17% | accept |
| total p99 | 0.21% | 2.01% | 2.00% | 0.15% | accept |
| commit p50-p99 | -2.50%- -1.14% | 0.09%-1.66% | mixed | mixed | withheld |
| publication p50-p99 | mixed | mixed | within ceilings | within ceilings | withheld (p95 accept) |

### One-percent guardrail (`ods_source_backed_one_percent_edit_save`)

| Statistic | A1 -> B1 reduction | A2 -> B2 reduction | Control drift | Candidate drift | Decision |
|---|---:|---:|---:|---:|---|
| lifecycle p50 | 3.54% | 2.54% | -0.43% | 0.60% | accept |
| lifecycle mean | 3.44% | 2.61% | -0.44% | 0.42% | accept |
| lifecycle p95 | 3.64% | 2.85% | -0.71% | 0.11% | accept |
| lifecycle p99 | 1.79% | 2.33% | -2.87% | -3.40% | accept |
| commit p50 | -0.23% | -0.58% | 0.28% | 0.64% | withheld |
| commit mean | -0.44% | -0.23% | 0.43% | 0.22% | withheld |
| commit p95 | 0.81% | 0.59% | 0.32% | 0.54% | accept |
| commit p99 | -9.28% | 3.05% | -3.28% | -14.19% | withheld |

### One-edit guardrail (`ods_source_backed_one_edit_save`)

| Statistic | A1 -> B1 reduction | A2 -> B2 reduction | Control drift | Candidate drift | Decision |
|---|---:|---:|---:|---:|---|
| lifecycle p50 | 1.25% | 3.31% | 0.77% | -1.34% | accept |
| lifecycle mean | 1.50% | 3.07% | 0.59% | -1.01% | accept |
| lifecycle p95 | 2.08% | 2.19% | -0.04% | -0.16% | accept |
| lifecycle p99 | -3.23% | -2.08% | 0.31% | -0.81% | withheld |
| commit p50 | -2.77% | -1.63% | 0.01% | -1.10% | withheld |
| commit mean | -2.98% | -1.83% | -0.01% | -1.13% | withheld |
| commit p95 | -2.78% | -1.37% | 1.09% | -0.30% | withheld |
| commit p99 | -13.90% | -6.99% | -0.86% | -6.88% | withheld |

### Source-backed open (`ods_file_source_open`) — change executes none of this code

| Statistic | A1 -> B1 reduction | A2 -> B2 reduction | Control drift | Candidate drift | Decision |
|---|---:|---:|---:|---:|---|
| p50 | -3.41% | -4.01% | -1.20% | -0.62% | withheld, rerun |
| mean | -3.40% | -3.84% | -1.23% | -0.80% | withheld, rerun |
| p95 | -1.48% | -1.61% | -0.91% | -0.78% | withheld |
| p99 | -4.50% | 3.01% | 1.89% | -5.43% | withheld |

### Eager open (`ods_file_eager_open`) — change executes none of this code

| Statistic | A1 -> B1 reduction | A2 -> B2 reduction | Control drift | Candidate drift | Decision |
|---|---:|---:|---:|---:|---|
| p50 | -2.28% | -1.67% | 0.51% | -0.09% | withheld |
| mean | -2.22% | -2.01% | 0.09% | -0.12% | withheld |
| p95 | -2.94% | -2.98% | -1.44% | -1.41% | withheld |
| p99 | 0.16% | -6.05% | -4.97% | 0.94% | withheld |

### Source-open rerun

The single permitted rerun reproduced the adverse pattern in full:
p50 -2.87%/-3.57%, mean -3.39%/-3.94%, p95 -4.85%/-7.20%, p99
-11.53%/-8.77% — both directions, all drifts within ceilings. The effect
is deterministic for this binary pair, not environment noise. Because the
fused protection parse executes zero instructions on the open path, this
is a per-binary-pair code-layout effect of the class documented in
changes 0197 and 0203 — reproduced, and larger than the sub-1.5%
documented-wobble band.

## Verdict

**Banked (re-verdict under the 0205 calibrated layout-noise floor).**
History: originally withheld and reverted byte-exact because source-open —
which executes none of the changed code — measured the candidate
2.87%-11.53% slower in both paired directions in the primary run AND in
the rerun, with source-identical eager-open (p50/mean/p95 1.67%-2.98%
slower both directions) and one-edit commit (p50/mean/p95 1.37%-2.98%
slower both directions) showing the same consistent pattern. Change 0205
then calibrated the per-binary-pair layout noise floor directly with
never-executed probe binaries: every one of those adverse magnitudes sits
within the measured floor for its phase (source-open floor 5.5%/4.5%/36%
for p50-mean/p95/p99; eager-open 3.0%; one-edit commit
3.7%/5.8%/17%). Under the refined banking rule (see
[`0205-layout-noise-floor-calibration.md`](0205-layout-noise-floor-calibration.md)),
within-floor adverse readings on phases executing no changed code are
layout readings and do not block. The implementation was restored
byte-exact from the preserved copy; the harness rebuild is bit-identical
to the measured candidate (SHA-256
`5f0dab648cdf8f693ec01c171bfb81b2776e9e5abfe117b5be85c42a5bd89f66`), the
`litchi-ods` suite passes (369 tests), and fmt/clippy/rustdoc/boundary
gates pass.

Claim scope (only statistics exceeding the calibrated floor are claimed):
**repeated-edit stage p50/mean/p95/p99 are 20.01%-25.86% lower** (stage
floor 3.8%/5.3%/11%). One-percent lifecycle p50/mean/p95 (2.54%-3.64%)
exceed the 2.4%/2.6% floor by at most ~0.25pp — within the floor's
estimation error, recorded as marginal and not claimed. Repeated-edit
total (0.21%-2.01% vs floor 1.8%/2.5%) and one-edit lifecycle
(1.25%-3.31% vs floor 2.2%/2.8%) are within floor — recorded as neutral.
All adverse readings listed above are recorded as within-floor layout
readings.

Methodology note: 0197, 0203, and 0204 all hit systematic per-binary-pair
layout effects of several percent on litchi-ods hot phases. The 0205
calibration quantified that floor and reclassified within-floor readings
on non-executed phases; adverse patterns on executed phases still block
unless cleared by the single permitted rerun.
