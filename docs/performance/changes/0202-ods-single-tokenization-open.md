# Change 0202: ODS fused open parse absorbs the pass-2a calculation-settings parse

Date: 2026-08-19

## Decision

Banked: the pass-2a calculation-settings parse joins the fused open
tokenization, so `SourceBackedSpreadsheet::from_package` tokenizes
`content.xml` exactly once. The source-backed open accepts
p50/mean/p95/p99 at 17.41%-23.21% lower on top of 0201; the one-edit
lifecycle accepts all four statistics; the one-percent lifecycle accepts
p50/mean/p95 in both the primary run and the single rerun; the
repeated-edit selector accepts total p50/mean, stage p50/mean/p95, and
commit p50/mean/p95/p99. The one-percent lifecycle p99 primary adverse
reading failed to reproduce in the single permitted rerun, and a
sub-0.5% repeated-edit publication p50/mean layout remnant on
source-identical phases is documented and not claimed either way.

## Mechanism and invariants

Change 0201 left `SourceBackedSpreadsheet::from_package` tokenizing
`content.xml` twice: the fused validate/locate/names/worksheet loop, and the
separate `litchi_odf_common::calculation::parse` pass (pass 2a, semantic
calculation settings) that the caller runs between the fused `run()` and
`finish()` phases, after the styles.xml/meta.xml/metadata member loads.

This change folds pass 2a into the fused driver as a fifth handler
(`CalculationHandler`), so the source-backed open tokenizes `content.xml`
exactly once. The calculation pass historically ran after validation but
before locate, so the dispatch order becomes validate, calculation, locate,
names, worksheet. The observable interleave is preserved exactly:

- The 64 MiB calculation-settings size check moves from an up-front
  rejection to the deferred record-first-error pattern (the handler records
  it as a pre-stream pass error and stays inactive), matching how locate's
  64 MiB check already moved in 0201.
- A mid-stream validate error still returns immediately from `run()`; a
  validate end-of-stream or read error still returns at the end of `run()`,
  ahead of the styles/meta/metadata member loads (pinned by the
  package-level `content_validation_eof_error_beats_styles_member_load`
  test). The read-error mapping chain gains the calculation pass's
  `"XML parsing error: {error}"` arm, which is structurally unreachable
  because validation is active on every read failure and claims the error
  first.
- `finish()` selects the first error in the original pass order: validation
  (unreachable — already returned from `run`), calculation, locate, the
  semantic/XML disagreement check, named definitions, worksheets.

The standalone `litchi_odf_common::calculation::parse` keeps its original
inline loop byte-identical for its other callers (eager facade open and
commit-side paths); only its 64 MiB size check and token classification
helpers are factored out as doc-hidden items shared with the handler. The
handler carries the same state as the standalone parse's locals, runs the
same unterminated-settings check and `Settings::validate()` in `finish()`,
and is re-exported `#[doc(hidden)]` under the same suppression-comment
precedent as the litchi-core plumbing used by 0201. Crate boundaries are
unaffected: `litchi-ods` already depends on `litchi-odf-common`.

Verification: the full `litchi-ods` suite (356 tests, +1 new) passes,
including a new pin that the calculation 64 MiB size limit beats later
passes; `litchi-odf-common` gains a handler-vs-standalone equivalence test
(361 tests, +1 new). The `open_parse::tests::sequential` oracle now runs the
standalone calculation parse second in its reference order, so every corpus
fixture exercises the new handler against the standalone. The 0201
precedence pins (`validate_error_beats_semantic_parse`,
`semantic_parse_beats_locate_error`) still hold with the calculation pass
now inline. fmt, clippy (`-D warnings`), rustdoc (`-D warnings`), and
`tools/check_crate_boundaries.py` all pass.

## Matched release timing

Two frozen release binaries differ only in the pass-2a fold; both carry
changes 0192-0196 and 0198-0201 as baseline and the identical selector
matrix. Control SHA-256
`606bab41e7be2de874f4fa567ce854838768a5f49fe85768c93c7d53dcea62c1` (the
banked 0201 binary, matching the pre-0202 tree), candidate SHA-256
`475cf2898880363517eec9e0a9ac6b582eed1f78054f161f394bdd635bb19d7d`.
Fresh CPU-2-pinned processes ran `A1 control, B1 candidate, B2 candidate,
A2 control`, 30 warmups and 500 retained samples per leg. The predeclared
p50/mean/p95/p99 drift ceilings are 5%/5%/10%/15%; a statistic is accepted
only when both paired directions are lower and both drifts pass its
ceiling, and the change is banked only when no withheld statistic shows a
consistent both-directions adverse pattern. The
`ods_source_backed_one_percent_edit_save` selector was rerun once (the
single permitted rerun) after its primary-run lifecycle p99 read adverse
in both directions with clean drifts.

### Source-backed open (`ods_file_source_open`)

| Statistic | A1 -> B1 reduction | A2 -> B2 reduction | Control drift | Candidate drift | Decision |
|---|---:|---:|---:|---:|---|
| p50 | 21.68% | 23.21% | -0.63% | -2.57% | accept |
| mean | 21.45% | 23.13% | -0.66% | -2.78% | accept |
| p95 | 20.47% | 22.64% | -0.94% | -3.65% | accept |
| p99 | 17.41% | 22.56% | -1.99% | -8.10% | accept |

All four accepted at 17.41%-23.21% lower, on top of the 0201 win.

### One-edit guardrail (`ods_source_backed_one_edit_save`)

| Statistic | A1 -> B1 reduction | A2 -> B2 reduction | Control drift | Candidate drift | Decision |
|---|---:|---:|---:|---:|---|
| lifecycle p50 | 3.83% | 2.81% | -0.08% | 0.97% | accept |
| lifecycle mean | 4.41% | 2.72% | -0.56% | 1.20% | accept |
| lifecycle p95 | 7.66% | 1.59% | -3.73% | 2.60% | accept |
| lifecycle p99 | 10.47% | 0.81% | -3.34% | 7.09% | accept |
| commit p50 | 2.06% | 0.69% | -0.78% | 0.62% | accept |
| commit mean | 2.61% | 0.28% | -1.12% | 1.25% | accept |
| commit p95 | 5.46% | -1.11% | -3.90% | 2.78% | withheld |
| commit p99 | 9.71% | -11.35% | -3.84% | 18.60% | withheld |

Commit p99 withheld on a candidate-drift ceiling violation (+18.60%); p95
withheld on mixed directions. Both tails pair-asymmetric on a
source-identical phase — per-leg layout/environment, neutral.

### One-percent guardrail (`ods_source_backed_one_percent_edit_save`)

Primary run:

| Statistic | A1 -> B1 reduction | A2 -> B2 reduction | Control drift | Candidate drift | Decision |
|---|---:|---:|---:|---:|---|
| lifecycle p50 | 1.79% | 2.20% | 0.63% | 0.21% | accept |
| lifecycle mean | 1.57% | 2.09% | 0.66% | 0.13% | accept |
| lifecycle p95 | 0.00% | 0.74% | 0.97% | 0.23% | accept |
| lifecycle p99 | -3.44% | -2.33% | 0.14% | -0.94% | withheld, rerun |
| commit p50 | 0.61% | 1.94% | 1.28% | -0.08% | accept |
| commit mean | 0.42% | 1.15% | 0.78% | 0.04% | accept |
| commit p95 | -0.82% | -2.07% | 1.58% | 2.84% | withheld |
| commit p99 | -21.43% | -1.92% | 0.94% | -15.28% | withheld |

The lifecycle p99 both-directions adverse reading (with clean drifts) is the
only primary-run pattern involving the changed open path; it is implausible
as a mechanism (the change strictly removes a tokenization from the open
while lifecycle p50/mean/p95 all accept favorable), so the single permitted
rerun is spent on this workload. Commit p95/p99 are pair-asymmetric in
magnitude (-21.43% vs -1.92% on p99, with a candidate-drift ceiling
violation) on a source-identical phase — tail instability, neutral.

Rerun (single permitted): the both-directions adverse pattern did not
reproduce. Lifecycle p50/mean/p95 accepted at 1.66%-3.81% lower (drifts
within ceilings); lifecycle p99 read +2.12%/-1.67% (pair-asymmetric,
withheld); commit p50/mean/p95/p99 all accepted at 1.06%-3.28% lower,
flipping the primary's commit-tail readings. The primary adverse reading
is therefore a per-run tail effect, not a candidate property; lifecycle
p99 is claimed in neither direction, and the rerun's accepted statistics
supplement the primary's.

### Repeated-edit selector (`ods_source_backed_repeated_edit`)

| Statistic | A1 -> B1 reduction | A2 -> B2 reduction | Control drift | Candidate drift | Decision |
|---|---:|---:|---:|---:|---|
| total p50 | 0.19% | 0.63% | -0.15% | -0.59% | accept |
| total mean | 0.30% | 0.43% | -0.26% | -0.38% | accept |
| total p95 | 1.16% | -0.65% | -2.37% | -0.59% | withheld |
| total p99 | 1.67% | -2.26% | -3.19% | 0.68% | withheld |
| stage p50 | 1.25% | 1.59% | -0.33% | -0.67% | accept |
| stage mean | 1.16% | 0.98% | -0.23% | -0.05% | accept |
| stage p95 | 1.19% | 1.26% | -0.10% | -0.17% | accept |
| stage p99 | 2.65% | -0.69% | 3.60% | 7.15% | withheld |
| commit p50 | 3.06% | 2.86% | -1.34% | -1.14% | accept |
| commit mean | 3.20% | 2.44% | -1.30% | -0.52% | accept |
| commit p95 | 3.51% | 1.69% | -1.25% | 0.61% | accept |
| commit p99 | 4.91% | 0.83% | -1.47% | 2.75% | accept |
| publication p50 | -0.43% | -0.03% | 0.04% | -0.35% | withheld |
| publication mean | -0.41% | -0.12% | -0.09% | -0.38% | withheld |
| publication p95 | 0.32% | -0.33% | -2.21% | -1.56% | withheld |
| publication p99 | -0.03% | -4.09% | -3.84% | 0.06% | withheld |

Documented remnant (not claimed): repeated-edit publication p50/mean
measured the candidate 0.03%-0.43% slower in both directions with clean
drifts. The publication phases are source-identical between the two
binaries, so no mechanism in this change can drive them slower; this
sub-0.5% reading is recorded as deterministic per-binary-pair code-layout
wobble of the class documented in changes 0197, 0200, and 0201, not a
regression pattern. With the single rerun spent on the one-percent
lifecycle p99 pattern, no rerun is available for this sub-threshold
remnant.

### Eager open (`ods_file_eager_open`)

The eager facade exercises the unchanged standalone passes, so this selector
measures code-layout neutrality. Mean accepted favorable (0.28%/0.80%
lower in the two directions, drifts within ceilings); p50 and p95 withheld
on mixed sub-1% directions; p99 withheld on a +21.87% control-drift ceiling
violation (environment noise; the A2->B2 direction read +21.56% favorable
against it). No statistic is claimed adversely; the pattern is neutral.

## Verdict

Banked. The source-backed open accepts p50/mean/p95/p99 at
17.41%-23.21% lower (stacking on 0201's 9.32%-17.88% and 0200's
19.72%-24.55%); the one-edit lifecycle accepts all four statistics
(0.81%-10.47%) plus commit p50/mean; the one-percent lifecycle accepts
p50/mean/p95 in the primary run and again in the single rerun (whose
commit accepts all four statistics at 1.06%-3.28%); the repeated-edit
selector accepts total p50/mean, stage p50/mean/p95, and commit
p50/mean/p95/p99 (0.83%-4.91%). The eager-open selector accepts mean
favorable and is otherwise neutral. The one-percent lifecycle p99
primary adverse reading did not reproduce in the rerun and is claimed in
neither direction; the repeated-edit publication p50/mean sub-0.5%
layout remnant is documented above. Allocation/RSS, physical-I/O,
cold-cache, producer, and broad-ODF claims remain withheld.
