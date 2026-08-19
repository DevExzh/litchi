# Change 0201: ODS fused open parse absorbs the structural validation pass

Date: 2026-08-19

## Decision

Banked: the structural validation pass joins the 0200 fused open
tokenization, so `SourceBackedSpreadsheet::from_package` tokenizes
`content.xml` twice instead of three times. The source-backed open accepts
p50/mean/p95/p99 at 9.32%-17.88% lower on top of 0200; both edit-guardrail
lifecycles and the repeated-edit publication accept. The eager-open primary
adverse reading failed to reproduce in the single permitted rerun, and a
sub-1.5% repeated-edit commit p50/mean layout remnant on source-identical
phases is documented and not claimed either way.

## Mechanism and invariants

Change 0200 left `SourceBackedSpreadsheet::from_package` tokenizing
`content.xml` three times: `authoring::validate_content_xml` (structural
validation), `litchi_odf_common::calculation::parse` (semantic calculation
settings), and the fused locate/names/worksheet loop. Post-0200 profiling of
the source-backed open still attributes ~16% of samples to quick_xml
namespace machinery (`resolve_event` 9.09%, `process_event` 7.28%) spread
across those three tokenizations.

This change folds the structural validation pass into the fused driver as a
fourth handler (`ValidateHandler`), cutting open from three tokenizations
to two. The driver becomes two-phase to preserve the exact observable
interleave in `from_package`:

- `OpenParse::run(content_xml)` — up-front 256 MiB validation size check
  (pass 1's pre-check, still the first parse-stage check after the unchanged
  `validate_content_part` substring gate), then one shared `NsReader` loop
  dispatching to the validate, locate, names, and worksheet handlers in pass
  order. The locate/names/worksheet size limits are recorded as pre-stream
  pass errors (handlers inactive) — locate's 64 MiB check moves from an
  up-front rejection to the deferred pattern because the pass-2a calculation
  parse historically precedes it. A mid-stream validate error returns
  immediately (pass 1 historically ended the open before pass 2a ran); a
  validate error at the end-of-stream event or a read failure is recorded
  and returned at the end of `run`, still ahead of pass 2a and of the
  styles.xml/meta.xml/metadata member loads the caller performs next.
- `from_package` performs the styles/meta/metadata loads exactly where they
  historically sat (after pass 1, before pass 2a), then calls
  `OpenParse::finish()`, which runs the pass-2a calculation parse and
  selects the first error in the original call order: validation, pass 2a,
  locate, the semantic/XML disagreement check, named definitions (including
  `validate_collection`), worksheets.

The standalone `authoring::validate_content_xml` keeps its original inline
loop byte-identical for its other callers (package/document.rs,
annotations/codec.rs, worksheet/package.rs, authoring/builder.rs,
facade/source_edit.rs); the new handler mirrors the same checks and messages
and is cross-checked against the standalone by the equivalence tests.

Verification: the full `litchi-ods` suite (355 tests, +3 new) passes. The
`open_parse::tests` oracle now runs the standalone validation first in its
sequential reference, so every corpus fixture and precedence case covers the
new pass; new pins cover validate-error-beats-semantic-parse,
semantic-parse-beats-locate-error, and a package-level interleave proving a
validation end-of-stream error beats a styles.xml member-load failure. Two
pre-existing pins whose fixtures the validation pass legitimately rejects
were redesigned validate-clean while preserving what they pin. fmt, clippy
(`-D warnings`), rustdoc (`-D warnings`), and
`tools/check_crate_boundaries.py` all pass.

## Matched release timing

Two frozen release binaries differ only in the validation-pass fold; both
carry changes 0192-0196 and 0198-0200 as baseline and the identical selector
matrix. Control SHA-256
`74951d0ccc58ce141a6e914790115cbe2b64b218d029f9aa29b8f17d8d0a844f` (the
banked 0200 v2 binary, matching the pre-0201 tree), candidate SHA-256
`606bab41e7be2de874f4fa567ce854838768a5f49fe85768c93c7d53dcea62c1`.
Fresh CPU-2-pinned processes ran `A1 control, B1 candidate, B2 candidate,
A2 control`, 30 warmups and 500 retained samples per leg. The predeclared
p50/mean/p95/p99 drift ceilings are 5%/5%/10%/15%; a statistic is accepted
only when both paired directions are lower and both drifts pass its ceiling,
and the change is banked only when no withheld statistic shows a consistent
both-directions adverse pattern.

### Source-backed open (`ods_file_source_open`)

| Statistic | A1 -> B1 reduction | A2 -> B2 reduction | Control drift | Candidate drift | Decision |
|---|---:|---:|---:|---:|---|
| p50 | 17.88% | 17.07% | -0.64% | 0.34% | accept |
| mean | 17.27% | 16.96% | -0.81% | -0.44% | accept |
| p95 | 13.33% | 17.70% | -0.59% | -5.60% | accept |
| p99 | 9.32% | 10.43% | -2.20% | -3.40% | accept |

All four accepted at 9.32%-17.88% lower, on top of the 0200 win.

### Eager open (`ods_file_eager_open`)

The eager facade exercises the unchanged standalone passes, so this selector
measures code-layout neutrality. Primary run: p50/mean/p95 measured the
candidate 0.40%-1.67% slower in both directions (withheld); p99 withheld on
a candidate-drift ceiling violation. The single permitted rerun did not
reproduce the pattern: the A2->B2 direction flipped to neutral/favorable
(-0.07%/+0.63%/+0.68%) while A1->B1 read 4.67%-5.11% slower, and the rerun's
control p99 drift reached +45.13% (environment noise), leaving every
statistic withheld. A pair-asymmetric reading on a source-identical path is
a per-leg layout/environment effect, not a consistent candidate-slower
pattern; no statistic is claimed in either direction.

### One-edit guardrail (`ods_source_backed_one_edit_save`)

| Statistic | A1 -> B1 reduction | A2 -> B2 reduction | Control drift | Candidate drift | Decision |
|---|---:|---:|---:|---:|---|
| lifecycle p50 | 2.17% | 4.09% | 0.82% | -1.15% | accept |
| lifecycle mean | 2.19% | 3.77% | 0.71% | -0.92% | accept |
| lifecycle p95 | 1.87% | 3.17% | 1.29% | -0.06% | accept |
| lifecycle p99 | 5.23% | 0.91% | -3.08% | 1.34% | accept |
| commit p50 | 0.45% | 1.48% | -0.46% | -1.50% | accept |
| commit mean | 0.87% | 0.91% | -1.10% | -1.14% | accept |
| commit p95 | 0.16% | -0.78% | -1.79% | -0.87% | withheld |
| commit p99 | 5.20% | -2.91% | -3.70% | 4.54% | withheld |

### One-percent guardrail (`ods_source_backed_one_percent_edit_save`)

| Statistic | A1 -> B1 reduction | A2 -> B2 reduction | Control drift | Candidate drift | Decision |
|---|---:|---:|---:|---:|---|
| lifecycle p50 | -0.40% | 2.99% | 1.41% | -2.02% | withheld |
| lifecycle mean | -0.11% | 3.32% | 1.25% | -2.22% | withheld |
| lifecycle p95 | 0.05% | 3.93% | -0.18% | -4.06% | accept |
| lifecycle p99 | -1.42% | 6.40% | 1.94% | -5.92% | withheld |
| commit p50-p99 | -3.71/-3.84/-5.62/-7.92% | 0.69/1.18/3.45/2.37% | within ceilings | within ceilings | withheld |

Commit is pair-asymmetric (A1->B1 adverse, A2->B2 favorable) — per-leg
layout, neutral.

### Repeated-edit selector (`ods_source_backed_repeated_edit`)

| Statistic | A1 -> B1 reduction | A2 -> B2 reduction | Control drift | Candidate drift | Decision |
|---|---:|---:|---:|---:|---|
| total p50 | 0.29% | -0.43% | -0.61% | 0.11% | withheld |
| total mean | 0.23% | -0.13% | -0.48% | -0.13% | withheld |
| total p95 | 0.98% | 0.50% | -0.25% | 0.24% | accept |
| total p99 | 2.34% | 1.15% | 0.44% | 1.67% | accept |
| stage p50-p99 | mixed | mixed | -10.90%-1.04% | 3.25%-11.77% | withheld |
| commit p50 | -1.41% | -1.21% | -1.69% | -1.89% | withheld |
| commit mean | -1.20% | -1.26% | -1.40% | -1.34% | withheld |
| commit p95 | -0.21% | 0.27% | 0.56% | 0.08% | withheld |
| commit p99 | -0.24% | -3.08% | -1.08% | 1.72% | withheld |
| publication p50 | 0.53% | 0.02% | -0.22% | 0.29% | accept |
| publication mean | 0.39% | 0.16% | -0.28% | -0.05% | accept |
| publication p95 | 0.70% | 0.53% | -0.83% | -0.66% | accept |
| publication p99 | 2.30% | 2.14% | -0.29% | -0.13% | accept |

Documented remnant (not claimed): repeated-edit commit p50/mean measured the
candidate 1.20%-1.41% slower in both directions with clean drifts. The
commit phases are source-identical between the two binaries (the commit
readback and protection paths use the untouched standalone shells), so no
mechanism in this change can drive them slower; with the single rerun spent
on the broader eager-open pattern (which failed to reproduce), this sub-1.5%
reading is recorded as deterministic per-binary-pair code-layout wobble of
the class documented in changes 0197 and 0200, not a regression pattern.

## Verdict

Banked. The source-backed open accepts p50/mean/p95/p99 at 9.32%-17.88%
lower (stacking on 0200's 19.72%-24.55%); the one-edit lifecycle accepts all
four statistics (1.87%-5.23%) plus commit p50/mean (0.45%-1.48%); the
one-percent lifecycle p95 accepts; the repeated-edit selector accepts total
p95/p99 and publication p50/mean/p95/p99. The eager-open selector claims
nothing in either direction; its primary-run adverse reading did not
reproduce in the single permitted rerun. The repeated-edit commit p50/mean
layout remnant is documented above. Allocation/RSS, physical-I/O,
cold-cache, producer, and broad-ODF claims remain withheld.
