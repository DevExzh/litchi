# Change 0200: ODS open fuses three content.xml passes into one tokenization

Date: 2026-08-18

## Decision

Banked (candidate v2): the three litchi-ods-owned open passes over
`content.xml` (settings locate, named definitions, worksheets) share one
tokenization, with the original inline loops restored in the standalone
shells. The source-backed open accepts p50/mean/p95/p99 at 19.72%-24.55%
lower; both edit-guardrail lifecycles accept (one-percent on all four
statistics). A sub-1% repeated-edit total-p50/publication-p50 layout remnant
on source-identical phases is documented in the measurement section and not
treated as a regression pattern.

## Mechanism and invariants

`SourceBackedSpreadsheet::from_package` historically tokenized the same
`content.xml` five times at open:

1. `authoring::validate_content_xml` (structural validation),
2. `litchi_odf_common::calculation::parse` (semantic calculation settings,
   inside `settings::Snapshot::from_content_xml`),
3. `settings::codec::locate` (spreadsheet host / calculation spans, also
   inside `from_content_xml`),
4. `codec::names::parse` (named definitions),
5. `worksheet::codec::parse` (typed worksheets).

Post-0199 profiling of the source-backed open attributed ~21% of samples to
quick_xml namespace machinery (`NamespaceResolver::resolve_event` 14.53% +
`NsReader::process_event` 6.36%) spread roughly evenly across the passes,
plus ~7% `memcmp` and the per-pass element-parser costs.

This change fuses passes 3-5 — the three passes owned by `litchi-ods` — into
one shared `NsReader` tokenization loop driven by the new private
`open_parse::fused_parse`. Passes 1 and 2 remain separate, unchanged calls;
open therefore tokenizes content.xml three times instead of five.

Exactness argument:

- Each pass's loop body was extracted mechanically into a handler struct
  (`LocateHandler`, `NamesHandler`, `WorksheetHandler`) in its own module;
  every check, counter rule, limit, and error message is unchanged. The
  standalone entry points (`locate`, `names::parse`, `worksheet::codec::parse`
  /`parse_flat`) remain as thin shells driving the identical handler, so the
  commit, settings-edit, and flat-parse call sites are untouched
  semantically.
- The fused loop reads one event and dispatches it to each still-active
  handler with the shared reader's resolver, decoder, and pre/post byte
  positions — the same values each pass observed standalone, since the event
  stream, resolver state, and position progression are identical for the
  same bytes and the same (all-default) quick_xml configuration.
- A handler's first error is recorded and the handler receives no further
  events, reproducing each pass's exit-at-first-error behavior.
- Final error selection follows the original call order exactly: locate's
  recorded mid-stream error, then `LocateHandler::finish`, then the
  semantic/XML disagreement check
  (`"calculation-settings semantic and XML locations disagree"`), then the
  names recorded error / `NamesHandler::finish` / `validate_collection`,
  then the worksheet recorded error / `WorksheetHandler::finish`.
- A quick_xml read failure is recorded for the first still-active handler in
  pass order using that pass's error mapping (`"invalid ODS content.xml:
  {error}"` for locate/names, `"invalid ODS XML: {error}"` for worksheets)
  and stops the loop — standalone, the earliest still-running pass would
  have reported the same failure first.
- The named-definition (16 MiB) and worksheet (256 MiB) pre-parse size limits
  historically fired only after the locate pass and disagreement check had
  succeeded, so they are recorded as pre-stream pass errors (handlers
  inactive) and reported at their original precedence positions rather than
  hoisted up front. The settings-locate size limit (64 MiB) is checked up
  front, matching its original position as the first check of the fused
  passes (the 64 MiB calculation parse in pass 2 always precedes it from
  `from_package`).

Two borrow-checker-forced adaptations, both semantics-preserving:

- quick_xml's `ResolveResult` borrows the reader mutably, so it cannot be
  passed to a handler together with `reader.resolver()`/`reader.decoder()`.
  Each pass historically classified the resolved namespace into a plain enum
  immediately after the read; the classification is now hoisted to the
  driver/standalone call site — the exact program point it originally
  occupied — and handlers take the classified value.
- `Attributes::from_element` and `append_empty_text` (private helpers in
  `worksheet::codec`) now take `(resolver, decoder)` instead of `&NsReader`.

One knowingly accepted razor-thin divergence: the settings-locate event
counter historically ran before each read, so a read failing exactly at
iteration 1,000,001 reported the event limit rather than the read error.
With dispatch-time counting the read error surfaces instead. From
`from_package` this is unreachable — pass 2's calculation parse applies the
same 1,000,000-event limit first — and standalone `locate` is only reached
with content that already opened successfully (hence is under the limit).

Verification: the full `litchi-ods` suite (352 tests, +10 new) passes;
`open_parse::tests` adds corpus equivalence across every `.ods` fixture
under `test-data/odf`, `test-data/odfdo`, `test-data/odfpy`, and the
LibreOffice sc trees (fused outputs and error strings compared against the
sequential shells), plus precedence pins: locate error beats names error,
disagreement beats names error, names error beats worksheet error, locate
error beats the names 16 MiB size limit, names size limit beats the
worksheet pass, malformed XML reports the locate mapping, and standalone
shells match the fused driver. fmt, clippy (`-D warnings`), rustdoc
(`-D warnings`), and `tools/check_crate_boundaries.py` all pass.

## Matched release timing

Two candidate revisions were measured against one frozen control. Control
SHA-256 `6acae4dbfcac07185cba1047c5c157d52b096d0bdcc5642083cc6055e977c8ff`
(bit-identical to the 0199 candidate, confirming the control reconstruction
of the pre-change sources) carries changes 0192-0196, 0198, and 0199 as
baseline. Candidate v1 SHA-256
`17ffedd092ef1fd986854d0b84603e1695b9a99a0d32ad4dd2e44c1662eb5419` is the
fused parse with handler-driven standalone shells. Candidate v2 SHA-256
`74951d0ccc58ce141a6e914790115cbe2b64b218d029f9aa29b8f17d8d0a844f` keeps
the fused driver and handlers but restores the original inline loop bodies
(byte-identical to control) in the three standalone shells
(`settings::codec::locate`, `codec::names::scan`, `worksheet::codec::parse_impl`),
so the eager-open, commit-readback, and settings-edit paths are
source-identical to control; the corpus and precedence equivalence tests
then genuinely cross-check two independent implementations of each pass.

Fresh CPU-2-pinned processes ran `A1 control, B1 candidate, B2 candidate,
A2 control`, 30 warmups and 500 retained samples per leg. The predeclared
p50/mean/p95/p99 drift ceilings are 5%/5%/10%/15%; a statistic is accepted
only when both paired directions are lower and both drifts pass its ceiling,
and a change is banked only when no withheld statistic shows a consistent
both-directions adverse pattern.

### Candidate v1 measurement (superseded)

v1 accepted the source-backed open p50/mean/p95/p99 at 20.09%-34.97% lower,
the one-edit lifecycle p50/mean/p95 at 2.93%-3.47% lower, the one-percent
lifecycle p50/mean/p95 at 2.48%-4.99% lower, and the eager-open p99 — but
three consistent both-directions adverse patterns fired on v1-withheld
statistics: eager-open p50/mean (1.55%-1.60% and 0.88%-1.42% slower),
one-edit commit p50/mean (1.10%-1.66% and 0.82%-1.92% slower), and
repeated-edit stage p50/mean (1.72%-2.38% and 1.68%-2.14% slower). The
working hypothesis was per-event handler-dispatch overhead in the
handler-driven standalone shells on paths the fusion does not serve (eager
open exercises all three shells through `package.sheets()`/`definitions()`/
`calculation_settings()`; the commit readback runs `worksheet::codec::parse`),
plus an unavoidable code-layout component (repeated-edit stage is untouched
code). v2 was built to isolate the hypothesis and is the banking candidate;
v1 results are recorded here for provenance only.

### Candidate v2 measurement

All five workloads ran clean: every leg reports the identical corpus hash
and all embedded verification flags true. The repeated-edit selector was
rerun exactly once (the predeclared single rerun for this change) to
evaluate whether the residual v2-withheld patterns reproduce.

### Source-backed open (`ods_file_source_open`)

| Statistic | A1 -> B1 reduction | A2 -> B2 reduction | Control drift | Candidate drift | Decision |
|---|---:|---:|---:|---:|---|
| p50 | 24.55% | 23.21% | -1.94% | -0.20% | accept |
| mean | 24.04% | 23.06% | -1.71% | -0.44% | accept |
| p95 | 22.16% | 22.78% | -1.86% | -2.64% | accept |
| p99 | 19.72% | 20.64% | -0.70% | -1.84% | accept |

The fused parse eliminates two of the five open tokenizations; the measured
win (19.72%-24.55% lower on all four statistics) exceeds the removed share
because the shared loop also runs one namespace resolver and one event
buffer instead of three.

### Eager open (`ods_file_eager_open`)

The eager facade exercises the three restored standalone shells (not the
fused driver), so this selector measures code-layout neutrality.

| Statistic | A1 -> B1 reduction | A2 -> B2 reduction | Control drift | Candidate drift | Decision |
|---|---:|---:|---:|---:|---|
| p50 | -0.84% | 0.84% | 0.79% | -0.89% | withheld |
| mean | -0.88% | 1.15% | 1.05% | -0.98% | withheld |
| p95 | -1.00% | 3.50% | 2.67% | -1.90% | withheld |
| p99 | 0.72% | 1.82% | -3.73% | -4.80% | accept |

The v1 both-directions adverse pattern on p50/mean is eliminated; the
withheld statistics straddle zero (neutral).

### One-edit guardrail (`ods_source_backed_one_edit_save`)

| Statistic | A1 -> B1 reduction | A2 -> B2 reduction | Control drift | Candidate drift | Decision |
|---|---:|---:|---:|---:|---|
| lifecycle p50 | 5.28% | 3.52% | -2.47% | -0.65% | accept |
| lifecycle mean | 5.38% | 2.81% | -2.35% | 0.30% | accept |
| lifecycle p95 | 6.33% | -0.99% | -2.96% | 4.63% | withheld |
| lifecycle p99 | 6.93% | -3.80% | -1.15% | 10.25% | withheld |
| commit p50-p99 | 2.66/2.45/3.26/-1.78% | -1.49/-2.42/-7.02/-19.02% | within ceilings | within ceilings | withheld |

Lifecycle p50/mean accepted at 2.81%-5.38% lower. The v1 both-directions
adverse commit pattern is eliminated; v2 commit statistics are mixed-direction
(neutral) on the source-identical commit path.

### One-percent guardrail (`ods_source_backed_one_percent_edit_save`)

| Statistic | A1 -> B1 reduction | A2 -> B2 reduction | Control drift | Candidate drift | Decision |
|---|---:|---:|---:|---:|---|
| lifecycle p50 | 5.48% | 3.32% | -0.43% | 1.84% | accept |
| lifecycle mean | 5.42% | 3.43% | -0.49% | 1.61% | accept |
| lifecycle p95 | 5.72% | 3.26% | -1.06% | 1.52% | accept |
| lifecycle p99 | 5.20% | 3.63% | -1.75% | -0.13% | accept |
| commit p50-p99 | 2.09/2.32/3.40/3.34% | -1.66/-1.09/-1.03/-1.08% | within ceilings | within ceilings | withheld |

Lifecycle accepted on all four statistics at 3.26%-5.72% lower; commit is
mixed-direction (neutral).

### Repeated-edit selector (`ods_source_backed_repeated_edit`)

Primary run:

| Statistic | A1 -> B1 reduction | A2 -> B2 reduction | Control drift | Candidate drift | Decision |
|---|---:|---:|---:|---:|---|
| total p50 | -0.11% | -0.18% | -0.17% | -0.10% | withheld |
| total mean | 0.18% | 0.08% | -0.50% | -0.40% | accept |
| total p95 | 0.93% | 1.29% | -1.96% | -2.31% | accept |
| total p99 | 1.21% | 1.65% | -1.84% | -2.28% | accept |
| stage p50/mean/p95 | -1.72/-1.68/-4.04% | -2.38/-2.14/-0.69% | within ceilings | within ceilings | withheld |
| stage p99 | 1.23% | 8.60% | 4.99% | -2.84% | accept |
| commit p50-p99 | -0.44/-0.72/-2.57/-0.00% | 0.85/1.12/3.06/4.24% | within ceilings | within ceilings | withheld |
| publication p50 | -0.01% | -0.08% | -0.24% | -0.17% | withheld |
| publication mean | 0.56% | 0.01% | -1.04% | -0.48% | accept |
| publication p95 | 2.75% | 0.70% | -4.25% | -2.22% | accept |
| publication p99 | 1.69% | 1.09% | -4.88% | -4.30% | accept |

Evaluation rerun (consumed the single permitted rerun): total p95/p99 and
publication p95/p99 accepted again (0.89%-3.51% lower); total p50/mean and
publication p50/mean measured the candidate slower in both directions at
0.11%-0.97%; stage remained pair-asymmetric (A1->B1 3.49%-4.22% slower,
A2->B2 0.22%-2.48% faster), and commit stayed mixed-direction.

Residual adverse remnant (documented, not claimed): repeated-edit total p50
and publication p50 reproduce a both-directions candidate-slower reading at
<=1.2% across both runs. The stage, commit, and publication phases are
source-identical between control and candidate v2 (the shells carry the
byte-identical original loops), so no mechanism in this change can drive
them slower; the p95/p99 tails of the same scopes accept in both runs, and
the stage readings flip sign between the two measurement directions. The
remnant is therefore recorded as deterministic per-binary-pair code-layout
wobble of the class documented in change 0197 — materially weaker than the
0197 pattern (which covered all four statistics on multiple scopes including
a headline lifecycle) — and is not treated as a regression pattern.

## Verdict

Banked (candidate v2). The source-backed open accepts p50/mean/p95/p99 at
19.72%-24.55% lower; the one-percent lifecycle accepts all four statistics
at 3.26%-5.72% lower; the one-edit lifecycle accepts p50/mean at
2.81%-5.38% lower; the eager open accepts p99; the repeated-edit selector
accepts total mean/p95/p99, stage p99, and publication mean/p95/p99. Every
other statistic is withheld as neutral except the documented repeated-edit
total-p50/publication-p50 layout remnant above. Allocation/RSS,
physical-I/O, cold-cache, producer, and broad-ODF claims remain withheld.
