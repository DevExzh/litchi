# 0729 probe review: public DOC lifecycle attribution

Status: read-only probe review, 2026-09-22. This review covers the phase-route
implementation appended to
[`probe/src/lib.rs`](probe/src/lib.rs), the current analyzer, and the
inherited oracle prefix. No build, native command, or production edit was
performed. The review does not accept a performance result.

## Scope and route equivalence

The four routes use the same source bytes, default DOC limits, replacement
text, public paragraph operation, output extraction, and owner-drop shape:

| Route | Public operation | Diagnostic work | Intended comparison |
| --- | --- | --- | --- |
| `ordinary-opaque` | ordinary `Snapshot::open` → `edit` → replacement → `commit` → output copy | none | whole-lifecycle baseline |
| `ordinary-split` | same ordinary operation | external `Instant` windows around open, edit construction, replacement, commit, and output copy | timer/split overhead and boundary residual |
| `profiled-empty` | same public operation through profiled open/commit APIs | empty semantic observers; no trace retained | profiled implementation control |
| `profiled-clock` | same public operation through profiled open/commit APIs | fixed-capacity timestamp collector | named semantic phase attribution |

The route functions now use matching inner scopes. The `commit` value is
dropped after the final output copy, and the source snapshot is dropped before
the outer whole timer ends, on every route. The final output `Vec` is moved
out of the timed scope and remains available for untimed hash, inventory, and
oracle validation. Public argument construction and oracle work are outside
the lifecycle timer. This is the correct ownership placement for a whole
public lifecycle measurement.

The profiled routes still have the production implementation difference
recorded in the source review: ordinary open keeps `_strict_editor` alive
through public-reader validation and source retention, while profiled open
drops it at the end of the strict-owner closure. Therefore
`ordinary-split` → `profiled-empty` measures that lifetime/code-path
difference as well as the empty observer route. Equal outer scopes do not make
ordinary and profiled phase timings a pure callback comparison.

`profiled-empty` is an empty-observer control, not an uninstrumented ordinary
control. Its callback may be optimized to little or no work in a release build;
the route remains useful for checking the profiled API path and its lifetime
behavior, but its whole-time delta must not be called observer dispatch cost
alone. The intended whole-time comparisons are:

* ordinary opaque → ordinary split: external timer and split-window cost;
* ordinary split → profiled empty: profiled implementation and lifetime
  difference, including empty observer calls; and
* profiled empty → profiled clock: fixed timestamp collector work on the
  profiled path.

None of these comparisons is a production speedup claim.

The compatibility `phase_ns` and `public_phase_ns` projections combine edit
construction and replacement under a `stage_ns` label. They are not required
for the new attribution and must not be used to claim a common-container
stage. They are slated for removal after the current build window; the
authoritative fields for this probe are `whole_ns`, `split`, and
`diagnostics` (plus the separate observer calibration field). The analyzer
should not require or rank the compatibility projections.

## Fixed timestamp collector

`ClockTrace` is bounded and allocation-free during the timed lifecycle. Each
trace owns a `[Option<RawPhaseEvent>; 32]` array, a length, and an overflow
flag. The changed DOC route expects six events for open and ten for commit,
so the capacity has margin. Every callback records:

* started/finished kind;
* the numeric semantic phase;
* success/error outcome; and
* `Instant` elapsed nanoseconds from the outer lifecycle start.

`trace_report` runs after the whole timer ends. It allocates the serialized
event and span vectors only then, checks overflow, stack balance, exact event
count, fixed phase order, successful outcomes, and one span per expected
phase. The analyzer additionally checks that every event timestamp is
monotonic across open and commit, lies inside `whole_ns`, and that the
attributed phase durations are bounded by their corresponding external open
or commit window. That is the right separation: fixed recording is timed,
event formatting and validation are untimed.

The collector records `DiagnosticEvent` only. The phase route calls
`open_bounded_profiled` and `commit_profiled`, whose CFB observers are the
library's empty defaults; it makes no claim about `CfbParseEvent` cardinality.
CFB event correctness remains covered by the format-level tests and the
source contract. Adding CFB timestamps would be a separate nested-observer
measurement and must not be added to the non-overlapping phase sum.

The collector's success contract is appropriate for the measured changed
route. It does not exercise an error or exact-no-op route. Those event
contracts remain inherited library-test evidence; the probe's timed samples
must not be presented as error/no-op coverage.

## Boundary residual and nested fractions

`split_timing` measures five sequential outer windows and reports:

```text
split_sum_ns = open_ns + edit_ns + replace_ns + commit_ns + output_copy_ns
whole_residual_ns = whole_ns - split_sum_ns
```

The residual is intentionally `i128`: independent `Instant` calls at every
boundary can make the sum exceed the outer interval for a small case. A
negative residual is a boundary-accounting observation, not a phase failure.
The current analyzer has a qualification bug here:

```python
assert all(type(v)==int and v>=0 for v in split.values())
```

That assertion includes `whole_residual_ns` and would reject the negative
residual that the Rust schema deliberately permits. It should require
nonnegative integer values only for the five duration fields and
`split_sum_ns`, require an integer (signed) residual, and then check
`whole_residual_ns == whole_ns - split_sum_ns`. The residual may produce a
negative percentage and should remain a separately labelled boundary value,
never be assigned to a format phase.

For `profiled-clock`, the analyzer computes each semantic span as a fraction
of the same sample's `whole_ns` and of its containing external `open_ns` or
`commit_ns` window. It checks the sum of spans for each owner is no greater
than that owner window, which prevents nested phases from being counted as
extra lifecycle time. CFB work, if later observed, would be nested inside
strict-owner validation and could not be added to this sum.

The parent fraction calculation should also guard against a zero
`open_ns`/`commit_ns` denominator before division. The current DOC cases are
expected to have nonzero owner windows, but the schema should fail clearly or
emit an unavailable fraction rather than raise a divide-by-zero exception.
The data currently checks global event monotonicity and owner-span totals;
direct timestamp containment against the exact external window endpoints is
inferred from the route scopes rather than encoded as endpoint timestamps.

The analyzer does not need to subtract either residual or nested fractions
from `whole_ns`. The whole timer is authoritative for lifecycle time; split
windows and semantic spans are descriptions with checked containment.

## Observer calibration field

`observer_clock_control_ns` is measured only in `profiled-clock`, after the
timed lifecycle has ended and after `whole_ns` is captured. It runs the same
fixed recorder over the expected six open plus ten changed-commit semantic
events. The field is emitted as an untimed calibration control. The analyzer
may report its distribution, but it must never subtract it from `whole_ns`, a
split window, or a semantic phase duration.

The calibration is not an exact cost model for the library callback: it
manually constructs the expected events and may be optimized differently from
the callback path. In particular, only `trace.len` is passed through
`black_box`; the recorded event-array writes are not otherwise consumed. A
release optimizer could eliminate dead array stores in the calibration while
the real callback's trace is later serialized, making the calibration a lower
bound or a noisy control. If this field is used as evidence of recorder cost,
the calibration should keep the event storage observable (for example by
passing the trace or a digest through `black_box`) while remaining outside the
workflow timer. It still must not become a subtraction.

## Profiled event validation and output checks

For every `profiled-clock` sample, the expected changed route is:

```text
open:   StrictOwnerValidation → PublicReaderValidation → SourceRetention
commit: Finish → StrictOwnerValidation → PublicReaderValidation
        → SourceRetention → Patch
```

The fixed collector and analyzer require each phase to start and finish once,
in this order, with success, and require balanced nesting. The open CFB parse
is not in the trace; no CFB fraction is reported. `profiled-empty` explicitly
requires no diagnostics and no calibration field, so it remains an honest
whole-time control rather than an empty attribution record.

The sample output is retained after timing and passed through the inherited
DOC oracle. The route validates output hash and stream inventory against the
expected public edit, then checks the semantic witness, changed UTF-16
length, untouched streams, directory metadata and CLSID policy, and the
negative controls. The expected output is generated once before timing, and
the probe does not call the oracle or inventory while the lifecycle timer is
running.

The new 0729 source prefix through the inherited `oracle_controls` function
is append-only relative to the 0728 probe. The raw directory oracle,
`oracle_for_output`, `changed_length_proof`, replacement derivation, semantic
witness, and negative controls are unchanged. `run_phase` invokes those same
functions for its expected output and each sample, while
`prior-oracle-contract.json`/qualification is intended to bind the witness
and case identities to the accepted 0728 contract. No new attribution code
should weaken that inherited oracle or replace it with event validation.

## Qualification disposition

The timed route structure is suitable after the signed-residual analyzer check
is corrected and compatibility aliases are removed or excluded from the
authoritative schema. Before freeze, retain the following interpretation:

* whole lifecycle and owner drops are timed consistently across routes;
* output ownership, inventory, semantic readback, and oracle work are outside
  the timed lifecycle;
* semantic diagnostic events are fixed-capacity, validated, and timestamped;
* `observer_clock_control_ns` is an untimed control and never a subtraction;
* nested phase fractions are bounded by the same-sample owner windows and CFB
  evidence is not double-counted; and
* ordinary/profiled lifetime differences remain a stated limitation.

This probe can produce descriptive public-DOC phase evidence once those
checks pass. It cannot by itself justify a production optimization,
allocation claim, or validated-render handoff.

## Resolution after final route review

The earlier signed-residual concern is resolved in favor of fail-closed
qualification. These five split windows are sequential, monotonic intervals
started after `whole_start` and finished before `whole_end`; the owner drops
also occur before `whole_end`. Their sum cannot exceed the enclosing whole
interval under this route. Although the Rust field remains signed as a
diagnostic representation, a negative `whole_residual_ns` is invalid evidence
for this probe and must fail qualification. The analyzer's nonnegative check
is therefore correct, and it now explicitly requires positive `open_ns` and
`commit_ns` before computing nested fractions.

The final route removes the compatibility phase aliases. The calibration now
passes the complete fixed recorder through `black_box` before its untimed
control ends, so the earlier dead-store-elimination concern is closed. The
observer calibration remains a reported control and is still never subtracted
from lifecycle or phase timings.

The final build, four quality checks, eight qualification routes, synthetic
analyzer/auditor integration, and inherited 0728 oracle checks are reported by
the packet owner as passing. Static review finds no further substantive issue
in the finalized attribution route. The ordinary/profiled lifetime
qualification and the absence of timed error/no-op samples remain documented
interpretation limits, not implementation failures.
