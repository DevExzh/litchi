# 0543 provisional event-cap boundary review

This is a static control-flow review of the 0543 candidate's
`MAX_SHARED_PROVISIONAL_EVENTS = 131_072` boundary. It does not change Rust
sources or run a build, test, or measurement. The coordinator is running a
separate fresh native guard with valid worksheet shapes at 160 × 160 (below
the cap), 164 × 164 (just above the cap), and 256 × 256 (above the cap). Those
fixtures remain below the 8 MiB shared-source ceiling and below the ordinary
one-million-event worksheet safety threshold; their receipts are the authority
for the retention decision.

## Current control flow

`source_stream_eligible` admits these plain UTF-8 sources because eligibility
is based on source bytes and marker absence, not on an event-count estimate.
`parse_source_with_observer` increments the event count for every reader event,
including EOF. It resolves the event and calls the validator before checking
the cap result. On event 131,073 it therefore performs one last bounded
validator callback and then returns `ProvisionalFailed` before the raw parser
transitions that event.

The caller drops both the reader and provisional parser state, drops the
validator, and runs the established full validation pass followed by the
authoritative raw parser. A valid source just above the cap consequently pays
the shared validator/parser work through the boundary and then two complete
authoritative traversals. Relative to the restored baseline's two complete
traversals, this is baseline work plus a discarded parser/validator prefix.
The 160 × 160 case exercises the one-traversal shared completion; the 164 ×
164 and 256 × 256 cases exercise the fallback cliff at different distances.

This branch is semantically safe. The fallback owns the final result and still
decides validation, parser, x14ac, source, and publication errors. The cap does
not publish the provisional store, and the unprocessed over-limit event is
re-read by the authoritative passes. The concern is valid-input performance:
the main 96 × 96 and 128 × 128 pilot shapes stay below the boundary and cannot
detect this cliff. The inherited cap tests use repeated comments to prove
fallback correctness; they do not measure a valid parser with a large cell
record state. The supplemental guard is therefore an admission control, not a
semantic regression test.

## Resource and error implications

The cap still bounds the speculative parser's event-driven state before a
complete validation EOF. At the boundary, the validator callback can inspect
the over-limit event before the cap branch returns, but its grammar stack is
bounded by the value-only worksheet grammar and the source itself is capped.
No event or namespace borrow escapes. The parser state and validator are
dropped before authoritative fallback, so the cap branch does not create a
new long-lived overlap.

All early cap outcomes retain the established error precedence. A validation
error on or after the boundary causes fallback full validation to own the
diagnostic; a parser transition error that would have occurred on the
unprocessed event is likewise rediscovered by the authoritative raw parser.
Only a complete valid document whose event count is at most 131,072 including
EOF avoids this branch; a complete valid document above the cap takes the
fallback described above. No x14ac or MCE ordering changes are implied by the
cap.

## Safe follow-up designs

The simplest way to eliminate the valid cap cliff is a conservative admission
precheck based on an input-byte upper bound. Every non-EOF quick-xml event
consumes source bytes, with EOF adding at most one event, so a formally
documented and tested bound of `content.len() + 1 <=
MAX_SHARED_PROVISIONAL_EVENTS` is sufficient to prove that the event cap cannot
trip. Sources outside that bound should bypass the shared reader before any
provisional parser allocation and take the established validation/raw path.
This preserves semantics and never adds a discarded prefix. The tradeoff is
substantial: XML cell records contain many bytes per event, so this byte bound
would conservatively reject sources such as ordinary 160 × 160 worksheets even
when their actual event count is below the cap. It should therefore be measured
as a separate candidate rather than silently folded into 0543. The invariant
must be backed by a quick-xml event-consumption test; an unchecked assumption
about event splitting is not a safe bound.

A less conservative follow-up can keep one validation pass after the cap while
discarding the parser immediately. The reader would continue only the bounded
value-only validator to EOF, then run the authoritative raw parser once from
the original source. On a valid source this costs one validator pass plus one
raw pass, with only the already-discarded parser prefix as overhead, instead of
the current prefix plus a second full validator and raw pass. It requires an
explicit decision that the completed shared validator owns a fully observed
validation result, or a second full validator on validation failure to retain
the current diagnostic contract. The latter preserves error ownership but
reduces the benefit for late invalid inputs. The continuation must never retain
the parser through EOF or materialize cells before validation accepts EOF; doing
so would defeat the provisional-state bound and validation-first rule.

Counting all events in a separate pre-scan would avoid the fallback cliff but
would add another source traversal to every eligible read and likely erase the
successful-path gain. Raising the cap toward the ordinary one-million-event
limit has the opposite risk: it postpones the cliff while increasing parser
state retained before validation, with no memory-budget proof. Neither is a
free 0543 fix. A future revision should compare the byte precheck, a tightly
bounded structural estimator, and the validator-only continuation against the
same valid, late-validator, and late-raw guards.

## Admission status

There is no new semantic source blocker in the current candidate. The
above-cap valid path is a concrete performance hazard and remains unresolved
until the coordinator's 160 × 160, 164 × 164, and 256 × 256 baseline/candidate
measurements are reviewed. Retention requires those results to show that the
boundary does not violate the unchanged valid workflow/planning and allocation
gates, with every cap-boundary adverse record individually classified. Even if
the supplemental guard passes, the byte-precheck or bounded-continuation
design should be carried as the next optimization if the valid cap cliff is
material.

The cap review remains within the OLE2/OOXML priority; ODF optimization stays
deferred until that goal is complete.
