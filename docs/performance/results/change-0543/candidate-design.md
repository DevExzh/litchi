# Change 0543 candidate: complete-result forwarding for XLSX source traversal

This is an unmeasured, isolated follow-up candidate for the selected
multi-sheet XLSX source path. It starts from the restored production source
and applies the complete 0542 candidate and its `next-candidate.patch`.
The inherited private cap-test file and test-module wiring remain part of the
coherent five-file candidate. The separate public post-EOF integration tests
are not included here; no live production source, build output, test output,
or benchmark output is part of this candidate.

The candidate keeps the conservative shared worksheet traversal introduced by
0542. `Snapshot::from_source_selected` uses it only when the original source
is valid UTF-8, within the 8 MiB source ceiling and the existing MCE input and
output limits, and contains no MCE namespace, `AlternateContent`, x14ac
namespace, or `dyDescent` marker. Every other source retains the established
validation-then-raw-parse path. The source payload, source identity and
version checks, cancellation checks, style and scalar validation, publication
rules, and worksheet relationship fence remain in their existing positions.

The shared driver gives each borrowed `quick_xml::NsReader` event to the
value-only validator before the ordinary raw parser transition. It returns a
provisional failure immediately for a reader error, observer rejection,
parser-transition error, or the 131,072-event provisional cap; the caller
drops the provisional validator and runs full validation followed by the
authoritative raw parser. The raw parser keeps its existing depth and event
limits, checked row/cell reservations, and typed allocation errors. No event
log, rewritten XML buffer, second source, public trait, unsafe block, archive
implementation dependency, or new runtime dependency is introduced.

After the observer accepts the complete source, including EOF and the
validator's root/depth condition, `Parser::finish_parse` returns an owned
`Result<Store>` in `SourceParseAttempt::Complete`. The caller consumes
`validator.finish()` before forwarding that result through
`raw::worksheet::complete_source_parse`. Thus a validation error still owns
the failure before materialization, while a post-EOF raw materialization error
can use the historical raw error and x14ac retry path without repeating full
validation and raw parsing. The late-validator cost remains an explicit risk:
validation failures found before EOF still take the immediate fallback branch.

The 0543 delta removes one redundant x14ac marker scan from the successful
completed shared path. That path is eligible only after the marker-free
predicate has already succeeded and the shared parser starts with empty
extension values, so `complete_source_parse` returns a successful owned store
without scanning the source again. For a completed parser error it still scans
for the historical marker condition and invokes x14ac capture only when the
condition is absent, preserving the established typed error precedence. The
ordinary raw facade continues to compute its extension state before parsing
and uses the same completion helper, so its behavior is unchanged.

The private `Validator` owns only bounded local-name stack entries, dialect
state, root/depth state, and one typed first error. Events and namespace views
are callback-borrowed and never escape. A provisional error drops the
validator before authoritative fallback. The complete branch necessarily
holds the validator's accepted finite state while `finish_parse` materializes
its owned result; this overlap, existing namespace/attribute allocations,
shared-formula work, and ordinary post-validation collections require fresh
allocation and invalid-input guard measurements. The finite source, event,
depth, parser-record, and scalar limits are resource bounds, not a process
memory budget or an OOM guarantee.

The candidate changes exactly these files:

- `crates/litchi-xlsx/src/cell_values/snapshot.rs`
- `crates/litchi-xlsx/src/cell_values/shared_traversal_tests.rs`
- `crates/litchi-xlsx/src/cell_values/validation.rs`
- `crates/litchi-xlsx/src/raw/worksheet/mod.rs`
- `crates/litchi-xlsx/src/raw/worksheet/codec.rs`

The generated `candidate.patch` is a full baseline-to-candidate patch for
these five files, including the inherited private cap tests. It leaves the
new public differential post-EOF `Complete(Err)` coverage to the separate
test agent. No 0542 timing, allocation, guard, or quality result is evidence
for this revision. Admission still requires the paired post-EOF error tests,
the retained early-failure/MCE/x14ac/UTF-8/cap matrix, fresh native and
allocation measurements, the unchanged invalid late-error gates, and final
quality checks before any runtime retention.

OLE2 and OOXML remain ahead of ODF in the performance program; ODF work stays
deferred until that optimization goal is complete.
