# Change 0360: bounded streaming MCE events

**Date:** 2026-09-01
**Status:** Implemented
**Performance claim:** none

## Decision

Add a new callback-scoped streaming Markup Compatibility and Extensibility
processor in `litchi-ooxml-common::mce`. The legacy byte-buffer processor is
unchanged. This is the semantic event foundation required before worksheet
MCE/x14ac processing can stop materializing rewritten XML.

The canonical API has independent raw-element and active-semantic observers.
The raw observer sees every syntactically valid `Start` and `Empty` element
after its namespace declarations resolve but before MCE filtering. The active
observer sees only the selected, visible event stream. Both surfaces borrow
bounded per-event scratch and cannot escape their HRTB callbacks.

## Processing contract

The stream processor preserves namespace scopes, unqualified-attribute rules,
`Ignorable`, `MustUnderstand`, `ProcessContent`, `PreserveElements`,
`PreserveAttributes`, `AlternateContent` Choice/Fallback selection, and opaque
extension preservation. Inactive branches are still parsed and structurally
validated rather than skipped with `read_to_end`.

Raw and active callbacks disable independently after their first error. The
processor continues to XML EOF, retaining both typed callback errors. Input,
XML, MCE, and limit failures are primary over callback errors. Callback panics
unwind normally and do not make a completion claim.

The new stream contract is deliberately stricter than the legacy borrowed
fast path. It requires matching end names and one complete root, rejects
non-whitespace content outside the root, validates custom references in hidden
branches, restricts XML declarations to the prolog, and rejects direct CDATA
inside `AlternateContent`.

## Resource boundary

`StreamLimits` adds finite event-count, per-event raw-byte, attribute,
attribute-byte, context-byte, and name-byte ceilings to the existing MCE
processing limits. The per-event meter bounds bytes consumed by each pinned
quick-XML `read_event_into` call, including delimiter lookahead. A fixed
three-byte prefix stage recognizes a split UTF-8 BOM without weakening the
event ceiling or double-charging total input. `Interrupted` retries are finite,
and invalid `BufRead::consume` use becomes a typed deferred error.

This is a bounded streaming design, not a fixed-memory or OOM-safe claim.
quick-XML internal parser state, decoded value allocation, collection overhead,
and callback-owned allocations are not fully represented by the context-byte
estimate. The API therefore makes no input-size-independent memory claim.

## Validation

Validation was deliberately serial and crate-scoped:

- focused streaming MCE tests: `11/11`;
- `litchi-ooxml-common` library tests: `223/223`;
- existing markup-compatibility integration: `1/1`.

The first focused compile exposed six lifetime/borrow and denied-pattern errors,
which were corrected before tests ran. Subsequent focused runs exposed valid
meter consumption across quick-XML event boundaries, the exact 129-byte text
event boundary including delimiter lookahead, and preserved-extension
classification; all were corrected before the passing run.

Every command used one Cargo process, `CARGO_BUILD_JOBS=1`, one test thread, an
8 GiB process ceiling, disabled incremental/debug compilation, and one on-disk
`change-0360` target. Its observed final/peak footprint was 267 MiB. Host
availability was approximately 14 GiB with 134 GiB disk free and exhausted
swap. No parallel build or OOM occurred. These are validation observations
only; `performance_claim: none`.

## Residual scope

XLSX does not consume the new event stream yet. The next batch must implement
the x14ac raw/active observer and worksheet event adapter while preserving
x14ac raw-marker precedence. Selected-cell scanning, bounded shared-string and
style lookup, and performance/resource measurement remain later work.
