# 0947 bounded source review

Reviewed 2026-10-07 against the current XML auditor and XLSX value-only
writer. This is a bounded, read-only source review. It makes no production
source, build, or benchmark change, and excludes iWork.

## Current proof boundary

`crates/xml-minifier/src/audit.rs:1003-1077` preserves the existing order:
identical bytes are audited once; otherwise common prefix and suffix are
computed, the original is audited completely, and one balanced non-root
window may prove the replacement. The initial difference-over-half gate is at
1060-1064. `WindowSearch` (`2188-2298`) tracks elements whose start is before
the first difference. `Window::prove` (`2321-2349`) rejects a candidate over
the byte limit, a candidate larger than half the replacement, a failed window
audit, counter arithmetic failure, or aggregate-limit overflow. Any missing
candidate or rejected candidate falls through to the complete replacement
audit at 1074-1076.

The public probe observed the retained twelve-pair outcomes: ten `Complete`
and two `Window`, with both sides passing. The classifications below are
source-derived from the retained bytes and those current predicates. They do
not claim that a runtime trace exposed a particular private
`WindowSearch::found` event or rejection reason.

For dense-sparse sheet 3, the first and last differing byte indices are 2755
and 5839, so the half-open differing enclosure is `[2755, 5840)` (3,085
bytes, 38.02% of the 8,115-byte replacement). The initial gate is therefore
false. The changed cells are `AG41` (`[2734,2764)`) and `BM89`
(`[5818,5848)`). No row can enclose both. The only non-root balanced element
that can cover the entire difference is the source `<sheetData>` range
`[159,8103)` (7,944 bytes, 97.89% of the replacement). Applying the current
`Window::prove` half-payload predicate classifies the complete result as a
source-derived `candidate_window_over_half` fallback.

The workbook insertion is source-derived `no_non_root_enclosure`: `calcPr` is
inserted after `</sheets>` and before `</workbook>`. The `<sheets>` element ends
before the common tail, while the document element is rejected at depth zero.

## Minimal coverage diagnostic

The diagnostic should report one row per XML Part with the following values:

```text
original_len, replacement_len, common_prefix, common_suffix
changed_cell_count, source/output span pairs, changed-byte coverage
current proof kind, current replacement audit bytes
candidate ranges/depths, eligibility, fallback reason
multi-window audit bytes, saved bytes, saved percent
original/replacement attributes, events, text bytes
removed/added counter totals, namespace and MCE boundary flags
```

For every ordered changed-cell pair, verify the source and replacement ranges
are in bounds, strictly ordered and disjoint, then compare every gap outside
the union of the ranges byte for byte. This matters when one generated cell
changes length and shifts every later output offset. Any uncovered gap,
ambiguous mapping, or unsupported boundary makes the model fall back to a
complete replacement audit.

Keep the complete original audit in every estimate. The selected replacement
work is the current window length, or the full replacement length for
`Complete`; the multi-window estimate is the sum of all eligible replacement
windows, or the full replacement length on fallback. Prefix/suffix comparison,
span validation, metadata reservation, and parser setup should be reported as
separate diagnostic work. A debug build's extra complete proof check is an
oracle and must not be counted as release scan work.

The retained span review gives these artifact-level coverage counts:

| Case | Full original bytes | Current selected replacement bytes | Hypothetical selected replacement bytes |
| --- | ---: | ---: | ---: |
| Medium 1% | 269,182 | 269,285 | 3,192 |
| Dense-sparse 1% | 508,857 | 502,623 | 5,532 |
| Medium one-edit | 63,718 | 535 | 535 |

The primary cases therefore have 266,093 and 497,091 potentially omitted
replacement bytes. These are coverage counts from source-derived spans, not
measured parser traffic, latency, allocation, or a speedup claim. The dense
model retains the existing 14-byte sheet-4 window and keeps the workbook's
complete replacement audit.

## Writer metadata and ownership

The writer has the needed source spans only transiently. `CellSlot.span` and
`CellFact.start/end` are in
`crates/litchi-xlsx/src/raw/worksheet/edit/codec/snapshot/model.rs:25-38,212-222`.
`write_replacement_row` computes the output `cell_start` immediately before
`write_cell` (`codec/snapshot/write/sheet_data.rs:355-414,709-763`). The fact
route has the same source information in `facts.rs:246-284`.

`OmittedCells` and `ValueOnlyRewrite` (`raw/worksheet/edit/package.rs:21-55`)
retain only unchanged output runs and a source pointer. They do not retain
changed source/output pairs. `Snapshot::from_rewritten_value_source`
(`cell_values/snapshot.rs:784-845`) drops facts after readback, and worksheet
publication (`1630-1675`) forwards only the detached payload. The semantic
transaction route passes `None` for facts (`workbook/edit/semantic/transaction.rs:2160-2187`),
while the cell-values source routes can receive facts (`cell_values/source.rs:827-853,1146-1209`).
Consequently, current production metadata cannot authorize a multi-window
publication proof. Any future metadata must remain private and carry source
allocation identity, lineage/version, Part identity, and freshness checks
through the operation-local audit boundary.

`crates/litchi-opc/src/xml_splice.rs:156-220,437-610` demonstrates checked
multiple source ranges, overlap rejection, and exact-gap assembly. It still
validates the assembled output completely and is not evidence that the
ordinary XLSX value-only route already has publication span provenance.

## Safety and validation requirements

Every proposed window must be a non-root balanced XML element beginning and
ending with markup. Its inherited depth and namespace bindings must match the
full original audit; its own namespace changes must close before the window
ends. Reject XML declaration, DTD/DOCTYPE, root/interstitial insertion,
unsupported namespace or MCE/`AlternateContent` shapes, mixed-content cases
whose state is not replayed, overlapping or out-of-order ranges, and any
changed bytes outside the pairs. Re-total all removed and added attributes,
events, and text bytes against the whole replacement and enforce byte,
depth, token, and aggregate limits. `State::within` and its explicit replay
fields are at `audit.rs:791-830`; the proof invariant and its future-check
obligation are at `1843-1895`; counter replacement is at `2087-2135`.

The original must still be scanned first. A failed original returns the same
`ReplacementError::Original`; any doubtful or failed replacement window must
run the complete replacement audit and return its canonical
`ReplacementError::Replacement` and offset. Freshness, cancellation, bounded
metadata reservations, quick-XML end-name checks, and the existing typed OPC
mapping at `source_backed.rs:12918-12966` remain in force. Planning-parser
success, reduced semantic readback, stopped-observer behavior, disabled
parser checks, and parsed-XML caches cannot authorize publication.

Before implementation, differential validation should cover the retained
medium/dense pairs, the one-edit control, length-changing values, multiple
windows in one row and across rows, root/interstitial insertions, malformed
original and replacement XML, namespace and MCE boundaries, all relevant
limits, stale or shifted span metadata, overlap/order errors, and allocation
refusal. Each accepted diagnostic estimate must be checked against a complete
replacement oracle; each rejected estimate must preserve the current proof
kind, typed error side, and canonical offset behavior.
