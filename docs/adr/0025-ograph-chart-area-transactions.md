# ADR 0025: OGraph chart-area snapshot transactions

Status: accepted

## Context

`[MS-OGRAPH]` section 2.4.21 defines `Chart` as one fixed 16-byte record:
`x`, `y`, `dx`, and `dy` are signed 16.16 fixed-point values. The current
`litchi-ograph::chart::Chart` model already decodes this record into `Rect`,
but its snapshot editor only replaced existing cache-cell payloads. Directly
mutating `Chart::set_rect` on a parsed snapshot correctly invalidated exact
replay, yet did not provide a safe, reversible edit path.

## Decision

Add `chart::transaction::chart_area`, a nested model/codec/validation/test
owner for a bounded chart-area edit. `Editor::set_rect` accepts only a zero
origin and nonnegative width and height, matching the normative `Chart`
constraints. Commit locates exactly one source `Chart` record, checks that its
decoded value still equals the source snapshot, and overwrites only its fixed
16-byte payload. The record kind, length, offset, ordering, and all unknown
records remain untouched.

`Patch` carries the optional chart-area change alongside cache changes. Its
inverse swaps the checked before/after rectangles, and applying it requires the
same source rectangle. Source validation occurs before physical publication;
invalid or conflicting edits return a typed error and produce no commit.

## Scope

This owner edits chart-area metadata only. It does not render, evaluate, resize
host objects, rewrite package relationships, or interpret unknown records.
Untouched parsed charts continue to replay their original allocation exactly.

## Verification

Focused tests cover fixed-record-only mutation, unknown-record preservation,
exact inverse replay, invalid semantic values, source conflicts, and no-op
publication. The implementation relies on the fixed `Chart` grammar in
`3rdparty/specs/[MS-OGRAPH]/2 Structures/2.4 Records.md` section 2.4.21.
