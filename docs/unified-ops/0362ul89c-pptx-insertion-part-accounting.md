# 0362ul89c — PPTX insertion catalog accounting

This return closes the source-native insertion part-boundary gap on the
combined 0362 tree.

## Exact effective catalog

Insertion admission now computes the effective materialized OPC Part set from
the retained physical catalog plus both topology namespaces: slide changes and
exclusive dependent-Part changes.  Names are folded with the OPC
case-insensitive rule, removals are applied before additions, and overlapping
entries count once with the final restored image present.  A missing physical
`before` witness still returns `PatchConflict`; duplicate witnesses do not
underflow the count.  The check remains before authored slide/output
allocation and reserves the bounded name indexes fallibly.

The common no-topology path keeps its existing metadata-only count.  Structural
removal followed by insertion therefore admits the actual final catalog rather
than charging an exclusive notes/media Part that has already left the output.

## Boundary coverage

The source integration regression builds a slide with an exclusive notes Part,
opens it under the exact original Part ceiling, removes that slide and inserts
a dependency-free replacement, serializes and reopens the result, and checks
the final catalog count.  It then reopens under the adjacent exact final
ceiling and verifies the next insertion is refused atomically.  A policy below
the original physical catalog is not used for ingress: source admission itself
must accept every original Part, so the tighter adjacent policy is selected
only after the first output is reopened.

No Cargo, native tools, or performance capture were run in this source-only
return.
