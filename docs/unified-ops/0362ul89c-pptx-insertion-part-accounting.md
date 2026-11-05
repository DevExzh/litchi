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
two dependency-free replacements in the same topology revision, serializes
and reopens the result, and checks the final catalog count.  This exact-bound
case is rejected by the former slide-only estimate on the second insertion.
The reopened final catalog is then tested at its exact original ceiling and a
third insertion is refused atomically.  A policy below the original physical
catalog is not used for ingress: source admission itself must accept every
original Part.

No Cargo, native tools, or performance capture were run in this source-only
return.

## Typed composition

The coordinator composes final 89d `7fb94ce6c052c173dfb8cb012c62d9c9882f9287`
after typed source merge `49f9b2fcac`. The bounded map remains proportional
to topology changes and dependents; no source-wide name set is restored.
Physical source-Part admission reports `Codec(Parts)`, distinct from the
opened edit's changed-Part budget, with the prospective count and configured
source maximum. Checked additions report true overflow; impossible removals
remain patch conflicts. The regression expects three slides after removing
one of two original slides and inserting two, and checks the exact observed
count on the next refusal. This source composition has not run under Cargo.
