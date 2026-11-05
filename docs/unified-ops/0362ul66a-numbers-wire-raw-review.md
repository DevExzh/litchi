# 0362ul66a — Numbers-wire raw policy review return

Parent: [0362ul66](0362ul66-numbers-wire-raw.md).

The independent review return closes the remaining raw-policy and consumer
gaps without a compatibility alias. Every tracked Numbers, iWork and fuzz
consumer now passes the selected typed `litchi_numbers_wire::Limits` profile
to BNC parsing, value decoding and encoding. The borrowed BNC rewrite and
clear operations use their canonical names (`rewrite_scalar`, `clear_value`,
`clear_comment`, `rewrite_formula_cache`, `rewrite_formula`,
`rewrite_formula_without_cache`, and
`rewrite_duration_format_identifier`) rather than a public `*_with_limit`
surface.

The raw table reader independently admits tile references, rows, present
cells, tile payload bytes, and row storage/offset spans before the owner's
aggregate budget. The owner budget remains separate and may be tighter. Pages
section limits, Keynote reference limits, and Numbers materialized-cell
limits are no longer mapped onto the raw repeated-item ceiling, so a semantic
cell budget cannot reject a valid wider table. Formula-state scans admit the
owner count before either pass and charge decoded packed values, including a
single ten-byte varint, as one item.

The shared field rebuild helper stages mapped replacements, preflights the
actual encoded total, and makes one bounded exact final reservation. Other
writers retain their existing bounded, fallible growth. Owned BNC encoding now
also checks field count and fixed widths before allocation; generated rewrite
readback uses the selected output profile without weakening untrusted source
ingress. Unknown fields and opaque tails remain byte-for-byte preserved.

Source-only regressions cover independent raw input and item caps with a
permissive owner budget, exact/one-under admission, packed-varint accounting,
formula-owner admission, and existing unknown-tail fidelity. Cargo, native and
rustdoc validation remain deferred to the coordinator's explicit heavy-build
slot; no performance claim is made by this return.
