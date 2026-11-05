# 0362ul94l — Common OPC clippy repair

This source-only follow-up starts from the committed 94k leaf tip
`5e8940041389d3b129f8b73cd6ede252bb6254be` (tree
`84130fd0c8f6c695031c65372f226dc036dc4bf0`).  The common and OPC compiler
rows and all six format/facade compiler rows passed; the private-v2 clippy
row found two dependency diagnostics.  This repair makes no acceptance claim
and does not rerun the gate.

The OPC package writer now uses `Relationships::is_empty()` for its generated
part-relationship test, and deferred payload comparison uses
`io::Error::other` for its fixed overflow diagnostic.  No behavior, error
classification, or public API changes are intended.

Only rustfmt, diff, and repository size/binary guards are run for this repair;
Cargo, rustdoc, compiler, native, and performance execution remain pending.
