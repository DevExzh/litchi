# 0362ul94n — XLSX metadata clippy repair

This source-only follow-up starts from the committed 94m tip
`305a7c0ee6d7b84680232a49d4eb92392dce8fe1` (tree
`663a875a5ab13eccc96f0c71753bbe610524513a`).  The strict common/OPC clippy
rows passed; the leaf row found two literal `ok_or_else` closures in the XLSX
metadata admission path.  This repair makes no acceptance claim and does not
rerun the gate.

Both closures now use `ok_or` with the same fixed `Error::Internal` values.
The error variants, messages, evaluation order, and no-allocation source path
are unchanged.

Only rustfmt, diff, and repository size/binary guards are run for this repair;
Cargo, rustdoc, compiler, native, and performance execution remain pending.
