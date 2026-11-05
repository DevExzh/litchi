# 0362ul94m — Common clippy repair

This source-only follow-up starts from the committed 94l tip
`53761f3bb4a66162618fdc2bfbc7afbeca80833e` (tree
`c03f78ccced200778935e38e33b4d6dba9d9acf7`).  The private-v2 common clippy
row found sixteen diagnostics in the properties transaction, VBA, and Web
Extensions planning code.  This repair makes no acceptance claim and does not
rerun the gate.

The changes replace literal `ok_or_else` closures with `ok_or`, remove one
needless `Ok(...?)` wrapper, elide an unnecessary iterator lifetime, and use
`contains` for the existing target-name membership check.  Error variants,
fixed messages, evaluation order, and allocation behavior are unchanged.

Only rustfmt, diff, and repository size/binary guards are run for this repair;
Cargo, rustdoc, compiler, native, and performance execution remain pending.
