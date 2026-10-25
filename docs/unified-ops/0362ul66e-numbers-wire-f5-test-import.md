# 0362ul66e — Numbers-wire F5 test import repair

Status: source READY for the scoped F5 retry.

The all-target Numbers-wire check reached test compilation and found one
missing test-module import: the existing limit-boundary assertion used
`WriteError` without importing the public table-writer error into its local
module.  The test now imports that existing type; no production behavior,
wire bytes, or limit policy changes.

The failed row remains preserved in the external F5 state directory and no
later row was started.  Rustfmt and source diff checks pass.
