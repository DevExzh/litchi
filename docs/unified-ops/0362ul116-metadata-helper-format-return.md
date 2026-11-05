# 0362ul116 — metadata helper formatting return

The fresh joint compiler packet on `fcc049963fde2cfd3c9d285264bfca476f3e93e2`
stopped at its first row: workspace rustfmt requires the new facade metadata
helper's signature on one line. This return changes only that whitespace;
property context, refusal classification and consumer behavior are unchanged.

The coordinator verified the terminal failed capture before editing source.
No Cargo row ran. The external packet is
`(deleted-coord-artifact)` under the
worktree parent. Its `rows.json` is 34,428 bytes, SHA-256
`8019e7e80d6d69a663909c6c5335835f7bd6fec7dd4813a0b36b3fab41021788`;
`result.json` is 36,143 bytes, SHA-256
`b767ec7afc51d04aa4bdb5590f246f65a3facddb165005f11f19005ba0f39707`.
The failed row took 4.839 seconds and left the target empty.

Validation: the corrected file passes the pinned workspace rustfmt command
with `--edition 2024 --config skip_children=true --check`; `git diff --check`
and commit size/binary guards pass. Fresh compiler, runtime and glossary
packets must bind the resulting source commit. This is not runtime or final
acceptance.
