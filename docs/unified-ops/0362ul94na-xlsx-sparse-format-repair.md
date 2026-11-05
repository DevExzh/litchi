# 0362ul94na — XLSX sparse-checkout formatter correction

The 94n source repair is semantically reviewed, but its sparse author checkout
omitted `rustfmt.toml`. Default rustfmt removed two match-arm trailing commas,
contrary to the repository configuration. Root detects this before preparing
or running a new validation packet.

This follow-up restores exactly those two commas using the candidate's tracked
configuration and an explicit `--config-path`. No other source changes. The
immutable preimage and formatter proof are retained externally under
`(deleted-coord-artifact)`. Future sparse worktrees must
hydrate the formatting configuration before formatting.

The previous actual capture remains failed; fresh validation is still required.
No Cargo/build/rustdoc or performance command runs for this correction.
