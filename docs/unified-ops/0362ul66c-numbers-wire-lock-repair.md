# 0362ul66c — Numbers-wire approved-lock repair

Status: source READY for the scoped F5 rerun.

The first actual F5 row stopped before compilation.  The approved workspace
lock `4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1`
contains the Numbers-wire package as it existed before the direct
`thiserror` dependency was added.  With `--offline --locked`, Cargo correctly
refused to mutate that byte-exact lock to add an already-resolved dependency.

The repair removes the redundant direct `thiserror` edge from
`litchi-numbers-wire/Cargo.toml` and implements the existing
`limits::Error` `Display` and `std::error::Error` contracts locally.  Error
variants, public names, messages, typed limit mapping, and all callers remain
unchanged; no compatibility alias or lockfile generation is involved.  This
keeps the approved root, perf, and native lock hashes exact while preserving
the crate's independent error type.

The changed Rust source passes `rustfmt --edition 2024 --check`.  The failed
row is preserved outside the checkout at the coordinator F5 state directory;
the scoped runner must restart from row 0 against the new clean source commit.
