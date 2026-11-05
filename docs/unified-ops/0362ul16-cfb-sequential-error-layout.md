# CFB sequential writer error layout

The unified gate stopped at `clippy-all-features` because
`SequentialWriteError::CandidateValidation` embedded the large `OleError`
directly. That made every `Result<_, SequentialWriteError>` operation carry a
large error layout and produced 45 `clippy::result_large_err` diagnostics in
`crates/litchi-cfb/src/writer/sequential.rs`.

The candidate-validation source is now boxed at its two construction sites.
The public variant still carries the complete `OleError`, progress and
`SequentialWriteReport`; only the exceptional source storage changes from an
inline value to one pointer. `Error::source()` exposes `source.as_ref()`, so
callers retain the same display text and can still downcast the error source to
`OleError`. No validation, limit, publication or durability branch changes.

The focused regression test asserts that the enum remains below the 128-byte
Clippy threshold, constructs a candidate-validation error with its complete
report and progress, verifies the source downcast, and checks the display
message. Existing writer and readback tests remain unchanged.

Validation ran from base `8f5f201712d2239dfc81cb557cd2938345130488` with the
working-tree patch that is now committed as `f9ee8613dff75df1058ee2df228ac9f84afedc8f`;
it was not a clean pre-edit `8f5f` checkout. The tested source file is now
`crates/litchi-cfb/src/writer/sequential.rs`, SHA-256
`c41bb4b3007f2b7fe9b80f777040929884c6cf88df760a3de769e8d2248318eb`. The
copied root lock is SHA-256
`4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1`.

The exact command outcomes and resource identity are recorded in the small
[validation receipt](results/0362ul16/validation-receipt.json). Every listed
command exited zero. The external Cargo target was
`(deleted-coord-artifact)`;
its post-validation `du -sb` size was 1,024,944,765 bytes (the sum of regular
file sizes was 1,182,800,909 bytes). Validation command output was sent to the
terminal rather than redirected, so no command-log path, SHA-256 or byte count
was captured; the receipt states those fields as uncaptured.

```text
cargo fmt --all --check                                      PASS
cargo test --offline --locked -p litchi-cfb --lib --tests    PASS (496 unit, all integration tests)
cargo clippy --offline --locked -p litchi-cfb --all-targets --all-features -- -D warnings PASS
cargo check --offline --locked -p litchi-doc -p litchi-xls -p litchi-ppt --all-targets --all-features PASS
cargo clippy --offline --locked -p litchi-doc -p litchi-xls -p litchi-ppt --all-targets --all-features -- -D warnings PASS
```

The checks used one Cargo job, one test thread, offline mode, the repository
memory-capped Cargo shim and the external target above. They establish build
and correctness readiness only; no ABBA or performance acceptance is claimed.
