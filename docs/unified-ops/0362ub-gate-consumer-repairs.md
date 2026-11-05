# 0362ub — Repair the detector example and resave consumer lock

The resumed gate on `ad1c6cede8` ran 52 checks: 48 passed, while ops,
facade, workspace tests and native-resave failed. The external summary and
failed logs are retained by size and SHA-256 in
[`results/0362ub/validation.json`](results/0362ub/validation.json), along with
the passing workspace documentation log. This records the actual gate
script, including the full ODS suite and iWork public-API tool unit tests;
it is not the final gate on the current assembly.

The detector doctest now imports its required `SourceLimits` type. The
native-resave lockfile now includes the ops SHA-256 dependency introduced by
content-bound replay. Offline Cargo metadata regeneration changes exactly
one dependency line, with no dependency version updates. Both corrections
are consumer migrations for the assembled API and change no runtime path.

On the combined `3c1d25de24` source with these corrections, all 20 facade
doctests pass. Both `cargo check --offline --locked --all-targets` commands
pass for `tools/native-resave/Cargo.toml` and
`tools/perf-baseline/Cargo.toml`. Full commands, elapsed times and log hashes
are in the validation receipt. No performance measurement applies to the
documentation import or lock dependency list.

The ops RTF consumers and XLSB refusal classifications are repaired by
`0641d926aa` and `759b09f0d9`; their focused receipts remain in 0362u/0362ua.
The ZIP prefill refusal repair is separately validated in 0362x and still
requires integration. Remaining raw-policy and performance work, and a
fresh full gate on the final combined source, stay required.
