# `litchi-numbers-wire` fuzzing

This is a standalone [`cargo-fuzz`](https://github.com/rust-fuzz/cargo-fuzz)
workspace for the borrowed legacy Numbers pre-BNC cell view. The target checks
that parsing never changes the caller-owned bytes, that every returned slice is
source-backed, and that known scalar and identifier fields retain their typed
meaning across versions zero through four. It also exercises truncation,
unsupported versions, unknown flags, opaque tails, and non-finite scalar
rejection.

The target has only two dependencies: `libfuzzer-sys` and the path dependency
on `litchi-numbers-wire`. It intentionally does not depend on the package or
editor crates, so malformed wire input cannot accidentally enter a higher
level compatibility path.

Install `cargo-fuzz` and a nightly toolchain, then run it from the repository root.
The command copies the small checked-in seed corpus into an owned temporary
corpus and removes that temporary directory when the run finishes, while also
keeping generated build and artifact output out of the repository tree:

```sh
cd crates/litchi-numbers-wire/fuzz
fuzz_root="$(mktemp -d "${TMPDIR:-/tmp}/litchi-numbers-wire-fuzz.XXXXXX")"
mkdir -p "$fuzz_root/corpus" "$fuzz_root/artifacts"
cp -R corpus/pre_bnc_cell/. "$fuzz_root/corpus/"

CARGO_TARGET_DIR="$fuzz_root/target" cargo +nightly fuzz run \
  pre_bnc_cell "$fuzz_root/corpus" -- \
  -artifact_prefix="$fuzz_root/artifacts/" \
  -max_len=65536 -timeout=10 -rss_limit_mb=2048
fuzz_status=$?
rm -r "$fuzz_root"
exit "$fuzz_status"
```

Useful preflight commands are:

```sh
cargo metadata --manifest-path crates/litchi-numbers-wire/fuzz/Cargo.toml --no-deps --format-version 1
cd crates/litchi-numbers-wire/fuzz
cargo +nightly fuzz list
```
