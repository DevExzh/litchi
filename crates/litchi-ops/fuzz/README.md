
Query plans: `cargo run --manifest-path crates/litchi-ops/fuzz/Cargo.toml --bin
smoke_queries` runs deterministic mutations of harvested selectors across three
fake format families. The `query_plan` libFuzzer target also compiles and runs
plans with bounded output. This validates the engine, not production adapters.

Batch payloads (record 0334): `batch_payload` feeds arbitrary bytes to
`batch::from_bytes` with a required `SourceLimits` profile (64 KiB, 64 items),
the OfficeCLI batch JSON ingress. Every input either fails with `corrupt_batch` or `limit_exceeded`, or
decodes to `BatchItem`s that validate, survive `into_owned`, lower their path
fields through `PathArg::parse` and their selector through `Query::parse` (to a
display fixed point or a typed path/selector refusal), and serialize to a
canonical wire form that decodes back to equal items. `batch_apply` (the
`apply` feature: DOCX, XLSX and PPTX adapters) runs every decoded batch of at
most 16 items against blank created documents, atomically and best effort with
`Unsupported::Warn`, and requires a commit or a typed refusal other than
`internal_error`; an atomic batch must behave the same when retried on the
same handle.

Seeds are in `corpus/batch_payload/`: hand-written cases mirroring the decoder
tests in `../src/batch/tests.rs`, and `officecli-*.json`, the first six items
of one batch from each OfficeCLI example script tracked under
`test-data/officecli/examples/` (OfficeCLI, Apache-2.0; see
`test-data/officecli/PROVENANCE.md`). `python3
crates/litchi-ops/fuzz/harvest_batch_seeds.py` regenerates the latter.

On the stable toolchain, `smoke_batch_payload` replays every seed, every
prefix and every single-bit flip through `batch_payload`, and with
`--features apply` a strided sample through `batch_apply`:

```sh
cargo run --release --manifest-path crates/litchi-ops/fuzz/Cargo.toml \
  --features apply --bin smoke_batch_payload
```

All targets type-check on stable, which is what the gate runs:

```sh
cargo check --manifest-path crates/litchi-ops/fuzz/Cargo.toml --bins
cargo check --manifest-path crates/litchi-ops/fuzz/Cargo.toml --bins --features apply
```

A libFuzzer campaign needs `cargo +nightly fuzz run batch_payload
corpus/batch_payload` from this directory, with build output outside the
checkout as in `../../litchi-core/fuzz/README.md`.

`regression-*.json` are libFuzzer finds of record 0334: a relative path whose
canonical spelling read back as sheet notation (finding F2) and long numbers
that drifted by ULPs across wire round trips (finding F1). Record 0338 fixed
both, and the harness now requires exact fixed points: the path display
reparses through `PathArg::parse` to the same path and spelling, and the wire
form decodes to equal items (numbers included) and re-serializes to the same
bytes.

Without nightly or `cargo-fuzz`, the libFuzzer targets still build and run on
the stable toolchain with LLVM's coverage pass (this is how record 0334 ran
them):

```sh
RUSTFLAGS="-Cpasses=sancov-module -Cllvm-args=-sanitizer-coverage-level=4 \
  -Cllvm-args=-sanitizer-coverage-inline-8bit-counters \
  -Cllvm-args=-sanitizer-coverage-pc-table \
  -Cllvm-args=-sanitizer-coverage-trace-compares \
  --cfg fuzzing -Cdebug-assertions -Coverflow-checks" \
cargo build --release --target x86_64-unknown-linux-gnu \
  --manifest-path crates/litchi-ops/fuzz/Cargo.toml --features apply \
  --bin batch_payload --bin batch_apply
# copy the corpus to a scratch directory first; libFuzzer adds to it
target/x86_64-unknown-linux-gnu/release/batch_payload "$scratch/corpus" \
  -artifact_prefix="$scratch/artifacts/" -max_total_time=300 -max_len=65536
```

`--target` keeps the flags off build scripts. There is no sanitizer, so only
panics (assertions) are detected.
