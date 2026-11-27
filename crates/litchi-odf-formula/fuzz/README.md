# OpenDocument Formula fuzz campaign

The `formula` target has four input modes selected by the low two bits of the
first byte:

- `0`: arbitrary UTF-8 MathML, with parse/serialize/validation/package
  round-trip properties for accepted trees;
- `1`: arbitrary Formula package bytes, with exact-provenance and reopen
  properties for accepted archives;
- `2`: a byte-decoded, bounded MathML 2 grammar derived independently from the
  W3C schema/signature examples; generated trees must parse, validate, serialize
  compactly, package, and fully reopen;
- `3`: one-rule grammar breakers which must remain rejected.

The checked-in `corpus/formula` directory is both the initial libFuzzer corpus
and the compact regression corpus. It contains the semantic shapes which have
historically exposed projection gaps, plus byte recipes for structured valid
and invalid generation.

## Reproducible commands

Replay the corpus without installing `cargo-fuzz`:

```sh
cargo run --release \
  --manifest-path crates/litchi-odf-formula/fuzz/Cargo.toml \
  --bin replay -- crates/litchi-odf-formula/fuzz/corpus/formula
```

Run a five-minute coverage-guided campaign after installing `cargo-fuzz`:

```sh
cargo install cargo-fuzz --locked
cd crates/litchi-odf-formula/fuzz
mkdir -p artifacts/formula
cargo fuzz run formula corpus/formula -- \
  -seed=20260810 -max_total_time=300 -timeout=10 -rss_limit_mb=2048 \
  -artifact_prefix=artifacts/formula/
```

Minimize the retained corpus after a successful campaign:

```sh
cd crates/litchi-odf-formula/fuzz
cargo fuzz cmin formula corpus/formula
```

Reproduce a saved failure with:

```sh
cd crates/litchi-odf-formula/fuzz
cargo fuzz run formula artifacts/formula/<artifact>
```

Do not commit `artifacts/` or generated coverage output. Commit only minimized
inputs that reproduce a distinct semantic or safety regression.
