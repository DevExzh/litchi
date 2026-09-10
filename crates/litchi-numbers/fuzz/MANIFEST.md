# Numbers fuzz manifest

`Cargo.toml` registers the owner-level `numbers_table_merges` target.  The
harness keeps the checked-in native merge fixture as a bounded source oracle,
then varies only semantic selectors and small package mutations.  It also
admits arbitrary package-shaped bytes through the same explicit physical and
semantic profiles.  Fuzz corpora and generated artifacts stay outside the
checkout.

| Registered target | Harness | Native source |
| --- | --- | --- |
| `numbers_table_merges` | `fuzz_targets/numbers_table_merges.rs` | `test-data/iwork/numbers/table-merges-native.numbers` |

The target does not use `Package::from_bytes` defaults.  Package ingress is
capped at 512 KiB, with at most 512 ZIP members, 256 KiB per member, 2 MiB
aggregate declared member bytes, and 512 KiB per decompressed IWA stream.
Semantic admission is capped at 16,384 objects, 8 sheets, 64 tables, and
4,096 references, 4,096 materialized cells, and 64 KiB of retained text.
Descriptor input is capped at 64 KiB.

Inspect the target without starting a campaign:

```sh
cargo metadata --manifest-path crates/litchi-numbers/fuzz/Cargo.toml --no-deps --format-version 1
cd crates/litchi-numbers/fuzz
cargo +nightly fuzz list
cargo +nightly fuzz check numbers_table_merges
```

Run a bounded sanitizer smoke with mutable corpus, artifacts, and build output
outside the checkout:

```sh
fuzz_root="$(mktemp -d "${TMPDIR:-/tmp}/litchi-numbers-merges-fuzz.XXXXXX")"
mkdir "$fuzz_root/corpus" "$fuzz_root/artifacts"
cleanup_fuzz_root() {
  if [ "${KEEP_FUZZ_CORPUS:-0}" = 1 ]; then
    printf 'retained temporary fuzz root: %s\n' "$fuzz_root"
  else
    rm -rf "$fuzz_root"
  fi
}
trap cleanup_fuzz_root EXIT

cd crates/litchi-numbers/fuzz
CARGO_TARGET_DIR="$fuzz_root/target" cargo +nightly fuzz run \
  numbers_table_merges "$fuzz_root/corpus" -- \
  -artifact_prefix="$fuzz_root/artifacts/" -runs=100 -max_len=65536 \
  -timeout=10 -rss_limit_mb=2048
```
