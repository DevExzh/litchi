# CFB fuzz manifest

`Cargo.toml` registers one target:

| Target | Harness | Checked-in corpus |
| --- | --- | --- |
| `parse_cfb` | `fuzz_targets/parse_cfb.rs` | none (arbitrary bytes) |

`parse_cfb` exercises OLE2 signature detection, header/FAT/directory parsing,
directory walking, and bounded stream lookups. Inputs larger than 4 MiB are
skipped rather than truncated, which keeps the parser's sector bookkeeping
and stream materialization bounded while preserving one unchanged source for
each iteration. The harness does not write files or retain generated corpus
entries in the checkout.

Inspect the registered target without starting a campaign:

```sh
cargo metadata --no-deps --format-version 1
cargo +nightly fuzz list
cargo +nightly fuzz check parse_cfb
```

For a bounded AddressSanitizer/libFuzzer smoke, keep the mutable corpus,
artifacts, and build output in a temporary directory:

```sh
fuzz_root="$(mktemp -d "${TMPDIR:-/tmp}/litchi-cfb-fuzz.XXXXXX")"
mkdir "$fuzz_root/corpus" "$fuzz_root/artifacts"
cleanup_fuzz_root() {
  if [ "${KEEP_FUZZ_CORPUS:-0}" = 1 ]; then
    printf 'retained temporary fuzz root: %s\n' "$fuzz_root"
  else
    rm -rf "$fuzz_root"
  fi
}
trap cleanup_fuzz_root EXIT
CARGO_TARGET_DIR="$fuzz_root/target" cargo +nightly fuzz run \
  parse_cfb "$fuzz_root/corpus" -- \
  -artifact_prefix="$fuzz_root/artifacts/" -runs=100 -max_len=4194304 \
  -timeout=10 -rss_limit_mb=2048
```

Generated corpus entries and artifacts are disposable campaign output; do not
copy them into this package unless a small, reviewable regression seed is
intentionally added.
