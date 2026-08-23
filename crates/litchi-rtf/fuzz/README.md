# RTF fuzz manifest

`Cargo.toml` registers one target:

| Target | Harness | Checked-in corpus |
| --- | --- | --- |
| `parse_rtf` | `fuzz_targets/parse_rtf.rs` | none (arbitrary bytes) |

`parse_rtf` sends each input through the compressed-RTF transport decoder and
the public `Document` byte parser. Inputs larger than 1 MiB are skipped rather
than truncated. Both paths use explicit finite expansion, token, binary, and
opaque-node limits; successful parses additionally visit document text and a
bounded paragraph selector. The harness does not write files or retain a
generated corpus in the checkout.

Inspect the registered target without starting a campaign:

```sh
cargo metadata --no-deps --format-version 1
cargo +nightly fuzz list
cargo +nightly fuzz check parse_rtf
```

For a bounded AddressSanitizer/libFuzzer smoke, keep the mutable corpus,
artifacts, and build output in a temporary directory:

```sh
fuzz_root="$(mktemp -d "${TMPDIR:-/tmp}/litchi-rtf-fuzz.XXXXXX")"
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
  parse_rtf "$fuzz_root/corpus" -- \
  -artifact_prefix="$fuzz_root/artifacts/" -runs=100 -max_len=1048576 \
  -timeout=10 -rss_limit_mb=2048
```

Generated corpus entries and artifacts are disposable campaign output; do not
copy them into this package unless a small, reviewable regression seed is
intentionally added.
