# Numbers table-merge fuzzing

`numbers_table_merges` exercises the focused Numbers package through its
selector-first merged-cell reader.  The target starts with the checked-in
Numbers 14.4 fixture containing a native `B11:C12` merge, then varies sheet
and table name/index selectors, out-of-range selectors, and missing names.
Each read is followed by exact source serialization so a successful read must
leave the package bytes unchanged.

The same bounded profile is used for one-byte mutations of the native ZIP and
for arbitrary inputs that begin with the ZIP signature.  These paths cover
malformed package and IWA framing while keeping descriptor input at 64 KiB,
package input at 512 KiB, and semantic cell/text retention bounded.  The
harness does not decode the merge wire format directly or depend on the
migration host crate.

Run a short sanitizer smoke from this directory with all generated state in a
temporary directory:

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

Use `cargo +nightly fuzz check numbers_table_merges` for a compile-only
preflight.  Generated corpus entries, artifacts, and standalone build output
must remain outside the checkout.
