# Format-detection fuzzing

`detect_format` exercises the public RTF signature detector with arbitrary
bytes. The detector is intentionally a small, non-allocating entry point, so
the corpus keeps one positive five-byte signature and one near-miss negative
input as reviewable smoke cases:

* `corpus/detect_format/rtf-prefix` starts with the minimal `{"\\rtf` marker;
* `corpus/detect_format/plain-prefix` is a non-RTF prefix of the same size.

The harness does not write files or retain decoded input. Keep local fuzzing
bounded with a maximum input length and put generated artifacts and build
output outside the checkout:

```sh
fuzz_root="$(mktemp -d "${TMPDIR:-/tmp}/litchi-core-fuzz.XXXXXX")"
mkdir "$fuzz_root/artifacts"
cleanup_fuzz_root() {
  if [ "${KEEP_FUZZ_CORPUS:-0}" = 1 ]; then
    printf 'retained temporary fuzz root: %s\n' "$fuzz_root"
  else
    rm -rf "$fuzz_root"
  fi
}
trap cleanup_fuzz_root EXIT
CARGO_TARGET_DIR="$fuzz_root/target" cargo +nightly fuzz run \
  detect_format corpus/detect_format -- \
  -artifact_prefix="$fuzz_root/artifacts/" -runs=100 -max_len=65536 \
  -timeout=10 -rss_limit_mb=2048
```

`cargo +nightly fuzz list` and `cargo +nightly fuzz check detect_format` can
be run from this directory to inspect the registered target without starting
a campaign.
