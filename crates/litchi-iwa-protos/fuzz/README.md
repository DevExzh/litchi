# Numbers tile-storage codec fuzzing

`numbers_tile_storage` sends one bounded, caller-owned byte source through both
tile entry points:

* `decode_tile_with_report` (scalar/report path); and
* `decode_tile_with_visitor` (streaming row path).

For a successful decode the target checks scalar equality, exact
`DecodeReport` equality, and source-order/count summaries from the streaming
callbacks. Every callback checks that each non-empty row payload points into
the unchanged source. The visitor retains only compact summaries for the
first `MAX_FIELDS / 32` rows (256 with the fixed limits below); pointer checks
continue for all later callbacks without retaining their payloads. When
decoding fails, it consumes no partial result and only checks that both strict
paths fail; a streaming visitor may have seen rows before the later error, as
specified by the visitor API. After strict scalar success, the target decodes
the same source with the generated Prost `tst::Tile` oracle and compares every
tile scalar plus every streamed row scalar, optional-field presence, payload,
and source-order ordinal. A Prost rejection after strict acceptance is a
fuzz failure.

The target accepts raw inputs up to 64 KiB. Checked-in corpus entries are
human-readable `hex:` recipes; the harness decodes those recipes before the
two calls. This keeps the corpus reviewable while ensuring the codec sees
the exact same binary source in each path. Invalid or oversized recipes are
skipped. The finite per-source policy is 8,192 fields, 256 KiB of work, 1,024
references, 64 KiB of text, and recursion depth 64.

## Corpus provenance

The seeds are hand-authored protobuf wire encodings from
`src/buffa-projections/TSTTableCellStorageArchive.proto` and the matching
`src/protos/TSTArchives.proto`; they are not copied from a private Numbers
document. `modern.hex` includes current/BNC row buffers and all optional tile
scalars. `pre_bnc.hex` contains only the required pre-BNC row buffers.
`empty.hex` is the required-field, zero-value tile. `wide.hex` has sixteen
wide-offset rows and 8,191-by-8,191 tile bounds. `unknown_fields.hex` appends
unknown scalar and length-delimited fields. The three `malformed_*.hex`
recipes are near-valid duplicate-required, overlong-varint, and truncated-row
inputs. The recipes are deterministic regression inputs; they do not claim to
represent a complete native Numbers corpus.

From this directory, list and type-check the target:

```sh
cargo +nightly fuzz list
cargo +nightly fuzz check numbers_tile_storage
```

Run a bounded nightly sanitizer smoke (the target itself is intentionally
short; increase `-runs` only for a longer local campaign). The seed corpus is
copied to a temporary directory so libFuzzer's corpus additions, artifacts,
and build output never land in this repository. Set `KEEP_FUZZ_CORPUS=1` to
retain that temporary directory for review; otherwise the exact temporary
directory is removed on exit.

```sh
fuzz_root="$(mktemp -d "${TMPDIR:-/tmp}/litchi-tile-storage-fuzz.XXXXXX")"
fuzz_corpus="$fuzz_root/corpus"
mkdir "$fuzz_corpus" "$fuzz_root/artifacts"
cp corpus/numbers_tile_storage/*.hex "$fuzz_corpus/"
cleanup_fuzz_corpus() {
  if [ "${KEEP_FUZZ_CORPUS:-0}" = 1 ]; then
    printf 'retained temporary fuzz root: %s\n' "$fuzz_root"
  else
    rm -rf "$fuzz_root"
  fi
}
trap cleanup_fuzz_corpus EXIT
CARGO_TARGET_DIR="$fuzz_root/target" cargo +nightly fuzz run \
  numbers_tile_storage "$fuzz_corpus" -- \
  -artifact_prefix="$fuzz_root/artifacts/" -runs=1 -max_len=65536 \
  -timeout=10 -rss_limit_mb=2048
```

`cargo +nightly fuzz run` is the sanitizer invocation. The corpus and target
do not write generated artifacts into the repository. If a run is interrupted,
the temporary directory remains recoverable under the system temporary
directory; set `KEEP_FUZZ_CORPUS=1` before the command to print and retain its
path for review, then inspect or move it before removing it.
