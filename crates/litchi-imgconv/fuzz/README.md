# litchi-imgconv fuzzing

Each target drives one bounded entry point of the metafile-image pipeline with
small explicit budgets, so one execution terminates quickly and a hang is a
finding rather than an accident:

| Target | Entry point | Budget |
| --- | --- | --- |
| `emf_svg` | `convert_metafile(Emf, Svg)` | `Limits::default()` |
| `emf_png` | `convert_metafile(Emf, Png)` and `(Emf, Auto)` | 256x256 render, 1 MiB output, 20 000 records |
| `wmf_svg` | `convert_metafile(Wmf, Svg)` | `Limits::default()` |
| `wmf_png` | `convert_metafile(Wmf, Png)` and `(Wmf, Auto)` | 256x256 render, 1 MiB output, 20 000 records |
| `pict_png` | `pict::PictParser` + `PictConverter` | 256x256 bitmap, 1 MiB output |
| `pict_convert` | `pict::convert_pict` | the function's default budgets |
| `blip_convert` | `image::Blip::parse` + `decode_data` + `to_png`/`to_svg` | 1 MiB BLIP, 256x256 render |
| `odraw_scan` | `image::scan`/`store`/`get`/`all`/`delay`/`record` | 20 000 records, 1 MiB record body |
| `emfplus_stream` | `EmfPlusRecordIter`, `validate_complete_stream`, `try_extract_emfplus_comment_body` | 1 MiB, 20 000 records |
| `emfplus_render` | `EmfPlusSvgRenderer::push_payload` + `finish`, and the same payload as one `EMR_COMMENT_EMFPLUS` inside a synthetic EMF | 1 MiB, 20 000 records |
| `dib_decode` | `dib::Dib::parse_with_usage` + `to_bmp`/`to_dynamic_image*`/`to_png`, and `parse_parts` | 256x256, 1 MiB |

The first input byte of `dib_decode` selects the color usage and alpha
interpretation; `emfplus_render` wraps arbitrary bytes as the payload of a
single `EMR_COMMENT_EMFPLUS` (MS-EMF 2.3.2.7) so EMF+ playback receives
mutation-shaped record streams from a valid outer metafile.

The targets also assert the one invariant that is cheap to check per
execution: encoded output never exceeds the requested `max_output_bytes`.

## Building

`cargo +nightly fuzz build` builds every target with libFuzzer and the address
sanitizer. The build output lives under `target/`, which is ignored.

## Running

Put build output, corpora and artifacts outside the checkout, and keep the
input length and wall time bounded. The commands below run from the repository
root; the corpus directory is seeded first (see below):

```sh
fuzz_root="$(mktemp -d "${TMPDIR:-/tmp}/litchi-imgconv-fuzz.XXXXXX")"
mkdir -p "$fuzz_root/artifacts" "$fuzz_root/corpus/emf_png"
find test-data -type f -iname '*.emf' -size -64k \
  -exec cp {} "$fuzz_root/corpus/emf_png/" \;
cd crates/litchi-imgconv/fuzz
CARGO_TARGET_DIR="$fuzz_root/target" cargo +nightly fuzz run \
  emf_png "$fuzz_root/corpus/emf_png" -- \
  -artifact_prefix="$fuzz_root/artifacts/" -max_total_time=2400 \
  -max_len=262144 -timeout=10 -rss_limit_mb=2048
```

`-fork=N` splits one campaign across `N` processes; adding
`-ignore_crashes=1 -ignore_ooms=1 -ignore_timeouts=1` keeps a campaign running
after a finding, so one run collects every distinct artifact. Keep `N` at or
below the number of cores the machine can spare. `cargo +nightly fuzz list` and
`cargo +nightly fuzz check <target>` inspect the registered targets without
starting a campaign.

## Seeding

Real metafiles make far better seeds than random bytes. Small files (below
64 KiB) from `test-data/` exercise the same records that Office documents
produce. Copy them into the corpus directory of `emf_svg`/`emf_png`
(`*.emf`) or `wmf_svg`/`wmf_png` (`*.wmf`); the `find` command above does this
for one target. Further seeds come from the metafiles an Office document
carries: extract them from your own documents and keep them below 64 KiB.

Targets with a structured input take these seeds:

- `emfplus_render` and `emfplus_stream` take bare EMF+ record streams:
  concatenate the `EMR_COMMENT_EMFPLUS` payloads of a dual-mode metafile in
  file order.
- `pict_png` and `pict_convert` take raw PICT data, with or without the
  512-byte header.
- `blip_convert` and `odraw_scan` take complete OfficeArt BLIP records.
- `dib_decode` takes one selector byte followed by the packed DIB whose pixels
  follow its header, masks and color table.
