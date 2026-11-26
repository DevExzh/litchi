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

## Seeding

Real metafiles make far better seeds than random bytes. Small files (below
64 KiB) from `test-data/` exercise the same records that Office documents
produce, and the sample below copies them into a corpus directory outside the
checkout so libFuzzer can grow the corpus without touching the repository:

```sh
mkdir -p "$fuzz_root/corpus/emf_png" "$fuzz_root/corpus/wmf_png"
find test-data -type f -iname '*.emf' -size -64k \
  -exec cp {} "$fuzz_root/corpus/emf_png/" \;
find test-data -type f -iname '*.wmf' -size -64k \
  -exec cp {} "$fuzz_root/corpus/wmf_png/" \;
```

`emfplus_render` and `emfplus_stream` take bare EMF+ record streams: concatenate
the `EMR_COMMENT_EMFPLUS` payloads of a dual-mode metafile in file order and
drop the result into their corpus directory. `blip_convert` and `odraw_scan`
take complete OfficeArt BLIP records; `dib_decode` takes one selector byte
followed by the packed DIB whose pixels follow its header, masks and color
table.

## Running

Put build output, corpora and artifacts outside the checkout, and keep the
input length and wall time bounded:

```sh
fuzz_root="$(mktemp -d "${TMPDIR:-/tmp}/litchi-imgconv-fuzz.XXXXXX")"
mkdir "$fuzz_root/artifacts" "$fuzz_root/corpus"
CARGO_TARGET_DIR="$fuzz_root/target" cargo +nightly fuzz run \
  emf_png "$fuzz_root/corpus/emf_png" -- \
  -artifact_prefix="$fuzz_root/artifacts/" -max_total_time=2400 \
  -max_len=262144 -timeout=10 -rss_limit_mb=2048
```

`-jobs` and `-workers` split one campaign across processes; keep the product
at or below the number of cores the machine can spare. `cargo +nightly fuzz
list` and `cargo +nightly fuzz check <target>` inspect the registered targets
without starting a campaign.
