# change 0581 evidence

Evidence for [`0581-opc-package-retention.md`](../../0581-opc-package-retention.md).
Design only; `performance_claim: none`. Every figure is a deterministic
peak-retained-byte or allocation count. **No timing was measured and none is
claimed.**

## How it was produced

All figures come from a **detached git worktree of
`32d25e08806d93f792ffd4954d83acc9db9c5301`** with an isolated `CARGO_TARGET_DIR`
in the session scratchpad, so that concurrent edits to `crates/soapberry-zip/`
and `crates/litchi-cfb/` in the main working tree could not link into the probe.
See [`environment.txt`](environment.txt).

The probe uses a process-global counting allocator with the same wrapper shape as
`tools/perf-baseline/src/bin/support/counting_allocator.rs`, and the same shape as
change 0578's probe, so the numbers are directly comparable with
[`../change-0578/`](../change-0578). The sink is sequential and non-seek: it
implements only `Write`, accepts each buffer whole, discards it, and folds an
FNV-1a digest of the complete output stream (hashing allocates nothing, so it does
not disturb the measurement).

Each region reports three quantities:

- `region_peak_bytes` — high-water live heap over the region, minus the live heap
  at region entry;
- `allocations` — allocation count over the region;
- `retained_bytes` — live-heap delta across the region. The region's product is
  kept alive by the caller, so this is what an open package still holds.

## Files

| file | what it is |
| --- | --- |
| [`opc-open-save-split-probe.rs`](opc-open-save-split-probe.rs) | the probe. Extends change 0578's `opc-save-peak-probe.rs` by splitting each save path into an open region and a publish region and by reporting retention. |
| [`opc-open-save-split-probe.Cargo.toml`](opc-open-save-split-probe.Cargo.toml) | its manifest; the two path dependencies point into the detached worktree. |
| [`environment.txt`](environment.txt) | revision, toolchain, CPU, kernel, allocator. |
| [`stage1-repro-0578-axis5.csv`](stage1-repro-0578-axis5.csv) | change 0578's axis-5 probe re-run **unmodified**. Reproduces `../change-0578/stage1-peak-e2e-opc.csv` byte for byte, including the output digests. |
| [`stage1-open-save-split-by-size.csv`](stage1-open-save-split-by-size.csv) | axis 1 — largest media member 64 KiB to 128 MiB, member count fixed. |
| [`stage1-open-save-split-by-count.csv`](stage1-open-save-split-by-count.csv) | axis 2 — 0 to 8,000 members of 4,096 bytes each, member size fixed. |
| [`stage1-open-save-split-real-fixtures.csv`](stage1-open-save-split-real-fixtures.csv) | axis 3 — six real corpus fixtures, unedited open-and-save. |
| [`gate-cargo-test-litchi-opc.log`](gate-cargo-test-litchi-opc.log) | `cargo test -p litchi-opc` at `32d25e088`: 664 passed, 0 failed, 1 ignored, 24 binaries. |
| [`gate-check-perf-claims.log`](gate-check-perf-claims.log) | `check_perf_claims.py --mode structural`, run twice — pristine `32d25e088` and the working tree with this record. Identical output, exit 2 both times, from a pre-existing unrelated `claim-0251` message. |

## CSV schema

`axis,axis_value,path,phase,archive_bytes,region_peak_bytes,allocations,retained_bytes,output_bytes,output_fnv1a64`

- `path` — `A_package_writer` (`OpcPackage::open` + `PackageWriter::write_to_stream`)
  or `B_source_backed` (`SourceBackedPackage::from_path` + overlay/topology publish).
- `phase` — `open`, `save`, `drop` for A; `open`, `save_and_drop` for B. The
  source-backed publish takes `self` by value, so B has no separate drop row and
  its publish region also covers dropping the package.
- `output_fnv1a64` — non-zero only on publish rows. **A and B agree on every one
  of the 19 fixtures.**

## Reproducing

```sh
git worktree add --detach /tmp/w 32d25e08806d93f792ffd4954d83acc9db9c5301
mkdir -p /tmp/probe/src && cp opc-open-save-split-probe.rs /tmp/probe/src/main.rs
sed 's|<worktree of [0-9a-f]*>|/tmp/w|' opc-open-save-split-probe.Cargo.toml > /tmp/probe/Cargo.toml
CARGO_TARGET_DIR=/tmp/t cargo build --release --manifest-path /tmp/probe/Cargo.toml
/tmp/t/release/opc-save-peak-probe /tmp/w/test-data/ooxml/docx/drawing.docx /tmp/work size
/tmp/t/release/opc-save-peak-probe /tmp/w/test-data/ooxml/docx/drawing.docx /tmp/work count
/tmp/t/release/opc-save-peak-probe /tmp/w/test-data/ooxml/docx/drawing.docx /tmp/work real \
  /tmp/w/test-data/poi/test-data/slideshow/ArtisticEffectSample.pptx \
  /tmp/w/test-data/poi/test-data/document/saut_page.docx \
  /tmp/w/test-data/poi/test-data/spreadsheet/no_drawing_patriarch.xlsx \
  /tmp/w/test-data/ooxml/xlsx/ConditionalFormattingSamples.xlsx \
  /tmp/w/test-data/poi/test-data/slideshow/EmbeddedVideo.pptx \
  /tmp/w/test-data/ooxml/docx/drawing.docx
```

The `size` axis builds a 128 MiB fixture and needs roughly 1 GiB of free disk and
about 600 MiB of peak RSS. Synthetic fixtures are deleted after each measurement.

## What is not here

No flame graph, `perf` counter, timing series, or RSS sample, because none was
taken. **No decompressed-payload-sum column**, which is why the record states the
`archive + Σ decompressed` shape as consistent with the ratios rather than as
directly measured. The record's C2 prediction column is arithmetic on these CSVs
(`archive_bytes + B open retained_bytes`), not a separate measurement, and is
labelled as such in the record.
