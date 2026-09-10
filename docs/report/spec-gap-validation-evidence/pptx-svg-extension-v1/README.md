# PPTX SVG extension validation

The picture inventory now recognizes the native SVG extension GUID
`{96DAC541-7B7A-43D3-8B79-37D633B846F1}` separately from the namespace of
`asvg:svgBlip`. Non-SVG producer extensions remain opaque through every
nested element. Typed-looking descendants cannot attach relationships or
re-enter the picture grammar. Genuine SVG extensions retain strict payload
and duplicate validation.

The three-file batch includes the parser, the source-backed SVG/MCE
preservation fixture, and performance-harness native-image expectations.
The original shapes fixture inventories its image; the LibreOffice resave's
malformed empty stretch remains a refusal. No rendering or native application
edit-acceptance claim is made. The local MS-ODRAWXML SVG section specifies
the child namespace; the native GUID is corroborated by local LibreOffice
`oox/source/export/drawingml.cxx` and POI `XSLFPictureShape.java` producer code.

The isolated validation base is `1b3f2c2d0c059e8e59272775a97d0567e14e67f2`.
The harness lock correction was separately committed in `23dc8248a` and is
excluded from this feature delta. Three exact source hashes are retained in
`source-manifest.json`. Root's final full gate passed 868 unit/integration
tests and 6 doctests (2 doctests ignored), strict all-target/all-feature
Clippy, warning-denied rustdoc, pinned formatting, and diff checks.
Independent source review cleared the final fixture and parser changes.

The included earlier harness checks passed 5 native-image tests and 353
library tests (1 ignored). Those checks cover the unchanged runtime parser
and harness-test source; the final third-file change corrects only the PPTX
integration fixture. The full root PPTX suite includes that final correction.
These are correctness gates, not performance measurements.

## Reproduction

Use Rust 1.95.0 and a fresh target directory. Restore `Cargo.lock.gz` as the
workspace `Cargo.lock` to retain dependency resolution. Root executed:

```sh
cargo +1.95.0 test -p litchi-pptx --all-features --no-fail-fast
cargo +1.95.0 clippy -p litchi-pptx --all-targets --all-features -- -D warnings
RUSTDOCFLAGS='-D warnings' cargo +1.95.0 doc -p litchi-pptx --all-features --no-deps
```

These root commands did not use `--offline` or `--locked`. An optional replay
may add `--locked` after restoring the supplied lock. The separate harness
lock is `perf-Cargo.lock.gz`; harness tests use
`cargo +1.95.0 test --manifest-path tools/perf-baseline/Cargo.toml --lib --locked`.
Allocator or timing baseline capture is outside this batch.
