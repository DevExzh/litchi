# Root iWork fuzzing

`parse_iwork` owns fuzz coverage for the supported `litchi::iwork` byte
coordinator. It exercises bounded package admission, format dispatch, all
three semantic projections, and the archive-free snapshot facade. The fuzz
package deliberately depends only on the root crate with the `iwork` feature;
it must not acquire a dependency on the legacy `litchi-iwa` migration host or
on an internal archive, protobuf, Buffa, or concrete-format crate.

`keynote_slide_text` is the focused title/body robustness target. It first
offers arbitrary bytes to the bounded Keynote package ingress, then uses those
same bytes as commands against the repository's native `basic.key` seed. This
second path bypasses the low survival rate of CRC-protected ZIP mutation while
still exercising both placeholder roles, semantic selectors, UTF-16 boundary
validation, set/clear/replace/insert/delete/no-op staging, exact-source patch
application, inversion, and content-redacted errors. It never writes a package
to disk.

`keynote_show_settings` is the focused presentation-settings target. It offers
arbitrary bytes to bounded Keynote package ingress and reuses a fixed prefix
for no-op, playback-only, slide-number/size rendering, and combined commands
against the native `basic.key` seed. The playback-only changed command asserts
that public commit diagnostics report no deleted root previews. It covers
strict reads, exact-source patch application and conflicts, inversion, typed
limits, content-redacted failures, public in-memory `write_to` verification,
and exact byte restoration without writing a package to disk.

`numbers_table_lock` is the focused interactive table-lock target. It offers
arbitrary bytes to checked Numbers package ingress and also interprets them as
bounded selector and lock-state commands against the native `basic.numbers`
seed. It covers name and index selectors, exact no-op and changed commits,
exact-source patch conflicts, inversion, and content-redacted errors without
writing a package to disk.

`numbers_names` is the focused atomic sheet/table-names target. It offers
arbitrary bytes to bounded Numbers package ingress and reuses a fixed command
prefix for no-op, sheet, table, and combined renames against the native
`basic.numbers` seed. It covers selector staging, Unicode names, typed finite
limits, content-redacted failures, exact-source patch application and
conflicts, inversion, and byte-exact restoration through public in-memory
`write_to` without writing a package to disk.

`pages_page_layout` is the focused Pages document-layout target. It offers
arbitrary bytes to checked Pages package ingress and reuses them as bounded
layout commands against the native `basic.pages` seed. It covers public layout
reads and validation, exact no-op and changed commits, exact-source patch
conflicts, inversion, content-redacted failures, and exact restoration without
writing a package to disk.

`pages_document_settings` is the focused combined Document and Footnotes
formatter target. It offers arbitrary bytes to checked Pages package ingress
and reuses a fixed prefix for bounded option and footnote commands against the
native `basic.pages` seed. It covers strict reads, exact no-ops, combined
changes, exact-source conflicts, inversion, typed limits, content-redacted
failures, and exact restoration without writing a package to disk.

`parse_iwork` uses tighter limits than the public defaults: 2 MiB of source
bytes, 512 package entries, 8 MiB per expanded entry and decoded IWA item,
32 MiB aggregate expanded bytes, 4,096 values of each semantic collection,
and 4 MiB of retained text. Keep its `-max_len` aligned with the 2 MiB source
ceiling so oversized mutations do not consume fuzzing time.

`keynote_slide_text` uses a narrower physical profile for its arbitrary-byte
path: 1 MiB of source bytes, 256 package entries, 2 MiB per expanded entry and
decoded IWA item, and 8 MiB aggregate expanded bytes. Its native `basic.key`
seed is embedded in the harness from the hash-verified source below; fuzzer
input supplies only bounded transaction commands, of which at most 1 KiB
becomes replacement text. Keep this target's `-max_len` at 4 KiB to
concentrate effort on deep-message operations.

`keynote_show_settings` reuses the same finite Keynote physical and semantic
profile. Settings commands consume only a fixed prefix; keep `-max_len` at 512
bytes so arbitrary ingress remains active while every input also reaches the
fixed native transaction.

`numbers_table_lock` accepts at most 512 KiB of source bytes, 128 package
entries, 1 MiB per expanded entry and decoded IWA item, and 4 MiB aggregate
expanded bytes. Its semantic profile admits at most 4,096 objects, 128 sheets,
512 tables, 8,192 references, 65,536 materialized cells, and 512 KiB of
retained text. Fuzzer-derived selector names are limited to 512 input bytes;
keep `-max_len` at 1 KiB so most work reaches the fixed native seed.

`numbers_names` uses the same finite Numbers physical and semantic profile.
Fuzzer-derived names are decoded lossily as UTF-8, reject NUL, and consume at
most 256 input bytes; keep `-max_len` at 1 KiB so malformed ingress and native
name transactions both receive every input.

`pages_page_layout` accepts at most 256 KiB of source bytes, 128 package
entries, 1 MiB per expanded entry and decoded IWA item, and 4 MiB aggregate
expanded bytes. Layout commands consume only a small fixed prefix; keep
`-max_len` at 512 bytes to retain malformed-ingress mutation while ensuring
every input also reaches the fixed native seed.

`pages_document_settings` reuses the same finite Pages physical profile.
Settings commands consume only a fixed prefix, so keep `-max_len` at 512 bytes
to combine malformed-ingress mutations with deterministic native transaction
coverage.

All targets currently share this package's single `litchi` dependency with
the `iwork` feature. Cargo unifies dependency features for the package, so the
focused Keynote, Numbers, and Pages binaries compile the complete root iWork
feature set rather than isolated concrete-format graphs. The focused targets'
source-level imports remain confined to the relevant public facade, but true
dependency isolation requires separate owner-specific fuzz packages or
dependencies.

## Native seed provenance

The seed documents already live in `test-data/iwork`; do not check in a second
binary copy. They were created, saved, closed, and reopened with the native
macOS applications, and their visible content is documented in
`test-data/iwork/README.md`.

| Corpus name | Source | Bytes | SHA-256 |
| --- | --- | ---: | --- |
| `basic.pages` | `test-data/iwork/pages/basic.pages` | 96,417 | `21107bc9323fba6f1589152454c0b0b0cc8e239313c6a369bc4a891116601b42` |
| `basic.numbers` | `test-data/iwork/numbers/basic.numbers` | 136,357 | `f225d5b1cd59e9da454f91a96fe8f81154bc31037c10029230e75d49b45fb693` |
| `basic.key` | `test-data/iwork/keynote/basic.key` | 500,058 | `3a3d07476b45b6e543bcfba75fe38a245434176dcb3565e34570b817708b9f42` |

From this directory, prepare the ignored local corpus and verify its hashes:

```sh
mkdir -p corpus/parse_iwork
cp ../../../test-data/iwork/pages/basic.pages corpus/parse_iwork/
cp ../../../test-data/iwork/numbers/basic.numbers corpus/parse_iwork/
cp ../../../test-data/iwork/keynote/basic.key corpus/parse_iwork/
shasum -a 256 corpus/parse_iwork/basic.pages \
  corpus/parse_iwork/basic.numbers corpus/parse_iwork/basic.key
```

Run the target with an address sanitizer and explicit process ceilings:

```sh
cargo +nightly fuzz run parse_iwork corpus/parse_iwork -- \
  -max_len=2097152 -timeout=10 -rss_limit_mb=2048
```

Run the focused Keynote target without a checked-in duplicate corpus:

```sh
cargo +nightly fuzz run keynote_slide_text -- \
  -max_len=4096 -timeout=10 -rss_limit_mb=2048
```

Run the focused Keynote show-settings target:

```sh
cargo +nightly fuzz run keynote_show_settings -- \
  -max_len=512 -timeout=10 -rss_limit_mb=2048
```

The `cargo +nightly fuzz run` commands are required for sanitizer-instrumented
coverage. A stable `cargo run --bin <target> -- -runs=...` invocation is only
a control-flow smoke test; on platforms without linked sanitizer runtimes it
may print missing-symbol warnings and is not sanitizer evidence.

Run the focused Numbers target without a checked-in duplicate corpus:

```sh
cargo +nightly fuzz run numbers_table_lock -- \
  -max_len=1024 -timeout=10 -rss_limit_mb=2048
```

Run the focused Numbers names target without a checked-in duplicate corpus:

```sh
cargo +nightly fuzz run numbers_names -- \
  -max_len=1024 -timeout=10 -rss_limit_mb=2048
```

Run the focused Pages target without a checked-in duplicate corpus:

```sh
cargo +nightly fuzz run pages_page_layout -- \
  -max_len=512 -timeout=10 -rss_limit_mb=2048
```

Run the focused Pages document-settings target:

```sh
cargo +nightly fuzz run pages_document_settings -- \
  -max_len=512 -timeout=10 -rss_limit_mb=2048
```

The native packages currently store ZIP members with CRC protection. Most
arbitrary payload mutations therefore fail during physical validation, which
is appropriate for the root ingress target. Valid deep-message mutation and
format-specific behavioral invariants belong in focused format-owner fuzz
targets rather than weakening package validation here.

## 2026-08-08 bounded sanitizer evidence

The three native seeds above were copied to a private temporary corpus and
their documented SHA-256 values were verified before this run. With
`cargo-fuzz 0.13.2` and `rustc 1.99.0-nightly (1a98b1e13 2026-08-07)`, the
root target completed this bounded AddressSanitizer/libFuzzer campaign:

```sh
cargo +nightly fuzz run parse_iwork /private/tmp/litchi-iwork-root-fuzz-1JVxsJ/corpus -- \
  -max_total_time=60 -max_len=2097152 -timeout=10 -rss_limit_mb=2048
```

It executed 152,219 inputs in 61 seconds, ending at coverage 7,454, feature
count 12,062, a 249-input / 69 MiB corpus, and 566 MiB RSS. There was no crash,
timeout, or out-of-memory finding. This records one bounded root-ingress
sanitizer run only; it is not evidence of exhaustive fuzzing or of the focused
deep-message campaigns still required from the format owners.
