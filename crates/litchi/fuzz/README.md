# Root iWork fuzzing

`parse_iwork` owns fuzz coverage for the supported `litchi::iwork` byte
coordinator. It exercises bounded package admission, format dispatch, all
three semantic projections, and the archive-free snapshot facade. The fuzz
package deliberately depends only on the root crate with the `iwork` feature;
it must not acquire a dependency on the legacy `litchi-iwa` migration host or
on an internal archive, protobuf, Buffa, or concrete-format crate.

The harness uses tighter limits than the public defaults: 2 MiB of source
bytes, 512 package entries, 8 MiB per expanded entry and decoded IWA item,
32 MiB aggregate expanded bytes, 4,096 values of each semantic collection,
and 4 MiB of retained text. Keep `-max_len` aligned with the 2 MiB source
ceiling so oversized mutations do not consume fuzzing time.

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
cargo fuzz run parse_iwork corpus/parse_iwork -- \
  -max_len=2097152 -timeout=10 -rss_limit_mb=2048
```

The native packages currently store ZIP members with CRC protection. Most
arbitrary payload mutations therefore fail during physical validation, which
is appropriate for the root ingress target. Valid deep-message mutation and
format-specific behavioral invariants belong in focused format-owner fuzz
targets rather than weakening package validation here.
