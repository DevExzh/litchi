# 0362ul94e — OOXML common compiler repair return

The actual compiler-first capture on `0395ca40f41961def43743e2003d60ccf871adca`
passed the changed-source formatting row and stopped at the common-package
compiler row.  This failure remains recorded; it is not relabeled as a pass.

The exact failing command was:

```text
cargo check --keep-going --offline --locked -q -p litchi-ooxml-common --all-targets --all-features
```

The capture exited 101 after 36.664550 seconds and left a 158,668,903-byte
target.  The formatting row passed in 2.099288 seconds.  The compiler output
contained ten diagnostics: the lifetime on `Debug for vba::Edit`, two
unnecessary `transaction::Error` qualifications in `properties.rs`, one
heterogeneous signed-test closure, two recipient `SourceLimitsBuilder`
arguments, and four structural-limit `SourceLimitsBuilder` arguments.

The immutable actual state is:

```text
(deleted-coord-artifact)
```

| Artifact | Bytes | SHA-256 |
|---|---:|---|
| `rows.json` | 12,095 | `5fbc7ccde72bf130e6d1582eadbdb2cb6afed48973edd6c54a12a98caf7ad8f7` |
| `result.json` | 13,179 | `1aa33d5fd58dcf5ee6e8dd8678f6149334715d4a0a1196720eeff30c2174a1d4` |
| `01-check-common-attempt-actual-f5-01.log` | 23,909 | `2d48a221ca7da6cda46a3451e553254817c9a10d50c9284f3bbc1a3fa4547c3e` |
| `01-check-common-attempt-actual-f5-01-cargo-targets.tsv` | 115 | `3bac3ecd6d2d92e2729407d62173806974755a4049e4b0104032cbde6e4a39ae` |

The source-only repair keeps the immutable `96e502277ff67bf700cd1e2d027340ba6e0eca2d`
changes and completes the two remaining recipient builders, using
`impl Debug for Edit<'_>` as suggested by the compiler.  No new compiler,
runtime, rustdoc, or performance command was run for this follow-up.
