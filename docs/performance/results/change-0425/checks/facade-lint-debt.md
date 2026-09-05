# 0425 facade strict Clippy debt

The explicit non-iWork root-facade check failed on the clean baseline revision
`340cc91ae2bdec338dfe7682b4b5d8c219a2d288`. This is a retained lint result,
not a production change or a performance result.

The recorded command was:

```text
cargo clippy --locked -p litchi --no-default-features --features legacy,ooxml,odf,rtf,sign,encryption,vba-inspection,formula,automatic-fonts,images,eval,web-functions,markdown,yaml --all-targets --no-deps -- -D warnings
```

Rust 1.98.1 ran with four Cargo jobs from 15:15:08.744 to 15:15:29.082 UTC
and exited 101. The [receipt](facade-clippy.json) records
`source_unchanged: true`; all 6,623 source hashes are identical before and
after the command. The five affected facade files also match their
`340cc91ae` blobs directly:

| File | SHA-256 |
| --- | --- |
| `crates/litchi/src/detection_smart/detected.rs` | `31adf609ea1191dc0ba6d14909de402db5b274139e14d0c6eabbf6210e19d7c2` |
| `crates/litchi/src/document/doc.rs` | `c08a4cca918d1c5032541e4a0c1d56533b6c11df78885739574219393f93fd78` |
| `crates/litchi/src/presentation/prs.rs` | `da8a28e38dad1e191133e401139698fb3cf26c11aebbddf615ea0814c5777bec` |
| `crates/litchi/src/sheet/functions.rs` | `dc5e41036428862caf4e08d6760b0d4d3ebc490bb5696fe37c006bf283cd45e2` |
| `crates/litchi/src/sheet/workbook.rs` | `540b165ebe50f5b0ce2db26d83733f3e49e97f0d60097686a711de5998695294` |

The log contains 18 unique diagnostics in seven categories. Its compiler
summaries report 14 errors for the library target and 18 for the library-test
target because the targets compile overlapping cfg-selected code; these are
not 32 distinct source findings.

| Category | Count | Locations | Review disposition |
| --- | ---: | --- | --- |
| `unusual_byte_groupings` | 1 | `sheet/workbook.rs:1983` | Small literal spelling cleanup; unrelated to the leaf migration. |
| `int_plus_one` | 1 | `sheet/workbook.rs:2083` | Test assertion spelling cleanup; preserve the existing bound oracle. |
| `needless_return` | 11 | `detection_smart/detected.rs:165,245,1580`; `document/doc.rs:764`; `presentation/prs.rs:697`; `sheet/functions.rs:265,268,272,276,280,286` | Defer for a cfg-matrix review. |
| `large_enum_variant` | 2 | `detection_smart/detected.rs:1993,2082` | Defer for an ownership/allocation design review. |
| `useless_conversion` | 1 | `detection_smart/detected.rs:2800` | Local cleanup; unrelated to the leaf migration. |
| `needless_match` | 1 | `presentation/prs.rs:113` | Local cleanup; unrelated to the leaf migration. |
| `field_reassign_with_default` | 1 | `document/doc.rs:2559` (test code) | Test-only cleanup; preserve the limit assertion. |

The `needless_return` reports sit in feature- and platform-gated compatibility
branches. The ODS/ODP helpers select different bodies when OOXML features are
enabled, and the Document, Presentation, and XLSB path facades split Unix or
Windows source-backed handling from portable byte fallback. A mechanical
return removal must be checked across those cfg combinations so fallback
ordering, source-version checks, and typed error mapping remain unchanged.

The two `large_enum_variant` reports are in private source-path handoff enums.
`PptxSourcePathDetection` carries a validated source-backed owner of at least
1,328 bytes, while another variant is 72 bytes; `DocumentSourcePathDetection`
has a largest variant of at least 1,464 bytes versus 32 bytes. Boxing would add
an allocation and change ownership and timing of source-backed handoff. This
is a deliberate layout/performance question, not a reason to rewrite the
facade or any iWork package during the compiler-directed chunk migration.

Keep this check failed and retain the findings as preexisting debt. A later
facade cleanup should fix the local spelling/test diagnostics, review the
cfg-aware returns under the full non-iWork feature matrix, and separately
measure any boxed source-owner design. No lint allowance or completion claim
is authorized here. Standalone performance and native-resave Clippy results
are separate checks and are not covered by this receipt.
