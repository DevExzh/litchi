# 0806 after-build readiness review

Review status: **PASS for the after-build handoff**. This is a bounded static
review of the retained application chain, source manifests, build receipts,
quality receipts, and the frozen before qualification witness. I did not run a
build, capture, Cargo command, profiler, or heavy replay. Main after timed
captures have not started, and the cross-format after build is still open.

## Application chain and source scope

The source revision remains `e3ff267ee3454e71d66f177f54f3cd05e0d9cce5` and
every application in the chain retains 9,196 source entries. The descriptor
links are exact:

`application.json` → `quality-amendment-application.json` →
`visibility-amendment-application.json` → `after-build-inputs.json`.

The first application changes the six recorded candidate paths: the five
helper modules and the shared OPC attribute test. The constructor quality
amendment changes exactly the five helper modules and preserves the shared test
byte-for-byte. The visibility amendment changes exactly the OLE helper and
restores its two public declarations. Comparing the original production source
map with the final visibility application gives exactly these six changed
paths; all other source entries retain their original hashes:

| Path | Original source SHA-256 | Final source SHA-256 |
| --- | --- | --- |
| `crates/litchi-ole-common/src/xml_attributes.rs` | `067bdb3d…` | `9076f5f9…` |
| `crates/litchi-opc/src/xml_attributes.rs` | `6c7e5aeb…` | `7409a007…` |
| `crates/litchi-opc/src/xml_attributes/tests.rs` | `6c7d0247…` | `e26c97c8…` |
| `crates/litchi-sign/src/xml_attributes.rs` | `7e389b5a…` | `02c0e5f9…` |
| `crates/litchi-xldm/src/xml_attributes.rs` | `7e389b5a…` | `02c0e5f9…` |
| `crates/xml-minifier/src/xml_attributes.rs` | `7e389b5a…` | `02c0e5f9…` |

The current workspace hashes for all six paths match the final application
source witness. The before and after frozen-input records are byte-identical,
including the architecture, inheritance, origin, analysis, policy, host,
probe-quality, profile, and quality driver hashes. No unrelated source or
architecture input entered the after build.

## Quality and build receipts

The six production quality rows in `quality-2/checks.json` all exit zero:
format, all-feature/all-target check, tests, warnings-denied Clippy, docs, and
crate-boundary validation. The test log contains 500 suites with 12,885 passed,
0 failed, and 89 ignored tests. The three after probe-quality rows also exit
zero, with 36 passed tests and no failures.

The after build has three successful locked offline rows for the native,
allocator, and profile binaries. Its source census matches the final
visibility application and `after-build-inputs.json`; its probe and lock
witnesses match the before build. The native and profile logs each contain the
known 11 unused allocation-metric warnings from compiling feature-separated
probe variants. The production warnings-denied Clippy row passed, and the
allocator build has no corresponding warning. These warnings do not change the
production quality result or add a resource claim.

## Qualification freeze and after-build chronology

`qualification-four-attr.json` is explicitly frozen at the
`after-before-qualification` stage with `after_before_qualification=true`,
`before_after_build=true`, and `timings_imported=false`. Its baseline source
and binaries match the sealed before build. It records semantic and
preservation oracles only; no timing values are imported into the after build.

The wrapper's `after-build-inputs.json` binds the exact qualification file, the
visibility application, and the final 9,196-entry source map. Its freeze time
is `2026-09-28T03:40:53.475166Z`; the after compilation begins at
`2026-09-28T03:40:53.702109Z`, after the completed quality-2 gates. The after
build source census equals this frozen source map, so the binaries were built
from the visibility-restored source after the before-only qualification was
sealed.

## Disposition

The source and build handoff is ready for the root-owned cross-format gate and
the subsequent main after captures. This review makes no timing, allocation,
profile, cross-format, or production-adoption claim. Those gates and the final
seal remain open.
