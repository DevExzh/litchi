# XLSX Data Model v1 evidence

This isolated candidate records the XLSX Data Model owner batch on committed HEAD `ce0be842b60a28db142b795389e24be5bde8ae07`. The feature delta is exactly ten owner paths, listed in `owner-manifest-v1.json`; the evidence files under this directory bind the source and validation results. The owner blobs are byte-identical to the source blobs used by the root-gated candidate `/var/tmp/litchi-xlsx-model-final-validation-v7-20260910-a`. A byte comparison of all 545 files under `crates/litchi-xlsx` is recorded in `source-audit.json`, so the root gate receipt applies directly to this candidate's package source tree.

The Custom Data and Connections v7 prerequisite is bound to `custom-source-manifest-v7.json`, sourced from `b2e2ff8e1fdabd4ab16a67ddf077e6968bf5e668`. All 17 frozen prerequisite hashes match current HEAD `ce0be842b60a28db142b795389e24be5bde8ae07` before the owner overlay. The 14 prerequisite paths outside the owner set remain byte-identical in the candidate. The three shared paths (`docs/FEATURE_MATRIX.md`, `src/lib.rs`, and `src/package.rs`) are the owner-integrated blobs whose hashes are listed in the owner manifest; their Custom and Survey API markers are checked in `source-audit.json`. The four committed Survey paths remain byte-identical to current HEAD.

The Data Model implementation validates the outer workbook/OPC contract: descriptor shape and bounded XML 1.0 strings, XLDM storage profile and payload limits, the fixed `/xl/model/item.data` part and content type, workbook relationship ownership, recognized table connection references, and bounded source-preserving XML/context publication. Opaque extension markup is retained lexically with namespace and inherited `xml:base`, `xml:lang`, and `xml:space` context, including nested base overrides and RFC 3986 reference resolution. Source/event spans, aggregate encoded XML size, attribute entities, and XML character validity are checked before materialization or output growth.

The binary XLDM payload remains opaque. Creation/import proves only the outer descriptor, storage, package, and relationship contract; it does not prove inner XLDM table, relationship, column, time-group, or dependency identity. A source-bound payload-only replacement with an unchanged typed descriptor is supported. `Transaction::set` refuses structural descriptor replacement without inner identity/dependency proof, including a changed descriptor supplied together with different payload bytes. This is an explicit safety boundary; implementing full inner XLDM identity proof remains open. Query evaluation, refresh, rendering, external connections, and native producer interoperability are outside this evidence.

The authoritative root gate passed 1,437 unit/integration tests and two doctests, with zero ignored tests. Strict all-target/all-feature Clippy with `-D warnings`, warning-denied rustdoc, formatting, and whitespace checks passed. The unchanged root receipt and compressed raw logs are included here. Evidence assembly did not rerun those gates; the recorded commands and results are the root receipt's results. No `--offline` or `--locked` flag was used by the recorded root commands.

## Reproduce the recorded gates

Run from this candidate root with Rust `1.95.0`; the exact command text is also in `commands.txt`:

```text
cargo +1.95.0 test -p litchi-xlsx --all-features
cargo +1.95.0 clippy -p litchi-xlsx --all-targets --all-features -- -D warnings
RUSTDOCFLAGS="-D warnings" cargo +1.95.0 doc -p litchi-xlsx --all-features --no-deps
rustfmt --edition 2024 --config skip_children=true --check <the 9 Rust owner paths>
git diff --check
```

`Cargo.lock.gz` is the lock snapshot used by the root-gated candidate and may be restored as `Cargo.lock` for a replay. Verify source binding with `owner-manifest-v1.json`, `custom-prerequisite-binding.json`, and `source-audit.json`. The compressed logs are `tests.log.gz`, `clippy.log.gz`, and `rustdoc.log.gz`; `root-final-gates-receipt.json` preserves the root result and artifact source hashes.
