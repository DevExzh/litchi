# 0362ul158c — PPT OfficeArt error classification

Date: 2026-10-26. Parent: 0362ul158. Status: source changes are ready for
coordinator verification; no Cargo, rustc, native, archive, or gate command
was run in this follow-up.

## Source binding

This change is on `uo2/ppt-hyperlink-crud-repair`, continuing committed
source `db70fac03b48d80e97a932bb2f4a2fa1802450e4`. It integrates the reviewed
external patch
`(deleted-coord-artifact)/ppt-officeart-classification.patch`
with SHA-256
`d39c5a54a4c13491ed4a9b098c329842cc4f36dc8ee8b4fe94cf29118c73024f`.

The external patch was prepared against the same `8aa693d5db` base and was
applied only to its three intended consumers. Its invalid-limit fixture was
corrected here to the realistic OfficeArt safety case `value: 65,
maximum: 64`.

## Repair

`litchi_ppt::odraw` now canonically re-exports the OfficeArt `Error`, raw
`LimitKind`, and image `ImageLimitKind` used by facade consumers. The PPT
classifier distinguishes typed OfficeArt resource exhaustion and invalid
policy construction from malformed content:

- `LimitExceeded` and `ImageLimitExceeded` classify as `limit_exceeded`;
- `InvalidLimit` classifies as `invalid_argument`;
- all other OfficeArt failures retain `corrupt_file`.

The typed leaf remains attached to the unified error source. No message
matching, capability waiver, or dependency change was introduced.

## Regression coverage

The new `litchi-ops` test
`error::tests::ppt_officeart_limits_are_not_reported_as_corrupt_files`
covers raw record exhaustion, image-file exhaustion, the realistic depth
invalid-limit case (`65/64`), malformed-shape fallback, and preservation of
the `Source::Ppt` typed leaf.

Root should run this test with the PPT feature together with the 158b focused
PPT tests, the 0294 hyperlink fidelity test, and the relevant CRUD/gate
filters through `run_safe_gate.run_step`. No compiled result or performance
claim is made in this record.

## Source-only verification

The following checks completed after integration:

```text
rustfmt --edition 2024 --check <changed Rust files>
git diff --check
```

No Cargo/native/archive/GO command was run. The coordinator owns exact source
binding, target bounds, process receipts, and compiled verification.
