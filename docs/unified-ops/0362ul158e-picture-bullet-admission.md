# 0362ul158e — retain picture-bullet image admission limits

Date: 2026-10-26. Parent: 0362ul158. Status: source changes are ready for
coordinator verification; no Cargo, rustc, native, archive, or gate command
was run in this follow-up.

## Source binding

This leaf continues the clean 158d source
`3194ef0204976cacfeb6c7239d28d503774bcd6d` on
`uo2/ppt-hyperlink-crud-repair`. The coordinator must bind the compiled run to
the resulting commit and keep 158b through 158e in the same source chain.

## Repair

`validate_officeart_record` now receives the effective
`PictureBulletLimits::image_limits()` policy. The existing single admission
pass uses that policy for the raw OfficeArt record parser and for
`Entry::parse`/`Blip::from_record`, covering the public constructor, source
parser, and bounded serialization without an extra validation or copy.

The previous `Default::default()` image policy was therefore removed. Later
`blip` and delayed-resolution reads continue to use the retained policy as
before.

## Regression coverage

The existing public integration regression
`public_picture_bullet_constructor_rejects_a_record_over_its_limit` remains in
place and checks the externally configurable encoded-record admission bound.
No artificial invalid configuration was added: the remaining image limits are
either already enforced by the outer admitted record or apply to image-store
operations that a picture-bullet OfficeArt record cannot validly enter.

## Source-only verification

The coordinator should run:

```text
rustfmt --edition 2024 --check crates/litchi-ppt/src/picture_bullets/record.rs
git diff --check
```

No Cargo/native/archive/GO command was run. Root owns exact source/hash
binding, target bounds, process receipts, and compiled verification.
