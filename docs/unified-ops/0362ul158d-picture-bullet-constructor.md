# 0362ul158d — restore checked picture-bullet construction

Date: 2026-10-26. Parent: 0362ul158. Status: source changes are ready for
coordinator verification; no Cargo, rustc, native, archive, or gate command
was run in this follow-up.

## Source binding

This leaf continues `uo2/ppt-hyperlink-crud-repair` at the committed 158c
source `e27d192427a630a28643f35609f1ca37bd43c819`. The coordinator must bind
the compiled run to the resulting exact commit and keep the preceding 158b
and 158c commits in the same source chain.

## Repair

Retaining `PictureBulletLimits` made the formerly four-public-field
`PictureBullet` impossible to construct outside `litchi-ppt`. The checked
public constructor

```text
PictureBullet::new(index, picture_type, unused, officeart_record, limits)
```

restores that capability without a public mutable policy field. It validates
the index, limits, OfficeArt framing, and preferred image kind, consumes the
caller-owned byte vector without copying, and retains the selected policy for
later BLIP resolution and serialization. Existing parsed bullets use a
private admitted constructor after their one existing validation pass, so
parsing does not double-validate or allocate.

Operational limits are excluded from `PictureBullet` and `OutlineTextBody`
content equality; their bytes and semantic fields remain value-compared.
This keeps the source policy from changing equality or readback comparisons.
No public limit-less raw alias or capability waiver was added.

## Regression coverage

The new external integration target
`crates/litchi-ppt/tests/picture_bullets.rs` has two tests:

- `public_picture_bullet_construction_reads_and_serializes_losslessly`
  constructs a PNG BLIP through the public API, checks zero-copy ownership,
  reads it, serializes its atom and collection, reparses the collection, and
  checks the exact OfficeArt payload and image readback;
- `public_picture_bullet_constructor_rejects_a_record_over_its_limit` checks
  the encoded-record limit before construction.

The unit regression
`picture_bullets::record::tests::blip_reuses_the_picture_bullet_admission_limits`
also confirms that equal content admitted under different policies remains
equal while later BLIP reads still use the retained policy.

Root should run these three focused tests through `run_safe_gate.run_step`,
then repeat the 158b/158c focused filters and the existing 0294 hyperlink
fidelity test. No compiled result or performance claim is made here.

## Source-only verification

The coordinator should run the source-only checks after committing:

```text
rustfmt --edition 2024 --check <changed Rust files>
git diff --check
```

No Cargo/native/archive/GO command was run. Root owns source/hash binding,
target bounds, process receipts, and compiled verification.
