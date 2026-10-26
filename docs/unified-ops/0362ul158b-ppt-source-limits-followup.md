# 0362ul158b — PPT source-limit follow-up

Date: 2026-10-26. Parent: 0362ul158. Status: source changes are ready for
coordinator verification; no Cargo, rustc, native, archive, or gate command
was run in this follow-up.

## Source binding

This leaf remains `uo2/ppt-hyperlink-crud-repair`, based on frozen C
`f8cd398a770f7363f628be40edb3ce4219566c16`. The preceding source-only
propagation commit is `8aa693d5db89d42683d291fcd75a4b350ab5d1e8`. The
coordinator must bind this record to the exact commit and use the approved
capped runner for compiled verification. The external operations-layer
OfficeArt classification patch is intentionally separate and must be merged
into the same coordinator follow-up before the combined gate.

## Repair

This follow-up closes the remaining limits-audit findings:

- `text_from_ppt_records`, OfficeArt drawing traversal, child scheduling,
  and retained outline bodies now carry the caller's depth ceiling and return
  typed `SourceLimitKind::Depth` errors when it is reached. The embedded
  record and OfficeArt host loops return typed `SourceLimitKind::Events`
  errors, including overflow, rather than converting exhaustion into
  `Corrupted`.
- Speaker-notes shape enumeration and body-text collection use the retained
  source event budget across the complete host walk. Body text is decoded
  through the bounded textbox path with the retained record and depth
  ceilings; it no longer reopens a shape through the default text projection.
- A parsed `PictureBullet` retains its `PictureBulletLimits`. Direct BLIPs,
  embedded FBSEs, and delayed FBSEs are reparsed with the intersected raw,
  BLIP, store-entry, and file-count policy. `to_record()` also revalidates
  under that retained policy, so a later read or write cannot silently reset
  to OfficeArt defaults.

No public `*_with_limits` alias, limit-less raw entry point, capability
waiver, or `3rdparty/` fixture was added. The separate operations patch is
responsible for preserving typed OfficeArt limit classification through the
facade; this leaf does not duplicate that change.

## Regression coverage

The source tests add or strengthen these focused cases:

- `litchi_ppt::odraw::tests::embedded_text_records_honor_the_caller_event_limit`
  asserts `Error::Limit { kind: Events, observed: 2, maximum: 1 }`.
- `litchi_ppt::odraw::tests::embedded_text_records_honor_the_caller_depth_limit`
  asserts `Error::Limit { kind: Depth, observed: 2, maximum: 1 }`.
- `litchi_ppt::odraw::tests::admitted_shapes_keep_host_loop_event_limits_typed`
  admits a real shape tree and then exercises nine host records against an
  eight-event limit, asserting the typed exhaustion after root admission.
- `litchi_ppt::outline_text::tests::body_text_honors_the_retained_source_depth_and_event_limits`
  covers outline-body reparsing under retained depth and event policies.
- `litchi_ppt::picture_bullets::record::tests::blip_reuses_the_picture_bullet_admission_limits`
  lowers the retained policy after exact-size admission and verifies that
  BLIP resolution returns `ImageLimitExceeded` instead of reopening with
  defaults.

The prior 0294 hyperlink fidelity regression remains unchanged. Root should
run the focused tests above, the PPT CRUD/hyperlink checks, and the external
typed-error classification test through `run_safe_gate.run_step`; this record
contains no compiled receipts or performance claims.

## Source-only verification

The following checks completed on this leaf:

```text
rustfmt --edition 2024 --check <changed Rust files>
git diff --check
```

No compiled result is claimed. The coordinator owns source/hash binding,
target bounds, process receipts, and all Cargo/native verification.
