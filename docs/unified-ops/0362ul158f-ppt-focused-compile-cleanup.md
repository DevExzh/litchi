# 0362ul158f — clean the bounded PPT text migration

Date: 2026-10-26. Parent: 0362ul158. Status: source changes are ready for
coordinator verification; no Cargo, rustc, native, archive, or gate command
was run in this follow-up.

## Source binding

This leaf continues committed 158e source
`0d4057395fbf68d9d862709438421b8295c47c06` on
`uo2/ppt-hyperlink-crud-repair`. The coordinator must bind the next compiled
run to the resulting exact commit and retain 158b through 158f in the same
source chain.

## Repair

The focused 158e compile exposed migration leftovers under the crate's strict
unused and qualification lints. This leaf:

- removes the private, default-policy `text_from_textbox_payload` helper and
  its stale crate-private re-export; every active consumer already uses the
  bounded `text_from_textbox_payload_bounded` entry point;
- removes unnecessary `crate::SourceLimits` qualification from the affected
  presentation and outline test helpers.

No lint waiver, limitless alias, or behavior change was added.

## Source-only verification

The following checks completed after the edits:

```text
rustfmt --edition 2024 --check \
  crates/litchi-ppt/src/presentation/model.rs \
  crates/litchi-ppt/src/odraw/codec.rs \
  crates/litchi-ppt/src/odraw/mod.rs \
  crates/litchi-ppt/src/outline_text.rs
git diff --check
```

No Cargo/native/archive/GO command was run. Root owns exact source/hash
binding, target bounds, process receipts, and all compiled verification.
