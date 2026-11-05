# 0362zaca — legacy raw boundary validation and typed-error repair

Status: coordinator follow-up to `0362zac`; source packet `29a1387a6f` is
repaired here before acceptance. This a-record covers only the concrete
correctness findings from the raw-boundary review. It does not defer or
reclassify the required DOC, PPT or XLS API migration.

## Repairs

`litchi_xls::raw::Reader::revision_log` now requires `&SourceLimits`. The
requested profile is combined component-wise with the profile retained at
ingress, so the lazy reader can narrow a budget but cannot widen one. The four
raw consumers (the three revision-log integration tests and the user-name
consumer) pass explicit limits. The source-limit unit test proves that a
wider request still resolves to the retained profile.

The new DOC, PPT and XLS foreign-error variants keep their typed payload and
`Error::source()` chain, but their newly exposed Display and Debug arms use
fixed, content-free labels. The DOC and PPT core projections delegate native
CFB/IO and unsupported cases where the existing core taxonomy can represent
them; generic foreign limit and parse failures use fixed projection messages
because `litchi_core::Error` has no foreign-source carrier. `litchi-ops`
classifies the original typed leaf before that projection and now delegates
Custom XML, VBA, signature and DOC DataSpaces variants. Its leaf classifiers
cover limits, IO, unsupported-code-page/policy, and patch-conflict categories.

## Validation

The preparation packet was initially checked against the coordinator's local
lock SHA-256 `52ed0cacac15b5e440349bfc0f4aca2665bcf50d59d4321624e596bb410783c6`.
Adding the direct `litchi-vba` edge required the local ignored lock to be
updated offline; no registry versions or checksums changed. The resulting
local lock is SHA-256
`4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1` and its
only diff is the `litchi-ops -> litchi-vba` package edge. The coordinator must
carry this lock edge when assembling the final staging pin.

With the updated local lock, this focused all-feature check passed:

```text
source (deleted-coord-artifact)
cargo check --locked --offline --all-features \
  -p litchi-ops -p litchi-doc -p litchi-ppt -p litchi-xls \
  --message-format=short
PASS — Finished dev profile in 27.46s
```

Focused tests passed:

```text
cargo test --offline --all-features -p litchi-doc --lib foreign_boundaries
  1 passed
cargo test --offline --all-features -p litchi-doc --lib core_projection_uses_typed_or_fixed_categories_without_foreign_text
  1 passed
cargo test --offline --all-features -p litchi-ppt --lib new_custom_xml_boundary_redacts_foreign_payload
  1 passed
cargo test --offline --all-features -p litchi-ppt --lib new_signature_boundary_keeps_typed_core_categories
  1 passed
cargo test --offline --all-features -p litchi-xls --lib new_foreign_boundaries_redact_payloads
  1 passed
cargo test --offline --all-features -p litchi-xls --lib constrained_profile_never_widens_ingress
  1 passed
cargo test --offline --all-features -p litchi-xls --test xls_revision_log
  3 passed
cargo test --offline --all-features -p litchi-xls --test xls_user_names package_owner_accepts_official_rrdhead_code_pages_and_keeps_unicode_user_name
  1 passed
cargo test --offline --all-features -p litchi-ops --lib legacy_
  10 passed
```

The ops test command was first run on the sparse checkout before its tracked
fixtures were materialized and reported four unrelated `No such file or
directory` failures. Materializing the seven small tracked fixtures and rerunning
the identical command produced the 10/10 result above; no test was skipped.
The affected fixtures are not source changes.

Default and signature-enabled ops feature combinations also passed:

```text
cargo check --offline -p litchi-ops --no-default-features --features doc,ppt,xls
cargo check --offline -p litchi-ops --no-default-features --features ppt,sign
PASS
```

Rustfmt and `git diff --check` passed for all changed Rust files. Full
workspace glossary, inventory and coordinator gate execution remains outside
this leaf packet.
