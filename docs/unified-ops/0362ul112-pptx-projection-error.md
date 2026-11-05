# 0362ul112 — PPTX capture projection error precedence

This source-only repair starts at `e130751fc503099c4d5fff74a5746b5099d625e1`
(`uo2/coordinator-replay-integration`). The isolated full PPTX runtime had one
opened-capture failure in
`capture_projection_preserves_relationship_and_part_identity_error_order`:
the malformed first slide was combined with a duplicate presentation target,
but strict catalog parsing ran before slide source admission and returned the
duplicate-target error first.

## Repair

Opened capture now obtains and bounds the tolerant catalog, admits the selected
slide source payloads under the retained `SourceLimits` profile, and only then
runs the strict slide-list contract used by edit snapshots. This keeps the
part-count reservation bound before capture allocation, preserves the existing
relationship/identity checks, and makes a selected malformed XML part retain
the documented content-free `Error::Invalid("PPTX XML is malformed")` refusal
instead of being masked by a later catalog defect. The existing relationship,
identity, root, notes, and deferred-name cases remain covered; the malformed
assertions now match the precise typed error variant and fixed reason.

## Source-only checks

`rustfmt --edition 2024 --check` passed for all four changed Rust files, and
`git diff --check` passed. Cargo, native/runtime qualification, rustdoc, and
the combined gate were not run on this leaf. No generated evidence, fixture,
executable, or `3rdparty/` path is included.
