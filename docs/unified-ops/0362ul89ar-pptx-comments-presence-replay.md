# 0362ul89ar — PPTX inactive collaboration replay limits

Date: 2026-10-25

## Qualification failure

The isolated source base was
`ef9470f7da08a131f31600041f9d562203d62bfa` (`test(pptx): record reviewed
typed limit golden`).  The retained full-feature qualification log is outside
Git at
`(deleted-coord-artifact)/0000.log`.
It is 548,761 bytes with SHA-256
`cfe5aa3f9c4082d67019cfc8285a121768d0a5e3063f8e0c064bfcbda60d3fef`.
The capture reported the following failure in this lane:

`comments::collaboration::tests::replay_presence_validator_ignores_inactive_typed_uri_bytes`

The other notes, opened, and slide-sync failures remain separate work and are
not relabeled here.

## Diagnosis and repair

`validate_replay_extensions` already performs a raw bounded scan and runs MCE
selection with the recipient's selected limits.  It then parsed the selected
XML with the ordinary comment parser using the same `max_string_bytes` value.
That parser's aggregate string counter includes namespace declarations and
opaque extension attributes, so a tight semantic string budget could be
consumed by structural/opaque bytes even when a typed presence payload was in
an inactive MCE branch.

Commit `ebc34fe305` adds a replay-only semantic validation pass.  Its ordinary
grammar parser remains bounded by `max_input_bytes`, which is already checked
against the processed output and therefore remains finite.  The selected
string policy is then applied explicitly to ordinary author/comment fields,
and the existing active typed collaboration validation still applies the
policy to `presenceInfo` values.  The raw scan and MCE pass continue to charge
inactive bytes for input, depth, node, event, and MCE limits; no source bytes
are rewritten and no duplicate/first-wins or allocation path was changed.

The exact source patch against the frozen base is retained outside the
repository at
`(deleted-coord-artifact)/comments-replay-89ar-ebc34fe305.patch`.
It is 3,063 bytes with SHA-256
`e79dc72e864a8ddef27833f3bbffd5f17ca19146248d76a9c5ab6c32321025cd`.

## Checks and status

No Cargo command, build, rustdoc, test, benchmark, or Clippy command was run.
The hydrated sparse worktree used the repository `.cargo/config.toml` path.
Source-only `rustfmt --edition 2024 --config skip_children=true --emit stdout`
completed, `git diff --check` passed, and the repository evidence-size and
binary guards were clean.  A fresh composed runtime run is still required
before acceptance.
