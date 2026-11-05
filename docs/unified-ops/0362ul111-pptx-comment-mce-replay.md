# 0362ul111 — PPTX collaboration replay MCE budget separation

This repair is based on `e130751fc503099c4d5fff74a5746b5099d625e1` and
addresses the two collaboration replay failures in the full-feature runtime:

- `comments::collaboration::tests::replay_presence_validator_ignores_inactive_typed_uri_bytes`
- `comments::collaboration::tests::replay_threading_validator_ignores_inactive_ordinary_values_after_mce`

The selected ordinary-comment parser was charging namespace declarations and
opaque extension attributes against the recipient's tight
`max_string_bytes`.  That made an active root fail at the 128-byte semantic
budget even though the oversized typed or ordinary values were in an MCE
branch that was discarded.  Replay now runs its existing processed grammar
pass with the already-bounded physical `max_input_bytes` ceiling, then applies
the selected semantic budget to active ordinary fields.  Typed presence values
continue to be checked by `parse_presence`; threading's typed numeric fields
remain structurally validated by `parse_threading`.

The raw bounded scan still runs before MCE processing.  Consequently malformed
authored XML and oversized raw attribute values retain physical admission
precedence, including when those values occur in an inactive branch.  The
change adds no second parse or unbounded clone; the existing one processed
projection and grammar pass remain bounded by the source policy.

No Cargo, build, runtime, rustdoc, native fixture, or benchmark command was run
on this source-only leaf.  Changed Rust files were checked with rustfmt and
`git diff --check`; repository evidence-size and binary guards were checked
before commit.
