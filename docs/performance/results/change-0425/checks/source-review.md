# 0425 source review

This is a read-only review of the three production-file changes against the
current worktree baseline. No Cargo command, test command, profiler, or CPU
workload was run by this reviewer. The recorded all-feature PPTX test receipt
reports 844 passed tests, and the recorded OPC/PPTX warning-denied Clippy
receipt passed with the source unchanged.

## Findings

No correctness blocker was found.

`crates/litchi-pptx/src/font/codec.rs:703-716` changes the EOT name decoder
from `chunks_exact(2)` to `as_chunks::<2>()`. The preceding
`eot_utf16` check rejects odd-sized name data at lines 705-708, so ignoring the
tuple remainder at line 712 cannot discard input. `validate_eot` also keeps the
bounded `eot_sized`/`eot_take` cursor checks and the final variable-header
boundary check. The UTF-16 iterator, invalid-surrogate error, allocation shape,
and error ordering are unchanged. This matches ADR 0022's strict EOT and
malformed-payload contract (lines 96-105, 127-138, and 161-171).

`crates/litchi-pptx/src/presentation_properties/metadata/protection/codec.rs:211-240`
has the same safe precondition. `decode_base64` rejects empty or non-four-byte
aligned input at lines 213-217 before allocating or entering the loop; every
iteration at line 219 is therefore a complete four-byte group. The sextet and
padding branches are unchanged, as are output sizing and the per-field length
checks in `parse_verifier`. The existing decoder's permissive padding details
are outside this mechanical change; this diff does not broaden or narrow them.

`crates/litchi-pptx/src/presentation/order.rs:301-315` now assigns
`self.source.source_version` directly while building the committed snapshot.
The field is `litchi_core::SourceVersion`, whose definition derives
`Clone, Copy` (`crates/litchi-core/src/source.rs:19-28`), so this is a value copy,
not a move of source state and it cannot shorten or extend a borrow. The commit
still copies the same lineage and relationships and retains the source snapshot
in the patch immediately afterward. The lifetime elision in
`binding_refs` at lines 672-682 is exactly equivalent to the removed explicit
`'a`: each returned `&str` remains tied to the input binding slice, and both
reference vectors are consumed synchronously by `reorder_slide_bindings` at
lines 263-275. No stale-source, allocation, or ordering behavior changes.

These conclusions are consistent with ADR 0003's immutable snapshots and
atomic commits (lines 8-25), ADR 0005's stable source identity/version and
bounded positional I/O (lines 8-23), and ADR 0006's non-mutating validation and
default protection enforcement (lines 8-17 and 91-94). The direct OPC/PPTX
ownership remains unchanged; no archive implementation or public type is added.

## Focused existing coverage

The recorded all-feature gate already covers these modules. If a focused rerun
is useful after rebasing, the highest-value existing selectors are:

- Font/EOT codec and graph: `font::tests::fresh_font_containers_are_structurally_checked`,
  `font::tests::rejects_malformed_xml_duplicates_and_caps`,
  `font::tests::package_writer_round_trips_strict_graph_and_schema_position`,
  and `font::tests::noncanonical_targets_and_every_main_profile_round_trip`.
  The first selector calls `Data::powerpoint` on the deterministic EOT fixture,
  which traverses all four EOT UTF-16 name fields; the malformed and graph tests
  retain rejection and publication coverage.
- Source order and `SourceVersion`: in the `source_backed_edit` integration
  target, `source_slide_order_move_changes_only_root_and_keeps_payloads_lazy`,
  `source_slide_order_noop_is_byte_exact_signed_and_bad_positions_are_refused`,
  `source_slide_order_refuses_stale_foreign_signed_cancelled_and_partial_output`,
  and `source_slide_order_rejects_limits_malformed_mce_and_invalid_owners`.
  Together these exercise order-only publication, no-op preservation, source
  identity/revision refusal, cancellation, bounds, and metadata refusals.
- Protection codec and callers: the private codec tests
  `parses_and_republishes_a_legacy_verifier`,
  `password_generation_is_explicitly_dependency_bound`, and
  `password_hash_matches_office_password_then_salt_order`, plus the
  `pptx_modification_verifier` integration test
  `presentation_modification_verifier_is_exposed_by_protection_owner`.
  The source-backed refusal selectors
  `direct_transition_refuses_extensions_sound_protection_and_unsafe_targets_atomically`
  and `source_backed_cross_copy_refuses_protection_opaque_and_trailing_data`
  cover the security boundary that consumes the parser.

No additional test or source change is required for this lint-driven batch.
