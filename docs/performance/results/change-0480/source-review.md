# 0480 DOCX projected-XML sharing source review

This review covers the two-line candidate change in
`crates/litchi-docx/src/source_backed/paragraph_copy.rs`: the changed-copy
publication branch now clones the `Arc<Vec<u8>>` owned by the projected
`Snapshot` and calls `write_part_overlay_shared_to_stream`. The review is
read-only and is a correctness hand-off, not a performance result.

## Ownership and precondition boundary

`Snapshot::xml` is already an immutable `Arc<Vec<u8>>` (paragraph_copy.rs:230-
238). The publication method performs the same complete precondition sequence
before selecting either output branch: it captures the current bounded
snapshot, checks the patch's artifact fingerprint, exact source XML and source
version in `Patch::apply`, rescans the projected XML under the patch limits,
captures the source artifact, and resolves the main Part (paragraph_copy.rs:
1014-1027). The candidate does not move any of those checks into the OPC
publisher or replace them with a byte-slice assumption.

`Patch::apply` has already performed the checked copy needed to construct the
projected snapshot and has scanned that allocation. Consequently the candidate
`Arc::clone(&target.xml)` is an infallible owner clone of validated immutable
bytes. It removes only the second publication-time byte allocation formerly
performed by `checked_clone`; it does not remove the checked allocation or XML
scan in patch application. The observable allocation-failure behavior at this
optional duplicate allocation changes deliberately: an otherwise valid
publication no longer fails solely because the extra duplicate cannot be
reserved.

## OPC equivalence

The owned and shared one-Part entry points both call the same private
`write_single_part_overlay_to_stream` implementation. The only difference is
the `into_shared` closure: `Arc::new` for the old `Vec<u8>` entry point versus
`identity` for the already shared entry point (litchi-opc/source_backed.rs:
7696-7727 and 7756-7828). Therefore both paths use the same:

* replacement length checks in `validate_overlay_limits`;
* selected-Part read and source-integrity checks;
* byte-identical comparison and exact-source no-op path;
* signature refusal and XML audit for changed XML Parts;
* preservation-plan construction, source re-reads, cancellation/context checks,
  output accounting, and sink/incomplete-output mapping.

The changed path stores the same `ChangedOverlayPayload::Shared` form either
way. The preservation writer clones that Arc into the regenerated ZIP entry
and uses the same bounded sink stack and final source-publication decision
(litchi-opc/source_backed.rs:9251-9435). Part names, content types,
relationships, physical member order, compression selection, and untouched
member copying are consequently unchanged.

## Drop and no-op behavior

For a changed publication, the local replacement Arc is retained by the
shared overlay until the selected member is regenerated. The publication then
retains `target` itself, so the projected snapshot remains valid after the
consuming OPC package and temporary overlay plan are dropped. If validation or
the sink fails, the temporary shared owners are dropped during unwinding and
the caller-owned output behavior remains the existing preflight/partial-output
contract.

The outer paragraph-copy no-op branch is unchanged: it writes the retained
`SourceArtifact` directly and never enters either overlay entry point. Its
`target` is still retained in the returned `Publication`. The lower shared
OPC no-op path also drops the decoded comparison payload before exact-source
publication, exactly as the owned route does (litchi-opc/source_backed.rs:
7780-7798). There is no premature drop or use-after-drop window in the
candidate.

## Existing coverage

The lower OPC tests already exercise the relevant shared route directly:

* `shared_single_overlay_matches_vec_changed_and_signed_noop_publication`
  compares changed output with the owned route and compares signed no-op output
  with the exact source;
* `shared_multi_overlay_matches_vec_and_reopens` checks output parity and
  reopened replacement bytes;
* `shared_overlay_preserves_duplicate_and_limit_refusals_before_output`
  checks pre-output refusal and unchanged sink state; and
* `shared_overlay_reports_partial_sink_failure_with_bounded_writes` checks
  incomplete-output accounting and bounded writes
  (litchi-opc/source_backed.rs:16545-16692).

The DOCX source-backed tests exercise the public caller path that now selects
the shared entry point. `copies_exact_fragment_at_first_middle_and_last_source_order_slots`
and `publication_raw_copies_unselected_members_and_inverse_restores_exact_artifact`
cover semantic output, raw preservation, and both inverse routes. The durable
patch test covers exact whole-artifact staleness.
`empty_edit_is_an_exact_noop_and_stale_inverse_writes_nothing` covers the
public no-op branch,
`changed_source_version_is_rejected_before_publication_output` covers source
rejection before sink output, and
`partial_sinks_complete_and_write_zero_fails_without_false_progress` covers
partial, zero-progress, and failing sinks
(`crates/litchi-docx/tests/source_backed_paragraph_copy.rs`: 113-236,
238-258, 601-721).

These tests cover the candidate's changed ownership route through the public
DOCX API while the OPC tests provide the direct old-versus-shared differential.
I found no meaningful functional regression test missing from this two-line
handoff; an allocation-count assertion would duplicate the benchmark's scope
and would not prove publication correctness.

## Review verdict

No source-level correctness blocker found. The candidate preserves the
validation, limits, source fencing, no-op behavior, physical ZIP preservation,
sink errors, and inverse publication contract while eliminating the redundant
publication replacement-byte copy. Root should still complete the candidate
build, focused tests, and measurement gates against the frozen source before
accepting the optimization result.
