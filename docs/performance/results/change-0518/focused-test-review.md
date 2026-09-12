# Focused test review — change-0518

This is a read-only audit of the frozen DOCX/OPC snapshot-reuse candidate. The
reviewed implementation and focused-test paths are:

* `crates/litchi-docx/src/source_backed.rs`
* `crates/litchi-docx/src/document/transaction.rs`
* `crates/litchi-docx/tests/source_backed_managed_document_edit.rs`
* `crates/litchi-opc/src/source_backed.rs`
* `crates/litchi-opc/src/source_xml_hint_tests.rs`
* `crates/litchi-opc/tests/source_xml_hint.rs`

The accepted constraints in `snapshot-reuse-design.md`, together with the
preceding OPC and source reviews, were used as the contract. No source edits,
builds, tests, or captures were performed while writing this review; the
receipts below record the prior focused runs.

## Assessment

The four DOCX additions and the 11 public plus four private OPC tests exercise
the admission, mismatch, freshness, error, and ownership guards. I found no
correctness blocker caused by an accidental pass. The DOCX success test alone
cannot prove that the internal reuse branch ran, because the old full snapshot
rebuild could produce the same output. The OPC public hit test supplies the
stronger observable evidence (retained allocation and no extra memory), while
the source trace below verifies the DOCX handoff and its fallback ordering.
Native and profile evidence remains the authority for the performance claim.

## Guard audit

| Guard | Focused evidence | Why the test is meaningful |
| --- | --- | --- |
| Managed DOCX changed publication, readback, physical preservation, and release | `managed_changed_publication_preserves_readback_and_releases_snapshot_budget` checks the changed commit source, changed paragraph readback, output main XML, untouched media and opaque members, live memory while the result and commit are retained, and zero memory after both are dropped. | It covers the public publication contract and the complete reservation lifecycle. The source path confirms that a matching `Patch::source()` is offered to the hint seam; output equality by itself is intentionally not treated as proof of a hit. |
| Foreign DOCX source with equal main XML | `managed_foreign_equal_main_xml_refuses_commit_before_output` opens two managed packages with equal main XML but different media and opaque members, publishes a commit from the first through the second, and requires `StaleSource`, an empty sink, and released budgets. | A byte-only or source-version-only shortcut could accept this fixture. Distinct `SourceLineage` values reach the identity gate, and the final `Patch::apply` check prevents output even though the main XML is equal. |
| DOCX source change after the current Part read begins | `managed_source_change_during_current_part_read_refuses_before_output` uses `SourceCacheLimits::new(1, 1)` to force a physical read, arms the mutable source only after the commit is captured, and advances the revision after bytes are copied. | The trigger is after a non-empty read, so the result tests the post-read freshness fence on the repeated current-proof read rather than a pre-cancelled or pre-stale source. The typed source error, empty sink, and zero memory preserve the existing before-output boundary. |
| DOCX cancellation during the current Part read | `managed_cancellation_during_current_part_read_refuses_before_output` uses the same one-byte cache and cancels from the source after a non-empty read. | This reaches the cancellation fence in the current proof read and verifies typed cancellation, output suppression, and cleanup. |
| OPC original hint hit | `matching_original_hint_reuses_exact_bytes_without_extra_budget_work` warms the comparison baseline, captures an original hint, and checks the returned bytes, pointer identity, unchanged memory reservation, and conservative comparison Work against the cold source-XML Work. | Pointer identity and unchanged memory while the hint is retained distinguish a reused proof from a fresh metadata/proof allocation. The test also proves that the normal public API still performs bounded accounting for the comparison. |
| OPC identity and policy admission | The private `each_identity_and_policy_field_is_required_for_a_hint_hit` perturbs lineage, version, Part URI, content type, and limits one at a time. Public tests cover equal-version foreign lineage, different Part/content type, and different limits. `original_bytes_are_checked_after_all_metadata_matches` supplies a valid but different allocation after all metadata fields match; `equal_bytes_in_a_distinct_allocation_still_take_the_hint_hit` covers the equal-byte comparison path. | These cases prevent any one omitted predicate, version-only authorization, pointer-only equality, or metadata-only acceptance from passing. The private field mutation is deliberate: those fields are opaque at the public boundary and the test is checking the owner’s predicate directly. |
| Derived OPC hint safety and reservation retention | `derived_hint_returns_current_original_and_keeps_derived_bytes_live` creates a checked splice through `checked_range`, `into_publication`, and `finish`, then passes it as a hint. It requires current original bytes, rejects the derived bytes as the current proof, and verifies that the derived reservation remains live until the derived token is dropped. | The candidate payload cannot authorize itself. This directly exercises the `payload == original` admission guard and the miss path’s fresh validation without silently discarding caller-owned derived memory. |
| OPC fallback validation, freshness, security, and Work refusal | `malformed_current_source_is_revalidated_after_foreign_hint_mismatch`, the before/during revision tests, the before/during cancellation test, `signed_and_encrypted_packages_refuse_hints_before_payload_transfer`, and the private `eligible_hint_comparison_consumes_its_bounded_work_charge`. | A mismatch is passed to the existing validator, source/context fences remain typed and interruptible, policy refusals happen before payload transfer, and an exhausted comparison budget is returned as a resource error instead of being hidden as a fallback. |

The source path matches these stimuli. OPC performs read-ahead disablement,
source/context checks, security checks, Part lookup and XML classification,
one normal `read_part`, and post-read checks before considering a hint. The
hint predicate requires current lineage and version, the embedded proof's
matching lineage and version, equivalent Part URI, content type, `ReadLimits`,
and original-payload ownership. It then charges bounded comparison Work,
checks allocation identity or compares bytes in interruptible chunks, and
fences the source/context again. Any mismatch passes the already-read
`PartData` to the normal full validator, so it does not authorize hint bytes or
introduce a second Part read.

DOCX supplies a hint only for a managed, source-authorized changed commit whose
retained `before` snapshot has the current source identity. After OPC returns a
current proof, `reuse_if_source_xml_matches` checks the snapshot identity and
exact bytes, using pointer-plus-length equality where possible and an exact
comparison for distinct allocations. Only that branch returns a clone of the
retained snapshot. Otherwise DOCX validates the fresh proof and reconstructs a
new snapshot before `Patch::apply` retains the stale-source decision.

## DOCX derived-commit boundary

There is no public API that constructs a reverse `Commit` from a derived
snapshot for the requested “derived inverse publication stale on original”
scenario. `publish_document_commit_to_stream` accepts `&Commit`; the public
`Patch::inverse()` returns a `Patch`, not a `Commit` wrapper. The separate
`publish_document_inverse_to_stream` API accepts a `DocumentPublication` and
restores the authenticated artifact, so it does not exercise the DOCX current
snapshot handoff.

The underlying OPC proof is intentionally opaque and
`SourceXmlPart::into_publication` rejects a derived payload from being edited a
second time. The normal managed DOCX edit path creates a derived candidate
snapshot, while `commit.patch().source()` remains the original source proof.
Constructing a test-only reverse `Commit` or exposing internal snapshot state
would therefore invent an unsupported public test API. No DOCX feature or
integration test is required for that case. The derived authorization guard is
covered directly at the OPC boundary, where the public
`source_xml_with_hint` method can receive a real derived hint.

## Recorded focused runs

The receipts were produced against unchanged source manifests and all have a
zero process exit code:

* `focused-docx-check/receipt.json` records `cargo test --locked -p
  litchi-docx --test source_backed_managed_document_edit --
  --test-threads=2`; its log reports 40 passed, 0 failed.
* `focused-opc-check/receipt-0.json` records `cargo test --locked -p
  litchi-opc --test source_xml_hint -- --test-threads=2`; its log reports 11
  passed, 0 failed.
* The initial private-unit selector in
  `focused-opc-check/receipt-1.json` completed successfully but its log reports
  0 passed and 391 filtered out. It did not cover the new private module at
  that source snapshot and is not coverage evidence.
* The corrected run in `focused-opc-unit-check/receipt.json` uses the same
  `--lib source_xml_hint_tests` selector after the module was registered; its
  log reports four private tests passed and 391 filtered out. The four names
  are all under `source_backed::source_xml_hint_tests`.

This review did not rerun any of these commands. The initial zero-test result
is retained as an audit correction rather than being counted as a passing
focused unit suite.

## Remaining limit and disposition

The focused tests do not instrument the private cache counter or count calls to
`read_part`; those properties are established by the source trace and belong
in the existing matched/mismatched profile and allocation gates. The tests do
show the observable distinction that matters at the public boundary: a valid
original hit keeps the proof allocation and reservation, while every foreign,
derived, stale, malformed, policy, or budget case follows its guarded path.

The four DOCX tests and the OPC 11-plus-4 focused tests are suitable release
gates for the frozen candidate. The absent DOCX derived-commit integration
test is an API boundary decision, not a coverage defect or a request for a new
feature.
