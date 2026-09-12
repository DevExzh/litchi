# DOCX publication snapshot reuse design

This is a bounded, read-only design review for the next OOXML optimization.
It is based on revision `afb62ab7a70859dad4a4b9c8eea91402d1ef4052`, the
current change-0518 plan, the accepted ADR manifest in this directory, the
DOCX/OPC source, and the scoped change-0517 publication profiles. No
production code, benchmark harness, build, or capture was changed by this
review.

## Finding

The repeated current snapshot is a real duplicate of work already completed by
the edit. A managed source-backed `document_snapshot` has already:

1. captured an OPC `SourceXmlPart` after source, package-security, content-type,
   part-size, and complete source-XML checks;
2. rejected the DOCX markup-compatibility projection and performed the bounded
   DOCX source-document scan; and
3. built the managed `Snapshot` indexes under the package execution context.

That snapshot is retained as `Patch::before` through the commit. Its source
identity contains the opened-package `SourceLineage`, `SourceVersion`, and
equivalent main-document `PartUri`; its `SourceXmlPart` retains the exact
original main-part bytes and source owner. Reconstructing a second current
`Snapshot` in `publish_document_commit_to_stream` repeats the DOCX scan and
index construction merely to supply the left side of `Patch::apply`.

The full reuse is coherent if the OPC owner first reauthorizes the retained
source proof against the current main Part. The reauthorization must read the
current decoded Part through the normal cache/source path, check its exact
bytes and immutable identity, and retain every source and security fence. A
lineage or version check alone is insufficient, and using the candidate as the
current snapshot would hide a stale source.

The recommended candidate is therefore a two-layer handoff:

* `litchi-opc` gets a narrow source-proof hint operation that performs the
  existing source XML Part preconditions and one current `read_part`, but can
  reuse an already validated original hint when its metadata and exact bytes
  match. A mismatch uses the same freshly read `PartData` to construct the
  existing fully validated `SourceXmlPart`; it must not read the Part twice.
* `litchi-docx` compares that returned current proof with `Patch::before`. On
  an exact identity/byte match it reuses `Patch::before` as the current
  `Snapshot`. Otherwise it builds the current snapshot from the already
  captured proof and lets the unchanged `Patch::apply` return
  `TransactionError::StaleSource` where it did before.

This removes the duplicate DOCX validation/index work on the normal managed
changed-document path while retaining the physical source read and the final
OPC destination proof. It is a stronger target than only calling the existing
`PartView::source_xml`, which would still repeat the OPC XML validation scan.

## Existing path and exact duplication

`Package::publish_document_commit_to_stream` currently calls
`main_document_snapshot("publish_document_commit_to_stream", commit.patch().changed())`
and then `commit.patch().apply(&current)`. For a changed, ordinary managed
source:

* `main_document_snapshot` calls `main.source_xml()`;
* `PartView::source_xml` disables read-ahead for publication, checks source and
  execution state, rejects encrypted and signed physical packages, checks XML
  classification, reads the main Part, and runs `SourceXmlPart::new`; and
* DOCX runs `ensure_source_document_xml` and
  `Snapshot::from_source_xml`, which allocates and scans a second managed
  paragraph/table/block index.

The edit that produced the commit already did the same source capture and
DOCX scan for `Patch::before`. The changed candidate is separately built by
`reconstruct_managed_paragraph_operations`, reparsed with
`Snapshot::from_source_xml_with_identity`, and semantically read back before
`Edit::commit` returns. `Edit::commit` retains that source-preserving candidate;
publication does not need to reparse it.

The scoped change-0517 Callgrind data gives the size of the opportunity. The
publication method's `main_document_snapshot` subtree was:

| workload | publication summary IR | current snapshot subtree IR | share |
| --- | ---: | ---: | ---: |
| p128, K1, owned, batch | 5,467,082 | 2,953,819 | 54.03% |
| p512, K1, owned, batch | 17,749,121 | 11,657,641 | 65.68% |
| p512, K32, owned, batch | 17,749,977 | 11,657,814 | 65.68% |

These are prior scoped instruction counts, not a current candidate result or
a speedup claim. They show why it is worth evaluating the complete proof reuse
instead of stopping after a smaller DOCX-only shortcut.

## Proposed OPC hint contract

Add a low-level, source-proof-preserving operation on `PartView`, for example:

```text
source_xml_with_hint(&self, hint: &SourceXmlPart) -> Result<SourceXmlPart>
```

The operation is still owned by `litchi-opc`; `SourceXmlPart` remains an
opaque, constructor-free proof type. The argument is an optimization hint and
never authorizes a transfer on its own.

The method should factor the common pre-read checks from
`SourceBackedPackage::source_xml_part` and then perform the following steps in
this order:

1. Disable read-ahead for publication, check the current source version and
   execution context, reject encrypted entries and signature infrastructure,
   resolve the catalog Part, and require the same XML classification.
2. Read the current Part once through `read_part`. This retains the existing
   cache admission, declared-size/decoded-size checks, PartBytes limit,
   source-version fences, cancellation checks, and managed `PartData` owner.
3. Check the source and execution state again before inspecting the result.
4. Accept the hint only when all of these hold:
   * the hint is an original source payload, not a derived splice payload;
   * its `SourceLineage` is the current package lineage;
   * its `SourceVersion` equals the current package version;
   * its Part URI is equivalent to the resolved destination Part;
   * its source content type equals the destination content type;
   * its retained `ReadLimits` equal the current package limits; and
   * its retained original bytes equal the just-read current Part bytes.
5. On a match, return a clone of the original hint. The clone must continue
   to describe the original source payload, with no derived payload
   reservation. No second XML validation is needed: this exact source was
   already admitted by the source XML validator and the DOCX scan at snapshot
   creation, and the immutable identity/limits match proves that evidence
   applies to the current Part.
6. On any non-error mismatch, call the existing
   `SourceXmlPart::from_source_parts` with the already-read `PartData`. That
   path performs the complete source XML validation exactly once and returns a
   fresh current proof. A derived hint must take this path; its `original`
   bytes are not allowed to stand in for its derived `payload`.
7. Check source and execution state before returning either proof.

The current `ReadAt` contract requires an adapter to advance its
`SourceVersion` whenever observable bytes change. The exact byte comparison is
still required: it protects the patch's source precondition and handles a
source that reports a different current allocation with the same version. A
cache hit is an exact immutable allocation under the existing source-cache
contract; a non-identical allocation takes the byte comparison path. A source
adapter that changes bytes without updating its version is already outside the
`ReadAt` contract, and the existing cache would have the same limitation.

`ReadLimits` is `Copy + Eq`, so comparing the complete value is cheap and
prevents a proof produced under one XML/Part policy from authorizing a package
with a different policy. `SourceLineage` is pointer identity, so equal caller
chosen `SourceVersion` values from two opened packages do not cross the proof
boundary.

The public low-level method is consistent with ADR 0001: it exposes no archive
implementation type, source generic, lock, or mutable byte escape hatch, and
it accepts only an OPC-owned proof that has no public constructor. If keeping
this seam private is preferred, the same operation must remain in `litchi-opc`
and be reached through an explicitly owned DOCX/OPC publication helper; DOCX
must not inspect `SourceXmlPart` fields itself.

## Proposed DOCX handoff

Add crate-private transaction helpers with no new ordinary CRUD surface:

```text
Snapshot::source_xml_hint() -> Option<SourceXmlPart>
Snapshot::source_identity_matches(lineage, version, partname) -> bool
Snapshot::matches_source_xml(lineage, version, partname, bytes) -> bool
```

These helpers should remain in `document::transaction`; they only expose a
proof and content-free identity checks to the sibling source-backed publisher.
They must not expose `SourceIdentity` fields publicly.

For a changed commit whose `Patch::before` has a source-backed hint, the
publisher should retain the current error and validation order:

1. Check `package.source_version()` and `package.check_execution()` before
   looking up the main Part. Resolve the Part and validate its DOCX main-part
   content type as today.
2. If the patch source identity does not match the current package lineage,
   version, or main Part URI, use the existing
   `main_document_snapshot` path. This preserves foreign-patch and signed
   package error precedence rather than allowing the hint operation to report
   a different error first.
3. Otherwise call `main.source_xml_with_hint(&hint)`. The OPC operation either
   returns the retained proof on an exact match or a fully validated fresh
   current proof after a mismatch.
4. Compare the returned proof's bytes and source identity with
   `Patch::before`. On a match, use `Patch::before.clone()` as `current`. This
   is the only branch that skips `ensure_source_document_xml` and
   `Snapshot::from_source_xml`; the proof and the retained snapshot are from
   the same immutable source identity and exact bytes.
5. On a mismatch, run `ensure_source_document_xml` on the returned fresh
   proof, build the current `Snapshot` with the existing constructor, and call
   the unchanged `Patch::apply`. This keeps stale bytes and foreign lineage
   as `TransactionError::StaleSource` while avoiding a second Part read.
6. Retain `document_policy::validate_changed_document`, candidate source
   proof checks, topology planning, source monitoring, and all output guards
   exactly where they are today.

The source-identity gate before the OPC hint operation matters. For example,
an equal-main-XML commit from a foreign package must continue through the
ordinary current-snapshot path and fail `Patch::apply` as stale. A signed local
package must continue to use the established signed/no-op behavior. A
source-version change must win before any hint or patch identity decision.

The no-op branch remains unchanged. It deliberately calls the non-authorized
current snapshot path so an exact signed or malformed source can still be
copied byte for byte under the existing contract. The fast path is only for a
changed commit with a source-backed `Patch::before`; a signed managed source
that fell back to `Snapshot::from_managed_part` has no source XML hint and uses
the existing path.

## Proof obligations

| Contract | Evidence required from the candidate implementation |
| --- | --- |
| Current source | Initial and final `source_version`/execution fences, plus the OPC hint's before/after read fences and the existing topology transfer fences. |
| Exact main Part | One normal `read_part`; exact `PartData` allocation or byte equality against the hint's original bytes; mismatch falls through to full validation and `Patch::apply`. |
| Lineage/version/Part | Source lineage pointer equality, exact `SourceVersion`, equivalent Part URI, and a DOCX `Snapshot` identity check before reuse. |
| Limits and budget | Current package `ReadLimits` and hint limits agree; the normal Part read retains PartBytes, cache, memory/object, work, cancellation, and source reservations; skipped scans are already retained by the before snapshot. |
| XML and DOCX safety | The hint can only be issued from a snapshot that passed `SourceXmlPart::new` and `ensure_source_document_xml`; a mismatch uses the original full validator and DOCX scan. |
| Signature/encryption | The hint operation repeats `disable_read_ahead_for_publication`, encrypted-entry rejection, signature-infrastructure rejection, XML classification, and context checks. No-op behavior remains on the old path. |
| Changed-document policy | Settings/protection/tracked-revision inspection remains in `validate_changed_document` after patch application. |
| Candidate correctness | Candidate reconstruction, `Snapshot::from_source_xml_with_identity`, selected paragraph semantic readback, and `Commit` construction remain unchanged. |
| Physical preservation | `SourceXmlPart::check_for_replacement`, destination original-byte comparison, topology limits, source monitoring, ZIP preservation, and sink handling remain unchanged. |
| Error precedence | Source changes and cancellation before output keep their existing typed errors; foreign/byte-mismatched hints fall back to the unchanged `Patch::apply` stale check. |

The skipped `ensure_source_document_xml` and layout scan are evidence reuse,
not a weaker validation mode. The candidate must document that the reused
snapshot is accepted only after the exact current Part proof and only because
the retained snapshot already passed both scans under the same source owner
and limits.

## Existing coverage and required additions

The following tests already cover important invariants and should remain green:

* `crates/litchi-docx/tests/source_backed_managed_document_edit.rs`:
  `managed_source_change_refuses_publication_before_sink_output`,
  `managed_cancellation_refuses_edit_and_publication_without_output`,
  `managed_signed_noop_and_equal_setter_preserve_exact_source`,
  `unmanaged_foreign_package_with_equal_main_xml_refuses_stale_commit_before_output`,
  `managed_disjoint_updates_publish_in_arbitrary_order_and_inverse_exact`,
  `managed_changed_publication_enforces_settings_protection_flags`,
  `managed_changed_publication_rejects_settings_mce_and_dtd_before_output`,
  and `managed_publication_reports_a_partial_sink_failure`.
* `crates/litchi-docx/tests/source_backed_managed_paragraph_batch.rs`:
  `managed_batch_forward_and_inverse_are_source_checked`,
  `managed_batch_noop_and_revert_restore_the_exact_source_owner`, and the
  cancellation/work/memory refusal tests.
* `crates/litchi-docx/tests/source_backed.rs`:
  `stale_commit_and_changed_source_fail_before_output` and
  `signed_changes_are_refused_but_signed_noops_are_exact`.
* `crates/litchi-docx/tests/sequential_text.rs`:
  `source_stale_before_during_and_after_output_takes_precedence` and
  `source_change_overrides_simultaneous_sink_failure`.
* `crates/litchi-opc/tests/source_xml_publication.rs` and
  `source_part_splice.rs`: source lineage, source-version, cancellation,
  limits, exact source ranges, signed no-op/change, and partial-output guards.

Add focused coverage for the new seam before accepting it:

1. A managed changed commit takes the hint hit, publishes the same output,
   retains selected-paragraph readback, and releases the before/candidate
   budget reservations after drops.
2. A foreign lineage with equal main XML, a different Part URI, a different
   content type, and a different `ReadLimits` value never takes the hit and
   retains the existing stale/error result before output.
3. A source revision change before and during the hinted Part read returns the
   existing typed source-change error with an empty sink; cancellation and
   memory/output limits retain their existing resource and cleanup behavior.
4. A same-version different decoded payload takes the fresh-proof fallback,
   then `Patch::apply` rejects it as stale. The test adapter must obey the
   normal source/cache contract; if it mutates without a version change, the
   test must explicitly document that it is outside `ReadAt` guarantees.
5. A derived `SourceXmlPart` hint is never treated as an original source
   payload. The OPC helper must use a fully validated fresh proof in that
   case, and no derived output reservation may be silently discarded.
6. Signed and encrypted inputs preserve their changed/no-op policy, including
   the existing exact signed no-op path. DTD, MCE, malformed XML, protection,
   tracked-revision, unsupported transfer, and sink-failure tests remain
   unchanged.

The mismatch tests are as important as the hit test: they prove that the hint
is merely a fast proof reuse and cannot turn a conflict into a publication.

## Measurement gate and decision

The candidate must be measured at the existing change-0518 publication
boundary with the same 24 native cases, two reversed-order campaigns, and the
separate six-case hardware lane. Add a Callgrind annotation or retained-symbol
check showing that a hit executes one `source_xml_with_hint` call and does not
enter `Snapshot::from_source_xml` for the current snapshot. The fallback path
must be profiled separately so its complete XML validation remains visible.

Compare publication phase, lifecycle, RSS, source/cache counters, managed
resource charges, output bytes, and all correctness fields. Do not mix
instrumented Callgrind or whole-child hardware measurements with the native
publication timer, and do not claim a speedup from the prior change-0517
instruction counts. Accept the optimization only if the source/resource
review and focused guards pass and a matched candidate campaign shows a
material improvement without a regression beyond the change-0518 thresholds.

If the low-level hint seam is rejected during implementation review, the
fallback candidate is to call the existing `main.source_xml()` once, compare
its exact proof with `Patch::before`, and reuse the retained DOCX `Snapshot`
only on a match. That partial candidate preserves the same proof matrix but
continues to pay the OPC XML validation scan. It should be treated as a
fallback measurement, not as evidence that the larger repeated current
snapshot work is solved.
