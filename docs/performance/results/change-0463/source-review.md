# 0463 frozen source review

This review covers the frozen 0463 candidate at source epoch
`d18406665f4fd9ad76cb0a2530bfe8e7459bf4ada874ea353b435a84ae4f921c` (7,030
captured files). The production change is limited to
`crates/litchi-odp/src/authoring/edit.rs` and
`crates/litchi-odp/src/authoring/mutable.rs`. The proof remains private to the
ordinary ODP authoring commit path. This is a correctness review; it makes no
performance-retention claim.

## Review result

No semantic blocker remains in the frozen source. `PublicationAuditProof` is
minted only after the bounded writer succeeds and its bytes are reopened as an
`OwnedPackage`. The proof requires both the candidate and source archive
owners, and the commit consumes it only while those exact owners are still in
use. A missing source owner, failed accounting, overflow, narrower final
limits, or either `Arc` identity mismatch falls through to the existing full
XML validator.

## Origin and limit proof

`MutablePresentation::serialize_bounded` writes `content.xml`, `styles.xml`,
`meta.xml`, and staged media through `PackageWriter::add_file_with_media_type`.
That writer path calls `verify_authored` before accepting every
XML-classified payload. `AuthoredXmlAccounting::record` uses the same broad
`xml_minifier::audit::package::is_xml_part` classification and records the
actual byte and part counts after each successful write. A non-XML core part,
counter overflow, or missing core member makes the proof ineligible.

Auxiliary source members go through `copy_auxiliary_files_from_except` and the
writer's `ExactSource` origin, so they are copied from the source package
bound by the proof. `preserve_source_manifest` either emits the exact source
manifest after inventory equivalence is checked or lets finalization generate
one. Generated manifests use the same default authored audit; the proof adds a
conservative 32 MiB manifest allowance and one part. A preserved manifest is
source-identical and is already excluded by the existing source-equality rule,
so this allowance is safe over-counting.

The captured writer profile is the default audit profile: bytes 32 MiB,
depth 256, events 1,000,000, attributes 250,000, one token 4 MiB, and text
bytes 16 MiB. The final ODP profile in `edit.rs` is bytes 128 MiB, depth 512,
events 1,000,000, attributes 250,000, one token 16 MiB, and text bytes 128
MiB. Every captured resource limit is therefore equal to or tighter than the
final limit. The checked aggregate byte and part bounds are compared with
`MAX_PACKAGE_BYTES` and `MAX_XML_PARTS`; uncertainty only disables the proof.

## Commit lifecycle and fallback

The initial changed-slide publication reparses the exact writer bytes before
the proof is created and still performs the staged-slide readback. The source
raw-reference precheck remains before the compactness decision. On a proof hit,
only `validate_compact_xml_parts` is skipped. Initial candidate semantic
readback remains before that decision. Embedded-media and removed-media checks,
RDF/chart/design/annotation readback, final snapshot selection and patch
construction retain their existing later positions.

Every later package replacement clears the proof: design, annotations, RDF,
charts, and semantic-content operations all set it to `None`. An exact no-op
still returns before serialization. Thus an intermediate serialization cannot
be reused after the package, source identity, or domain state changes. The
source-less construction test also confirms that a serialization without a
strong source owner cannot mint a proof.

The proof branch introduces no new refusal or error path after successful
serialization. Writer errors, candidate reopen errors, readback errors, and
the old validator's errors retain their existing order. Exact copied XML may
remain noncompact under the same source-equality rule as before; authored and
generated XML must have passed the writer audit first.

The proof is intentionally scoped to the current `MutablePresentation` path.
If a future implementation adds generated XML, checked splices, encryption,
or signing to this path, it must either extend the origin accounting or force
the existing validator. The current bounded writer path has no such producer.

## Focused source coverage

The eight new mutable proof tests cover:

- candidate and source `Arc` identity, including equal-byte replacement owners;
- exact aggregate byte/part boundaries and overflow fallback;
- XML media classification and the manifest bound;
- each narrower document-limit dimension;
- source-less proof rejection;
- a BOM/commented, indented noncompact manifest with exact noncompact
  `vendor/foreign.xml` and opaque auxiliary copy; and
- generated manifest plus XML-classified staged media.

The two edit tests cover the authored fast-path hit for a changed slide and
proof invalidation/fallback after an RDF rewrite. Existing integration tests
also cover noncompact referenced XML preservation, raw custom-manifest and
opaque-member preservation, manifest regeneration after media topology
changes, and exact auxiliary copying (`presentation_edit.rs:157`,
`odp_rich_content.rs:657`, `odp_rich_content.rs:746`,
`odp_rich_content.rs:780`, and `odf-common` writer tests around
`auxiliary_copy_preserves_noncompact_source_rdf_exactly`).

## Receipts

- `checks/owner-clippy-r2.json`: PASS, all targets/all features with
  warnings denied, source unchanged at the final epoch.
- `checks/owner-tests-r1.json`: PASS, release/all-features `litchi-odp`
  suite, 381 passed, 0 failed, 0 ignored, source unchanged at the final
  epoch. This includes the eight mutable and two edit tests above and the
  integration/doctest coverage.
- `git diff --check`: clean at review time.

An earlier owner-test receipt in the same directory reported 194 passing and
one failure caused by the newly added test's invalid RDF predicate fixture.
The fixture was corrected to a namespace-qualified predicate; the final
epoch above passed without a production-source change.

This review does not claim a measured speedup or recommend retaining the
candidate on performance grounds. Those conclusions belong to the separate
harness and retention gates.
