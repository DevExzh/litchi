# 0517 shared OPC source-XML publication review

Status: **approved for semantic safety; focused tests reviewed with no blocker**.

Review target: the working-tree change to
`crates/litchi-opc/src/source_backed.rs` described by
`candidate-plan.json`, against base `201f65094ff0ca6186fbf1e2ffa6b3641290253b`.
The production change replaces the late
`SourceXmlPart::check_for_publication` calls for changed source-XML
replacements and source-XML additions with `check_source_state`, while
retaining the earlier complete checks. `git diff --check` passed. This review
did not build or run tests, as requested.

I reviewed the publication map and protocol, the source-proof owners in
`litchi-opc`, and accepted ADRs 0001, 0003, 0005, 0006, 0010, 0011, and 0024.

## Proof that the removed scans are duplicates

The replacement path still calls
`SourceXmlPart::check_for_replacement` in the pending-replacement loop
(`source_backed.rs:6605-6613`). That call checks destination freshness,
source lineage and version, destination Part identity, exact original
destination bytes, content type, destination `ReadLimits`, complete source
XML validity, and source state after validation.

The source-XML addition path still calls
`SourceXmlPart::check_for_publication` before topology planning
(`source_backed.rs:6529-6538`). It checks source lineage authority, source
freshness/version and cancellation, destination content type and limits, and
complete source XML validity. `try_add_source_xml_part` derives the addition
content type from the proof, so the later call has no independently mutable
destination type to validate.

`SourceXmlPart` retains its payload in `Arc<Vec<u8>>` and its original in
immutable `PartData`; after `XmlSplicePublication::finish` issues a derived
value, there is no public or internal mutation path for the payload. The
destination `ReadLimits` are a `Copy` value held by the consuming package.
Therefore the content-type, destination-limit, XML, and lineage checks above
remain authoritative after the topology planning between the two call sites.

## Safety and failure behavior

`check_source_state` retains the late source fence: it checks the source
version, the source execution context, and cancellation immediately before
the proof is registered for transfer. The existing
`TransferSourceCheckedSink` still checks every bounded sink write and flush,
and the final transfer-source check remains in place. The destination source
is still fenced by `SourceCheckedSink` and `finish_source_publication`.

The change does not move the first output write earlier. All topology,
signature, encryption, opaque-member, relationship, output-limit, and source
proof decisions still happen before the preservation writer is called. If a
source or cancellation failure occurs after bytes have been accepted, the
existing writer and `IncompleteOutput { written, source }` mapping are
unchanged. Signed-source policy checks also precede the reviewed late calls
and are untouched.

The expected accounting change is lower cumulative `Resource::Work` and lower
transient XML-validation memory because one duplicate scan is removed. The
initial full validation still charges the destination limits and all XML
validation work once; the late state-only fence performs no XML parse or
second limit charge. This is compatible with monotonic budget accounting, but
the candidate evidence and focused tests should accept the new lower charge
instead of asserting the old duplicate-scan total.

No public signature, source ownership, preservation, cancellation, signature,
or partial-sink contract is changed by this diff. I found no semantic blocker.
The candidate remains conditional on the existing first full checks staying in
place and on the planned focused OOXML regression suite and matched
performance campaign passing.

## Focused-test review

The new work-accounting test is meaningful: it measures the source execution
context after capture and requires exactly one `SOURCE_XML.len()` charge during
the addition publication. A second complete late scan would charge the same
payload again, so this assertion distinguishes the candidate from the old
path. The tightened depth and event assertions also verify that the retained
initial destination-limit proof remains the error owner and emits no output.

The version and cancellation tests reach the late fence as intended. Each
`check_source_state()` calls the provider through `ensure_current_public()`
once; its later comparison uses the captured `SourceSnapshot::version()` token
and does not call the provider again. The initial publication proof therefore
uses two provider version calls, and `arm_after_versions(2)` triggers the
third call at the transfer-boundary `check_source_state()`. The tests verify
that a source change or cancellation is rejected before the first output byte.
These tests primarily protect the retained late freshness/cancellation fence;
the work-charge assertion is the test that distinguishes removal of the
duplicate XML scan.
