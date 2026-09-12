# Change-0518 live source and proof review

This review covers the applied DOCX/OPC snapshot-reuse integration in the
working tree. The accepted ADR manifest and its review matrix remain the
governing constraints. The reviewed source paths are:

* `crates/litchi-docx/src/source_backed.rs`
* `crates/litchi-docx/src/document/transaction.rs`
* `crates/litchi-opc/src/source_backed.rs`

No source edits, builds, tests, or captures were performed for this review.

## Verdict

The live source/proof integration passes this bounded review with no source
or ownership blocker. It is ready for the independent test and matched
performance gates. This review does not by itself retain the optimization:
runtime guards, exact-output checks, and native/instruction/allocation evidence
still determine whether the candidate is accepted.

## DOCX handoff

`publish_document_commit_to_stream` passes `Patch::source()` to the private
`main_document_snapshot_with_before` path while retaining the existing
`main_document_snapshot` wrapper for all other callers. The fast path is
entered only for a managed, source-authorized operation whose `before`
snapshot has the current `SourceLineage`, `SourceVersion`, and equivalent main
Part URI. A snapshot without a retained `SourceXmlPart`, a foreign patch, or
an identity mismatch uses the old `main.source_xml()` path.

The current Part is therefore authorized before reuse. The returned proof is
checked again against the patch source's identity and bytes by
`Snapshot::reuse_if_source_xml_matches`. A matching immutable allocation uses
pointer-plus-length equality; distinct allocations use an exact byte compare.
Only this branch returns `before.clone()`, so the existing `Patch::apply`
source check remains authoritative. A proof mismatch runs
`ensure_source_document_xml` and `Snapshot::from_source_xml` on the proof
returned from the same current read; `Patch::apply` then reports
`StaleSource` as before.

The no-op path still passes `source_authorized = false`, so it ignores the
hint and retains the exact-copy behavior. A signed-source refusal still falls
through to the established `main.data()` and managed-snapshot path. Candidate
reconstruction, candidate reparse/readback, changed-document policy, and the
topology/output branches are outside the changed handoff and remain in their
previous order.

## OPC proof boundary

`PartView::source_xml_with_hint` delegates to the existing source XML capture
owner. The common path still performs, in order:

1. read-ahead disable/drain, source freshness, and execution-context checks;
2. encrypted-entry and signature-infrastructure refusals;
3. catalog lookup and XML Part classification;
4. one normal bounded `read_part` call; and
5. post-read source and execution checks.

An eligible hint must also prove both its stored and embedded source snapshot
lineage/version, equivalent Part URI, exact current content type, equal
`ReadLimits`, and original (non-derived) payload ownership. The fresh
`PartData` is checked against the retained original allocation or compared in
bounded chunks. Every chunk checks source/context state, and the full decoded
length consumes `Resource::Work` even when allocation identity avoids the byte
walk. A final source/context fence runs before a hit can return the cloned
proof.

An ineligible hint, a derived splice payload, a length mismatch, or a byte
mismatch never authorizes its bytes. The same `PartData` is passed to
`SourceXmlPart::from_source_parts`, retaining metadata admission, PartBytes
limits, complete bounded XML validation, and final source/context checks. No
second Part read is introduced. On a hit, the returned clone retains the
original proof's source owner and reservations; the just-read `PartData` is
dropped after comparison. On a miss, the fresh proof owns the normal current
payload and reservations.

## Error, security, and budget ordering

The identity gate prevents a foreign patch from changing the error order by
submitting its proof to OPC. Source/version/context checks remain before the
handoff and after the current proof. A current source change or cancellation
therefore remains a typed source/execution error. Current malformed XML still
goes through the full OPC validator before DOCX stale application. Signed and
encrypted checks remain before payload authorization, while the signed no-op
fallback is unchanged. Changed-document policy remains after `Patch::apply`
and before output.

The normal read retains cache, decoded-size, PartBytes, memory/object, input,
source, and cancellation accounting. Eligible hints add a cumulative full-byte
work charge and a bounded comparison; if that charge is refused, the typed
resource error is preserved instead of silently treating the hint as a miss.
The hit does not allocate a second DOCX index or admission token. The miss
retains the fresh snapshot's full admission, and the original patch snapshot
remains alive until the caller drops the commit.

The OPC seam exposes only the existing opaque `SourceXmlPart` proof. It adds no
archive handle, mutable byte access, global cache, provider identity, or
ordinary DOCX CRUD surface, satisfying the ownership and API-layer constraints
in the accepted ADR review.

## Remaining acceptance gates

The source review finds no blocker, but the candidate still needs focused
runtime coverage for a hint hit, foreign lineage and Part identity, source
revision changes before/during the read, cancellation and tight Work limits,
derived hints, malformed/DTD/MCE input, signed/encrypted packages, exact
no-ops, stale output suppression, and budget cleanup. Matched native,
Callgrind, allocation, and hardware lanes must separately verify that the hit
skips the repeated current DOCX scan while preserving output bytes, source
diagnostics, and lifecycle results. A mismatch path should remain visible in
the profile so the full validator is not accidentally removed.
