# OPC source XML hint review

This is an independent, read-only review of the proposed 0518 optimization. It
examines the OPC owner and the DOCX handoff at revision
`afb62ab7a70859dad4a4b9c8eea91402d1ef4052`. No production source, build, or
capture was changed for this review. The adjacent
[`snapshot-reuse-design.md`](snapshot-reuse-design.md) describes the DOCX
transaction side; this note records the lower-level proof and budget contract
that that design depends on.

## Finding

The optimization is safe and has no ADR blocker. `litchi-opc` can accept an
existing `SourceXmlPart` as an optimization hint while still reading and
authorizing the current main Part. The hint may suppress the repeated generic
source-XML audit only after the package has performed its normal source,
security, read-ahead, Part classification, limits, cache, ZIP, and freshness
checks and the fresh decoded bytes equal the hint's retained **original**
bytes under the same source identity and `ReadLimits`.

The highest-impact result requires the DOCX owner to use that proof to reuse
`Patch::source()` (the `before` snapshot). Merely calling the existing
`PartView::source_xml()` through a new wrapper would avoid no DOCX scan and
would still repeat the OPC XML validator. Once the current proof is exact, the
DOCX publisher can use the immutable `before` snapshot as its current snapshot;
otherwise it constructs the current snapshot from the already-read fresh proof
and leaves `Patch::apply` as the stale-source authority.

The recommended low-level seam is one narrow method on the public OPC view:

```text
PartView::source_xml_with_hint(&self, hint: &SourceXmlPart)
    -> Result<SourceXmlPart>
```

`SourceBackedPackage::source_xml_part` should own the implementation, with the
existing `source_xml()` path delegating to the same common helper without a
hint. The proof type stays opaque and constructor-free. There is no global
cache, archive type escape, or mutable byte handle in this proposal. For the
minimal contract, only a hint whose `payload` is still the original
allocation may take the hit; a derived splice is a miss and follows the full
validator path.

## Opportunity and current path

The scoped 0517 Callgrind profile already identifies the duplicate work. The
publication method's inclusive instruction count and its current snapshot
subtree were:

| workload | publication method | `main_document_snapshot` | share |
| --- | ---: | ---: | ---: |
| p128, K1, owned, batch | 5,467,082 | 2,953,819 | 54.03% |
| p512, K1, owned, batch | 17,749,121 | 11,657,641 | 65.68% |
| p512, K32, owned, batch | 17,749,977 | 11,657,814 | 65.68% |

These are prior scoped profile observations, not a candidate speedup claim.
The authoritative profile analysis for the selected p128 case reports
5,468,004 inclusive instructions for
`litchi_docx::source_backed::Package::publish_document_commit_to_stream`,
with 2,954,680 in `main_document_snapshot` (54.03%). See
`profile-preflight/profile-analysis.json` in this change directory.

For a managed changed document, the current call chain is:

1. `Package::publish_document_commit_to_stream` asks
   `main_document_snapshot` to reconstruct the current source.
2. `PartView::source_xml` performs the OPC source capture and
   `SourceXmlPart::new` performs the bounded source XML audit.
3. DOCX runs `ensure_source_document_xml` and
   `Snapshot::from_source_xml`, which scans and indexes paragraphs, tables,
   block controls, and conformance.
4. `Patch::apply` compares that newly built snapshot with the patch's
   immutable `before` snapshot.

The edit that produced the commit already did steps 2 and 3 for the same
source bytes and retained the result in `Patch::source()`. A successful hint
hit can therefore remove the second OPC audit and the second DOCX semantic
scan while retaining one fresh physical Part read and all publication fences.

## What `SourceXmlPart` actually proves

`SourceXmlPart` is not just a byte buffer. Its fields at
`crates/litchi-opc/src/xml_splice.rs:48-66` are the ownership and proof
boundary:

| field | role in a hint decision |
| --- | --- |
| `source: SourceSnapshot` | The positional `ReadAt`, captured source version and length, execution context, read monitor, and budget diagnostic owners. |
| `source_lineage` | Process-local package lineage; equality is source-owner identity, not a caller version token. |
| `source_version` | Version observed when the proof was issued. |
| `source_partname` | The source Part URI, compared with the destination using the existing URI equivalence rule. |
| `source_content_type` | The validated OPC content type of that Part. |
| `original: PartData` | The exact decoded source allocation retained by the proof. |
| `payload: Arc<Vec<u8>>` | The bytes currently published by the proof; these differ from `original` after a splice. |
| `payload_reservation` | Optional managed reservation for a derived assembled payload. |
| `_metadata_reservation` | Shared managed reservation for the proof's retained metadata. |
| `limits: ReadLimits` | The XML/Part policy under which the original proof was admitted. |

`bytes()` returns `payload`, so it is deliberately the wrong comparison for a
generic hint. `XmlSplicePublication::finish` keeps `original` unchanged,
allocates a new `payload`, reserves that output, and validates the assembled
bytes. `is_original_payload` currently identifies the original state by
`Arc::ptr_eq(payload, original.shared_bytes())`; see
`crates/litchi-opc/src/xml_splice.rs:228-238` and `:510-581`.

The minimal contract should require `is_original_payload()` for a hit. The
DOCX publisher passes the `before` proof, which is already an original source
proof. A derived splice remains useful as a caller-owned candidate, but it
must take the ordinary fallback path; its `payload` must never authorize a
current source read. A later owner-reviewed extension could normalize a
derived hint to its `original` allocation, but that would need separate
reservation and regression coverage and is not needed for the DOCX hot path.

The type has no public constructor and no public mutable byte access. A
managed `PartData` cannot escape as a bare `Arc`; an unmanaged allocation is
still immutable while owned by the proof. The exact byte comparison remains
necessary as the patch precondition even though source adapters are required
to advance `SourceVersion` when observable bytes change.

## Package lineage and identity

`SourceBackedPackage::from_read_at_inner` creates one
`SourceLineage(Arc::new(()))` for each opened package. Every clone of that
package's internal `SourceSnapshot` shares the token, while two packages with
equal source bytes and equal caller-supplied `SourceVersion` values receive
different tokens. `SourceLineage` equality is pointer equality at
`crates/litchi-opc/src/source_backed.rs:2704-2718`; version equality alone is
therefore insufficient.

A hint hit must require all of the following, in addition to a current source
and execution-context check:

1. The hint's internal snapshot lineage agrees with its stored
   `source_lineage`, and that lineage is the current package lineage.
2. The hint's stored `source_version` equals the current package version and
   the hint snapshot reports the same version.
3. `hint.source_partname` is equivalent to the resolved current Part URI.
4. The hint's content type is exactly the current Part content type.
5. `hint.limits == self.limits`. `ReadLimits` derives `PartialEq` and `Eq`
   (`crates/litchi-opc/src/limits.rs:113`), so this check covers the complete
   source policy without relying on a partial field list.
6. The fresh `PartData` bytes equal `hint.original.as_bytes()` exactly.

The current package's catalog is immutable after opening, but checking Part
URI, content type, and limits explicitly documents why the proof cannot be
transferred across a future package or policy boundary. A mismatch is an
ordinary optimization miss, not a hint error.

## Required OPC revalidation boundary

The existing `source_xml_part` sequence at
`crates/litchi-opc/src/source_backed.rs:5881-5908` should remain authoritative
through the fresh Part read:

1. Disable and drain read-ahead for exact publication.
2. Check source freshness and the execution context.
3. Refuse encrypted entries.
4. Refuse signature infrastructure unless the caller takes the established
   explicit signed-source policy path.
5. Resolve the catalog Part and require the same XML Part classification.
6. Read the Part once through `read_part`.
7. Check source freshness and the execution context again.

`read_part` is not a cheap unchecked byte lookup. Its cache path performs the
source and context fences; a managed load admits declared Part bytes, memory
and object reservations, and cold `Resource::Work`; the ZIP reader checks
decoded length and physical metadata; the returned `PartData` retains the
appropriate payload owner. A cache hit still checks the source and context.
The hint helper must call this path and must not reuse `hint.original` before
the current Part read.

After step 7, the hint helper should perform the identity, original-payload,
and original-byte checks. A hit can return a clone of the original hint. A
miss (including a derived hint) must pass the same fresh `PartData` to
`SourceXmlPart::from_source_parts`, not call `read_part` again. That constructor retains the current behavior at
`crates/litchi-opc/src/xml_splice.rs:83-127`: metadata admission,
`ContentType` validation, `PartBytes` checking, complete bounded XML
validation, and final source/context fences.

The match branch may reuse the already validated content-type token from the
immutable catalog after exact equality. It must retain the hint's metadata
reservation and avoid creating a second metadata allocation. Since derived
hints are misses in this minimal contract, their payload reservation remains
with the caller while the fresh `PartData` is passed to the ordinary
constructor. For an original hint, the fresh `PartData` handle can be dropped
after comparison; the retained original proof remains the ownership of the
returned value.

This boundary is deliberately narrower than a snapshot cache:

* It skips only generic validation of bytes proven identical to an existing
  source proof.
* It never skips the fresh source read, source version checks, ZIP/decoded-size
  checks, cache admission, PartBytes check, signature/encryption policy, or
  read-ahead transition.
* It never skips validation of a derived candidate. `check_for_publication`
  and `check_for_replacement` must continue to validate `payload`, compare the
  destination's original bytes, and enforce destination limits; see
  `crates/litchi-opc/src/xml_splice.rs:240-302`.

## Budget and cancellation accounting

The hit changes how much repeated work is performed, but it cannot turn an
unadmitted operation into a successful one.

* The fresh `read_part` retains its existing `InputBytes`/`PartBytes`, memory,
  object, cold-work, cancellation, and source-freshness behavior. The
  `PartData` produced for the comparison is live until the comparison finishes.
* The helper must recheck `PartBytes` against the current `ReadLimits` before a
  hit. This keeps the source proof's explicit limit boundary visible even if a
  future `read_part` cache path changes.
* If the fresh allocation is pointer-identical to `hint.original`, an
  allocation-free `PartData::shares_allocation_with` fast path is valid under
  the immutable cache contract. The helper should still admit a conservative
  `Resource::Work` charge for the full original length before deciding the
  pointer fast path. That preserves the old source-capture budget boundary
  even though the pointer check avoids CPU comparison. If the allocations are
  distinct, compare the bytes in cancellation-aware bounded chunks, charging
  the bytes examined (or the same conservative full-length bound) to
  `Resource::Work` and checking source/context state around and periodically
  during the comparison. A plain uninterruptible `slice == slice` over a
  large Part would create a new cancellation and budget hole.
* The existing source XML constructor charges the parser workspace and XML
  validation work. A hit intentionally removes that duplicate parser work;
  work performed by the fresh read and the exact comparison remains admitted.
  A byte-mismatch miss necessarily pays the comparison work and then the one
  full validator pass on the same fresh bytes; it must not read the Part a
  second time or allocate a second fresh payload.
* Cloning an original hint shares `_metadata_reservation` and its original
  `PartData`; no new output allocation or global retained cache entry is
  introduced. A derived hint is not cloned on a hit, so its
  `payload_reservation` remains owned by the caller and the fallback's fresh
  proof receives its own normal metadata reservation.
* Any cancellation, resource limit, allocation failure, or source change
  observed during comparison is returned as its typed error. It must not be
  converted into a mismatch and silently sent through a less authoritative
  path.

## DOCX handoff and snapshot reuse

`Snapshot::XmlStorage::Source` stores an `Arc<SourceXmlPart>` and a separate
`SourceIdentity` containing lineage, version, and main-Part URI
(`crates/litchi-docx/src/document/transaction.rs:795-834`).
`Snapshot::from_source_xml` has already performed the bounded DOCX scan and
retains the resulting indexes (`:973-1010`). `Snapshot::same_source` compares
both bytes and identity, and `Patch::apply` uses that exact check
(`:4643-4657`).

The DOCX owner should add only crate-private helpers needed for this handoff,
for example a source-hint accessor and a method that compares a returned
`SourceXmlPart` with the snapshot's bytes and the current package identity.
The helpers should not expose `SourceIdentity` fields or OPC internals to the
public API.

For a changed, managed, source-authorized commit, the intended order is:

1. Keep the current `source_version`, execution-context, main-Part lookup, and
   DOCX content-type checks before looking at the hint.
2. Use the patch source's source identity as a cheap gate. If lineage, version,
   or Part identity differs, use the existing current-snapshot path so foreign
   patches retain their existing stale/error behavior.
3. Pass `commit.patch().source().source_xml()` to
   `main.source_xml_with_hint`. The signed-source refusal must continue to
   select the existing `main.data()`/managed fallback, and encrypted or other
   errors must retain their current mapping.
4. Compare the returned proof with the `before` snapshot using exact bytes and
   current lineage/version/Part identity. On a hit, use
   `commit.patch().source().clone()` as the current snapshot and run the
   unchanged `Patch::apply`; it returns the already-built target for a changed
   patch.
5. On a miss, run `ensure_source_document_xml` and
   `Snapshot::from_source_xml` once on the returned fresh proof, then call the
   unchanged `Patch::apply`. A byte or identity mismatch remains
   `TransactionError::StaleSource` after the same current-source checks as
   before.
6. Leave candidate reconstruction, changed-document policy, source splice
   validation, topology planning, source monitoring, and sink/output fences in
   their current positions.

The no-op publication path remains unhinted. Its exact-copy behavior is needed
for the established signed/no-op policy and does not pay for a changed
transaction's source proof. A source-authorized path that fell back to a
managed `PartData` because the package is signed has no reusable
`SourceXmlPart`; it stays on the existing branch.

## Error precedence

The helper should preserve the following precedence rather than treating a
hint as an alternate source:

| point of failure | required result |
| --- | --- |
| Read-ahead drain, initial source/version, or context check | Existing typed source, cancellation, or execution error. |
| Encrypted entries or signature infrastructure | Existing overlay refusal or `SignedSourceRequiresExplicitPolicy`, before payload read. |
| Part lookup or XML classification | Existing `PartNotFound`/overlay error. |
| Fresh cache admission, ZIP read, declared/decoded size, PartBytes, source change, or cancellation | Existing `OpcError` from the authoritative current read. |
| Post-read source/context fence | Source or execution error wins before hint inspection. |
| Hint lineage/version/Part/content-type/limits/byte mismatch | No error; construct a fully validated proof from the same fresh `PartData`. |
| Comparison work/cancellation/source fence failure | Typed resource, cancellation, or source error; do not downgrade to a miss. |
| Full fallback XML validation | Existing malformed XML, UTF-8, namespace, DTD, depth, attribute, and limit errors. |
| Final source/context fence | Existing source or execution error before the proof escapes. |

If the current bytes equal a valid hint's `original`, the current bytes cannot
simultaneously be malformed under the same limits. If they differ, the full
validator remains the authority, including its old malformed-input error
precedence before DOCX stale-patch application. A source that mutates without
advancing its required `SourceVersion` is outside the `ReadAt` contract and
has the same cache limitation as the existing path; the exact comparison still
protects against distinct allocations and stale patch identities within the
contract.

## ADR assessment

No amendment is required:

* ADR 0003 requires immutable cheap-to-share snapshots and exact source
  authorization. Reusing an already scanned `before` snapshot after a fresh
  exact source proof satisfies both; it does not mutate or weaken a patch.
* ADR 0005 requires positional source freshness, lazy payload ownership, and
  hierarchical budgets. The helper preserves the physical read and all
  resource fences, and charges comparison work while removing duplicate
  parser/index work.
* ADR 0006 requires non-mutating validation and explicit signed/encrypted
  policy. Security classification remains before the hint and derived output
  validation remains unchanged.
* ADR 0010 and ADR 0011 put OPC physical-package ownership below the format
  facade. The hint method belongs in `litchi-opc`; DOCX receives only the
  opaque `SourceXmlPart` and its own snapshot helper.

The only real blocker would be a policy decision that every source capture
must repeat XML validation even when the same immutable source proof and
limits are retained. No accepted ADR states that requirement. If maintainers
choose that policy later, the hint can still save the DOCX scan only through a
separate explicitly approved snapshot-reuse API, but that would reject the
highest-impact two-layer optimization described here.

## Focused verification before measurement

The implementation should add focused tests before a new performance claim:

1. Same-package original hint, warm cache, and an evicted/cold cache both
   perform the fresh read and return exact original bytes.
2. A derived hint whose `payload` differs from `original` never takes the hit;
   it uses one fresh full-validation fallback, retains the candidate hint's
   output reservation, and leaves the candidate hint unchanged.
3. Equal bytes and equal `SourceVersion` from a foreign package, a different
   Part URI, a different content type, or different `ReadLimits` never hit and
   use one fresh full-validation fallback.
4. A fresh byte mismatch, malformed fresh XML, source revision during the
   read/comparison, cancellation, and exhausted comparison work preserve the
   existing typed error and empty-output behavior.
5. Read-ahead is drained before the hinted read; encrypted and signed package
   policy remains unchanged; candidate `check_for_replacement` still validates
   assembled payloads.
6. A managed DOCX changed commit with unchanged current source reuses
   `Patch::source()` and produces the same bytes and semantic result. A changed
   source builds the fresh snapshot and remains stale through `Patch::apply`.
   Foreign, signed, malformed, protection, MCE, DTD, cancellation, and sink
   failure cases remain covered.

Only after these gates should the matched and mismatched native campaigns and
the scoped Callgrind profiles be rerun. The old 54–66% subtree share motivates
the work but is not evidence for a new candidate speedup.

## Read-only review of the applied OPC seam

The current unbuilt production diff adds the public
`PartView::source_xml_with_hint` wrapper and factors
`SourceBackedPackage::source_xml_part_with_hint` behind the existing
`source_xml_part` path (`crates/litchi-opc/src/source_backed.rs:3487-3503` and
`:5899-5967`). It matches the minimal contract above:

* read-ahead is disabled before the source/context fences, security checks,
  Part lookup, XML classification, and one normal `read_part`;
* the hint is eligible only for matching lineage, version, Part URI, content
  type, limits, and `Arc::ptr_eq(payload, original)`;
* the fresh source/context fences are retained before and after the hint
  decision, and a hit returns the original hint clone without a second
  metadata/parser allocation;
* the complete original-source validator receives the same fresh `PartData`
  on every miss, including derived hints, so there is no second read; and
* the eligible path charges `data.len()` to `Resource::Work` before accepting
  either pointer or byte equality. A work-limit failure is returned as the
  typed execution/resource error and is not converted into a fallback;
* pointer equality uses the existing `PartData` allocation-identity helper,
  while distinct allocations are compared in 64 KiB chunks with source and
  context checks before each chunk and a final fence.

This is approved as the minimal owner API and does not require an ADR change.
The latest diff closes the earlier interruptibility concern with bounded
comparison chunks. Its intentional budget/error boundary should remain
documented: the full comparison Work charge is admitted after the fresh read
and before byte equality, so an exhausted Work budget returns the typed
resource error without attempting fallback validation. This avoids a second
read or a second failed Work admission. It can report Work exhaustion before
the fallback validator's UTF-8/XML error in the narrow case where an eligible
hint's fresh bytes differ and the budget is already exhausted; this is a
conservative, typed failure of the newly required equality operation. If exact
legacy malformed-input precedence is required even under an exhausted Work
budget, that policy would need an explicit preflight tradeoff; it is not an
ADR ownership blocker.

## Read-only review of the applied DOCX handoff

The current DOCX diff adds `main_document_snapshot_with_before` and passes
`commit.patch().source()` only from changed publication. It performs the
source-identity gate before calling the OPC hint API, preserves the signed
fallback, and reuses the retained `before` snapshot only after the returned
proof matches the current lineage, version, Part URI, and exact bytes
(`crates/litchi-docx/src/source_backed.rs:1298-1303` and `:1745-1823`). Other
snapshot callers still use the unchanged wrapper with no hint.

`Snapshot::source_identity_matches` and
`Snapshot::reuse_if_source_xml_matches` remain crate-private and do not expose
`SourceIdentity` or OPC fields. The source-backed snapshot clone retains its
existing managed admission and parsed indexes. Candidate reconstruction,
`Patch::apply`, changed-document policy, topology replacement, signed/no-op
handling, and final package source/context fences remain in their existing
positions. This handoff is approved and has no ADR or ownership blocker.

There is one bounded efficiency caveat for the mismatch path. When the OPC
hint misses and returns a freshly allocated proof, `reuse_if_source_xml_matches`
performs a byte comparison and then `Patch::apply` performs its established
`Snapshot::same_source` comparison again before returning `StaleSource`. There
is no second Part read or payload allocation, and the normal successful hint
hit compares the same retained allocation by pointer, so this does not reduce
the expected hit-path win. If stale-source workloads become material, the
owner could return an internal hit marker or let `Patch::apply` be the sole
byte comparison on the miss path. The current duplicate comparison is a
performance follow-up rather than a correctness or ADR blocker.
