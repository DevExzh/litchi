# 0519 SourceXmlPart proof review

This is an independent, read-only review of the planned
`SourceXmlPart::check_for_publication` fast path on base
`45d71cb6f0cd5d003c544b61c748552a513b0245`. The proposed branch is:

```text
existing content-type/lineage/state/PartBytes checks
    if destination_limits == self.limits:
        consume_work_from_context(source.context_ref(), payload.len())
    else:
        validate_source_xml(..., destination_limits, ...)
existing final source-state fence
```

No production files, builds, or captures were changed for this review. The
0519 priority remains OLE2/OOXML; ODF is deferred.

## Verdict

The planned branch is sound under the current `SourceXmlPart` invariants. A
payload issued by this crate has already passed the complete source XML
validator under the limits stored in the proof. `ReadLimits` has exact
`PartialEq/Eq` semantics and contains every limit consulted by
`validate_source_xml`. Thus an exact destination content-type match plus
`destination_limits == self.limits` is sufficient to reuse the XML result.

The branch must remain inside `SourceXmlPart::check_for_publication`, retain
its current checks and error order, and charge one payload-length Work amount.
It must not be exposed as a caller-supplied “already validated” boolean. The
existing DOCX candidate reparse/readback, replacement authorization, package
security checks, topology checks, transfer fences, and output checks remain
independent obligations. I found no ADR or ownership blocker.

The only implementation caveat is an intentional resource-policy change:
equal-policy reuse no longer reserves the transient XML parser workspace. That
is safe because the branch performs no parser allocation, but it can turn a
previous temporary `Resource::Memory` refusal into success. Tests and the
publication documentation should record this as the expected result.

## Complete construction and mutation inventory

The production source audit found one real constructor, one forwarding
constructor, and one payload mutation. All other uses move or clone an opaque
proof.

| Path | Operation | Validation or mutation consequence |
| --- | --- | --- |
| `crates/litchi-opc/src/xml_splice.rs:55-64` | `SourceXmlPart` fields | `source`, source identity, `original`, `payload`, reservations, and `limits` are private to the crate. `Arc<Vec<u8>>` and `PartData` expose immutable borrows/handles. |
| `xml_splice.rs:83-127` | `SourceXmlPart::new` | Fences source/context, reserves metadata, validates `PartBytes`, runs complete `validate_source_xml` with its `limits`, fences again, and installs `payload = original.shared_bytes()`. |
| `xml_splice.rs:1399-1414` | `from_source_parts` | `pub(crate)` forwarding wrapper; its only production caller passes the current package's `self.limits` after XML classification and one fresh `PartData` read. It cannot bypass `new`. |
| `xml_splice.rs:215-229` | `into_publication` | Moves the proof and permits edits only while `payload` is the original allocation. Derived payloads cannot be edited again. |
| `xml_splice.rs:512-585` | `XmlSplicePublication::finish` | No edits returns the already validated original proof. Edited output is size checked, Work charged, output-memory reserved, assembled, source fenced, fully validated with the stored limits, fenced again, then installed as a new immutable `Arc<Vec<u8>>` with its output reservation. |
| `source_backed.rs:3485-3506` | `PartView::source_xml` and hint wrapper | The OPC owner performs current read/security/classification checks. Hint hits clone only a matching original proof; derived or mismatched hints go through `from_source_parts`. |
| `source_backed.rs:1521-1549` | `try_replace_source_xml_part` | Checks target URI/XML classification, shares `payload` by `Arc`, and stores the whole proof for later destination authorization. It does not mutate bytes. |
| `source_backed.rs:1627-1649` | `try_add_source_xml_part` | Checks destination name/XML classification, retains the source content type, and stores an `Arc<SourceXmlPart>`. Foreign source lineage is intentionally allowed for additions. |
| `source_backed.rs:6610-6618` | Addition preflight | Calls `check_for_publication` with destination content type and limits before any output. |
| `source_backed.rs:6682-6699` | Replacement preflight | Freshly reads destination bytes, then `check_for_replacement` checks destination lineage/version, equivalent Part URI, exact original bytes, and publication policy. |
| `source_backed.rs:7663-7696` | Changed replacement materialization | A source proof skips only the authored compactness audit; it retains the proof token and its payload. |
| `source_backed.rs:7801-7837` | Transfer/output | Rechecks source state, enables source monitoring, shares the immutable payload with the ZIP entry, and retains final source/context/output fences. |

The only assignments to `SourceXmlPart::payload` and
`payload_reservation` in production are the two assignments at the end of
`finish`. The only assignments to the identity fields, `original`, or `limits`
are construction assignments. The internal `source_xml_hint_tests` module
mutates private fields to test fail-closed hint admission, but no non-test
production path does so. `SourceXmlPart` has no public struct-literal path,
mutable byte access, `DerefMut`, or arbitrary-payload constructor.

`PartData` itself wraps an immutable `CachedPayload`; `shared_bytes` clones its
Arc and `shares_allocation_with` only observes identity. The cache may evict a
payload, but a retained `PartData` reservation and Arc keep its allocation
alive. A source proof therefore remains immutable even if the package cache
changes state.

## Limits and proof invariants

`SourceBackedPackage` captures one `ReadLimits` value at construction and
stores it by value. Every production source XML capture passes that package
value through `from_source_parts` to `SourceXmlPart::new`. The proof copies
that exact value into `SourceXmlPart::limits`; there is no setter or production
mutation. `ReadLimits` is `Copy`, `PartialEq`, and `Eq`, and its private fields
cover input/archive, PartBytes, aggregate Part, XML events, XML depth, XML
attribute bytes, and the other package ceilings.

The XML acceptance function is determined by the immutable payload, the
`PackURI` used for diagnostics, and the `ReadLimits`/fixed validator constants.
The part name does not change acceptance. For a replacement,
`check_for_replacement` already requires an equivalent destination Part URI;
for an addition, `try_add_source_xml_part` already checks that the destination
name is XML-classified for the retained content type. The destination content
type comparison in `check_for_publication` remains exact.

The issuance invariant is complete for both payload classes:

* an original payload passed `new`;
* an edited payload passed `finish`; and
* a no-op `finish` returns the original proof rather than manufacturing a new
  unvalidated value.

This makes `payload_reservation` unsuitable as a validity marker. Managed
original proofs have no derived payload reservation, managed edited proofs do,
and unmanaged edited proofs can also have no reservation. All three have the
same XML-proof invariant. A future constructor that can issue arbitrary bytes
would require a private validation-state marker before this optimization could
remain sound.

Source identity is separate from limits. `SourceLineage` uses process-local
Arc identity, while `SourceVersion` records the observed source revision. The
proof stores both, and `check_source_state` verifies the source snapshot and
context. Replacement publication additionally compares the proof lineage and
version with the destination snapshot and checks the fresh destination
original bytes. Equal version values from different packages therefore cannot
authorize a replacement. Additions deliberately permit foreign proofs, so
their source and destination checks must remain separate.

## Safe branch boundary

The branch should be placed only after the current checks, in this order:

1. For a replacement, `check_for_replacement` must first run destination
   freshness, destination lineage/version, equivalent Part URI, and exact
   original-byte checks. An addition has no destination original-byte check
   because its foreign-proof transfer is intentional.
2. `check_for_publication` must retain exact content-type equality, internal
   proof-lineage consistency, and `check_source_state`.
3. The destination `PartBytes` check must run even when the two `ReadLimits`
   values compare equal.
4. Equal limits select `consume_work_from_context(source.context_ref(),
   payload.len())`; unequal limits select the existing
   `validate_source_xml` call with destination limits.
5. The final `check_source_state` fence must remain after either branch.

This boundary reuses only XML validity. It does not reuse destination
authorization, source freshness, or any semantic DOCX state. It is safe for an
original proof and for a `finish` proof because both payload classes have the
same issuance invariant. It is also structurally safe for a foreign addition:
the validator's part-name argument affects diagnostics, not acceptance, while
the source and destination security checks remain outside the parser result.

## Remaining validation topology

The source and candidate checks are ordered as follows:

1. `new` validates the original source XML.
2. `finish` validates the complete assembled candidate under the source
   policy, if a splice was performed.
3. DOCX builds or reuses its semantic snapshot only after the existing current
   source and identity gates. Candidate `Snapshot::from_source_xml...`,
   `Patch::apply`, changed-document policy, and semantic readback are outside
   this OPC helper and must remain unchanged.
4. Topology publication calls `check_for_replacement` for existing Parts or
   `check_for_publication` for additions. The latter currently repeats the
   complete XML validator under destination limits.
5. The changed loop skips `validate_overlay_xml` for a source-authorized
   replacement because its proof already owns complete XML validation.
6. Transfer checks source state and monitoring, then the writer performs the
   existing physical and output checks.

The planned branch removes only step 4's duplicate parser pass when the
destination policy exactly equals the stored proof policy. It does not remove
step 2, step 3, or any final source/output step.

The 0518 candidate profile confirms the target. The remaining candidate XML
validator is one `SourceXmlPart::check_for_publication` call: approximately
1.19M Ir for p128 and 4.73M Ir for p512. No new profile is claimed here.

## Security, freshness, and error review

The following guards remain effective with the branch:

* `write_topology_to_stream` disables read-ahead, checks the destination source
  and execution context, and refuses encrypted entries before addition or
  replacement proof transfer;
* signed-source policy remains in its existing topology position, and source
  XML capture already refuses signed/encrypted source packages;
* `check_for_replacement` still checks destination freshness, lineage, version,
  equivalent Part URI, and exact original bytes before publication policy;
* `check_for_publication` still checks exact content type, internal lineage
  authority, source version/context, destination `PartBytes`, and the final
  source/context fence;
* topology aggregate Part/archive/relationship/physical-member limits and
  output reservations remain unchanged; and
* source-token transfer still rechecks freshness/cancellation, enables
  publication monitoring, and keeps the existing sink/output fences.

The planned `consume_work_from_context` call is cancellation-aware because
`ExecutionContext::consume` checks cancellation before charging. The existing
pre-branch and final `check_source_state` calls still cover cancellation and
source revision. The equal-policy path is constant-time with respect to the
payload and therefore does not need the parser's per-event cancellation
checks; it performs no payload I/O or parser allocation.

The current error order is preserved through the branch. For replacements,
the destination freshness/identity/original-byte checks occur before the
publication checks below; for additions, that destination row is not
applicable:

| Order | Existing check or result | 0519 requirement |
| --- | --- | --- |
| Replacement preflight | Destination freshness, lineage/version, equivalent Part URI, and exact original-byte comparison | Return the existing typed source/overlay error before publication-policy reuse. |
| Publication 1 | Content-type mismatch or inconsistent proof lineage | Return the existing typed overlay error before Work or XML reuse. |
| Publication 2 | Source revision or cancellation before Work | Return the existing typed source/execution error before transfer. |
| Publication 3 | Destination `PartBytes` check | Retain it even though equal limits make a valid proof's length redundant. |
| Publication 4a | Equal limits | Charge payload-length Work; omit only transient parser work/memory. A Work refusal remains a typed resource error. |
| Publication 4b | Unequal limits | Run the existing destination validator and retain its UTF-8, DTD, namespace, depth, event, attribute, and malformed-input errors. |
| Publication 5 | Final source/context fence | Retain it before publication continues. |

For a valid proof, skipping the second UTF-8/parser pass cannot hide malformed
bytes: the original or candidate payload could only have entered the type after
the complete validator succeeded. In-place source mutation without a version
change is outside the existing `ReadAt`/`SourceVersion` contract and is not
made worse by this branch; the destination original-byte check still protects
replacement publication within that contract.

One subtle but intentional change is memory error behavior. `validate_source_xml`
temporarily reserves parser workspace after its Work charge. The equal-policy
branch does not need that workspace, so a low-memory context may succeed where
the old duplicate parser admission failed. Retained `PartData`, metadata,
edited-payload, topology, and output reservations still apply. The branch must
not drop or detach any retained reservation to obtain this result.

For additions, the proof's source context is the context used by the current
validator and by the planned Work charge. The destination package still runs
its own context/security/topology checks. A foreign source proof is not treated
as a destination replacement because `check_for_replacement` is not called for
additions; equal limits authorize reuse of the XML result only, not a foreign
destination identity.

## Recommended API and documentation boundary

The planned implementation is the smallest coherent owner change. No public
API or caller assertion is needed. Keep the branch inside
`SourceXmlPart::check_for_publication`, and make the method decide from its
stored proof and the destination arguments. `check_for_replacement` remains
the owner of destination identity/original-byte authorization.

The public `try_add_source_xml_part` documentation currently says the
publisher “rechecks ... XML well-formedness.” After this change that should be
worded as “rechecks or reuses the immutable XML proof under destination
limits,” so callers do not mistake proof reuse for unchecked transfer. The
`SourceXmlPart` issuance documentation should continue to state that every
returned payload is fully validated.

Charging `payload.len()` through `consume_work_from_context` is the right
budget boundary. It preserves the existing
`topology_source_xml_addition_charges_one_xml_validation_pass` expectation,
keeps repeated cloned proofs bounded by cumulative Work, and still removes the
quick-xml event/namespace/parser-memory cost. It also preserves cancellation
checking through `ExecutionContext::consume`.

## Required tests

The next implementation should add or update focused coverage for:

1. An original source XML replacement with equal limits and exact content type:
   output reopens, preserves bytes, and no second parser allocation is
   required.
2. A derived `finish` replacement with equal limits: the assembled candidate
   publishes unchanged, retains its payload reservation, and still bypasses
   only the authored compactness audit.
3. A source XML addition with equal limits/content type, including a foreign
   source lineage, proving source freshness and destination security checks
   still run.
4. Every unequal XML policy exercised by the existing depth/event tests (and a
   different PartBytes or attribute limit), proving the full destination
   validator remains selected and returns its typed limit error.
5. Content-type mismatch, foreign replacement lineage, equivalent-Part mismatch,
   and stale destination original bytes, proving identity/error precedence is
   unchanged and no output is emitted.
6. A managed Work budget with exactly one publication payload-length charge and
   a budget one unit short, proving the typed Work refusal occurs before output.
7. Cancellation/source-version changes before the proof check, after the
   initial check, and at the transfer fence. The existing addition tests arm
   these boundaries and should retain their expected check count.
8. A managed memory profile where retained proof/output reservations fit but the
   duplicate parser workspace would not; equal-policy publication should use no
   transient parser reservation and still release all retained resources.
9. Signed/encrypted source and destination packages, topology/output-limit
   refusals, and exact no-op behavior.
10. DOCX candidate construction, semantic reparse/readback, `Patch::apply`,
    exact output, and source-version guards, proving this low-level shortcut
    does not replace the format-owner checks.

The existing `destination_depth_and_event_limits_refuse_source_xml_additions`
and `topology_source_xml_addition_charges_one_xml_validation_pass` tests are
particularly important: the former exercises the unequal-policy fallback, and
the latter fixes the intended cumulative Work accounting for the equal-policy
branch.

## Disposition

Approve the planned 0519 OPC change for implementation and focused testing.
It reuses an immutable proof under an equal policy, preserves all authorization
and semantic fences, and removes the one measured duplicate validator from the
current topology path. The memory-admission relaxation is real and should be
documented/tested as a consequence of doing no parser allocation. No ADR
amendment or cross-format change is required.

## Live diff approval and coverage

The applied working-tree diff was reviewed at base HEAD
`45d71cb6f0cd5d003c544b61c748552a513b0245`. The tracked production changes are
exactly the planned branch in `xml_splice.rs` and the corresponding public
`try_add_source_xml_part` documentation in `source_backed.rs`; the new test
module is declared under `#[cfg(test)]` at the end of `xml_splice.rs`. No
production Rust change beyond that branch was found.

Hashes of the reviewed files at this point are:

| File | SHA-256 |
| --- | --- |
| `crates/litchi-opc/src/xml_splice.rs` | `9d629e4a56ed406a8c3db2346495fd4b787e79891a19b5c0ece576408a0e5b30` |
| `crates/litchi-opc/src/source_backed.rs` | `6f37f9a1e2ecd0435cc811b56a3e92d4bbdb0b41bcdf755478d04f2455f8ea48` |
| `crates/litchi-opc/src/publication_proof_tests.rs` | `952747c865ad19c03086538c2e83100207fdfed682780f9367ab0404e2c89dad` |

The production diff applies the branch after the existing publication checks:

```text
content type -> proof lineage -> source/context state -> destination PartBytes
    equal ReadLimits: consume source-context Work(payload.len)
    different ReadLimits: validate_source_xml under destination limits
final source/context state fence
```

The public addition documentation now accurately states that XML validity is
reused under identical limits and fully checked under different limits. The
replacement path and its destination identity/original-byte checks are
unchanged.

The new internal `publication_proof_tests.rs` contains eight tests:

| Test | Coverage |
| --- | --- |
| `equal_original_and_derived_proofs_cover_replacement_and_addition` | Exercises equal-policy original replacement, derived `finish` replacement, and addition; each retains exactly one payload-sized Work charge. |
| `equal_limit_proof_hit_needs_no_parser_memory_but_keeps_work_bounded` | Saturates remaining managed memory and proves equal-policy reuse needs no transient parser workspace while still charging Work. |
| `different_limits_still_require_the_validator_memory_reservation` | Changes a non-XML `max_parts` field, proving full `ReadLimits` equality is required and the unequal path still admits parser memory. |
| `lower_xml_limits_take_the_full_validator_path` | Uses a lower XML depth limit and retains the typed `ReadLimit` fallback plus Work charge. |
| `proof_work_refusal_is_typed_and_does_not_partially_charge` | Confirms equal-policy Work exhaustion returns the typed resource error without a partial cumulative charge. |
| `content_type_and_lineage_guards_precede_proof_reuse` | Confirms content-type and internally inconsistent lineage failures occur before Work reuse. |
| `source_revision_invalidates_a_previously_valid_publication_proof` | Confirms source version changes invalidate an otherwise valid proof. |
| `destination_identity_is_checked_for_replacements` | Confirms equal-version foreign package lineage cannot authorize a replacement. |

Existing public and source-backed tests continue to cover actual topology
publication, foreign additions, source cancellation/version fences, signed and
encrypted policy, destination XML depth/event limits, exact output/readback,
reservation cleanup, and cumulative addition Work. The focused source review
did not execute those tests or the new module; root's running compile/test
receipts remain authoritative for execution status.

**Approval at the reviewed hash:** the implementation has no correctness,
security, provenance, or ownership blocker. The only expected behavior change
is lower transient parser-memory demand on an equal-policy proof hit. Retain
the full fallback for unequal limits and retain every semantic, source,
destination, security, topology, transfer, and output fence.
