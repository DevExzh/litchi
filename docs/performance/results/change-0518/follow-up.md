# 0518 follow-up: reusing the topology XML proof

This is a read-only owner review of the remaining XML-validation edge in the
0518 candidate. It uses the frozen source and recorded profiles; no production
files, builds, or captures were changed for this follow-up. OLE2/OOXML remains
the active optimization goal and ODF remains deferred.

## Finding

There is a real, bounded optimization opportunity in
`SourceXmlPart::check_for_publication`. The candidate still enters
`validate_source_xml` once during topology publication. The p512 candidate
profile attributes 4,726,138 Ir to that one call, and the p128 profiles
attribute about 1.19M Ir. The call is the single positive XML-validator edge in
the candidate topology path; the large DOCX snapshot scan has already been
removed from that publication method.

For a `SourceXmlPart` issued by this crate, the XML parse is already a proof:

* `SourceXmlPart::new` validates the retained original bytes with the Part's
  source `ReadLimits` before it returns (`crates/litchi-opc/src/xml_splice.rs:83-127`).
* `XmlSplicePublication::finish` validates the complete assembled payload with
  those same stored limits before installing its immutable `Arc<Vec<u8>>`
  (`crates/litchi-opc/src/xml_splice.rs:510-584`).
* The fields are crate-private, there is no public constructor for arbitrary
  payloads, and the only payload replacement is the checked `finish` path.
  An unmanaged finish still validates; its absence of a reservation is not a
  reason to treat it differently.
* `ReadLimits` derives `PartialEq, Eq` and includes every XML and PartBytes
  ceiling (`crates/litchi-opc/src/limits.rs:113`).

Therefore, once the destination content type is exactly equal to the proof's
stored content type and the destination `ReadLimits` equal the proof's stored
limits, the destination XML parser cannot produce a new structural result for
the same immutable payload. The existing destination `PartBytes` check should
still run because it establishes error ordering and protects the invariant if
the owner changes later.

This is a proof-reuse opportunity, not a security or ADR blocker. The minimal
owner change belongs in `litchi-opc`, at the `SourceXmlPart` publication
boundary. It should remove only the duplicate XML parser pass after all of the
checks below. It should not remove the DOCX candidate reparse/readback or any
package security check.

## Current validation topology

The remaining path is:

| Stage | Current operation | What it proves | Keep? |
| --- | --- | --- | --- |
| Source capture | `SourceXmlPart::new` | Current original bytes are bounded and well formed under source limits. | Yes. |
| Source edit | `XmlSplicePublication::finish` | The complete assembled candidate is bounded and well formed under source limits; source state is fenced before and after. | Yes. |
| DOCX semantic owner | `Snapshot::from_source_xml_with_identity` and the existing candidate semantic readback | DOCX grammar, indexes, changed-document policy, and selected-operation semantics. | Yes. |
| Topology preflight | `SourceXmlPart::check_for_replacement` then `check_for_publication` (`source_backed.rs:6685-6694`, `xml_splice.rs:240-297`) | Destination identity, original bytes, content type, source state, destination limits, and XML validity. | Keep every check; conditionally omit only the duplicate parser. |
| Topology changed loop | Source proof causes `validate_overlay_xml` to be skipped (`source_backed.rs:7674-7695`). | Avoids the authored compactness audit because the source proof owns the complete XML audit. | Yes. |
| Transfer and sink | Source/context fences, monitoring, output limits, ZIP preservation, and final sink checks. | Source freshness and physical publication safety. | Yes. |

The current topology call graph makes the duplication explicit:

1. `finish` validates the assembled payload with `self.source.limits`.
2. DOCX has already parsed the candidate into the target snapshot before the
   commit is published. On the source-reuse path it may use the retained
   `before` layout only after an exact current-source check; that handoff is a
   separate semantic optimization.
3. `write_topology_to_stream` reads the destination Part and calls
   `check_for_replacement`, which calls `check_for_publication`.
4. `check_for_publication` repeats `validate_source_xml` under the destination
   limits. This is the one positive candidate validator shown in
   `profile-after-r2/*-exclusive.txt` (for example the p512 k1 row).
5. The later changed loop deliberately does not parse a source-authorized
   replacement again, and the transfer path only fences its source/context.

## Safe reuse boundary

The parser can be skipped only after the existing checks establish the same
publication policy. The order should remain:

1. Compare the stored source content type with the destination content type.
   The comparison is exact and must remain an error before any success path.
2. Check the stored source-lineage authority, source version, and execution
   context with `check_source_state`.
3. For an existing replacement, keep `destination.ensure_current_public`, the
   destination lineage/version check, equivalent destination Part URI check,
   and the fresh exact comparison between `SourceXmlPart::original` and the
   destination Part bytes. These checks are in `check_for_replacement` and
   prove that the proof belongs to this current destination Part.
4. Check `destination_limits` against the payload length using
   `ReadResource::PartBytes`, even when `destination_limits == self.limits`.
5. If `destination_limits == self.limits`, reuse the XML validity established
   by `new` or `finish`. If they differ, call `validate_source_xml` exactly as
   today under the destination policy.
6. Run the existing final `check_source_state` fence before the proof is
   accepted or transferred.

The equality test is sufficient for the parser's policy because `ReadLimits`
contains the XML event, depth, attribute, relationship, and PartBytes ceilings
used by this validator. Exact content-type equality keeps the XML
classification and source proof's content-type contract intact. The payload is
immutable after the proof is issued, and a clone preserves its source lineage,
version, Part URI, original allocation, payload allocation, and reservations.

The same XML-proof reasoning applies to a `SourceXmlPart` addition when its
stored limits equal the destination limits. `try_add_source_xml_part` is
explicitly allowed to carry a proof from another source-backed package, so the
addition path has no destination original-byte comparison or destination
lineage check. That is not a reason to call the bytes unvalidated: the source
proof remains valid under the equal parser policy. It is a reason to keep the
addition's source freshness/context fence and all destination package guards.
If the owner wants the smallest first patch, it can optimize the common
replacement path and leave additions on the full `check_for_publication` path;
that is conservative, but it is not required for XML correctness when the
stored limits and content type are equal.

## What must not be skipped

Reusing the prior XML proof must not turn a `SourceXmlPart` into an unchecked
byte buffer. The following checks remain authoritative:

* `disable_read_ahead_for_publication`, initial and post-read source/version
  checks, cache admission, one fresh destination read, decoded-size limits,
  and current `PartBytes` checks;
* encrypted-entry refusal and signed-package policy before payload transfer;
* source-lineage/version consistency and cancellation/context checks, including
  the final transfer-boundary fence and `monitor_publication`;
* replacement destination Part identity and exact original bytes;
* content-type equality and all destination topology, aggregate Part, archive,
  relationship, physical-member, and output limits;
* DOCX `Snapshot::from_source_xml_with_identity` when a new candidate snapshot
  is needed, `Patch::apply`, changed-document policy, and the existing
  candidate semantic reparse/readback; and
* ZIP preservation, output-budgeted sink checks, and final source/output
  fences.

The DOCX snapshot reuse already has its own exact-byte and identity gate. A
topology XML-proof reuse must not be used to bypass that gate. Conversely, the
DOCX semantic parse must not be treated as a substitute for OPC's source XML
proof for callers that use `SourceTopologyPlan` directly.

## Budget and error precedence

The current `check_for_publication` does more than parse: it checks destination
PartBytes, charges the XML payload to `Resource::Work`, temporarily admits the
parser's working memory, checks context during parsing, and fences source state
after parsing. The proposed fast branch should preserve the operation's
bounded accounting deliberately.

The recommended behavior is:

* retain the destination PartBytes check before the branch;
* charge one conservative payload-length `Resource::Work` amount for each
  publication attempt, even when the structural parser is reused; and
* omit only the temporary parser memory and event loop when the stored limits
  and content type match. A Work refusal remains the typed resource error and
  is not converted into a fallback or a successful transfer.

Charging the byte amount preserves the existing `topology_source_xml_addition_charges_one_xml_validation_pass`
  contract (`source_backed.rs:12732-12769`) and prevents cloned source proofs
  from making repeated publications free of cumulative Work. The operation is
  still cheaper because it avoids quick-xml event, namespace, and temporary
  allocation work. If a future policy intentionally makes proof reuse free of
  Work, that must be an explicit budget decision with new tests; it should not
  happen as an incidental parser shortcut.

The error order should be preserved as follows:

| Condition | Required result |
| --- | --- |
| Destination/source identity, content type, or original-byte mismatch | Existing typed source/overlay error before any reused-proof success. |
| Source revision or cancellation before, during, or after the destination read | Existing source or execution error. |
| Destination PartBytes or aggregate topology limit | Existing destination limit error before transfer. |
| Equal limits/content type and valid immutable `SourceXmlPart` payload | Charge bounded Work, skip only the duplicate parser, then run the final source fence. |
| Different limits/content type | Existing full destination XML validation and its UTF-8, DTD, namespace, depth, event, attribute, and malformed-input errors. |
| Signed/encrypted/security/output failure | Existing policy or output error, regardless of prior XML proof. |

The private constructors make malformed payloads unavailable through the normal
`SourceXmlPart` API. A source change or cancellation can still occur after a
prior validation, so the pre/post `check_source_state` fences remain necessary.
The proof does not authorize bytes from a different current source, and it
does not authorize a destination package that fails its own signature,
encryption, topology, or output policy.

## Owner API recommendation

No new public API is needed. Keep `SourceXmlPart` opaque and put the decision
inside `litchi-opc`:

* factor the existing publication checks into a private helper that accepts an
  internal “reuse the source-policy XML proof” condition, or make
  `check_for_publication` select that condition automatically when the exact
  content type and `ReadLimits` match;
* preserve the full validation branch for unequal limits; and
* retain `check_for_replacement` as the owner of destination lineage, Part
  identity, and original-byte checks before it enters the helper.

Automatic selection in `check_for_publication` is coherent for additions too,
because the source proof's private issuance invariant covers original and
assembled payloads. A replacement-only helper is an acceptable lower-risk
first step if the owner wants to avoid changing the deliberately cross-package
addition path. In either design, callers must not receive a boolean that lets a
DOCX or topology caller assert “already validated” for an arbitrary byte
buffer; the proof type itself supplies that authority.

No extra `payload_reservation` or pointer test should be used as the validity
marker. Original payloads, derived payloads, and unmanaged derived payloads
all have the same validation invariant, while reservation presence only
describes managed retained memory. If future constructors can issue unvalidated
payloads, add an explicit private validation-state marker before extending the
fast branch.

## ADR and acceptance assessment

There is no ADR blocker in this follow-up:

* immutable source and candidate proofs remain owned by OPC and DOCX;
* equal `ReadLimits` makes the prior XML result applicable to the destination
  parser policy;
* source freshness, cancellation, Work, PartBytes, signature/encryption,
  topology, and output guards remain in place; and
* DOCX candidate reconstruction, final semantic reparse/readback, and
  `Patch::apply` remain unchanged.

The actual design boundary is the distinction between an immutable XML proof
and destination authorization. The former can be reused under equal policy;
the latter still needs every destination check. Foreign source additions make
that distinction visible, but they do not block reusing the XML result itself.

The next implementation owner should add focused coverage for:

1. an existing source-authorized replacement from `finish` with equal limits,
   showing one Work charge and no second XML parser pass;
2. an original (unspliced) replacement and a source XML addition with equal
   limits/content type;
3. each unequal XML limit, content-type mismatch, foreign replacement lineage,
   stale destination bytes, and source-version/cancellation fence;
4. malformed or DTD source capture still failing in `new`, and malformed input
   still taking the full validation path whenever policy differs;
5. signed/encrypted packages and output/topology limit refusals; and
6. DOCX candidate semantic reparse/readback and exact output/readback guards,
   proving that the topology shortcut does not change the candidate handoff.

The recorded profiles prove the opportunity, not a new speedup claim. The
current 54–66% publication reduction and approximately unchanged topology are
valid for the frozen candidate; this follow-up recommends a separately tested
OPC optimization for the remaining validator and does not alter the current
batch's acceptance decision.
