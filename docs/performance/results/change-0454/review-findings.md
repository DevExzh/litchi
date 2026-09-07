# 0454 review findings

Read-only review of the pending `source_backed.rs` and PPTX cross-copy changes
against `docs/GOAL.md` and the accepted ADRs. No production source was edited by
the reviewer.

## Frozen-source re-review (current status)

The following status supersedes the draft findings below. It is based on the
current `xml_splice.rs`, `source_backed.rs`,
`source_cross_copy.rs`, and the recorded integrated outcome. This reviewer did
not edit production source or run a build in this pass.

### Remaining blockers

No verified production blocker remains in the latest frozen source reviewed
here. The final native inventory and the corrected bytes/range pilot receipts
pass; any coordinator reruns are final-gate bookkeeping rather than an open
review finding. The pre-fix compactness receipt is historical evidence.

### Historical blocker notes (superseded)

1. **Closed: relationship declaration grammar is validated.** The current
   `Event::Decl` branch at `source_backed.rs:1270-1278` checks placement and
   delegates the payload to `validate_source_declaration`, so malformed version,
   encoding, standalone, ordering, and duplicate-field values are refused. The
   focused mutations at `source_backed.rs:12860-12875` retain the regression
   evidence.

2. **Closed: canonical relationship output has a pre-allocation reservation.**
   The current serializer reservation is described by
   `canonical_relationship_serializer_memory_bound` at
   `source_backed.rs:1381-1408`. It is computed from the exact canonical XML
   length, a transient output allowance, sorted relationship references, the
   reference Vec, and fixed serializer state before `try_to_xml_bytes()`.
   Existing-member source comparison reserves it at `6851-6857`; changed
   existing canonical output reserves and retains it at `6866-6905`; new-member
   canonical output does the same at `7077-7113`. Existing source materialization
   remains reserved at `6818-6827`, while noncanonical append staging retains
   separate source and append reservations at `6981-7074`.

3. **Closed: the final native inventory passes the source-preserving cases.**
   `final-native-inventory-r2.json` records a pass with 189 image-slide rows,
   and `final-outcome-comparison.json` records 185 unchanged rows plus four
   published rows: the LibreOffice smoke test slide 0 and POI slides 16, 17,
   and 33. The four outputs replace the old missing-name refusal with
   `PUBLISHED` results, including the smoke-test output at 42,948 bytes; the
   prior `NotCompact(FormattingWhitespace, offset 55)` receipt is historical.

<!--
The original pre-freeze wording is retained below as evidence. Its first two
items are superseded by the current declaration and serializer fixes.
-->
<!--
1. **The relationship append parser placement-checks declarations but does not
   validate their grammar.** At `source_backed.rs:1260-1268`, `Event::Decl(_)`
   only checks position and then accepts any version, encoding, or standalone
   value. A relationship member beginning with
   `<?xml version="1.1" encoding="UTF-16"?>` (or a character-reference form
   of `1.0`) can therefore enter the lexical append path. The generic source
   validator correctly rejects these through `validate_source_declaration` at
   `xml_splice.rs:888-954`; the relationship parser must apply the same strict
   policy or deliberately refuse declarations. The focused mutations at
   `source_backed.rs:12600-12615` already encode the expected refusal.

2. **Canonical relationship output is allocated without a size-specific
   reservation.** In the canonical branch at
   `source_backed.rs:6750-6790`, `relationship_xml_working_memory_bound` is
   reserved from the original member's `declared_bytes`, then
   `owner_relationships.try_to_xml_bytes()` allocates the changed canonical
   XML. The final `xml.len()` is checked only after that allocation. A small
   canonical source can receive additions or longer replacement fields within
   the relationship limits, making the generated output much larger than
   `declared_bytes`; the retained source-sized reservation does not cover that
   output allocation. Reserve the generated size (and any overlapping
   serializer working state) before `try_to_xml_bytes`, or derive and test a
   bound from the destination relationship limits.

3. **The last recorded integrated native run was still refused by the authored
   compactness gate.** The recorded comparison has 189 rows, 185 unchanged,
   and four cases changed from the old missing-name refusal to
   `Error: Opc(XmlPublication { part: "/ppt/presentation.xml", source:
   NotCompact(FormattingWhitespace, offset 55) })`. The exact evidence is in
   `integrated-outcome-comparison.json`. This is a verified failure of the last
   run, but the source-preserving changes landed after that receipt; acceptance
   still requires a fresh native run proving the intended cases pass.
-->

### Closed in the current source snapshot

- **Empty-child depth admission is closed.** The relationship `Event::Empty`
  branch now checks the child depth, and the focused shallow-depth regression
  covers it.
- **Aggregate relationship-tag accounting is closed.** The parser no longer
  applies the per-attribute limit to the complete `Start`/`Empty` event; the
  combined per-attribute regression covers several individually valid
  attributes.
- **Relationship namespace lexical normalization is closed.** Namespace
  declaration bytes are now required to be their normalized literal value,
  with an escaped-alias refusal regression. The previous auxiliary-binding
  concern is historical.
- **The source and metadata budget changes are closed for this review.** The
  explicit resolver/hash-bucket bound, shared source payload/metadata, and
  pre-clone metadata reservation are present; parent-reported gauge assertions
  cover repeated issuance versus `Arc` clones.
- **Authored fragment capacity is now charged.** `replace` reserves
  `fragment.bytes.capacity()` at `xml_splice.rs:451-455`, so caller-owned spare
  capacity cannot bypass the managed memory budget. The public regression
  `managed_source_xml_refuses_retained_fragment_capacity_over_budget` in
  `crates/litchi-opc/tests/source_xml_publication.rs:484-513` asserts typed
  memory refusal, unchanged output, and zero gauges after cleanup.
- **The first edit-vector growth floor is now conservative.** The capacity
  metadata uses `.max(4)` at `xml_splice.rs:432-437`, covering the four-element
  minimum `try_reserve(1)` allocation for this edit type; later standard growth
  remains covered by the per-edit and capacity charges.

### Final evidence and scope limitations

- The native proof is a self-pair inventory: each selector uses the same
  unmodified archive as source and destination. It demonstrates the candidate
  source-preserving path and final publication outputs for those four cases,
  but does not cover a distinct source/destination package pair.
- The passing external pilots are provider lanes, one direct-bytes and one
  logical-range, over the pinned LibreOffice QA fixture. The Rust workload and
  independent ZIP/XML oracle pass; the smoke-test output is 42,948 bytes with
  37 members and exact copied slide/image checks. No LibreOffice, PowerPoint,
  or other external application opened and saved the candidate output, so this
  evidence does not claim application round-trip interoperability or a speedup
  against the historical refusal.
- Three failed pilot epochs remain under
  `external-pilot-attempts/` as audit evidence. Their failures were verifier
  assumptions about content-type insertion before a preserved newline, the
  `image1-copy1.png` policy, and an explicit PNG override; the corrected final
  pilots pass and no production-library change was made for those oracle fixes.

### Performance acceptance review

The retained `measurements.json` contains 15 provider timing flags whose
absolute change exceeds 5%; none is hidden by the acceptance summary. They are
grouped below with their measured direction:

- Range/media-rich R1: `open_source.p99` **-8.97%**.
- Range/media-rich R2: `open_source.p99` **+5.84%**,
  `open_destination.p99` **+9.93%**, and `open.p99` **+7.87%**.
- Range/plain R2: `open_source.p95/p99` **-14.38%/-21.06%**,
  `open_destination.p95/p99` **-7.01%/-20.23%**, `open.p95/p99`
  **-12.04%/-13.90%**, `plan.p99` **-8.01%**,
  `publication.p95/p99` **-10.13%/-15.58%**, and
  `api_sum.p95/p99` **-10.28%/-13.10%**.

These flags are phase-tail observations. The matched whole-API p50 changes
stay within about 1.3%, while the opposite-direction source-tail result in
media-rich R1 and the negative plain-corpus R2 tails do not support a stable
candidate regression or improvement claim. The measured source-preserving
bytes/range path remains a needed capability enabler for the lossless and
provider coverage above; the phase flags do not justify removing it.

The retained controls use 30 samples after three warmups in fresh children,
with deterministic bootstrap intervals for medians. Logical read, budget, and
work counters are stable across matched lanes, and no RSS comparison crosses
the 5% review threshold. RSS covers the whole child, including setup and the
retained sample loop; allocation counters are unavailable. The records cannot
separate scheduler effects from other run-to-run causes, so they support
measured descriptive flags only and no scheduler-causality, allocator, cold
cache, physical-I/O, scaling, or speedup conclusion.

<!--
The former fourth blocker is retained in the historical section below. Do not
copy it back into the active list unless a fresh run reproduces it.
-->

<!-- Historical wording retained below for exact pre-freeze evidence. -->
<!--
4. **The integrated native target is still refused by the authored compactness
   gate.** The recorded comparison has 189 rows, 185 unchanged, and four cases
   changed from the old missing-name refusal to
   `Error: Opc(XmlPublication { part: "/ppt/presentation.xml", source:
   NotCompact(FormattingWhitespace, offset 55) })`. The exact evidence is in
   `integrated-outcome-comparison.json`. The source-preserving presentation
   splice is therefore not yet demonstrated for the intended native cases;
   generic `verify_authored` must remain bypassed only for a complete source
   proof whose lexical replacement is validated.
-->

### Current implementation notes

- The generic validator now rejects literal `<` in attribute values, `]]>` in
  text/CDATA, malformed comments and QNames, reserved `xmlns` prefixes, loose
  declarations, and depth-limit violations for both `Start` and `Empty`.
  `xml_splice.rs:1000-1145` and the focused tests at `1279-1364` cover these
  cases.
- Namespace normalization parity in the generic validator is now fail-closed:
  namespace declaration raw bytes must equal the normalized UTF-8 value at
  `xml_splice.rs:1032-1040`. The escaped namespace alias appears in the
  regression table at `1285`; the old concern is closed for this path.
- `SourceXmlPart` now shares `Arc` metadata/payload and its metadata reservation
  is acquired before cloning catalog strings (`xml_splice.rs:87-118`). The
  source range proof is non-`Clone`, and edit metadata/capacity is charged at
  `394-457`; the earlier deep-clone and uncharged-edit findings are closed by
  the frozen implementation and its parent-reported gauge assertions.
- Destination content type, destination XML/Part limits, source lineage/version,
  encrypted-entry refusal, and final source fences are checked before source
  XML transfer. The early encrypted gate is at `source_backed.rs:6223-6233`,
  source-XML preflight is at `6254-6265`, and replacement provenance checks are
  at `6327-6338`.
- `TopologyPartPayload::SourceXml` is retained through `Arc<SourceXmlPart>` and
  generic authored compactness checks are skipped only after the source proof
  validates the assembled XML. The ordinary authored path remains audited.
- The relationship helper's old heuristic has been replaced by an explicit
  quick-xml-0.41 resolver/attribute/hash-bucket bound at
  `source_backed.rs:918-980`; the noncanonical append path retains both source
  and append reservations at `6981-7074`; canonical serializer reservations
  now cover both existing and new-member output as documented above.

The review reproductions previously used `/tmp/qxml_probe` and
`/tmp/litchi_probe`; both were removed. The allowed `/tmp/native-copy` and
`/tmp/opc-fuzz` directories were not touched. No new temporary files were
created in this re-review.

The sections below preserve the pre-freeze findings and draft receipts as
historical evidence. Their “in-progress” statuses are superseded by the
current closures and pending validation above.

## Historical findings from pre-freeze review

1. **Closed in the current source snapshot: the original relationship lexical
   admission bug was fixed.** The earlier
   `parse_noncanonical_relationship_source` implementation used
   `quick_xml::NsReader` only as a tokenizer and accepted malformed QNames and
   comments. The final splice was reparsed by this parser but did not pass
   through `validate_overlay_xml`.

   A minimal real `SourceBackedPackage` reproduction used this relationship
   member:

   ```xml
   <bad?:Relationships xmlns:bad?="http://schemas.openxmlformats.org/package/2006/relationships"><bad?:Relationship Id="rExisting" Type="urn:litchi:test" Target="target.xml"/></bad?:Relationships>
   ```

   In the old source epoch, one planned external relationship returned `Ok(())`
   and emitted:

   ```xml
   <bad?:Relationships xmlns:bad?="http://schemas.openxmlformats.org/package/2006/relationships"><bad?:Relationship Id="rExisting" Type="urn:litchi:test" Target="target.xml"/><bad?:Relationship Id="rAdded" Type="urn:litchi:test" Target="https://added.invalid/" TargetMode="External"/></bad?:Relationships>
   ```

   The generated QName was invalid XML. A second old-epoch reproduction with
   `<!--bad--comment-->` inside the root also returned `Ok(())` and preserved
   the invalid comment. The current parser validates NCNames, reserved
   namespace bindings, comments, PIs, controls, and declaration placement.
   Keep regression tests for these cases; the remaining relationship lexical gap
   is the literal `<` in quoted attribute values described below.

2. **The relationship append memory reservation still lacks a complete bound
   proof.** The current code reserves `8 * declared_bytes` for source
   materialization and parser work, then reserves `2 * new_len + append_len`
   for the noncanonical append path. Those factors are documented as
   conservative, but the implementation does not derive them from the actual
   `NsReader` resolver metadata, `HashSet` buckets, owned IDs, temporary decoded
   attribute `String`s, canonical XML output, and allocator capacity overhead.
   The resolver permits up to 256 declarations per start element, so the
   metadata can be material relative to a short source. Keep all overlapping
   reservations (the current canonical path does retain the source reservation)
   and either charge these structures explicitly or document and test a bound
   that covers their worst-case capacities; do not treat the current constants
   alone as a memory proof.

3. **Native unnamed-slide cases reach a generic compactness refusal.**
   The integrated native outcome comparison records four cases changing from
   the old missing-name refusal to
   `XmlPublication(NotCompact(FormattingWhitespace, offset 55))` for
   `/ppt/presentation.xml`:
   [integrated-outcome-comparison.json](integrated-outcome-comparison.json).
   The PPTX planner stages a valid source-preserving presentation edit, but the
   generic topology replacement path applies `xml_minifier::audit::verify_authored`
   to formatted native XML. A preservation-aware lexical replacement/validation
   path is required for the intended native coverage.

## Historical in-progress XML provenance findings

These findings are from a read-only inspection of the in-progress
`crates/litchi-opc/src/xml_splice.rs`; no build or reproducer was run while the
file was being edited.

1. **Namespace-result handling is corrected in the current snapshot.** The
   workspace uses quick-xml 0.41, where
   `NamespaceResolver::set_max_declarations_per_element` and
   `BytesText/CData::xml_content(XmlVersion)` are valid APIs. The current
   validator rejects `ResolveResult::Unknown`, leaves valid `Unbound` results
   alone, and checks element/attribute QNames and namespace declarations
   explicitly. The earlier quick-xml 0.38 API blocker was a review error and is
   withdrawn.

2. **The source XML validator still is not a complete XML grammar check.**
   Comments, predefined/numeric references, controls, duplicate attributes,
   QNames, and reserved namespace bindings now have checks, but quick-xml's
   attribute iterator only scans to the matching quote. A literal `<` in a
   quoted attribute therefore reaches `validate_xml_string`, which permits it.
   Ordinary text also never rejects the XML-forbidden `]]>` sequence. Both
   cases can cross the source-publication boundary and need explicit lexical
   checks and regression tests. The relationship append parser has the same
   attribute-value gap.

3. **Fragment size admission is now present for the public fragment forms.**
   The current API exposes only `AuthoredXmlFragment::markup` and `text`, both
   capped by `MAX_FRAGMENT_BYTES`; the earlier `start_tag`/`end_tag` concern is
   obsolete. The public `Debug` implementation still exposes the complete
   fragment bytes (see finding 14).

4. **Source limits are not transferred to the destination policy.** A
   `SourceXmlPart` keeps the source `ReadLimits`, and its publication check
   validates XML events/depth/attributes with those limits. Topology publication
   only checks destination byte totals for source-XML additions/replacements;
   a destination with stricter XML limits can therefore receive a token that
   violates its own XML policy.

5. **Cross-package source-XML admission is still late.** The current topology
   writer handles `TopologyPartPayload::SourceXml` in the append loop and keeps
   its source snapshot in the monitored transfer sink, so the earlier
   replacement-only/final-fence observation is closed. It still performs the
   first source-XML addition check only after relationship and content-types
   planning has allocated working state, and it does not apply destination XML
   limits (finding 4). Move the source/context and destination-policy checks
   into the preflight phase.

6. **The XML parser working bound is still unproven.** The current code attempts
   an `8 * source_bytes` validation reservation. Namespace resolver buffers,
   event/attribute vectors, duplicate-key storage, and decoded strings are not
   charged individually; the payload plus heuristic reservation is not evidence
   that every transient allocation fits the advertised budget.

7. **Cloning a public range proof bypasses its reservation.** `XmlSourceRange`
   derives `Clone`, which allocates a new `expected: Vec<u8>` but only clones
   the same `_reservation: Arc<Reservation>`. Repeated proof clones can grow a
   managed working set without charging each allocation. Make the proof
   non-Clone, share expected bytes, or charge a custom clone.

8. `SourceXmlPart` also derives `Clone` while deep-copying its `PackURI` and
   `ContentType` strings without a new reservation. If the public snapshot is
   intended to be cheap-to-share, those metadata fields need shared ownership
   or a bounded/custom clone path.

9. `XmlSplicePublication` has no aggregate edit/count ceiling and reserves only
   fragment byte lengths. The `Vec<XmlSpliceEdit>` capacity and per-edit
   metadata remain uncharged, so many tiny non-overlapping edits can consume
   substantially more memory than their byte reservation. Add an edit-count or
   metadata reservation bounded before each insertion.

10. The source capture and splice finish paths fence the source before XML
   validation but not after it. A source version or context cancellation can
   change during that parse, yet `source_xml()`/`finish()` can return `Ok` with
   a stale token; the later topology check may refuse it, but the issuance
   boundary should close with a final source/context check as well.

11. `Event::Decl` is placement-checked only. The validator first requires
    UTF-8 bytes but does not validate the declaration's version/encoding/
    standalone values, so a UTF-8 payload can advertise `encoding="UTF-16"`
    and be interpreted differently by a downstream XML consumer.

12. The Start/Empty event admission currently applies
    `max_xml_attribute_bytes` to the entire raw tag span before checking each
    attribute. That is a policy mismatch with the per-attribute limit: a tag
    with several individually valid attributes can be refused, while the raw
    span check does not replace the required per-attribute lexical checks.

13. `ReadLimitsBuilder::max_xml_depth` accepts values above quick-xml's
    `NamespaceResolver` `u16` nesting counter. A deeply nested source can then
    panic in debug or wrap namespace scope in release before the validator's
    depth check. Clamp the policy or use a resolver without this narrower
    counter.

14. `XmlSourceRange` and `AuthoredXmlFragment` derive `Debug` and expose their
    full byte vectors; `XmlSplicePublication` exposes staged fragments
    transitively. `SourceXmlPart` already uses a redacting custom formatter, so
    these public provenance values should report lengths/offsets rather than XML
    payload contents.

15. The topology writer rejects an encrypted destination only after relationship,
    content-types, and source-XML planning has already allocated working state.
    Encrypted-source refusal should happen before any nonempty topology plan
    reads or stages payloads, preserving the encryption boundary and avoiding a
    budget-consuming refusal path.

16. The generic source validator does not reject the reserved `xmlns` element
    prefix. `validate_source_qname` checks lexical QName shape and rejects only
    `ResolveResult::Unknown`; quick-xml resolves the built-in `xmlns` prefix to
    the XMLNS namespace, so an element such as `<xmlns:x/>` can pass the generic
    source validator even though XML Namespaces reserves `xmlns` for namespace
    declarations. The OPC relationship parser rejects this particular shape via
    its OPC-namespace check, but the public generic source-XML boundary does not.
    Add a reserved-prefix check and a regression test.

## Test and evidence issues

- The updated test suite now accepts absent and empty names and intentionally
  retains source spelling/spacing (including `<p:cSld >`) to exercise the
  source-preserving proof path.
- `write_topology_to_stream` rustdoc contains “Existing Existing” at the
  relationship paragraph.
- The historical eager-reopen failure involving `https://before.invalid/` was
  caused by a fixture that omitted `TargetMode="External"`; the current
  default-namespace fixture includes it. The focused OPC library run observed
  324 passing tests, including the append tests. No namespace-rebinding or
  attribute-normalization parity defect was found: both paths use the same
  quick-xml resolution/normalization behavior.
- A later draft strict-clippy receipt (run by the parent agent; not run by this
  reviewer) failed before source freeze with five compile errors: two accesses
  to private `SourceXmlPart.source` fields, one missing `XmlVersion` import/use,
  and two borrows of a temporary `attribute.key.as_ref()` value. These are
  draft-epoch compile blockers owned by the coder; the current snapshot has the
  quick-xml 0.41 namespace-cap API and `XmlVersion` import, so the withdrawn
  dependency/API claim must not be used as a finding.

## Temporary review artifacts

The reproductions were created outside the repository under `/tmp/qxml_probe`
and `/tmp/litchi_probe`; both were removed after recording these findings.
The reviewer did not modify the allowed `/tmp/native-copy` or
`/tmp/opc-fuzz` directories.
