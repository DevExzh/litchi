# PPTX `inkAction` package-owner design

Status: bounded design and specification revalidation. This document does not
add production code, choose an unverified OPC constant, or claim native Office
acceptance. It closes the design portion of audit priority 7 after the shared
DrawingML action profile was added.

Review date: 2026-09-12.

The existing `litchi-drawingml::ink::actions` reader, detached builder, and
source-backed editor are a useful neutral value layer. They read and preserve
one `iact:actions` XML value, retain opaque InkML boundaries, and provide
semantic action selectors, reversible source-checked patches, and bounded
construction. They do not identify a PresentationML owner, resolve an OPC
relationship, allocate a part, or execute an action. The package owner belongs
in `litchi-pptx`; its generic Content-part boundary is established by the
checked ECMA-376 and PresentationML rules. A producer-specific MIME/path
profile and native interoperability claim remain evidence-gated.

## Decision

The checked specifications establish the PresentationML anchor and the action
XML vocabulary. ECMA-376 also establishes the generic Content-part contract:
`p:contentPart` uses the exact custom-XML relationship for the package dialect
(Strict or Transitional), the target is an internal XML part, and `text/xml`
is the MIME value to declare when that XML format has no explicit MIME type;
consumers identify the contents by root namespace. A package still must have
an effective `[Content_Types].xml` declaration; a missing content-type mapping
is not treated as an implicit `text/xml` fallback. What remains absent is an
**action-specific** MIME, path convention, or producer profile. The
implementation therefore has two layers:

1. The normative generic owner accepts the supported InkAction MCE anchor, the
   exact custom-XML relationship selected by the package dialect, an internal
   XML target, an effective `[Content_Types].xml` declaration, and the
   `iact:actions` root. The initial generic content-type policy is a declared
   `text/xml`; another explicit XML MIME is accepted only when the caller or a
   producer profile declares it supported. A producer-specific
   `ActionPartContract` is not required merely to read or edit an existing
   generic Content part.
2. An optional evidence-backed producer profile records an observed MIME,
   target-name convention, or other Office interoperability rule. It is
   required before claiming native acceptance or selecting a producer-specific
   fresh-part policy. A direct `p:contentPart` or an `iact:actions` root found
   without the normative ink-action anchor is still insufficient to expose a
   typed PPTX action object.

This is deliberately stricter than treating the shared XML profile as proof of
host ownership. The existing InkML owner is a separate contract and cannot
fill the missing action fields by analogy.

| Area | Established result | Design consequence |
| --- | --- | --- |
| Action XML | `iact:actions` in `http://schemas.microsoft.com/office/powerpoint/2014/inkAction`, with the `[MS-ODRAWXML]` §2.21 grammar | Reuse the shared detached profile after package closure has been validated |
| PresentationML anchor | Ink extension is an MCE `Choice` in the fixed MCE namespace `http://schemas.openxmlformats.org/markup-compatibility/2006`, requiring the PowerPoint 2010 main URI and the 2014 `inkAction` URI; its child is `p:contentPart`; fallback is `p:pic` | The owner is the enclosing shape tree/group and its selected MCE branch, not the action XML alone; this MCE namespace is unchanged in Strict packages |
| Generic relation | ISO/IEC 29500 `contentPart` requires the package dialect's custom-XML relationship type | Infer Strict versus Transitional from package conformance/root namespaces, require the exact matching URI, and reject a mismatch; do not derive it from the InkML owner |
| Action part content type | No action-specific MIME is enumerated; ECMA-376 supplies `text/xml` as the value to declare when no explicit XML MIME exists and says the root namespace identifies the content | Require an effective `[Content_Types].xml` declaration of `text/xml` or another explicitly supported XML MIME; missing mapping is an error and `application/inkml+xml` is not inferred |
| Action target path | No fixed path is specified and no native action fixture was found | Resolve the explicit relationship target and retain its URI; a fresh producer path requires an explicit policy and carries no native-acceptance claim |
| Owner cardinality | Generic Content parts are zero or more and are reached by explicit owner relationships; no one-to-one target rule was found | Model incoming references as a graph; allow a target to be shared and check all owners before removal |
| Execution/rendering | Not supplied by the shared profile or this design | No playback, recognition, rendering, or action execution API |

## Normative revalidation

The following local sources were read rather than inferred from the current
implementation.

### PresentationML anchor

Local `[MS-PPTX]` §2.2.3.1, `3rdparty/specs/[MS-PPTX]/2 Structures/2.2 Extensions.md`,
extends `spTree` and `grpSp` with an `mc:AlternateContent` whose table has:

- a `Choice` whose resolved `Requires` list is exactly the two distinct URIs
  `http://schemas.microsoft.com/office/powerpoint/2010/main` and
  `http://schemas.microsoft.com/office/powerpoint/2014/inkAction`, with no
  additional namespace;
- a `p:contentPart` child in that selected choice; and
- a `p:pic` child in `mc:Fallback`.

The `Requires` values are namespace URIs after XML namespace resolution. A
prefix such as `p14` or `iact` has no identity by itself. The exact-two rule
rejects duplicate or additional required namespaces. A reader must retain
the complete `AlternateContent` source, including every inactive choice and the
fallback, even when an active branch is available. The active projection used
by a normal MCE helper is not enough for a lossless editor because it discards
source that the fallback or another consumer may need.

The MCE namespace itself does not vary with package conformance. ECMA-376
Part 3, fifth edition (December 2015), §7.1 (Markup Compatibility and
Extensibility), in the local source archive
`3rdparty/specs/ECMA-376/ECMA-376-3_5th_edition_december_2015.zip`, requires
all MCE elements and attributes to use
`http://schemas.openxmlformats.org/markup-compatibility/2006`. This fixed URI
therefore applies to both Transitional and Strict packages, including
`AlternateContent`, `Choice`, `Fallback`, and `mc:Ignorable`. Strict
conformance changes the PresentationML namespace and the applicable
relationship URI; it does not substitute a second Strict-only MCE namespace.
The owner scanner must accept the fixed MCE URI under a Strict PML root and
must not classify a `purl.oclc.org/ooxml/markup-compatibility/2006` URI as the
normative Strict MCE namespace.

The companion `[MS-PPTX]` Part Enumerations file has no InkAction part
enumeration. Its absence is evidence about the checked local contract, not a
proof that no producer-specific contract exists elsewhere.

### Action XML

Local `[MS-ODRAWXML]` §2.21 and its schema appendix §5.19 define the action
namespace and XML content model. In the bounded profile, `iact:actions` has
`lengthUnit` and `timeUnit`, an optional first `inkml:definitions`, and an
unbounded source-order choice of `actionGroup` and `action`. Actions contain
properties followed by action data or data groups and carry `type` and
`startTime`. The reserved `add`, `remove`, and `transform` conventions and
custom strings are structural/semantic data for the shared profile.

Those sections do **not** define an action-specific OPC content type, a target
path, or a separate PPTX part enumeration. The schema import of InkML does not
turn the action root into the separately enumerated Ink Content Part. The
generic ECMA-376 Content-part rules still apply because the PresentationML
anchor is `contentPart`; they supply the relationship and generic XML/MIME
boundary, while §2.21 supplies the action root classification.

### Generic content-part and OPC rules

The local ECMA-376 Part 1 fifth-edition PDF in
`3rdparty/specs/ECMA-376/ECMA-376-1_5th_edition_december_2016.zip` was checked
at §15.2.4 (Content Part), §17.3.3.2 (`contentPart`), and §13.3.8 (Slide
Part). The resulting boundaries are:

- a Content part is XML with a root determined by its content type and may use
  any supported XML MIME type; where a format has no explicit MIME type,
  `text/xml` shall be used and the consumer should determine the contents by
  the root namespace;
- a `contentPart` anchor carries `r:id`, and its relationship type is the
  dialect's custom-XML relationship type;
- Content parts are zero or more and are reached by explicit relationships from
  PresentationML slide, slide-layout, slide-master, notes-slide, notes-master,
  or handout-master owners (as applicable to the package grammar);
- a Content part is an internal target; it has no implicit or explicit
  relationships to other ECMA-376-defined parts; and
- the generic rules do not say that a target has exactly one inbound
  relationship. The graph must therefore retain all inbound edges and must
  not delete a shared target after removing one anchor.

For a Strict package the relationship URI is exactly
`http://purl.oclc.org/ooxml/officeDocument/relationships/customXml`. The
Transitional Microsoft Ink Content Part enumeration uses
`http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml`,
which is the existing `litchi-opc::constants::relationship_type::CUSTOM_XML`.
The action owner infers the package dialect from the package's conformance and
root namespace declarations, then requires the exact matching URI. A
relationship URI from the other dialect is a mismatch; no InkML analogy or
producer-specific profile is needed to select between these normative
substitutions.

This relationship substitution is independent of MCE namespace resolution:
the owner uses the fixed ECMA-376 MCE URI above in either dialect, while it
uses the Strict or Transitional PresentationML and custom-XML relationship
URIs according to the package conformance.

The local `[MS-ODRAWXML]` §2.1.4 Ink Content Part enumeration is concrete but
limited to an InkML part:

| Field | Enumerated InkML value |
| --- | --- |
| Content type | `application/inkml+xml` |
| Relationship | Transitional `.../relationships/customXml` |
| Root | `ink` in `http://www.w3.org/2003/InkML` |
| Owner | an explicit `contentPart` relationship |

The existing `crates/litchi-pptx/src/presentation/embedded/ink/` path and
`tests/pptx_ink_annotations.rs` implement and test that InkML contract. They
do not contain `iact:actions` and are not action-package evidence.

## Owner boundary and proposed package model

The package owner should be a new `litchi-pptx` surface, for example
`presentation::embedded::ink_actions`, layered as follows:

```text
PPTX owner (litchi-pptx)
  -> owner-part XML and MCE branch catalog
  -> owner .rels and [Content_Types].xml
  -> explicit internal target and inbound-reference graph
  -> generic Content-part contract + optional producer profile
  -> litchi-drawingml::ink::actions::Profile / Edit / Prepared
```

The shared crate remains independent of OPC and archive types. It continues to
provide the detached value, source spans, opaque definitions/transforms/traces,
and action-level semantic selectors. It must not learn a PPTX path or
relationship mutator.

The first implementation batch should be **slide-owner read and edit only**.
The generic ECMA owner list is normative and is broader than this batch;
layout, master, notes, and handout owner APIs are deferred implementation
seams, not normative ownership uncertainties. Those parts remain in the
generic source-preserving graph until their owner-specific code and tests are
added. A future slice can enable each additional owner explicitly without
changing the graph model.

An inventory entry should retain at least:

- the checked slide selector and owner-part kind;
- a zero-based source ordinal and semantic ordinal for the anchor;
- the enclosing `AlternateContent` source span, selected branch identity, all
  inactive branch bytes, and fallback bytes;
- the raw `r:id`, relationship type, target mode, resolved target part name,
  declared content type, and target bytes;
- the action root namespace/name classification and shared-profile source;
- all inbound edges to the target and any unexpected target relationships; and
- a source revision/fingerprint covering the owner XML, owner relationships,
  target bytes, content-type declaration, and selected MCE metadata.

Physical values are retained for diagnostics and publication authorization.
They are not the ordinary semantic identity exposed to callers.

## Classification and preservation matrix

Classification must happen before the shared action profile is exposed. A
namespace-aware raw scanner should record branch spans and resolved names, then
parse only the selected candidate. It must not serialize an MCE projection and
reconstruct the source.

| Source shape | Typed result | Preservation and edit rule |
| --- | --- | --- |
| `AlternateContent` under a supported `spTree`/`grpSp`, exactly one selected `Choice` whose resolved `Requires` list is exactly the two required URIs with no extras, selected child is `p:contentPart`, the package-dialect custom-XML relation is internal, the target has an effective declared `text/xml` (or explicitly supported XML) content type, and target root is `iact:actions` | Typed action anchor | Read the shared profile. A replacement may touch the target bytes and the selected anchor only after full closure validation. |
| Direct `p:contentPart` without the ink-action MCE owner | Generic opaque content part, even if its target happens to have an `iact:actions` root | Preserve the anchor, target, relationship, and content type. Return `UnsupportedOwnerShape` for a typed action request. Do not infer ownership from the shared fragment. |
| Generic content-part `AlternateContent` whose `Choice` requires the PowerPoint 2010 main URI but not the 2014 `inkAction` URI | Generic opaque content-part branch | Preserve the whole MCE and resolve it only through the generic content-part reader. It is not promoted to an action anchor merely because it contains `p:contentPart`. |
| MCE with unknown/unresolved `Requires`, an additional required namespace, duplicate supported choices, missing choice, or malformed branch structure | Opaque MCE plus a typed unsupported/ambiguous diagnostic | Retain the complete enclosing MCE bytes. Do not select a branch by prefix, guessed order, or projection. |
| Supported choice with missing/duplicate/unresolved `r:id`, external target, missing target, wrong relation URI, missing content-type declaration, non-XML/unsupported content type, or root mismatch | Opaque graph plus a typed closure error | Keep source unchanged. Do not treat missing `[Content_Types].xml` data as `text/xml`, fall back to the InkML contract, or repair the package during save. |
| Supported action choice with `mc:Fallback` `p:pic` | Typed action anchor and retained fallback | Edit only the selected action branch; replay the fallback byte-for-byte. A fallback is not an alternate action payload. |
| Target has unknown XML only inside a shared-profile admitted opaque boundary (`inkml:definitions`, `iact:transform`, `inkml:trace`, or `inkml:traceView`) | Typed structural profile with bounded opaque payloads | Retain the complete admitted element and lexical source. An edit that needs to reinterpret an opaque reference or identity refuses until its closure is modeled. |
| Target has an unknown direct `iact` child, an unknown schema-owned attribute, invalid action ordering, or another direct action-grammar violation | `ActionGrammarError` typed refusal | Retain the target bytes through the generic source-preserving path when possible, but do not expose the invalid subtree as an opaque typed action payload or perform a scalar profile edit. |
| Target has custom action/data/property scalar strings admitted by the shared profile | Typed structural profile with those scalar values | Preserve their lexical values; they are not an excuse to treat an unknown direct element as opaque. |
| Target has an unknown outbound relationship | Typed read and scalar profile edit with a diagnostic | Preserve the relationship and bytes exactly. Any graph-changing edit or deletion that could affect its unknown closure refuses until that closure is modeled. |
| Target has an outbound relationship to a known ECMA-376-defined part or another known Content-part violation | Typed conformance refusal | Preserve the source for generic opening/saving if possible, but do not publish the typed action view or an edit that would rely on the violated graph. |

For an unsupported or malformed candidate, the package may still be opened and
saved through the existing generic source-preservation path if that path can
retain it. The typed action API must fail closed. Validation never changes the
source, and ordinary save never repairs it.

## Selectors and identity

The package selector must carry both semantic and source coordinates without
making a physical relationship ID the public identity.

```text
InkActionAnchorSelector::Semantic {
    owner: SlideSelector,
    ordinal: usize,
    context: AnchorFingerprint,
}

InkActionAnchorSelector::Source {
    owner: SlideSelector,
    ordinal: usize,
    context: AnchorFingerprint,
}
```

The definitions are:

- **Semantic ordinal** is zero-based depth-first document order among active,
  contract-valid ink-action anchors in the selected owner part. An anchor is
  counted once even if its target is shared by another anchor. This is the
  normal selector for an action view and is resolved against the immutable base
  snapshot.
- **Source ordinal** is zero-based document order among raw ink-action MCE
  anchor spans in that owner part, including spans that are currently
  unsupported or inactive. It identifies a source-preservation location, not a
  license to reinterpret an unknown branch. Resolving it must still verify the
  context fingerprint and branch classification.
- `AnchorFingerprint` covers the enclosing owner kind/path, selected branch
  namespace set, source span hash, and expected relationship/content closure.
  It is a conflict check, not a value substituted into the Office file.
- `r:id`, target part URI, relationship type, and part-name strings remain
  diagnostics and internal closure coordinates. They are not accepted as the
  ordinary public selector.

The action inside a selected target uses the existing shared
`ActionSelector`: its depth-first `Ordinal` is the default, with direct-root
and group-relative forms where a caller already has the semantic structure.
Source spans are retained for diagnostics and exact patching. Any identity
changing `xml:id`/`ref` edit remains refused until the package owner models the
complete reference closure.

Source and semantic ordinals must never be recomputed against a partly staged
package to make an operation appear to succeed. Every operation resolves on
the base snapshot, records the base context, stages its changes, and lets the
post-write inventory report the new ordinals.

## Read, edit, and commit lifecycle

### Read and inventory

1. Open the owner part under the existing XML and package limits.
2. Scan raw owner XML for `mc:AlternateContent`, `mc:Choice`, `mc:Fallback`,
   `spTree`, `grpSp`, and `p:contentPart` spans. Resolve namespace URIs, never
   prefix strings.
3. Build the active semantic projection for classification while retaining all
   original branch bytes and source spans.
4. Resolve each selected `r:id` in that owner's `.rels`; require the exact
   custom-XML relationship for the inferred package dialect, an internal
   target, an effective `[Content_Types].xml` declaration of `text/xml` (or
   another explicitly supported XML MIME), and a target member before typed
   classification. A missing declaration is an error, not an implicit
   `text/xml` value.
5. Check the target root and generic content-type policy. Only then call the
   shared `read_profile`; retain the exact target bytes alongside the profile.
6. Build inbound-reference indexes for the target. Preserve unknown outgoing
   relationships as diagnostics for reads and scalar profile edits; reject a
   known ECMA-376 graph violation, and reject graph-changing edits or deletion
   when an unknown outgoing closure would be affected.

The generic `content_parts` implementation is a useful seam for raw anchor
and relationship retention, but its opaque payload API is not the typed action
owner. Its MCE projection cannot replace the action snapshot's complete raw
branch catalog.

### Existing-target profile edit

The first write operation should be `replace_profile` or a delegated shared
profile edit on an already classified target. It should:

1. resolve the package anchor by semantic selector and verify the source
   fingerprint;
2. clone the shared profile edit and apply action/property/data operations by
   the shared semantic selectors;
3. finish the target XML under both package and action-profile limits;
4. stage only the target bytes and any selected anchor bytes that actually need
   changing; retain the inactive MCE choice and fallback exactly;
5. validate the complete relationship/content-type/root closure; and
6. reopen the complete candidate package, inventory the same semantic selector,
   and compare the requested action state before publication.

Fresh action-part creation, anchor insertion, and anchor removal are separate
operations. A generic package-conformant writer may use `text/xml` under an
explicit generic-content policy, but that choice carries no native Office
acceptance claim. A producer-compatible writer should wait for an evidence
profile. In either case, insertion must author the full MCE shape, allocate an
explicit relationship and content-type declaration, and use a caller-supplied
or owner-local source policy for any new part name. It must not copy the InkML
writer's `/ppt/ink/inkN.xml` convention without evidence.

### Graph and deletion rules

The transaction treats the following as one dependency closure:

```text
owner XML MCE span
  -> owner .rels r:id
  -> internal target part
  -> [Content_Types].xml declaration
  -> inbound-reference set and any target relationship closure
```

One owner part may contain zero or more action anchors. A target may have zero,
one, or multiple inbound relationships in the graph even though a particular
anchor has one `r:id`. Removing an anchor therefore defaults to detaching that
edge. Deleting the target, its content-type override, or another inbound edge
requires a separate explicit disposition (`retain`, `cascade`, `retarget`, or
typed refusal); the owner must not guess. Unknown outbound relationships are
retained and diagnosed during reads and scalar profile edits. They block a
graph-changing edit or deletion until their closure is modeled; a known
ECMA-376-defined outbound violation is a typed conformance refusal.

The shared action profile's `xml:id`, `ref`, and transform/path conventions
also form a semantic closure. If a requested operation would change an ID or
remove a referenced data node and the complete reference graph is not modeled,
return a typed refusal before staging bytes.

### Commit, patch, inverse, and stale checks

The package transaction follows ADR 0003:

- `commit()` returns a new immutable package snapshot, a reversible package
  patch, and diagnostics; it never mutates the base snapshot;
- an exact semantic no-op reuses the original source allocation and bytes and
  does not perform a redundant candidate reopen;
- a real commit stages a clone, applies bounded replacements, validates the
  full package graph, reopens the candidate, and publishes atomically only
  after typed readback matches;
- `Patch::inverse()` restores the exact accepted source artifact in memory;
  patch application requires the exact source revision/artifact and the same
  owner/anchor context; and
- signed-source policy, source provenance, and any package read options are
  part of authorization. There is no last-writer-wins path.

A patch is stale if the owner XML, `.rels`, target bytes, content-type
declaration, inbound-reference graph, or enclosing MCE branch changed after
the snapshot. A changed namespace declaration that changes the resolved
`Requires` set is also stale. A stale or ambiguous source returns a typed
conflict; it is never silently rebased by ordinal or relationship ID.

## Bounds and failure behavior

The owner should use one hierarchical limit object whose package component is
stricter than or equal to the existing package limits and whose action component
is stricter than or equal to the shared profile limits. The initial design
reuses the following ceilings rather than creating an unbounded second parser:

| Resource | Initial ceiling to carry into the owner |
| --- | ---: |
| Owner/slide XML | existing content-part owner limit (32 MiB) |
| One action target | shared action source limit (16 MiB) |
| One generic content payload | existing content-part payload limit (64 MiB) |
| Aggregate retained payload bytes | existing package aggregate limit (256 MiB) |
| Action XML depth/nodes | shared profile limits (depth 128, nodes 100,000) |
| Action attributes and scalar values | shared profile limits (256 attributes per element, 256-byte typed tokens, and 1 MiB attribute values) |
| Namespace declarations per element | shared profile ceiling (256 declarations) carried by the owner scanner and action parser |
| Action records/groups | shared profile limits (65,536 actions and 16,384 groups) |
| Action anchors/relationship edges | new owner counters, initially 4,096 anchors and 4,096 payload relationships |
| MCE branches and namespace declarations | owner counters charged per retained raw branch and declaration |

The exact constants should be named in the future owner rather than copied as
unrelated literals. The `256` namespace-declaration ceiling applies per XML
element, including MCE and action payload elements; the owner must carry it
through raw scanning instead of relying on an unbounded namespace resolver. A
shared target is charged once for retained target bytes and once per reference
edge for graph metadata. Every size, depth, node, relationship, allocation,
and integer operation is checked before retention; fallible allocation and
cancellation are required. Limits apply to raw source and to staged output,
so an opaque unknown branch cannot bypass the budget.

Limits and errors should distinguish at least: source too large, target too
large, aggregate payload too large, MCE branch/namespace limit, graph/edge
limit, action profile limit, malformed XML, unsupported owner/contract,
external/missing target, wrong relationship/content type/root, ambiguous MCE,
and stale source. Validation must not mutate and normal save must not repair.

## Native fixture and evidence status

No native or independently validated PPTX containing the action namespace was
available in the inspected corpus. The retained raw-marker scan in
`ink-action-profile/packaging-scan.json` covers 1,021 candidates (1,003 valid
packages), 42,578 members, and 42,117 successfully read members; it found no
`iact:`, `inkAction`, action-namespace, or `2014/inkAction` marker. The scan
records 18 malformed/unreadable candidates and is bounded negative evidence,
not proof of absence or Office acceptance. Its receipt hash is
`f5734d000672b00b6cc8e925026a554c9c009d0c20a223838ea561272c453fe9`.

The parent agent's independent member survey of the `test-data` and vendored
LibreOffice/PPTX/PPTM corpus likewise found no `inkAction`, `/ink`, `inkml`, or
`application/inkml` marker in 9,021 scanned XML/relationship members. The two
malformed LibreOffice inputs
`sd/.../fail/ofz46160-1.pptx` and `filter/.../empty.pptx` are excluded from the
negative result. These searches do not establish a hidden escaped or
producer-specific owner contract.

The smallest evidence-completing fixture must be retained byte-for-byte with a
package hash and producer/version metadata and include:

1. `[Content_Types].xml`, the owner part, owner `.rels`, target member, and
   exact target root bytes;
2. an `AlternateContent` with arbitrary prefixes, `mc:Ignorable`, both
   required choice URIs, one `p:contentPart`, and a real `p:pic` fallback;
3. at least one direct action, one grouped action, a property, and an opaque
   definition/transform/trace boundary;
4. exact target mode, relationship URI, declared content type, target path,
   and any target outgoing relationships; and
5. negative variants for wrong content type/relationship, missing or duplicate
   `r:id`, inactive-only action choice, external/missing target, duplicate
   supported choice, shared target/orphan deletion, malformed MCE, stale source,
   exact no-op, and inverse patch.

Until this fixture or an additional normative binding exists, the following
are intentionally unsupported or uncertain:

- any producer-specific action MIME beyond the generic `text/xml` fallback,
  and any observed `[Content_Types].xml` default/override convention;
- an action-specific relationship URI or relationship variant beyond the
  generic dialect-specific `contentPart` custom-XML rule (the Strict and
  Transitional substitutions are normative and selected from package
  conformance);
- producer-specific action target path/name allocation and action/InkML
  co-location (an existing target URI is always resolved from its relationship);
- native Office producer acceptance and round-trip behavior;
- complete InkML semantic validation, trace interpretation, recognition,
  playback, rendering, and execution.

## Implementation and review gates

The bounded work can be split without changing the contract:

1. **Evidence gate:** record any producer `ActionPartContract` and fixture;
   reject implementation that substitutes the InkML contract or guesses a
   producer path. The generic `text/xml` reader/writer policy remains separate
   and must not be presented as native Office interoperability.
2. **Coder:** add a `litchi-pptx` owner snapshot/inventory and source-preserving
   MCE scanner; reuse the shared profile only after relationship, content type,
   root, and target mode checks.
3. **Reviewer:** compare every constant and owner kind against `[MS-PPTX]`,
   `[MS-ODRAWXML]`, ECMA-376, the retained fixture, and ADRs 0003, 0006, 0007,
   0011, and 0024. Verify that no physical ID leaks into the ordinary API.
4. **Tester:** add exact replay/no-op/inverse/stale tests, graph closure and
   shared-target tests, all MCE preservation and refusal cases, package limits,
   malformed XML, external targets, and bounded fuzz inputs. Reopen every real
   candidate and compare semantic readback.
5. **Profiler:** measure named read/edit scenarios with the repository's
   explicit performance context and retained bounds. Report source bytes,
   allocations, peak memory, and wall time for the fixture; make no native or
   rendering performance claim from this design.

These gates implement the preservation, bounded-resource, source-checked,
selector-first, atomic-publication goals in `docs/GOAL.md` and the accepted
snapshot/patch, compatibility, Office-owner, and crate-topology ADRs. They do
not broaden the shared action profile into a package owner or an execution
engine.

## References

- `[MS-PPTX]` §2.2.3.1: `3rdparty/specs/[MS-PPTX]/2 Structures/2.2 Extensions.md`.
- `[MS-PPTX]` Part Enumerations: `3rdparty/specs/[MS-PPTX]/2 Structures/2.1 Part Enumerations.md`.
- `[MS-ODRAWXML]` §2.1.4: `3rdparty/specs/[MS-ODRAWXML]/2 Structures/2.1 Part Enumerations.md`.
- `[MS-ODRAWXML]` §2.21: `3rdparty/specs/[MS-ODRAWXML]/2 Structures/2.21 http---schemas.microsoft.com-office-powerpoint-2014-inkAction.md`.
- `[MS-ODRAWXML]` §5.19 schema: `3rdparty/specs/[MS-ODRAWXML]/5 Appendix A - Full XML Schemas/5.19 http---schemas.microsoft.com-office-powerpoint-2014-inkAction Schema.md`.
- ECMA-376 Part 1 fifth edition, §§13.3.8, 15.2.4, and 17.3.3.2, in `3rdparty/specs/ECMA-376/ECMA-376-1_5th_edition_december_2016.zip`.
- Shared value/editor: `crates/litchi-drawingml/src/ink/actions.rs` and `actions_edit.rs`.
- Existing generic graph seam: `crates/litchi-pptx/src/presentation/embedded/content_parts/`.
- Separate InkML owner: `crates/litchi-pptx/src/presentation/embedded/ink/` and `crates/litchi-pptx/tests/pptx_ink_annotations.rs`.
- Existing packaging evidence: `docs/report/spec-gap-validation-evidence/ink-action-profile/packaging-evidence.md` and `host-next-steps.md`.
- Accepted constraints: `docs/GOAL.md`, `docs/adr/0003-snapshots-edits-and-patches.md`, `docs/adr/0006-validation-security-and-compatibility.md`, `docs/adr/0007-office-object-models.md`, `docs/adr/0011-ooxml-physical-package-ownership.md`, and `docs/adr/0024-current-topology.md`.
