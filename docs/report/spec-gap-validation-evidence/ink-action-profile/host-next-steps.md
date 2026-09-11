# PPTX InkAction host integration: evidence and next steps

Status: design evidence for a remaining audit row. This note does not claim
PPTX InkAction package ownership, native interoperability, action execution, or
completion of the shared ink row. It records the smallest next work that can be
implemented once the part packaging is established by a real producer fixture
or an additional normative source.

Review date: 2026-09-11. The current shared profile was reviewed after the
namespace, XML-character, and empty-property corrections. The profile is a
bounded structural reader/replayer; it is not a package owner and it does not
execute or render actions.

## What the local specifications establish

The action vocabulary is specified in `[MS-ODRAWXML]` section 2.21,
`3rdparty/specs/[MS-ODRAWXML]/2 Structures/2.21 http---schemas.microsoft.com-office-powerpoint-2014-inkAction.md`:

- `iact:actions` is `CT_Actions`. It has optional first `inkml:definitions`,
  then an unbounded choice of `actionGroup` and `action`, required
  `lengthUnit` and `timeUnit`, and optional `xml:id`.
- `action` has zero or more `property` elements followed by zero or more
  `actionData` or `actionDataGroup` elements. `type` and `startTime` are
  required; `startTime` is `xsd:decimal`.
- `actionData` permits an optional first `transform`, followed by `trace` or
  `traceView` children. Its `xml:id`, `name`, and `ref` attributes are optional
  (the schema defaults `name` to `stroke` and `ref` to the empty value).
  `actionDataGroup` and `actionGroup` each require at least one child.
- `property` requires `name`, optionally carries `value` (default `ink`), and
  has an empty content model. Thus text, whitespace character references, and
  nonempty CDATA are invalid inside `property`, even though indentation around sibling
  elements is valid. The shared strict profile now enforces this boundary.
- The prose assigns reserved conventions: `add` uses one `stroke` or `path`
  data child, `remove` uses one `stroke` child, and `transform` uses two
  children named `target` and `path`; reserved property names include
  `dataType` and `style`. Custom action, data, property, and value strings
  remain possible. Those conventions and reference closure are a semantic
  validation layer beyond the current inert structural profile.

The PresentationML package placement is only partly specified by the local
documents. `[MS-PPTX]` section 2.2.3.1, `3rdparty/specs/[MS-PPTX]/2 Structures/2.2 Extensions.md`,
states that a slide or group `spTree` can contain an `mc:AlternateContent`
whose `Choice` requires both
`http://schemas.microsoft.com/office/powerpoint/2010/main` and
`http://schemas.microsoft.com/office/powerpoint/2014/inkAction` and contains a
PresentationML `contentPart`; its `Fallback` is a `pic`. The host therefore
owns an MCE branch, a slide relationship, and a target part as one closure.
Prefix spelling is not identity: the choice and action namespace must be
resolved by URI, while the inactive choice/fallback bytes remain preservation
data.

`[MS-ODRAWXML]` section 2.1.4, `3rdparty/specs/[MS-ODRAWXML]/2 Structures/2.1 Part Enumerations.md`,
does define an **Ink Content Part** with content type
`application/inkml+xml`, source relationship
`http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml`,
and an `inkml:ink` root. That section requires the part to be the target of an
explicit internal relationship from a `contentPart` owner. It does not define a
second part enumeration for an `iact:actions` root. Section 2.21 defines the
action XML vocabulary but, in the checked local copy, does not bind it to a
distinct content type, relationship URI, or fixed `/ppt/...` location.

This leaves an evidence gate that must not be filled with a filename or a
guessed constant. The existing writer's `/ppt/ink/inkN.xml` plus
`customXml` plus `application/inkml+xml` convention is proven only for an
`inkml:ink` payload by `crates/litchi-pptx/src/presentation/embedded/ink/` and
`crates/litchi-pptx/tests/pptx_ink_annotations.rs`. It is not evidence that an
`iact:actions` root uses the same part contract. The host must first obtain a
PowerPoint/Office package (or a further normative statement) containing
`iact:actions`, then record its `[Content_Types].xml`, slide `.rels`, target
member name, MCE anchor, root namespace, and provenance hash. Until then, an
action candidate should be classified from its explicit relationship,
declared content type, and root namespace, and an unknown combination should
remain opaque or be refused as unsupported. The implementation must not scan
all `/ppt/ink/` members or infer ownership from a part name.

## Correct public owner and graph boundary

The package owner belongs in `litchi-pptx`, because it resolves slides,
PresentationML `contentPart` anchors, MCE branch selection, OPC relationship
direction, content types, target parts, and save/readback. The shared
`litchi-drawingml` crate should own only the detached `iact:actions` value and
its bounded codec. A shared codec must not add package relationship mutation or
invent a PPTX path.

The current surfaces show the required split:

- `presentation::embedded::ink::{load_slide,store_slide}` is a low-level
  InkML inventory/storage helper. It requires `application/inkml+xml`,
  `customXml`, and the `inkml:ink` summary, and its writer allocates
  `/ppt/ink/inkN.xml`. It is not an InkAction owner.
- `Presentation::content_parts()` and
  `presentation::embedded::content_parts` already expose an opaque,
  source-retaining slide content-part graph. They intentionally do not parse
  or interpret the related XML, so they cannot be the typed action facade by
  themselves.
- The ordinary facade should eventually expose a typed inventory below the
  `Presentation`/`Slide` owner (for example, a slide-selected
  `ink_actions()` view) and a separate source-backed editor. Selectors should
  be checked slide position or producer-visible slide name plus semantic
  content-part order. `r:id`, part URI, and relationship type can be returned
  as diagnostics, but should not be the ordinary selector vocabulary.
- The source-backed path must use the existing immutable source catalog and
  exact-source publication model. A low-level `Package::opc()` mutation or
  generic payload replacement is not sufficient for an ordinary action API,
  because it can bypass MCE fallback preservation, relationship closure, and
  typed readback.

For an action inventory, each result needs at least the slide selector and
anchor position, selected MCE branch, exact anchor bytes or a bounded source
span, relationship id/type/target mode, resolved target part name, declared
content type, root classification, and the retained action source. Physical
metadata is useful for diagnostics and publication authorization; it should
not become the semantic identity exposed to callers.

## Snapshot, edit, and publication contract

An action snapshot should capture one complete dependency closure:

1. the slide XML source and source revision;
2. the active `contentPart` anchor and its enclosing `AlternateContent` span,
   including the inactive branch/fallback source needed for exact preservation;
3. the slide relationship record and resolved target part;
4. the target content type and exact action bytes; and
5. any additional relationships or same-part references discovered by the
   selected action profile.

An edit can then provide typed operations such as replacing the detached
action profile, adding an action content part at a checked anchor, or removing
one only after incoming-reference and fallback disposition checks. It must
stage a candidate package, validate the relationship/content-type/root
contract, serialize the changed action, reopen the complete candidate under
the retained limits, and read back the selected action projection before
returning a commit. The commit should carry a reversible patch and diagnostics
consistent with ADR 0003 (`docs/adr/0003-snapshots-edits-and-patches.md`).
Publication must refuse a stale source, ambiguous MCE branch, duplicate
relationship id, missing target, external target, wrong relationship type,
wrong content type, or root/namespace mismatch. A removal that would orphan a
shared target or discard an incoming reference needs an explicit disposition;
the host must not guess between cascade, detach, retarget, and retain.

The preservation rules are equally important:

- An exact semantic no-op should retain the complete input artifact and all
  ZIP records byte-for-byte. An unchanged action payload should retain its
  original declaration, prefixes, comments, CDATA/opaque InkML payloads, and
  lexical whitespace.
- A changed action may use canonical serialization for the changed semantic
  value, but the MCE inactive branch, fallback picture, unrelated slide
  markup, relationships, and unrelated parts remain source-owned bytes.
- `definitions`, `trace`, `traceView`, and `transform` are currently bounded
  opaque payload boundaries in the shared profile. A host edit must retain
  them exactly unless a future typed operation explicitly changes that
  subtree. It must not imply full InkML schema validation or trace rendering.
- The target part must stay attached to the selected owner. If a future
  fixture demonstrates that action and InkML data share a part or have
  cross-part references, the transaction must model that closure and refuse
  unsafe deletion rather than silently creating an orphan.

## Detached action construction is the first implementation batch

The host cannot provide safe action CRUD around an immutable replay-only
profile. The first production batch should therefore complete detached action
construction and mutation in `litchi-drawingml`, before PPTX package wiring:

- Add a bounded `actions::Draft`/`actions::Prepared` family mirroring the existing
  detached InkML `Draft`/`Prepared` contract. The draft owns root units and
  ordered direct `Action`/`ActionGroup` entries. Action drafts own typed
  `ActionType`, checked `xsd:decimal` lexical `startTime`, optional `xml:id`,
  ordered properties, and ordered data children.
- Represent `ActionData` with the optional first `transform`, then ordered
  `trace`/`traceView` opaque payloads; represent data/action groups with their
  required nonempty children. Keep custom names and values losslessly bounded.
  A builder must reject impossible ordering/cardinality at the operation that
  creates it, rather than constructing an invalid intermediate value.
- Separate structural validity from semantic action conventions. A structural
  profile can accept custom action types and names while checking the §2.21
  sequence. A strict semantic profile may additionally enforce the reserved
  `add`/`remove`/`transform` data-name and count rules and prove `ref` closure.
  The API and diagnostics must say which profile was requested; accepting an
  opaque trace payload is not evidence that its InkML contents are valid.
- Finish serialization with preflight byte/resource checks, fallible output
  allocation, the chosen namespace policy, and an immediate `read_profile`
  readback. Publish a `Prepared` value only when the requested typed projection
  equals the readback projection. Source-backed edits may carry untouched
  opaque spans from the original source; fresh builders should require an
  explicit bounded opaque payload value.
- Provide transaction-scoped `set`, `add`, `move_before`, `clear`, and
  `remove` operations with selectors based on semantic order or stable
  detached identity. Do not expose a catalog-string-only or relationship-id
  mutation mode. Identity-changing edits must update any modeled `xml:id`/ref
  closure or return a typed refusal.

This batch should add detached tests for namespace aliases, declaration and
XML-character rules, property empty-content rejection, element-only whitespace
references, duplicate attributes, action/data ordering, one-over limits,
custom values, opaque payload retention, canonical write/readback, no-op
source replay, and inverse edits. It still must not claim action execution,
recognition, rendering, or native PowerPoint acceptance.

## Resource and performance requirements

The shared profile currently bounds one source at 16 MiB, XML depth at 128,
nodes at 100,000, attributes per element at 256, scalar attribute values at
1 MiB, actions at 65,536, and action groups at 16,384. The PPTX owner must add
package and slide aggregate budgets rather than resetting those limits for
each alias or anchor. The existing InkML host limits provide a reasonable
starting point: 32 MiB slide XML, 16 MiB per target, 256 MiB aggregate target
bytes, and 4,096 content-part anchors, subject to a named action-specific
policy and one-time charging of shared targets.

The implementation should preflight raw part and slide sizes before copying or
reserving, use fallible vectors/output buffers, and parse only selected action
parts when the caller requests them. A source-backed handle must retain its
cache reservation while its action bytes or opaque spans are live. No latency,
allocation, RSS, or throughput claim follows from these bounds; any such claim
requires representative measurements under ADR 0005
(`docs/adr/0005-io-memory-and-performance.md`).

## Fixture and acceptance plan

The existing `pptx_ink_annotations.rs` fixture is useful for the generic
InkML path only. It proves an `inkml:ink` part at `/ppt/ink/ink1.xml` with a
`customXml` relationship and `application/inkml+xml`; it contains no
`iact:actions` payload and cannot establish the action package contract.

The next evidence artifact should be a preserved native or independently
validated package containing, at minimum:

1. a slide `spTree` or `grpSp` with the active `AlternateContent` choice and
   `contentPart`, plus a real `pic` fallback;
2. both required choice namespace URIs, arbitrary prefix spellings, and any
   `mc:Ignorable` declarations used by the producer;
3. the owning slide `.rels`, `[Content_Types].xml`, target part name, declared
   content type, and relationship target mode;
4. an `iact:actions` root with a declaration, units, direct and grouped
   actions, properties, optional definitions, and at least one opaque
   trace/traceView/transform boundary; and
5. malformed variants or controlled mutations for wrong root, wrong type,
   missing relationship, duplicate `r:id`, inactive-only choice, external
   target, orphaned target, and stale-source publication.

Record producer/version, acquisition path, package SHA-256, and every member
used to establish the relationship/content-type/location rule. The synthetic
detached action profile tests can establish XML grammar and readback but are
not package-owner or native-interoperability evidence. Package acceptance
should cover exact no-op and inverse bytes, active/inactive MCE preservation,
namespace aliases, relation/content-type/root mismatch refusals, source
retention, bounded aggregate limits, and selected action projection after
reopen.

## Explicit remaining nonclaims

The shared profile and this design note do not claim a PowerPoint InkAction
part content type, relationship URI, or canonical package path until evidence
binds those values. They do not claim that `application/inkml+xml` may contain
an `iact:actions` root merely because `[MS-ODRAWXML]` names an Ink Content Part.
They do not claim full InkML trace/context/brush validation, action execution,
recognition, rendering, animation, or native Office acceptance. The audit row
therefore remains open for host ownership, detached action CRUD, semantic
reference validation, and producer evidence.
