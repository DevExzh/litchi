# ADR 0021: Typed DOCX glossary ownership

- Status: Accepted
- Date: 2026-08-03

## Context

The OOXML migration host owned the complete WordprocessingML glossary-document
grammar, building-block CRUD, and auxiliary OPC graph. The public model exposed
long format-prefixed names, mutable field bags, numeric positions as mutation
keys, raw relationship and part identifiers beside ordinary semantic values,
and graph replacement that could delete a glossary-owned part without proving
exclusive inbound ownership. Whole-catalog cloning made every entry edit scale
with unrelated building-block bodies.

Glossary entries are developer-facing objects identified by their
producer-visible names. Their bodies and auxiliary resources can be large and
must remain inert: reading or copying a building block does not insert it into a
document, resolve fields, load fonts, decode media, activate hyperlinks, or
execute embedded content. Physical relationship identifiers remain necessary
for lossless low-level graph work but are not suitable ordinary CRUD selectors.

The checked-in `[MS-OI29500]` sections 2.1.310 through 2.1.312 and
`[MS-OE376]` sections 2.1.314 through 2.1.317 record Word-specific glossary
behavior. In particular, the latter scopes extra authored-entry properties,
name, and GUID-value requirements to Word 2007, Word 2007 SP1, and Word 2007
SP2. The former records last-occurrence behavior and permits an empty
`docParts` list. Readers must retain valid producer states without making
incomplete fresh authoring easy.

## Decision

`litchi-docx::glossary` is the sole owner of the bounded glossary XML grammar,
typed catalog, and package graph service. Its ordinary vocabulary uses short
contextual names: `Catalog`, `Entry`, `Props`, `Name`, `Category`, `Gallery`,
`Id`, `Kind`, `Insert`, and `Conformance`. Physical graph replacement is a
separate low-level `glossary::raw` capability whose `Graph`, `Part`, and `Rel`
types do not leak into ordinary catalog CRUD.

Exact producer-visible names are the primary selectors. One deterministic,
bounded NFD/default-case-fold/NFD identity is computed once per name and kept
behind a private catalog index shared by lookup, add, put, replace, rename,
remove, and reorder validation. A producer catalog may contain duplicate
display names; those entries remain inspectable and make semantic lookup
explicitly ambiguous rather than making the whole catalog unreadable. New
conflicting names are rejected. Checked source-order positions remain available through `at`,
`replace_at`, `rename_at`, `remove_at`, and `move_at` for repair and import. A
missing semantic name is `Ok(None)` for lookup and removal, while name-first
reorder reports `Ok(false)`; ambiguous input and invalid positions are typed
errors. No public `Index`
implementation can panic, and ordinary selectors never expose a native
relationship ID.

`Entry::new` requires a checked name and validated inert `docPartBody` payload,
so the Word 2007 authored-entry requirements are represented at construction.
The reader remains capable of retaining valid producer states permitted by the
base standard: `docPartPr` and its name may be absent, and `<guid>` may be
present without `guid/@w:val`. The reader also accepts Word's native
present-but-empty `<w:types/>` producer state; typed fresh authoring cannot emit
that state. Fresh semantic authoring still requires a name.
`docPartPr` children are accepted in any schema-valid order and typed projection
follows Word's last-occurrence behavior. Untouched direct producer entries retain
their bounded original subtree, including inactive or ignorable MCE content,
when an unrelated catalog entry changes; a deliberately changed entry is
serialized from its typed active projection. `Id` validates the building-block
GUID and writes its canonical braced uppercase form. Independent kinds and
insertion behaviors are compact typed bitflags; other optional properties remain
grouped under `Props`. Entry payloads move across catalog edits; raw resource
payloads share immutable ownership while package publication borrows the
recovery graph.
Edits that add, replace, or rename content reserve and validate the candidate
before publication instead of cloning unrelated entry bodies. Rename changes
the checked `Name` in place and invalidates only that entry's producer snapshot.
Checked Strict and Transitional serialized sizes are cached on the private
entry state and in catalog totals, then updated by deltas, so repeated CRUD does
not reparse or replan unrelated entries. Removal, clearing, and reorder publish
only after their checked preconditions pass.

The package verbs are `load`, consuming `put`, and `remove`. A package-aware
load privately binds relationship-bearing catalog XML to its complete validated
physical graph. Each relationship-bearing entry and background subtree carries
the same unforgeable private lineage token. Semantic insertion, replacement,
background editing, and publication compare that token and reject missing or
foreign lineage before mutation, preventing equal `r:id` spellings from
silently selecting unrelated destination resources. Editing a body clears its
lineage; explicit cross-package resource transfer uses `glossary::raw`.

The package verbs validate one internal main-document glossary relationship,
matching Strict or Transitional XML and relationship families, the required
content type, every internal target, auxiliary-part bounds, and package-wide
inbound ownership before a change. Relationship validation is role-driven:
the role assigned by the incoming edge determines the legal outgoing edge,
target mode, target content profile, and next role. An arbitrary embedded
payload therefore cannot gain chart permissions merely by selecting a chart
content type. Internal hyperlinks are references, not glossary-owned
dependencies, and removal never deletes their targets. A physical
`/word/glossary/` directory is only a producer convention and never establishes
ownership; ownership comes from the validated typed relationship closure.
Generic control payloads remain inert and may be retained, while only a control
with the exact ActiveX descriptor content type may own an ActiveX binary. The
matrix includes settings recipient data, chart theme overrides, both 2011 and
2012 chart-style relationship families, the chart/chart-drawing cycle,
chart-drawing Custom XML, and diagram hyperlinks.
Package-root, other-owner, duplicate, invalid-mode, orphan-glossary-content-type,
mixed-dialect, wrong-content-type, shared, and dangling graphs are rejected
before mutation.
Strict content also rejects Transitional relationship namespaces and lexical
forms plus active VML.
An unchanged bound catalog loaded from the destination, or a canonical/exact raw
no-op, returns before reserialization and retains the producer's original part
names, bytes, and signatures. Shared raw payload identity is checked before byte
comparison. A real change first prepares its standalone graph data, then clones
the package and invalidates signatures on that candidate before package mutation
and round-trip validation. Only the final assignment publishes the candidate to
the original.

Low-level removal moves out the complete `raw::Graph`, including auxiliary
payloads, root-owner metadata, and relationship metadata; it never reports a
catalog-only return value after destroying the rest of the graph. Low-level
publication borrows that graph, so validation or destination failure cannot
consume the caller's only recovery copy. Cross-package publication rebases the
owner target against the destination main part and allocates a new relationship
ID on collision while preserving source metadata when it remains valid. Raw
parts use shared immutable payload allocations. Per-part and aggregate payload,
relationship-count, and metadata budgets are checked before publication. Node,
attribute, content-token, depth, namespace, DOM-allocation, projected-XML, and
producer-snapshot budgets apply to glossary XML and opaque semantic XML
subtrees; auxiliary bytes remain inert and are bounded by size, graph topology,
and content profile. Producer snapshots are bounded serialized subtrees rather
than retained DOMs. They carry relationship references from active and inactive
MCE branches so raw identifiers cannot be rebound through projection. OPC
manifests, relationship parts, and signature infrastructure paths or content
types cannot masquerade as glossary-owned parts.

Opaque XML is re-emitted through namespace-aware nodes when conformance must be
changed. Namespace URI text and unrelated attribute values are never rewritten
by global string substitution. Namespace scopes use shared persistent frames,
so inherited declarations are not cloned into or emitted by every descendant;
each extracted opaque root receives the scope it needs and all extracted output
shares one aggregate ceiling. Parsing also has a streaming expected-constant-time
namespace resolver. XML 1.0 forbidden characters are rejected in parsed and
authored values, and carriage returns serialize as character references so they
remain distinct from line feeds. Strict publication rejects VML at the selected
dialect gate without making a valid Transitional catalog unreadable. MCE
preprocessing uses the same 32 MiB input/output boundary as the glossary codec.
External links and all auxiliary resources remain inert.

Fresh semantic publication seeds the four glossary-local resources used by
desktop Word templates: styles, settings, font table, and web settings. When a
matching main-document resource has no outgoing relationships, its immutable
payload is shared; otherwise a bounded minimal resource is authored. Canonical
part names are preferred, but the root and all four resources allocate checked
free names when unrelated package parts already occupy those paths. The raw
layer remains explicit and does not silently add resources. The host's
`Package::new_template()` creates the native DOTX container expected for
AutoText authoring; ordinary `Package::new()` remains a DOCX document.

The OOXML migration host exposes short package/document adapters and the
canonical module itself as a contextual re-export. It deletes its duplicate
implementation and long compatibility exports.

## Consequences

- Building-block create, read, update, delete, clear, reorder, and complete
  graph replacement have one concrete owner and one selector policy.
- Ordinary callers work with names or checked positions; only focused low-level
  code sees physical OPC graph metadata.
- Move-first entry edits avoid copies proportional to unrelated bodies. Exact
  package no-ops preserve producer allocations and signatures; no public lock
  or async-runtime type is introduced.
- Glossary bodies remain validated opaque XML; auxiliary formats remain bounded
  inert content. Untouched direct producer entries retain unsupported content
  across unrelated CRUD; a targeted rewrite of that same entry retains only
  modeled active semantics.
  This capability does not insert building blocks into document stories,
  render them, evaluate fields, or activate linked/embedded content.
- Ownership and reduced cloning are structural properties. Throughput, cache,
  latency, and allocation improvements require representative measurement and
  are not inferred from this refactor.

## Verification

Verification must cover Unicode-caseless semantic CRUD and checked positions;
Strict and Transitional parse/write; empty producer catalogs; Word 2007 authored
entry requirements; exact no-op/signature preservation; namespace-URI text;
all accepted auxiliary relationship families; shared inbound, orphan,
duplicate, external, dangling, mixed-dialect, wrong-content-type, collision,
limit, and allocation failures; atomic package create/update/remove; fixture
round trips; warning-denied Clippy and rustdoc; formatting; manifest sorting;
and executable crate-boundary policy.

Native verification on the tested desktop Microsoft Word for macOS build first
observed a DOCX without the later four-resource seed. Word opened it without
repair in Compatibility Mode, did not expose the custom entry, and removed its
glossary graph on resave. That negative result is consistent with templates
being the native AutoText authoring container, but the subsequent seeded DOTX
also changed the resource graph and therefore is not a controlled causal
comparison. The seeded DOTX opened without repair. Word's AutoText dialog
displayed the exact `Litchi AutoText` entry; inserting it placed `Litchi reusable
native building block` in the real document. Word saved the template, ZIP
integrity passed, and the canonical owner reverse-read the catalog entry and its
building-block payload from the Office-saved copy. That reverse read also fixed
the reader policy around Word's native `<w:types/>` rewrite. The Word UI
observation separately confirmed the inserted body text. This verifies one
Transitional text-only AutoText
open/discover/insert/resave path in the observed Word Compatibility Mode on that
build. Images, fields, linked or embedded resources, Strict documents, other
Office versions, and performance remain outside the observed native scope.
