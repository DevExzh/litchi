# ADR 0019: Typed DOCX web-settings ownership

- Status: Accepted
- Date: 2026-08-03

## Context

The OOXML migration host owned more than three thousand lines of
WordprocessingML web-settings parsing, authoring, package discovery, and
mutation. It also embedded a second producer template and exposed long
host-specific names. The concrete `litchi-docx` crate therefore did not own a
complete format capability, and callers had to manipulate a cached mutable
model whose later save could rewrite the part.

Web settings contain several independent option families, recursive framesets,
HTML division metadata, and frame relationships. Raw relationship identifiers
are necessary for a focused low-level frame link but are not suitable selectors
for ordinary division CRUD. Strict and Transitional packages also use distinct
namespace and relationship families that must not be mixed.

## Decision

`litchi-docx::web` is the sole owner of the bounded web-settings XML grammar,
semantic model, and OPC graph service. Its contextual vocabulary is
`Settings`, `Conformance`, `Key`, `Id`, `Twips`, `Div`, `Borders`, `Border`,
`Frameset`, `Child`, `Frame`, `SplitBar`, `Color`, `Layout`, `Scrollbar`, and
`Screen`.

The shared Word color theme vocabulary lives at `litchi_docx::color::Theme`,
so paragraph and web settings no longer duplicate that enum.

`Settings::get`, `add`, `put`, `remove`, and `move_to` make a division's
producer-visible ID the primary selector. `Key::Index` remains a checked raw
source-order selector for repair and import. Nested divisions use the same
policy. Missing semantic lookups return `Ok(None)`, ambiguity and invalid
positions are typed errors, and no selector indexes or unwinds. Recursive
frameset authoring and collection growth reserve fallibly before publication.
`Div::new` requires a nonzero numeric `Id` and installs all four required
signed-twips margins. The reader accepts every schema-valid `OnOff` spelling,
while the writer emits explicit numeric values for the `blockQuote` and
`bodyDiv` role markers because desktop Word rejects their otherwise valid
empty true form.

The short `Frame::rel` and `set_rel` methods deliberately expose only the inert
relationship token stored by Word; the ordinary model does not fetch or
activate its target.

The physical verbs are `load`, consuming `put`, and `remove`. They validate the
document owner, exact relationship multiplicity, internal target, content
type, Strict/Transitional agreement, bounded XML, and frame relationships.
Replacing or deleting a part with another inbound owner is refused. A semantic
and conformance no-op retains the producer's exact part bytes and package
signatures. A real edit emits deterministic modeled XML only after a semantic
round trip succeeds; only an actual commit invalidates signatures.

New DOCX packages obtain their default web-settings bytes from
`litchi_docx::web::write`. The migration host's parser, dirty cache, writer,
template accessor, and duplicate source/generated resources are removed. Its
remaining adapter exposes `Package::{web, put_web, remove_web}` and
`Document::web`, and re-exports the canonical `web` and `color` modules rather
than defining compatibility aliases.

## Consequences

- Web-settings create, read, update, remove, absence, no-op, and nested
  semantic CRUD now have one concrete owner and concise checked entry points.
- Loaded settings are independent owned values. `put_web` moves the completed
  value into the package boundary; the public facade exposes neither a lock
  wrapper nor unchecked numeric relationship IDs.
- An unchanged or semantic no-op preserves exact source bytes. A real modeled
  edit canonicalizes the supported grammar and can discard ignored or unknown
  extension markup. Source-surgical preservation across such an edit remains
  explicit follow-up work.
- Frame relationship CRUD beyond retaining or replacing the inert token is not
  yet a high-level semantic link API. The migration host still owns the wider
  DOCX package and document model, so its dependency on `litchi-docx` remains
  recorded migration debt.
- Border style is a bounded checked token rather than an exhaustive enum. This
  safely retains the schema vocabulary without pretending the current facade
  provides effect-specific semantic behavior for every border style.
- Compact ownership and removal of duplicate code are structural facts, not
  evidence of lower latency, allocation, or cache pressure. Those claims
  require representative measurements.

## Verification

Owner tests cover semantic and checked numeric division CRUD, recursive
framesets, Strict and Transitional serialization, exact no-ops, graph create/
update/remove, shared inbound refusal, bounds, malformed and nonempty scalar
elements, failure atomicity, and the shared color vocabulary. Host tests cover
the canonical producer bytes, package/document facades, relationship retention,
Strict graph handling, body-edit byte preservation, duplicate/external graph
refusal, a real POI fixture, and exact producer-asset parity.

Producer web-settings parts in the repository's Office packages always give a
division a border and spell `bodyDiv` as `w:val="1"`; borders nevertheless
remain optional in the model.

The host heading writer emits the style IDs `Heading1` through `Heading9`
(not the display name `Heading 1`) and installs every accepted built-in
heading style; each carries its structural `w:outlineLvl` (zero through
eight). A save/reopen regression checks `Title` and all nine heading IDs.
The `docx_owner_native_smoke` example produces an artifact with scalar web
settings, a body division, a block-quote division, and a Heading 1 for native
Word open-and-inspect checks.
