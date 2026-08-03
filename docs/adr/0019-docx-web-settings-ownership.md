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
refusal, and a real POI fixture. The complete owner gate passes 43 unit tests,
two public API tests, and one doctest. Focused host gates pass two owner
integrations, seven legacy-surface replacements, and four shared-color/
underline regressions. Exact producer-asset parity also passes. Warning-denied
Clippy and rustdoc are green for both owner and focused host, together with
formatting, diff, panic-name, stale-name, and crate-boundary checks.

The native gate first separated the generated part into baseline, scalar-only,
plain division, body-division, and block-quote artifacts. Word opened the
baseline, scalar-only, bordered plain division, and borderless plain division
without repair. It requested recovery for both `<w:bodyDiv/>` and
`<w:blockQuote/>`. A scan of 60 checked-in Office packages found 54 web-settings
parts, nine division-bearing parts, and 138 divisions: all 138 had a border,
129 had `bodyDiv`, and all 129 producer markers used `w:val="1"`; no producer
`blockQuote` example was present. The independent borderless artifact proves
that borders remain optional despite that producer convention.

That same native run exposed a host writer defect: `add_heading(_, 1)` used the
display name `Heading 1` as a style ID, and levels four through nine were
accepted without matching default catalog entries. The writer now emits
`Heading1` through `Heading9` and installs every accepted built-in style. A
closed `Outline::{H1, ..., H9}` type supplies the required wire levels zero
through eight, and the style writer emits `w:outlineLvl`. A save/reopen
regression proves that `Title` plus all nine heading IDs resolve with the exact
structural outline levels in the saved style catalog.

After those corrections, `owner_native_smoke` generated a document containing
scalar web settings, a body division, a block-quote division, a Heading 1, and
ordinary text. Desktop Word for macOS opened `web-settings-owner.docx` without
a recovery prompt, rendered both paragraphs, and selected the native Heading 1
style for the heading. This is open-and-inspect evidence for that Transitional
artifact only. It does not establish Office edit/resave preservation, a Strict
native round trip, other Word versions, source-surgical extension editing, or
measured performance.
