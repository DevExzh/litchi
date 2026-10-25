# 0362ul81 — PPTX structural metadata policies

- Date: 2026-10-25.
- Status: source-ready for independent review; native Cargo validation remains
  with the coordinator's isolated build slot.
- Base: `fda2b5d223452548d92201ec4945af9e4d5835a5`.

This slice completes the explicit policy migration for the eight PPTX
structural metadata families: custom shows, guides, sections, presentation
structure, color maps, handout masters, templates and protection.  Their raw
parsers, package loaders, source snapshots and writers now require the shared
`presentation_properties::metadata::Limits` value.  Package-aware operations
derive that value from the retained `SourceLimits`; detached callers select it
explicitly.  There are no default-only twins, deprecated aliases or public
`_with_limits` spellings.

The shared policy bounds input and output bytes, XML depth and nodes, decoded
strings and attributes, extension and family counts, slide references and
relationships.  MCE processing is bound to the selected policy.  Writers and
source splices check the exact final size before publication.  Source-preserving
families retain opaque XML, strict/transitional namespaces, relationship
mappings, custom-show IDs and section GUIDs.  Detached handout and protection
models intentionally emit their documented canonical typed forms; unknown
markup is outside those detached models and is not claimed as opaque round-trip
fidelity.

The package edit families keep the `Edit` → `commit` shared triad with
`litchi_core::edit::{Commit, Diagnostics}` and typed `LimitKind`.  Dropping an
edit discards staged changes.  Replay compares equal content and relationship
topology rather than patch lineage.  A no-op retains the recipient's policy
and exact bytes; a changed payload is validated under the recipient policy
before publication.  No new full-owner rescan or proportional clone is
introduced on the no-op path.

Tracked consumers in tests and examples were migrated to required limits,
including package structure readers, handout/protection examples, and raw
color-map, guide, section and custom-show fixtures.  The minimal and full
PPTX glossary profiles register the required-limit signatures, and the living
API glossary records the eight-family contract.

The common 61 physical source scan and no-op base check remain separate
integration returns; this slice does not claim that composition closed.

Source-only checks for this record are `rustfmt --check`, `git diff --check`,
glossary JSON parsing and focused source/API sweeps.  Cargo tests, rustdoc,
clippy and ABBA measurements are intentionally outstanding for the
coordinator's independent run; no binary or large performance evidence is
committed.
