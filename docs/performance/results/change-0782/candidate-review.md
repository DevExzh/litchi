# Independent candidate review

Review status: **bounded approval for the coordinator's quality and paired
measurement gates; no static correctness blocker found**. This review covers
the selected 0782 patch against base
`f1df64d15a8f2711e554f58acb413b06d56ef1ca` and the final formatted applied
archive. It does not approve production adoption or claim a performance
result.

## Review basis

I read `design.md`, `source-review.md`, and the frozen candidate
`implementation-notes.md`, including the common-workflow regression guard. I
also checked the applicable ADR constraints identified by the design: API
layers and crate topology (0001/0002), borrowed immutable views (0003), memory
and measured performance (0005), validation and deterministic output (0006),
migration gates (0008), and current topology (0024). The review was
read-only: no Cargo, compiler, formatter, native, profiler, or performance
command was run by this reviewer.

The archive has the expected twelve-file sorted allowlist. The selected patch
passes `git apply --check --whitespace=error`; every source/control hash in
the candidate archive receipt matches. The final `candidate/applied/` copy is
byte-identical to the twelve current formatted source files. Compared with the
selected draft, its only differences are rustfmt line wrapping in
`core/tests.rs`, `escher/semantic.rs`, and `escher/tests/groups.rs`.

## Ownership and lifetime findings

- The private `UserShapeData<'text>::text` field changes from
  `Option<String>` to `Option<&'text str>`. The public
  `ShapeProperties.text: Option<String>` owner is unchanged.
- Plain conversion receives `&'text WritableShape` and uses
  `props.text.as_deref()`. The resulting slice is tied to the source writer
  shape and remains valid through the synchronous serialization call.
- Both fresh package routes build `Vec<UserShapeData<'_>>` from slide
  references and consume it before the source borrow ends. The lifetime is
  carried through drawing, group, shape, property, and validation helpers.
- `ChildShape`, `GroupShape`, and `GroupChild` carry the same lifetime through
  nested group construction and encoding. These are crate-private types, so no
  public source generic or public API change is introduced.
- The production call graph has no non-empty plain `UserShapeData` value that
  needs to detach from its source. Notes/default construction remains
  lifetime-safe, and the manually authored test fixture now uses a static
  string. No unsafe code, cache, dependency, or global state is added.

## Mutation and output findings

- Rich paragraphs remain owned: conversion clones them before assigning PP9
  smart-tag run identifiers and building extension records.
- Centered and other non-left plain text still follows the existing owned
  `Paragraph::new(text.clone()).align(...)` fallback. The candidate does not
  borrow into a paragraph that can be mutated.
- The shape encoder passes the optional slice directly to the existing `&str`
  textbox writer. The text codec is unchanged, so the plain bytes and refusal
  ordering have no source-level change from this representation alone.
- The selected change removes an owned alternative rather than adding a
  fallback. That is safe for the inspected private production call graph, but
  complete quality gates must still confirm all crate-internal callers compile
  and preserve their behavior.

## Focused coverage

The candidate adds a pointer/length assertion for the borrowed source slice, a
target-layout assertion showing `Option<&str>` has slice layout and is smaller
than `Option<String>`, a centered fallback text assertion, and a grouped
borrowed-text test that compares the nested textbox payload with the existing
direct encoder and reads it back. The existing smart-tag test also checks that
rich conversion leaves source PP9 identifiers unchanged and exercises both
`write_to` and `save` output paths.

The layout assertion is evidence about the changed field representation on the
current target. It is not a whole-`UserShapeData` size claim or a causal
performance attribution. The unchanged complete PPT text golden suite remains
the required broader check for ASCII, Unicode, empty, record-boundary,
interaction, rich, long-text, notes, and both serialization routes.

## Adoption conditions

The design's guard remains binding: adoption requires a useful public-case
improvement, no new retained-memory cost, and no persistent control
regression. A paired p50 median above +5% whose 95% interval remains above
zero counts against adoption in any of the ten cases. Allocation reduction by
itself cannot justify a common-workflow slowdown. The coordinator must retain
all ten cases, tails, RSS, allocation metrics, spread, and paired flags before
making that decision. This static review does not establish any of those
runtime claims, native interoperability, or complete CRUD support.
