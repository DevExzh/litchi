# Independent candidate review

Review status: **bounded approval for build and gate validation; no static
correctness blocker found**. This review covers the frozen candidate against
`fdca3e63037ac83ff27d4bf64a0f157b971b45f8`. It does not approve production
adoption or claim a performance result.

## Review basis

I read the 0781 `design.md`, `source-review.md`, and
`implementation-notes.md`, together with ADR 0001 (API layers), ADR 0002
(crate topology), ADR 0003 (borrowed views and ownership), ADR 0004 (semantic
API conventions), ADR 0005 (memory and measurement), ADR 0006 (validation and
deterministic serialization), ADR 0008 (verification gates), ADR 0024 (current
topology), and the ADR index. The review was read-only: no candidate source
was applied and no Cargo, formatter, native, or profiler command was run.

The archive control evidence is intact. `git apply --check` and the strict
whitespace check both pass. The byte counts and SHA-256 values recorded in
`implementation-notes.md` match all twelve archived source files plus the two
candidate control artifacts.

## Ownership and lifetime findings

- `UserShapeData<'text>::text` is the only changed ownership boundary. The
  private type changes from `Option<String>` to `Option<Cow<'text, str>>` and
  retains its existing `Clone`/`Default` behavior.
- `convert_shape_to_escher_with_sound_mapping` receives
  `&'text WritableShape` and uses `Cow::Borrowed` only for the plain branch
  where no paragraph vector exists. The returned borrow is therefore tied to
  the source shape, not to a temporary conversion buffer.
- Both fresh package routes (`Writer::save` and `Writer::write_to`) construct
  the intermediate `Vec<UserShapeData<'_>>` from slide references and consume
  it synchronously while the source writer remains borrowed. The Escher
  drawing, shape, property, and validation helpers carry the lifetime through
  without storing it beyond the call.
- The grouped path propagates the same lifetime through `ChildShape`,
  `GroupShape`, and `GroupChild`, including nested groups and their codec and
  validation helpers. The types are crate-private, so this does not expose a
  source generic through the public writer API. Existing notes construction
  and manually authored shapes continue to use owned/default values.
- The Escher encoder consumes `Cow` through `as_ref()` and otherwise follows
  the existing text record writer. `Cow::Owned` therefore has the old output
  behavior, while the borrowed branch changes only allocation ownership.

## Mutation and semantic findings

- Rich paragraphs remain owned: the conversion clones the paragraph vector
  before assigning PP9 run identifiers and building smart-tag extension data.
  The focused smart-tag assertion checks that source runs retain unset PP9
  identifiers after conversion.
- Non-left alignment remains the existing owned fallback through
  `Paragraph::new(text.clone()).align(...)`; the candidate does not try to
  borrow text into a paragraph that may be changed.
- Plain text with no paragraphs, including empty text and text routed through
  the existing interaction encoder, reaches the same `&str` record encoder.
  No source mutation, global cache, unsafe code, dependency, or public API
  change is introduced.

## Focused coverage present in the archive

The candidate adds a source-pointer/length assertion for the plain borrow, an
alignment fallback assertion that checks paragraph text, a rich smart-tag
round-trip covering both `write_to` and `save`, and a grouped borrowed-text
encoding/readback case. The existing PPT text golden suite remains the needed
broader byte-equivalence coverage for ASCII, Unicode, empty, record-boundary,
rich, and long text; it is deliberately unchanged in this archive.

## Conditions before adoption

The candidate is ready for the coordinator to apply only after the corrected
Unicode oracle has frozen and the baseline qualification has completed. The
coordinator still needs to compile and format the applied source, run the
focused and complete PPT gates, compare both serialization routes for exact
output/refusal/semantic parity, and perform the planned paired native and
allocation measurements. A successful static review cannot establish any of
those runtime or performance claims. If any gate exposes a compiler, byte,
semantic, or lifetime discrepancy, this review should be superseded by the
resulting failure evidence rather than treated as approval.

## Applied-archive follow-up

After application and `rustfmt`, I compared all twelve files under
`candidate/applied/` with the reviewed `candidate/files/` archive. Eight are
byte-identical. The four differences are whitespace-only line wrapping in the
PP9 source assertion (`core/tests.rs`), `visit_group` (`codec/validation.rs`),
`push_shape` (`escher/semantic.rs`), and the grouped textbox binding
(`escher/tests/groups.rs`). There is no changed token or semantic divergence
from the reviewed candidate; the applied archive still contains exactly the
same twelve paths. This follow-up does not change the pre-existing conditions
for compilation, exact-output gates, or performance adoption.
