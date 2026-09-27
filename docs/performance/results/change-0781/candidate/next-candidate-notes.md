# 0781 narrower follow-up: borrowed slice representation

This is a read-only audit of the frozen `Cow` candidate against
`fdca3e63037ac83ff27d4bf64a0f157b971b45f8`. It does not replace or amend the
measured candidate, its applied archive, or its measurement records. No
Cargo, compiler, formatter, native probe, or profiler command was run.

## What currently needs owned text

The internal `UserShapeData.text` field is read by the Escher shape encoder;
the current call sites do not need to mutate it or keep it after the
synchronous drawing write.

- The only production assignment of non-empty `UserShapeData.text` is the
  plain branch in `writer/core/codec.rs`. With the archived candidate it
  always creates `Cow::Borrowed` from `ShapeProperties.text`.
- `writer/escher/model.rs` supplies `None` in `Default`, and the notes writer
  constructs a default shape before assigning owned paragraph data. The notes
  path does not assign `UserShapeData.text`.
- The rich conversion path clones `ShapeProperties.paragraphs` because PP9
  smart-tag identifiers are assigned to the converted copy. That ownership is
  in `paragraphs`, independent of `UserShapeData.text`.
- The non-left alignment fallback creates an owned `Paragraph` with
  `Paragraph::new(text.clone())`; it needs that paragraph ownership and is not
  a reason for an owned plain-text field.
- The encoder only passes plain text to the existing `&str` textbox writer.
  It does not mutate, append to, or store the `UserShapeData.text` value.

The remaining direct `UserShapeData` text construction is in crate tests: the
legacy case uses a string literal and the grouped case already borrows from a
local source string. There is no current production caller that constructs a
dynamic owned `String` inside `UserShapeData` and then retains that value.
The public `ShapeProperties.text: Option<String>` remains the owner at the
writer API boundary.

## Concrete narrower candidate

The next representation can therefore be `Option<&'text str>` instead of
`Option<Cow<'text, str>>`. Keep the existing `'text` lifetime propagation
through `UserShapeData`, `ChildShape`, `GroupShape`, `GroupChild`, package
vectors, drawing helpers, validation, and the synchronous writer calls. The
conversion's plain branch would return `props.text.as_deref()` directly. The
shape encoder would pass the optional slice directly to the existing text
writer. Tests can use `Some("Test")` and `Some(source.as_str())` without an
owned variant.

This is coherent for the current internal construction graph because every
non-empty converted plain value already aliases a live `WritableShape`, and
the package consumes each intermediate vector before that source borrow ends.
If a future crate-private caller needs to build a detached Escher shape, it
would need an explicit owned intermediate or a separate owned model; the
current source audit found no such caller, so adding that fallback now would
broaden the change.

## Layout and branch hypothesis

The coordinator's paired result for the `Cow` candidate is mixed: payload
write/lifecycle improved by about 25.55%/19.19%, Unicode improved by about 8%,
while the many-shape short-text case regressed by about 12.12% write and
10.55% lifecycle consistently across the six blocks. Those observations do
not establish a cause. They do identify a concrete representation question
for a bounded follow-up.

`Cow<str>` carries an owned `String` alternative and a variant distinction
even though the current production conversion never selects that alternative.
An `Option<&str>` is a single optional slice representation. Removing the
unused owned alternative and its variant handling could reduce the
intermediate shape layout and the plain-path branch work, especially for
many short-text shapes, while retaining the large-text aliasing benefit.
The exact `size_of` and generated code must be checked after application; this
is a layout/branch hypothesis, not a causal attribution of the measured
regression or a predicted result.

If this follow-up is tried, preserve the current candidate's byte, semantic,
refusal, rich-text mutation, group, and source lifetime gates, then rerun the
same paired matrix. Do not infer adoption from the representation inspection
alone.
