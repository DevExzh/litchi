# 0898 source review: emitted QName output spans

This review covers the two-file emitted-name candidate against the frozen
0898 source at `8d517a7861`. It is a source and focused-test review only. It
approves the candidate for the quality, differential, resource/error, and
matched-measurement qualification gates; it does not approve production
adoption or make a performance claim.

## Source finding

`write_start` now returns a `Range<usize>` for the qualified
name after the complete start tag succeeds. Both the source-copy path and the
rebuilt path record the offset immediately after writing `<`, and both call
`checked_range` after writing the final `>`. The rebuilt path writes
`q.as_bytes()` at that offset; the source-copy path uses the same validated
source spelling, so `q.len()` describes the recorded name bytes. Hoisted
declarations and filtered attributes are written after the name and do not
split the span.

The range is an offset into the append-only output, rather than a borrowed
slice. Later `Vec` growth can move the allocation without changing the offset.
`checked_range` uses checked addition and verifies the end against the current
output length. `extend_from_within` checks subtraction and the source end,
calls the existing fallible output reservation, and only then invokes the safe
`Vec` copy. No unsafe code, source lifetime escape, new public API, unchecked
slice, or omitted name/attribute validation was added.

The close path keeps the existing three operations and their limit order:
`</`, the stored name span, and `>`. This is present for both ordinary end
events and empty elements. The span is published only after the corresponding
start write succeeds, so a failed start cannot leave a frame with an invalid
closing name. The output remains append-only throughout the codec; the only
direct output-buffer accesses in the candidate are the span's internal offset
calculation and checked copy.

Active-state handling is explicit. An active opaque extension or emitted
element receives `Mode::Emit(range)`. An emitted element below an inactive
branch receives `Mode::Inactive`, avoiding both the old owned name and a close
operation. `Skip`, `Unwrap`, `Alt`, and `Branch` retain their control modes.
The `c.opaque` path computes its range only when its parent is active, so an
opaque extension below a dropped wrapper cannot leak bytes. Namespace
inheritance and hoisting continue to use the pre-existing `emitted` boolean,
and the range is taken only after any hoisted start-tag content is complete.

## Test review

`an_emitted_name_span_survives_output_reallocation` keeps the root name live
while 64 child empty elements normalize `&#34;` to `&quot;`, forcing output
growth. It checks the complete bytes, the exact output limit, and the one-byte
under-limit refusal. The empty-child closes exercise the new span copy while
the root close verifies that the original range survives `Vec`
capacity growth.

`inactive_and_opaque_modes_keep_their_exact_byte_branches` covers a preserved
element, an unwrapped element, a dropped element containing an opaque
extension, an active opaque extension, and an alternate-content fallback in one
document. The dropped wrapper's opaque subtree must remain silent, while the
active opaque subtree preserves its branch. Existing focused tests continue to
cover namespace hoisting, redeclaration shadowing, alternate-content branch
selection, empty elements, copied and rebuilt starts, duplicate names,
malformed names, and exact output bounds.

The source review found no candidate-specific correctness issue. The frozen
semantic differential and the root quality gates remain the authorities for
the complete corpus and build/test result; this review does not substitute for
those checks.

## Final source-review receipt

Review base: source selected from `506aa2edae6666d9c76ae6827c2f5941a8e47896`
with current evaluation head `8d517a7861`.

The reviewed candidate files are bound by these SHA-256 identities:

* `crates/litchi-ooxml-common/src/mce/codec.rs`:
  `89579ff2a159282a64b395c0c5e5131044e3aed718075c43fe3d1992aec4c003`
* `crates/litchi-ooxml-common/src/mce/tests.rs`:
  `ea9bc287274bcfba6eee8dbb6728e87fb20d57b526d78441d7126938d0aab166`

The frozen candidate source receipt records the same hashes in
`candidate-overlay/`, with `production_adopted: false`. Source-review status:
**approved for the 0898 qualification gates**. This is not production
adoption approval and does not claim a whole-producer speedup.
