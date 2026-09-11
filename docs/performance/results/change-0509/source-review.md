# 0509 ODT sink buffer reuse source review

Reviewed 2026-09-11 against the frozen `text.rs` production diff, its four
private `PendingSinkBlocks` tests, and the focused `sequential_text.rs`
integration additions. This is a read-only correctness review. No build or
test was run for this review, and no Rust source was changed.

## Verdict

I found no concrete correctness blocker for the bounded reuse candidate. The
change keeps one operation-local `Option<String>` spare in
`PendingSinkBlocks`. It does not change the public API, parser event handling,
frontier slots, writer call boundary, or source revision checks.

## Resource and allocation invariants

`SinkTextBudget` is created once before the sink parser's event loop and is
never moved into an active block or reset when `take_reusable()` supplies a
buffer. Every visible text contribution still charges that same monotone
budget before append:

* `Text` and `CData` events compute the normalized UTF-8 length, charge it,
  then materialize and append through `append_sink_precharged`;
* `text:s`, `text:tab`, and `text:line-break` charge their decoded spaces or
  control byte; and
* character and named references charge the decoded value before appending.

`append_sink_precharged` still performs the fallible capacity reservation and
does not alter the budget. Reusing an existing allocation therefore cannot
make the cumulative `MAX_TEXT_BYTES` ceiling per-block or capacity-based.

`recycle` reads `String::capacity()` before clearing the value and drops the
value when that actual capacity exceeds `MAX_SINK_TEXT_REUSE_CAPACITY` (4,096
bytes). A short string with a large allocation is consequently not retained.
For accepted values it clears the string and keeps the largest available
spare; it performs no fallible allocation. The spare belongs to one parser
operation and is dropped with that `PendingSinkBlocks`, so there is no global
or cross-document retention pool. Active blocks may still grow beyond 4,096
bytes while being parsed, but such a value is discarded after emission.

The private tests
`pending_sink_reuses_successfully_emitted_text_after_clearing` and
`pending_sink_drops_short_value_with_oversized_capacity` directly cover empty
reuse and the capacity-versus-length distinction. The latter constructs a
one-byte value with capacity 4,097 and verifies that no spare remains.

## Semantic order and publication boundary

`complete` still stores the finished value in its assigned slot and emits only
the contiguous front of the start-order queue. It calls
`writer.write_object(TextObjectKind::Paragraph, &value)` before incrementing
`next_emit` or recycling the value. A nested block that completes while its
outer block is pending therefore cannot become a spare early or alter output
order. Once the outer slot completes, both values are emitted in their
existing order; the smaller nested spare cannot replace the larger outer
spare. `pending_sink_keeps_larger_spare_after_nested_completion` checks this
with outer slot 0, nested slot 1, and the expected `outer\n` output.

Normal `Start` blocks take the spare after their slot and attribute validation;
empty blocks continue to use `String::new()`. This leaves the existing empty
object policy and all nested depth/suppression handling unchanged. The new
integration case in
[`sequential_text.rs`](../../../../crates/litchi-odt/tests/sequential_text.rs#L75)
compares exact owned and source-backed output across a short block, an empty
block, nested content, a 4,097-byte block, and a following `tail` block. Its
expected separators and six object count make stale contents and start-order
publication observable.

## Errors and progress

If `write_object` returns a sink, output-byte, or object-limit error, `?`
returns before `next_emit` changes or the failed value is recycled. A
successful writer call remains the only path to recycling. The existing
`SequentialTextWriter` therefore continues to report accepted bytes and
completed objects at the same boundaries, including partial sink writes.

The private
`pending_sink_does_not_recycle_after_writer_failure` test checks that a
failing writer leaves the spare empty and `next_emit` unchanged. The new
integration test
`sink_failure_after_emitted_block_preserves_progress_and_text` forces a
successful `one`, separator, and one byte of the next value, then asserts the
exact `one|t` sink contents, five accepted bytes, one completed object, and a
typed sink error. Existing integration coverage continues to assert object
and output-byte limits after the first paragraph, malformed input after a
valid emitted block, and source-change precedence with truthful prior
progress.

No new operation can fail between a successful writer call and `recycle`:
`String::clear`, capacity inspection, and `Option` replacement are
infallible. The checked `next_emit` increment remains before recycling, as in
the prior frontier implementation; its overflow is unreachable while
`MAX_TEXT_BLOCKS` bounds reservations.

## Coverage observations

The frozen tests directly cover successful reuse, actual-capacity rejection,
nested completion deferral, writer failure, exact owned/source-backed bytes,
empty blocks, oversized content, output/object limits, malformed tails, and
source staleness. I found no missing test that exposes a semantic or progress
regression in this candidate.

There is no dedicated test that drives the internal 64 MiB decoded-text
ceiling across multiple blocks while a reused spare is present. Static review
shows the required property: the single `SinkTextBudget` remains outside the
active block and every append path charges it. This is a modest evidence gap,
not a source-level blocker; adding such a fixture would mainly duplicate the
existing budget checks while making the test comparatively large.

The 4,096-byte equality boundary is also not exercised as a separate fixture,
but the implementation checks actual capacity with `>` and the tests cover
both retained capacity below the bound and a one-byte value above it. The
admitted profile and retention record support this bounded candidate only;
they do not establish a general peak-memory or latency claim.

