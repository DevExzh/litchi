# Change 0301: PPT selected-slide decoding

Status: implemented

`performance_claim: none`

## Scope

Native PPT positional reads now have an additive `Presentation::slide_at`
operation. It resolves one logical directory entry and parses only that slide;
the existing `Presentation::slides` API remains unchanged.

`Presentation::text` processes directory entries one at a time and appends
each selected slide's existing `Slide::text` result directly to the aggregate
string. Its ordering, separators, fallback text, and error filtering remain
unchanged.

The facade's native PPT positional operation uses this bounded path instead of
building the complete temporary `Vec<Slide>` before selecting an item.

## Evidence boundary

This change claims only removal of the temporary all-slide collection for a
selected query, avoidance of unrelated slide-record parsing for that query,
and one-slide-at-a-time aggregate processing. It makes no claim about total
RSS, peak memory, allocator traffic, I/O, latency, throughput, parser grammar,
or strictness.

The complete Document and Current User streams, parser, persist mapping,
slide directory, and NotesIndex behavior remain outside this change. This is
not a source-backed or freshness/cancellation change, and it does not alter
notes/master handling, editing, or saving.
