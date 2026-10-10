# One PPTX slide publication across bounded chunks

Records 0156–0157 measured nonlinear homogeneous command batching because each
bounded native setter call reparsed and republished the entire slide. The native
primary eager and source edits now accept `set_shape_text_chunks` and
`set_shape_property_chunks`; focused source slide edits expose the same
operations without a slide selector. The facade submits each contiguous
same-slide shape run through these operations, including alternating text-only
and geometry updates.

Each input chunk retains the existing 256-target and aggregate text-byte
limits. Chunk count and total selection count are bounded by the default shape
reader's 100,000-shape ceiling; total replacement text is also bounded by the
output-byte budget. Admission precedes descriptor flattening. All targets are
resolved against one immutable slide state, and duplicate or overlapping
selections across chunks are refused. Text/geometry validation, lexical
preservation and semantic readback use the shared native codecs. One candidate
is published after the complete operation succeeds. Existing single-chunk
setters and their limits remain available.

The facade preserves command-level results, no-op reporting, barriers and
wire-order failure attribution. New benchmark routes ending in
`chunk_publication_v2` compare the bulk native operation against one atomic
facade batch containing a separate command per shape. The prior
`command_batch_v1` harness remains unchanged.

## Verification

The native PPTX/encryption suite passed 1,688 tests/doctests. The combined
DOCX/XLSX/PPTX/encryption facade suite passed 1,108. New native coverage checks
1,000-shape byte equivalence with sequential bounded setters, 10,000-shape
edits, source immutability, eager/source/focused exact no-ops, cross-chunk
duplicates and invalid payloads, empty chunk admission, per-chunk and aggregate
bounds, managed reservation refusal/retention, and transported replay/inverse.
Two roughly 4 MiB text values test aggregate text beyond one chunk's 8 MiB
budget, for both text and combined property operations. The facade regression
checks all 1,000 alternating text/geometry results in one parsed output scene.
Existing wire-order error and transaction-barrier tests remain green.

Native/facade and standalone harness all-target Clippy passed with warnings
denied. The full PPTX glossary profile passed 70 vocabulary bans. Repository
boundaries passed for 66 workspace packages and 268 dependency declarations.

## Measurement

Paired measurements are pending.
