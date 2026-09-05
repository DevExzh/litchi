# 0425 baseline-reproduced test corrections review

This is a read-only review of the three test-only corrections against baseline
`340cc91ae`. The locked baseline receipts are
`checks/baseline-odp-locked.{log,json}`,
`checks/baseline-odt-locked.{log,json}`, and
`checks/baseline-xlsb-locked.{log,json}`. No Cargo command, test, build,
profiler, or CPU workload was run by this reviewer.

## Findings

No blocker was found. Each edit addresses the recorded baseline failure in the
test oracle or fixture construction; none changes the corresponding
production implementation.

### ODP limit-before-read fixture

The locked ODP failure was in
`source_facade_rejects_oversized_content_before_materialization`: the helper
first changed the central `content.xml` uncompressed size to 256 MiB plus one,
then called `member_range`, so the ZIP reader correctly rejected the malformed
central size before the source-backed limit assertion could run. The corrected
`oversized_content_package` in
`crates/litchi-odp/tests/source_backed.rs:214-250` obtains the valid archive's
compressed payload range before mutating that central size and returns the
range alongside the malformed bytes. The test at lines 429-445 then checks the
family-limit error and verifies that no observed source range overlaps the
content payload.

This preserves the intended proof: the source can inspect the archive metadata
and reject the declared content size before reading the compressed payload.
The range is taken from the unmodified archive, so the oracle does not need to
parse the intentionally invalid central metadata. The encrypted oversized
case remains independently covered by the existing test and is unchanged.

### ODT descriptor refusal and atomicity

The locked ODT failure showed that the former test's “size fallback” expectation
was incompatible with the actual preservation boundary: changing either the
compressed-size or uncompressed-size field at descriptor offsets 8 or 12
caused `Pictures/deflated.bin` to be rejected. The corrected test at
`crates/litchi-odt/tests/generic_content_publication.rs:605-621` treats all
three descriptor mutations—CRC at offset 4 and the two sizes at offsets 8 and
12—as malformed input. For each mutation it captures `document.to_bytes()`
before the attempted definition change, requires a refusal naming the affected
member, and checks the bytes are identical afterward.

`corrupt_deflated_descriptor` still derives the descriptor from the preserved
local span, asserts the descriptor signature, and flips exactly one field byte
(`generic_content_publication.rs:368-386`). This keeps the fixture tied to the
16-byte signed data-descriptor layout and avoids accidentally testing a random
member. The revised assertions now cover malformed CRC and both malformed size
fields with the same refusal and no-partial-publication contract. The separate
payload-corruption test remains unchanged and continues to exercise changed
payload verification.

### XLSB chart-sheet census and worksheet transfers

The locked XLSB failure was the old loop's single index-space assumption: it
iterated `0..worksheet_count()` but passed each worksheet ordinal directly to
the full workbook-catalog `sheet_drawing` lookup. The first corpus fixture has
one worksheet at catalog position 0 and one chart sheet at catalog position 1;
its drawing anchor is therefore not reached by the old loop. The independent
raw fixture census and diagnosis are recorded in
`checks/xlsb-failure-analysis.md`.

The corrected test at
`crates/litchi-xlsb/tests/workbook_structure_edit.rs:784-902` keeps the spaces
separate. Its four fixture rows map worksheet ordinals to catalog positions and
assert per-fixture worksheet counts, parsed drawing-anchor counts, and transfer
counts. It explicitly checks the chart sheet at catalog position 1 and its one
anchor. The totals are six parsed anchors (including the chart-sheet anchor)
and five supported worksheet-transfer anchors. Each worksheet anchor still
calls `transfer_drawing_object` with the worksheet ordinal, publishes the patch,
saves, reopens, and requires exactly one target anchor. Thus the test does not
expand the transfer API to chart sheets while it does retain the complete
six-anchor parser census.

The mapping agrees with the public API split: `sheet_drawing` selects a full
catalog position, while `transfer_drawing_object` accepts a worksheet ordinal
and resolves it internally. The correction therefore avoids both the original
false omission and an invalid attempt to transfer a chart-sheet drawing
through the worksheet-only operation.

## Disposition

The three corrections are coherent test-oracle/fixture repairs. The ODP test
retains the limit-before-payload-read assertion, the ODT test now covers all
malformed descriptor fields with exact atomicity, and the XLSB test preserves
the six-anchor parsed/raw census plus five supported transfer/reopen cases.
No production change is indicated by these baseline failures.
