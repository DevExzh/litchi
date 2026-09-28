# 0824 archived candidate: outer-whitespace compaction witness

This archive records a candidate against base `7b268927bfe3ed9c9000bfd567fc4ce656ec0cd9`. It is not applied to production. The only candidate source copies are `before/transaction.rs`, `before/xml.rs`, `after/transaction.rs`, and `after/xml.rs` in this directory.

## Candidate

`compact_changed_slide_xml` returns the existing compacted byte vector together with a transient `root_unchanged` witness. During its one existing `quick_xml::Reader` pass it records the source and emitted byte ranges for the single document element. `ReaderOrigin::offset(reader.buffer_position())` makes those ranges correct for a leading UTF-8 BOM. A witness is true only when both ranges exist, both `get` calls succeed, and the complete source and emitted root slices compare equal.

The pass also checks the source and emitted spans for every declaration, processing instruction, and comment outside the root. The check is conservative: a missing/invalid position or any byte difference makes the witness false. The existing event behavior remains otherwise unchanged. The only accepted outside-root removal is ASCII-whitespace `Text`; BOM, declaration, PI, and comment bytes are copied. CDATA, entity references, non-whitespace outside-root text, DTD, malformed roots, and all existing limits/refusals retain their old paths.

`compact_changed_slides` retains the existing sequence. A weak staged-read hit still means only the exact staged allocation was previously read as a Scene. A weak miss still performs `compaction_scene(after)` before compaction. After compaction, byte-identical output keeps the existing zero-read fast path. If bytes differ and `root_unchanged` is true, the compacted Scene reread is skipped. If the witness is false, the old `compaction_scene(after)`/`compaction_scene(compacted)` comparison remains in the same order. Debug builds rederive shape semantics directly on witness hits; those calls do not increment the production read counter.

## Proof obligations

Exact root equality preserves every shape and all root-contained bytes: element and attribute order, qualified names and namespace declarations, MCE directives and branches, unknown extensions, text/CDATA/entity bytes, `xml:space`, and internal lexical whitespace. XML declarations, PIs, comments, and the BOM are separately byte-checked or copied exactly. Therefore the only possible semantic difference is removal of outer ASCII whitespace, which the Scene scanner and MCE processor ignore. The compacted input is no larger, so input/output byte limits cannot newly fail; root depth, node, shape, and retained-text counts are unchanged.

The candidate never reuses spans from the old Scene. Final capture still reads the published compacted bytes, so offsets shifted by removed prolog whitespace remain safe. Staged Scene validation, compactor refusals, dependency checks, patch construction, and publication atomicity are unchanged. No transaction or snapshot cache is added; the witness and four optional positions live only for one compaction call.

## Expected read reduction

For a changed payload whose compaction removes only outer whitespace, the existing weak-hit path performs two redundant compaction Scene reads after the staged edit read; this candidate performs zero. A cold/weak-miss path retains the initial staged validation and changes two reads to one. Byte-identical compaction keeps its existing weak-hit zero-read/cold one-read behavior. Root-content whitespace or attribute-quote normalization makes the root witness false and retains the old two-read semantic comparison.

The opportunity is tied to the 0822 real-file profile: direct public edit p50 was 1.428437 ms; exact-owner `Scene::read_with` self-leaf samples were 111/108. The independent profile basis counted `Scene::read_with` in 636/624 Scene-qualified samples and `compaction_scene` in 301/304 of those overlapping stacks; these are inclusive stack-presence counts, not total calls or phase fractions. The rejected 0823 reader-transport trial improved the real row 2.394% (ratio 0.976062, interval 0.974798–0.984336), below the frozen 3% gate; this candidate removes repeated semantic reads in commit compaction and does not alter reader transport.

## Archive-only checks

The archive adds direct XML cases for declaration/PI/comment preservation, BOM and empty roots, unknown/MCE markup, root-content whitespace, quote normalization, entities, invalid outside-root text, DTDs, multiple roots, and unterminated roots. Transaction tests cover byte-identical compaction, root-different fallback, weak-hit `2 -> 0`, weak-miss `2 -> 1`, generated XML with CRLF after the declaration, and the admitted `test-data/ooxml/pptx/shapes.pptx` fixture. The real fixture test requires every witness hit with an accepted original Scene to have equal original and compacted `shape_semantics`; the CRLF test asserts the staged commit performs zero compaction Scene reads and removes the declaration-to-root CRLF.

No Cargo, formatter, native probe, Python analysis, or test command was run for this archive. Production source was not edited; the `before` copies were taken byte-for-byte from the base worktree before the candidate edits.
