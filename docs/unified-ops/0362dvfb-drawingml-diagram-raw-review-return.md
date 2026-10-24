# 0362dvfb — DrawingML diagram raw review return

- Status: source repair and regressions are committed for coordinator review;
  focused Cargo validation is pending a newly granted validation slot.
- Parent source: `e98076db975242b74e0653f06d6094a1e63bb6cb`.
- This packet has no API signature change, so existing DOCX, PPTX, ops and
  SmartArt-harness consumers remain migrated to the canonical limits APIs.
- The canonical `0362dvf` record now consistently says that the focused
  DrawingML run passed 245 unit tests (the old trailing 240 was a stale
  contradiction).

## Concrete review findings and repairs

The `inventory_font_cache_limits` re-review identified seven parser/authoring
gaps. The source repair addresses each at its owning path:

1. `diagram/model.rs::SmartArt::parse_data_xml` now counts raw `dgm:pt` and
   `dgm:cxn` records against `Limits::points_value()` and
   `Limits::connections_value()`, reports the typed observed/max limits, uses
   `data_text_bytes_value()` for point text, and checks reconstructed semantic
   depth against `node_depth_value()` before descending.
2. A paired empty `dgm:dataModel` now requires the direct `dgm:ptLst` grammar;
   the existing self-closing-root rejection remains. The parser rejects
   nested/duplicate data-model roots and misplaced `pt`, `ptLst`, `cxn` and
   `cxnLst` elements instead of silently ignoring them.
3. `RawConnection` retains `srcOrd`. Parent and root sibling vectors are
   sorted by that source order with a stable point-order tie-breaker, so valid
   sibling order survives parsing.
4. Empty content points are retained, allowing an empty parent with children
   to remain in the semantic tree. A `parOf` connection whose source or
   destination is not the document/content graph now returns `Invalid` rather
   than promoting or dropping the child.
5. `diagram/data/codec.rs::DiagramDataModel::parse` rejects a nested
   `dgm:dataModel` root explicitly.
6. `diagram/edit.rs::read_drawing` requires exactly one direct `dsp:spTree`,
   so a paired `dsp:drawing` without a shape tree is rejected while a valid
   empty shape tree remains readable.
7. `diagram/authoring.rs::Spec::parse_outline` checks custom node-count,
   semantic-depth and node-text limits before each `Node` allocation; invalid
   XML text is also rejected before insertion.

## Regression coverage

The packet adds behavior tests for:

- independent points, connections, data-text and semantic-depth limits;
- paired-empty data models, misplaced structural elements, nested roots and
  dangling parent connections;
- empty-parent retention and `srcOrd` sibling ordering;
- the paired drawing-root/no-`spTree` case;
- nested typed `DiagramDataModel` roots; and
- authoring node/depth/text policies before tree growth.

No fixture, binary, performance artifact or `3rdparty/` reference was added.
No Cargo command was run for this source-only repair before a new validation
slot was granted.

## Root-order follow-up

The follow-up review found one deterministic-order defect in the first repair:
root entries were sorted only by `srcOrd`, so equal document-level orders could
be reordered by `sort_unstable`. `diagram/model.rs` now uses the same
`(srcOrd, point-index)` key for roots and child vectors. A regression supplies
three roots whose points and connections are in different orders but whose
`srcOrd` values are equal, and asserts stable point order.

The immediately following audit found that an unconditional second
`dataModel` guard accidentally rejected every normal paired root. That guard
was removed; nested-root rejection remains in the depth-aware root branch. A
new regression now generates a nonempty SmartArt data part with a paired root,
parses it successfully and checks its node text. This source correction is
committed but remains untested by Cargo until the validation slot is granted.
