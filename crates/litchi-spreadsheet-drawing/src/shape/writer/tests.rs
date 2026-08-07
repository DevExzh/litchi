#![allow(clippy::unwrap_used, reason = "test assertions require direct values")]

use super::{Emitter, ShapeSpec};
use crate::shape::{Anchor, CellMarker, EditAs, Emu};
use litchi_drawingml::geom::Preset;

#[test]
fn emitter_writes_a_fresh_shape() {
    let anchor = Anchor::TwoCell {
        from: CellMarker {
            column: 0,
            column_offset: Emu(0),
            row: 0,
            row_offset: Emu(0),
        },
        to: CellMarker {
            column: 1,
            column_offset: Emu(0),
            row: 1,
            row_offset: Emu(0),
        },
        edit_as: EditAs::TwoCell,
    };
    let mut xml = String::new();
    Emitter::new(1)
        .write_anchored_shape(
            &mut xml,
            &ShapeSpec::text_box("Title", anchor, Preset::Rect, "Hello"),
        )
        .unwrap();
    assert!(xml.contains("Hello"));
}
