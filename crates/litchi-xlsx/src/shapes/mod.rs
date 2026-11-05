//! Semantic `SpreadsheetDrawing` shape facade.
//!
//! The shared owner provides the bounded XML grammar and semantic model;
//! this host retains only workbook/package relationship traversal.

mod package;

pub use litchi_spreadsheet_drawing::shape::{
    Anchor, AnchoredObject, Aspect, Autofit, Body, CellMarker, ClientData, Columns, ConnectionEnd,
    ConnectionShape, Coordinate32, Direction, EditAs, EditAsError, Emu, EmuExtent, EmuOffset,
    Geometry, Group, GroupTransform, Insets, Limits, NonVisual, Object, OleObject, Opaque,
    Paragraph, Preset, Properties, Run, Shape, Shapes, TextSize, Underline, Unknown,
    UnknownAttribute, UnknownElement, VerticalAnchor, Wrap, read, writer,
};
pub use package::{load_shapes, load_sheet_shapes};
