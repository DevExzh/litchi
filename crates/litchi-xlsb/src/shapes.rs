//! Semantic SpreadsheetDrawing shape facade.
//!
//! XLSB worksheet streams point to the same standard `xdr` XML parts as XLSX;
//! their bounded grammar and semantic model are owned by the shared crate.

pub use litchi_spreadsheet_drawing::shape::{
    Anchor, AnchoredObject, Aspect, Autofit, Body, CellMarker, ClientData, Columns, ConnectionEnd,
    ConnectionShape, Coordinate32, Direction, EditAs, EditAsError, Emu, EmuExtent, EmuOffset,
    Geometry, Group, GroupTransform, Insets, Limits, NonVisual, Object, OleObject, Opaque,
    Paragraph, Preset, Properties, Run, Shape, Shapes, TextSize, Underline, Unknown,
    UnknownAttribute, UnknownElement, VerticalAnchor, Wrap, read, writer,
};
