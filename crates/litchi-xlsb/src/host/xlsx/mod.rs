//! DrawingML support local to the XLSB package layer.
//!
//! XLSB drawings reuse the SpreadsheetDrawing and chart XML grammars. These
//! types live here so the raw BIFF12 kernel remains independent of package
//! services and the XLSB crate does not depend on the XLSX crate.

pub mod chart;
pub mod pivot_chart;
pub mod shapes;
pub mod views;
pub mod writer;

pub use chart::{
    Chart, ChartAnchor, ChartExternalDataPart, ChartExternalDataTarget, ChartUserShapesPart,
    Relationship, RelationshipTarget,
};
pub use litchi_drawingml::geom::Preset;
pub use shapes::{
    AnchoredObject, Autofit, Body, CellMarker, ClientData, Columns, ConnectionEnd, ConnectionShape,
    Coordinate32, Direction, DrawingObject, DrawingOleObject, EditAs, EditAsError, Emu, EmuExtent,
    EmuOffset, Geometry, Group, GroupTransform, Insets, NonVisual, OleObjectAspect, Paragraph,
    Properties, Run, Shape, ShapeAnchor, Shapes, TextSize, Underline, VerticalAnchor, Wrap,
};
pub use writer::{ConnectionEndSpec, ConnectionShapeSpec, DrawingObjectSpec, GroupSpec, ShapeSpec};
