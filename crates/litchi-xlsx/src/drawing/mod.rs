//! Layered `SpreadsheetDrawing` ownership.
//!
//! `model` owns the contextual object inventory and `codec` owns the
//! namespace-aware, bounded XML reader. Shape authoring and full `DrawingML`
//! text parsing remain in their existing `shapes` owner; the text facade here
//! reuses [`litchi_drawingml`] without copying its vocabulary.

mod codec;
mod model;
pub mod source;
mod svg;

pub use super::chart::Anchor;
pub use codec::parse;
pub use model::{Chart, Drawing, Object, Picture, Unknown, UnknownKind, text};
pub use source::{
    ByteRange, DrawingDialect, ElementRange, PictureSource, RelationshipDialect,
    RelationshipReference, ScanLimits, SourceDrawing, SvgOwner, SvgOwnerState,
};
pub use svg::{PictureSelector, SvgInput};

/// Complete `SpreadsheetDrawing` anchor geometry.
pub use litchi_spreadsheet_drawing::shape::{
    Anchor as DrawingAnchor, CellMarker, EditAs, Emu, EmuExtent, EmuOffset,
};

#[cfg(test)]
mod source_tests;
#[cfg(test)]
mod tests;
