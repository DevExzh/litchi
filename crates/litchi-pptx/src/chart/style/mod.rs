//! Layered Office 2013 chart-style semantics.

pub mod package;

pub use litchi_drawingml::chart::style::{
    Color, ColorDocument, ColorInfo, ColorKind, ColorMethod, ColorValue, Document, Entry,
    EntryKind, FontIndex, FontReference, Info, MarkerLayout, MarkerSymbol, Payload, Reference,
    Transform, TransformKind, Variation,
};
pub use package::{
    COLOR_CONTENT_TYPE, COLOR_RELATIONSHIP_TYPE, ColorPart, Part, STYLE_CONTENT_TYPE,
    STYLE_RELATIONSHIP_TYPE, discover,
};
