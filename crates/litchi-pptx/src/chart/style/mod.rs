//! Layered Office 2013 chart-style semantics.

pub mod codec;
pub mod model;
pub mod package;

pub use codec::{
    COLOR_CONTENT_TYPE, COLOR_RELATIONSHIP_TYPE, STYLE_CONTENT_TYPE, STYLE_RELATIONSHIP_TYPE,
};
pub use model::*;
pub use package::discover;

/// Borrowed chart-style companion part.
pub struct Part<'a> {
    pub(crate) part: &'a dyn litchi_opc::part::Part,
}

/// Borrowed chart-color-style companion part.
pub struct ColorPart<'a> {
    pub(crate) part: &'a dyn litchi_opc::part::Part,
}
