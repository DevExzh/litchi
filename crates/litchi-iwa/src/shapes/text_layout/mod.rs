//! Typed frame-level layout for text inside shapes and ordinary text boxes.
//!
//! iWork stores vertical alignment, edge insets, and autosizing on the shape
//! style. Text-box columns are another independently composable shape-style
//! property; document-body columns use a separate column-style archive.

mod native;
mod style;

pub(crate) use style::{reset_shape_text_layout, set_shape_text_layout, shape_text_layout};
impl From<litchi_iwa_common::text::layout::Error> for crate::Error {
    fn from(error: litchi_iwa_common::text::layout::Error) -> Self {
        Self::ParseError(error.to_string())
    }
}
