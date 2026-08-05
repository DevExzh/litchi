//! Migration aliases for the common table-cell layout vocabulary.

pub use litchi_iwa_common::table::cell::layout::{
    Error as TableCellLayoutError, Inset, Insets, Layout, TextWrap, VerticalAlignment,
};

/// Contextual facade name for [`Inset`].
pub type TableCellInset = Inset;
/// Contextual facade name for [`Insets`].
pub type TableCellInsets = Insets;
/// Contextual facade name for [`Layout`].
pub type TableCellLayout = Layout;
/// Contextual facade name for [`TextWrap`].
pub type TableCellTextWrap = TextWrap;
/// Contextual facade name for [`VerticalAlignment`].
pub type TableCellVerticalAlignment = VerticalAlignment;

impl From<TableCellLayoutError> for crate::Error {
    fn from(error: TableCellLayoutError) -> Self {
        Self::ParseError(error.to_string())
    }
}
