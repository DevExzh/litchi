//! Cell fill and text layout of one body-table cell (record 0371).
//!
//! [`crate::Package::body_table_cell_appearance`] reads a cell's effective
//! fill and layout; [`crate::Package::edit_body_table_cell_appearance`]
//! stages fills and layouts and their resets. A change gives the cell a
//! private style variation (or rewrites its own), so every other property
//! stays inherited from the table's styles.

pub use crate::package::body_table_cell_appearance::{
    Commit, Diagnostics, Edit, Error, LimitKind, Patch, Refusal,
};
pub use litchi_iwa_common::shape::fill::Fill;
pub use litchi_iwa_common::table::cell::layout::{
    Inset, Insets, Layout, TextWrap, VerticalAlignment,
};

/// One cell's effective appearance.
#[derive(Debug, Clone, PartialEq)]
#[non_exhaustive]
pub struct Appearance {
    /// The fill the cell shows (`Fill::None` when no style sets one).
    pub fill: Fill,
    /// Text wrapping, vertical alignment and insets.
    pub layout: Layout,
}
