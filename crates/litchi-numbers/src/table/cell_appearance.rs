//! Cell fill, text layout and borders of one table cell (record 0412b).
//!
//! [`crate::Package::table_cell_appearance`] reads a cell's effective fill,
//! layout and explicit edge strokes;
//! [`crate::Package::edit_table_cell_appearance`] stages fills, layouts,
//! their resets and edge strokes. A change gives the cell a private style
//! variation (or rewrites its own), so every other property stays inherited
//! from the table's styles. The engine is the one the Pages and Keynote
//! families run (`litchi_iwa_table::cell_appearance`).

pub use crate::package::table_cell_appearance::{
    Commit, Diagnostics, Edit, Error, LimitKind, Patch, Refusal,
};
pub use litchi_iwa_common::shape::fill::Fill;
pub use litchi_iwa_common::shape::stroke::Stroke;
pub use litchi_iwa_common::table::cell::layout::{
    Inset, Insets, Layout, TextWrap, VerticalAlignment,
};
pub use litchi_iwa_common::table::cell::{BorderSide, Borders};

/// One cell's effective appearance.
#[derive(Debug, Clone, PartialEq)]
#[non_exhaustive]
pub struct Appearance {
    /// The fill the cell shows (`Fill::None` when no style sets one).
    pub fill: Fill,
    /// Text wrapping, vertical alignment and insets.
    pub layout: Layout,
    /// The explicit edge strokes (`None`: the table style's edge, or an
    /// explicitly cleared one).
    pub borders: Borders,
}
