//! Neutral table vocabulary shared by concrete iWork format owners.

pub mod appearance;
pub mod axis;
pub mod cell;
/// Compact, archive-free cell coordinates and A1 selectors.
pub mod coordinate;
/// Checked, archive-free row, column, and point-size semantics.
pub mod dimension;
/// Lossless, archive-free table header and footer settings.
pub mod headers;
pub mod lock;
/// Shared sparse cell values, table extents, builders, and bounded views.
pub mod model;
/// Checked, archive-free table sort semantics.
pub mod sort;
/// Lossless, archive-free table title settings.
pub mod title;
