//! Archive-free semantic Pages table values.

/// Effective style, gridline, banding, and row-sizing settings for a body table.
pub mod appearance;
/// Checked, archive-free row and column dimension values.
pub mod dimension;
/// Lossless header, footer, freeze, and print-repetition settings.
pub mod headers;
/// Archive-free hidden-row and hidden-column values.
pub mod hidden_axes;
/// Interactive lock state for a body-attached table.
pub mod lock;
/// Archive-free merged-cell geometry and topology algebra.
pub mod merge;
/// Archive-free table values together with sparse cell comments.
pub mod read {
    pub use litchi_iwa_common::table::read::*;
}
/// Validated, archive-free names for Pages body tables.
pub mod name;
/// Persisted row-order rules for a body-attached table.
pub mod sort;
pub mod title;
