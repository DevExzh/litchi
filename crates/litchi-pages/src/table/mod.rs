//! Archive-free semantic Pages table values.

/// Checked, archive-free row and column dimension values.
pub mod dimension;
/// Lossless header, footer, freeze, and print-repetition settings.
pub mod headers;
/// Interactive lock state for a body-attached table.
pub mod lock;
/// Persisted row-order rules for a body-attached table.
pub mod sort;
pub mod title;
