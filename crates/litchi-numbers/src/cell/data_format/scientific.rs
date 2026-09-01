//! Scientific-notation display values and selector-first transactions.
//!
//! Scientific notation is a distinct semantic family even though Numbers
//! stores its fixed precision beside the plain Number and Percentage formats.
//! The nominal type remains compatible with the decimal value module while
//! this namespace owns the focused Scientific transaction entry point.

pub use super::number::{FixedDecimalPlaces, Scientific};

/// Selector-first exact-source transactions for one existing table cell's
/// explicit Scientific format.
pub mod transaction {
    pub use crate::package::table_cell_scientific_format::{
        Commit, Diagnostics, Edit, Error, LimitKind, Patch, Path,
    };
}
