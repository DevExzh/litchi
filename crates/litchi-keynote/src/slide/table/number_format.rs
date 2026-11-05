//! Decimal Number display formats for existing Keynote table cells.
//!
//! The Keynote-local semantic value mirrors the neutral iWork Number domain.
//! Native table identifiers,
//! format-list keys, BNC records, and archive bytes remain private to the
//! Keynote package adapter.

/// Failure from a checked Number-format value constructor (the shared
/// number-format error).
pub use litchi_iwa_common::table::cell::data_format::number::Error as NumberError;
/// The shared iWork decimal Number display format and its parts (D4,
/// record 0368): the same nominal types as Numbers'
/// `cell::data_format::number`.
pub use litchi_iwa_common::table::cell::data_format::number::{
    DecimalPlaces, FixedDecimalPlaces, MAX_DECIMAL_PLACES, NegativeStyle, Number,
    ThousandsSeparator,
};
/// A checked semantic table-cell coordinate: the shared iWork grid
/// coordinate (`litchi_iwa_common::table::coordinate::At`, the same type as
/// [`crate::slide::table::cells::At`]). The package adapter checks it
/// against the selected table's declared extent before it inspects any cell
/// storage.
pub use litchi_iwa_common::table::coordinate::At;
/// Checked coordinate-construction failure (the shared coordinate error).
pub use litchi_iwa_common::table::coordinate::Error as AtError;

/// Exact-source transactions for one existing slide-table cell.
pub mod transaction {
    pub use crate::package::slide_table_cell_number_format::{
        Commit, Diagnostics, Edit, Error, LimitKind, Patch, Path, Refusal,
    };
}
