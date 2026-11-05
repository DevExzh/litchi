//! Shared numeric-control values for Slider and Stepper cells.
//!
//! The values are the neutral `litchi_iwa_common::table::cell::data_format::control`
//! vocabulary (record 0368), re-exported here at their established path.

pub use litchi_iwa_common::table::cell::data_format::control::{
    CellControl, Checkbox, DisplayFormat, Error, NotCellControl, PopUpMenu, Range, Result, Slider,
    StarRating, Stepper,
};

/// Exact-source selector-first transactions for one table cell's control.
///
/// The package adapter owns the rooted table/control graph, native wire
/// preservation, and publication. This semantic namespace exposes only the
/// archive-free transaction handles; native identifiers, generated messages,
/// and package bytes never cross the boundary.
pub mod transaction {
    pub use crate::package::table_cell_control::{
        Commit, Diagnostics, Edit, Error, LimitKind, Patch, Path,
    };
}
