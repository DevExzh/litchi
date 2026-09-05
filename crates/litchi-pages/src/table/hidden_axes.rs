//! Archive-free hidden-row and hidden-column values for Pages body tables.
//!
//! The semantic positions are shared with the other iWork table owners. A
//! The `HiddenAxes` value contains only typed zero-based positions and never
//! retains a package, archive member, protobuf message, wire payload, or
//! native object identifier. Pages package adapters are responsible for
//! resolving those positions against a concrete table and validating their
//! bounds before they cross the archive boundary.

/// One zero-based row or column position in a table.
pub use litchi_iwa_common::table::axis::AxisIndex;
/// Validation failures while constructing a hidden-axis collection.
pub use litchi_iwa_common::table::axis::Error;
/// Canonical, duplicate-free hidden rows and columns.
pub use litchi_iwa_common::table::axis::HiddenAxes;
/// Result type for constructing hidden rows and columns.
pub use litchi_iwa_common::table::axis::Result;

/// Exact-source transactions for one rooted Pages body table's hidden axes.
pub mod transaction {
    pub use crate::package::body_table_hidden_axes::{
        BodyTableHiddenAxesCommit as Commit, BodyTableHiddenAxesDiagnostics as Diagnostics,
        BodyTableHiddenAxesEdit as Edit, BodyTableHiddenAxesError as Error,
        BodyTableHiddenAxesLimitKind as LimitKind, BodyTableHiddenAxesPatch as Patch,
        BodyTableHiddenAxesPath as Path,
    };
}

#[cfg(test)]
mod tests {
    use std::mem::size_of;

    use super::{AxisIndex, Error, HiddenAxes};

    #[test]
    fn body_table_hidden_axes_are_sorted_and_duplicate_free() {
        let hidden = HiddenAxes::new([AxisIndex::column(2), AxisIndex::row(3), AxisIndex::row(1)])
            .unwrap_or_else(|error| panic!("valid hidden axes: {error}"));

        assert_eq!(
            hidden.as_slice(),
            [AxisIndex::row(1), AxisIndex::row(3), AxisIndex::column(2)]
        );
        assert!(hidden.contains(AxisIndex::row(3)));
        assert!(!hidden.contains(AxisIndex::column(1)));
        assert_eq!(hidden.iter().collect::<Vec<_>>(), hidden.as_slice());
    }

    #[test]
    fn body_table_hidden_axes_reject_duplicate_positions() {
        assert_eq!(
            HiddenAxes::new([AxisIndex::column(2), AxisIndex::column(2)]),
            Err(Error::Duplicate {
                axis: AxisIndex::column(2),
            })
        );
    }

    #[test]
    fn body_table_hidden_axes_have_no_native_state() {
        assert_eq!(size_of::<HiddenAxes>(), size_of::<Box<[AxisIndex]>>());
        assert_eq!(AxisIndex::row(4).index(), 4);
        assert_eq!(AxisIndex::column(7).index(), 7);
        assert_eq!(AxisIndex::row(4).to_string(), "row 4");
        assert_eq!(AxisIndex::column(7).to_string(), "column 7");
    }
}
