//! Dependency-free hidden-row and hidden-column vocabulary for tables.
//!
//! Native hidden-state archives, stable axis UUIDs, package traversal, and
//! wire mutation remain in the concrete iWork adapters. This module owns only
//! the checked semantic positions and their deterministic collection.

use std::fmt;

/// One zero-based row or column position in a table.
#[allow(
    clippy::module_name_repetitions,
    reason = "AxisIndex distinguishes a table position from the enclosing axis module"
)]
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub enum AxisIndex {
    /// A zero-based row index.
    Row(usize),
    /// A zero-based column index.
    Column(usize),
}

impl AxisIndex {
    /// Address one row by zero-based index.
    #[must_use]
    pub const fn row(index: usize) -> Self {
        Self::Row(index)
    }

    /// Address one column by zero-based index.
    #[must_use]
    pub const fn column(index: usize) -> Self {
        Self::Column(index)
    }

    /// Return the zero-based index within this position's axis.
    #[must_use]
    pub const fn index(self) -> usize {
        match self {
            Self::Row(index) | Self::Column(index) => index,
        }
    }
}

impl fmt::Display for AxisIndex {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Row(index) => write!(formatter, "row {index}"),
            Self::Column(index) => write!(formatter, "column {index}"),
        }
    }
}

/// Validation failures for a hidden-axis collection.
#[derive(Clone, Copy, Debug, PartialEq, Eq, thiserror::Error)]
pub enum Error {
    /// The same row or column was supplied more than once.
    #[error("hidden table axes contain duplicate position {axis}")]
    Duplicate {
        /// The repeated semantic position.
        axis: AxisIndex,
    },
}

/// Result type for hidden-axis collection construction.
pub type Result<T> = std::result::Result<T, Error>;

/// Canonical, duplicate-free set of user-hidden table axes.
///
/// Positions are stored in row-then-column order, followed by ascending
/// position within each axis. The boxed slice keeps the immutable collection
/// to one allocation without retaining vector capacity.
#[repr(transparent)]
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct HiddenAxes {
    axes: Box<[AxisIndex]>,
}

impl HiddenAxes {
    /// No rows or columns are hidden.
    #[must_use]
    pub fn empty() -> Self {
        Self { axes: Box::new([]) }
    }

    /// Construct a sorted hidden-axis set, rejecting duplicate positions.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Duplicate`] when the input contains the same row or
    /// column position more than once.
    #[must_use = "use the validated hidden-axis set or handle its validation error"]
    pub fn new(input: impl IntoIterator<Item = AxisIndex>) -> Result<Self> {
        let mut axes = input.into_iter().collect::<Vec<_>>();
        axes.sort_unstable();
        if let Some(axis) = axes
            .windows(2)
            .find_map(|pair| (pair[0] == pair[1]).then_some(pair[0]))
        {
            return Err(Error::Duplicate { axis });
        }
        Ok(Self {
            axes: axes.into_boxed_slice(),
        })
    }

    /// Borrow the canonical row-then-column positions.
    #[must_use]
    pub fn as_slice(&self) -> &[AxisIndex] {
        &self.axes
    }

    /// Iterate over canonical row-then-column positions.
    #[must_use]
    pub fn iter(&self) -> impl ExactSizeIterator<Item = AxisIndex> + '_ {
        self.axes.iter().copied()
    }

    /// Return whether every table axis is visible.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.axes.is_empty()
    }

    /// Return whether one row or column is hidden.
    #[must_use]
    pub fn contains(&self, axis: AxisIndex) -> bool {
        self.axes.binary_search(&axis).is_ok()
    }
}

impl Default for HiddenAxes {
    fn default() -> Self {
        Self::empty()
    }
}

#[cfg(test)]
mod tests {
    use std::mem::size_of;

    use super::{AxisIndex, Error, HiddenAxes};

    #[test]
    fn hidden_axes_are_sorted_and_compact() {
        let hidden = HiddenAxes::new([AxisIndex::column(2), AxisIndex::row(3), AxisIndex::row(1)])
            .unwrap_or_else(|error| panic!("valid hidden axes: {error}"));

        assert_eq!(size_of::<HiddenAxes>(), size_of::<Box<[AxisIndex]>>());
        assert_eq!(
            hidden.as_slice(),
            [AxisIndex::row(1), AxisIndex::row(3), AxisIndex::column(2)]
        );
        assert_eq!(hidden.iter().len(), 3);
        assert!(hidden.contains(AxisIndex::row(3)));
        assert!(!hidden.contains(AxisIndex::column(1)));
    }

    #[test]
    fn hidden_axes_reject_duplicate_positions_with_typed_error() {
        assert_eq!(
            HiddenAxes::new([AxisIndex::column(2), AxisIndex::column(2)]),
            Err(Error::Duplicate {
                axis: AxisIndex::column(2),
            })
        );
    }

    #[test]
    fn axis_index_helpers_are_stable() {
        assert_eq!(AxisIndex::row(4).index(), 4);
        assert_eq!(AxisIndex::column(7).index(), 7);
        assert_eq!(AxisIndex::row(4).to_string(), "row 4");
        assert_eq!(AxisIndex::column(7).to_string(), "column 7");
        assert!(HiddenAxes::empty().is_empty());
        assert_eq!(HiddenAxes::default(), HiddenAxes::empty());
    }
}
