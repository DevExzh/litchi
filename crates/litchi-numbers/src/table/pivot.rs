//! Pivot-table vocabulary: the aggregate categories a pivot group-by
//! exposes (records 0421 and 0424).
//!
//! A [`Category`](crate::table::pivot::Category) can be referenced from an
//! authored formula through
//! [`Expression::pivot_category`](crate::formula::Expression::pivot_category).

use std::fmt;

use thiserror::Error as ThisError;

/// The summary function a pivot group-by applies to an aggregated column
/// (`TST.ColumnAggregateArchive.agg_type`).
///
/// No source names Numbers' summary-function codes yet: the host only
/// copies the code, the protobuf schema declares it as a bare `uint32`, and
/// no tracked workbook holds an enabled group-by. Every code is therefore
/// [`Aggregate::Unnamed`] until the macOS run queued by record 0424 builds
/// a native pivot with each "Summarize by" function; the named variants
/// (`Sum`, `Average`, `Count`, …) are added from that evidence. Two
/// categories with the same summary function compare equal.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Aggregate {
    /// A summary function without native evidence for its name.
    Unnamed(UnnamedAggregate),
}

/// An opaque summary function whose native code is not yet named (see
/// [`Aggregate`]). It compares by its native code, which it does not expose.
#[derive(Clone, Copy, PartialEq, Eq, Hash)]
pub struct UnnamedAggregate {
    code: u32,
}

impl fmt::Debug for UnnamedAggregate {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("UnnamedAggregate")
            .finish_non_exhaustive()
    }
}

impl Aggregate {
    pub(crate) const fn from_code(code: u32) -> Self {
        Self::Unnamed(UnnamedAggregate { code })
    }

    pub(crate) const fn code(self) -> u32 {
        match self {
            Self::Unnamed(unnamed) => unnamed.code,
        }
    }
}

/// The native identity of a category: its group-by, aggregated column and
/// group node UUIDs as `(lower, upper)`. Never exposed.
pub(crate) type CategoryKey = ((u64, u64), (u64, u64), (u64, u64));

/// One pivot aggregate category: a group node of an enabled pivot group-by
/// and one of its aggregated columns (`Package::pivot_categories`).
///
/// A category identifies its group-by, column and group node privately, so
/// a formula can reference it
/// ([`Expression::pivot_category`](crate::formula::Expression::pivot_category));
/// it exposes no native identifier.
#[derive(Clone, PartialEq, Eq, Hash)]
pub struct Category {
    key: CategoryKey,
    label: Option<String>,
    group_level: i32,
    aggregate: Aggregate,
}

impl Category {
    pub(crate) const fn new(
        key: CategoryKey,
        label: Option<String>,
        group_level: i32,
        aggregate: Aggregate,
    ) -> Self {
        Self {
            key,
            label,
            group_level,
            aggregate,
        }
    }

    pub(crate) const fn key(&self) -> CategoryKey {
        self.key
    }

    /// The group's label: its group value as text, `Grand Total` for the
    /// root group, or `None` when the group has no value.
    #[must_use]
    pub fn label(&self) -> Option<&str> {
        self.label.as_deref()
    }

    /// The group's depth in the group-by tree (0 for the root group).
    #[must_use]
    pub const fn group_level(&self) -> i32 {
        self.group_level
    }

    /// The summary function of the aggregated column.
    #[must_use]
    pub const fn aggregate(&self) -> Aggregate {
        self.aggregate
    }
}

impl fmt::Debug for Category {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Category")
            .field("label", &self.label)
            .field("group_level", &self.group_level)
            .field("aggregate", &self.aggregate)
            .finish_non_exhaustive()
    }
}

/// A pivot-category read failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// The package holds several calculation engines.
    #[error("the Numbers package holds several calculation engines")]
    InvalidSource,
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn aggregates_compare_by_code_without_exposing_it() {
        let first = Aggregate::from_code(1);
        assert_eq!(first, Aggregate::from_code(1));
        assert_ne!(first, Aggregate::from_code(2));
        assert_eq!(first.code(), 1);
        let debug = format!("{first:?}");
        assert_eq!(debug, "Unnamed(UnnamedAggregate { .. })");
        let category = Category::new(((1, 2), (3, 4), (5, 6)), None, 0, first);
        assert!(!format!("{category:?}").contains('3'));
    }
}
