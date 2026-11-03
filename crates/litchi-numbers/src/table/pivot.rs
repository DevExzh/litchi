//! Pivot-table vocabulary: the aggregate categories a pivot group-by
//! exposes (record 0421).

use thiserror::Error as ThisError;

/// One pivot aggregate category: a group node of an enabled pivot group-by
/// and one of its aggregated columns (`Package::pivot_categories`).
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct Category {
    label: Option<String>,
    group_level: i32,
    aggregate_type: u32,
}

impl Category {
    pub(crate) const fn new(label: Option<String>, group_level: i32, aggregate_type: u32) -> Self {
        Self {
            label,
            group_level,
            aggregate_type,
        }
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

    /// Numbers' summary-function code of the aggregated column
    /// (`TST.ColumnAggregateArchive.agg_type`).
    #[must_use]
    pub const fn aggregate_type(&self) -> u32 {
        self.aggregate_type
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
