//! Archive-free iWork table sort semantics.
//!
//! The IWA adapter owns protobuf decoding, native wire preservation, and
//! package transactions. This module owns only the checked values that make
//! up a table sort order, independent of any concrete format crate.

use std::collections::BTreeSet;
use std::fmt;

const ENTIRE_TABLE: i32 = 0;
const ROW_RANGE: i32 = 1;
const ASCENDING: i32 = 0;
const DESCENDING: i32 = 1;

/// Failures returned while constructing a table sort value.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Error {
    /// A column index cannot be represented by the native `uint32` field.
    ColumnIndexOverflow {
        /// The rejected platform-sized index.
        index: usize,
    },
    /// A native column index cannot be represented by this platform's `usize`.
    NativeColumnIndexOverflow {
        /// The rejected native index.
        index: u32,
    },
    /// A selected-row range is empty or inverted.
    InvalidRowRange {
        /// Inclusive range start.
        start: usize,
        /// Exclusive range end.
        end: usize,
    },
    /// A sort order contains no rules.
    EmptyOrder,
    /// A sort order contains the same column more than once.
    DuplicateColumn {
        /// The repeated physical column.
        column: usize,
    },
    /// A native sort scope is not known to this semantic model.
    UnknownScope {
        /// The rejected native scope.
        value: i32,
    },
    /// A native sort direction is not known to this semantic model.
    UnknownDirection {
        /// The rejected native direction.
        value: i32,
    },
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::ColumnIndexOverflow { index } => write!(
                formatter,
                "iWork table sort column index {index} exceeds the native u32 range"
            ),
            Self::NativeColumnIndexOverflow { index } => write!(
                formatter,
                "iWork table sort native column index {index} exceeds usize"
            ),
            Self::InvalidRowRange { start, end } => write!(
                formatter,
                "iWork selected-row sort range {start}..{end} must be non-empty"
            ),
            Self::EmptyOrder => {
                formatter.write_str("iWork table sort order must contain at least one rule")
            },
            Self::DuplicateColumn { column } => write!(
                formatter,
                "iWork table sort order cannot contain column {column} more than once"
            ),
            Self::UnknownScope { value } => {
                write!(
                    formatter,
                    "iWork table sort order has unknown scope {value}"
                )
            },
            Self::UnknownDirection { value } => {
                write!(
                    formatter,
                    "iWork table sort rule has unknown direction {value}"
                )
            },
        }
    }
}

impl std::error::Error for Error {}

/// Result type for checked table sort values.
pub type Result<T> = std::result::Result<T, Error>;

/// Rows targeted by a persisted table sort configuration.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq, Hash)]
pub enum Scope {
    /// Apply the rules to every body row, excluding headers and footers.
    #[default]
    EntireTable,
    /// Apply the rules to the rows selected in the document view.
    SelectedRows,
}

impl Scope {
    /// Decode a native sort scope.
    ///
    /// # Errors
    ///
    /// Returns [`Error::UnknownScope`] for a native value outside the known
    /// iWork sort-scope domain.
    pub const fn from_native(value: i32) -> Result<Self> {
        match value {
            ENTIRE_TABLE => Ok(Self::EntireTable),
            ROW_RANGE => Ok(Self::SelectedRows),
            other => Err(Error::UnknownScope { value: other }),
        }
    }

    /// Return the native sort scope value.
    #[must_use]
    pub const fn native_value(self) -> i32 {
        match self {
            Self::EntireTable => ENTIRE_TABLE,
            Self::SelectedRows => ROW_RANGE,
        }
    }
}

/// A non-empty, body-relative half-open row range for selected-row sorting.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub struct RowRange {
    start: usize,
    end: usize,
}

impl RowRange {
    /// Construct a non-empty body-relative range `[start, end)`.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidRowRange`] when `start` is not less than
    /// `end`.
    pub const fn new(start: usize, end: usize) -> Result<Self> {
        if start >= end {
            return Err(Error::InvalidRowRange { start, end });
        }
        Ok(Self { start, end })
    }

    /// Return the inclusive body-relative start row.
    #[must_use]
    pub const fn start(self) -> usize {
        self.start
    }

    /// Return the exclusive body-relative end row.
    #[must_use]
    pub const fn end(self) -> usize {
        self.end
    }

    /// Return the number of selected body rows.
    #[must_use]
    pub const fn len(self) -> usize {
        self.end - self.start
    }

    /// Return whether this range is empty.
    #[must_use]
    pub const fn is_empty(self) -> bool {
        self.start == self.end
    }
}

/// A validated zero-based physical column index used by a table sort rule.
#[derive(Clone, Copy, Debug, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct ColumnIndex(usize);

impl ColumnIndex {
    /// Construct a native-compatible zero-based column index.
    ///
    /// # Errors
    ///
    /// Returns [`Error::ColumnIndexOverflow`] when `index` does not fit
    /// The native `uint32` representation.
    pub const fn new(index: usize) -> Result<Self> {
        if index > u32::MAX as usize {
            return Err(Error::ColumnIndexOverflow { index });
        }
        Ok(Self(index))
    }

    /// Decode a native zero-based column index without truncation.
    ///
    /// # Errors
    ///
    /// Returns [`Error::NativeColumnIndexOverflow`] when the native value
    /// cannot be represented by this platform's `usize`.
    pub fn from_native(index: u32) -> Result<Self> {
        let compact = usize::try_from(index)
            .map_err(|_conversion| Error::NativeColumnIndexOverflow { index })?;
        Self::new(compact)
    }

    /// Return the zero-based column index.
    #[must_use]
    pub const fn get(self) -> usize {
        self.0
    }

    /// Return the native `uint32` column index.
    #[must_use]
    #[allow(
        clippy::cast_possible_truncation,
        reason = "ColumnIndex::new enforces the native u32 bound"
    )]
    pub const fn native_value(self) -> u32 {
        self.0 as u32
    }
}

impl TryFrom<usize> for ColumnIndex {
    type Error = Error;

    fn try_from(index: usize) -> Result<Self> {
        Self::new(index)
    }
}

impl From<ColumnIndex> for usize {
    fn from(index: ColumnIndex) -> Self {
        index.get()
    }
}

/// Sort direction for one table column.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum Direction {
    /// Sort low-to-high, alphabetically A-to-Z, or oldest-to-newest.
    Ascending,
    /// Sort high-to-low, alphabetically Z-to-A, or newest-to-oldest.
    Descending,
}

impl Direction {
    /// Decode a native sort direction.
    ///
    /// # Errors
    ///
    /// Returns [`Error::UnknownDirection`] for a native value outside the
    /// known iWork sort-direction domain.
    pub const fn from_native(value: i32) -> Result<Self> {
        match value {
            ASCENDING => Ok(Self::Ascending),
            DESCENDING => Ok(Self::Descending),
            other => Err(Error::UnknownDirection { value: other }),
        }
    }

    /// Return the native sort direction value.
    #[must_use]
    pub const fn native_value(self) -> i32 {
        match self {
            Self::Ascending => ASCENDING,
            Self::Descending => DESCENDING,
        }
    }
}

/// One sort-configuration rule in priority order.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub struct Rule {
    column: ColumnIndex,
    direction: Direction,
}

impl Rule {
    /// Construct a rule for one physical table column.
    #[must_use]
    pub const fn new(column: ColumnIndex, direction: Direction) -> Self {
        Self { column, direction }
    }

    /// Return the column selected by this rule.
    #[must_use]
    pub const fn column(self) -> ColumnIndex {
        self.column
    }

    /// Return this rule's direction.
    #[must_use]
    pub const fn direction(self) -> Direction {
        self.direction
    }
}

/// An ordered, non-empty table sort-rule configuration.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Order {
    scope: Scope,
    rules: Vec<Rule>,
}

impl Order {
    /// Construct a full-table sort-rule configuration.
    ///
    /// # Errors
    ///
    /// Returns [`Error::EmptyOrder`] for an empty rule sequence or
    /// [`Error::DuplicateColumn`] when a column occurs more than once.
    pub fn new(rules: impl IntoIterator<Item = Rule>) -> Result<Self> {
        Self::with_scope(Scope::EntireTable, rules)
    }

    /// Construct a selected-row sort-rule configuration.
    ///
    /// # Errors
    ///
    /// Returns [`Error::EmptyOrder`] for an empty rule sequence or
    /// [`Error::DuplicateColumn`] when a column occurs more than once.
    pub fn selected_rows(rules: impl IntoIterator<Item = Rule>) -> Result<Self> {
        Self::with_scope(Scope::SelectedRows, rules)
    }

    /// Construct a sort-rule configuration with an explicit scope.
    ///
    /// # Errors
    ///
    /// Returns [`Error::EmptyOrder`] for an empty rule sequence or
    /// [`Error::DuplicateColumn`] when a column occurs more than once.
    pub fn with_scope(scope: Scope, rule_iter: impl IntoIterator<Item = Rule>) -> Result<Self> {
        let rules = rule_iter.into_iter().collect::<Vec<_>>();
        if rules.is_empty() {
            return Err(Error::EmptyOrder);
        }
        let mut columns = BTreeSet::new();
        for rule in &rules {
            let column = rule.column.get();
            if !columns.insert(rule.column) {
                return Err(Error::DuplicateColumn { column });
            }
        }
        Ok(Self { scope, rules })
    }

    /// Return the persisted sort scope.
    #[must_use]
    pub const fn scope(&self) -> Scope {
        self.scope
    }

    /// Borrow the rules in native evaluation order.
    #[must_use]
    pub fn rules(&self) -> &[Rule] {
        &self.rules
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn native_scope_and_direction_values_are_lossless() {
        assert_eq!(Scope::from_native(0), Ok(Scope::EntireTable));
        assert_eq!(Scope::from_native(1), Ok(Scope::SelectedRows));
        assert_eq!(Scope::EntireTable.native_value(), 0);
        assert_eq!(Scope::SelectedRows.native_value(), 1);
        assert_eq!(Direction::from_native(0), Ok(Direction::Ascending));
        assert_eq!(Direction::from_native(1), Ok(Direction::Descending));
        assert_eq!(Direction::Ascending.native_value(), 0);
        assert_eq!(Direction::Descending.native_value(), 1);
        assert!(matches!(
            Scope::from_native(9),
            Err(Error::UnknownScope { value: 9 })
        ));
        assert!(matches!(
            Direction::from_native(9),
            Err(Error::UnknownDirection { value: 9 })
        ));
    }

    #[test]
    fn column_index_preserves_native_bounds() {
        assert_eq!(ColumnIndex::new(0).unwrap().native_value(), 0);
        assert_eq!(
            ColumnIndex::new(u32::MAX as usize).unwrap().native_value(),
            u32::MAX
        );
        if let Ok(too_large) = usize::try_from(u64::from(u32::MAX) + 1) {
            assert!(matches!(
                ColumnIndex::new(too_large),
                Err(Error::ColumnIndexOverflow { index }) if index == too_large
            ));
        }
    }

    #[test]
    fn row_range_and_order_validate_their_invariants() {
        assert!(matches!(
            RowRange::new(0, 0),
            Err(Error::InvalidRowRange { start: 0, end: 0 })
        ));
        assert!(RowRange::new(2, 5).is_ok());

        let column = ColumnIndex::new(1).unwrap();
        let rule = Rule::new(column, Direction::Ascending);
        assert!(matches!(Order::new([]), Err(Error::EmptyOrder)));
        assert!(matches!(
            Order::new([rule, Rule::new(column, Direction::Descending)]),
            Err(Error::DuplicateColumn { column: 1 })
        ));
        assert_eq!(
            Order::selected_rows([rule]).unwrap().scope(),
            Scope::SelectedRows
        );
        let empty = RowRange { start: 2, end: 2 };
        assert!(empty.is_empty());
    }
}
