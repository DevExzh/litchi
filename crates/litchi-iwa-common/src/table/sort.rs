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

/// Archive-free planning primitives used by concrete physical table owners.
///
/// This module deliberately knows nothing about IWA objects, native row IDs,
/// cell storage, or package mutation. It owns only the checked, stable mapping
/// from body-row destinations to body-row sources. Concrete format adapters
/// remain responsible for decoding semantic keys and moving every row-affine
/// storage structure transactionally.
#[doc(hidden)]
pub mod planning {
    use std::cmp::Ordering;
    use std::fmt;

    use super::Order;

    /// A checked zero-based offset within the body rows selected for sorting.
    #[derive(Clone, Copy, Debug, PartialEq, Eq, PartialOrd, Ord, Hash)]
    pub struct BodyRowOffset(usize);

    impl BodyRowOffset {
        /// Construct an offset within the selected body-row slice.
        #[must_use]
        pub const fn new(offset: usize) -> Self {
            Self(offset)
        }

        /// Return the zero-based body-row offset.
        #[must_use]
        pub const fn get(self) -> usize {
            self.0
        }
    }

    /// Failures produced while validating or allocating a stable row plan.
    #[derive(Debug, Clone, Copy, PartialEq, Eq)]
    #[non_exhaustive]
    pub enum Error {
        /// A row key does not contain exactly one scalar per configured rule.
        KeyArity {
            /// Body-relative row offset.
            row: BodyRowOffset,
            /// Number of configured rules.
            expected: usize,
            /// Number of supplied key values.
            actual: usize,
        },
        /// One rule encounters different scalar domains across body rows.
        MixedScalarKinds {
            /// Body-relative row offset containing the mismatched scalar.
            row: BodyRowOffset,
            /// Zero-based rule position.
            rule: usize,
        },
        /// Temporary storage for the plan could not be reserved.
        Allocation {
            /// Number of body rows requested by the plan.
            rows: usize,
        },
        /// The internally produced source mapping is not a complete bijection.
        InvalidPermutation,
    }

    impl fmt::Display for Error {
        fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
            match self {
                Self::KeyArity {
                    row,
                    expected,
                    actual,
                } => write!(
                    formatter,
                    "iWork body-row sort key {} contains {actual} values; expected {expected}",
                    row.get()
                ),
                Self::MixedScalarKinds { row, rule } => write!(
                    formatter,
                    "iWork body-row sort rule {rule} changes scalar kind at row {}",
                    row.get()
                ),
                Self::Allocation { rows } => write!(
                    formatter,
                    "failed to allocate an iWork body-row sort plan for {rows} rows"
                ),
                Self::InvalidPermutation => {
                    formatter.write_str("iWork body-row sort plan is not a permutation")
                },
            }
        }
    }

    impl std::error::Error for Error {}

    /// A validated stable source mapping for one selected body-row slice.
    ///
    /// Element `d` stores the original body-row offset that must be written to
    /// destination offset `d`. All offsets are relative to the selected slice;
    /// concrete storage adapters add the physical body-row start exactly once.
    #[derive(Clone, Debug, PartialEq, Eq)]
    pub struct RowPermutation {
        sources_by_destination: Vec<BodyRowOffset>,
    }

    impl RowPermutation {
        /// Return whether at least one row changes destination.
        #[must_use]
        pub fn changes_rows(&self) -> bool {
            self.sources_by_destination
                .iter()
                .enumerate()
                .any(|(destination, source)| source.get() != destination)
        }

        /// Return the number of rows covered by this plan.
        #[must_use]
        pub fn len(&self) -> usize {
            self.sources_by_destination.len()
        }

        /// Return whether the plan covers no rows.
        #[must_use]
        pub fn is_empty(&self) -> bool {
            self.sources_by_destination.is_empty()
        }

        /// Borrow source offsets in destination order.
        #[must_use]
        pub fn sources_by_destination(&self) -> &[BodyRowOffset] {
            &self.sources_by_destination
        }

        /// Build the inverse destination mapping in source order.
        ///
        /// # Errors
        ///
        /// Returns [`Error::Allocation`] when the inverse cannot be reserved,
        /// or [`Error::InvalidPermutation`] if the stored map is not bijective.
        pub fn destinations_by_source(&self) -> Result<Vec<BodyRowOffset>, Error> {
            let rows = self.sources_by_destination.len();
            let mut destinations = Vec::new();
            destinations
                .try_reserve_exact(rows)
                .map_err(|_allocation| Error::Allocation { rows })?;
            let unassigned = BodyRowOffset::new(rows);
            destinations.resize(rows, unassigned);

            for (destination, source) in self.sources_by_destination.iter().copied().enumerate() {
                let Some(slot) = destinations.get_mut(source.get()) else {
                    return Err(Error::InvalidPermutation);
                };
                if *slot != unassigned {
                    return Err(Error::InvalidPermutation);
                }
                *slot = BodyRowOffset::new(destination);
            }

            if destinations.contains(&unassigned) {
                return Err(Error::InvalidPermutation);
            }
            Ok(destinations)
        }
    }

    /// Plan one deterministic stable ordering over already-decoded row keys.
    ///
    /// `kind` assigns a comparable scalar domain to each key value. Every row
    /// must carry exactly one value per rule, and values at the same rule
    /// position must have one consistent domain. `compare` performs the
    /// domain-specific comparison after those invariants are checked.
    ///
    /// The implementation uses an allocation-free unstable sort plus the
    /// original body-row offset as an explicit final key. This produces stable
    /// duplicate ordering without a second scratch allocation.
    ///
    /// # Errors
    ///
    /// Returns a typed validation error for malformed key matrices, or an
    /// allocation error before any caller-owned state is mutated.
    pub fn stable_body_row_permutation<K, Kind>(
        keys_by_body_row: &[Vec<K>],
        order: &Order,
        kind: impl Fn(&K) -> Kind,
        compare: impl Fn(&K, &K) -> Ordering,
    ) -> Result<RowPermutation, Error>
    where
        Kind: Eq,
    {
        let rules = order.rules();
        for (row, keys) in keys_by_body_row.iter().enumerate() {
            if keys.len() != rules.len() {
                return Err(Error::KeyArity {
                    row: BodyRowOffset::new(row),
                    expected: rules.len(),
                    actual: keys.len(),
                });
            }
        }

        if let Some(first) = keys_by_body_row.first() {
            for rule in 0..rules.len() {
                let expected = kind(&first[rule]);
                for (row, keys) in keys_by_body_row.iter().enumerate().skip(1) {
                    if kind(&keys[rule]) != expected {
                        return Err(Error::MixedScalarKinds {
                            row: BodyRowOffset::new(row),
                            rule,
                        });
                    }
                }
            }
        }

        let rows = keys_by_body_row.len();
        let mut sources_by_destination = Vec::new();
        sources_by_destination
            .try_reserve_exact(rows)
            .map_err(|_allocation| Error::Allocation { rows })?;
        sources_by_destination.extend((0..rows).map(BodyRowOffset::new));
        sources_by_destination.sort_unstable_by(|left, right| {
            let left_keys = &keys_by_body_row[left.get()];
            let right_keys = &keys_by_body_row[right.get()];
            for ((left_key, right_key), rule) in left_keys.iter().zip(right_keys).zip(rules) {
                let ordering = match rule.direction() {
                    super::Direction::Ascending => compare(left_key, right_key),
                    super::Direction::Descending => compare(left_key, right_key).reverse(),
                };
                if ordering != Ordering::Equal {
                    return ordering;
                }
            }
            left.cmp(right)
        });

        Ok(RowPermutation {
            sources_by_destination,
        })
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

    #[test]
    fn stable_row_planner_applies_priority_direction_and_source_ties() {
        use planning::stable_body_row_permutation;

        #[derive(Clone, Copy, PartialEq, Eq)]
        enum Kind {
            Number,
            Text,
        }

        #[derive(Clone, Copy)]
        enum Key<'a> {
            Number(i32),
            Text(&'a str),
        }

        let order = Order::new([
            Rule::new(ColumnIndex::new(0).unwrap(), Direction::Ascending),
            Rule::new(ColumnIndex::new(1).unwrap(), Direction::Descending),
        ])
        .unwrap();
        let keys = vec![
            vec![Key::Text("pear"), Key::Number(1)],
            vec![Key::Text("apple"), Key::Number(1)],
            vec![Key::Text("apple"), Key::Number(2)],
            vec![Key::Text("apple"), Key::Number(2)],
        ];
        let permutation = stable_body_row_permutation(
            &keys,
            &order,
            |value| match value {
                Key::Number(_) => Kind::Number,
                Key::Text(_) => Kind::Text,
            },
            |left, right| match (left, right) {
                (Key::Number(left), Key::Number(right)) => left.cmp(right),
                (Key::Text(left), Key::Text(right)) => left.cmp(right),
                _ => std::cmp::Ordering::Equal,
            },
        )
        .unwrap();

        assert_eq!(
            permutation
                .sources_by_destination()
                .iter()
                .map(|offset| offset.get())
                .collect::<Vec<_>>(),
            vec![2, 3, 1, 0]
        );
        assert_eq!(
            permutation
                .destinations_by_source()
                .unwrap()
                .iter()
                .map(|offset| offset.get())
                .collect::<Vec<_>>(),
            vec![3, 2, 0, 1]
        );
        assert!(permutation.changes_rows());
    }

    #[test]
    fn stable_row_planner_rejects_key_shape_and_kind_drift() {
        use planning::{Error as PlanError, stable_body_row_permutation};

        let order = Order::new([Rule::new(
            ColumnIndex::new(0).unwrap(),
            Direction::Ascending,
        )])
        .unwrap();
        let malformed = vec![vec![1_u8], vec![]];
        assert!(matches!(
            stable_body_row_permutation(&malformed, &order, |_| 0_u8, u8::cmp),
            Err(PlanError::KeyArity {
                row,
                expected: 1,
                actual: 0,
            }) if row.get() == 1
        ));

        let mixed = vec![vec![(0_u8, 3_i32)], vec![(1_u8, 2_i32)]];
        assert!(matches!(
            stable_body_row_permutation(
                &mixed,
                &order,
                |value| value.0,
                |left, right| left.1.cmp(&right.1),
            ),
            Err(PlanError::MixedScalarKinds { row, rule: 0 }) if row.get() == 1
        ));
    }

    #[test]
    fn stable_row_planner_preserves_exact_identity_and_empty_plans() {
        use planning::stable_body_row_permutation;

        let order = Order::new([Rule::new(
            ColumnIndex::new(0).unwrap(),
            Direction::Ascending,
        )])
        .unwrap();
        let empty: Vec<Vec<u8>> = Vec::new();
        let empty_plan = stable_body_row_permutation(&empty, &order, |_| (), u8::cmp).unwrap();
        assert!(empty_plan.is_empty());
        assert!(!empty_plan.changes_rows());

        let sorted = vec![vec![1_u8], vec![1], vec![2]];
        let plan = stable_body_row_permutation(&sorted, &order, |_| (), u8::cmp).unwrap();
        assert_eq!(plan.len(), 3);
        assert!(!plan.changes_rows());
        assert_eq!(
            plan.sources_by_destination(),
            &[
                planning::BodyRowOffset::new(0),
                planning::BodyRowOffset::new(1),
                planning::BodyRowOffset::new(2),
            ]
        );
    }
}
