//! Duration display values.

use std::fmt;

/// Selector-first exact-source transactions for one existing table cell's
/// explicit Duration display format.
pub mod transaction {
    pub use crate::package::table_cell_duration_format::{
        Commit, Diagnostics, Edit, Error, LimitKind, Patch, Path,
    };
}

/// Presentation style for a duration.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum Style {
    /// Display selected units as colon-separated fields.
    Colon,
    /// Display compact unit symbols.
    #[default]
    Abbreviated,
    /// Display complete unit names.
    FullNames,
}

/// A unit supported by a duration display.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub enum Unit {
    /// Weeks.
    Weeks,
    /// Days.
    Days,
    /// Hours.
    Hours,
    /// Minutes.
    Minutes,
    /// Seconds.
    Seconds,
    /// Milliseconds.
    Milliseconds,
}

impl Unit {
    /// Return the coarse-to-fine semantic rank of this unit.
    const fn rank(self) -> u8 {
        match self {
            Self::Weeks => 0,
            Self::Days => 1,
            Self::Hours => 2,
            Self::Minutes => 3,
            Self::Seconds => 4,
            Self::Milliseconds => 5,
        }
    }

    fn from_rank(rank: u8) -> Self {
        match rank {
            0 => Self::Weeks,
            1 => Self::Days,
            2 => Self::Hours,
            3 => Self::Minutes,
            4 => Self::Seconds,
            5 => Self::Milliseconds,
            _ => unreachable!(),
        }
    }
}

/// Errors returned by duration range construction.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Error {
    /// The largest unit is finer than the smallest unit.
    ReversedRange { largest: Unit, smallest: Unit },
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::ReversedRange { largest, smallest } => write!(
                formatter,
                "duration range is reversed: largest unit {largest:?} follows smallest unit {smallest:?}"
            ),
        }
    }
}

impl std::error::Error for Error {}

/// Result returned by checked duration constructors.
pub type Result<T> = std::result::Result<T, Error>;

/// Inclusive range of displayed duration units.
///
/// A range is always contiguous in coarse-to-fine order. It therefore needs
/// no heap storage, and iterating it can never yield more than six values.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct UnitRange {
    largest: Unit,
    smallest: Unit,
}

impl UnitRange {
    /// The complete range from weeks through milliseconds.
    pub const ALL: Self = Self {
        largest: Unit::Weeks,
        smallest: Unit::Milliseconds,
    };

    /// The common range from hours through milliseconds.
    pub const HOURS_TO_MILLISECONDS: Self = Self {
        largest: Unit::Hours,
        smallest: Unit::Milliseconds,
    };

    /// Validates and constructs a range from largest through smallest unit.
    ///
    /// # Errors
    ///
    /// Returns [`Error::ReversedRange`] when the arguments are not ordered
    /// from coarse to fine.
    pub const fn new(largest: Unit, smallest: Unit) -> Result<Self> {
        if largest.rank() > smallest.rank() {
            return Err(Error::ReversedRange { largest, smallest });
        }
        Ok(Self { largest, smallest })
    }

    /// The complete range from weeks through milliseconds.
    #[must_use]
    pub const fn all() -> Self {
        Self::ALL
    }

    /// The common range from hours through milliseconds.
    #[must_use]
    pub const fn hours_to_milliseconds() -> Self {
        Self::HOURS_TO_MILLISECONDS
    }

    /// Returns the largest displayed unit.
    #[must_use]
    pub const fn largest(self) -> Unit {
        self.largest
    }

    /// Returns the smallest displayed unit.
    #[must_use]
    pub const fn smallest(self) -> Unit {
        self.smallest
    }

    /// Return the inclusive number of units in this range.
    #[must_use]
    pub const fn len(self) -> usize {
        (self.smallest.rank() - self.largest.rank()) as usize + 1
    }

    /// Return whether this range contains no units.
    ///
    /// Validated inclusive ranges always contain both endpoints, so this is
    /// necessarily `false`.
    #[must_use]
    pub const fn is_empty(self) -> bool {
        false
    }

    /// Return whether this range contains the given unit.
    #[must_use]
    pub const fn contains(self, unit: Unit) -> bool {
        self.largest.rank() <= unit.rank() && unit.rank() <= self.smallest.rank()
    }

    /// Return an allocation-free iterator over units from largest to smallest.
    #[must_use]
    pub const fn iter(self) -> UnitRangeIter {
        UnitRangeIter {
            front: self.largest.rank(),
            back: self.smallest.rank() + 1,
        }
    }
}

impl Default for UnitRange {
    fn default() -> Self {
        Self::all()
    }
}

/// Allocation-free iterator returned by [`UnitRange::iter`].
///
/// The iterator is exact-size, double-ended, and fused. Its length is always
/// bounded by the six supported duration units.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct UnitRangeIter {
    front: u8,
    back: u8,
}

impl Iterator for UnitRangeIter {
    type Item = Unit;

    fn next(&mut self) -> Option<Self::Item> {
        if self.front == self.back {
            return None;
        }
        let rank = self.front;
        self.front += 1;
        Some(Unit::from_rank(rank))
    }

    fn size_hint(&self) -> (usize, Option<usize>) {
        let length = usize::from(self.back.saturating_sub(self.front));
        (length, Some(length))
    }
}

impl DoubleEndedIterator for UnitRangeIter {
    fn next_back(&mut self) -> Option<Self::Item> {
        if self.front == self.back {
            return None;
        }
        self.back -= 1;
        Some(Unit::from_rank(self.back))
    }
}

impl ExactSizeIterator for UnitRangeIter {
    fn len(&self) -> usize {
        usize::from(self.back.saturating_sub(self.front))
    }
}

impl std::iter::FusedIterator for UnitRangeIter {}

impl IntoIterator for UnitRange {
    type Item = Unit;
    type IntoIter = UnitRangeIter;

    fn into_iter(self) -> Self::IntoIter {
        self.iter()
    }
}

/// Automatic or explicit duration units.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Units {
    /// Let the spreadsheet select visible units while retaining its range.
    Automatic(UnitRange),
    /// Display exactly the selected range.
    Custom(UnitRange),
}

impl Units {
    /// Construct an automatic unit-selection policy for a validated range.
    #[must_use]
    pub const fn automatic(range: UnitRange) -> Self {
        Self::Automatic(range)
    }

    /// Construct a fixed unit-selection policy for a validated range.
    #[must_use]
    pub const fn custom(range: UnitRange) -> Self {
        Self::Custom(range)
    }

    /// Returns the persisted unit range.
    #[must_use]
    pub const fn range(self) -> UnitRange {
        match self {
            Self::Automatic(range) | Self::Custom(range) => range,
        }
    }

    /// Returns whether the spreadsheet selects visible units automatically.
    #[must_use]
    pub const fn is_automatic(self) -> bool {
        matches!(self, Self::Automatic(_))
    }

    /// Returns whether the selected unit range is fixed explicitly.
    #[must_use]
    pub const fn is_custom(self) -> bool {
        matches!(self, Self::Custom(_))
    }

    /// Replace the range while preserving automatic or custom selection.
    #[must_use]
    pub const fn with_range(self, range: UnitRange) -> Self {
        match self {
            Self::Automatic(_) => Self::Automatic(range),
            Self::Custom(_) => Self::Custom(range),
        }
    }
}

impl Default for Units {
    fn default() -> Self {
        Self::Automatic(UnitRange::all())
    }
}

/// Duration display format.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub struct Duration {
    style: Style,
    units: Units,
}

impl Duration {
    /// Constructs a duration format from a style and unit policy.
    #[must_use]
    pub const fn new(style: Style, units: Units) -> Self {
        Self { style, units }
    }

    /// Constructs an automatic-unit duration format.
    #[must_use]
    pub const fn automatic(style: Style) -> Self {
        Self::new(style, Units::automatic(UnitRange::all()))
    }

    /// Constructs a fixed-unit duration format.
    #[must_use]
    pub const fn custom(style: Style, range: UnitRange) -> Self {
        Self::new(style, Units::custom(range))
    }

    /// Returns the presentation style.
    #[must_use]
    pub const fn style(self) -> Style {
        self.style
    }

    /// Returns the automatic or fixed unit policy.
    #[must_use]
    pub const fn units(self) -> Units {
        self.units
    }

    /// Replace the presentation style.
    #[must_use]
    pub const fn with_style(mut self, style: Style) -> Self {
        self.style = style;
        self
    }

    /// Replace the automatic or fixed unit policy.
    #[must_use]
    pub const fn with_units(mut self, units: Units) -> Self {
        self.units = units;
        self
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn ranges_reject_reverse_order() {
        assert!(matches!(
            UnitRange::new(Unit::Seconds, Unit::Hours),
            Err(Error::ReversedRange {
                largest: Unit::Seconds,
                smallest: Unit::Hours,
            })
        ));
    }

    #[test]
    fn duration_values_round_trip_through_accessors() {
        let range = UnitRange::hours_to_milliseconds();
        let value = Duration::custom(Style::Abbreviated, range);
        assert_eq!(value.style(), Style::Abbreviated);
        assert_eq!(value.units(), Units::Custom(range));
        assert!(!value.units().is_automatic());
        assert_eq!(value.units().range(), range);
    }

    #[test]
    fn ranges_iterate_in_bounded_coarse_to_fine_order() {
        let range = UnitRange::new(Unit::Hours, Unit::Seconds).unwrap();
        assert_eq!(range.len(), 3);
        assert!(range.contains(Unit::Hours));
        assert!(range.contains(Unit::Minutes));
        assert!(range.contains(Unit::Seconds));
        assert!(!range.contains(Unit::Days));
        assert!(!range.contains(Unit::Milliseconds));

        let mut iterator = range.iter();
        assert_eq!(iterator.len(), 3);
        assert_eq!(iterator.size_hint(), (3, Some(3)));
        assert_eq!(iterator.next(), Some(Unit::Hours));
        assert_eq!(iterator.len(), 2);
        assert_eq!(iterator.next_back(), Some(Unit::Seconds));
        assert_eq!(iterator.size_hint(), (1, Some(1)));
        assert_eq!(iterator.next(), Some(Unit::Minutes));
        assert_eq!(iterator.next(), None);
        assert_eq!(iterator.next_back(), None);
        assert_eq!(iterator.len(), 0);

        let mut reverse = range.into_iter().rev();
        assert_eq!(reverse.next(), Some(Unit::Seconds));
        assert_eq!(reverse.next(), Some(Unit::Minutes));
        assert_eq!(reverse.next(), Some(Unit::Hours));
        assert_eq!(reverse.next(), None);
    }

    #[test]
    fn range_constants_are_canonical_and_bounded() {
        assert_eq!(
            UnitRange::all(),
            UnitRange::new(Unit::Weeks, Unit::Milliseconds).unwrap()
        );
        assert_eq!(
            UnitRange::hours_to_milliseconds(),
            UnitRange::new(Unit::Hours, Unit::Milliseconds).unwrap()
        );
        assert_eq!(UnitRange::all().len(), 6);
        assert_eq!(UnitRange::hours_to_milliseconds().len(), 4);
    }

    #[test]
    fn builders_preserve_the_other_duration_setting() {
        let range = UnitRange::new(Unit::Days, Unit::Minutes).unwrap();
        let value = Duration::automatic(Style::Colon)
            .with_style(Style::FullNames)
            .with_units(Units::custom(range));
        assert_eq!(value.style(), Style::FullNames);
        assert_eq!(value.units(), Units::Custom(range));

        let automatic = value.units().with_range(UnitRange::all());
        assert_eq!(automatic, Units::Custom(UnitRange::all()));
        assert!(automatic.is_custom());
    }
}
