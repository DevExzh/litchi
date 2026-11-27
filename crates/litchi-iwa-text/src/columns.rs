//! Archive-free text-frame column layouts.

use std::num::NonZeroU32;

/// Maximum number of columns in either supported layout.
///
/// The native representation stores every column after the first in a
/// repeated field.  Keeping the semantic value bounded prevents a malformed
/// or accidental authoring input from turning that field into an unbounded
/// allocation at the archive boundary.
pub const MAX_COLUMNS: u32 = 256;

/// Maximum number of columns in a variable-width layout.
pub const MAX_VARIABLE_COLUMNS: usize = MAX_COLUMNS as usize;

/// Why a text-column value was rejected.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Error {
    /// The equal-width column count was zero.
    ZeroCount,
    /// A layout exceeds [`MAX_COLUMNS`].
    TooManyColumns,
    /// The gap is not finite.
    GapNotFinite,
    /// The gap is negative.
    GapNegative,
    /// The variable-width column width is not finite.
    WidthNotFinite,
    /// The variable-width column width is not positive.
    WidthNonPositive,
    /// A variable-width layout has no column after its first column.
    MissingFollowing,
    /// A variable-width layout exceeds [`MAX_VARIABLE_COLUMNS`].
    TooManyVariableColumns,
}

impl std::fmt::Display for Error {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        let message = match self {
            Self::ZeroCount => "text column count must be non-zero",
            Self::TooManyColumns => "text column count exceeds the supported column limit",
            Self::GapNotFinite => "text column gap must be finite",
            Self::GapNegative => "text column gap must be non-negative",
            Self::WidthNotFinite => "text column width must be finite",
            Self::WidthNonPositive => "text column width must be positive",
            Self::MissingFollowing => "variable text columns must contain at least two columns",
            Self::TooManyVariableColumns => {
                "variable text columns exceed the supported column limit"
            },
        };
        formatter.write_str(message)
    }
}

impl std::error::Error for Error {}

/// Validated, non-zero number of equal-width text columns.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
#[repr(transparent)]
pub struct Count(NonZeroU32);

impl Count {
    /// The default single-column layout.
    pub const ONE: Self = Self(NonZeroU32::MIN);

    /// Validate an equal-width column count.
    ///
    /// # Errors
    ///
    /// Returns [`Error::ZeroCount`] for zero and [`Error::TooManyColumns`]
    /// when the count exceeds [`MAX_COLUMNS`].
    pub fn new(count: u32) -> Result<Self, Error> {
        if count > MAX_COLUMNS {
            return Err(Error::TooManyColumns);
        }
        NonZeroU32::new(count).map(Self).ok_or(Error::ZeroCount)
    }

    /// Return the native count value.
    #[must_use]
    pub const fn get(self) -> u32 {
        self.0.get()
    }
}

impl TryFrom<u32> for Count {
    type Error = Error;

    fn try_from(value: u32) -> Result<Self, Self::Error> {
        Self::new(value)
    }
}

impl From<Count> for u32 {
    fn from(value: Count) -> Self {
        value.get()
    }
}

/// Finite, non-negative distance between adjacent columns, in points.
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
#[repr(transparent)]
pub struct Gap(f32);

impl Gap {
    /// An explicit zero-point gap.
    pub const ZERO: Self = Self(0.0);

    /// Validate a point distance.
    ///
    /// # Errors
    ///
    /// Returns [`Error::GapNotFinite`] for non-finite values and
    /// [`Error::GapNegative`] for negative values, including negative zero.
    pub fn from_points(points: f32) -> Result<Self, Error> {
        if !points.is_finite() {
            return Err(Error::GapNotFinite);
        }
        if points < 0.0 || (points == 0.0 && points.is_sign_negative()) {
            return Err(Error::GapNegative);
        }
        Ok(Self(points))
    }

    /// Return the point distance.
    #[must_use]
    pub const fn points(self) -> f32 {
        self.0
    }
}

/// Finite, positive width of one variable-width column, in points.
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
#[repr(transparent)]
pub struct Width(f32);

impl Width {
    /// Validate a point width.
    ///
    /// # Errors
    ///
    /// Returns [`Error::WidthNotFinite`] for non-finite values and
    /// [`Error::WidthNonPositive`] for zero or negative values.
    pub fn from_points(points: f32) -> Result<Self, Error> {
        if !points.is_finite() {
            return Err(Error::WidthNotFinite);
        }
        if points <= 0.0 {
            return Err(Error::WidthNonPositive);
        }
        Ok(Self(points))
    }

    /// Return the point width.
    #[must_use]
    pub const fn points(self) -> f32 {
        self.0
    }
}

/// Equal-width column layout with an optional explicit inter-column gap.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Equal {
    count: Count,
    gap: Option<Gap>,
}

impl Equal {
    /// Construct an equal-width layout from validated values.
    #[must_use]
    pub const fn new(count: Count, gap: Option<Gap>) -> Self {
        Self { count, gap }
    }

    /// Return the number of columns.
    #[must_use]
    pub const fn count(self) -> Count {
        self.count
    }

    /// Return the explicit gap, or `None` for the native automatic gap.
    #[must_use]
    pub const fn gap(self) -> Option<Gap> {
        self.gap
    }
}

/// One variable-width column after the first, including its preceding gap.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Following {
    gap: Gap,
    width: Width,
}

impl Following {
    /// Construct one following column from validated values.
    #[must_use]
    pub const fn new(gap: Gap, width: Width) -> Self {
        Self { gap, width }
    }

    /// Return the preceding gap.
    #[must_use]
    pub const fn gap(self) -> Gap {
        self.gap
    }

    /// Return this column's width.
    #[must_use]
    pub const fn width(self) -> Width {
        self.width
    }
}

/// Variable-width column layout in first-plus-following order.
#[derive(Debug, Clone, PartialEq)]
pub struct Variable {
    first_width: Width,
    following: Box<[Following]>,
}

impl Variable {
    /// Construct a variable-width layout.
    ///
    /// The first column is represented separately to match the native
    /// layout, while the remaining columns share one compact boxed slice.
    ///
    /// # Errors
    ///
    /// Returns [`Error::MissingFollowing`] when no second column is supplied,
    /// or [`Error::TooManyVariableColumns`] when the layout exceeds the
    /// bounded column budget.
    pub fn new(first_width: Width, following: Vec<Following>) -> Result<Self, Error> {
        if following.is_empty() {
            return Err(Error::MissingFollowing);
        }
        if following.len().saturating_add(1) > MAX_VARIABLE_COLUMNS {
            return Err(Error::TooManyVariableColumns);
        }
        Ok(Self {
            first_width,
            following: following.into_boxed_slice(),
        })
    }

    /// Return the first column's width.
    #[must_use]
    pub const fn first_width(&self) -> Width {
        self.first_width
    }

    /// Borrow the following columns without copying them.
    #[must_use]
    pub fn following(&self) -> &[Following] {
        &self.following
    }

    /// Return the total number of columns.
    #[must_use]
    pub fn count(&self) -> usize {
        self.following.len().saturating_add(1)
    }
}

/// Equal-width or variable-width text-frame columns.
#[derive(Debug, Clone, PartialEq)]
pub enum Columns {
    /// All columns share one width determined by the text frame.
    Equal(Equal),
    /// Each column has an explicit width and preceding gap.
    Variable(Variable),
}

impl Columns {
    /// Construct an equal-width layout.
    #[must_use]
    pub const fn equal(count: Count, gap: Option<Gap>) -> Self {
        Self::Equal(Equal::new(count, gap))
    }

    /// Construct a variable-width layout.
    ///
    /// # Errors
    ///
    /// Returns the validation error from [`Variable::new`].
    pub fn variable(first_width: Width, following: Vec<Following>) -> Result<Self, Error> {
        Variable::new(first_width, following).map(Self::Variable)
    }

    /// Return the total number of columns.
    #[must_use]
    pub fn count(&self) -> usize {
        match self {
            Self::Equal(columns) => columns.count().get() as usize,
            Self::Variable(columns) => columns.count(),
        }
    }
}

impl Default for Columns {
    fn default() -> Self {
        Self::equal(Count::ONE, None)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn scalars_reject_invalid_values() {
        assert_eq!(Count::new(0), Err(Error::ZeroCount));
        assert_eq!(Count::new(MAX_COLUMNS + 1), Err(Error::TooManyColumns));
        assert_eq!(Gap::from_points(-0.01), Err(Error::GapNegative));
        assert_eq!(Gap::from_points(-0.0), Err(Error::GapNegative));
        assert_eq!(Gap::from_points(f32::NAN), Err(Error::GapNotFinite));
        assert_eq!(Width::from_points(0.0), Err(Error::WidthNonPositive));
        assert_eq!(
            Width::from_points(f32::INFINITY),
            Err(Error::WidthNotFinite)
        );
    }

    #[test]
    fn equal_and_variable_layouts_are_compact_and_round_trip_semantically() {
        let equal = Columns::equal(
            Count::new(3).unwrap(),
            Some(Gap::from_points(12.0).unwrap()),
        );
        assert_eq!(equal.count(), 3);

        let variable = Columns::variable(
            Width::from_points(72.0).unwrap(),
            vec![Following::new(
                Gap::from_points(8.0).unwrap(),
                Width::from_points(96.0).unwrap(),
            )],
        )
        .unwrap();
        assert_eq!(variable.count(), 2);
    }

    #[test]
    fn variable_layouts_are_bounded_and_need_a_following_column() {
        let width = Width::from_points(72.0).unwrap();
        assert_eq!(
            Variable::new(width, Vec::new()),
            Err(Error::MissingFollowing)
        );
        let following = vec![Following::new(Gap::ZERO, width); MAX_VARIABLE_COLUMNS];
        assert_eq!(
            Variable::new(width, following),
            Err(Error::TooManyVariableColumns)
        );
    }
}
