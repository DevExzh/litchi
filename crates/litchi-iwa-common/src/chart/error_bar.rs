//! Archive-free semantic values for native chart error bars.
//!
//! The IWA adapter owns protobuf decoding, native field numbers, graph
//! lookup, and package mutation. This module owns only the checked values
//! exchanged at that boundary. Unknown native direction and value-type
//! discriminants remain available for a lossless read/write cycle.

use std::fmt;

const NATIVE_DIRECTION_BOTH: i32 = 1;
const NATIVE_DIRECTION_POSITIVE: i32 = 2;
const NATIVE_DIRECTION_NEGATIVE: i32 = 3;

const NATIVE_FIXED_VALUE: i32 = 1;
const NATIVE_PERCENTAGE: i32 = 2;
const NATIVE_STANDARD_DEVIATION: i32 = 3;
const NATIVE_STANDARD_ERROR: i32 = 4;
const NATIVE_CUSTOM_VALUES: i32 = 5;

const DEFAULT_FIXED_VALUE: f32 = 10.0;
const DEFAULT_PERCENTAGE: u8 = 5;
const DEFAULT_STANDARD_DEVIATIONS: u32 = 1;
const MAX_EXACT_F32_INTEGER: u32 = 1_u32 << f32::MANTISSA_DIGITS;

/// Maximum number of custom magnitudes retained for either side of a series.
pub const MAX_CUSTOM_VALUES_PER_SIDE: usize = 1_000_000;

/// Validation failures for chart error-bar values.
#[derive(Clone, Copy, Debug, PartialEq, Eq, thiserror::Error)]
pub enum Error {
    /// A fixed error magnitude was NaN or infinite.
    #[error("chart fixed error-bar value must be finite")]
    FixedValueNonFinite,
    /// A fixed error magnitude was zero or negative.
    #[error("chart fixed error-bar value must be positive")]
    FixedValueNonPositive,
    /// A percentage was outside the native inclusive domain.
    #[error("chart error-bar percentage must be in {minimum}..={maximum}, got {value}")]
    PercentageOutOfRange {
        /// Supplied percentage.
        value: u8,
        /// Smallest accepted percentage.
        minimum: u8,
        /// Largest accepted percentage.
        maximum: u8,
    },
    /// A standard-deviation count was outside the exactly representable
    /// native `f32` integer domain.
    #[error(
        "chart error-bar standard-deviation count must be in {minimum}..={maximum}, got {value}"
    )]
    StandardDeviationCountOutOfRange {
        /// Supplied count.
        value: u32,
        /// Smallest accepted count.
        minimum: u32,
        /// Largest accepted count.
        maximum: u32,
    },
    /// A custom magnitude was NaN or infinite.
    #[error("custom chart error-bar value must be finite")]
    CustomValueNonFinite,
    /// A custom magnitude was negative.
    #[error("custom chart error-bar value must be nonnegative")]
    CustomValueNegative,
    /// A custom-value side exceeded its bounded semantic budget.
    #[error("custom chart error-bar {side} values exceed {maximum}; observed at least {count}")]
    TooManyCustomValues {
        /// Side whose values exceeded the budget.
        side: Side,
        /// Number of values observed or promised by a lower bound.
        count: usize,
        /// Maximum accepted values on one side.
        maximum: usize,
    },
    /// The combined custom-value allocation could not be reserved.
    #[error("custom chart error-bar values could not be allocated")]
    AllocationFailed,
    /// A known native value type was supplied through the unknown form.
    #[error("known chart error-bar value type {native_value} must use its named variant")]
    KnownValueTypeAsUnknown {
        /// Known native value-type discriminant.
        native_value: i32,
    },
}

/// Result type for chart error-bar value construction.
pub type Result<T> = std::result::Result<T, Error>;

/// The side whose custom values are being validated.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum Side {
    /// Positive error magnitudes.
    Positive,
    /// Negative error magnitudes.
    Negative,
}

impl fmt::Display for Side {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::Positive => "positive",
            Self::Negative => "negative",
        })
    }
}

/// Which side of each data point receives an error bar.
//
// The native discriminant is stored directly so known and future values have
// the same four-byte layout and unknown values can be written without a
// lossy enum projection.
#[repr(transparent)]
#[derive(Clone, Copy, PartialEq, Eq, Hash)]
pub struct Direction(i32);

/// The recognized native error-bar directions.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum Kind {
    /// Draw error bars above and below each value.
    PositiveAndNegative,
    /// Draw only the positive error bar.
    PositiveOnly,
    /// Draw only the negative error bar.
    NegativeOnly,
}

impl Direction {
    /// Draw error bars above and below each value.
    #[allow(
        non_upper_case_globals,
        reason = "enum-style associated constants are the ergonomic focused API"
    )]
    pub const PositiveAndNegative: Self = Self(NATIVE_DIRECTION_BOTH);
    /// Draw only the positive error bar.
    #[allow(
        non_upper_case_globals,
        reason = "enum-style associated constants are the ergonomic focused API"
    )]
    pub const PositiveOnly: Self = Self(NATIVE_DIRECTION_POSITIVE);
    /// Draw only the negative error bar.
    #[allow(
        non_upper_case_globals,
        reason = "enum-style associated constants are the ergonomic focused API"
    )]
    pub const NegativeOnly: Self = Self(NATIVE_DIRECTION_NEGATIVE);

    /// Decode the native direction without discarding future values.
    #[must_use]
    pub const fn from_native(value: i32) -> Self {
        Self(value)
    }

    /// Return the native direction discriminant.
    #[must_use]
    pub const fn native_value(self) -> i32 {
        self.0
    }

    /// Return the recognized direction, if this is a known native value.
    #[must_use]
    pub const fn kind(self) -> Option<Kind> {
        match self.0 {
            NATIVE_DIRECTION_BOTH => Some(Kind::PositiveAndNegative),
            NATIVE_DIRECTION_POSITIVE => Some(Kind::PositiveOnly),
            NATIVE_DIRECTION_NEGATIVE => Some(Kind::NegativeOnly),
            _ => None,
        }
    }

    /// Whether this value is an unrecognized native direction.
    #[must_use]
    pub const fn is_unsupported(self) -> bool {
        self.kind().is_none()
    }
}

impl fmt::Debug for Direction {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self.kind() {
            Some(kind) => kind.fmt(formatter),
            None => formatter.debug_tuple("Unsupported").field(&self.0).finish(),
        }
    }
}

/// A validated positive finite fixed error magnitude.
#[repr(transparent)]
#[derive(Clone, Copy, Debug, PartialEq, PartialOrd)]
pub struct FixedValue(f32);

impl FixedValue {
    /// Native inspector default.
    pub const DEFAULT: Self = Self(DEFAULT_FIXED_VALUE);

    /// Validate and construct a fixed error magnitude.
    ///
    /// # Errors
    ///
    /// Returns [`Error::FixedValueNonFinite`] for NaN or infinity, and
    /// [`Error::FixedValueNonPositive`] for zero or negative values.
    #[must_use = "use the validated fixed value or handle its validation error"]
    pub fn new(value: f32) -> Result<Self> {
        if !value.is_finite() {
            return Err(Error::FixedValueNonFinite);
        }
        if value <= 0.0 {
            return Err(Error::FixedValueNonPositive);
        }
        Ok(Self(value))
    }

    /// Return the fixed error magnitude.
    #[must_use]
    pub const fn value(self) -> f32 {
        self.0
    }
}

impl Default for FixedValue {
    fn default() -> Self {
        Self::DEFAULT
    }
}

impl TryFrom<f32> for FixedValue {
    type Error = Error;

    fn try_from(value: f32) -> Result<Self> {
        Self::new(value)
    }
}

/// A validated percentage accepted by the iWork error-bar inspector.
#[repr(transparent)]
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub struct Percentage(u8);

impl Percentage {
    /// Smallest accepted percentage.
    pub const MINIMUM: u8 = 1;
    /// Largest accepted percentage.
    pub const MAXIMUM: u8 = 100;
    /// Native inspector default.
    pub const DEFAULT: Self = Self(DEFAULT_PERCENTAGE);

    /// Validate and construct an error percentage.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PercentageOutOfRange`] when `value` is outside the
    /// inclusive native percentage range.
    #[must_use = "use the validated percentage or handle its validation error"]
    pub const fn new(value: u8) -> Result<Self> {
        if value < Self::MINIMUM || value > Self::MAXIMUM {
            return Err(Error::PercentageOutOfRange {
                value,
                minimum: Self::MINIMUM,
                maximum: Self::MAXIMUM,
            });
        }
        Ok(Self(value))
    }

    /// Return the percentage as an integer.
    #[must_use]
    pub const fn value(self) -> u8 {
        self.0
    }
}

impl Default for Percentage {
    fn default() -> Self {
        Self::DEFAULT
    }
}

impl TryFrom<u8> for Percentage {
    type Error = Error;

    fn try_from(value: u8) -> Result<Self> {
        Self::new(value)
    }
}

/// A validated integral count used by standard-deviation error bars.
#[repr(transparent)]
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub struct StandardDeviationCount(u32);

impl StandardDeviationCount {
    /// Smallest accepted count.
    pub const MINIMUM: u32 = 1;
    /// Largest consecutive integer represented exactly by native `f32`.
    pub const MAXIMUM: u32 = MAX_EXACT_F32_INTEGER;
    /// Native inspector default.
    pub const DEFAULT: Self = Self(DEFAULT_STANDARD_DEVIATIONS);

    /// Validate and construct a standard-deviation count.
    ///
    /// # Errors
    ///
    /// Returns [`Error::StandardDeviationCountOutOfRange`] when `value` is
    /// outside the native exactly representable integer range.
    #[must_use = "use the validated count or handle its validation error"]
    pub const fn new(value: u32) -> Result<Self> {
        if value < Self::MINIMUM || value > Self::MAXIMUM {
            return Err(Error::StandardDeviationCountOutOfRange {
                value,
                minimum: Self::MINIMUM,
                maximum: Self::MAXIMUM,
            });
        }
        Ok(Self(value))
    }

    /// Return the number of standard deviations.
    #[must_use]
    pub const fn value(self) -> u32 {
        self.0
    }
}

impl Default for StandardDeviationCount {
    fn default() -> Self {
        Self::DEFAULT
    }
}

impl TryFrom<u32> for StandardDeviationCount {
    type Error = Error;

    fn try_from(value: u32) -> Result<Self> {
        Self::new(value)
    }
}

/// One validated finite, nonnegative custom error magnitude.
#[repr(transparent)]
#[derive(Clone, Copy, Debug, PartialEq, PartialOrd)]
pub struct CustomValue(f64);

impl CustomValue {
    /// Validate and construct a custom error magnitude.
    ///
    /// # Errors
    ///
    /// Returns [`Error::CustomValueNonFinite`] for NaN or infinity, and
    /// [`Error::CustomValueNegative`] for negative values.
    #[must_use = "use the validated custom value or handle its validation error"]
    pub fn new(value: f64) -> Result<Self> {
        if !value.is_finite() {
            return Err(Error::CustomValueNonFinite);
        }
        if value < 0.0 {
            return Err(Error::CustomValueNegative);
        }
        Ok(Self(value))
    }

    /// Return the custom error magnitude.
    #[must_use]
    pub const fn value(self) -> f64 {
        self.0
    }
}

impl TryFrom<f64> for CustomValue {
    type Error = Error;

    fn try_from(value: f64) -> Result<Self> {
        Self::new(value)
    }
}

/// A checked, lossless unrecognized native error-bar value type.
#[repr(transparent)]
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub struct Unknown(i32);

impl Unknown {
    /// Construct an unrecognized native value-type discriminant.
    ///
    /// # Errors
    ///
    /// Returns [`Error::KnownValueTypeAsUnknown`] when `native_value` is a
    /// recognized native value type.
    #[must_use = "use the validated unknown type or handle its validation error"]
    pub const fn new(native_value: i32) -> Result<Self> {
        if matches!(
            native_value,
            NATIVE_FIXED_VALUE
                | NATIVE_PERCENTAGE
                | NATIVE_STANDARD_DEVIATION
                | NATIVE_STANDARD_ERROR
                | NATIVE_CUSTOM_VALUES
        ) {
            return Err(Error::KnownValueTypeAsUnknown { native_value });
        }
        Ok(Self(native_value))
    }

    /// Return the unrecognized native value-type discriminant.
    #[must_use]
    pub const fn native_value(self) -> i32 {
        self.0
    }
}

impl TryFrom<i32> for Unknown {
    type Error = Error;

    fn try_from(value: i32) -> Result<Self> {
        Self::new(value)
    }
}

/// Per-point positive and negative custom error magnitudes.
//
// Both sides share one boxed slice. The positive prefix length is kept
// separately, so the semantic value uses one allocation instead of one per
// side while preserving borrowed positive/negative views.
#[derive(Clone, Debug, PartialEq)]
pub struct CustomValues {
    values: Box<[CustomValue]>,
    positive_len: usize,
}

impl CustomValues {
    /// Construct and validate custom positive and negative magnitudes.
    ///
    /// # Errors
    ///
    /// Returns a scalar validation error for an invalid magnitude,
    /// [`Error::TooManyCustomValues`] when either side exceeds its bounded
    /// semantic budget, or [`Error::AllocationFailed`] when storage cannot
    /// be reserved.
    #[must_use = "use the validated custom values or handle their validation error"]
    pub fn new(
        positive: impl IntoIterator<Item = f64>,
        negative: impl IntoIterator<Item = f64>,
    ) -> Result<Self> {
        Self::from_results(
            positive.into_iter().map(CustomValue::new),
            negative.into_iter().map(CustomValue::new),
        )
    }

    /// Combine already-validated custom values without copying their scalar
    /// payloads into separate side allocations.
    ///
    /// # Errors
    ///
    /// Returns [`Error::TooManyCustomValues`] when either side exceeds its
    /// bounded semantic budget, or [`Error::AllocationFailed`] when storage
    /// cannot be reserved.
    #[must_use = "use the combined custom values or handle their validation error"]
    pub fn from_validated(
        positive: impl IntoIterator<Item = CustomValue>,
        negative: impl IntoIterator<Item = CustomValue>,
    ) -> Result<Self> {
        Self::from_results(
            positive.into_iter().map(Ok::<CustomValue, Error>),
            negative.into_iter().map(Ok::<CustomValue, Error>),
        )
    }

    /// Positive custom magnitudes in native point order.
    #[must_use]
    pub fn positive(&self) -> &[CustomValue] {
        &self.values[..self.positive_len]
    }

    /// Negative custom magnitudes in native point order.
    #[must_use]
    pub fn negative(&self) -> &[CustomValue] {
        &self.values[self.positive_len..]
    }

    /// Whether both custom-value sides are empty.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.values.is_empty()
    }

    fn from_results<P, N>(positive: P, negative: N) -> Result<Self>
    where
        P: IntoIterator<Item = Result<CustomValue>>,
        N: IntoIterator<Item = Result<CustomValue>>,
    {
        let mut positive_iter = positive.into_iter();
        let mut negative_iter = negative.into_iter();
        let positive_lower = positive_iter.size_hint().0;
        let negative_lower = negative_iter.size_hint().0;
        check_lower_bound(positive_lower, Side::Positive)?;
        check_lower_bound(negative_lower, Side::Negative)?;

        let initial_capacity = positive_lower
            .checked_add(negative_lower)
            .ok_or(Error::AllocationFailed)?;
        let mut values = Vec::new();
        values
            .try_reserve_exact(initial_capacity)
            .map_err(|_allocation_error| Error::AllocationFailed)?;
        append_results(&mut values, &mut positive_iter, Side::Positive)?;
        let positive_len = values.len();
        append_results(&mut values, &mut negative_iter, Side::Negative)?;
        Ok(Self {
            values: values.into_boxed_slice(),
            positive_len,
        })
    }
}

impl Default for CustomValues {
    fn default() -> Self {
        Self {
            values: Box::new([]),
            positive_len: 0,
        }
    }
}

/// Complete semantic error-bar configuration for one native chart series.
#[derive(Clone, Debug, PartialEq, Default)]
#[non_exhaustive]
pub enum Series {
    /// Do not draw error bars.
    #[default]
    None,
    /// Use one fixed magnitude for every point.
    FixedValue {
        /// Sides on which bars are drawn.
        direction: Direction,
        /// Fixed error magnitude.
        value: FixedValue,
    },
    /// Derive each error from a percentage of its point value.
    Percentage {
        /// Sides on which bars are drawn.
        direction: Direction,
        /// Percentage of each point value.
        percentage: Percentage,
    },
    /// Derive errors from an integral number of standard deviations.
    StandardDeviation {
        /// Sides on which bars are drawn.
        direction: Direction,
        /// Number of standard deviations.
        deviations: StandardDeviationCount,
    },
    /// Derive errors from the standard error.
    StandardError {
        /// Sides on which bars are drawn.
        direction: Direction,
    },
    /// Use custom positive and negative magnitudes in native point order.
    CustomValues {
        /// Sides on which bars are drawn.
        direction: Direction,
        /// Per-point magnitudes.
        values: CustomValues,
    },
    /// Preserve an unrecognized native value-derivation type.
    Unsupported {
        /// Sides on which bars are drawn.
        direction: Direction,
        /// Unrecognized native type discriminant.
        native_type: Unknown,
    },
}

impl Series {
    /// Native inspector defaults for fixed-value errors.
    #[must_use]
    pub const fn fixed_value(direction: Direction) -> Self {
        Self::FixedValue {
            direction,
            value: FixedValue::DEFAULT,
        }
    }

    /// Native inspector defaults for percentage errors.
    #[must_use]
    pub const fn percentage(direction: Direction) -> Self {
        Self::Percentage {
            direction,
            percentage: Percentage::DEFAULT,
        }
    }

    /// Native inspector defaults for standard-deviation errors.
    #[must_use]
    pub const fn standard_deviation(direction: Direction) -> Self {
        Self::StandardDeviation {
            direction,
            deviations: StandardDeviationCount::DEFAULT,
        }
    }

    /// Construct standard-error settings with no type-specific value.
    #[must_use]
    pub const fn standard_error(direction: Direction) -> Self {
        Self::StandardError { direction }
    }

    /// Construct custom-value settings.
    #[must_use]
    pub const fn custom_values(direction: Direction, values: CustomValues) -> Self {
        Self::CustomValues { direction, values }
    }

    /// Construct settings for an unrecognized native value type.
    ///
    /// # Errors
    ///
    /// Returns [`Error::KnownValueTypeAsUnknown`] when `native_type` is a
    /// recognized native value type.
    #[must_use = "use the validated error-bar series or handle its validation error"]
    pub fn unsupported(direction: Direction, native_type: i32) -> Result<Self> {
        Ok(Self::Unsupported {
            direction,
            native_type: Unknown::new(native_type)?,
        })
    }

    /// Return the direction for visible settings, or `None` when hidden.
    #[must_use]
    pub const fn direction(&self) -> Option<Direction> {
        match self {
            Self::None => None,
            Self::FixedValue { direction, .. }
            | Self::Percentage { direction, .. }
            | Self::StandardDeviation { direction, .. }
            | Self::StandardError { direction }
            | Self::CustomValues { direction, .. }
            | Self::Unsupported { direction, .. } => Some(*direction),
        }
    }

    /// Return the native value-type discriminant for visible settings.
    #[must_use]
    pub const fn native_type(&self) -> Option<i32> {
        match self {
            Self::None => None,
            Self::FixedValue { .. } => Some(NATIVE_FIXED_VALUE),
            Self::Percentage { .. } => Some(NATIVE_PERCENTAGE),
            Self::StandardDeviation { .. } => Some(NATIVE_STANDARD_DEVIATION),
            Self::StandardError { .. } => Some(NATIVE_STANDARD_ERROR),
            Self::CustomValues { .. } => Some(NATIVE_CUSTOM_VALUES),
            Self::Unsupported { native_type, .. } => Some(native_type.native_value()),
        }
    }

    /// Validate the lossless unknown discriminant before an adapter writes it.
    ///
    /// # Errors
    ///
    /// Returns [`Error::KnownValueTypeAsUnknown`] if an unsupported variant
    /// contains a recognized native value type.
    pub fn validate(&self) -> Result<()> {
        if let Self::Unsupported { native_type, .. } = self {
            Unknown::new(native_type.native_value())?;
        }
        Ok(())
    }
}

fn check_lower_bound(lower: usize, side: Side) -> Result<()> {
    if lower > MAX_CUSTOM_VALUES_PER_SIDE {
        return Err(Error::TooManyCustomValues {
            side,
            count: lower,
            maximum: MAX_CUSTOM_VALUES_PER_SIDE,
        });
    }
    Ok(())
}

fn append_results<I>(values: &mut Vec<CustomValue>, iterator: &mut I, side: Side) -> Result<()>
where
    I: Iterator<Item = Result<CustomValue>>,
{
    for (side_len, value) in iterator.enumerate() {
        if side_len == MAX_CUSTOM_VALUES_PER_SIDE {
            return Err(Error::TooManyCustomValues {
                side,
                count: side_len.saturating_add(1),
                maximum: MAX_CUSTOM_VALUES_PER_SIDE,
            });
        }
        if values.len() == values.capacity() {
            values
                .try_reserve(1)
                .map_err(|_allocation_error| Error::AllocationFailed)?;
        }
        values.push(value?);
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use std::mem::{align_of, size_of};

    use super::{
        CustomValue, CustomValues, Direction, Error, FixedValue, Kind, MAX_CUSTOM_VALUES_PER_SIDE,
        Percentage, Series, Side, StandardDeviationCount, Unknown,
    };

    #[test]
    fn scalar_values_are_compact_and_strictly_validated() {
        assert_eq!(size_of::<Direction>(), 4);
        assert_eq!(align_of::<Direction>(), 4);
        assert_eq!(size_of::<FixedValue>(), 4);
        assert_eq!(size_of::<Percentage>(), 1);
        assert_eq!(size_of::<StandardDeviationCount>(), 4);
        assert_eq!(size_of::<CustomValue>(), 8);
        assert_eq!(size_of::<Unknown>(), 4);
        assert_eq!(FixedValue::new(0.0), Err(Error::FixedValueNonPositive));
        assert_eq!(FixedValue::new(f32::NAN), Err(Error::FixedValueNonFinite));
        assert_eq!(
            Percentage::new(0),
            Err(Error::PercentageOutOfRange {
                value: 0,
                minimum: Percentage::MINIMUM,
                maximum: Percentage::MAXIMUM,
            })
        );
        assert!(StandardDeviationCount::new(0).is_err());
        assert_eq!(
            CustomValue::new(f64::INFINITY),
            Err(Error::CustomValueNonFinite)
        );
        assert_eq!(CustomValue::new(-1.0), Err(Error::CustomValueNegative));
    }

    #[test]
    fn direction_and_unknown_type_preserve_native_values_without_aliases() {
        assert_eq!(
            Direction::PositiveAndNegative.kind(),
            Some(Kind::PositiveAndNegative)
        );
        assert_eq!(Direction::from_native(9_001).native_value(), 9_001);
        assert!(Direction::from_native(9_001).is_unsupported());
        assert_eq!(Unknown::new(9_001).unwrap().native_value(), 9_001);
        assert_eq!(
            Unknown::new(1),
            Err(Error::KnownValueTypeAsUnknown { native_value: 1 })
        );
        assert!(Series::unsupported(Direction::PositiveOnly, 2).is_err());
        let unknown = Series::unsupported(Direction::from_native(9_001), 9_001).unwrap();
        assert_eq!(unknown.direction(), Some(Direction::from_native(9_001)));
        assert_eq!(unknown.native_type(), Some(9_001));
    }

    #[test]
    fn custom_values_use_one_allocation_and_keep_side_order() {
        let values = CustomValues::new([1.0, 2.0], [0.5, 1.5]).unwrap();
        assert_eq!(
            values.positive(),
            [
                CustomValue::new(1.0).unwrap(),
                CustomValue::new(2.0).unwrap()
            ]
        );
        assert_eq!(
            values.negative(),
            [
                CustomValue::new(0.5).unwrap(),
                CustomValue::new(1.5).unwrap()
            ]
        );
        assert!(!values.is_empty());
        assert_eq!(
            size_of::<CustomValues>(),
            size_of::<Box<[CustomValue]>>() + size_of::<usize>()
        );
        assert_eq!(size_of::<Series>(), 32);
    }

    #[test]
    fn custom_value_limits_are_checked_before_large_lower_bound_allocations() {
        let too_many = std::iter::repeat_n(1.0, MAX_CUSTOM_VALUES_PER_SIDE + 1);
        assert_eq!(
            CustomValues::new(too_many, []),
            Err(Error::TooManyCustomValues {
                side: Side::Positive,
                count: MAX_CUSTOM_VALUES_PER_SIDE + 1,
                maximum: MAX_CUSTOM_VALUES_PER_SIDE,
            })
        );
    }
}
