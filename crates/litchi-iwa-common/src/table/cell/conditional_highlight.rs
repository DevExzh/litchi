//! Archive-free conditional-highlight values for table cells.
//!
//! Native predicate identifiers, formula graphs, wire decoding, and package
//! mutation remain in the concrete iWork format owners. This module contains
//! only the validated semantic values exchanged at that boundary.

use std::num::NonZeroU32;

use chrono::NaiveDate;

use crate::color::Rgba;

const APPLE_EPOCH_YEAR: i32 = 2001;
const APPLE_EPOCH_MONTH: u32 = 1;
const APPLE_EPOCH_DAY: u32 = 1;
const SECONDS_PER_DAY: f64 = 86_400.0;

/// Validation failures for conditional-highlight values.
#[derive(Clone, Copy, Debug, PartialEq, Eq, thiserror::Error)]
pub enum Error {
    /// A numeric operand was not finite.
    #[error("conditional-highlight numbers must be finite")]
    NumberNonFinite,
    /// Numeric bounds were supplied in descending order.
    #[error("conditional-highlight range bounds must be ordered")]
    RangeReversed,
    /// A date's Apple-epoch seconds were not finite.
    #[error("conditional-highlight dates must be finite")]
    DateNonFinite,
    /// A date's Apple-epoch seconds were not representable with the required
    /// day-offset check.
    #[error("conditional-highlight date is out of range")]
    DateOutOfRange,
    /// A date was not aligned to midnight.
    #[error("conditional-highlight dates must be Apple-epoch midnight values")]
    DateNotMidnight,
    /// Gregorian date components did not form a valid date.
    #[error("conditional-highlight date is not valid")]
    DateInvalid,
    /// Date bounds were supplied in descending order.
    #[error("conditional-highlight date range bounds must be ordered")]
    DateRangeReversed,
    /// A relative-date period had a zero count.
    #[error("conditional-highlight date periods must be nonzero")]
    PeriodZero,
    /// A relative-date period could overflow its native month count.
    #[error("conditional-highlight date period is too large")]
    PeriodTooLarge,
    /// A text operand was empty.
    #[error("conditional-highlight text must not be empty")]
    TextEmpty,
    /// A style did not override any visual property.
    #[error("conditional-highlight style must override at least one property")]
    StyleEmpty,
}

/// Result type for conditional-highlight value construction.
pub type Result<T> = std::result::Result<T, Error>;

/// Finite numeric operand used by a comparison condition.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Number(f64);

impl Number {
    /// Construct a finite numeric operand.
    ///
    /// # Errors
    ///
    /// Returns [`Error::NumberNonFinite`] when `value` is NaN or infinite.
    pub fn new(value: f64) -> Result<Self> {
        if !value.is_finite() {
            return Err(Error::NumberNonFinite);
        }
        Ok(Self(value))
    }

    /// Return the numeric value.
    #[must_use]
    pub const fn get(self) -> f64 {
        self.0
    }
}

/// Ordered finite bounds used by a numeric comparison condition.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Range {
    lower: Number,
    upper: Number,
}

impl Range {
    /// Construct ordered inclusive bounds.
    ///
    /// # Errors
    ///
    /// Returns [`Error::RangeReversed`] when `lower` is greater than `upper`.
    pub fn new(lower: Number, upper: Number) -> Result<Self> {
        if lower.get() > upper.get() {
            return Err(Error::RangeReversed);
        }
        Ok(Self { lower, upper })
    }

    /// Return the lower bound.
    #[must_use]
    pub const fn lower(self) -> Number {
        self.lower
    }

    /// Return the upper bound.
    #[must_use]
    pub const fn upper(self) -> Number {
        self.upper
    }
}

/// Calendar date stored as whole seconds from Apple's 2001 epoch.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Date(f64);

impl Date {
    /// Construct a date from Apple-epoch seconds at midnight.
    ///
    /// # Errors
    ///
    /// Returns an error when the value is non-finite, outside the supported
    /// range, or not aligned to midnight.
    pub fn new(apple_seconds: f64) -> Result<Self> {
        if !apple_seconds.is_finite() {
            return Err(Error::DateNonFinite);
        }
        if !(apple_seconds + SECONDS_PER_DAY).is_finite() {
            return Err(Error::DateOutOfRange);
        }
        if apple_seconds.rem_euclid(SECONDS_PER_DAY) != 0.0 {
            return Err(Error::DateNotMidnight);
        }
        Ok(Self(apple_seconds))
    }

    /// Construct a date from Gregorian calendar components.
    ///
    /// # Errors
    ///
    /// Returns [`Error::DateInvalid`] when the components do not form a valid
    /// Gregorian date.
    pub fn from_ymd(year: i32, month: u32, day: u32) -> Result<Self> {
        let date = NaiveDate::from_ymd_opt(year, month, day).ok_or(Error::DateInvalid)?;
        let epoch = NaiveDate::from_ymd_opt(APPLE_EPOCH_YEAR, APPLE_EPOCH_MONTH, APPLE_EPOCH_DAY)
            .ok_or(Error::DateInvalid)?;
        #[allow(
            clippy::cast_precision_loss,
            reason = "The supported NaiveDate range is safely represented by epoch seconds"
        )]
        let days = date.signed_duration_since(epoch).num_days() as f64;
        Ok(Self(days * SECONDS_PER_DAY))
    }

    /// Return the Apple-epoch seconds.
    #[must_use]
    pub const fn apple_seconds(self) -> f64 {
        self.0
    }
}

/// Inclusive ordered calendar-date bounds used by a date-range condition.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct DateRange {
    lower: Date,
    upper: Date,
}

impl DateRange {
    /// Construct ordered inclusive bounds.
    ///
    /// # Errors
    ///
    /// Returns [`Error::DateRangeReversed`] when `lower` is later than `upper`.
    pub fn new(lower: Date, upper: Date) -> Result<Self> {
        if lower.apple_seconds() > upper.apple_seconds() {
            return Err(Error::DateRangeReversed);
        }
        Ok(Self { lower, upper })
    }

    /// Return the lower date.
    #[must_use]
    pub const fn lower(self) -> Date {
        self.lower
    }

    /// Return the upper date.
    #[must_use]
    pub const fn upper(self) -> Date {
        self.upper
    }
}

/// Calendar unit used by a relative-date condition.
#[repr(u8)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum PeriodUnit {
    /// Calendar days.
    Days,
    /// Seven-day calendar weeks.
    Weeks,
    /// Calendar months.
    Months,
    /// Three-calendar-month quarters.
    Quarters,
    /// Twelve-calendar-month years.
    Years,
}

impl PeriodUnit {
    const fn month_multiplier(self) -> Option<u32> {
        match self {
            Self::Days | Self::Weeks => None,
            Self::Months => Some(1),
            Self::Quarters => Some(3),
            Self::Years => Some(12),
        }
    }
}

/// Positive quantity and calendar unit used by a relative-date condition.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Period {
    count: NonZeroU32,
    unit: PeriodUnit,
}

impl Period {
    /// Construct a nonzero period whose native month quantity cannot overflow.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PeriodZero`] for a zero count or
    /// [`Error::PeriodTooLarge`] when the native month quantity would overflow.
    pub fn new(count: u32, unit: PeriodUnit) -> Result<Self> {
        let nonzero = NonZeroU32::new(count).ok_or(Error::PeriodZero)?;
        if unit
            .month_multiplier()
            .is_some_and(|multiplier| nonzero.get().checked_mul(multiplier).is_none())
        {
            return Err(Error::PeriodTooLarge);
        }
        Ok(Self {
            count: nonzero,
            unit,
        })
    }

    /// Return the positive period count.
    #[must_use]
    pub const fn count(self) -> u32 {
        self.count.get()
    }

    /// Return the calendar unit.
    #[must_use]
    pub const fn unit(self) -> PeriodUnit {
        self.unit
    }
}

/// Direction of an exact date offset relative to today.
#[repr(u8)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum OffsetDirection {
    /// Before today.
    Ago,
    /// After today.
    FromNow,
}

/// Exact relative date expressed as a period before or after today.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Offset {
    period: Period,
    direction: OffsetDirection,
}

impl Offset {
    /// Construct a relative-date offset.
    #[must_use]
    pub const fn new(period: Period, direction: OffsetDirection) -> Self {
        Self { period, direction }
    }

    /// Return the period.
    #[must_use]
    pub const fn period(self) -> Period {
        self.period
    }

    /// Return the direction.
    #[must_use]
    pub const fn direction(self) -> OffsetDirection {
        self.direction
    }
}

/// Non-empty text operand used by a text condition.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct Text(Box<str>);

impl Text {
    /// Construct a non-empty text operand.
    ///
    /// # Errors
    ///
    /// Returns [`Error::TextEmpty`] when `value` is empty.
    pub fn new(value: impl Into<Box<str>>) -> Result<Self> {
        let text = value.into();
        if text.is_empty() {
            return Err(Error::TextEmpty);
        }
        Ok(Self(text))
    }

    /// Borrow the text operand.
    #[must_use]
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

/// Condition evaluated against the cell carrying the rule.
#[derive(Debug, Clone, PartialEq)]
pub enum Condition {
    CellIsBlank,
    CellIsNotBlank,
    CheckboxIsChecked,
    CheckboxIsNotChecked,
    BooleanIsTrue,
    BooleanIsFalse,
    NumberIsPositive,
    NumberIsNegative,
    DateIsToday,
    DateIsYesterday,
    DateIsTomorrow,
    DateIs(Date),
    DateIsBefore(Date),
    DateIsAfter(Date),
    DateIsBetween(DateRange),
    DateIsInNext(Period),
    DateIsInLast(Period),
    DateIsOffsetFromToday(Offset),
    EqualTo(Number),
    NotEqualTo(Number),
    GreaterThan(Number),
    GreaterThanOrEqualTo(Number),
    LessThan(Number),
    LessThanOrEqualTo(Number),
    Between(Range),
    NotBetween(Range),
    TextEqualTo(Text),
    TextNotEqualTo(Text),
    TextStartsWith(Text),
    TextDoesNotStartWith(Text),
    TextEndsWith(Text),
    TextDoesNotEndWith(Text),
    TextContains(Text),
    TextDoesNotContain(Text),
}

impl Condition {
    /// Return the operand for a single-number condition.
    #[must_use]
    pub const fn single_operand(&self) -> Option<Number> {
        match self {
            Self::EqualTo(value)
            | Self::NotEqualTo(value)
            | Self::GreaterThan(value)
            | Self::GreaterThanOrEqualTo(value)
            | Self::LessThan(value)
            | Self::LessThanOrEqualTo(value) => Some(*value),
            Self::CellIsBlank
            | Self::CellIsNotBlank
            | Self::CheckboxIsChecked
            | Self::CheckboxIsNotChecked
            | Self::BooleanIsTrue
            | Self::BooleanIsFalse
            | Self::NumberIsPositive
            | Self::NumberIsNegative
            | Self::DateIsToday
            | Self::DateIsYesterday
            | Self::DateIsTomorrow
            | Self::DateIs(_)
            | Self::DateIsBefore(_)
            | Self::DateIsAfter(_)
            | Self::DateIsBetween(_)
            | Self::DateIsInNext(_)
            | Self::DateIsInLast(_)
            | Self::DateIsOffsetFromToday(_)
            | Self::Between(_)
            | Self::NotBetween(_)
            | Self::TextEqualTo(_)
            | Self::TextNotEqualTo(_)
            | Self::TextStartsWith(_)
            | Self::TextDoesNotStartWith(_)
            | Self::TextEndsWith(_)
            | Self::TextDoesNotEndWith(_)
            | Self::TextContains(_)
            | Self::TextDoesNotContain(_) => None,
        }
    }

    /// Return the range for a two-number condition.
    #[must_use]
    pub const fn range(&self) -> Option<Range> {
        match self {
            Self::Between(range) | Self::NotBetween(range) => Some(*range),
            Self::CellIsBlank
            | Self::CellIsNotBlank
            | Self::CheckboxIsChecked
            | Self::CheckboxIsNotChecked
            | Self::BooleanIsTrue
            | Self::BooleanIsFalse
            | Self::NumberIsPositive
            | Self::NumberIsNegative
            | Self::DateIsToday
            | Self::DateIsYesterday
            | Self::DateIsTomorrow
            | Self::DateIs(_)
            | Self::DateIsBefore(_)
            | Self::DateIsAfter(_)
            | Self::DateIsBetween(_)
            | Self::DateIsInNext(_)
            | Self::DateIsInLast(_)
            | Self::DateIsOffsetFromToday(_)
            | Self::EqualTo(_)
            | Self::NotEqualTo(_)
            | Self::GreaterThan(_)
            | Self::GreaterThanOrEqualTo(_)
            | Self::LessThan(_)
            | Self::LessThanOrEqualTo(_)
            | Self::TextEqualTo(_)
            | Self::TextNotEqualTo(_)
            | Self::TextStartsWith(_)
            | Self::TextDoesNotStartWith(_)
            | Self::TextEndsWith(_)
            | Self::TextDoesNotEndWith(_)
            | Self::TextContains(_)
            | Self::TextDoesNotContain(_) => None,
        }
    }

    /// Borrow the text operand for a text condition.
    #[must_use]
    pub fn text(&self) -> Option<&Text> {
        match self {
            Self::TextEqualTo(value)
            | Self::TextNotEqualTo(value)
            | Self::TextStartsWith(value)
            | Self::TextDoesNotStartWith(value)
            | Self::TextEndsWith(value)
            | Self::TextDoesNotEndWith(value)
            | Self::TextContains(value)
            | Self::TextDoesNotContain(value) => Some(value),
            Self::CellIsBlank
            | Self::CellIsNotBlank
            | Self::CheckboxIsChecked
            | Self::CheckboxIsNotChecked
            | Self::BooleanIsTrue
            | Self::BooleanIsFalse
            | Self::NumberIsPositive
            | Self::NumberIsNegative
            | Self::DateIsToday
            | Self::DateIsYesterday
            | Self::DateIsTomorrow
            | Self::DateIs(_)
            | Self::DateIsBefore(_)
            | Self::DateIsAfter(_)
            | Self::DateIsBetween(_)
            | Self::DateIsInNext(_)
            | Self::DateIsInLast(_)
            | Self::DateIsOffsetFromToday(_)
            | Self::EqualTo(_)
            | Self::NotEqualTo(_)
            | Self::GreaterThan(_)
            | Self::GreaterThanOrEqualTo(_)
            | Self::LessThan(_)
            | Self::LessThanOrEqualTo(_)
            | Self::Between(_)
            | Self::NotBetween(_) => None,
        }
    }

    /// Return the date for a single-date condition.
    #[must_use]
    pub const fn date(&self) -> Option<Date> {
        match self {
            Self::DateIs(value) | Self::DateIsBefore(value) | Self::DateIsAfter(value) => {
                Some(*value)
            },
            Self::CellIsBlank
            | Self::CellIsNotBlank
            | Self::CheckboxIsChecked
            | Self::CheckboxIsNotChecked
            | Self::BooleanIsTrue
            | Self::BooleanIsFalse
            | Self::NumberIsPositive
            | Self::NumberIsNegative
            | Self::DateIsToday
            | Self::DateIsYesterday
            | Self::DateIsTomorrow
            | Self::DateIsBetween(_)
            | Self::DateIsInNext(_)
            | Self::DateIsInLast(_)
            | Self::DateIsOffsetFromToday(_)
            | Self::EqualTo(_)
            | Self::NotEqualTo(_)
            | Self::GreaterThan(_)
            | Self::GreaterThanOrEqualTo(_)
            | Self::LessThan(_)
            | Self::LessThanOrEqualTo(_)
            | Self::Between(_)
            | Self::NotBetween(_)
            | Self::TextEqualTo(_)
            | Self::TextNotEqualTo(_)
            | Self::TextStartsWith(_)
            | Self::TextDoesNotStartWith(_)
            | Self::TextEndsWith(_)
            | Self::TextDoesNotEndWith(_)
            | Self::TextContains(_)
            | Self::TextDoesNotContain(_) => None,
        }
    }

    /// Return the date range for a date-range condition.
    #[must_use]
    pub const fn date_range(&self) -> Option<DateRange> {
        match self {
            Self::DateIsBetween(range) => Some(*range),
            Self::CellIsBlank
            | Self::CellIsNotBlank
            | Self::CheckboxIsChecked
            | Self::CheckboxIsNotChecked
            | Self::BooleanIsTrue
            | Self::BooleanIsFalse
            | Self::NumberIsPositive
            | Self::NumberIsNegative
            | Self::DateIsToday
            | Self::DateIsYesterday
            | Self::DateIsTomorrow
            | Self::DateIs(_)
            | Self::DateIsBefore(_)
            | Self::DateIsAfter(_)
            | Self::DateIsInNext(_)
            | Self::DateIsInLast(_)
            | Self::DateIsOffsetFromToday(_)
            | Self::EqualTo(_)
            | Self::NotEqualTo(_)
            | Self::GreaterThan(_)
            | Self::GreaterThanOrEqualTo(_)
            | Self::LessThan(_)
            | Self::LessThanOrEqualTo(_)
            | Self::Between(_)
            | Self::NotBetween(_)
            | Self::TextEqualTo(_)
            | Self::TextNotEqualTo(_)
            | Self::TextStartsWith(_)
            | Self::TextDoesNotStartWith(_)
            | Self::TextEndsWith(_)
            | Self::TextDoesNotEndWith(_)
            | Self::TextContains(_)
            | Self::TextDoesNotContain(_) => None,
        }
    }

    /// Return the period for a relative-date condition.
    #[must_use]
    pub const fn date_period(&self) -> Option<Period> {
        match self {
            Self::DateIsInNext(period) | Self::DateIsInLast(period) => Some(*period),
            Self::CellIsBlank
            | Self::CellIsNotBlank
            | Self::CheckboxIsChecked
            | Self::CheckboxIsNotChecked
            | Self::BooleanIsTrue
            | Self::BooleanIsFalse
            | Self::NumberIsPositive
            | Self::NumberIsNegative
            | Self::DateIsToday
            | Self::DateIsYesterday
            | Self::DateIsTomorrow
            | Self::DateIs(_)
            | Self::DateIsBefore(_)
            | Self::DateIsAfter(_)
            | Self::DateIsBetween(_)
            | Self::DateIsOffsetFromToday(_)
            | Self::EqualTo(_)
            | Self::NotEqualTo(_)
            | Self::GreaterThan(_)
            | Self::GreaterThanOrEqualTo(_)
            | Self::LessThan(_)
            | Self::LessThanOrEqualTo(_)
            | Self::Between(_)
            | Self::NotBetween(_)
            | Self::TextEqualTo(_)
            | Self::TextNotEqualTo(_)
            | Self::TextStartsWith(_)
            | Self::TextDoesNotStartWith(_)
            | Self::TextEndsWith(_)
            | Self::TextDoesNotEndWith(_)
            | Self::TextContains(_)
            | Self::TextDoesNotContain(_) => None,
        }
    }

    /// Return the offset for an exact relative-date condition.
    #[must_use]
    pub const fn date_offset(&self) -> Option<Offset> {
        match self {
            Self::DateIsOffsetFromToday(offset) => Some(*offset),
            Self::CellIsBlank
            | Self::CellIsNotBlank
            | Self::CheckboxIsChecked
            | Self::CheckboxIsNotChecked
            | Self::BooleanIsTrue
            | Self::BooleanIsFalse
            | Self::NumberIsPositive
            | Self::NumberIsNegative
            | Self::DateIsToday
            | Self::DateIsYesterday
            | Self::DateIsTomorrow
            | Self::DateIs(_)
            | Self::DateIsBefore(_)
            | Self::DateIsAfter(_)
            | Self::DateIsBetween(_)
            | Self::DateIsInNext(_)
            | Self::DateIsInLast(_)
            | Self::EqualTo(_)
            | Self::NotEqualTo(_)
            | Self::GreaterThan(_)
            | Self::GreaterThanOrEqualTo(_)
            | Self::LessThan(_)
            | Self::LessThanOrEqualTo(_)
            | Self::Between(_)
            | Self::NotBetween(_)
            | Self::TextEqualTo(_)
            | Self::TextNotEqualTo(_)
            | Self::TextStartsWith(_)
            | Self::TextDoesNotStartWith(_)
            | Self::TextEndsWith(_)
            | Self::TextDoesNotEndWith(_)
            | Self::TextContains(_)
            | Self::TextDoesNotContain(_) => None,
        }
    }
}

/// Visual overrides applied when a condition matches.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Style {
    fill: Option<Rgba>,
    text_color: Option<Rgba>,
    bold: bool,
}

impl Style {
    /// Construct a style that overrides at least one visual property.
    ///
    /// # Errors
    ///
    /// Returns [`Error::StyleEmpty`] when no visual property is overridden.
    pub fn new(fill: Option<Rgba>, text_color: Option<Rgba>, bold: bool) -> Result<Self> {
        if fill.is_none() && text_color.is_none() && !bold {
            return Err(Error::StyleEmpty);
        }
        Ok(Self {
            fill,
            text_color,
            bold,
        })
    }

    /// Construct a fill-only style.
    #[must_use]
    pub const fn with_fill(fill: Rgba) -> Self {
        Self {
            fill: Some(fill),
            text_color: None,
            bold: false,
        }
    }

    /// Construct a text-color-only style.
    #[must_use]
    pub const fn with_text_color(text_color: Rgba) -> Self {
        Self {
            fill: None,
            text_color: Some(text_color),
            bold: false,
        }
    }

    /// Return the optional fill override.
    #[must_use]
    pub const fn fill(self) -> Option<Rgba> {
        self.fill
    }

    /// Return the optional text-color override.
    #[must_use]
    pub const fn text_color(self) -> Option<Rgba> {
        self.text_color
    }

    /// Return whether matching text is bold.
    #[must_use]
    pub const fn bold(self) -> bool {
        self.bold
    }
}

/// One ordered condition and its visual style.
#[derive(Debug, Clone, PartialEq)]
pub struct Rule {
    /// The condition evaluated against the cell.
    pub condition: Condition,
    /// The visual overrides applied on a match.
    pub style: Style,
}

impl Rule {
    /// Construct one conditional-highlight rule.
    #[must_use]
    pub const fn new(condition: Condition, style: Style) -> Self {
        Self { condition, style }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::color::{RgbColorSpace, Rgba};

    #[test]
    fn numeric_operands_and_styles_reject_empty_or_non_finite_values() {
        assert!(Number::new(f64::NAN).is_err());
        assert!(Number::new(f64::INFINITY).is_err());
        assert!(Style::new(None, None, false).is_err());
        let lower = Number::new(7.0).unwrap_or_else(|error| panic!("valid number: {error}"));
        let upper = Number::new(3.0).unwrap_or_else(|error| panic!("valid number: {error}"));
        assert!(Range::new(lower, upper).is_err());
        assert!(Text::new("").is_err());
        let text = Text::new("Grain").unwrap_or_else(|error| panic!("valid text: {error}"));
        assert_eq!(text.as_str(), "Grain");

        let color = Rgba::new(0.9, 0.2, 0.1, 1.0, RgbColorSpace::Srgb)
            .unwrap_or_else(|error| panic!("valid color: {error}"));
        assert_eq!(Style::with_fill(color).fill(), Some(color));
    }

    #[allow(
        clippy::float_cmp,
        reason = "The epoch conversion is expected to preserve this exact day value"
    )]
    #[test]
    fn date_operands_are_midnight_aligned_and_ranges_are_ordered() {
        let date =
            Date::from_ymd(2026, 7, 27).unwrap_or_else(|error| panic!("valid date: {error}"));
        assert_eq!(date.apple_seconds(), 806_803_200.0);
        assert_eq!(
            Date::new(date.apple_seconds()).unwrap_or_else(|error| panic!("valid date: {error}")),
            date
        );
        assert!(Date::from_ymd(2026, 2, 29).is_err());
        assert!(Date::new(f64::NAN).is_err());
        assert!(Date::new(date.apple_seconds() + 1.0).is_err());

        let earlier =
            Date::from_ymd(2026, 7, 26).unwrap_or_else(|error| panic!("valid date: {error}"));
        assert!(DateRange::new(earlier, date).is_ok());
        assert!(DateRange::new(date, earlier).is_err());
    }

    #[test]
    fn date_periods_are_positive_typed_and_overflow_checked() {
        assert!(Period::new(0, PeriodUnit::Days).is_err());
        assert!(Period::new(u32::MAX, PeriodUnit::Years).is_err());
        let period = Period::new(3, PeriodUnit::Quarters)
            .unwrap_or_else(|error| panic!("valid period: {error}"));
        assert_eq!(period.count(), 3);
        assert_eq!(period.unit(), PeriodUnit::Quarters);

        let offset = Offset::new(period, OffsetDirection::Ago);
        assert_eq!(offset.period(), period);
        assert_eq!(offset.direction(), OffsetDirection::Ago);
    }
}
