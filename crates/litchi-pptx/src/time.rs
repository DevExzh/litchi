//! Checked PowerPoint 2010 universal time offsets.
//!
//! [`Offset`] implements the `p14:ST_UniversalTimeOffset` grammar from
//! `[MS-PPTX]` section 2.3.4.6. Values are stored as canonical decimal
//! milliseconds, so equivalent source spellings compare, order, and hash by
//! the represented duration rather than by their original unit or spelling.

use std::cmp::Ordering;
use std::fmt;
use std::str::FromStr;
use std::time::Duration;

/// Maximum accepted byte length of a unit-bearing producer spelling.
///
/// The Microsoft grammar is unbounded. A finite limit is required when parsing
/// untrusted package XML. This preserves the existing 64-byte event-codec
/// limit for producer spellings that carry a unit.
pub const MAX_BYTES: usize = 64;

/// Maximum size of the normalized millisecond representation.
///
/// Unit conversion can add at most a small number of decimal zeroes. The wider
/// canonical bound lets every accepted 64-byte value serialize and reparse;
/// extended input is accepted only for unitless canonical milliseconds.
pub const MAX_CANONICAL_BYTES: usize = MAX_BYTES + 16;

/// A unit accepted by PowerPoint's universal-time-offset grammar.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
#[repr(u8)]
pub enum Unit {
    Hour,
    Minute,
    Second,
    Millisecond,
    Microsecond,
    Nanosecond,
}

impl Unit {
    /// Return the unit suffix used by `[MS-PPTX]`.
    pub const fn suffix(self) -> &'static str {
        match self {
            Self::Hour => "h",
            Self::Minute => "min",
            Self::Second => "s",
            Self::Millisecond => "ms",
            Self::Microsecond => "µs",
            Self::Nanosecond => "ns",
        }
    }

    /// Return `(multiplier, decimal exponent)` in milliseconds.
    const fn millisecond_parts(self) -> (u8, i32) {
        match self {
            Self::Hour => (36, 5),
            Self::Minute => (6, 4),
            Self::Second => (1, 3),
            Self::Millisecond => (1, 0),
            Self::Microsecond => (1, -3),
            Self::Nanosecond => (1, -6),
        }
    }

    fn from_suffix(value: &str) -> Result<Self, ParseError> {
        match value {
            "h" => Ok(Self::Hour),
            "min" => Ok(Self::Minute),
            "s" => Ok(Self::Second),
            "ms" => Ok(Self::Millisecond),
            "µs" => Ok(Self::Microsecond),
            "ns" => Ok(Self::Nanosecond),
            _ => Err(ParseError::InvalidUnit),
        }
    }
}

impl fmt::Display for Unit {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(self.suffix())
    }
}

impl FromStr for Unit {
    type Err = ParseError;

    fn from_str(value: &str) -> Result<Self, Self::Err> {
        Self::from_suffix(value)
    }
}

/// Failure to parse or losslessly convert a universal time offset.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ParseError {
    /// No lexical value was supplied.
    Empty,
    /// The lexical or normalized representation exceeded a safety limit.
    TooLong { len: usize, max: usize },
    /// The numeric component did not match `1*DIGIT ["." 1*DIGIT]`.
    InvalidNumber,
    /// The optional unit was not one of the six Microsoft-defined units.
    InvalidUnit,
    /// The exact value cannot be represented by [`Duration`].
    NotRepresentableAsDuration,
}

impl fmt::Display for ParseError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Empty => formatter.write_str("universal time offset is empty"),
            Self::TooLong { len, max } => write!(
                formatter,
                "universal time offset uses {len} bytes; the safety limit is {max}"
            ),
            Self::InvalidNumber => formatter.write_str(
                "universal time offset number must match DIGIT+ with an optional non-empty fractional part",
            ),
            Self::InvalidUnit => formatter.write_str(
                "universal time offset unit must be h, min, s, ms, µs, or ns",
            ),
            Self::NotRepresentableAsDuration => formatter.write_str(
                "universal time offset cannot be represented exactly by std::time::Duration",
            ),
        }
    }
}

impl std::error::Error for ParseError {}

#[derive(Clone, PartialEq, Eq, Hash)]
enum Repr {
    Zero,
    /// Canonical decimal milliseconds with no suffix, redundant leading zeroes,
    /// or trailing fractional zeroes.
    Decimal(Box<str>),
}

/// A non-negative, exact PowerPoint universal time offset.
///
/// The internal representation is private so its storage can evolve without
/// exposing lexical strings in the semantic API. [`Self::as_str`] returns a
/// canonical, specification-valid millisecond spelling for codecs and logs.
#[derive(Clone, PartialEq, Eq, Hash)]
#[must_use]
pub struct Offset(Repr);

impl Offset {
    /// Zero milliseconds.
    pub const ZERO: Self = Self(Repr::Zero);

    /// Construct an integral value in a Microsoft-defined unit.
    pub fn new(value: u64, unit: Unit) -> Self {
        if value == 0 {
            return Self::ZERO;
        }
        let (multiplier, exponent) = unit.millisecond_parts();
        Self::from_digits(value.to_string(), multiplier, exponent)
    }

    /// Construct an integral millisecond value.
    #[inline]
    pub fn ms(value: u64) -> Self {
        Self::new(value, Unit::Millisecond)
    }

    /// Construct an integral hour value.
    #[inline]
    pub fn hours(value: u64) -> Self {
        Self::new(value, Unit::Hour)
    }

    /// Construct an integral minute value.
    #[inline]
    pub fn mins(value: u64) -> Self {
        Self::new(value, Unit::Minute)
    }

    /// Construct an integral second value.
    #[inline]
    pub fn secs(value: u64) -> Self {
        Self::new(value, Unit::Second)
    }

    /// Construct an integral microsecond value.
    #[inline]
    pub fn micros(value: u64) -> Self {
        Self::new(value, Unit::Microsecond)
    }

    /// Construct an integral nanosecond value.
    #[inline]
    pub fn nanos(value: u64) -> Self {
        Self::new(value, Unit::Nanosecond)
    }

    /// Parse a decimal number paired with an explicit unit.
    pub fn decimal(value: &str, unit: Unit) -> Result<Self, ParseError> {
        if value.is_empty() {
            return Err(ParseError::Empty);
        }
        let len = value
            .len()
            .checked_add(unit.suffix().len())
            .ok_or(ParseError::TooLong {
                len: usize::MAX,
                max: MAX_BYTES,
            })?;
        if len > MAX_BYTES {
            return Err(ParseError::TooLong {
                len,
                max: MAX_BYTES,
            });
        }
        let mut lexical = String::with_capacity(len);
        lexical.push_str(value);
        lexical.push_str(unit.suffix());
        Self::try_from(lexical)
    }

    /// Parse an `[MS-PPTX]` universal-time-offset lexical value.
    #[inline]
    pub fn parse(value: &str) -> Result<Self, ParseError> {
        value.parse()
    }

    /// Return whether this value is zero.
    #[inline]
    pub const fn is_zero(&self) -> bool {
        matches!(self.0, Repr::Zero)
    }

    /// Return the canonical decimal-millisecond spelling.
    pub fn as_str(&self) -> &str {
        match &self.0 {
            Repr::Zero => "0",
            Repr::Decimal(value) => value,
        }
    }

    /// Convert to [`Duration`] when the value has nanosecond precision and is
    /// within `Duration`'s range.
    #[inline]
    pub fn duration(&self) -> Result<Duration, ParseError> {
        Duration::try_from(self)
    }

    fn parse_owned(mut value: String) -> Result<Self, ParseError> {
        if value.is_empty() {
            return Err(ParseError::Empty);
        }
        let source_len = value.len();
        if source_len > MAX_CANONICAL_BYTES {
            return Err(ParseError::TooLong {
                len: source_len,
                max: MAX_CANONICAL_BYTES,
            });
        }

        let unit_start = value
            .find(|character: char| !character.is_ascii_digit() && character != '.')
            .unwrap_or(value.len());
        let suffix = &value[unit_start..];
        let unit = if suffix.is_empty() {
            Unit::Millisecond
        } else {
            Unit::from_suffix(suffix)?
        };
        if source_len > MAX_BYTES && !suffix.is_empty() {
            return Err(ParseError::TooLong {
                len: source_len,
                max: MAX_BYTES,
            });
        }
        // Unit-bearing spellings are already rejected above. For an extended
        // unitless spelling, retain only whether the original bytes were
        // canonical instead of cloning the source just for a later equality
        // check. The parser is bounded, but this path is still exercised for
        // every producer value over MAX_BYTES.
        let extended_source_is_canonical =
            source_len <= MAX_BYTES || is_canonical_millisecond_spelling(&value);
        value.truncate(unit_start);

        let (fraction_digits, decimal_index) = validate_number(&value)?;
        if let Some(index) = decimal_index {
            value.remove(index);
        }

        let leading_zeroes = value.bytes().take_while(|byte| *byte == b'0').count();
        if leading_zeroes == value.len() {
            return enforce_extended_canonical(
                Self::ZERO,
                source_len,
                extended_source_is_canonical,
            );
        }
        if leading_zeroes != 0 {
            value.drain(..leading_zeroes);
        }

        let (multiplier, unit_exponent) = unit.millisecond_parts();
        let fraction_digits =
            i32::try_from(fraction_digits).map_err(|_| ParseError::InvalidNumber)?;
        let exponent = unit_exponent
            .checked_sub(fraction_digits)
            .ok_or(ParseError::InvalidNumber)?;
        let offset = Self::from_digits(value, multiplier, exponent);
        if offset.as_str().len() > MAX_CANONICAL_BYTES {
            return Err(ParseError::TooLong {
                len: offset.as_str().len(),
                max: MAX_CANONICAL_BYTES,
            });
        }
        enforce_extended_canonical(offset, source_len, extended_source_is_canonical)
    }

    fn from_digits(mut digits: String, multiplier: u8, mut exponent: i32) -> Self {
        normalize_trailing_zeroes(&mut digits, &mut exponent);
        if multiplier != 1 {
            digits = multiply_decimal_digits(&digits, multiplier);
            normalize_trailing_zeroes(&mut digits, &mut exponent);
        }
        Self(Repr::Decimal(format_decimal_milliseconds(digits, exponent)))
    }
}

impl Default for Offset {
    fn default() -> Self {
        Self::ZERO
    }
}

impl fmt::Debug for Offset {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_tuple("Offset")
            .field(&self.as_str())
            .finish()
    }
}

impl fmt::Display for Offset {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(self.as_str())
    }
}

impl FromStr for Offset {
    type Err = ParseError;

    fn from_str(value: &str) -> Result<Self, Self::Err> {
        if value.len() > MAX_CANONICAL_BYTES {
            return Err(ParseError::TooLong {
                len: value.len(),
                max: MAX_CANONICAL_BYTES,
            });
        }
        Self::try_from(value.to_owned())
    }
}

impl TryFrom<String> for Offset {
    type Error = ParseError;

    fn try_from(value: String) -> Result<Self, Self::Error> {
        Self::parse_owned(value)
    }
}

impl Ord for Offset {
    fn cmp(&self, other: &Self) -> Ordering {
        let (left_whole, left_fraction) = decimal_parts(self.as_str());
        let (right_whole, right_fraction) = decimal_parts(other.as_str());
        left_whole
            .len()
            .cmp(&right_whole.len())
            .then_with(|| left_whole.cmp(right_whole))
            .then_with(|| compare_fraction(left_fraction, right_fraction))
    }
}

impl PartialOrd for Offset {
    fn partial_cmp(&self, other: &Self) -> Option<Ordering> {
        Some(self.cmp(other))
    }
}

impl From<Duration> for Offset {
    fn from(value: Duration) -> Self {
        let seconds = value.as_secs();
        let nanoseconds = value.subsec_nanos();
        if seconds == 0 && nanoseconds == 0 {
            return Self::ZERO;
        }

        if seconds == 0 {
            let mut digits = nanoseconds.to_string();
            let mut exponent = -6;
            normalize_trailing_zeroes(&mut digits, &mut exponent);
            return Self(Repr::Decimal(format_decimal_milliseconds(digits, exponent)));
        }

        let mut digits = seconds.to_string();
        let nanos = nanoseconds.to_string();
        for _ in nanos.len()..9 {
            digits.push('0');
        }
        digits.push_str(&nanos);
        let mut exponent = -6;
        normalize_trailing_zeroes(&mut digits, &mut exponent);
        Self(Repr::Decimal(format_decimal_milliseconds(digits, exponent)))
    }
}

impl TryFrom<&Offset> for Duration {
    type Error = ParseError;

    fn try_from(value: &Offset) -> Result<Self, Self::Error> {
        let (whole, fraction) = decimal_parts(value.as_str());
        if fraction.len() > 6 {
            return Err(ParseError::NotRepresentableAsDuration);
        }

        let mut coefficient = whole
            .parse::<u128>()
            .map_err(|_| ParseError::NotRepresentableAsDuration)?;
        for digit in fraction.bytes() {
            coefficient = coefficient
                .checked_mul(10)
                .and_then(|current| current.checked_add(u128::from(digit - b'0')))
                .ok_or(ParseError::NotRepresentableAsDuration)?;
        }
        for _ in fraction.len()..6 {
            coefficient = coefficient
                .checked_mul(10)
                .ok_or(ParseError::NotRepresentableAsDuration)?;
        }

        let seconds = coefficient / 1_000_000_000;
        let nanoseconds = coefficient % 1_000_000_000;
        let seconds = u64::try_from(seconds).map_err(|_| ParseError::NotRepresentableAsDuration)?;
        let nanoseconds =
            u32::try_from(nanoseconds).map_err(|_| ParseError::NotRepresentableAsDuration)?;
        Ok(Duration::new(seconds, nanoseconds))
    }
}

impl TryFrom<Offset> for Duration {
    type Error = ParseError;

    fn try_from(value: Offset) -> Result<Self, Self::Error> {
        Self::try_from(&value)
    }
}

fn validate_number(value: &str) -> Result<(usize, Option<usize>), ParseError> {
    if value.is_empty() {
        return Err(ParseError::InvalidNumber);
    }
    let mut decimal_index = None;
    for (index, byte) in value.bytes().enumerate() {
        match byte {
            b'0'..=b'9' => {},
            b'.' if decimal_index.is_none() => decimal_index = Some(index),
            _ => return Err(ParseError::InvalidNumber),
        }
    }
    if decimal_index == Some(0) || decimal_index == value.len().checked_sub(1) {
        return Err(ParseError::InvalidNumber);
    }
    let fraction_digits = decimal_index.map_or(0, |index| value.len() - index - 1);
    Ok((fraction_digits, decimal_index))
}

fn enforce_extended_canonical(
    offset: Offset,
    source_len: usize,
    source_is_canonical: bool,
) -> Result<Offset, ParseError> {
    if source_len > MAX_BYTES && !source_is_canonical {
        return Err(ParseError::TooLong {
            len: source_len,
            max: MAX_BYTES,
        });
    }
    Ok(offset)
}

/// Return whether a validated unitless value already has the canonical
/// millisecond spelling emitted by [`Offset::as_str`].
///
/// This deliberately checks lexical shape only. The numeric grammar is still
/// validated by [`validate_number`], preserving its error precedence for
/// malformed extended input while avoiding an owned copy of the source.
fn is_canonical_millisecond_spelling(value: &str) -> bool {
    if value == "0" {
        return true;
    }

    let (whole, fraction) = value
        .split_once('.')
        .map_or((value, None), |(whole, fraction)| (whole, Some(fraction)));
    if whole.is_empty()
        || (whole != "0" && whole.starts_with('0'))
        || !whole.bytes().all(|byte| byte.is_ascii_digit())
    {
        return false;
    }

    let Some(fraction) = fraction else {
        return true;
    };
    !fraction.is_empty()
        && fraction.bytes().all(|byte| byte.is_ascii_digit())
        && !fraction.ends_with('0')
}

fn normalize_trailing_zeroes(digits: &mut String, exponent: &mut i32) {
    while digits.ends_with('0') {
        digits.pop();
        *exponent = exponent.saturating_add(1);
    }
}

fn multiply_decimal_digits(digits: &str, multiplier: u8) -> String {
    let mut reversed = Vec::with_capacity(digits.len() + 2);
    let mut carry = 0u16;
    for digit in digits.bytes().rev() {
        let value = u16::from(digit - b'0') * u16::from(multiplier) + carry;
        reversed.push((value % 10) as u8);
        carry = value / 10;
    }
    while carry != 0 {
        reversed.push((carry % 10) as u8);
        carry /= 10;
    }
    let mut output = String::with_capacity(reversed.len());
    for digit in reversed.into_iter().rev() {
        output.push(char::from(b'0' + digit));
    }
    output
}

fn format_decimal_milliseconds(mut digits: String, exponent: i32) -> Box<str> {
    if exponent >= 0 {
        for _ in 0..exponent {
            digits.push('0');
        }
        return digits.into_boxed_str();
    }

    let fractional_places = usize::try_from(exponent.unsigned_abs()).unwrap_or(usize::MAX);
    if fractional_places < digits.len() {
        let index = digits.len() - fractional_places;
        digits.insert(index, '.');
        return digits.into_boxed_str();
    }

    let zeroes = fractional_places.saturating_sub(digits.len());
    let mut output = String::with_capacity(digits.len().saturating_add(2).saturating_add(zeroes));
    output.push_str("0.");
    for _ in 0..zeroes {
        output.push('0');
    }
    output.push_str(&digits);
    output.into_boxed_str()
}

fn decimal_parts(value: &str) -> (&str, &str) {
    value.split_once('.').unwrap_or((value, ""))
}

fn compare_fraction(left: &str, right: &str) -> Ordering {
    left.bytes()
        .chain(std::iter::repeat(b'0'))
        .zip(right.bytes().chain(std::iter::repeat(b'0')))
        .take(left.len().max(right.len()))
        .find_map(|(left, right)| (left != right).then(|| left.cmp(&right)))
        .unwrap_or(Ordering::Equal)
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::collections::HashSet;

    #[test]
    fn equivalent_units_normalize_for_equality_order_and_hashing() {
        let values = ["1s", "1000", "1000ms", "1000000µs", "1000000000ns"]
            .map(|value| Offset::parse(value).unwrap());
        assert!(values.windows(2).all(|pair| pair[0] == pair[1]));
        assert_eq!(values[0].as_str(), "1000");
        assert_eq!(values.into_iter().collect::<HashSet<_>>().len(), 1);

        assert_eq!(
            Offset::parse("1h").unwrap(),
            Offset::parse("60min").unwrap()
        );
        assert_eq!(Offset::parse("1h").unwrap().as_str(), "3600000");
        assert!(Offset::parse("999.9ms").unwrap() < Offset::parse("1s").unwrap());
    }

    #[test]
    fn grammar_and_fractional_precision_are_exact() {
        assert_eq!(Offset::parse("001.500s").unwrap().as_str(), "1500");
        assert_eq!(
            Offset::parse("0.0000001ns").unwrap().as_str(),
            "0.0000000000001"
        );
        assert_eq!(
            Offset::decimal("1.25", Unit::Minute).unwrap().as_str(),
            "75000"
        );

        for invalid in ["", ".1s", "1.", "1..2s", "-1s", "+1s", "1e3", "1MS"] {
            assert!(Offset::parse(invalid).is_err(), "accepted {invalid:?}");
        }
        assert!("".parse::<Unit>().is_err());
        assert_eq!(Offset::parse("1").unwrap(), Offset::ms(1));
        assert!(matches!(
            Offset::parse(&"1".repeat(MAX_CANONICAL_BYTES + 1)),
            Err(ParseError::TooLong { .. })
        ));
        assert!(matches!(
            Offset::decimal(&"1".repeat(MAX_BYTES), Unit::Second),
            Err(ParseError::TooLong { .. })
        ));

        let largest_lexical = format!("{}h", "9".repeat(MAX_BYTES - 1));
        let smallest_lexical = format!("0.{}1ns", "0".repeat(MAX_BYTES - 5));
        for lexical in [largest_lexical, smallest_lexical] {
            assert_eq!(lexical.len(), MAX_BYTES);
            let offset = Offset::parse(&lexical).unwrap();
            assert!(offset.as_str().len() <= MAX_CANONICAL_BYTES);
            assert_eq!(Offset::parse(offset.as_str()).unwrap(), offset);
        }

        let maximum_canonical = "9".repeat(MAX_CANONICAL_BYTES);
        assert_eq!(
            Offset::parse(&maximum_canonical).unwrap().as_str(),
            maximum_canonical
        );
        assert!(matches!(
            Offset::parse(&format!("{}h", "9".repeat(MAX_CANONICAL_BYTES - 1))),
            Err(ParseError::TooLong { .. })
        ));
        assert!(matches!(
            Offset::parse(&format!("0{}", "9".repeat(MAX_BYTES))),
            Err(ParseError::TooLong { .. })
        ));
    }

    #[test]
    fn extended_unitless_values_keep_only_canonical_spellings() {
        let canonical = format!("1{}", "0".repeat(MAX_BYTES));
        assert_eq!(canonical.len(), MAX_BYTES + 1);
        assert_eq!(Offset::parse(&canonical).unwrap().as_str(), canonical);

        let canonical_fraction = format!("0.{}1", "0".repeat(MAX_BYTES - 2));
        assert_eq!(canonical_fraction.len(), MAX_BYTES + 1);
        assert_eq!(
            Offset::parse(&canonical_fraction).unwrap().as_str(),
            canonical_fraction
        );

        let noncanonical = format!("0{}", "1".repeat(MAX_BYTES));
        assert!(matches!(
            Offset::parse(&noncanonical),
            Err(ParseError::TooLong { len, max })
                if len == MAX_BYTES + 1 && max == MAX_BYTES
        ));
    }

    #[test]
    fn owned_parse_and_display_round_trip_canonical_value() {
        let value = Offset::try_from(String::from("2500ms")).unwrap();
        assert_eq!(value.to_string(), "2500");
        assert_eq!(Offset::parse(&value.to_string()).unwrap(), value);
        assert_eq!(Offset::new(2, Unit::Second), Offset::ms(2000));
        assert_eq!(Offset::secs(2), Offset::ms(2000));
        assert_eq!(Offset::hours(1), Offset::mins(60));
        assert_eq!(Offset::micros(1), Offset::nanos(1000));
        assert!(Offset::ZERO.is_zero());
    }

    #[test]
    fn duration_conversions_are_exact_and_checked() {
        let duration = Duration::new(u64::MAX, 123_456_789);
        let offset = Offset::from(duration);
        assert_eq!(offset.duration().unwrap(), duration);
        assert_eq!(Offset::from(Duration::from_nanos(1)).as_str(), "0.000001");
        assert_eq!(
            Offset::parse("0.000001").unwrap().duration().unwrap(),
            Duration::from_nanos(1)
        );
        assert!(matches!(
            Offset::parse("0.0000001").unwrap().duration(),
            Err(ParseError::NotRepresentableAsDuration)
        ));
        assert!(matches!(
            Offset::parse(&format!("{}0s", u64::MAX))
                .unwrap()
                .duration(),
            Err(ParseError::NotRepresentableAsDuration)
        ));
    }
}
