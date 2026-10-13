//! Numbers: the .NET `NumberStyles` parses OfficeCLI applies with the
//! invariant culture, exact decimal rounding for `0.##`-style output, and
//! .NET's shortest round-trip rendering of doubles.

use std::fmt::{self, Write as _};

use super::error::{ErrorKind, ValueError, What};
use super::text::{StackStr, is_number_space};

/// A .NET `NumberStyles` composite, restricted to the flags OfficeCLI uses.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct Styles(u16);

impl Styles {
    const LEADING_WHITE: u16 = 1;
    const TRAILING_WHITE: u16 = 1 << 1;
    const LEADING_SIGN: u16 = 1 << 2;
    const TRAILING_SIGN: u16 = 1 << 3;
    const PARENTHESES: u16 = 1 << 4;
    const DECIMAL_POINT: u16 = 1 << 5;
    const THOUSANDS: u16 = 1 << 6;
    const EXPONENT: u16 = 1 << 7;
    const CURRENCY: u16 = 1 << 8;

    /// `NumberStyles.Integer`: `[ws][sign]digits[ws]`.
    pub(super) const INTEGER: Self =
        Self(Self::LEADING_WHITE | Self::TRAILING_WHITE | Self::LEADING_SIGN);
    /// `NumberStyles.Float`: `[ws][sign][digits][.digits][e[sign]digits][ws]`.
    pub(super) const FLOAT: Self = Self(Self::INTEGER.0 | Self::DECIMAL_POINT | Self::EXPONENT);
    /// `NumberStyles.Float | NumberStyles.AllowThousands`, the style of
    /// .NET's `double.TryParse(text, provider, out value)` overload.
    pub(super) const FLOAT_THOUSANDS: Self = Self(Self::FLOAT.0 | Self::THOUSANDS);
    /// `NumberStyles.Any`: also trailing signs, parentheses and the
    /// invariant currency symbol `¤`.
    pub(super) const ANY: Self =
        Self(Self::FLOAT_THOUSANDS.0 | Self::TRAILING_SIGN | Self::PARENTHESES | Self::CURRENCY);

    const fn has(self, flag: u16) -> bool {
        self.0 & flag != 0
    }
}

/// The invariant culture's currency symbol.
const CURRENCY: &[u8] = "\u{a4}".as_bytes();

/// The parts of a number the scanner accepted.
struct Scan<'a> {
    negative: bool,
    /// Digits, group separators, the decimal point and the exponent, as one
    /// contiguous slice of the input.
    core: &'a str,
    grouped: bool,
}

/// .NET's `Number.TryParseNumber` state machine for the invariant culture,
/// requiring the whole input to be consumed. Words such as `Infinity` and
/// `NaN`, which .NET accepts as a fallback, are not numbers here: every
/// OfficeCLI caller rejects non-finite values anyway.
fn scan(text: &str, styles: Styles) -> Option<Scan<'_>> {
    let bytes = text.as_bytes();
    let mut index = 0;
    let mut sign = false;
    let mut parens = false;
    let mut currency = false;
    let mut negative = false;

    // Leading white space, sign, parenthesis and currency symbol. White space
    // is refused after a sign unless a currency symbol followed it
    // ("- 5" fails, as the invariant culture's negative pattern is "-n").
    loop {
        let byte = bytes.get(index).copied();
        let white = byte.is_some_and(is_number_space);
        if !white || !styles.has(Styles::LEADING_WHITE) || (sign && !currency) {
            if styles.has(Styles::LEADING_SIGN) && !sign && matches!(byte, Some(b'+' | b'-')) {
                sign = true;
                negative = byte == Some(b'-');
            } else if byte == Some(b'(') && styles.has(Styles::PARENTHESES) && !sign {
                sign = true;
                parens = true;
                negative = true;
            } else if styles.has(Styles::CURRENCY)
                && !currency
                && bytes[index..].starts_with(CURRENCY)
            {
                currency = true;
                index += CURRENCY.len() - 1;
            } else {
                break;
            }
        }
        index += 1;
    }

    // Digits, group separators (after a digit, before the decimal point) and
    // one decimal point.
    let core_start = index;
    let mut digits = false;
    let mut decimal = false;
    let mut grouped = false;
    loop {
        match bytes.get(index) {
            Some(byte) if byte.is_ascii_digit() => digits = true,
            Some(b'.') if styles.has(Styles::DECIMAL_POINT) && !decimal => decimal = true,
            Some(b',') if styles.has(Styles::THOUSANDS) && digits && !decimal => grouped = true,
            _ => break,
        }
        index += 1;
    }
    if !digits {
        return None;
    }

    // An exponent is consumed only when digits follow `e[sign]`.
    if styles.has(Styles::EXPONENT) && matches!(bytes.get(index), Some(b'e' | b'E')) {
        let mut cursor = index + 1;
        if matches!(bytes.get(cursor), Some(b'+' | b'-')) {
            cursor += 1;
        }
        if bytes.get(cursor).is_some_and(u8::is_ascii_digit) {
            while bytes.get(cursor).is_some_and(u8::is_ascii_digit) {
                cursor += 1;
            }
            index = cursor;
        }
    }
    let core_end = index;

    // Trailing white space, sign, closing parenthesis and currency symbol.
    loop {
        let byte = bytes.get(index).copied();
        let white = byte.is_some_and(is_number_space);
        if !white || !styles.has(Styles::TRAILING_WHITE) {
            if styles.has(Styles::TRAILING_SIGN) && !sign && matches!(byte, Some(b'+' | b'-')) {
                sign = true;
                negative = byte == Some(b'-');
            } else if byte == Some(b')') && parens {
                parens = false;
            } else if styles.has(Styles::CURRENCY)
                && !currency
                && bytes[index..].starts_with(CURRENCY)
            {
                currency = true;
                index += CURRENCY.len() - 1;
            } else {
                break;
            }
        }
        index += 1;
    }
    if parens || index != bytes.len() {
        return None;
    }
    Some(Scan {
        negative,
        core: &text[core_start..core_end],
        grouped,
    })
}

/// Parses a double in `styles`. Overflow yields an infinity, as in .NET; the
/// callers decide whether a non-finite value is acceptable.
pub(super) fn parse_f64(text: &str, styles: Styles) -> Option<f64> {
    let scan = scan(text, styles)?;
    let magnitude = if scan.grouped {
        parse_ungrouped(scan.core)?
    } else {
        scan.core.parse::<f64>().ok()?
    };
    Some(if scan.negative { -magnitude } else { magnitude })
}

/// Parses a numeral with its group separators removed. Numerals of up to 128
/// characters are rebuilt on the stack; longer grouped numerals allocate.
fn parse_ungrouped(core: &str) -> Option<f64> {
    let mut buffer = StackStr::<128>::new();
    if core.split(',').all(|part| buffer.write_str(part).is_ok()) {
        return buffer.as_str().parse().ok();
    }
    core.replace(',', "").parse().ok()
}

/// Parses a 64-bit integer in `styles`, which must not admit a decimal point,
/// group separators or an exponent.
pub(super) fn parse_i64(text: &str, styles: Styles) -> Option<i64> {
    let scan = scan(text, styles)?;
    let mut magnitude: u64 = 0;
    for byte in scan.core.bytes() {
        if !byte.is_ascii_digit() {
            return None;
        }
        magnitude = magnitude
            .checked_mul(10)?
            .checked_add(u64::from(byte - b'0'))?;
    }
    if scan.negative {
        (magnitude <= 1 << 63).then(|| 0_i64.wrapping_sub_unsigned(magnitude))
    } else {
        i64::try_from(magnitude).ok()
    }
}

/// Parses a 32-bit integer in `NumberStyles.Integer` (.NET `int.TryParse`).
pub(super) fn parse_i32(text: &str) -> Option<i32> {
    parse_i64(text, Styles::INTEGER).and_then(|value| i32::try_from(value).ok())
}

/// OfficeCLI's `NumericText.HasValidThousandsGrouping`: every comma sits
/// after a digit, before the first decimal point, and is followed by exactly
/// three digits.
fn has_valid_thousands_grouping(text: &str) -> bool {
    let bytes = text.as_bytes();
    let dot = bytes.iter().position(|&byte| byte == b'.');
    bytes
        .iter()
        .enumerate()
        .filter(|&(_, &byte)| byte == b',')
        .all(|(index, _)| {
            dot.is_none_or(|dot| index < dot)
                && index > 0
                && bytes[index - 1].is_ascii_digit()
                && index + 3 < bytes.len()
                && bytes[index + 1..=index + 3].iter().all(u8::is_ascii_digit)
                && !bytes.get(index + 4).is_some_and(u8::is_ascii_digit)
        })
}

/// "This text is a number", as OfficeCLI decides it for cell values,
/// selector comparison and header-row predicates (`NumericText.TryParse`):
/// `NumberStyles.Any` with the invariant culture, refusing malformed
/// thousands groups such as `1,5`. The result can be infinite when the
/// numeral overflows.
///
/// ```
/// use litchi_ops::value::parse_numeric_text;
///
/// assert_eq!(parse_numeric_text("1,234.5"), Some(1234.5));
/// assert_eq!(parse_numeric_text(" (42) "), Some(-42.0));
/// assert_eq!(parse_numeric_text("1,5"), None);
/// assert_eq!(parse_numeric_text("12pt"), None);
/// ```
#[must_use]
pub fn parse_numeric_text(text: &str) -> Option<f64> {
    if has_valid_thousands_grouping(text) {
        parse_f64(text, Styles::ANY)
    } else {
        None
    }
}

/// Parses a finite number the way OfficeCLI's `SafeParseDouble` does:
/// `NumberStyles.Float` with the invariant culture. White space around the
/// number, a leading sign, a decimal point and an exponent are accepted;
/// group separators, `NaN` and infinities are not.
///
/// # Errors
///
/// Returns an `invalid_value` [`ValueError`] when the text is not such a
/// number or the number is not finite.
///
/// ```
/// use litchi_ops::value::parse_number;
///
/// assert_eq!(parse_number(" -1.5e2 ")?, -150.0);
/// assert!(parse_number("1,000").is_err());
/// # Ok::<(), litchi_ops::Error>(())
/// ```
pub fn parse_number(text: &str) -> Result<f64, ValueError<'_>> {
    match parse_f64(text, Styles::FLOAT) {
        Some(value) if value.is_finite() => Ok(value),
        Some(_) => Err(ValueError::new(What::Number, ErrorKind::NotFinite, text)),
        None if text.trim().is_empty() => {
            Err(ValueError::new(What::Number, ErrorKind::Empty, text))
        },
        None => Err(ValueError::new(What::Number, ErrorKind::Syntax, text)),
    }
}

/// Parses a 64-bit integer in .NET's `NumberStyles.Integer`: white space
/// around the digits and a leading sign. OfficeCLI's `SafeParseInt` reads
/// 32-bit integers; a property that needs that bound checks it after
/// parsing.
///
/// # Errors
///
/// Returns an `invalid_value` [`ValueError`] for anything else, including
/// decimals and values outside the 64-bit range.
///
/// ```
/// use litchi_ops::value::parse_integer;
///
/// assert_eq!(parse_integer(" +42 ")?, 42);
/// assert!(parse_integer("4.0").is_err());
/// # Ok::<(), litchi_ops::Error>(())
/// ```
pub fn parse_integer(text: &str) -> Result<i64, ValueError<'_>> {
    match parse_i64(text, Styles::INTEGER) {
        Some(value) => Ok(value),
        None if text.trim().is_empty() => {
            Err(ValueError::new(What::Integer, ErrorKind::Empty, text))
        },
        None => Err(ValueError::new(What::Integer, ErrorKind::Syntax, text)),
    }
}

/// An exact decimal printed without trailing zeros, the way .NET's custom
/// numeric format `0.##` (with one `#` per allowed decimal) prints a value
/// that was rounded half away from zero.
///
/// OfficeCLI renders lengths, font sizes and spacing this way (`2.54cm`,
/// `10.5pt`, `1.1042x`). Building the decimal from an exact ratio keeps the
/// result independent of floating-point error.
///
/// ```
/// use litchi_ops::value::Decimal;
///
/// // 265 twips of "auto" line spacing, printed with four decimals.
/// let spacing = Decimal::from_ratio(265, 240, 4).expect("a positive denominator");
/// assert_eq!(spacing.to_string(), "1.1042");
/// assert_eq!(Decimal::from_ratio(-1, 1000, 2).expect("valid").to_string(), "-0");
/// ```
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Decimal {
    negative: bool,
    magnitude: u128,
    scale: u8,
}

impl Decimal {
    /// The largest number of decimals [`Decimal::from_ratio`] accepts.
    pub const MAX_DECIMALS: u8 = 18;

    /// Rounds `numerator / denominator` half away from zero to at most
    /// `max_decimals` decimals.
    ///
    /// A negative value that rounds to zero keeps its sign and prints as
    /// `-0`, as .NET does. Returns `None` when `denominator` is zero or
    /// `max_decimals` exceeds [`Decimal::MAX_DECIMALS`].
    #[must_use]
    pub fn from_ratio(numerator: i64, denominator: i64, max_decimals: u8) -> Option<Self> {
        if denominator == 0 || max_decimals > Self::MAX_DECIMALS {
            return None;
        }
        let negative = numerator != 0 && (numerator < 0) != (denominator < 0);
        // In 64 bits when the scaled numerator fits, as for every length a
        // workbook or deck holds; the same rounding in 128 bits otherwise.
        // `remainder >= denominator - remainder` is `2 * remainder >=
        // denominator` without overflow.
        let power = 10_u64.pow(u32::from(max_decimals));
        let (numerator, denominator) = (numerator.unsigned_abs(), denominator.unsigned_abs());
        let magnitude = match numerator.checked_mul(power) {
            Some(scaled) => {
                let (quotient, remainder) = (scaled / denominator, scaled % denominator);
                u128::from(quotient + u64::from(remainder >= denominator - remainder))
            },
            None => {
                let scaled = u128::from(numerator) * u128::from(power);
                let denominator = u128::from(denominator);
                let (quotient, remainder) = (scaled / denominator, scaled % denominator);
                quotient + u128::from(remainder >= denominator - remainder)
            },
        };
        Some(Self::normalized(negative, magnitude, max_decimals))
    }

    /// Builds `±magnitude × 10^-scale` with trailing zeros removed.
    pub(super) const fn normalized(negative: bool, mut magnitude: u128, mut scale: u8) -> Self {
        if magnitude <= u64::MAX as u128 {
            // 64-bit division where the magnitude allows it.
            #[allow(
                clippy::cast_possible_truncation,
                reason = "the magnitude was just checked to fit 64 bits"
            )]
            let mut small = magnitude as u64;
            while scale > 0 && small % 10 == 0 {
                small /= 10;
                scale -= 1;
            }
            magnitude = small as u128;
        } else {
            while scale > 0 && magnitude % 10 == 0 {
                magnitude /= 10;
                scale -= 1;
            }
        }
        Self {
            negative,
            magnitude,
            scale,
        }
    }

    /// `true` when the rounded value is zero (of either sign).
    #[must_use]
    pub const fn is_zero(self) -> bool {
        self.magnitude == 0
    }

    /// The value as the double a correctly rounded parse of its text reads,
    /// computed without the text: a magnitude of at most 2^53 and a power of
    /// ten of at most 10^22 are exact doubles, so one division rounds
    /// correctly (Clinger's fast path). `None` beyond that range.
    pub(super) fn to_f64_exact(self) -> Option<f64> {
        const POWERS: [f64; 23] = [
            1e0, 1e1, 1e2, 1e3, 1e4, 1e5, 1e6, 1e7, 1e8, 1e9, 1e10, 1e11, 1e12, 1e13, 1e14, 1e15,
            1e16, 1e17, 1e18, 1e19, 1e20, 1e21, 1e22,
        ];
        let magnitude = u64::try_from(self.magnitude)
            .ok()
            .filter(|magnitude| *magnitude <= 1 << 53)?;
        let power = POWERS.get(usize::from(self.scale))?;
        #[allow(
            clippy::cast_precision_loss,
            reason = "the magnitude is at most 2^53, exact as a double"
        )]
        let value = magnitude as f64 / power;
        Some(if self.negative { -value } else { value })
    }

    /// `true` for a negative value, including a negative value that rounded
    /// to zero.
    #[must_use]
    pub const fn is_negative(self) -> bool {
        self.negative
    }
}

impl fmt::Display for Decimal {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        if self.negative {
            formatter.write_char('-')?;
        }
        let unit = 10_u128.pow(u32::from(self.scale));
        match u64::try_from(self.magnitude / unit) {
            Ok(whole) => write!(formatter, "{whole}")?,
            Err(_) => write!(formatter, "{}", self.magnitude / unit)?,
        }
        if self.scale > 0 {
            // Below `unit`, at most 10^18: always 64 bits.
            let fraction = u64::try_from(self.magnitude % unit).map_err(|_| fmt::Error)?;
            write!(
                formatter,
                ".{fraction:0width$}",
                width = usize::from(self.scale)
            )?;
        }
        Ok(())
    }
}

/// Renders a double the way .NET Core 3.0 and later print one with
/// `ToString()` and the invariant culture: the shortest digits that round
/// trip, positional for decimal exponents from -4 to 14 and scientific
/// (`1E+15`, `1E-05`) outside that range.
///
/// ```
/// use litchi_ops::value::DotNetDouble;
///
/// assert_eq!(DotNetDouble(1.5).to_string(), "1.5");
/// assert_eq!(DotNetDouble(1e15).to_string(), "1E+15");
/// assert_eq!(DotNetDouble(0.0001).to_string(), "0.0001");
/// assert_eq!(DotNetDouble(0.00001).to_string(), "1E-05");
/// ```
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct DotNetDouble(pub f64);

impl fmt::Display for DotNetDouble {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        write_double(formatter, self.0)
    }
}

/// Writes `value` as [`DotNetDouble`] describes.
pub(super) fn write_double(out: &mut impl fmt::Write, value: f64) -> fmt::Result {
    if value.is_nan() {
        return out.write_str("NaN");
    }
    if value.is_infinite() {
        return out.write_str(if value < 0.0 { "-Infinity" } else { "Infinity" });
    }
    if value == 0.0 {
        return out.write_str(if value.is_sign_negative() { "-0" } else { "0" });
    }
    // Rust's `{:e}` prints the shortest round-trip digits as `d.ddde±x`.
    let mut scientific = StackStr::<40>::new();
    write!(scientific, "{:e}", value.abs())?;
    let (mantissa, exponent) = scientific.as_str().split_once('e').ok_or(fmt::Error)?;
    let exponent: i32 = exponent.parse().map_err(|_| fmt::Error)?;
    let (lead, rest) = mantissa.split_once('.').unwrap_or((mantissa, ""));
    if value < 0.0 {
        out.write_char('-')?;
    }
    if !(-4..15).contains(&exponent) {
        out.write_str(lead)?;
        if !rest.is_empty() {
            out.write_char('.')?;
            out.write_str(rest)?;
        }
        let sign = if exponent < 0 { '-' } else { '+' };
        return write!(out, "E{sign}{:02}", exponent.unsigned_abs());
    }
    // `lead` is one digit; `digits` is `lead` followed by `rest`.
    let integer_digits = usize::try_from(exponent + 1).unwrap_or(0);
    if exponent >= 0 {
        let all = rest.len() + 1;
        if all <= integer_digits {
            out.write_str(lead)?;
            out.write_str(rest)?;
            for _ in all..integer_digits {
                out.write_char('0')?;
            }
        } else {
            out.write_str(lead)?;
            out.write_str(&rest[..integer_digits - 1])?;
            out.write_char('.')?;
            out.write_str(&rest[integer_digits - 1..])?;
        }
    } else {
        out.write_str("0.")?;
        for _ in 0..(exponent.unsigned_abs() - 1) {
            out.write_char('0')?;
        }
        out.write_str(lead)?;
        out.write_str(rest)?;
    }
    Ok(())
}
