//! Deterministic, allocation-free parser for the scalar `VALUE` function.
//!
//! The parser deliberately receives borrowed text and owns no evaluator state.
//! The caller charges the source bytes and performs cancellation checks around
//! this bounded scan.  This profile is en-US: it accepts the invariant integer
//! and exponent forms, en-US decimal/grouping/currency/percent forms, mixed
//! fractions, ISO dates/times/datetimes, and the en-US date spellings from ODF
//! 1.4 §6.13.34.
//!
//! Two-digit years use an explicit pivot: `00..=29` means 2000..=2029 and
//! `30..=99` means 1930..=1999.  The pivot is part of the evaluator profile;
//! it does not consult the host clock or locale.

use super::super::ScalarError;
use super::super::calendar::{
    CivilDate, MAX_CALENDAR_YEAR, MAX_DATE_SERIAL_EXCLUSIVE, MIN_CALENDAR_YEAR, MIN_DATE_SERIAL,
    MONTH_LONG, MONTH_SHORT, SECONDS_PER_DAY, serial_from_civil,
};

const TWO_DIGIT_YEAR_PIVOT: u32 = 30;
// Normal grouped values use the exact `fast_float2` conversion after commas
// have been removed.  Larger inputs use caller-supplied scratch through
// `parse_value_with_scratch`; the scalar wrapper remains intentionally bounded
// so it never hides an uncharged allocation.
const MAX_NORMALIZED_BYTES: usize = 1024;

/// Parse one borrowed `VALUE` text according to the deterministic en-US
/// profile.  No heap allocation is performed by this function.
#[inline]
pub(in super::super) fn parse_value(text: &str) -> Result<f64, ScalarError> {
    let mut scratch = [0_u8; MAX_NORMALIZED_BYTES];
    parse_value_with_scratch(text, &mut scratch)
}

/// Return whether the caller should reserve evaluator-owned normalization
/// scratch before parsing this source.  Comma-free values take the direct
/// `fast_float2` path and need no scratch even when they are long.
#[inline]
pub(in super::super) fn needs_scratch(text: &str) -> bool {
    text.len() > MAX_NORMALIZED_BYTES && text.contains(',')
}

/// Parse `VALUE` text using caller-owned bounded scratch for decimal
/// normalization.  The inspection evaluator reserves this buffer against its
/// text/parser budget before calling this helper; the short `parse_value`
/// wrapper above is retained for scalar callers and tests.  For a grouped
/// decimal, the caller must provide at least the input byte length (the exact
/// normalized requirement is no larger); an undersized buffer returns the
/// numeric-domain error instead of rounding a truncated prefix.
pub(in super::super) fn parse_value_with_scratch(
    text: &str,
    scratch: &mut [u8],
) -> Result<f64, ScalarError> {
    let text = trim_ascii_space(text);
    if text.is_empty() {
        return Err(ScalarError::Value);
    }

    if let Some(result) = parse_date_or_time(text) {
        return result;
    }
    if let Some(result) = parse_fraction(text) {
        return result;
    }
    parse_number(text, scratch)
}

/// Parse the date-oriented profile used by `DATEVALUE` and DateParam.
///
/// Date and datetime spellings are admitted through the shared borrowed
/// grammar.  A clock-only spelling is deliberately rejected, and the only
/// fallback is the numeric portion of the VALUE grammar (grouping, currency,
/// percent, sign, exponent, and mixed-fraction forms are retained; dates,
/// clocks, slashes, and colons are not).
pub(in super::super) fn parse_date_value(text: &str) -> Result<f64, ScalarError> {
    let mut scratch = [0_u8; MAX_NORMALIZED_BYTES];
    parse_date_value_with_scratch(text, &mut scratch)
}

/// Parse `DATEVALUE` text with caller-owned numeric normalization scratch.
pub(in super::super) fn parse_date_value_with_scratch(
    text: &str,
    scratch: &mut [u8],
) -> Result<f64, ScalarError> {
    let text = trim_ascii_space(text);
    if text.is_empty() {
        return Err(ScalarError::Value);
    }
    if let Some((kind, result)) = parse_date_time_candidate(text, DateDomain::DateTime) {
        return match kind {
            DateTimeKind::Date | DateTimeKind::DateTime => {
                let value = result?;
                let value = validate_serial(value, DateDomain::DateTime)?;
                Ok(value.floor())
            },
            DateTimeKind::Time => Err(ScalarError::Value),
        };
    }
    let value = parse_numeric_fallback_with_scratch(text, scratch)?;
    Ok(validate_serial(value, DateDomain::DateTime)?.floor())
}

/// Parse the time-oriented profile used by `TIMEVALUE` and TimeParam.
///
/// A combined datetime contributes only its fractional day.  Date-only text
/// is rejected before the numeric fallback is attempted.
pub(in super::super) fn parse_time_value(text: &str) -> Result<f64, ScalarError> {
    let mut scratch = [0_u8; MAX_NORMALIZED_BYTES];
    parse_time_value_with_scratch(text, &mut scratch)
}

/// Parse `TIMEVALUE` text with caller-owned numeric normalization scratch.
pub(in super::super) fn parse_time_value_with_scratch(
    text: &str,
    scratch: &mut [u8],
) -> Result<f64, ScalarError> {
    let text = trim_ascii_space(text);
    if text.is_empty() {
        return Err(ScalarError::Value);
    }
    if let Some((kind, result)) = parse_date_time_candidate(text, DateDomain::DateTime) {
        return match kind {
            DateTimeKind::Time => result,
            DateTimeKind::DateTime => {
                let value = result?;
                validate_serial(value, DateDomain::DateTime)?;
                // Keep the clock conversion independent of the large date
                // serial.  Subtracting `floor()` from a date-time serial can
                // lose several ulps when the date is around 40,000.
                parse_datetime_time_fraction(text)
            },
            DateTimeKind::Date => Err(ScalarError::Value),
        };
    }
    parse_numeric_fallback_with_scratch(text, scratch)
}

/// Parse the numeric-only portion of the established VALUE grammar.  Date and
/// clock forms are selected by the caller before this function; mixed
/// fractions are still admitted here, followed by grouping, currency,
/// percent, sign, and exponent behavior from `parse_number`.
pub(in super::super) fn parse_numeric_fallback_with_scratch(
    text: &str,
    scratch: &mut [u8],
) -> Result<f64, ScalarError> {
    let text = trim_ascii_space(text);
    if text.is_empty() {
        return Err(ScalarError::Value);
    }
    if let Some(result) = parse_fraction(text) {
        return result;
    }
    parse_number(text, scratch)
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum DateDomain {
    Value,
    DateTime,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum DateTimeKind {
    Date,
    Time,
    DateTime,
}

fn parse_date_time_candidate(
    text: &str,
    domain: DateDomain,
) -> Option<(DateTimeKind, Result<f64, ScalarError>)> {
    let bytes = text.as_bytes();

    // ISO dates are checked before locale slash/dash forms.  A ten-byte
    // candidate is a date; a suffix after the separator makes it a datetime
    // even when that suffix is malformed and must return #VALUE!.
    if bytes.len() >= 10 && bytes.get(4) == Some(&b'-') && bytes.get(7) == Some(&b'-') {
        let kind = if bytes.len() == 10 {
            DateTimeKind::Date
        } else {
            DateTimeKind::DateTime
        };
        return Some((kind, parse_iso_date_time_with_domain(text, domain)));
    }

    if bytes.contains(&b':') {
        if let Some(result) = parse_non_iso_datetime_with_domain(text, domain) {
            return Some((DateTimeKind::DateTime, result));
        }
        return Some((DateTimeKind::Time, parse_clock(text)));
    }

    let slash_count = bytes.iter().filter(|byte| **byte == b'/').count();
    let dash_count = bytes.iter().filter(|byte| **byte == b'-').count();
    let has_exponent_marker = bytes.iter().any(|byte| matches!(*byte, b'e' | b'E'));
    if slash_count == 2 || (dash_count == 2 && !has_exponent_marker) {
        return Some((
            DateTimeKind::Date,
            parse_numeric_date_with_domain(text, domain),
        ));
    }

    if bytes
        .iter()
        .any(|byte| byte.is_ascii_alphabetic() && !matches!(*byte, b'e' | b'E'))
    {
        return Some((
            DateTimeKind::Date,
            parse_english_date_with_domain(text, domain),
        ));
    }

    None
}

fn trim_ascii_space(text: &str) -> &str {
    let bytes = text.as_bytes();
    let mut start = 0_usize;
    let mut end = bytes.len();
    while start < end && bytes[start].is_ascii_whitespace() {
        start += 1;
    }
    while end > start && bytes[end - 1].is_ascii_whitespace() {
        end -= 1;
    }
    // VALUE input is UTF-8.  The boundaries found above are ASCII bytes, so
    // they remain valid UTF-8 boundaries.
    &text[start..end]
}

fn parse_date_or_time(text: &str) -> Option<Result<f64, ScalarError>> {
    parse_date_time_candidate(text, DateDomain::Value).map(|(_, result)| result)
}

fn parse_non_iso_datetime_with_domain(
    text: &str,
    domain: DateDomain,
) -> Option<Result<f64, ScalarError>> {
    let bytes = text.as_bytes();

    // A literal T is unambiguous for the listed locale date forms.
    if let Some(separator) = bytes
        .iter()
        .rposition(|byte| *byte == b'T' || *byte == b't')
    {
        let date_text = trim_ascii_space(&text[..separator]);
        let time_text = &text[separator + 1..];
        if !date_text.is_empty()
            && !time_text.is_empty()
            && time_text.contains(':')
            && let Ok(date) = parse_date_text_with_domain(date_text, domain)
        {
            return Some(match parse_clock(time_text) {
                Ok(time) => finite_sum(date, time),
                Err(error) => Err(error),
            });
        }
    }

    // For a space separator, the final token is the clock.  Trying the split
    // from the right preserves the spaces within alphabetic month forms.
    let separator = bytes.iter().rposition(u8::is_ascii_whitespace)?;
    let date_text = trim_ascii_space(&text[..separator]);
    let time_text = trim_ascii_space(&text[separator + 1..]);
    if date_text.is_empty() || time_text.is_empty() || !time_text.contains(':') {
        return None;
    }
    Some(combine_date_and_time_with_domain(
        date_text, time_text, domain,
    ))
}

fn parse_datetime_time_fraction(text: &str) -> Result<f64, ScalarError> {
    let bytes = text.as_bytes();
    if let Some(separator) = bytes
        .iter()
        .rposition(|byte| *byte == b'T' || *byte == b't')
    {
        return parse_clock(&text[separator + 1..]);
    }
    let separator = bytes.iter().rposition(u8::is_ascii_whitespace);
    let Some(separator) = separator else {
        return Err(ScalarError::Value);
    };
    parse_clock(trim_ascii_space(&text[separator + 1..]))
}

fn combine_date_and_time_with_domain(
    date_text: &str,
    time_text: &str,
    domain: DateDomain,
) -> Result<f64, ScalarError> {
    let date = parse_date_text_with_domain(date_text, domain)?;
    finite_sum(date, parse_clock(time_text)?)
}

fn parse_date_text_with_domain(date_text: &str, domain: DateDomain) -> Result<f64, ScalarError> {
    if date_text.as_bytes().contains(&b'/') || date_text.as_bytes().contains(&b'-') {
        parse_numeric_date_with_domain(date_text, domain)
    } else {
        parse_english_date_with_domain(date_text, domain)
    }
}

fn parse_iso_date_time_with_domain(text: &str, domain: DateDomain) -> Result<f64, ScalarError> {
    let bytes = text.as_bytes();
    if bytes.len() < 10
        || !all_ascii_digits(&bytes[0..4])
        || bytes[4] != b'-'
        || !all_ascii_digits(&bytes[5..7])
        || bytes[7] != b'-'
        || !all_ascii_digits(&bytes[8..10])
    {
        return Err(ScalarError::Value);
    }

    let year = parse_digits_u32(&bytes[0..4]).ok_or(ScalarError::Value)?;
    let month = parse_digits_u32(&bytes[5..7]).ok_or(ScalarError::Value)?;
    let day = parse_digits_u32(&bytes[8..10]).ok_or(ScalarError::Value)?;
    let date = serial_date_with_domain(year as i32, month, day, domain)?;
    if bytes.len() == 10 {
        return Ok(date);
    }

    let separator = *bytes.get(10).ok_or(ScalarError::Value)?;
    if separator != b'T' && separator != b't' && separator != b' ' {
        return Err(ScalarError::Value);
    }
    let time = parse_clock(&text[11..])?;
    finite_sum(date, time)
}

fn parse_numeric_date_with_domain(text: &str, domain: DateDomain) -> Result<f64, ScalarError> {
    let bytes = text.as_bytes();
    let separator = if bytes.contains(&b'/') { b'/' } else { b'-' };
    let first = bytes
        .iter()
        .position(|byte| *byte == separator)
        .ok_or(ScalarError::Value)?;
    let second = bytes
        .iter()
        .enumerate()
        .skip(first + 1)
        .find_map(|(index, byte)| (*byte == separator).then_some(index))
        .ok_or(ScalarError::Value)?;
    if bytes[first + 1..second].contains(&if separator == b'/' { b'-' } else { b'/' })
        || bytes[second + 1..].contains(&separator)
    {
        return Err(ScalarError::Value);
    }

    let month = parse_date_component(&text[..first]).ok_or(ScalarError::Value)?;
    let day = parse_date_component(&text[first + 1..second]).ok_or(ScalarError::Value)?;
    let year_text = &text[second + 1..];
    let year = if separator == b'/' {
        parse_year_component(year_text)
    } else if year_text.len() == 4 && all_ascii_digits(year_text.as_bytes()) {
        parse_digits_u32(year_text.as_bytes()).map(|year| year as i32)
    } else {
        None
    }
    .ok_or(ScalarError::Value)?;
    serial_date_with_domain(year, month, day, domain)
}

fn parse_english_date_with_domain(text: &str, domain: DateDomain) -> Result<f64, ScalarError> {
    let mut cursor = 0_usize;
    let (first, first_comma) = next_date_token(text, &mut cursor).ok_or(ScalarError::Value)?;
    let (second, second_comma) = next_date_token(text, &mut cursor).ok_or(ScalarError::Value)?;
    let (third, third_comma) = next_date_token(text, &mut cursor).ok_or(ScalarError::Value)?;
    if next_date_token(text, &mut cursor).is_some() || third_comma {
        return Err(ScalarError::Value);
    }

    // mmm(mmmm) DD, YYYY
    if let Some(month) = month_number(first) {
        if first_comma {
            return Err(ScalarError::Value);
        }
        if let Some(day) = parse_date_component(second)
            && second_comma
            && let Some(year) = parse_long_year_component(third)
        {
            return serial_date_with_domain(year, month, day, domain);
        }
    }

    // DD mmm(mmmm) YYYY
    if !first_comma
        && !second_comma
        && let Some(day) = parse_date_component(first)
        && let Some(month) = month_number(second)
        && let Some(year) = parse_long_year_component(third)
    {
        return serial_date_with_domain(year, month, day, domain);
    }

    Err(ScalarError::Value)
}

fn next_date_token<'a>(text: &'a str, cursor: &mut usize) -> Option<(&'a str, bool)> {
    let bytes = text.as_bytes();
    while *cursor < bytes.len() && bytes[*cursor].is_ascii_whitespace() {
        *cursor += 1;
    }
    let start = *cursor;
    while *cursor < bytes.len() && !bytes[*cursor].is_ascii_whitespace() && bytes[*cursor] != b',' {
        *cursor += 1;
    }
    if start == *cursor {
        return None;
    }
    let token = &text[start..*cursor];
    while *cursor < bytes.len() && bytes[*cursor].is_ascii_whitespace() {
        *cursor += 1;
    }
    let comma = if bytes.get(*cursor) == Some(&b',') {
        *cursor += 1;
        true
    } else {
        false
    };
    Some((token, comma))
}

fn parse_clock(text: &str) -> Result<f64, ScalarError> {
    let bytes = text.as_bytes();
    let (hour, mut cursor) = parse_clock_component(bytes, 0)?;
    if bytes.get(cursor) != Some(&b':') {
        return Err(ScalarError::Value);
    }
    cursor += 1;
    let (minute, next) = parse_clock_component(bytes, cursor)?;
    cursor = next;

    let mut second = 0.0_f64;
    if bytes.get(cursor) == Some(&b':') {
        cursor += 1;
        let (whole_second, next) = parse_clock_component(bytes, cursor)?;
        if whole_second > 59 {
            return Err(ScalarError::Value);
        }
        second = whole_second as f64;
        cursor = next;
        if bytes.get(cursor) == Some(&b'.') {
            cursor += 1;
            let start = cursor;
            while bytes.get(cursor).is_some_and(u8::is_ascii_digit) {
                cursor += 1;
            }
            if start == cursor {
                return Err(ScalarError::Value);
            }
            second += fraction_digits(&bytes[start..cursor]);
        }
    }
    if cursor != bytes.len() || hour > 23 || minute > 59 || !second.is_finite() {
        return Err(ScalarError::Value);
    }
    Ok((hour as f64 * 3_600.0 + minute as f64 * 60.0 + second) / SECONDS_PER_DAY)
}

fn parse_clock_component(bytes: &[u8], start: usize) -> Result<(u32, usize), ScalarError> {
    let mut cursor = start;
    while bytes.get(cursor).is_some_and(u8::is_ascii_digit) && cursor - start < 2 {
        cursor += 1;
    }
    if cursor == start {
        return Err(ScalarError::Value);
    }
    let value = parse_digits_u32(&bytes[start..cursor]).ok_or(ScalarError::Value)?;
    Ok((value, cursor))
}

fn fraction_digits(bytes: &[u8]) -> f64 {
    let mut value = 0.0_f64;
    // Horner evaluation from the least significant digit preserves the
    // decimal place order without allocating a prefixed `0.` string.
    for byte in bytes.iter().rev() {
        value = (value + f64::from(*byte - b'0')) / 10.0;
    }
    value
}

fn parse_fraction(text: &str) -> Option<Result<f64, ScalarError>> {
    let bytes = text.as_bytes();
    let slash = bytes.iter().position(|byte| *byte == b'/')?;
    if bytes[slash + 1..].contains(&b'/') {
        return Some(Err(ScalarError::Value));
    }

    let mut left = &text[..slash];
    let denominator_text = &text[slash + 1..];
    if denominator_text.is_empty()
        || !all_ascii_digits(denominator_text.as_bytes())
        || denominator_text.as_bytes()[0] == b'0'
    {
        return Some(Err(ScalarError::Value));
    }
    let denominator = parse_digits_u32(denominator_text.as_bytes()).unwrap_or(0);
    if denominator == 0 || denominator > 99 {
        return Some(Err(ScalarError::Value));
    }

    let mut negative = false;
    if let Some(rest) = left.strip_prefix('-') {
        negative = true;
        left = rest;
    } else if let Some(rest) = left.strip_prefix('+') {
        left = rest;
    }
    if left.is_empty() {
        return Some(Err(ScalarError::Value));
    }

    let (whole, numerator_text) =
        if let Some(space) = left.as_bytes().iter().position(|byte| *byte == b' ') {
            let whole_text = &left[..space];
            let numerator_text = &left[space + 1..];
            if whole_text.is_empty()
                || numerator_text.is_empty()
                || whole_text.as_bytes().iter().any(u8::is_ascii_whitespace)
                || numerator_text
                    .as_bytes()
                    .iter()
                    .any(u8::is_ascii_whitespace)
            {
                return Some(Err(ScalarError::Value));
            }
            let whole = match parse_unsigned_decimal(whole_text) {
                Ok(value) => value,
                Err(error) => return Some(Err(error)),
            };
            (whole, numerator_text)
        } else {
            (0.0, left)
        };
    if !all_ascii_digits(numerator_text.as_bytes()) {
        return Some(Err(ScalarError::Value));
    }
    let numerator = match parse_unsigned_decimal(numerator_text) {
        Ok(value) => value,
        Err(error) => return Some(Err(error)),
    };
    let result = whole + numerator / f64::from(denominator);
    if !result.is_finite() {
        return Some(Err(ScalarError::Number));
    }
    Some(Ok(if negative { -result } else { result }))
}

fn parse_number(text: &str, scratch: &mut [u8]) -> Result<f64, ScalarError> {
    let mut text = text;
    let mut negative = false;
    let mut parenthesized = false;

    if text.starts_with('(') || text.ends_with(')') {
        if !(text.starts_with('(') && text.ends_with(')')) {
            return Err(ScalarError::Value);
        }
        parenthesized = true;
        text = trim_ascii_space(&text[1..text.len() - 1]);
        if text.is_empty() {
            return Err(ScalarError::Value);
        }
    }

    if let Some(rest) = text.strip_prefix('-') {
        if parenthesized {
            return Err(ScalarError::Value);
        }
        negative = true;
        text = rest;
    } else if let Some(rest) = text.strip_prefix('+') {
        if parenthesized {
            return Err(ScalarError::Value);
        }
        text = rest;
    }
    if text.is_empty() {
        return Err(ScalarError::Value);
    }

    if let Some(rest) = text.strip_prefix('$') {
        text = rest;
    }
    if text.is_empty() || text.as_bytes().contains(&b'$') {
        return Err(ScalarError::Value);
    }

    let mut percent = false;
    if let Some(rest) = text.strip_suffix('%') {
        percent = true;
        text = rest;
    }
    if text.is_empty() || text.as_bytes().contains(&b'%') {
        return Err(ScalarError::Value);
    }

    let mut value = if text.as_bytes().contains(&b',') {
        parse_grouped_number(text, scratch)?
    } else {
        parse_plain_number(text)?
    };
    if negative ^ parenthesized {
        value = -value;
    }
    if percent {
        value /= 100.0;
    }
    if !value.is_finite() {
        return Err(ScalarError::Number);
    }
    Ok(value)
}

fn parse_unsigned_decimal(text: &str) -> Result<f64, ScalarError> {
    if text.is_empty() || !all_ascii_digits(text.as_bytes()) {
        return Err(ScalarError::Value);
    }
    // `fast_float2` performs the correctly rounded conversion for arbitrary
    // length ungrouped decimal input. Grouped input is handled separately so
    // it can use evaluator-owned normalization scratch.
    parse_plain_number(text)
}

fn parse_plain_number(text: &str) -> Result<f64, ScalarError> {
    if !valid_decimal_syntax(text.as_bytes(), false) {
        return Err(ScalarError::Value);
    }
    // Syntax was validated above, so a parser failure at this point is an
    // accepted lexical form outside the finite Number domain (for example a
    // huge exponent), rather than a malformed-text error.
    let value = fast_float2::parse::<f64, _>(text).map_err(|_| ScalarError::Number)?;
    if value.is_finite() {
        Ok(value)
    } else {
        Err(ScalarError::Number)
    }
}

fn parse_grouped_number(text: &str, scratch: &mut [u8]) -> Result<f64, ScalarError> {
    let bytes = text.as_bytes();
    if !valid_decimal_syntax(bytes, true) {
        return Err(ScalarError::Value);
    }

    // Keep grouped values on the same correctly rounded conversion path as
    // ungrouped values. The caller owns this scratch, so its capacity is
    // charged by the evaluator rather than being a hidden retained allocation.
    let source_needs_scratch = needs_scratch(text);
    let normalized_length = bytes.iter().filter(|byte| **byte != b',').count();
    if source_needs_scratch && normalized_length > scratch.len() {
        return Err(ScalarError::Number);
    }
    let mut length = 0_usize;
    for byte in bytes.iter().copied().filter(|byte| *byte != b',') {
        let Some(slot) = scratch.get_mut(length) else {
            // The scalar wrapper deliberately has a fixed stack scratch. The
            // production inspection bridge supplies a budgeted scratch large
            // enough for the complete normalized source, so this branch is a
            // defensive formula Number refusal rather than an approximation.
            return Err(ScalarError::Number);
        };
        *slot = byte;
        length += 1;
    }
    let normalized = std::str::from_utf8(&scratch[..length]).map_err(|_| ScalarError::Value)?;
    let value = fast_float2::parse::<f64, _>(normalized).map_err(|_| ScalarError::Number)?;
    value
        .is_finite()
        .then_some(value)
        .ok_or(ScalarError::Number)
}

fn valid_decimal_syntax(bytes: &[u8], grouped: bool) -> bool {
    if bytes.is_empty() {
        return false;
    }
    let mut cursor = 0_usize;
    let mut integer_digits = 0_usize;
    let mut first_group_digits = 0_usize;
    let mut group_digits = 0_usize;
    let mut comma_groups = 0_usize;
    let mut saw_comma = false;
    let mut saw_decimal = false;
    let mut fraction_digits = 0_usize;
    let mut saw_digit = false;

    while cursor < bytes.len() && bytes[cursor].is_ascii_digit() {
        integer_digits += 1;
        first_group_digits += 1;
        group_digits += 1;
        saw_digit = true;
        cursor += 1;
    }
    while grouped && cursor < bytes.len() && bytes[cursor] == b',' {
        if (comma_groups == 0 && group_digits == 0) || (comma_groups > 0 && group_digits != 3) {
            return false;
        }
        saw_comma = true;
        comma_groups += 1;
        group_digits = 0;
        cursor += 1;
        while cursor < bytes.len() && bytes[cursor].is_ascii_digit() {
            integer_digits += 1;
            group_digits += 1;
            saw_digit = true;
            cursor += 1;
        }
        if group_digits == 0 {
            return false;
        }
    }
    if saw_comma && (first_group_digits == 0 || group_digits != 3) {
        return false;
    }
    if cursor < bytes.len() && bytes[cursor] == b'.' {
        saw_decimal = true;
        cursor += 1;
        while cursor < bytes.len() && bytes[cursor].is_ascii_digit() {
            fraction_digits += 1;
            saw_digit = true;
            cursor += 1;
        }
        if fraction_digits == 0 {
            return false;
        }
    }
    if !saw_digit || (saw_decimal && integer_digits == 0 && fraction_digits == 0) {
        return false;
    }
    if cursor < bytes.len() && (bytes[cursor] == b'e' || bytes[cursor] == b'E') {
        cursor += 1;
        if bytes.get(cursor) == Some(&b'+') || bytes.get(cursor) == Some(&b'-') {
            cursor += 1;
        }
        let start = cursor;
        while cursor < bytes.len() && bytes[cursor].is_ascii_digit() {
            cursor += 1;
        }
        if start == cursor {
            return false;
        }
    }
    cursor == bytes.len()
}

fn serial_date_with_domain(
    year: i32,
    month: u32,
    day: u32,
    domain: DateDomain,
) -> Result<f64, ScalarError> {
    if !(1..=12).contains(&month) {
        return Err(ScalarError::Value);
    }
    if !(MIN_CALENDAR_YEAR..=MAX_CALENDAR_YEAR).contains(&year) {
        return Err(ScalarError::Number);
    }
    let date = CivilDate::new(year, month, day).ok_or(ScalarError::Value)?;
    let serial = serial_from_civil(date).ok_or(ScalarError::Number)?;
    // Profile serials are below three million, so this conversion is exact.
    let serial = serial as f64;
    validate_serial(serial, domain)
}

fn validate_serial(value: f64, domain: DateDomain) -> Result<f64, ScalarError> {
    let minimum = match domain {
        DateDomain::Value => 0.0,
        DateDomain::DateTime => MIN_DATE_SERIAL as f64,
    };
    let maximum = match domain {
        DateDomain::Value | DateDomain::DateTime => MAX_DATE_SERIAL_EXCLUSIVE as f64,
    };
    if !value.is_finite() || value < minimum || value >= maximum {
        return Err(ScalarError::Number);
    }
    Ok(value)
}

fn finite_sum(left: f64, right: f64) -> Result<f64, ScalarError> {
    let value = left + right;
    value
        .is_finite()
        .then_some(value)
        .ok_or(ScalarError::Number)
}

fn parse_date_component(text: &str) -> Option<u32> {
    if text.is_empty() || text.len() > 2 || !all_ascii_digits(text.as_bytes()) {
        return None;
    }
    parse_digits_u32(text.as_bytes())
}

fn parse_year_component(text: &str) -> Option<i32> {
    if !(text.len() == 2 || text.len() == 4) || !all_ascii_digits(text.as_bytes()) {
        return None;
    }
    let year = parse_digits_u32(text.as_bytes())?;
    if text.len() == 2 {
        Some(if year < TWO_DIGIT_YEAR_PIVOT {
            2_000 + year as i32
        } else {
            1_900 + year as i32
        })
    } else {
        Some(year as i32)
    }
}

fn parse_long_year_component(text: &str) -> Option<i32> {
    if text.len() != 4 || !all_ascii_digits(text.as_bytes()) {
        return None;
    }
    parse_digits_u32(text.as_bytes()).map(|year| year as i32)
}

fn month_number(text: &str) -> Option<u32> {
    for (index, name) in MONTH_SHORT.iter().enumerate() {
        if text.eq_ignore_ascii_case(name) || text.eq_ignore_ascii_case(MONTH_LONG[index]) {
            return Some(index as u32 + 1);
        }
    }
    None
}

fn all_ascii_digits(bytes: &[u8]) -> bool {
    !bytes.is_empty() && bytes.iter().all(u8::is_ascii_digit)
}

fn parse_digits_u32(bytes: &[u8]) -> Option<u32> {
    let mut value = 0_u32;
    for byte in bytes {
        value = value
            .checked_mul(10)?
            .checked_add(u32::from(*byte - b'0'))?;
    }
    Some(value)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn close(actual: f64, expected: f64) {
        assert!(
            (actual - expected).abs() < 1.0e-12,
            "{actual} != {expected}"
        );
    }

    #[test]
    fn parses_invariant_decimal_exponent_percent_and_currency() {
        close(parse_value("+12").unwrap(), 12.0);
        close(parse_value("-1.25e2").unwrap(), -125.0);
        close(parse_value("1e3").unwrap(), 1_000.0);
        close(parse_value("-1e-3").unwrap(), -0.001);
        close(parse_value("-1E-3").unwrap(), -0.001);
        close(parse_value("-1,234e-3").unwrap(), -1.234);
        close(parse_value("-$1,234e-3").unwrap(), -1.234);
        close(parse_value("12.5%").unwrap(), 0.125);
        close(parse_value("$1,234.56").unwrap(), 1_234.56);
        close(parse_value("1234,567").unwrap(), 1_234_567.0);
        close(parse_value("1234,567,890").unwrap(), 1_234_567_890.0);
        close(parse_value("($1,234.56)").unwrap(), -1_234.56);
        assert_eq!(
            parse_value("1,000,000,000,000,000,128").unwrap().to_bits(),
            1_000_000_000_000_000_128_f64.to_bits()
        );
    }

    #[test]
    fn parses_mixed_and_simple_fractions() {
        close(parse_value("1 1/2").unwrap(), 1.5);
        close(parse_value("-1 1/2").unwrap(), -1.5);
        close(parse_value("1/4").unwrap(), 0.25);
        assert_eq!(parse_value("1 1/0"), Err(ScalarError::Value));
        assert_eq!(parse_value("1 1/02"), Err(ScalarError::Value));
        assert_eq!(parse_value("1 1 /2"), Err(ScalarError::Value));
    }

    #[test]
    fn date_time_numeric_fallback_keeps_mixed_fraction_profile() {
        assert_eq!(parse_date_value("2 1/2").unwrap(), 2.0);
        assert_eq!(parse_time_value("2 1/2").unwrap(), 2.5);
        assert_eq!(parse_time_value("-2 1/2").unwrap(), -2.5);
        assert_eq!(parse_date_value("-2 1/2").unwrap(), -3.0);
    }

    #[test]
    fn timevalue_datetime_preserves_clock_precision() {
        let expected = (12.0 * 3_600.0 + 34.0 * 60.0 + 56.5) / SECONDS_PER_DAY;
        assert_eq!(
            parse_time_value("2006-05-21 12:34:56.5").unwrap().to_bits(),
            expected.to_bits()
        );
    }

    #[test]
    fn parses_times_iso_and_english_dates() {
        close(parse_value("2:00").unwrap(), 2.0 / 24.0);
        close(
            parse_value("12:34:56.5").unwrap(),
            (12.0 * 3600.0 + 34.0 * 60.0 + 56.5) / 86400.0,
        );
        close(
            parse_value("12:34:56.12").unwrap(),
            (12.0 * 3600.0 + 34.0 * 60.0 + 56.12) / 86400.0,
        );
        close(
            parse_value("12:34:56.001").unwrap(),
            (12.0 * 3600.0 + 34.0 * 60.0 + 56.001) / 86400.0,
        );
        assert!(parse_value("12:34:59.999999999999999999999999999999").is_ok());
        assert_eq!(parse_value("12:34:60.1"), Err(ScalarError::Value));
        close(parse_value("2020-01-02").unwrap(), 43_832.0);
        close(
            parse_value("2020-01-02T03:04:05.5").unwrap(),
            43_832.0 + (3.0 * 3600.0 + 4.0 * 60.0 + 5.5) / 86400.0,
        );
        close(
            parse_value("5/21/2006").unwrap(),
            parse_value("May 21, 2006").unwrap(),
        );
        close(
            parse_value("29 October 2006").unwrap(),
            parse_value("Oct 29, 2006").unwrap(),
        );
        close(
            parse_value("October 29, 2006 12:00").unwrap(),
            parse_value("Oct 29, 2006").unwrap() + 0.5,
        );
    }

    #[test]
    fn two_digit_year_pivot_is_stable() {
        assert_eq!(parse_year_component("29"), Some(2029));
        assert_eq!(parse_year_component("30"), Some(1930));
        close(
            parse_value("5/21/06").unwrap(),
            parse_value("5/21/2006").unwrap(),
        );
    }

    #[test]
    fn rejects_malformed_forms_and_nonfinite_numbers() {
        assert_eq!(parse_value("1,23"), Err(ScalarError::Value));
        assert_eq!(parse_value("2020-02-30"), Err(ScalarError::Value));
        assert_eq!(parse_value("25:00"), Err(ScalarError::Value));
        assert_eq!(parse_value("1e999"), Err(ScalarError::Number));
        assert_eq!(parse_value("1899-12-29"), Err(ScalarError::Number));
        assert_eq!(parse_value("not a value"), Err(ScalarError::Value));
    }

    #[test]
    fn long_grouped_decimal_uses_complete_caller_scratch() {
        let mut source = String::from("0");
        for _ in 0..400 {
            source.push_str(",000");
        }
        source.push_str(",001.000000000000000111022302462515654042363166809082031251");

        assert!(needs_scratch(&source));
        assert!(!needs_scratch("1,234"));
        assert!(!needs_scratch(
            "1".repeat(MAX_NORMALIZED_BYTES + 1).as_str()
        ));

        let mut scratch = vec![0_u8; source.len()];
        let exact = parse_value_with_scratch(&source, &mut scratch).unwrap();
        assert_eq!(exact.to_bits(), 1.0000000000000002_f64.to_bits());

        // A scalar call keeps its fixed stack bound and refuses this input;
        // it must never silently round a truncated prefix.
        assert_eq!(parse_value(&source), Err(ScalarError::Number));
    }
}
