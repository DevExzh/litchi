#![expect(
    clippy::shadow_reuse,
    reason = "parser bindings are intentionally refined after validation"
)]
//! Shared bounded validation for inert conflict markup.

use crate::{Error, Result};

use super::model::{Id, Limits};

const MAX_SOURCE_BYTES: usize = 64 * 1024 * 1024;
const MAX_CONFLICTS: usize = 1_000_000;
const MAX_RANGES: usize = 1_000_000;
const MAX_TEXT_BYTES: usize = 32 * 1024 * 1024;
const MAX_TEXT_SEGMENTS: usize = 2_000_000;
const MAX_DEPTH: usize = 1_024;
const MAX_ATTRIBUTES: usize = 256;
const MAX_EVENTS: usize = 4_000_000;
const MAX_METADATA_BYTES: usize = 32 * 1024 * 1024;
const MAX_OPEN_RANGES: usize = 1_000_000;
const MAX_ATTRIBUTE_BYTES: usize = 1024 * 1024;
const MAX_OUTPUT_BYTES: usize = 64 * 1024 * 1024;
const MAX_STORIES: usize = 4_096;
const MAX_TOTAL_STORY_BYTES: usize = 512 * 1024 * 1024;
const MAX_TOTAL_CONFLICTS: usize = 4_000_000;
const MAX_TOTAL_METADATA_BYTES: usize = 128 * 1024 * 1024;
const MAX_TOTAL_TEXT_SEGMENTS: usize = 8_000_000;
const MAX_TOTAL_RANGES: usize = 4_000_000;
const MAX_RELATIONSHIPS_PER_STORY: usize = 65_536;
const MAX_TOTAL_RELATIONSHIPS: usize = 1_000_000;
const MAX_TOPOLOGY_BYTES: usize = 128 * 1024 * 1024;
const MAX_AUTHOR_CHARS: usize = 255;
const MAX_DATE_CHARS: usize = 128;

/// Validate an `ST_MarkupId` value after lexical integer parsing.
pub(crate) fn validate_id(value: i32) -> Result<()> {
    if value == -1 {
        return Err(Error::Invalid("conflict markup ID -1 is reserved".into()));
    }
    Ok(())
}

/// Validate parsed author/date metadata without normalizing lexical values.
pub(crate) fn validate_metadata(id: Id, author: &str, date: Option<&str>) -> Result<()> {
    validate_id(id.get())?;
    validate_author(author)?;
    if let Some(date) = date {
        validate_date(date)?;
    }
    Ok(())
}

/// Validate an author attribute value.
pub(crate) fn validate_author(author: &str) -> Result<()> {
    if author.chars().count() > MAX_AUTHOR_CHARS {
        return Err(Error::Invalid(format!(
            "conflict author exceeds {MAX_AUTHOR_CHARS} characters"
        )));
    }
    validate_xml_characters(author, "conflict author")
}

/// Validate a retained lexical date without parsing or canonicalizing it.
pub(crate) fn validate_date(date: &str) -> Result<()> {
    if date.chars().count() > MAX_DATE_CHARS {
        return Err(Error::Invalid(format!(
            "conflict date exceeds {MAX_DATE_CHARS} characters"
        )));
    }
    validate_xml_characters(date, "conflict date")?;
    if !is_xsd_datetime(date) {
        return Err(Error::Invalid(
            "conflict date is not a valid xsd:dateTime lexical value".into(),
        ));
    }
    Ok(())
}

/// Reject XML-forbidden scalar values before they can become emitted markup.
pub(crate) fn validate_xml_characters(value: &str, field: &str) -> Result<()> {
    if value.chars().any(|ch| {
        !matches!(ch, '\u{9}' | '\u{A}' | '\u{D}' | '\u{20}'..='\u{D7FF}' | '\u{E000}'..='\u{FFFD}' | '\u{10000}'..='\u{10FFFF}')
    }) {
        return Err(Error::Invalid(format!("{field} contains an XML-forbidden character")));
    }
    Ok(())
}

/// Validate finite, nonzero, hard-capped parser limits.
pub(crate) fn validate_limits(limits: Limits) -> Result<()> {
    let structural = [
        (
            "max_source_bytes",
            limits.max_source_bytes,
            MAX_SOURCE_BYTES,
        ),
        ("max_events", limits.max_events, MAX_EVENTS),
        ("max_depth", limits.max_depth, MAX_DEPTH),
        ("max_attributes", limits.max_attributes, MAX_ATTRIBUTES),
    ];
    for (name, value, maximum) in structural {
        if value == 0 || value > maximum {
            return Err(Error::Invalid(format!(
                "conflict limit {name} must be in 1..={maximum}, got {value}"
            )));
        }
    }
    let quotas = [
        ("max_conflicts", limits.max_conflicts, MAX_CONFLICTS),
        ("max_ranges", limits.max_ranges, MAX_RANGES),
        ("max_text_bytes", limits.max_text_bytes, MAX_TEXT_BYTES),
        (
            "max_text_segments",
            limits.max_text_segments,
            MAX_TEXT_SEGMENTS,
        ),
        (
            "max_metadata_bytes",
            limits.max_metadata_bytes,
            MAX_METADATA_BYTES,
        ),
        ("max_open_ranges", limits.max_open_ranges, MAX_OPEN_RANGES),
        (
            "max_attribute_bytes",
            limits.max_attribute_bytes,
            MAX_ATTRIBUTE_BYTES,
        ),
        (
            "max_output_bytes",
            limits.max_output_bytes,
            MAX_OUTPUT_BYTES,
        ),
        ("max_stories", limits.max_stories, MAX_STORIES),
        (
            "max_total_story_bytes",
            limits.max_total_story_bytes,
            MAX_TOTAL_STORY_BYTES,
        ),
        (
            "max_total_conflicts",
            limits.max_total_conflicts,
            MAX_TOTAL_CONFLICTS,
        ),
        (
            "max_total_metadata_bytes",
            limits.max_total_metadata_bytes,
            MAX_TOTAL_METADATA_BYTES,
        ),
        (
            "max_total_text_segments",
            limits.max_total_text_segments,
            MAX_TOTAL_TEXT_SEGMENTS,
        ),
        (
            "max_total_ranges",
            limits.max_total_ranges,
            MAX_TOTAL_RANGES,
        ),
        (
            "max_relationships_per_story",
            limits.max_relationships_per_story,
            MAX_RELATIONSHIPS_PER_STORY,
        ),
        (
            "max_total_relationships",
            limits.max_total_relationships,
            MAX_TOTAL_RELATIONSHIPS,
        ),
        (
            "max_topology_bytes",
            limits.max_topology_bytes,
            MAX_TOPOLOGY_BYTES,
        ),
    ];
    for (name, value, maximum) in quotas {
        if value > maximum {
            return Err(Error::Invalid(format!(
                "conflict limit {name} must be in 0..={maximum}, got {value}"
            )));
        }
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn semantic_zero_quotas_are_valid() {
        let limits = Limits {
            max_conflicts: 0,
            max_ranges: 0,
            max_open_ranges: 0,
            max_text_bytes: 0,
            max_metadata_bytes: 0,
            max_total_conflicts: 0,
            max_total_ranges: 0,
            max_total_metadata_bytes: 0,
            ..Limits::default()
        };
        assert!(validate_limits(limits).is_ok());
    }

    #[test]
    fn structural_zero_budgets_are_rejected() {
        let limits = Limits {
            max_events: 0,
            ..Limits::default()
        };
        assert!(validate_limits(limits).is_err());
    }
}

fn is_xsd_datetime(value: &str) -> bool {
    let bytes = value.as_bytes();
    let mut cursor = usize::from(bytes.first() == Some(&b'-'));
    let year_start = cursor;
    while cursor < bytes.len() && bytes[cursor].is_ascii_digit() {
        cursor += 1;
    }
    let year_len = cursor - year_start;
    if year_len < 4
        || bytes[year_start..cursor].iter().all(|digit| *digit == b'0')
        || cursor >= bytes.len()
        || bytes[cursor] != b'-'
    {
        return false;
    }
    let year = number_mod(&bytes[year_start..cursor], 400);
    cursor += 1;
    let month = take_fixed_number(bytes, &mut cursor, 2);
    if !take_byte(bytes, &mut cursor, b'-') {
        return false;
    }
    let day = take_fixed_number(bytes, &mut cursor, 2);
    if !take_byte(bytes, &mut cursor, b'T') {
        return false;
    }
    let hour = take_fixed_number(bytes, &mut cursor, 2);
    if !take_byte(bytes, &mut cursor, b':') {
        return false;
    }
    let minute = take_fixed_number(bytes, &mut cursor, 2);
    if !take_byte(bytes, &mut cursor, b':') {
        return false;
    }
    let second = take_fixed_number(bytes, &mut cursor, 2);
    let (Some(month), Some(day), Some(hour), Some(minute), Some(second)) =
        (month, day, hour, minute, second)
    else {
        return false;
    };
    if month == 0
        || month > 12
        || day == 0
        || day > days_in_month(year, month)
        || minute > 59
        || second > 59
        || hour > 24
        || (hour == 24 && (minute != 0 || second != 0))
    {
        return false;
    }
    if cursor < bytes.len() && bytes[cursor] == b'.' {
        cursor += 1;
        let fraction_start = cursor;
        while cursor < bytes.len() && bytes[cursor].is_ascii_digit() {
            cursor += 1;
        }
        if cursor == fraction_start || hour == 24 {
            return false;
        }
    }
    if cursor == bytes.len() {
        return true;
    }
    if bytes[cursor] == b'Z' {
        return cursor + 1 == bytes.len();
    }
    if !matches!(bytes[cursor], b'+' | b'-') {
        return false;
    }
    cursor += 1;
    let zone_hour = take_fixed_number(bytes, &mut cursor, 2);
    if !take_byte(bytes, &mut cursor, b':') {
        return false;
    }
    let zone_minute = take_fixed_number(bytes, &mut cursor, 2);
    matches!((zone_hour, zone_minute), (Some(hours), Some(minutes)) if hours <= 14 && minutes <= 59 && (hours != 14 || minutes == 0) && cursor == bytes.len())
}

fn take_byte(bytes: &[u8], cursor: &mut usize, expected: u8) -> bool {
    if bytes.get(*cursor) == Some(&expected) {
        *cursor += 1;
        true
    } else {
        false
    }
}

fn take_fixed_number(bytes: &[u8], cursor: &mut usize, width: usize) -> Option<u32> {
    let end = cursor.checked_add(width)?;
    let value = parse_number(bytes.get(*cursor..end)?)?;
    *cursor = end;
    Some(value)
}

fn parse_number(bytes: &[u8]) -> Option<u32> {
    if bytes.is_empty() || !bytes.iter().all(u8::is_ascii_digit) {
        return None;
    }
    bytes.iter().try_fold(0_u32, |value, digit| {
        value.checked_mul(10)?.checked_add(u32::from(digit - b'0'))
    })
}

fn number_mod(bytes: &[u8], modulus: u32) -> u32 {
    bytes.iter().fold(0, |value, digit| {
        (value * 10 + u32::from(digit - b'0')) % modulus
    })
}

fn days_in_month(year_mod_400: u32, month: u32) -> u32 {
    match month {
        1 | 3 | 5 | 7 | 8 | 10 | 12 => 31,
        4 | 6 | 9 | 11 => 30,
        2 if year_mod_400.is_multiple_of(400)
            || (year_mod_400.is_multiple_of(4) && !year_mod_400.is_multiple_of(100)) =>
        {
            29
        },
        2 => 28,
        _ => 0,
    }
}
