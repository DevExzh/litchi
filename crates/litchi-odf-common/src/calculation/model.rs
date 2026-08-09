//! Semantic calculation-settings values.

use litchi_core::{Error, Result};
use std::num::NonZeroUsize;

/// Whether formula iteration is enabled for the document.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum IterationStatus {
    Enable,
    Disable,
}

impl IterationStatus {
    pub(crate) fn parse(value: &str) -> Result<Self> {
        match value {
            "enable" => Ok(Self::Enable),
            "disable" => Ok(Self::Disable),
            _ => Err(Error::InvalidFormat(format!(
                "invalid table:iteration status '{value}'"
            ))),
        }
    }

    pub(crate) const fn as_str(self) -> &'static str {
        match self {
            Self::Enable => "enable",
            Self::Disable => "disable",
        }
    }
}

/// The document's formula null date.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct NullDate {
    /// Whether `table:value-type="date"` was explicitly present.
    pub value_type_date: bool,
    /// XML Schema date lexical value, preserved without timezone normalization.
    pub date_value: Option<String>,
}

/// Formula iteration limits and status.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct Iteration {
    pub status: Option<IterationStatus>,
    pub steps: Option<NonZeroUsize>,
    /// XML Schema double lexical value, preserving `INF`, `-INF`, and `NaN`.
    pub maximum_difference: Option<String>,
}

/// Spreadsheet-wide formula calculation settings.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct Settings {
    pub case_sensitive: Option<bool>,
    pub precision_as_shown: Option<bool>,
    pub search_criteria_must_apply_to_whole_cell: Option<bool>,
    pub automatic_find_labels: Option<bool>,
    pub use_regular_expressions: Option<bool>,
    pub use_wildcards: Option<bool>,
    pub null_year: Option<NonZeroUsize>,
    pub null_date: Option<NullDate>,
    pub iteration: Option<Iteration>,
}

impl Settings {
    /// Validate all lexical values before they cross an XML boundary.
    ///
    /// # Errors
    ///
    /// Returns an error when a retained XML Schema lexical value is malformed
    /// or exceeds the configured attribute limit.
    pub fn validate(&self) -> Result<()> {
        if let Some(date_value) = self
            .null_date
            .as_ref()
            .and_then(|null_date| null_date.date_value.as_deref())
            && (date_value.len() > super::MAX_ATTRIBUTE_BYTES || !is_xsd_date(date_value))
        {
            return Err(Error::InvalidFormat(format!(
                "invalid calculation null date '{date_value}'"
            )));
        }
        if let Some(maximum_difference) = self
            .iteration
            .as_ref()
            .and_then(|iteration| iteration.maximum_difference.as_deref())
            && (maximum_difference.len() > super::MAX_ATTRIBUTE_BYTES
                || !is_xsd_double(maximum_difference))
        {
            return Err(Error::InvalidFormat(format!(
                "invalid iteration maximum difference '{maximum_difference}'"
            )));
        }
        Ok(())
    }
}

pub(crate) fn is_xsd_double(value: &str) -> bool {
    if matches!(value, "INF" | "-INF" | "NaN") {
        return true;
    }
    let bytes = value.as_bytes();
    let mut index = usize::from(matches!(bytes.first(), Some(b'+' | b'-')));
    let integer_start = index;
    while bytes.get(index).is_some_and(u8::is_ascii_digit) {
        index += 1;
    }
    let has_integer = index > integer_start;
    let mut has_fraction = false;
    if bytes.get(index) == Some(&b'.') {
        index += 1;
        let start = index;
        while bytes.get(index).is_some_and(u8::is_ascii_digit) {
            index += 1;
        }
        has_fraction = index > start;
    }
    if !has_integer && !has_fraction {
        return false;
    }
    if matches!(bytes.get(index), Some(b'e' | b'E')) {
        index += 1;
        if matches!(bytes.get(index), Some(b'+' | b'-')) {
            index += 1;
        }
        let start = index;
        while bytes.get(index).is_some_and(u8::is_ascii_digit) {
            index += 1;
        }
        if index == start {
            return false;
        }
    }
    index == bytes.len()
}

pub(crate) fn is_xsd_date(value: &str) -> bool {
    if !value.is_ascii() {
        return false;
    }
    let date_with_timezone = if let Some(date) = value.strip_suffix('Z') {
        date
    } else if value.len() >= 6 {
        let split = value.len() - 6;
        let suffix = &value[split..];
        let timezone = matches!(suffix.as_bytes().first(), Some(b'+' | b'-'))
            && suffix.as_bytes().get(3) == Some(&b':')
            && suffix[1..3].bytes().all(|byte| byte.is_ascii_digit())
            && suffix[4..6].bytes().all(|byte| byte.is_ascii_digit())
            && suffix[1..3].parse::<u8>().is_ok_and(|hour| hour <= 14)
            && suffix[4..6].parse::<u8>().is_ok_and(|minute| minute <= 59)
            && (suffix[1..3] != *"14" || suffix[4..6] == *"00");
        if timezone { &value[..split] } else { value }
    } else {
        value
    };
    let unsigned_date = date_with_timezone
        .strip_prefix('-')
        .unwrap_or(date_with_timezone);
    let Some((year, rest)) = unsigned_date.split_once('-') else {
        return false;
    };
    let Some((month, day)) = rest.split_once('-') else {
        return false;
    };
    if year.len() < 4
        || !year.bytes().all(|byte| byte.is_ascii_digit())
        || year.bytes().all(|byte| byte == b'0')
        || (year.len() > 4 && year.starts_with('0'))
        || month.len() != 2
        || day.len() != 2
    {
        return false;
    }
    let (Ok(parsed_month), Ok(parsed_day)) = (month.parse::<u8>(), day.parse::<u8>()) else {
        return false;
    };
    let leap =
        decimal_mod(year, 4) == 0 && (decimal_mod(year, 100) != 0 || decimal_mod(year, 400) == 0);
    let days = match parsed_month {
        1 | 3 | 5 | 7 | 8 | 10 | 12 => 31,
        4 | 6 | 9 | 11 => 30,
        2 if leap => 29,
        2 => 28,
        _ => return false,
    };
    (1..=days).contains(&parsed_day)
}

fn decimal_mod(value: &str, modulus: u16) -> u16 {
    value.bytes().fold(0, |remainder, byte| {
        (remainder * 10 + u16::from(byte - b'0')) % modulus
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn accepts_lossless_lexical_forms() {
        for value in ["0", ".5", "5.", "-1.25E-3", "INF", "-INF", "NaN"] {
            assert!(is_xsd_double(value), "{value}");
        }
        for value in ["", ".", "inf", "1E", "1 2"] {
            assert!(!is_xsd_double(value), "{value}");
        }
        for value in ["1899-12-30", "2026-07-14Z", "2026-07-14+08:00"] {
            assert!(is_xsd_date(value), "{value}");
        }
        for value in [
            "2026-07-14+14:01",
            "0000-01-01",
            "02026-01-01",
            "2026-02-29",
        ] {
            assert!(!is_xsd_date(value), "{value}");
        }
        assert!(is_xsd_date("2024-02-29"));
    }

    #[test]
    fn rejects_oversized_lexical_values() {
        let value = "1".repeat(super::super::MAX_ATTRIBUTE_BYTES + 1);
        assert!(
            Settings {
                iteration: Some(Iteration {
                    maximum_difference: Some(value),
                    ..Iteration::default()
                }),
                ..Settings::default()
            }
            .validate()
            .is_err()
        );
    }
}
