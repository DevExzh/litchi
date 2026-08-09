//! A1 notation conversion for checked coordinate values.

use super::model::{CellCoord, CellRange};
use litchi_core::{Error, Result};
use std::fmt;
use std::str::FromStr;

impl CellCoord {
    /// Convert this checked coordinate to A1 notation.
    #[must_use]
    pub fn to_a1(&self) -> String {
        format!("{}{}", digit_to_alpha(self.column()), self.row() + 1)
    }
}

impl FromStr for CellCoord {
    type Err = Error;

    /// Parse a cell coordinate from A1 notation.
    fn from_str(value: &str) -> Result<Self> {
        let mut alpha = String::new();
        let mut rest_start = 0;

        for (index, character) in value.char_indices() {
            if character.is_ascii_alphabetic() {
                alpha.push(character);
                rest_start = index + 1;
            } else {
                break;
            }
        }

        if alpha.is_empty() {
            return Err(Error::Other(format!("No column letter found in '{value}'")));
        }

        let numeric = &value[rest_start..];
        if numeric.is_empty() {
            return Err(Error::Other(format!("No row number found in '{value}'")));
        }

        let column = alpha_to_digit(&alpha)?;
        let row: usize = numeric.parse().map_err(|_parse_error| {
            Error::Other(format!("Failed to parse row number from '{numeric}'"))
        })?;
        if row == 0 {
            return Err(Error::Other("Row number must be >= 1".to_string()));
        }

        CellCoord::new(column, row - 1)
    }
}

impl fmt::Display for CellCoord {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(formatter, "{}", self.to_a1())
    }
}

impl FromStr for CellRange {
    type Err = Error;

    /// Parse an inclusive cell range from `A1:B3` notation.
    fn from_str(value: &str) -> Result<Self> {
        let Some((start_text, end_text)) = value.split_once(':') else {
            return Err(Error::Other(format!(
                "Invalid range format '{value}', expected 'A1:B3'"
            )));
        };
        if end_text.contains(':') {
            return Err(Error::Other(format!(
                "Invalid range format '{value}', expected 'A1:B3'"
            )));
        }

        let start = start_text.trim().parse::<CellCoord>()?;
        let end = end_text.trim().parse::<CellCoord>()?;
        CellRange::new(start, end)
    }
}

impl fmt::Display for CellRange {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(formatter, "{}:{}", self.start(), self.end())
    }
}

/// Convert an alphabetic A1 column label to a zero-based index.
///
/// The conversion is case-insensitive and uses checked arithmetic, so an
/// adversarially long label returns an error instead of overflowing.
///
/// # Errors
///
/// Returns an error when the label contains a non-alphabetic character or its
/// value does not fit in `usize`.
///
/// # Examples
///
/// ```
/// use litchi_odf_common::coordinates::alpha_to_digit;
///
/// assert_eq!(alpha_to_digit("A").unwrap(), 0);
/// assert_eq!(alpha_to_digit("Z").unwrap(), 25);
/// assert_eq!(alpha_to_digit("AA").unwrap(), 26);
/// assert_eq!(alpha_to_digit("AB").unwrap(), 27);
/// ```
pub fn alpha_to_digit(alpha: &str) -> Result<usize> {
    if alpha.is_empty() || !alpha.bytes().all(|byte| byte.is_ascii_alphabetic()) {
        return Err(malformed_column(alpha));
    }

    let mut bytes = alpha.bytes();
    let Some(first) = bytes.next() else {
        return Err(malformed_column(alpha));
    };
    let mut column = usize::from(first.to_ascii_uppercase() - b'A');
    for byte in bytes {
        let value = usize::from(byte.to_ascii_uppercase() - b'A');
        column = column
            .checked_add(1)
            .and_then(|value_before_digit| value_before_digit.checked_mul(26))
            .and_then(|value_before_digit| value_before_digit.checked_add(value))
            .ok_or_else(|| oversized_column(alpha))?;
    }

    Ok(column)
}

/// Convert a zero-based column index to alphabetic A1 notation.
///
/// The loop uses the zero-based form directly, avoiding the overflow that a
/// naïve `digit + 1` implementation encounters at the edge of `usize`.
#[must_use]
pub fn digit_to_alpha(mut digit: usize) -> String {
    let mut column = String::new();
    loop {
        let remainder = u8::try_from(digit % 26).unwrap_or_default();
        column.push(char::from(b'A' + remainder));
        if digit < 26 {
            break;
        }
        digit = digit / 26 - 1;
    }
    column.chars().rev().collect()
}

fn malformed_column(alpha: &str) -> Error {
    Error::Other(format!(
        "Column value '{alpha}' is malformed, must contain only letters"
    ))
}

fn oversized_column(alpha: &str) -> Error {
    Error::Other(format!(
        "Column value '{alpha}' exceeds the platform coordinate range"
    ))
}
