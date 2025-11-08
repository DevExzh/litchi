//! Cell coordinate conversion utilities (A1 notation).
//!
//! This module provides utilities for converting between cell coordinates in different formats:
//! - A1 notation (e.g., "A1", "B3", "AA10")
//! - Numeric coordinates (column, row as integers)
//! - Range notation (e.g., "A1:B3")
//!
//! # Implementation Status
//!
//! ✅ COMPLETED: Alpha to digit conversion (A -> 0, B -> 1, AA -> 26, etc.)
//! ✅ COMPLETED: Digit to alpha conversion (0 -> A, 1 -> B, 26 -> AA, etc.)
//! ✅ COMPLETED: A1 notation parsing and formatting
//! ✅ COMPLETED: Range parsing (A1:B3)
//!
//! # References
//!
//! - odfdo: `3rdparty/odfdo/src/odfdo/utils/coordinates.py`

use crate::Result;
use std::fmt;
use std::str::FromStr;

/// Convert alphabetic column to numeric (0-indexed)
///
/// # Arguments
///
/// * `alpha` - Column in alphabetic notation (e.g., "A", "Z", "AA")
///
/// # Returns
///
/// Column index (0-indexed): A=0, B=1, ..., Z=25, AA=26, etc.
///
/// # Examples
///
/// ```
/// use litchi::odf::coordinates::alpha_to_digit;
///
/// assert_eq!(alpha_to_digit("A").unwrap(), 0);
/// assert_eq!(alpha_to_digit("Z").unwrap(), 25);
/// assert_eq!(alpha_to_digit("AA").unwrap(), 26);
/// assert_eq!(alpha_to_digit("AB").unwrap(), 27);
/// ```
pub fn alpha_to_digit(alpha: &str) -> Result<usize> {
    if alpha.is_empty() || !alpha.chars().all(|c| c.is_ascii_alphabetic()) {
        return Err(crate::Error::Other(format!(
            "Column value '{}' is malformed, must contain only letters",
            alpha
        )));
    }

    let mut column = 0usize;
    for c in alpha.to_uppercase().chars() {
        let val = (c as u32 - b'A' as u32 + 1) as usize;
        column = column * 26 + val;
    }

    Ok(column - 1)
}

/// Convert numeric column to alphabetic notation (0-indexed)
///
/// # Arguments
///
/// * `digit` - Column index (0-indexed)
///
/// # Returns
///
/// Column in alphabetic notation: 0=A, 1=B, ..., 25=Z, 26=AA, etc.
///
/// # Examples
///
/// ```
/// use litchi::odf::coordinates::digit_to_alpha;
///
/// assert_eq!(digit_to_alpha(0), "A");
/// assert_eq!(digit_to_alpha(25), "Z");
/// assert_eq!(digit_to_alpha(26), "AA");
/// assert_eq!(digit_to_alpha(27), "AB");
/// ```
pub fn digit_to_alpha(mut digit: usize) -> String {
    let mut column = String::new();
    digit += 1; // Convert from 0-indexed to 1-indexed for calculation

    while digit > 0 {
        let c = ((digit - 1) % 26) as u8;
        column.insert(0, (b'A' + c) as char);
        digit = (digit - 1) / 26;
    }

    column
}

/// Cell coordinates (column, row) both 0-indexed
///
/// # Examples
///
/// ```
/// use litchi::odf::coordinates::CellCoord;
///
/// let coord = CellCoord::new(0, 0); // A1
/// assert_eq!(coord.to_string(), "A1");
///
/// let coord = CellCoord::new(1, 2); // B3
/// assert_eq!(coord.to_string(), "B3");
///
/// let coord: CellCoord = "AA10".parse().unwrap();
/// assert_eq!(coord.column(), 26);
/// assert_eq!(coord.row(), 9);
/// ```
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct CellCoord {
    column: usize,
    row: usize,
}

impl CellCoord {
    /// Create a new cell coordinate
    ///
    /// # Arguments
    ///
    /// * `column` - Column index (0-indexed)
    /// * `row` - Row index (0-indexed)
    #[inline]
    pub const fn new(column: usize, row: usize) -> Self {
        Self { column, row }
    }

    /// Get column index (0-indexed)
    #[inline]
    pub const fn column(&self) -> usize {
        self.column
    }

    /// Get row index (0-indexed)
    #[inline]
    pub const fn row(&self) -> usize {
        self.row
    }

    /// Convert to A1 notation string
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi::odf::coordinates::CellCoord;
    ///
    /// let coord = CellCoord::new(0, 0);
    /// assert_eq!(coord.to_a1(), "A1");
    ///
    /// let coord = CellCoord::new(26, 9);
    /// assert_eq!(coord.to_a1(), "AA10");
    /// ```
    pub fn to_a1(&self) -> String {
        format!("{}{}", digit_to_alpha(self.column), self.row + 1)
    }
}

impl FromStr for CellCoord {
    type Err = crate::Error;

    /// Parse cell coordinate from A1 notation
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi::odf::coordinates::CellCoord;
    ///
    /// let coord: CellCoord = "A1".parse().unwrap();
    /// assert_eq!(coord.column(), 0);
    /// assert_eq!(coord.row(), 0);
    ///
    /// let coord: CellCoord = "B3".parse().unwrap();
    /// assert_eq!(coord.column(), 1);
    /// assert_eq!(coord.row(), 2);
    ///
    /// let coord: CellCoord = "AA10".parse().unwrap();
    /// assert_eq!(coord.column(), 26);
    /// assert_eq!(coord.row(), 9);
    /// ```
    fn from_str(s: &str) -> Result<Self> {
        // Extract alpha part
        let mut alpha = String::new();
        let mut rest_start = 0;

        for (i, c) in s.char_indices() {
            if c.is_ascii_alphabetic() {
                alpha.push(c);
                rest_start = i + 1;
            } else {
                break;
            }
        }

        if alpha.is_empty() {
            return Err(crate::Error::Other(format!(
                "No column letter found in '{}'",
                s
            )));
        }

        // Extract numeric part
        let numeric = &s[rest_start..];
        if numeric.is_empty() {
            return Err(crate::Error::Other(format!(
                "No row number found in '{}'",
                s
            )));
        }

        let column = alpha_to_digit(&alpha)?;
        let row: usize = numeric.parse().map_err(|_| {
            crate::Error::Other(format!("Failed to parse row number from '{}'", numeric))
        })?;

        if row == 0 {
            return Err(crate::Error::Other("Row number must be >= 1".to_string()));
        }

        Ok(Self::new(column, row - 1))
    }
}

impl fmt::Display for CellCoord {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{}", self.to_a1())
    }
}

/// Cell range (start cell, end cell)
///
/// # Examples
///
/// ```
/// use litchi::odf::coordinates::{CellCoord, CellRange};
///
/// let range = CellRange::new(
///     CellCoord::new(0, 0),
///     CellCoord::new(1, 2),
/// );
/// assert_eq!(range.to_string(), "A1:B3");
///
/// let range: CellRange = "A1:B3".parse().unwrap();
/// assert_eq!(range.start().column(), 0);
/// assert_eq!(range.start().row(), 0);
/// assert_eq!(range.end().column(), 1);
/// assert_eq!(range.end().row(), 2);
/// ```
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct CellRange {
    start: CellCoord,
    end: CellCoord,
}

impl CellRange {
    /// Create a new cell range
    ///
    /// # Arguments
    ///
    /// * `start` - Start cell coordinate
    /// * `end` - End cell coordinate
    #[inline]
    pub const fn new(start: CellCoord, end: CellCoord) -> Self {
        Self { start, end }
    }

    /// Get start cell coordinate
    #[inline]
    pub const fn start(&self) -> CellCoord {
        self.start
    }

    /// Get end cell coordinate
    #[inline]
    pub const fn end(&self) -> CellCoord {
        self.end
    }

    /// Get the number of columns in the range
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi::odf::coordinates::CellRange;
    ///
    /// let range: CellRange = "A1:C3".parse().unwrap();
    /// assert_eq!(range.width(), 3);
    /// ```
    #[inline]
    pub fn width(&self) -> usize {
        if self.end.column >= self.start.column {
            self.end.column - self.start.column + 1
        } else {
            0
        }
    }

    /// Get the number of rows in the range
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi::odf::coordinates::CellRange;
    ///
    /// let range: CellRange = "A1:C3".parse().unwrap();
    /// assert_eq!(range.height(), 3);
    /// ```
    #[inline]
    pub fn height(&self) -> usize {
        if self.end.row >= self.start.row {
            self.end.row - self.start.row + 1
        } else {
            0
        }
    }
}

impl FromStr for CellRange {
    type Err = crate::Error;

    /// Parse cell range from A1:B3 notation
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi::odf::coordinates::CellRange;
    ///
    /// let range: CellRange = "A1:B3".parse().unwrap();
    /// assert_eq!(range.start().column(), 0);
    /// assert_eq!(range.start().row(), 0);
    /// assert_eq!(range.end().column(), 1);
    /// assert_eq!(range.end().row(), 2);
    /// ```
    fn from_str(s: &str) -> Result<Self> {
        let parts: Vec<&str> = s.split(':').collect();

        if parts.len() != 2 {
            return Err(crate::Error::Other(format!(
                "Invalid range format '{}', expected 'A1:B3'",
                s
            )));
        }

        let start = CellCoord::from_str(parts[0].trim())?;
        let end = CellCoord::from_str(parts[1].trim())?;

        Ok(Self::new(start, end))
    }
}

impl fmt::Display for CellRange {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{}:{}", self.start.to_a1(), self.end.to_a1())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_alpha_to_digit() {
        assert_eq!(alpha_to_digit("A").unwrap(), 0);
        assert_eq!(alpha_to_digit("B").unwrap(), 1);
        assert_eq!(alpha_to_digit("Z").unwrap(), 25);
        assert_eq!(alpha_to_digit("AA").unwrap(), 26);
        assert_eq!(alpha_to_digit("AB").unwrap(), 27);
        assert_eq!(alpha_to_digit("AZ").unwrap(), 51);
        assert_eq!(alpha_to_digit("BA").unwrap(), 52);

        // Case insensitive
        assert_eq!(alpha_to_digit("a").unwrap(), 0);
        assert_eq!(alpha_to_digit("aa").unwrap(), 26);

        // Errors
        assert!(alpha_to_digit("").is_err());
        assert!(alpha_to_digit("A1").is_err());
        assert!(alpha_to_digit("1A").is_err());
    }

    #[test]
    fn test_digit_to_alpha() {
        assert_eq!(digit_to_alpha(0), "A");
        assert_eq!(digit_to_alpha(1), "B");
        assert_eq!(digit_to_alpha(25), "Z");
        assert_eq!(digit_to_alpha(26), "AA");
        assert_eq!(digit_to_alpha(27), "AB");
        assert_eq!(digit_to_alpha(51), "AZ");
        assert_eq!(digit_to_alpha(52), "BA");
    }

    #[test]
    fn test_round_trip() {
        for i in 0..100 {
            let alpha = digit_to_alpha(i);
            let digit = alpha_to_digit(&alpha).unwrap();
            assert_eq!(digit, i);
        }
    }

    #[test]
    fn test_cell_coord_parse() {
        let coord: CellCoord = "A1".parse().unwrap();
        assert_eq!(coord.column(), 0);
        assert_eq!(coord.row(), 0);

        let coord: CellCoord = "B3".parse().unwrap();
        assert_eq!(coord.column(), 1);
        assert_eq!(coord.row(), 2);

        let coord: CellCoord = "AA10".parse().unwrap();
        assert_eq!(coord.column(), 26);
        assert_eq!(coord.row(), 9);

        // Errors
        assert!("A0".parse::<CellCoord>().is_err()); // Row must be >= 1
        assert!("1A".parse::<CellCoord>().is_err()); // No column
        assert!("A".parse::<CellCoord>().is_err()); // No row
    }

    #[test]
    fn test_cell_coord_display() {
        let coord = CellCoord::new(0, 0);
        assert_eq!(coord.to_string(), "A1");

        let coord = CellCoord::new(1, 2);
        assert_eq!(coord.to_string(), "B3");

        let coord = CellCoord::new(26, 9);
        assert_eq!(coord.to_string(), "AA10");
    }

    #[test]
    fn test_cell_range_parse() {
        let range: CellRange = "A1:B3".parse().unwrap();
        assert_eq!(range.start().column(), 0);
        assert_eq!(range.start().row(), 0);
        assert_eq!(range.end().column(), 1);
        assert_eq!(range.end().row(), 2);

        let range: CellRange = "AA10:AB20".parse().unwrap();
        assert_eq!(range.start().column(), 26);
        assert_eq!(range.start().row(), 9);
        assert_eq!(range.end().column(), 27);
        assert_eq!(range.end().row(), 19);

        // Errors
        assert!("A1".parse::<CellRange>().is_err()); // No colon
        assert!("A1:".parse::<CellRange>().is_err()); // Missing end
        assert!(":B3".parse::<CellRange>().is_err()); // Missing start
    }

    #[test]
    fn test_cell_range_display() {
        let range = CellRange::new(CellCoord::new(0, 0), CellCoord::new(1, 2));
        assert_eq!(range.to_string(), "A1:B3");
    }

    #[test]
    fn test_cell_range_dimensions() {
        let range: CellRange = "A1:C3".parse().unwrap();
        assert_eq!(range.width(), 3);
        assert_eq!(range.height(), 3);

        let range: CellRange = "B2:E5".parse().unwrap();
        assert_eq!(range.width(), 4);
        assert_eq!(range.height(), 4);
    }
}
