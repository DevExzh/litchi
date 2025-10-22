//! Number format definitions and utilities.

/// Number format information.
///
/// Excel number formats control how cell values are displayed.
/// This includes both built-in formats (IDs 0-163) and custom formats.
#[derive(Debug, Clone)]
pub struct NumberFormat {
    /// Format ID
    pub id: u32,
    /// Format code (e.g., "General", "0.00", "mm/dd/yyyy")
    pub code: String,
}

impl NumberFormat {
    /// Create a new number format.
    #[inline]
    pub fn new(id: u32, code: String) -> Self {
        Self { id, code }
    }

    /// Check if this is a built-in format (ID < 164).
    #[inline]
    pub fn is_builtin(&self) -> bool {
        self.id < 164
    }

    /// Check if this format represents a date/time format.
    ///
    /// This uses heuristics to detect date/time formats based on
    /// the format code string.
    pub fn is_date_format(&self) -> bool {
        is_date_format(&self.code)
    }
}

/// Check if a format code represents a date/time format.
///
/// This function uses the same logic as calamine's `detect_custom_number_format`.
pub fn is_date_format(format: &str) -> bool {
    let mut escaped = false;
    let mut is_quote = false;
    let mut brackets = 0u8;
    let mut prev = ' ';
    let mut hms = false;
    let mut ap = false;

    for s in format.chars() {
        match (s, escaped, is_quote, ap, brackets) {
            (_, true, ..) => escaped = false, // if escaped, ignore
            ('_' | '\\', ..) => escaped = true,
            ('"', _, true, _, _) => is_quote = false,
            (_, _, true, _, _) => (), // inside quotes, skip
            ('"', _, _, _, _) => is_quote = true,
            (';', ..) => return false, // first format only
            ('[', ..) => brackets += 1,
            (']', .., 1) if hms => return false, // TimeDelta, not DateTime
            (']', ..) => brackets = brackets.saturating_sub(1),
            ('a' | 'A', _, _, false, 0) => ap = true,
            ('p' | 'm' | '/' | 'P' | 'M', _, _, true, 0) => return true,
            ('d' | 'm' | 'h' | 'y' | 's' | 'D' | 'M' | 'H' | 'Y' | 'S', _, _, false, 0) => {
                return true;
            },
            _ => {
                if hms && s.eq_ignore_ascii_case(&prev) {
                    // ok ...
                } else {
                    hms = prev == '[' && matches!(s, 'm' | 'h' | 's' | 'M' | 'H' | 'S');
                }
            },
        }
        prev = s;
    }
    false
}

/// Get the format code for a built-in number format ID.
///
/// Returns `None` if the ID is not a recognized built-in format.
/// Built-in formats are Excel's standard formats (0-163).
#[allow(dead_code)] // Reserved for future use
pub(crate) fn builtin_format_code(id: u32) -> Option<&'static str> {
    match id {
        0 => Some("General"),
        1 => Some("0"),
        2 => Some("0.00"),
        3 => Some("#,##0"),
        4 => Some("#,##0.00"),
        9 => Some("0%"),
        10 => Some("0.00%"),
        11 => Some("0.00E+00"),
        12 => Some("# ?/?"),
        13 => Some("# ??/??"),
        14 => Some("mm-dd-yy"),
        15 => Some("d-mmm-yy"),
        16 => Some("d-mmm"),
        17 => Some("mmm-yy"),
        18 => Some("h:mm AM/PM"),
        19 => Some("h:mm:ss AM/PM"),
        20 => Some("h:mm"),
        21 => Some("h:mm:ss"),
        22 => Some("m/d/yy h:mm"),
        37 => Some("#,##0 ;(#,##0)"),
        38 => Some("#,##0 ;[Red](#,##0)"),
        39 => Some("#,##0.00;(#,##0.00)"),
        40 => Some("#,##0.00;[Red](#,##0.00)"),
        45 => Some("mm:ss"),
        46 => Some("[h]:mm:ss"),
        47 => Some("mmss.0"),
        48 => Some("##0.0E+0"),
        49 => Some("@"),
        _ => None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_is_date_format() {
        assert!(is_date_format("DD/MM/YY"));
        assert!(is_date_format("H:MM:SS;@"));
        assert!(is_date_format("m\"M\"d\"D\";@"));
        assert!(is_date_format("[$-404]e\"\\xfc\"m\"\\xfc\"d\"\\xfc\""));
        assert!(is_date_format("ha/p\\\\m"));

        assert!(!is_date_format("#,##0\\ [$\\u20bd-46D]"));
        assert!(!is_date_format(
            "\"Y: \"0.00\"m\";\"Y: \"-0.00\"m\";\"Y: <num>m\";@"
        ));
        assert!(!is_date_format("#,##0\\ [$''u20bd-46D]"));
        assert!(!is_date_format("\"$\"#,##0_);[Red](\"$\"#,##0)"));
        assert!(!is_date_format("0_ ;[Red]\\-0\\ "));
        assert!(!is_date_format("\\Y000000"));
        assert!(!is_date_format("#,##0.0####\" YMD\""));
        assert!(!is_date_format("[h]:mm:ss")); // TimeDelta
        assert!(!is_date_format("[ss]")); // TimeDelta
        assert!(!is_date_format("[m]")); // TimeDelta
    }

    #[test]
    fn test_builtin_format_code() {
        assert_eq!(builtin_format_code(0), Some("General"));
        assert_eq!(builtin_format_code(14), Some("mm-dd-yy"));
        assert_eq!(builtin_format_code(22), Some("m/d/yy h:mm"));
        assert_eq!(builtin_format_code(999), None);
    }
}
