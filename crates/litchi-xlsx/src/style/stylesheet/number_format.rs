//! Number format definitions and utilities.

use std::fmt;

/// First `numFmtId` available to user-defined number formats.
///
/// ECMA-376 §18.8.30 lists the built-in formats whose `formatCode` is
/// implied rather than saved; IDs `0..=163` are reserved for built-in and
/// locale-specific formats, and Excel assigns user-defined formats starting
/// at 164 (the same convention as Apache POI's
/// `FIRST_USER_DEFINED_FORMAT_INDEX`).
pub const FIRST_CUSTOM_ID: u32 = 164;

/// Maximum length of a format code in Unicode scalar values.
///
/// `formatCode` is `ST_Xstring` (ECMA-376 §22.9.2.19) and therefore
/// schema-unbounded, but Excel rejects format codes longer than 255
/// characters; enforcing the application limit keeps authored formats
/// storable by Excel.
const MAX_CODE_LENGTH: usize = 255;

/// Error returned when a number format fails authoring validation.
///
/// Deserialization never produces this error: [`NumberFormat::from_raw`]
/// preserves authored file content losslessly (ADR-0006).
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum InvalidNumberFormat {
    /// The format code is empty; `formatCode` must name a display format.
    EmptyCode,
    /// The format code exceeds Excel's 255-character limit.
    CodeTooLong {
        /// Number of Unicode scalar values in the rejected code.
        length: usize,
    },
    /// A user-defined format used an ID from the reserved built-in range
    /// `0..=163`; custom formats must use [`FIRST_CUSTOM_ID`] or above.
    ReservedId {
        /// The rejected reserved identifier.
        id: u32,
    },
}

impl fmt::Display for InvalidNumberFormat {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::EmptyCode => formatter.write_str("number-format code must not be empty"),
            Self::CodeTooLong { length } => write!(
                formatter,
                "number-format code has {length} characters, exceeding Excel's \
                 {MAX_CODE_LENGTH}-character limit"
            ),
            Self::ReservedId { id } => write!(
                formatter,
                "number-format ID {id} is in the reserved built-in range 0..={}",
                FIRST_CUSTOM_ID - 1
            ),
        }
    }
}

impl std::error::Error for InvalidNumberFormat {}

/// Number format information.
///
/// Excel number formats control how cell values are displayed.
/// This includes both built-in formats (IDs below [`FIRST_CUSTOM_ID`]) and
/// custom formats. Construction through [`NumberFormat::new`] and
/// [`NumberFormat::custom`] validates the format code (ADR-0004); the
/// stylesheet deserializer uses [`NumberFormat::from_raw`] to preserve
/// authored content losslessly.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct NumberFormat {
    id: u32,
    code: String,
}

impl NumberFormat {
    /// Create a number format with a validated format code.
    ///
    /// # Errors
    ///
    /// Returns [`InvalidNumberFormat::EmptyCode`] for an empty code and
    /// [`InvalidNumberFormat::CodeTooLong`] for a code above Excel's
    /// 255-character limit.
    pub fn new(id: u32, code: String) -> Result<Self, InvalidNumberFormat> {
        if code.is_empty() {
            return Err(InvalidNumberFormat::EmptyCode);
        }
        let length = code.chars().count();
        if length > MAX_CODE_LENGTH {
            return Err(InvalidNumberFormat::CodeTooLong { length });
        }
        Ok(Self { id, code })
    }

    /// Create a user-defined number format, reserving the built-in ID range.
    ///
    /// # Errors
    ///
    /// Returns [`InvalidNumberFormat::ReservedId`] when `id` is below
    /// [`FIRST_CUSTOM_ID`], in addition to the code errors of
    /// [`NumberFormat::new`].
    pub fn custom(id: u32, code: String) -> Result<Self, InvalidNumberFormat> {
        if id < FIRST_CUSTOM_ID {
            return Err(InvalidNumberFormat::ReservedId { id });
        }
        Self::new(id, code)
    }

    /// Construct without validation for the stylesheet deserializer.
    ///
    /// Real files may declare format codes the authoring API rejects
    /// (ECMA-376 §18.8.30 marks unlisted IDs as implementation-defined);
    /// ADR-0006 requires reading to preserve such content rather than newly
    /// fail, so this constructor stays crate-private and the writer echoes
    /// the stored values verbatim.
    pub(crate) fn from_raw(id: u32, code: String) -> Self {
        Self { id, code }
    }

    /// Format ID.
    #[must_use]
    #[inline]
    pub const fn id(&self) -> u32 {
        self.id
    }

    /// Format code (e.g., "General", "0.00", "mm/dd/yyyy").
    #[must_use]
    #[inline]
    pub fn code(&self) -> &str {
        &self.code
    }

    /// Check if this is a built-in format (ID below [`FIRST_CUSTOM_ID`]).
    #[must_use]
    #[inline]
    pub const fn is_builtin(&self) -> bool {
        self.id < FIRST_CUSTOM_ID
    }

    /// Check if this format represents a date/time format.
    ///
    /// This uses heuristics to detect date/time formats based on
    /// the format code string.
    #[must_use]
    pub fn is_date_format(&self) -> bool {
        is_date_format(&self.code)
    }
}

/// Check if a format code represents a date/time format.
///
/// This function uses the same logic as the internal `detect_custom_number_format`.
#[must_use]
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
#[allow(
    dead_code,
    reason = "reserved for parsing additional built-in number-format families"
)]
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
    fn test_number_format_new() {
        let format = NumberFormat::new(14, "mm-dd-yy".to_string());
        assert_eq!(
            format.map(|value| (value.id(), value.code().to_string())),
            Ok((14, "mm-dd-yy".to_string()))
        );
    }

    #[test]
    fn test_number_format_validation() {
        assert_eq!(
            NumberFormat::new(164, String::new()),
            Err(InvalidNumberFormat::EmptyCode)
        );

        let oversized = "0".repeat(256);
        assert_eq!(
            NumberFormat::new(164, oversized),
            Err(InvalidNumberFormat::CodeTooLong { length: 256 })
        );
        assert!(NumberFormat::new(164, "0".repeat(255)).is_ok());

        assert_eq!(
            NumberFormat::custom(0, "General".to_string()),
            Err(InvalidNumberFormat::ReservedId { id: 0 })
        );
        assert_eq!(
            NumberFormat::custom(163, "0.00".to_string()),
            Err(InvalidNumberFormat::ReservedId { id: 163 })
        );
        assert!(NumberFormat::custom(FIRST_CUSTOM_ID, "0.00".to_string()).is_ok());
        assert_eq!(
            InvalidNumberFormat::ReservedId { id: 14 }.to_string(),
            "number-format ID 14 is in the reserved built-in range 0..=163"
        );
    }

    #[test]
    fn test_number_format_is_builtin() {
        let builtin = NumberFormat::from_raw(100, "0.00".to_string());
        assert!(builtin.is_builtin());

        let custom = NumberFormat::from_raw(164, "Custom Format".to_string());
        assert!(!custom.is_builtin());

        let custom2 = NumberFormat::from_raw(200, "Another Custom".to_string());
        assert!(!custom2.is_builtin());
    }

    #[test]
    fn test_number_format_is_date_format() {
        let date_format = NumberFormat::from_raw(14, "mm-dd-yy".to_string());
        assert!(date_format.is_date_format());

        let time_format = NumberFormat::from_raw(20, "h:mm".to_string());
        assert!(time_format.is_date_format());

        let number_format = NumberFormat::from_raw(1, "0".to_string());
        assert!(!number_format.is_date_format());
    }

    #[test]
    fn test_number_format_clone() {
        let format = NumberFormat::from_raw(2, "0.00".to_string());
        let format2 = format.clone();
        assert_eq!(format.id(), format2.id());
        assert_eq!(format.code(), format2.code());
    }

    #[test]
    fn test_is_date_format() {
        assert!(is_date_format("DD/MM/YY"));
        assert!(is_date_format("H:MM:SS;@"));
        assert!(is_date_format("m\"M\"d\"D\";@"));
        assert!(is_date_format("[$-404]e\"\\xfc\"m\"\\xfc\"d\"\\xfc\""));
        assert!(is_date_format("ha/p\\m"));
        assert!(is_date_format("yyyy-mm-dd"));
        assert!(is_date_format("mm/dd/yyyy"));
        assert!(is_date_format("h:mm:ss"));
        assert!(is_date_format("AM/PM"));
        assert!(is_date_format("d-mmm-yy"));

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
        assert!(!is_date_format("General"));
        assert!(!is_date_format("0.00"));
        assert!(!is_date_format("#,##0"));
    }

    #[test]
    fn test_is_date_format_edge_cases() {
        // Empty string
        assert!(!is_date_format(""));

        // Just quotes
        assert!(!is_date_format("\"\""));

        // Escaped characters
        assert!(!is_date_format("\\d\\m\\y"));

        // Semicolon (multiple formats)
        assert!(!is_date_format("0.00;-0.00"));
    }

    #[test]
    fn test_builtin_format_code() {
        assert_eq!(builtin_format_code(0), Some("General"));
        assert_eq!(builtin_format_code(1), Some("0"));
        assert_eq!(builtin_format_code(2), Some("0.00"));
        assert_eq!(builtin_format_code(3), Some("#,##0"));
        assert_eq!(builtin_format_code(4), Some("#,##0.00"));
        assert_eq!(builtin_format_code(9), Some("0%"));
        assert_eq!(builtin_format_code(10), Some("0.00%"));
        assert_eq!(builtin_format_code(11), Some("0.00E+00"));
        assert_eq!(builtin_format_code(12), Some("# ?/?"));
        assert_eq!(builtin_format_code(13), Some("# ??/??"));
        assert_eq!(builtin_format_code(14), Some("mm-dd-yy"));
        assert_eq!(builtin_format_code(15), Some("d-mmm-yy"));
        assert_eq!(builtin_format_code(16), Some("d-mmm"));
        assert_eq!(builtin_format_code(17), Some("mmm-yy"));
        assert_eq!(builtin_format_code(18), Some("h:mm AM/PM"));
        assert_eq!(builtin_format_code(19), Some("h:mm:ss AM/PM"));
        assert_eq!(builtin_format_code(20), Some("h:mm"));
        assert_eq!(builtin_format_code(21), Some("h:mm:ss"));
        assert_eq!(builtin_format_code(22), Some("m/d/yy h:mm"));
        assert_eq!(builtin_format_code(37), Some("#,##0 ;(#,##0)"));
        assert_eq!(builtin_format_code(38), Some("#,##0 ;[Red](#,##0)"));
        assert_eq!(builtin_format_code(39), Some("#,##0.00;(#,##0.00)"));
        assert_eq!(builtin_format_code(40), Some("#,##0.00;[Red](#,##0.00)"));
        assert_eq!(builtin_format_code(45), Some("mm:ss"));
        assert_eq!(builtin_format_code(46), Some("[h]:mm:ss"));
        assert_eq!(builtin_format_code(47), Some("mmss.0"));
        assert_eq!(builtin_format_code(48), Some("##0.0E+0"));
        assert_eq!(builtin_format_code(49), Some("@"));

        // Non-existent format codes
        assert_eq!(builtin_format_code(5), None);
        assert_eq!(builtin_format_code(999), None);
        assert_eq!(builtin_format_code(163), None);
        assert_eq!(builtin_format_code(164), None);
    }
}
