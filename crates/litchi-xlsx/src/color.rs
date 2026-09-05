//! Checked `SpreadsheetML` color scalars.
//!
//! `SpreadsheetML` serializes RGB values as four hexadecimal bytes in ARGB
//! order.  Keeping this scalar in the XLSX owner crate gives styles,
//! conditional formatting, and future worksheet owners one exact wire type
//! instead of duplicating loosely validated strings.

use std::fmt;
use std::str::FromStr;

/// A checked four-byte `SpreadsheetML` ARGB value.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Rgb {
    alpha: u8,
    red: u8,
    green: u8,
    blue: u8,
}

impl Rgb {
    /// Construct an opaque RGB value.
    #[must_use]
    pub const fn new(red: u8, green: u8, blue: u8) -> Self {
        Self {
            alpha: 0xFF,
            red,
            green,
            blue,
        }
    }

    /// Construct an ARGB value with an explicit alpha byte.
    #[must_use]
    pub const fn argb(alpha: u8, red: u8, green: u8, blue: u8) -> Self {
        Self {
            alpha,
            red,
            green,
            blue,
        }
    }

    /// Return the alpha byte.
    #[must_use]
    pub const fn alpha(self) -> u8 {
        self.alpha
    }

    /// Return the red byte.
    #[must_use]
    pub const fn red(self) -> u8 {
        self.red
    }

    /// Return the green byte.
    #[must_use]
    pub const fn green(self) -> u8 {
        self.green
    }

    /// Return the blue byte.
    #[must_use]
    pub const fn blue(self) -> u8 {
        self.blue
    }

    /// Parse a hexadecimal color in `RRGGBB` or canonical `AARRGGBB` form.
    ///
    /// The six-digit form is an authoring convenience treated as opaque
    /// (alpha `FF`); the wire form produced by [`fmt::Display`] is always
    /// eight digits, so a valid eight-digit input round-trips
    /// byte-identically.
    ///
    /// # Errors
    ///
    /// Returns [`ParseRgbError`] unless `value` is exactly six or eight
    /// hexadecimal digits.
    pub fn from_hex(value: &str) -> Result<Self, ParseRgbError> {
        if value.len() == 8 {
            return value.parse();
        }
        if value.len() != 6 {
            return Err(ParseRgbError);
        }
        let mut components = [0u8; 3];
        for (index, pair) in value.as_bytes().as_chunks::<2>().0.iter().enumerate() {
            let component = parse_hex_pair(pair).ok_or(ParseRgbError)?;
            let slot = components.get_mut(index).ok_or(ParseRgbError)?;
            *slot = component;
        }
        let [red, green, blue] = components;
        Ok(Self::new(red, green, blue))
    }
}

impl fmt::Display for Rgb {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(
            formatter,
            "{:02X}{:02X}{:02X}{:02X}",
            self.alpha, self.red, self.green, self.blue
        )
    }
}

impl FromStr for Rgb {
    type Err = ParseRgbError;

    fn from_str(value: &str) -> Result<Self, Self::Err> {
        if value.len() != 8 {
            return Err(ParseRgbError);
        }
        let mut components = [0u8; 4];
        for (index, pair) in value.as_bytes().as_chunks::<2>().0.iter().enumerate() {
            let component = parse_hex_pair(pair).ok_or(ParseRgbError)?;
            let slot = components.get_mut(index).ok_or(ParseRgbError)?;
            *slot = component;
        }
        let [alpha, red, green, blue] = components;
        Ok(Self::argb(alpha, red, green, blue))
    }
}

fn parse_hex_pair(pair: &[u8]) -> Option<u8> {
    let [high, low] = pair else {
        return None;
    };
    hex_nibble(*high)?
        .checked_mul(16)?
        .checked_add(hex_nibble(*low)?)
}

const fn hex_nibble(value: u8) -> Option<u8> {
    match value {
        b'0'..=b'9' => Some(value - b'0'),
        b'a'..=b'f' => Some(value - b'a' + 10),
        b'A'..=b'F' => Some(value - b'A' + 10),
        _ => None,
    }
}

/// Error returned when an RGB token is not exactly eight hexadecimal digits.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ParseRgbError;

impl fmt::Display for ParseRgbError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str("invalid SpreadsheetML RGB color")
    }
}

impl std::error::Error for ParseRgbError {}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn opaque_rgb_uses_ff_alpha_and_canonical_display() {
        let value = Rgb::new(0, 0xA1, 0xFF);
        assert_eq!(value, Rgb::argb(0xFF, 0, 0xA1, 0xFF));
        assert_eq!(value.alpha(), 0xFF);
        assert_eq!(value.red(), 0);
        assert_eq!(value.green(), 0xA1);
        assert_eq!(value.blue(), 0xFF);
        assert_eq!(value.to_string(), "FF00A1FF");
    }

    #[test]
    fn parses_exact_argb_wire_order_case_insensitively() {
        assert_eq!(
            "8000a1FF".parse::<Rgb>(),
            Ok(Rgb::argb(0x80, 0, 0xA1, 0xFF))
        );
        assert_eq!("8000a1FF".parse::<Rgb>().unwrap().to_string(), "8000A1FF");
    }

    #[test]
    fn rejects_non_argb_lengths_and_non_hex_bytes() {
        for value in ["", "00A1FF", "0000A1FF00", "#FF00A1FF", "FF00A1FG"] {
            assert!(value.parse::<Rgb>().is_err(), "accepted {value:?}");
        }
    }

    #[test]
    fn parse_error_is_specific_and_std_error_compatible() {
        let error = "00A1FF".parse::<Rgb>().unwrap_err();
        assert_eq!(error.to_string(), "invalid SpreadsheetML RGB color");
        let source: &dyn std::error::Error = &error;
        assert!(source.source().is_none());
    }
}
