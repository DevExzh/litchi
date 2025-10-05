/// Shared types and utilities used across OOXML formats.
///
/// This module provides common functionality that can be reused across
/// different Office file formats (docx, xlsx, pptx).
use std::fmt;

/// Length in English Metric Units (EMU).
///
/// EMU is the base unit for measurements in OOXML:
/// - 914,400 EMU = 1 inch
/// - 360,000 EMU = 1 centimeter
/// - 36,000 EMU = 1 millimeter
/// - 12,700 EMU = 1 point
/// - 635 EMU = 1 twip (1/20 of a point)
///
/// # Examples
///
/// ```rust
/// use litchi::ooxml::Length;
///
/// let width = Length::from_inches(1.0);
/// assert_eq!(width.emu(), 914400);
/// assert_eq!(width.inches(), 1.0);
/// ```
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct Length(i64);

impl Length {
    /// EMUs per inch
    pub const EMUS_PER_INCH: i64 = 914_400;
    /// EMUs per centimeter
    pub const EMUS_PER_CM: i64 = 360_000;
    /// EMUs per millimeter
    pub const EMUS_PER_MM: i64 = 36_000;
    /// EMUs per point
    pub const EMUS_PER_PT: i64 = 12_700;
    /// EMUs per twip (1/20 of a point)
    pub const EMUS_PER_TWIP: i64 = 635;

    /// Create a new Length from EMU value.
    ///
    /// # Examples
    ///
    /// ```rust
    /// use litchi::ooxml::Length;
    ///
    /// let len = Length::new(914400);
    /// assert_eq!(len.emu(), 914400);
    /// ```
    #[inline]
    pub const fn new(emu: i64) -> Self {
        Self(emu)
    }

    /// Create a Length from inches.
    ///
    /// # Examples
    ///
    /// ```rust
    /// use litchi::ooxml::Length;
    ///
    /// let width = Length::from_inches(0.5);
    /// assert_eq!(width.emu(), 457200);
    /// ```
    #[inline]
    pub fn from_inches(inches: f64) -> Self {
        Self((inches * Self::EMUS_PER_INCH as f64) as i64)
    }

    /// Create a Length from centimeters.
    #[inline]
    pub fn from_cm(cm: f64) -> Self {
        Self((cm * Self::EMUS_PER_CM as f64) as i64)
    }

    /// Create a Length from millimeters.
    #[inline]
    pub fn from_mm(mm: f64) -> Self {
        Self((mm * Self::EMUS_PER_MM as f64) as i64)
    }

    /// Create a Length from points.
    #[inline]
    pub fn from_pt(pt: f64) -> Self {
        Self((pt * Self::EMUS_PER_PT as f64) as i64)
    }

    /// Create a Length from twips (1/20 of a point).
    #[inline]
    pub fn from_twips(twips: f64) -> Self {
        Self((twips * Self::EMUS_PER_TWIP as f64) as i64)
    }

    /// Get the length in EMU.
    #[inline]
    pub const fn emu(self) -> i64 {
        self.0
    }

    /// Get the length in inches.
    #[inline]
    pub fn inches(self) -> f64 {
        self.0 as f64 / Self::EMUS_PER_INCH as f64
    }

    /// Get the length in centimeters.
    #[inline]
    pub fn cm(self) -> f64 {
        self.0 as f64 / Self::EMUS_PER_CM as f64
    }

    /// Get the length in millimeters.
    #[inline]
    pub fn mm(self) -> f64 {
        self.0 as f64 / Self::EMUS_PER_MM as f64
    }

    /// Get the length in points.
    #[inline]
    pub fn pt(self) -> f64 {
        self.0 as f64 / Self::EMUS_PER_PT as f64
    }

    /// Get the length in twips.
    #[inline]
    pub fn twips(self) -> i64 {
        (self.0 as f64 / Self::EMUS_PER_TWIP as f64).round() as i64
    }
}

impl fmt::Display for Length {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{}emu", self.0)
    }
}

/// RGB color value.
///
/// Represents a color using red, green, and blue components,
/// each in the range 0-255.
///
/// # Examples
///
/// ```rust
/// use litchi::ooxml::RGBColor;
///
/// let red = RGBColor::new(255, 0, 0);
/// let from_hex = RGBColor::from_hex("FF0000").unwrap();
/// assert_eq!(red, from_hex);
/// ```
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct RGBColor {
    /// Red component (0-255)
    pub r: u8,
    /// Green component (0-255)
    pub g: u8,
    /// Blue component (0-255)
    pub b: u8,
}

impl RGBColor {
    /// Create a new RGB color.
    ///
    /// # Examples
    ///
    /// ```rust
    /// use litchi::ooxml::RGBColor;
    ///
    /// let color = RGBColor::new(60, 47, 128);
    /// ```
    #[inline]
    pub const fn new(r: u8, g: u8, b: u8) -> Self {
        Self { r, g, b }
    }

    /// Create an RGB color from a hex string.
    ///
    /// The string should be 6 hex digits without a leading '#'.
    ///
    /// # Examples
    ///
    /// ```rust
    /// use litchi::ooxml::RGBColor;
    ///
    /// let color = RGBColor::from_hex("3C2F80").unwrap();
    /// assert_eq!(color.r, 0x3C);
    /// assert_eq!(color.g, 0x2F);
    /// assert_eq!(color.b, 0x80);
    /// ```
    pub fn from_hex(hex: &str) -> Result<Self, &'static str> {
        if hex.len() != 6 {
            return Err("Hex color must be exactly 6 characters");
        }

        let r =
            u8::from_str_radix(&hex[0..2], 16).map_err(|_| "Invalid hex digit in red component")?;
        let g = u8::from_str_radix(&hex[2..4], 16)
            .map_err(|_| "Invalid hex digit in green component")?;
        let b = u8::from_str_radix(&hex[4..6], 16)
            .map_err(|_| "Invalid hex digit in blue component")?;

        Ok(Self { r, g, b })
    }

    /// Convert to a hex string representation.
    ///
    /// # Examples
    ///
    /// ```rust
    /// use litchi::ooxml::RGBColor;
    ///
    /// let color = RGBColor::new(60, 47, 128);
    /// assert_eq!(color.to_hex(), "3C2F80");
    /// ```
    pub fn to_hex(&self) -> String {
        format!("{:02X}{:02X}{:02X}", self.r, self.g, self.b)
    }
}

impl fmt::Display for RGBColor {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(
            f,
            "RGBColor({:02X}, {:02X}, {:02X})",
            self.r, self.g, self.b
        )
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_length_conversions() {
        let len = Length::from_inches(1.0);
        assert_eq!(len.emu(), 914400);
        assert!((len.inches() - 1.0).abs() < 1e-6);

        let len = Length::from_cm(1.0);
        assert_eq!(len.emu(), 360000);
        assert!((len.cm() - 1.0).abs() < 1e-6);

        let len = Length::from_pt(72.0);
        assert_eq!(len.emu(), 72 * 12700);
        assert!((len.pt() - 72.0).abs() < 1e-6);
    }

    #[test]
    fn test_rgb_color() {
        let color = RGBColor::new(60, 47, 128);
        assert_eq!(color.to_hex(), "3C2F80");

        let from_hex = RGBColor::from_hex("3C2F80").unwrap();
        assert_eq!(color, from_hex);

        assert!(RGBColor::from_hex("GGGGGG").is_err());
        assert!(RGBColor::from_hex("FF00").is_err());
    }
}
