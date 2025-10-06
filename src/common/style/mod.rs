/// Common style and formatting types.
///
/// This module provides unified style types used across different Office formats.

use std::fmt;

/// RGB color representation.
///
/// Represents a color using red, green, and blue components, each in the range 0-255.
///
/// # Examples
///
/// ```rust
/// use litchi::common::RGBColor;
///
/// // Create a red color
/// let red = RGBColor::new(255, 0, 0);
///
/// // Create from hex string
/// let blue = RGBColor::from_hex("0000FF").unwrap();
/// ```
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
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
    /// # Arguments
    ///
    /// * `r` - Red component (0-255)
    /// * `g` - Green component (0-255)
    /// * `b` - Blue component (0-255)
    ///
    /// # Examples
    ///
    /// ```rust
    /// use litchi::common::RGBColor;
    ///
    /// let color = RGBColor::new(255, 128, 0); // Orange
    /// ```
    #[inline]
    pub const fn new(r: u8, g: u8, b: u8) -> Self {
        Self { r, g, b }
    }

    /// Create an RGB color from a hex string.
    ///
    /// # Arguments
    ///
    /// * `hex` - Hex color string (e.g., "FF0000" or "#FF0000")
    ///
    /// # Examples
    ///
    /// ```rust
    /// use litchi::common::RGBColor;
    ///
    /// let red = RGBColor::from_hex("FF0000").unwrap();
    /// let blue = RGBColor::from_hex("#0000FF").unwrap();
    /// ```
    pub fn from_hex(hex: &str) -> Option<Self> {
        let hex = hex.trim_start_matches('#');
        if hex.len() != 6 {
            return None;
        }

        let r = u8::from_str_radix(&hex[0..2], 16).ok()?;
        let g = u8::from_str_radix(&hex[2..4], 16).ok()?;
        let b = u8::from_str_radix(&hex[4..6], 16).ok()?;

        Some(Self::new(r, g, b))
    }

    /// Convert to hex string (without # prefix).
    ///
    /// # Examples
    ///
    /// ```rust
    /// use litchi::common::RGBColor;
    ///
    /// let color = RGBColor::new(255, 0, 0);
    /// assert_eq!(color.to_hex(), "FF0000");
    /// ```
    pub fn to_hex(&self) -> String {
        format!("{:02X}{:02X}{:02X}", self.r, self.g, self.b)
    }
}

impl fmt::Display for RGBColor {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "#{}", self.to_hex())
    }
}

/// Length measurement with units.
///
/// Represents a measurement value used for dimensions, positions, etc.
/// Office formats primarily use EMUs (English Metric Units).
///
/// # Examples
///
/// ```rust
/// use litchi::common::Length;
///
/// // Create from EMUs
/// let length = Length::from_emus(914400); // 1 inch
///
/// // Convert to different units
/// let inches = length.inches();
/// let cm = length.cm();
/// ```
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Length {
    /// Value in EMUs (English Metric Units)
    /// 1 inch = 914,400 EMUs
    /// 1 cm = 360,000 EMUs
    emus: i64,
}

impl Length {
    /// Create a length from EMUs (English Metric Units).
    ///
    /// EMUs are the native unit used in Office Open XML formats.
    /// - 1 inch = 914,400 EMUs
    /// - 1 cm = 360,000 EMUs
    ///
    /// # Examples
    ///
    /// ```rust
    /// use litchi::common::Length;
    ///
    /// let length = Length::from_emus(914400); // 1 inch
    /// ```
    #[inline]
    pub const fn from_emus(emus: i64) -> Self {
        Self { emus }
    }

    /// Create a length from inches.
    ///
    /// # Examples
    ///
    /// ```rust
    /// use litchi::common::Length;
    ///
    /// let length = Length::from_inches(1.0);
    /// ```
    #[inline]
    pub fn from_inches(inches: f64) -> Self {
        Self {
            emus: (inches * 914400.0) as i64,
        }
    }

    /// Create a length from centimeters.
    ///
    /// # Examples
    ///
    /// ```rust
    /// use litchi::common::Length;
    ///
    /// let length = Length::from_cm(2.54); // ~1 inch
    /// ```
    #[inline]
    pub fn from_cm(cm: f64) -> Self {
        Self {
            emus: (cm * 360000.0) as i64,
        }
    }

    /// Get the value in EMUs.
    #[inline]
    pub const fn emus(&self) -> i64 {
        self.emus
    }

    /// Convert to inches.
    #[inline]
    pub fn inches(&self) -> f64 {
        self.emus as f64 / 914400.0
    }

    /// Convert to centimeters.
    #[inline]
    pub fn cm(&self) -> f64 {
        self.emus as f64 / 360000.0
    }

    /// Convert to points (1/72 inch).
    #[inline]
    pub fn points(&self) -> f64 {
        self.inches() * 72.0
    }
}

impl fmt::Display for Length {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{:.2}\"", self.inches())
    }
}

