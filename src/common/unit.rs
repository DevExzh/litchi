//! Unit conversion utilities.
//!
//! This module provides conversion utilities for various length units used in documents.
//! Based on reference implementations from odfdo library.
//!

use crate::Result;
use std::cmp::Ordering;
use std::fmt;
use std::str::FromStr;

pub const EMUS_PER_INCH: i64 = 914_400;
pub const EMUS_PER_CM: i64 = 360_000;
pub const EMUS_PER_MM: i64 = 36_000;
pub const EMUS_PER_PT: i64 = 12_700;
pub const EMUS_PER_TWIP: i64 = 635;
pub const PPT_MASTER_UNITS_PER_INCH: i64 = 576;

#[inline]
pub fn pt_to_emu_f64(pt: f64) -> i64 {
    (pt * EMUS_PER_PT as f64) as i64
}

#[inline]
pub fn pt_to_emu_i32(pt: i32) -> i32 {
    pt.saturating_mul(EMUS_PER_PT as i32)
}

#[inline]
pub fn pt_f32_to_emu_u32(pt: f32) -> u32 {
    (pt * EMUS_PER_PT as f32) as u32
}

#[inline]
pub fn emu_to_pt_f64(emu: i64) -> f64 {
    emu as f64 / EMUS_PER_PT as f64
}

#[inline]
pub fn px_to_emu(px: u32, dpi: u32) -> i64 {
    ((px as f64) * EMUS_PER_INCH as f64 / dpi as f64) as i64
}

#[inline]
pub fn emu_to_px(emu: i64, dpi: u32) -> u32 {
    ((emu as f64) * dpi as f64 / EMUS_PER_INCH as f64) as u32
}

#[inline]
pub fn px_to_emu_96(px: u32) -> i64 {
    px_to_emu(px, 96)
}

#[inline]
pub fn emu_to_px_96(emu: i64) -> u32 {
    emu_to_px(emu, 96)
}

#[inline]
pub fn twip_to_emu_i64(twips: i64) -> i64 {
    twips.saturating_mul(EMUS_PER_TWIP)
}

#[inline]
pub fn emu_to_twip_i64(emu: i64) -> i64 {
    (emu as f64 / EMUS_PER_TWIP as f64).round() as i64
}

#[inline]
pub fn emu_u32_to_ppt_master_u32(emu: u32) -> u32 {
    ((emu as u64 * PPT_MASTER_UNITS_PER_INCH as u64) / EMUS_PER_INCH as u64) as u32
}

#[inline]
pub fn emu_i32_to_ppt_master_i16_round(emu: i32) -> i16 {
    ((emu as f64) * PPT_MASTER_UNITS_PER_INCH as f64 / EMUS_PER_INCH as f64).round() as i16
}

#[inline]
pub fn ppt_master_i64_to_emu_i32(master: i64) -> i32 {
    ((master * EMUS_PER_INCH) / PPT_MASTER_UNITS_PER_INCH) as i32
}

/// Supported length units
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LengthUnit {
    /// Millimeter
    Millimeter,
    /// Centimeter
    Centimeter,
    /// Meter
    Meter,
    /// Kilometer
    Kilometer,
    /// Point (1/72 inch)
    Point,
    /// Pica (1/6 inch)
    Pica,
    /// Inch
    Inch,
    /// Foot
    Foot,
    /// Mile
    Mile,
    /// Pixel
    Pixel,
}

impl LengthUnit {
    /// Get the unit abbreviation
    #[inline]
    pub fn as_str(&self) -> &'static str {
        match self {
            Self::Millimeter => "mm",
            Self::Centimeter => "cm",
            Self::Meter => "m",
            Self::Kilometer => "km",
            Self::Point => "pt",
            Self::Pica => "pc",
            Self::Inch => "in",
            Self::Foot => "ft",
            Self::Mile => "mi",
            Self::Pixel => "px",
        }
    }

    /// Parse unit from string
    fn from_str_internal(s: &str) -> Option<Self> {
        match s {
            "mm" => Some(Self::Millimeter),
            "cm" => Some(Self::Centimeter),
            "m" => Some(Self::Meter),
            "km" => Some(Self::Kilometer),
            "pt" => Some(Self::Point),
            "pc" => Some(Self::Pica),
            "in" | "inch" => Some(Self::Inch),
            "ft" => Some(Self::Foot),
            "mi" => Some(Self::Mile),
            "px" => Some(Self::Pixel),
            _ => None,
        }
    }
}

impl FromStr for LengthUnit {
    type Err = crate::Error;

    fn from_str(s: &str) -> Result<Self> {
        Self::from_str_internal(s)
            .ok_or_else(|| crate::Error::Other(format!("Unknown length unit '{}'", s)))
    }
}

impl fmt::Display for LengthUnit {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{}", self.as_str())
    }
}

/// Length value with unit
///
/// Represents a length measurement with a numeric value and a unit.
/// Supports parsing from strings (e.g., "2.5cm", "10pt") and conversion between units.
///
/// # Examples
///
/// ```
/// use litchi::common::unit::{Length, LengthUnit};
///
/// // Parse from string
/// let length = "2.5cm".parse::<Length>().unwrap();
/// assert_eq!(length.value(), 2.5);
/// assert_eq!(length.unit(), LengthUnit::Centimeter);
///
/// // Convert to pixels (at 96 DPI)
/// let pixels = length.to_pixels(96).unwrap();
/// assert!(pixels.value() > 0.0);
///
/// // Create from value and unit
/// let length = Length::new(10.0, LengthUnit::Point);
/// assert_eq!(length.to_string(), "10pt");
/// ```
#[derive(Debug, Clone, Copy)]
pub struct Length {
    value: f64,
    unit: LengthUnit,
}

impl Length {
    /// Create a new length measurement
    ///
    /// # Arguments
    ///
    /// * `value` - Numeric value
    /// * `unit` - Length unit
    #[inline]
    pub fn new(value: f64, unit: LengthUnit) -> Self {
        Self { value, unit }
    }

    /// Get the numeric value
    #[inline]
    pub fn value(&self) -> f64 {
        self.value
    }

    /// Get the unit
    #[inline]
    pub fn unit(&self) -> LengthUnit {
        self.unit
    }

    /// Convert to pixels with given DPI
    ///
    /// # Arguments
    ///
    /// * `dpi` - Dots per inch (typically 72, 96, or 300)
    ///
    /// # Returns
    ///
    /// `Ok(Length)` with unit Pixel, or `Err` if conversion fails
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi::common::unit::{Length, LengthUnit};
    ///
    /// let inch = Length::new(1.0, LengthUnit::Inch);
    /// let pixels = inch.to_pixels(96).unwrap();
    /// assert_eq!(pixels.value() as i32, 96);
    ///
    /// let cm = Length::new(2.54, LengthUnit::Centimeter);
    /// let pixels = cm.to_pixels(96).unwrap();
    /// assert_eq!(pixels.value() as i32, 96);
    /// ```
    pub fn to_pixels(&self, dpi: u32) -> Result<Self> {
        let dpi_f64 = dpi as f64;

        let pixel_value = match self.unit {
            LengthUnit::Pixel => self.value,
            LengthUnit::Inch => self.value * dpi_f64,
            LengthUnit::Centimeter => self.value / 2.54 * dpi_f64,
            LengthUnit::Millimeter => self.value / 25.4 * dpi_f64,
            LengthUnit::Meter => self.value / 0.0254 * dpi_f64,
            LengthUnit::Kilometer => self.value / 0.0000254 * dpi_f64,
            LengthUnit::Point => self.value / 72.0 * dpi_f64,
            LengthUnit::Pica => self.value / 6.0 * dpi_f64,
            LengthUnit::Foot => self.value * 12.0 * dpi_f64,
            LengthUnit::Mile => self.value * 63360.0 * dpi_f64,
        };

        Ok(Self::new(pixel_value, LengthUnit::Pixel))
    }

    /// Convert to inches
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi::common::unit::{Length, LengthUnit};
    ///
    /// let cm = Length::new(2.54, LengthUnit::Centimeter);
    /// let inches = cm.to_inches().unwrap();
    /// assert!((inches.value() - 1.0).abs() < 0.001);
    /// ```
    pub fn to_inches(&self) -> Result<Self> {
        let inch_value = match self.unit {
            LengthUnit::Inch => self.value,
            LengthUnit::Centimeter => self.value / 2.54,
            LengthUnit::Millimeter => self.value / 25.4,
            LengthUnit::Meter => self.value / 0.0254,
            LengthUnit::Kilometer => self.value / 0.0000254,
            LengthUnit::Point => self.value / 72.0,
            LengthUnit::Pica => self.value / 6.0,
            LengthUnit::Foot => self.value * 12.0,
            LengthUnit::Mile => self.value * 63360.0,
            LengthUnit::Pixel => {
                return Err(crate::Error::Other(
                    "Cannot convert pixels to inches without DPI information".to_string(),
                ));
            },
        };

        Ok(Self::new(inch_value, LengthUnit::Inch))
    }

    /// Convert to centimeters
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi::common::unit::{Length, LengthUnit};
    ///
    /// let inch = Length::new(1.0, LengthUnit::Inch);
    /// let cm = inch.to_centimeters().unwrap();
    /// assert!((cm.value() - 2.54).abs() < 0.001);
    /// ```
    pub fn to_centimeters(&self) -> Result<Self> {
        let cm_value = match self.unit {
            LengthUnit::Centimeter => self.value,
            LengthUnit::Inch => self.value * 2.54,
            LengthUnit::Millimeter => self.value / 10.0,
            LengthUnit::Meter => self.value * 100.0,
            LengthUnit::Kilometer => self.value * 100000.0,
            LengthUnit::Point => self.value / 72.0 * 2.54,
            LengthUnit::Pica => self.value / 6.0 * 2.54,
            LengthUnit::Foot => self.value * 12.0 * 2.54,
            LengthUnit::Mile => self.value * 63360.0 * 2.54,
            LengthUnit::Pixel => {
                return Err(crate::Error::Other(
                    "Cannot convert pixels to centimeters without DPI information".to_string(),
                ));
            },
        };

        Ok(Self::new(cm_value, LengthUnit::Centimeter))
    }
}

impl FromStr for Length {
    type Err = crate::Error;

    /// Parse length from string (e.g., "2.5cm", "10pt")
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi::common::unit::{Length, LengthUnit};
    ///
    /// let length = "2.5cm".parse::<Length>().unwrap();
    /// assert_eq!(length.value(), 2.5);
    /// assert_eq!(length.unit(), LengthUnit::Centimeter);
    ///
    /// let length = "10pt".parse::<Length>().unwrap();
    /// assert_eq!(length.value(), 10.0);
    /// assert_eq!(length.unit(), LengthUnit::Point);
    /// ```
    fn from_str(s: &str) -> Result<Self> {
        let mut digits = String::new();
        let mut non_digits = String::new();
        let mut seen_dot = false;

        for c in s.chars() {
            if c.is_ascii_digit() || (c == '.' && !seen_dot) || (c == '-' && digits.is_empty()) {
                if c == '.' {
                    seen_dot = true;
                }
                digits.push(c);
            } else {
                non_digits.push(c);
            }
        }

        if digits.is_empty() {
            return Err(crate::Error::Other(format!(
                "No numeric value found in '{}'",
                s
            )));
        }

        let value: f64 = digits.parse().map_err(|_| {
            crate::Error::Other(format!("Failed to parse numeric value from '{}'", s))
        })?;

        let unit = if non_digits.is_empty() {
            LengthUnit::Centimeter // Default to cm
        } else {
            LengthUnit::from_str(&non_digits)?
        };

        Ok(Self::new(value, unit))
    }
}

impl fmt::Display for Length {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{}{}", self.value, self.unit.as_str())
    }
}

impl PartialEq for Length {
    fn eq(&self, other: &Self) -> bool {
        if self.unit != other.unit {
            // Try to convert to common unit for comparison
            if let (Ok(a), Ok(b)) = (self.to_inches(), other.to_inches()) {
                (a.value - b.value).abs() < 1e-10
            } else {
                false
            }
        } else {
            (self.value - other.value).abs() < 1e-10
        }
    }
}

impl PartialOrd for Length {
    fn partial_cmp(&self, other: &Self) -> Option<Ordering> {
        if self.unit != other.unit {
            // Try to convert to common unit for comparison
            if let (Ok(a), Ok(b)) = (self.to_inches(), other.to_inches()) {
                a.value.partial_cmp(&b.value)
            } else {
                None
            }
        } else {
            self.value.partial_cmp(&other.value)
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_length() {
        let length = "2.5cm".parse::<Length>().unwrap();
        assert_eq!(length.value(), 2.5);
        assert_eq!(length.unit(), LengthUnit::Centimeter);

        let length = "10pt".parse::<Length>().unwrap();
        assert_eq!(length.value(), 10.0);
        assert_eq!(length.unit(), LengthUnit::Point);

        let length = "1.5in".parse::<Length>().unwrap();
        assert_eq!(length.value(), 1.5);
        assert_eq!(length.unit(), LengthUnit::Inch);

        // Negative values
        let length = "-5mm".parse::<Length>().unwrap();
        assert_eq!(length.value(), -5.0);
        assert_eq!(length.unit(), LengthUnit::Millimeter);
    }

    #[test]
    fn test_to_pixels() {
        let inch = Length::new(1.0, LengthUnit::Inch);
        let pixels = inch.to_pixels(96).unwrap();
        assert_eq!(pixels.value() as i32, 96);

        let cm = Length::new(2.54, LengthUnit::Centimeter);
        let pixels = cm.to_pixels(96).unwrap();
        assert_eq!(pixels.value() as i32, 96);

        let pt = Length::new(72.0, LengthUnit::Point);
        let pixels = pt.to_pixels(96).unwrap();
        assert_eq!(pixels.value() as i32, 96);
    }

    #[test]
    fn test_to_inches() {
        let cm = Length::new(2.54, LengthUnit::Centimeter);
        let inches = cm.to_inches().unwrap();
        assert!((inches.value() - 1.0).abs() < 0.001);

        let mm = Length::new(25.4, LengthUnit::Millimeter);
        let inches = mm.to_inches().unwrap();
        assert!((inches.value() - 1.0).abs() < 0.001);

        let pt = Length::new(72.0, LengthUnit::Point);
        let inches = pt.to_inches().unwrap();
        assert!((inches.value() - 1.0).abs() < 0.001);
    }

    #[test]
    fn test_to_centimeters() {
        let inch = Length::new(1.0, LengthUnit::Inch);
        let cm = inch.to_centimeters().unwrap();
        assert!((cm.value() - 2.54).abs() < 0.001);

        let mm = Length::new(10.0, LengthUnit::Millimeter);
        let cm = mm.to_centimeters().unwrap();
        assert!((cm.value() - 1.0).abs() < 0.001);
    }

    #[test]
    fn test_comparison() {
        let cm = Length::new(2.54, LengthUnit::Centimeter);
        let inch = Length::new(1.0, LengthUnit::Inch);
        assert_eq!(cm, inch);

        let mm = Length::new(25.4, LengthUnit::Millimeter);
        assert_eq!(mm, inch);
    }

    #[test]
    fn test_display() {
        let length = Length::new(2.5, LengthUnit::Centimeter);
        assert_eq!(length.to_string(), "2.5cm");

        let length = Length::new(10.0, LengthUnit::Point);
        assert_eq!(length.to_string(), "10pt");
    }
}
