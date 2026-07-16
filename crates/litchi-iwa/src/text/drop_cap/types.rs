//! Strict public types for paragraph Drop Caps.

use std::num::NonZeroU32;

use crate::{Error, Result};

const MIN_DROP_CAP_LINES: u8 = 2;
const MAX_DROP_CAP_LINES: u8 = 50;

/// UTF-16 position at which a paragraph begins.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash, Default)]
pub struct ParagraphStart(u32);

impl ParagraphStart {
    pub const ZERO: Self = Self(0);

    /// Construct a paragraph start from an iWork UTF-16 character index.
    pub fn from_utf16_index(index: usize) -> Result<Self> {
        u32::try_from(index).map(Self).map_err(|_| {
            Error::InvalidFormat("paragraph start exceeds the u32 UTF-16 range".to_owned())
        })
    }

    /// Return the iWork UTF-16 character index.
    pub const fn utf16_index(self) -> u32 {
        self.0
    }
}

/// Number of text lines occupied by a Drop Cap.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct DropCapLineCount(u8);

impl DropCapLineCount {
    pub const DEFAULT: Self = Self(3);

    /// Construct the app-supported line count, from 2 through 50.
    pub fn new(lines: u8) -> Result<Self> {
        if !(MIN_DROP_CAP_LINES..=MAX_DROP_CAP_LINES).contains(&lines) {
            return Err(Error::InvalidFormat(format!(
                "Drop Cap line count must be between {MIN_DROP_CAP_LINES} and {MAX_DROP_CAP_LINES}"
            )));
        }
        Ok(Self(lines))
    }

    pub const fn get(self) -> u8 {
        self.0
    }
}

impl Default for DropCapLineCount {
    fn default() -> Self {
        Self::DEFAULT
    }
}

/// Number of leading characters enlarged by a Drop Cap.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct DropCapCharacterCount(NonZeroU32);

impl DropCapCharacterCount {
    pub const DEFAULT: Self = Self(NonZeroU32::MIN);

    pub fn new(characters: u32) -> Result<Self> {
        NonZeroU32::new(characters).map(Self).ok_or_else(|| {
            Error::InvalidFormat("Drop Cap character count must be nonzero".to_owned())
        })
    }

    pub const fn get(self) -> u32 {
        self.0.get()
    }
}

impl Default for DropCapCharacterCount {
    fn default() -> Self {
        Self::DEFAULT
    }
}

/// Number of lines by which a Drop Cap is raised above its paragraph.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash, Default)]
pub struct DropCapRaisedLines(u8);

impl DropCapRaisedLines {
    pub const NONE: Self = Self(0);

    pub fn new(lines: u8) -> Result<Self> {
        if lines > MAX_DROP_CAP_LINES {
            return Err(Error::InvalidFormat(format!(
                "Drop Cap raised-line count must not exceed {MAX_DROP_CAP_LINES}"
            )));
        }
        Ok(Self(lines))
    }

    pub const fn get(self) -> u8 {
        self.0
    }
}

/// How body text wraps beside a Drop Cap.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum DropCapWrap {
    #[default]
    Rectangular,
    Contour,
    None,
}

impl DropCapWrap {
    pub(crate) const fn native_value(self) -> i32 {
        match self {
            Self::Rectangular => 0,
            Self::Contour => 1,
            Self::None => 2,
        }
    }

    pub(crate) fn from_native_value(value: i32) -> Result<Self> {
        match value {
            0 => Ok(Self::Rectangular),
            1 => Ok(Self::Contour),
            2 => Ok(Self::None),
            _ => Err(Error::InvalidFormat(format!(
                "unsupported native iWork Drop Cap wrap type {value}"
            ))),
        }
    }
}

/// Extra typographic spacing beside a Drop Cap, in points.
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd, Default)]
pub struct DropCapPadding(f64);

impl DropCapPadding {
    pub const ZERO: Self = Self(0.0);

    pub fn from_points(points: f64) -> Result<Self> {
        finite_nonnegative(points, "Drop Cap padding").map(Self)
    }

    pub const fn points(self) -> f64 {
        self.0
    }
}

/// Fraction of the Drop Cap positioned outside the paragraph margin.
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd, Default)]
pub struct DropCapOutdent(f64);

impl DropCapOutdent {
    pub const NONE: Self = Self(0.0);

    pub fn from_ratio(ratio: f64) -> Result<Self> {
        finite_unit_interval(ratio, "Drop Cap outdent").map(Self)
    }

    pub const fn ratio(self) -> f64 {
        self.0
    }
}

/// Fractional corner radius retained by the native text Drop Cap model.
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
pub struct DropCapCornerRadius(f64);

impl DropCapCornerRadius {
    pub const APP_DEFAULT: Self = Self(0.2);

    pub fn from_ratio(ratio: f64) -> Result<Self> {
        finite_unit_interval(ratio, "Drop Cap corner radius").map(Self)
    }

    pub const fn ratio(self) -> f64 {
        self.0
    }
}

impl Default for DropCapCornerRadius {
    fn default() -> Self {
        Self::APP_DEFAULT
    }
}

/// Scale applied to the Drop Cap glyphs by iWork.
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
pub struct DropCapCharacterScale(f64);

impl DropCapCharacterScale {
    pub const APP_DEFAULT: Self = Self(0.8);

    pub fn from_ratio(ratio: f64) -> Result<Self> {
        if !ratio.is_finite() || ratio <= 0.0 {
            return Err(Error::InvalidFormat(
                "Drop Cap character scale must be finite and positive".to_owned(),
            ));
        }
        Ok(Self(ratio))
    }

    pub const fn ratio(self) -> f64 {
        self.0
    }
}

impl Default for DropCapCharacterScale {
    fn default() -> Self {
        Self::APP_DEFAULT
    }
}

/// A plain-text Drop Cap model supported by Pages, Numbers, and Keynote.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct ParagraphDropCap {
    pub lines: DropCapLineCount,
    pub characters: DropCapCharacterCount,
    pub raised_lines: DropCapRaisedLines,
    pub wrap: DropCapWrap,
    pub padding: DropCapPadding,
    pub outdent: DropCapOutdent,
    pub corner_radius: DropCapCornerRadius,
    pub character_scale: DropCapCharacterScale,
}

impl ParagraphDropCap {
    pub const fn new(lines: DropCapLineCount, characters: DropCapCharacterCount) -> Self {
        Self {
            lines,
            characters,
            raised_lines: DropCapRaisedLines::NONE,
            wrap: DropCapWrap::Rectangular,
            padding: DropCapPadding::ZERO,
            outdent: DropCapOutdent::NONE,
            corner_radius: DropCapCornerRadius::APP_DEFAULT,
            character_scale: DropCapCharacterScale::APP_DEFAULT,
        }
    }

    pub const fn with_raised_lines(mut self, raised_lines: DropCapRaisedLines) -> Self {
        self.raised_lines = raised_lines;
        self
    }

    pub const fn with_wrap(mut self, wrap: DropCapWrap) -> Self {
        self.wrap = wrap;
        self
    }

    pub const fn with_padding(mut self, padding: DropCapPadding) -> Self {
        self.padding = padding;
        self
    }

    pub const fn with_outdent(mut self, outdent: DropCapOutdent) -> Self {
        self.outdent = outdent;
        self
    }

    pub const fn with_corner_radius(mut self, corner_radius: DropCapCornerRadius) -> Self {
        self.corner_radius = corner_radius;
        self
    }

    pub const fn with_character_scale(mut self, character_scale: DropCapCharacterScale) -> Self {
        self.character_scale = character_scale;
        self
    }
}

impl Default for ParagraphDropCap {
    fn default() -> Self {
        Self::new(DropCapLineCount::DEFAULT, DropCapCharacterCount::DEFAULT)
    }
}

/// One Drop Cap attached to a paragraph start in a text storage.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct ParagraphDropCapPlacement {
    pub paragraph_start: ParagraphStart,
    pub drop_cap: ParagraphDropCap,
}

fn finite_nonnegative(value: f64, label: &str) -> Result<f64> {
    if !value.is_finite() || value < 0.0 {
        return Err(Error::InvalidFormat(format!(
            "{label} must be finite and nonnegative"
        )));
    }
    Ok(if value == 0.0 { 0.0 } else { value })
}

fn finite_unit_interval(value: f64, label: &str) -> Result<f64> {
    if !value.is_finite() || !(0.0..=1.0).contains(&value) {
        return Err(Error::InvalidFormat(format!(
            "{label} must be finite and between zero and one"
        )));
    }
    Ok(if value == 0.0 { 0.0 } else { value })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn drop_cap_scalars_are_strict_and_canonical() {
        assert!(DropCapLineCount::new(1).is_err());
        assert!(DropCapLineCount::new(2).is_ok());
        assert!(DropCapLineCount::new(50).is_ok());
        assert!(DropCapLineCount::new(51).is_err());
        assert!(DropCapCharacterCount::new(0).is_err());
        assert!(DropCapRaisedLines::new(51).is_err());
        assert!(DropCapPadding::from_points(f64::NAN).is_err());
        assert!(DropCapPadding::from_points(-1.0).is_err());
        assert!(DropCapOutdent::from_ratio(1.01).is_err());
        assert!(DropCapCornerRadius::from_ratio(-0.01).is_err());
        assert!(DropCapCharacterScale::from_ratio(0.0).is_err());
        assert_eq!(
            DropCapOutdent::from_ratio(-0.0).unwrap().ratio().to_bits(),
            0.0_f64.to_bits()
        );
        if usize::BITS > u32::BITS {
            assert!(ParagraphStart::from_utf16_index(usize::MAX).is_err());
        }
        assert!(DropCapWrap::from_native_value(3).is_err());
    }
}
