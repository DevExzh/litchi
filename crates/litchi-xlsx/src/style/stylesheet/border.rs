//! Typed SpreadsheetML border definitions.

use std::fmt;
use std::str::FromStr;

pub use crate::color::{ParseRgbError, Rgb};

/// SpreadsheetML namespace and edge-name convention.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Conformance {
    Transitional,
    Strict,
}

/// A visible SpreadsheetML border line.
///
/// An absent line is represented by `Option<Side>::None`; keeping `None` out
/// of this enum prevents contradictory values such as a styled side whose
/// style is `none`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum Line {
    Thin,
    Medium,
    Dashed,
    Dotted,
    Thick,
    Double,
    Hair,
    MediumDashed,
    DashDot,
    MediumDashDot,
    DashDotDot,
    MediumDashDotDot,
    SlantDashDot,
}

impl Line {
    /// Return the SpreadsheetML token without allocating.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Thin => "thin",
            Self::Medium => "medium",
            Self::Dashed => "dashed",
            Self::Dotted => "dotted",
            Self::Thick => "thick",
            Self::Double => "double",
            Self::Hair => "hair",
            Self::MediumDashed => "mediumDashed",
            Self::DashDot => "dashDot",
            Self::MediumDashDot => "mediumDashDot",
            Self::DashDotDot => "dashDotDot",
            Self::MediumDashDotDot => "mediumDashDotDot",
            Self::SlantDashDot => "slantDashDot",
        }
    }

    pub(crate) fn from_xml(value: &str) -> Result<Option<Self>, ParseLineError> {
        if value == "none" {
            Ok(None)
        } else {
            value.parse().map(Some)
        }
    }
}

impl fmt::Display for Line {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(self.as_str())
    }
}

impl FromStr for Line {
    type Err = ParseLineError;

    fn from_str(value: &str) -> Result<Self, Self::Err> {
        match value {
            "thin" => Ok(Self::Thin),
            "medium" => Ok(Self::Medium),
            "dashed" => Ok(Self::Dashed),
            "dotted" => Ok(Self::Dotted),
            "thick" => Ok(Self::Thick),
            "double" => Ok(Self::Double),
            "hair" => Ok(Self::Hair),
            "mediumDashed" => Ok(Self::MediumDashed),
            "dashDot" => Ok(Self::DashDot),
            "mediumDashDot" => Ok(Self::MediumDashDot),
            "dashDotDot" => Ok(Self::DashDotDot),
            "mediumDashDotDot" => Ok(Self::MediumDashDotDot),
            "slantDashDot" => Ok(Self::SlantDashDot),
            _ => Err(ParseLineError),
        }
    }
}

/// Error returned when a border-line token is outside the fixed domain.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ParseLineError;

impl fmt::Display for ParseLineError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str("invalid SpreadsheetML border line")
    }
}

impl std::error::Error for ParseLineError {}

/// Checked SpreadsheetML tint in the inclusive range `-1.0..=1.0`.
#[derive(Clone, Copy, PartialEq, Eq, Hash)]
pub struct Tint(u64);

impl Tint {
    pub fn new(value: f64) -> Result<Self, TintError> {
        if !value.is_finite() || !(-1.0..=1.0).contains(&value) {
            return Err(TintError);
        }
        let canonical = if value == 0.0 { 0.0 } else { value };
        Ok(Self(canonical.to_bits()))
    }

    #[must_use]
    pub const fn get(self) -> f64 {
        f64::from_bits(self.0)
    }
}

impl fmt::Debug for Tint {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.debug_tuple("Tint").field(&self.get()).finish()
    }
}

impl TryFrom<f64> for Tint {
    type Error = TintError;

    fn try_from(value: f64) -> Result<Self, Self::Error> {
        Self::new(value)
    }
}

/// Error returned when a tint is non-finite or outside `-1.0..=1.0`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct TintError;

impl fmt::Display for TintError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str("SpreadsheetML tint must be finite and between -1 and 1")
    }
}

impl std::error::Error for TintError {}

/// Typed SpreadsheetML color with an optional checked tint.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Color {
    /// A present `<color>` element with no explicit base value.
    Default {
        tint: Option<Tint>,
    },
    Rgb {
        value: Rgb,
        tint: Option<Tint>,
    },
    Theme {
        index: u32,
        tint: Option<Tint>,
    },
    Indexed {
        index: u32,
        tint: Option<Tint>,
    },
    Auto {
        enabled: bool,
        tint: Option<Tint>,
    },
}

impl Color {
    /// Preserve a color element that relies on the consumer's default base.
    #[must_use]
    pub const fn default_base() -> Self {
        Self::Default { tint: None }
    }

    #[must_use]
    pub const fn rgb(red: u8, green: u8, blue: u8) -> Self {
        Self::Rgb {
            value: Rgb::new(red, green, blue),
            tint: None,
        }
    }

    #[must_use]
    pub const fn argb(alpha: u8, red: u8, green: u8, blue: u8) -> Self {
        Self::Rgb {
            value: Rgb::argb(alpha, red, green, blue),
            tint: None,
        }
    }

    #[must_use]
    pub const fn from_rgb(value: Rgb) -> Self {
        Self::Rgb { value, tint: None }
    }

    #[must_use]
    pub const fn theme(index: u32) -> Self {
        Self::Theme { index, tint: None }
    }

    #[must_use]
    pub const fn indexed(index: u32) -> Self {
        Self::Indexed { index, tint: None }
    }

    /// Create an automatic color with `auto="1"`.
    #[must_use]
    pub const fn auto() -> Self {
        Self::Auto {
            enabled: true,
            tint: None,
        }
    }

    /// Preserve an explicit SpreadsheetML automatic-color boolean.
    #[must_use]
    pub const fn auto_value(enabled: bool) -> Self {
        Self::Auto {
            enabled,
            tint: None,
        }
    }

    #[must_use]
    pub const fn with_tint(self, value: Tint) -> Self {
        match self {
            Self::Default { .. } => Self::Default { tint: Some(value) },
            Self::Rgb { value: rgb, .. } => Self::Rgb {
                value: rgb,
                tint: Some(value),
            },
            Self::Theme { index, .. } => Self::Theme {
                index,
                tint: Some(value),
            },
            Self::Indexed { index, .. } => Self::Indexed {
                index,
                tint: Some(value),
            },
            Self::Auto { enabled, .. } => Self::Auto {
                enabled,
                tint: Some(value),
            },
        }
    }

    #[must_use]
    pub const fn tint(self) -> Option<Tint> {
        match self {
            Self::Default { tint }
            | Self::Rgb { tint, .. }
            | Self::Theme { tint, .. }
            | Self::Indexed { tint, .. }
            | Self::Auto { tint, .. } => tint,
        }
    }
}

/// One visible border side.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct Side {
    /// Line style.
    pub line: Line,
    /// Optional typed border color.
    pub color: Option<Color>,
}

impl Side {
    /// Create a border side without an explicit color.
    #[must_use]
    pub const fn new(line: Line) -> Self {
        Self { line, color: None }
    }

    #[must_use]
    pub const fn with_color(mut self, color: Color) -> Self {
        self.color = Some(color);
        self
    }
}

/// Direction of a diagonal cell border.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum Dir {
    Up,
    Down,
    Both,
}

impl Dir {
    pub(crate) const fn from_flags(up: bool, down: bool) -> Option<Self> {
        match (up, down) {
            (false, false) => None,
            (true, false) => Some(Self::Up),
            (false, true) => Some(Self::Down),
            (true, true) => Some(Self::Both),
        }
    }

    #[must_use]
    pub const fn is_up(self) -> bool {
        matches!(self, Self::Up | Self::Both)
    }

    #[must_use]
    pub const fn is_down(self) -> bool {
        matches!(self, Self::Down | Self::Both)
    }
}

/// A diagonal border.
///
/// Public construction requires both a side and direction. Parsing may retain
/// a partial but schema-valid source state, exposed through the optional
/// accessors and serialized back without inventing data.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct Diagonal {
    side: Option<Side>,
    dir: Option<Dir>,
}

impl Diagonal {
    #[must_use]
    pub const fn new(side: Side, dir: Dir) -> Self {
        Self {
            side: Some(side),
            dir: Some(dir),
        }
    }

    #[must_use]
    pub fn side(&self) -> Option<&Side> {
        self.side.as_ref()
    }

    #[must_use]
    pub const fn dir(&self) -> Option<Dir> {
        self.dir
    }

    #[must_use]
    pub const fn is_visible(&self) -> bool {
        self.side.is_some() && self.dir.is_some()
    }

    pub(crate) fn from_parts(side: Option<Side>, dir: Option<Dir>) -> Option<Self> {
        (side.is_some() || dir.is_some()).then_some(Self { side, dir })
    }
}

/// Border information for a cell.
#[derive(Debug, Clone, Default, PartialEq, Eq, Hash)]
pub struct Border {
    /// Strict logical leading edge.
    pub start: Option<Side>,
    /// Strict logical trailing edge.
    pub end: Option<Side>,
    /// Transitional physical left edge.
    pub left: Option<Side>,
    /// Transitional physical right edge.
    pub right: Option<Side>,
    pub top: Option<Side>,
    pub bottom: Option<Side>,
    /// Borders between columns in a multi-cell region.
    pub vertical: Option<Side>,
    /// Borders between rows in a multi-cell region.
    pub horizontal: Option<Side>,
    pub diagonal: Option<Diagonal>,
    /// Explicit `outline` value; absence retains the schema default.
    pub outline: Option<bool>,
}

impl Border {
    #[inline]
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    #[inline]
    #[must_use]
    pub fn has_borders(&self) -> bool {
        self.start.is_some()
            || self.end.is_some()
            || self.left.is_some()
            || self.right.is_some()
            || self.top.is_some()
            || self.bottom.is_some()
            || self.vertical.is_some()
            || self.horizontal.is_some()
            || self.diagonal.as_ref().is_some_and(Diagonal::is_visible)
    }

    pub(crate) fn set_diagonal_side(&mut self, side: Option<Side>) {
        let dir = self.diagonal.as_ref().and_then(Diagonal::dir);
        self.diagonal = Diagonal::from_parts(side, dir);
    }

    pub(crate) fn set_diagonal_dir(&mut self, dir: Option<Dir>) {
        let side = self
            .diagonal
            .as_mut()
            .and_then(|diagonal| diagonal.side.take());
        self.diagonal = Diagonal::from_parts(side, dir);
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn line_tokens_round_trip_without_allocating() {
        let lines = [
            Line::Thin,
            Line::Medium,
            Line::Dashed,
            Line::Dotted,
            Line::Thick,
            Line::Double,
            Line::Hair,
            Line::MediumDashed,
            Line::DashDot,
            Line::MediumDashDot,
            Line::DashDotDot,
            Line::MediumDashDotDot,
            Line::SlantDashDot,
        ];
        for line in lines {
            assert_eq!(line.as_str().parse::<Line>(), Ok(line));
            assert_eq!(line.to_string(), line.as_str());
        }
        assert!("invalid".parse::<Line>().is_err());
        assert_eq!(Line::from_xml("none"), Ok(None));
    }

    #[test]
    fn rgb_and_tint_are_checked() {
        assert_eq!(Rgb::new(0, 0xA1, 0xFF).to_string(), "FF00A1FF");
        assert_eq!(
            "8000a1FF".parse::<Rgb>(),
            Ok(Rgb::argb(0x80, 0, 0xA1, 0xFF))
        );
        assert_eq!(Rgb::argb(0x80, 0, 0xA1, 0xFF).to_string(), "8000A1FF");
        assert!("00A1FF".parse::<Rgb>().is_err());
        assert!("#00A1FF".parse::<Rgb>().is_err());
        assert!(Tint::new(-1.0).is_ok());
        assert!(Tint::new(1.0).is_ok());
        assert!(Tint::new(1.1).is_err());
        assert!(Tint::new(f64::NAN).is_err());
    }

    #[test]
    fn authored_diagonal_is_always_complete() {
        let diagonal = Diagonal::new(Side::new(Line::Hair), Dir::Both);
        assert!(diagonal.is_visible());
        assert_eq!(diagonal.dir(), Some(Dir::Both));
        assert_eq!(diagonal.side().map(|side| side.line), Some(Line::Hair));

        let partial = Diagonal::from_parts(None, Some(Dir::Up));
        assert!(partial.as_ref().is_some_and(|value| !value.is_visible()));
    }

    #[test]
    fn border_reports_only_visible_sides() {
        let mut border = Border::new();
        border.set_diagonal_dir(Some(Dir::Up));
        assert!(!border.has_borders());
        border.diagonal = Some(Diagonal::new(Side::new(Line::Thin), Dir::Up));
        assert!(border.has_borders());
    }
}
