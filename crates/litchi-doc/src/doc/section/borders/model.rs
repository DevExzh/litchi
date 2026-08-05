//! Checked semantic values for Word section page borders.

use std::fmt;

/// Pages in a section to which its page borders apply.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ApplyTo {
    AllPages,
    FirstPage,
    AllButFirstPage,
}

/// Z-order of page borders relative to the section contents.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Depth {
    InFront,
    Behind,
}

/// Reference from which a page border's spacing is measured.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Offset {
    Text,
    PageEdge,
}

/// Validation failure for a page-border value.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Error {
    /// `Brc80.dptSpace` is a five-bit value.
    InvalidSpacing(u8),
    /// `Brc80.brcType` art codes occupy the inclusive range `0x40..=0xE3`.
    InvalidArt(u8),
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::InvalidSpacing(value) => write!(
                formatter,
                "page-border spacing {value} exceeds the 31-point Brc80 limit"
            ),
            Self::InvalidArt(value) => write!(
                formatter,
                "page-border art code {value:#04x} is outside the Brc80 art range"
            ),
        }
    }
}

impl std::error::Error for Error {}

/// A validated Word page-border art code.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Art(u8);

impl Art {
    /// Return the `BrcType` art code in the inclusive range `0x40..=0xE3`.
    pub fn code(self) -> u8 {
        self.0
    }
}

impl TryFrom<u8> for Art {
    type Error = Error;

    fn try_from(value: u8) -> Result<Self, Self::Error> {
        if (0x40..=0xE3).contains(&value) {
            Ok(Self(value))
        } else {
            Err(Error::InvalidArt(value))
        }
    }
}

/// Line or image style of a Word 97 `Brc80` page border.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Style {
    Single,
    Double,
    Thick,
    Dotted,
    Dashed,
    DotDash,
    DotDotDash,
    Triple,
    ThinThickSmallGap,
    ThickThinSmallGap,
    ThinThickThinSmallGap,
    ThinThickMediumGap,
    ThickThinMediumGap,
    ThinThickThinMediumGap,
    ThinThickLargeGap,
    ThickThinLargeGap,
    ThinThickThinLargeGap,
    Wave,
    DoubleWave,
    DashSmallGap,
    DashDotStroked,
    ThreeDEmboss,
    ThreeDEngrave,
    Art(Art),
}

/// Palette color selected by a Word `Ico` value.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Color {
    Automatic,
    Black,
    Blue,
    Cyan,
    Green,
    Magenta,
    Red,
    Yellow,
    White,
    DarkBlue,
    DarkCyan,
    DarkGreen,
    DarkMagenta,
    DarkRed,
    DarkYellow,
    DarkGray,
    LightGray,
}

/// One section page-border edge decoded from `Brc80`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Border {
    pub style: Style,
    /// Width in eighths of a point. Values below two render as two.
    pub width_eighth_points: u8,
    pub color: Color,
    /// Distance from text or the page edge, in points.
    pub spacing_points: u8,
    pub shadow: bool,
    pub frame: bool,
}

impl Border {
    /// Validate fields whose domains are wider in Rust than in `Brc80`.
    pub fn validate(self) -> Result<(), Error> {
        if self.spacing_points > 31 {
            return Err(Error::InvalidSpacing(self.spacing_points));
        }
        Ok(())
    }
}

/// Page borders and shared placement controls for one section.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Borders {
    pub top: Option<Border>,
    pub left: Option<Border>,
    pub bottom: Option<Border>,
    pub right: Option<Border>,
    pub apply_to: ApplyTo,
    pub depth: Depth,
    pub offset_from: Offset,
}

impl Default for Borders {
    fn default() -> Self {
        Self {
            top: None,
            left: None,
            bottom: None,
            right: None,
            apply_to: ApplyTo::AllPages,
            depth: Depth::InFront,
            offset_from: Offset::Text,
        }
    }
}

impl Borders {
    /// Validate every present border edge before it is stored or encoded.
    pub fn validate(self) -> Result<(), Error> {
        for border in [self.top, self.left, self.bottom, self.right]
            .into_iter()
            .flatten()
        {
            border.validate()?;
        }
        Ok(())
    }
}
