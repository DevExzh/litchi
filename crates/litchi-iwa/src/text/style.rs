//! Text and Paragraph Styling Information
//!
//! iWork documents support rich text with character-level and paragraph-level styling.

use super::paragraph_tabs::ParagraphTabStops;

/// Positive character size in typographic points.
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
pub struct TextPointSize(f32);

impl TextPointSize {
    /// Default size used by scratch iWork text styles.
    pub const TWELVE: Self = Self(12.0);

    /// Construct a finite character size greater than zero.
    pub fn from_points(points: f32) -> crate::Result<Self> {
        if !points.is_finite() || points <= 0.0 {
            return Err(crate::Error::InvalidFormat(
                "text point size must be finite and greater than zero".to_owned(),
            ));
        }
        Ok(Self(points))
    }

    /// Return the character size in typographic points.
    pub const fn points(self) -> f32 {
        self.0
    }
}

impl Default for TextPointSize {
    fn default() -> Self {
        Self::TWELVE
    }
}

/// Effective uniform character formatting attached to a paragraph style.
///
/// Pages, Numbers, and Keynote store whole-paragraph direct font formatting
/// in the paragraph-style inheritance graph. Partially formatted ranges use a
/// separate character-style table and are intentionally not represented by
/// this uniform value.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct TextStyle {
    /// Effective font size.
    pub point_size: TextPointSize,
    /// Whether bold emphasis is enabled.
    pub bold: bool,
    /// Whether italic emphasis is enabled.
    pub italic: bool,
}

impl TextStyle {
    /// Construct character formatting with no emphasis.
    pub const fn new(point_size: TextPointSize) -> Self {
        Self {
            point_size,
            bold: false,
            italic: false,
        }
    }

    /// Enable or disable bold emphasis.
    pub const fn with_bold(mut self, bold: bool) -> Self {
        self.bold = bold;
        self
    }

    /// Enable or disable italic emphasis.
    pub const fn with_italic(mut self, italic: bool) -> Self {
        self.italic = italic;
        self
    }
}

impl Default for TextStyle {
    fn default() -> Self {
        Self::new(TextPointSize::TWELVE)
    }
}

/// Underline treatment stored by native iWork character styles.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum TextUnderline {
    #[default]
    None,
    Single,
    Double,
    Wavy,
}

impl TextUnderline {
    pub(crate) const fn native_value(self) -> i32 {
        match self {
            Self::None => 0,
            Self::Single => 1,
            Self::Double => 2,
            Self::Wavy => 3,
        }
    }

    pub(crate) fn from_native_value(value: i32) -> crate::Result<Self> {
        match value {
            0 => Ok(Self::None),
            1 => Ok(Self::Single),
            2 => Ok(Self::Double),
            3 => Ok(Self::Wavy),
            _ => Err(crate::Error::InvalidFormat(format!(
                "unsupported native iWork underline type {value}"
            ))),
        }
    }
}

/// Strikethrough treatment stored by native iWork character styles.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum TextStrikethrough {
    #[default]
    None,
    Single,
    Double,
    Triple,
}

impl TextStrikethrough {
    pub(crate) const fn native_value(self) -> i32 {
        match self {
            Self::None => 0,
            Self::Single => 1,
            Self::Double => 2,
            Self::Triple => 3,
        }
    }

    pub(crate) fn from_native_value(value: i32) -> crate::Result<Self> {
        match value {
            0 => Ok(Self::None),
            1 => Ok(Self::Single),
            2 => Ok(Self::Double),
            3 => Ok(Self::Triple),
            _ => Err(crate::Error::InvalidFormat(format!(
                "unsupported native iWork strikethrough type {value}"
            ))),
        }
    }
}

/// Effective uniform underline and strikethrough formatting.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub struct TextDecorations {
    pub underline: TextUnderline,
    pub strikethrough: TextStrikethrough,
}

impl TextDecorations {
    pub const NONE: Self = Self {
        underline: TextUnderline::None,
        strikethrough: TextStrikethrough::None,
    };

    pub const fn new(underline: TextUnderline, strikethrough: TextStrikethrough) -> Self {
        Self {
            underline,
            strikethrough,
        }
    }
}

/// Effective uniform capitalization applied by a native iWork character style.
///
/// Title Case and Start Case share iWork's native titled mode. iWork marks
/// Title Case with a separate linguistic-boundaries flag and leaves that flag
/// unset for Start Case.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum TextCapitalization {
    #[default]
    None,
    AllCaps,
    SmallCaps,
    TitleCase,
    StartCase,
}

impl TextCapitalization {
    pub(crate) const fn native_value(self) -> i32 {
        match self {
            Self::None => 0,
            Self::AllCaps => 1,
            Self::SmallCaps => 2,
            Self::TitleCase | Self::StartCase => 3,
        }
    }

    pub(crate) const fn uses_linguistics(self) -> Option<bool> {
        match self {
            Self::TitleCase => Some(true),
            _ => None,
        }
    }

    pub(crate) const fn native_override_count(self) -> u32 {
        match self {
            Self::TitleCase => 2,
            _ => 1,
        }
    }

    pub(crate) fn from_native_value(
        value: i32,
        uses_linguistics: Option<bool>,
    ) -> crate::Result<Self> {
        match (value, uses_linguistics.unwrap_or(false)) {
            (0, false) => Ok(Self::None),
            (1, false) => Ok(Self::AllCaps),
            (2, false) => Ok(Self::SmallCaps),
            (3, true) => Ok(Self::TitleCase),
            (3, false) => Ok(Self::StartCase),
            (0..=2, true) => Err(crate::Error::InvalidFormat(
                "native iWork linguistic capitalization is not title case".to_owned(),
            )),
            _ => Err(crate::Error::InvalidFormat(format!(
                "unsupported native iWork capitalization type {value}"
            ))),
        }
    }
}

/// Effective uniform baseline script applied by a native iWork character style.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum TextScript {
    /// Keep characters on the normal text baseline.
    #[default]
    Normal,
    /// Raise and resize characters as superscript.
    Superscript,
    /// Lower and resize characters as subscript.
    Subscript,
}

impl TextScript {
    pub(crate) const fn native_value(self) -> i32 {
        match self {
            Self::Normal => 0,
            Self::Superscript => 1,
            Self::Subscript => 2,
        }
    }

    pub(crate) fn from_native_value(value: i32) -> crate::Result<Self> {
        match value {
            0 => Ok(Self::Normal),
            1 => Ok(Self::Superscript),
            2 => Ok(Self::Subscript),
            _ => Err(crate::Error::InvalidFormat(format!(
                "unsupported native iWork text script {value}"
            ))),
        }
    }
}

/// Uniform vertical displacement from the text baseline in typographic points.
///
/// This is independent of [`TextScript`]: iWork can apply a custom baseline
/// shift while retaining normal, superscript, or subscript formatting.
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
pub struct TextBaselineShift(f32);

impl TextBaselineShift {
    /// No custom displacement from the inherited text baseline.
    pub const ZERO: Self = Self(0.0);

    /// Construct a finite baseline shift.
    ///
    /// Positive values raise text and negative values lower it.
    pub fn from_points(points: f32) -> crate::Result<Self> {
        if !points.is_finite() {
            return Err(crate::Error::InvalidFormat(
                "text baseline shift must be finite".to_owned(),
            ));
        }
        Ok(Self(points))
    }

    /// Return the signed displacement in typographic points.
    pub const fn points(self) -> f32 {
        self.0
    }
}

impl Default for TextBaselineShift {
    fn default() -> Self {
        Self::ZERO
    }
}

/// Uniform paragraph properties currently supported by the shared text editor.
#[derive(Debug, Clone, PartialEq, Default)]
pub struct ParagraphStyle {
    /// Native paragraph alignment.
    pub alignment: TextAlignment,
    /// Native line-spacing mode and amount.
    pub line_spacing: ParagraphLineSpacing,
    /// Space inserted before and after the paragraph.
    pub spacing: ParagraphSpacing,
    /// First-line, left, and right paragraph indentation.
    pub indents: ParagraphIndents,
    /// Explicit ruler tab stops inherited by the paragraph.
    pub tab_stops: ParagraphTabStops,
}

impl ParagraphStyle {
    /// Create a new default paragraph style
    pub fn new() -> Self {
        Self {
            alignment: TextAlignment::Natural,
            line_spacing: ParagraphLineSpacing::default(),
            spacing: ParagraphSpacing::default(),
            indents: ParagraphIndents::default(),
            tab_stops: ParagraphTabStops::default(),
        }
    }
}

/// Nonnegative paragraph indentation in typographic points.
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd, Default)]
pub struct ParagraphIndentPoints(f32);

impl ParagraphIndentPoints {
    pub const ZERO: Self = Self(0.0);

    /// Construct a finite, nonnegative paragraph indentation distance.
    pub fn from_points(points: f32) -> crate::Result<Self> {
        if !points.is_finite() || points < 0.0 {
            return Err(crate::Error::InvalidFormat(
                "paragraph indentation must be finite and nonnegative".to_owned(),
            ));
        }
        Ok(if points == 0.0 {
            Self::ZERO
        } else {
            Self(points)
        })
    }

    /// Return the distance in typographic points.
    pub const fn points(self) -> f32 {
        self.0
    }
}

/// Native first-line, left, and right paragraph indentation.
///
/// `first_line` and `left` are both absolute distances from the left text
/// boundary. A hanging indent therefore has `first_line < left`.
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub struct ParagraphIndents {
    pub first_line: ParagraphIndentPoints,
    pub left: ParagraphIndentPoints,
    pub right: ParagraphIndentPoints,
}

impl ParagraphIndents {
    pub const NONE: Self = Self {
        first_line: ParagraphIndentPoints::ZERO,
        left: ParagraphIndentPoints::ZERO,
        right: ParagraphIndentPoints::ZERO,
    };

    pub const fn new(
        first_line: ParagraphIndentPoints,
        left: ParagraphIndentPoints,
        right: ParagraphIndentPoints,
    ) -> Self {
        Self {
            first_line,
            left,
            right,
        }
    }
}

/// Nonnegative paragraph spacing in typographic points.
///
/// Pages, Numbers, and Keynote clamp negative paragraph spacing to zero.
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd, Default)]
pub struct ParagraphSpacingPoints(f32);

impl ParagraphSpacingPoints {
    pub const ZERO: Self = Self(0.0);

    /// Construct a finite, nonnegative paragraph spacing distance.
    pub fn from_points(points: f32) -> crate::Result<Self> {
        if !points.is_finite() || points < 0.0 {
            return Err(crate::Error::InvalidFormat(
                "paragraph spacing must be finite and nonnegative".to_owned(),
            ));
        }
        Ok(if points == 0.0 {
            Self::ZERO
        } else {
            Self(points)
        })
    }

    /// Return the distance in typographic points.
    pub const fn points(self) -> f32 {
        self.0
    }
}

/// Space inserted before and after a paragraph.
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub struct ParagraphSpacing {
    pub before: ParagraphSpacingPoints,
    pub after: ParagraphSpacingPoints,
}

impl ParagraphSpacing {
    pub const NONE: Self = Self {
        before: ParagraphSpacingPoints::ZERO,
        after: ParagraphSpacingPoints::ZERO,
    };

    pub const fn new(before: ParagraphSpacingPoints, after: ParagraphSpacingPoints) -> Self {
        Self { before, after }
    }
}

/// Positive multiplier used by relative paragraph line spacing.
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
pub struct ParagraphLineSpacingMultiple(f32);

impl ParagraphLineSpacingMultiple {
    pub const SINGLE: Self = Self(1.0);
    pub const ONE_POINT_TWO: Self = Self(1.2);
    pub const ONE_POINT_FIVE: Self = Self(1.5);
    pub const DOUBLE: Self = Self(2.0);

    /// Construct a finite multiplier greater than zero.
    pub fn new(value: f32) -> crate::Result<Self> {
        if !value.is_finite() || value <= 0.0 {
            return Err(crate::Error::InvalidFormat(
                "paragraph line-spacing multiplier must be finite and greater than zero".to_owned(),
            ));
        }
        Ok(Self(value))
    }

    /// Return the multiplier represented by this value.
    pub const fn get(self) -> f32 {
        self.0
    }
}

impl Default for ParagraphLineSpacingMultiple {
    fn default() -> Self {
        Self::SINGLE
    }
}

/// Positive line-spacing distance in typographic points.
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
pub struct ParagraphLineSpacingPoints(f32);

impl ParagraphLineSpacingPoints {
    /// Construct a finite point distance greater than zero.
    pub fn from_points(points: f32) -> crate::Result<Self> {
        if !points.is_finite() || points <= 0.0 {
            return Err(crate::Error::InvalidFormat(
                "paragraph line-spacing distance must be finite and greater than zero".to_owned(),
            ));
        }
        Ok(Self(points))
    }

    /// Return the distance in typographic points.
    pub const fn points(self) -> f32 {
        self.0
    }
}

/// Native iWork line-spacing modes.
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum ParagraphLineSpacing {
    /// Scale the font's natural line height by a positive multiplier.
    Relative(ParagraphLineSpacingMultiple),
    /// Use at least the specified point distance.
    AtLeast(ParagraphLineSpacingPoints),
    /// Use exactly the specified point distance.
    Exactly(ParagraphLineSpacingPoints),
    /// Cap the line distance at the specified point distance.
    Maximum(ParagraphLineSpacingPoints),
    /// Add the specified point distance between adjacent lines.
    Between(ParagraphLineSpacingPoints),
}

impl Default for ParagraphLineSpacing {
    fn default() -> Self {
        Self::Relative(ParagraphLineSpacingMultiple::SINGLE)
    }
}

/// Horizontal paragraph alignment used by native iWork paragraph styles.
///
/// `Natural` follows the paragraph writing direction. The four explicit
/// variants correspond to the controls exposed by Pages, Numbers, and Keynote.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum TextAlignment {
    #[default]
    Natural,
    Left,
    Center,
    Right,
    Justified,
}

impl TextAlignment {
    pub(crate) const fn native_value(self) -> i32 {
        match self {
            Self::Natural => 0,
            Self::Right => 1,
            Self::Center => 2,
            Self::Justified => 3,
            Self::Left => 4,
        }
    }

    pub(crate) fn from_native_value(value: i32) -> crate::Result<Self> {
        match value {
            0 => Ok(Self::Natural),
            1 => Ok(Self::Right),
            2 => Ok(Self::Center),
            3 => Ok(Self::Justified),
            4 => Ok(Self::Left),
            _ => Err(crate::Error::InvalidFormat(format!(
                "unsupported native iWork paragraph alignment {value}"
            ))),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn text_style_scalars_are_strict() {
        assert!(TextPointSize::from_points(0.0).is_err());
        assert!(TextPointSize::from_points(-1.0).is_err());
        assert!(TextPointSize::from_points(f32::NAN).is_err());
        assert_eq!(
            TextStyle::new(TextPointSize::from_points(19.5).unwrap())
                .with_bold(true)
                .with_italic(true),
            TextStyle {
                point_size: TextPointSize::from_points(19.5).unwrap(),
                bold: true,
                italic: true,
            }
        );
    }

    #[test]
    fn text_decoration_native_values_are_strict_and_reversible() {
        for underline in [
            TextUnderline::None,
            TextUnderline::Single,
            TextUnderline::Double,
            TextUnderline::Wavy,
        ] {
            assert_eq!(
                TextUnderline::from_native_value(underline.native_value()).unwrap(),
                underline
            );
        }
        for strikethrough in [
            TextStrikethrough::None,
            TextStrikethrough::Single,
            TextStrikethrough::Double,
            TextStrikethrough::Triple,
        ] {
            assert_eq!(
                TextStrikethrough::from_native_value(strikethrough.native_value()).unwrap(),
                strikethrough
            );
        }
        assert!(TextUnderline::from_native_value(4).is_err());
        assert!(TextStrikethrough::from_native_value(-1).is_err());
    }

    #[test]
    fn test_paragraph_style() {
        let para = ParagraphStyle::new();
        assert_eq!(para.alignment, TextAlignment::Natural);
        assert_eq!(
            para.line_spacing,
            ParagraphLineSpacing::Relative(ParagraphLineSpacingMultiple::SINGLE)
        );
        assert_eq!(para.spacing, ParagraphSpacing::NONE);
        assert_eq!(para.indents, ParagraphIndents::NONE);
        assert!(para.tab_stops.is_empty());
    }

    #[test]
    fn test_alignment_parsing() {
        assert_eq!(
            TextAlignment::from_native_value(0).unwrap(),
            TextAlignment::Natural
        );
        assert_eq!(
            TextAlignment::from_native_value(1).unwrap(),
            TextAlignment::Right
        );
        assert_eq!(
            TextAlignment::from_native_value(2).unwrap(),
            TextAlignment::Center
        );
        assert_eq!(
            TextAlignment::from_native_value(3).unwrap(),
            TextAlignment::Justified
        );
        assert_eq!(
            TextAlignment::from_native_value(4).unwrap(),
            TextAlignment::Left
        );
        assert!(TextAlignment::from_native_value(999).is_err());
    }

    #[test]
    fn line_spacing_scalars_are_strict() {
        assert_eq!(
            ParagraphLineSpacingMultiple::new(1.5).unwrap(),
            ParagraphLineSpacingMultiple::ONE_POINT_FIVE
        );
        assert!(ParagraphLineSpacingMultiple::new(0.0).is_err());
        assert!(ParagraphLineSpacingMultiple::new(f32::NAN).is_err());
        assert!(ParagraphLineSpacingPoints::from_points(0.0).is_err());
        assert!(ParagraphLineSpacingPoints::from_points(f32::INFINITY).is_err());
    }

    #[test]
    fn paragraph_spacing_points_are_strict_and_allow_zero() {
        assert_eq!(
            ParagraphSpacingPoints::from_points(0.0).unwrap(),
            ParagraphSpacingPoints::ZERO
        );
        assert_eq!(
            ParagraphSpacingPoints::from_points(-0.0)
                .unwrap()
                .points()
                .to_bits(),
            0.0_f32.to_bits()
        );
        assert_eq!(
            ParagraphSpacingPoints::from_points(12.5).unwrap().points(),
            12.5
        );
        assert!(ParagraphSpacingPoints::from_points(-0.1).is_err());
        assert!(ParagraphSpacingPoints::from_points(f32::NAN).is_err());
        assert!(ParagraphSpacingPoints::from_points(f32::INFINITY).is_err());
    }

    #[test]
    fn paragraph_indent_points_are_strict_and_allow_hanging_indents() {
        let hanging = ParagraphIndents::new(
            ParagraphIndentPoints::from_points(8.0).unwrap(),
            ParagraphIndentPoints::from_points(24.0).unwrap(),
            ParagraphIndentPoints::from_points(12.0).unwrap(),
        );
        assert!(hanging.first_line < hanging.left);
        assert_eq!(
            ParagraphIndentPoints::from_points(-0.0)
                .unwrap()
                .points()
                .to_bits(),
            0.0_f32.to_bits()
        );
        assert!(ParagraphIndentPoints::from_points(-0.1).is_err());
        assert!(ParagraphIndentPoints::from_points(f32::NAN).is_err());
        assert!(ParagraphIndentPoints::from_points(f32::INFINITY).is_err());
    }
}
