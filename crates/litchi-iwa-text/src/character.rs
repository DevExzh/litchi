//! Dependency-free character-style vocabulary for shared iWork text values.
//!
//! Native protobuf fields, archive traversal, and inheritance remain in the
//! concrete IWA adapter. This module contains only the validated semantic
//! values exchanged by Pages, Numbers, and Keynote.

/// Validation failures produced while constructing character-style values.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Error {
    /// The point size is not finite.
    PointSizeNonFinite,
    /// The point size is zero or negative.
    PointSizeNonPositive,
    /// The baseline shift is not finite.
    BaselineShiftNonFinite,
    /// The character spacing percentage is not finite.
    CharacterSpacingNonFinite,
    /// The character spacing percentage is outside its supported range.
    CharacterSpacingOutOfRange,
}

impl std::fmt::Display for Error {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        let message = match self {
            Self::PointSizeNonFinite | Self::PointSizeNonPositive => {
                "text point size must be finite and greater than zero"
            },
            Self::BaselineShiftNonFinite => "text baseline shift must be finite",
            Self::CharacterSpacingNonFinite | Self::CharacterSpacingOutOfRange => {
                "text character spacing must be finite and between -40% and 400%"
            },
        };
        formatter.write_str(message)
    }
}

impl std::error::Error for Error {}

/// Result type for character-style construction.
pub type Result<T> = std::result::Result<T, Error>;

/// Positive character size in typographic points.
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
pub struct TextPointSize(f32);

impl TextPointSize {
    /// Default size used by scratch iWork text styles.
    pub const TWELVE: Self = Self(12.0);

    /// Construct a finite character size greater than zero.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PointSizeNonFinite`] for NaN or infinite input and
    /// [`Error::PointSizeNonPositive`] for zero or negative input.
    pub fn from_points(points: f32) -> Result<Self> {
        if !points.is_finite() {
            return Err(Error::PointSizeNonFinite);
        }
        if points <= 0.0 {
            return Err(Error::PointSizeNonPositive);
        }
        Ok(Self(points))
    }

    /// Return the character size in typographic points.
    #[must_use]
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
    #[must_use]
    pub const fn new(point_size: TextPointSize) -> Self {
        Self {
            point_size,
            bold: false,
            italic: false,
        }
    }

    /// Enable or disable bold emphasis.
    #[must_use]
    pub const fn with_bold(mut self, bold: bool) -> Self {
        self.bold = bold;
        self
    }

    /// Enable or disable italic emphasis.
    #[must_use]
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

/// Strikethrough treatment stored by native iWork character styles.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum TextStrikethrough {
    #[default]
    None,
    Single,
    Double,
    Triple,
}

/// Effective uniform underline and strikethrough formatting.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub struct TextDecorations {
    /// Underline treatment.
    pub underline: TextUnderline,
    /// Strikethrough treatment.
    pub strikethrough: TextStrikethrough,
}

impl TextDecorations {
    /// No underline or strikethrough.
    pub const NONE: Self = Self {
        underline: TextUnderline::None,
        strikethrough: TextStrikethrough::None,
    };

    /// Construct a combined decoration value.
    #[must_use]
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
    ///
    /// # Errors
    ///
    /// Returns [`Error::BaselineShiftNonFinite`] for NaN or infinite input.
    pub fn from_points(points: f32) -> Result<Self> {
        if !points.is_finite() {
            return Err(Error::BaselineShiftNonFinite);
        }
        Ok(Self(points))
    }

    /// Return the signed displacement in typographic points.
    #[must_use]
    pub const fn points(self) -> f32 {
        self.0
    }
}

impl Default for TextBaselineShift {
    fn default() -> Self {
        Self::ZERO
    }
}

/// Uniform extra spacing between text characters, expressed as a percentage.
///
/// Pages, Numbers, and Keynote expose this native tracking value as
/// “Character Spacing” and constrain it to -40% through 400%.
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
pub struct TextCharacterSpacing(f32);

impl TextCharacterSpacing {
    /// Smallest percentage accepted by the iWork applications.
    pub const MINIMUM_PERCENT: f32 = -40.0;
    /// Largest percentage accepted by the iWork applications.
    pub const MAXIMUM_PERCENT: f32 = 400.0;
    const PERCENT_SCALE: f32 = 100.0;

    /// No extra spacing between characters.
    pub const NORMAL: Self = Self(0.0);
    /// Tightest character spacing accepted by the iWork applications.
    pub const MINIMUM: Self = Self(Self::MINIMUM_PERCENT / Self::PERCENT_SCALE);
    /// Widest character spacing accepted by the iWork applications.
    pub const MAXIMUM: Self = Self(Self::MAXIMUM_PERCENT / Self::PERCENT_SCALE);

    /// Construct character spacing from a percentage in the inclusive
    /// `-40.0..=400.0` range.
    ///
    /// # Errors
    ///
    /// Returns [`Error::CharacterSpacingNonFinite`] for NaN or infinite input
    /// and [`Error::CharacterSpacingOutOfRange`] for a value outside the
    /// inclusive supported range.
    pub fn from_percent(percent: f32) -> Result<Self> {
        if !percent.is_finite() {
            return Err(Error::CharacterSpacingNonFinite);
        }
        if !(Self::MINIMUM_PERCENT..=Self::MAXIMUM_PERCENT).contains(&percent) {
            return Err(Error::CharacterSpacingOutOfRange);
        }
        Ok(Self(percent / Self::PERCENT_SCALE))
    }

    /// Return the character spacing percentage.
    #[must_use]
    pub const fn percent(self) -> f32 {
        self.0 * Self::PERCENT_SCALE
    }
}

impl Default for TextCharacterSpacing {
    fn default() -> Self {
        Self::NORMAL
    }
}

/// Uniform ligature policy applied by a native iWork character style.
///
/// The names describe iWork's native behavior. The applications present
/// these as “Use None”, “Use Default”, and “Use All”, respectively. Even the
/// first policy retains ligatures required by the writing system.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum TextLigatures {
    /// Keep only ligatures required by the writing system (“Use None”).
    RequiredOnly,
    /// Use the font's standard ligatures (“Use Default”).
    #[default]
    Standard,
    /// Use every ligature supported by the font (“Use All”).
    All,
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn text_style_scalars_are_strict() -> Result<()> {
        assert_eq!(
            TextPointSize::from_points(0.0),
            Err(Error::PointSizeNonPositive)
        );
        assert_eq!(
            TextPointSize::from_points(-1.0),
            Err(Error::PointSizeNonPositive)
        );
        assert_eq!(
            TextPointSize::from_points(f32::NAN),
            Err(Error::PointSizeNonFinite)
        );

        let point_size = TextPointSize::from_points(19.5)?;
        assert_eq!(
            TextStyle::new(point_size).with_bold(true).with_italic(true),
            TextStyle {
                point_size,
                bold: true,
                italic: true,
            }
        );
        Ok(())
    }

    #[test]
    fn character_style_defaults_are_native_defaults() {
        assert_eq!(TextPointSize::default(), TextPointSize::TWELVE);
        assert_eq!(TextStyle::default(), TextStyle::new(TextPointSize::TWELVE));
        assert_eq!(TextDecorations::default(), TextDecorations::NONE);
        assert_eq!(TextBaselineShift::default(), TextBaselineShift::ZERO);
        assert_eq!(
            TextCharacterSpacing::default(),
            TextCharacterSpacing::NORMAL
        );
        assert_eq!(TextCapitalization::default(), TextCapitalization::None);
        assert_eq!(TextScript::default(), TextScript::Normal);
        assert_eq!(TextLigatures::default(), TextLigatures::Standard);
    }

    #[test]
    fn baseline_and_character_spacing_values_are_strict() {
        assert_eq!(
            TextBaselineShift::from_points(f32::NAN),
            Err(Error::BaselineShiftNonFinite)
        );
        assert_eq!(
            TextCharacterSpacing::from_percent(f32::NAN),
            Err(Error::CharacterSpacingNonFinite)
        );
        assert_eq!(
            TextCharacterSpacing::from_percent(-40.01),
            Err(Error::CharacterSpacingOutOfRange)
        );
        assert_eq!(
            TextCharacterSpacing::from_percent(400.01),
            Err(Error::CharacterSpacingOutOfRange)
        );
    }
}
