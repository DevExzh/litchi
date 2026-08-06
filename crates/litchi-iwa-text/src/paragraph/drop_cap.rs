//! Archive-free paragraph drop-cap values shared by iWork formats.

use std::num::NonZeroU32;

use crate::position::TextPosition;

const MIN_LINES: u8 = 2;
const MAX_LINES: u8 = 50;

/// Validation failures produced while constructing drop-cap values.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Error {
    /// The number of occupied lines is outside the native inspector domain.
    LinesOutOfRange,
    /// A drop cap must enlarge at least one character.
    CharactersZero,
    /// The raised-line count is outside the native inspector domain.
    RaisedLinesOutOfRange,
    /// A padding value is not finite.
    PaddingNotFinite,
    /// A padding value is negative.
    PaddingNegative,
    /// An outdent value is not finite.
    OutdentNotFinite,
    /// An outdent value is outside the inclusive unit interval.
    OutdentOutOfRange,
    /// A corner-radius value is not finite.
    CornerRadiusNotFinite,
    /// A corner-radius value is outside the inclusive unit interval.
    CornerRadiusOutOfRange,
    /// A character-scale value is not finite.
    CharacterScaleNotFinite,
    /// A character-scale value is zero or negative.
    CharacterScaleNonPositive,
}

impl std::fmt::Display for Error {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        let message = match self {
            Self::LinesOutOfRange => "drop-cap line count must be between 2 and 50",
            Self::CharactersZero => "drop-cap character count must be nonzero",
            Self::RaisedLinesOutOfRange => "drop-cap raised-line count must not exceed 50",
            Self::PaddingNotFinite | Self::PaddingNegative => {
                "drop-cap padding must be finite and nonnegative"
            },
            Self::OutdentNotFinite | Self::OutdentOutOfRange => {
                "drop-cap outdent must be finite and between zero and one"
            },
            Self::CornerRadiusNotFinite | Self::CornerRadiusOutOfRange => {
                "drop-cap corner radius must be finite and between zero and one"
            },
            Self::CharacterScaleNotFinite | Self::CharacterScaleNonPositive => {
                "drop-cap character scale must be finite and positive"
            },
        };
        formatter.write_str(message)
    }
}

impl std::error::Error for Error {}

/// Result type for drop-cap semantic values.
pub type Result<T> = std::result::Result<T, Error>;

/// Validated number of text lines occupied by a drop cap.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
#[repr(transparent)]
pub struct LineCount(u8);

impl LineCount {
    /// The native default line count.
    pub const DEFAULT: Self = Self(3);

    /// Construct a line count in the app-supported range, 2 through 50.
    ///
    /// # Errors
    ///
    /// Returns [`Error::LinesOutOfRange`] when `lines` is outside `2..=50`.
    pub fn new(lines: u8) -> Result<Self> {
        if !(MIN_LINES..=MAX_LINES).contains(&lines) {
            return Err(Error::LinesOutOfRange);
        }
        Ok(Self(lines))
    }

    /// Return the native line count.
    #[must_use]
    pub const fn get(self) -> u8 {
        self.0
    }
}

impl Default for LineCount {
    fn default() -> Self {
        Self::DEFAULT
    }
}

/// Validated number of leading characters enlarged by a drop cap.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
#[repr(transparent)]
pub struct CharacterCount(NonZeroU32);

impl CharacterCount {
    /// The native default character count.
    pub const DEFAULT: Self = Self(NonZeroU32::MIN);

    /// Construct a nonzero character count.
    ///
    /// # Errors
    ///
    /// Returns [`Error::CharactersZero`] when `characters` is zero.
    pub fn new(characters: u32) -> Result<Self> {
        NonZeroU32::new(characters)
            .map(Self)
            .ok_or(Error::CharactersZero)
    }

    /// Return the native character count.
    #[must_use]
    pub const fn get(self) -> u32 {
        self.0.get()
    }
}

impl Default for CharacterCount {
    fn default() -> Self {
        Self::DEFAULT
    }
}

/// Validated number of lines by which a drop cap is raised.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash, Default)]
#[repr(transparent)]
pub struct RaisedLines(u8);

impl RaisedLines {
    /// No raised lines.
    pub const NONE: Self = Self(0);

    /// Construct a raised-line count in the native domain.
    ///
    /// # Errors
    ///
    /// Returns [`Error::RaisedLinesOutOfRange`] when `lines` exceeds 50.
    pub fn new(lines: u8) -> Result<Self> {
        if lines > MAX_LINES {
            return Err(Error::RaisedLinesOutOfRange);
        }
        Ok(Self(lines))
    }

    /// Return the native raised-line count.
    #[must_use]
    pub const fn get(self) -> u8 {
        self.0
    }
}

/// How body text wraps beside a drop cap.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
#[repr(u8)]
pub enum Wrap {
    /// Wrap text in a rectangular box around the drop cap.
    #[default]
    Rectangular,
    /// Follow the drop-cap contour.
    Contour,
    /// Do not wrap body text beside the drop cap.
    None,
}

/// Extra typographic spacing beside a drop cap, in points.
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd, Default)]
#[repr(transparent)]
pub struct Padding(f64);

impl Padding {
    /// No extra spacing.
    pub const ZERO: Self = Self(0.0);

    /// Construct finite, nonnegative padding.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PaddingNotFinite`] for a non-finite value or
    /// [`Error::PaddingNegative`] for a negative value.
    pub fn from_points(points: f64) -> Result<Self> {
        if !points.is_finite() {
            return Err(Error::PaddingNotFinite);
        }
        if points < 0.0 {
            return Err(Error::PaddingNegative);
        }
        Ok(Self(canonical_zero(points)))
    }

    /// Return the padding in typographic points.
    #[must_use]
    pub const fn points(self) -> f64 {
        self.0
    }
}

/// Fraction of a drop cap positioned outside the paragraph margin.
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd, Default)]
#[repr(transparent)]
pub struct Outdent(f64);

impl Outdent {
    /// No outdent.
    pub const NONE: Self = Self(0.0);

    /// Construct an outdent ratio in the inclusive unit interval.
    ///
    /// # Errors
    ///
    /// Returns a typed [`Error`] when `ratio` is non-finite or outside
    /// `0.0..=1.0`.
    pub fn from_ratio(ratio: f64) -> Result<Self> {
        unit_interval(ratio, Error::OutdentNotFinite, Error::OutdentOutOfRange).map(Self)
    }

    /// Return the outdent ratio.
    #[must_use]
    pub const fn ratio(self) -> f64 {
        self.0
    }
}

/// Fractional corner radius retained by the native text drop-cap model.
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
#[repr(transparent)]
pub struct CornerRadius(f64);

impl CornerRadius {
    /// The default corner radius used by iWork.
    pub const APP_DEFAULT: Self = Self(0.2);

    /// Construct a corner-radius ratio in the inclusive unit interval.
    ///
    /// # Errors
    ///
    /// Returns a typed [`Error`] when `ratio` is non-finite or outside
    /// `0.0..=1.0`.
    pub fn from_ratio(ratio: f64) -> Result<Self> {
        unit_interval(
            ratio,
            Error::CornerRadiusNotFinite,
            Error::CornerRadiusOutOfRange,
        )
        .map(Self)
    }

    /// Return the corner-radius ratio.
    #[must_use]
    pub const fn ratio(self) -> f64 {
        self.0
    }
}

impl Default for CornerRadius {
    fn default() -> Self {
        Self::APP_DEFAULT
    }
}

/// Scale applied to drop-cap glyphs by iWork.
#[derive(Debug, Clone, Copy, PartialOrd, PartialEq)]
#[repr(transparent)]
pub struct CharacterScale(f64);

impl CharacterScale {
    /// The default glyph scale used by iWork.
    pub const APP_DEFAULT: Self = Self(0.8);

    /// Construct a finite, positive glyph scale.
    ///
    /// # Errors
    ///
    /// Returns a typed [`Error`] when `ratio` is non-finite or non-positive.
    pub fn from_ratio(ratio: f64) -> Result<Self> {
        if !ratio.is_finite() {
            return Err(Error::CharacterScaleNotFinite);
        }
        if ratio <= 0.0 {
            return Err(Error::CharacterScaleNonPositive);
        }
        Ok(Self(ratio))
    }

    /// Return the glyph scale ratio.
    #[must_use]
    pub const fn ratio(self) -> f64 {
        self.0
    }
}

impl Default for CharacterScale {
    fn default() -> Self {
        Self::APP_DEFAULT
    }
}

/// A plain-text drop-cap model shared by Pages, Numbers, and Keynote.
#[derive(Debug, Clone, Copy, PartialEq)]
#[repr(C)]
pub struct DropCap {
    /// Extra spacing beside the glyphs.
    pub padding: Padding,
    /// Fraction of the glyphs outside the paragraph margin.
    pub outdent: Outdent,
    /// Fractional corner radius of the native glyph shape.
    pub corner_radius: CornerRadius,
    /// Scale applied to the glyphs.
    pub character_scale: CharacterScale,
    /// Number of leading characters enlarged.
    pub characters: CharacterCount,
    /// Number of text lines occupied by the glyphs.
    pub lines: LineCount,
    /// Number of lines raised above the paragraph.
    pub raised_lines: RaisedLines,
    /// Body-text wrapping policy.
    pub wrap: Wrap,
}

impl DropCap {
    /// Construct a drop cap with native defaults for optional controls.
    #[must_use]
    pub const fn new(lines: LineCount, characters: CharacterCount) -> Self {
        Self {
            padding: Padding::ZERO,
            outdent: Outdent::NONE,
            corner_radius: CornerRadius::APP_DEFAULT,
            character_scale: CharacterScale::APP_DEFAULT,
            characters,
            lines,
            raised_lines: RaisedLines::NONE,
            wrap: Wrap::Rectangular,
        }
    }

    /// Set the number of raised lines.
    #[must_use]
    pub const fn with_raised_lines(mut self, raised_lines: RaisedLines) -> Self {
        self.raised_lines = raised_lines;
        self
    }

    /// Set the body-text wrapping policy.
    #[must_use]
    pub const fn with_wrap(mut self, wrap: Wrap) -> Self {
        self.wrap = wrap;
        self
    }

    /// Set the extra typographic padding.
    #[must_use]
    pub const fn with_padding(mut self, padding: Padding) -> Self {
        self.padding = padding;
        self
    }

    /// Set the margin outdent ratio.
    #[must_use]
    pub const fn with_outdent(mut self, outdent: Outdent) -> Self {
        self.outdent = outdent;
        self
    }

    /// Set the native corner-radius ratio.
    #[must_use]
    pub const fn with_corner_radius(mut self, corner_radius: CornerRadius) -> Self {
        self.corner_radius = corner_radius;
        self
    }

    /// Set the glyph scale.
    #[must_use]
    pub const fn with_character_scale(mut self, character_scale: CharacterScale) -> Self {
        self.character_scale = character_scale;
        self
    }
}

impl Default for DropCap {
    fn default() -> Self {
        Self::new(LineCount::DEFAULT, CharacterCount::DEFAULT)
    }
}

/// One drop cap attached to a paragraph boundary.
#[derive(Debug, Clone, Copy, PartialEq)]
#[repr(C)]
pub struct Placement {
    /// UTF-16 paragraph boundary at which the drop cap begins.
    pub paragraph: TextPosition,
    /// Drop-cap semantics applied at the boundary.
    pub drop_cap: DropCap,
}

impl Placement {
    /// Construct one drop-cap placement at a validated paragraph boundary.
    #[must_use]
    pub const fn new(paragraph: TextPosition, drop_cap: DropCap) -> Self {
        Self {
            paragraph,
            drop_cap,
        }
    }
}

fn canonical_zero(value: f64) -> f64 {
    if value == 0.0 { 0.0 } else { value }
}

fn unit_interval(value: f64, non_finite: Error, out_of_range: Error) -> Result<f64> {
    if !value.is_finite() {
        return Err(non_finite);
    }
    if !(0.0..=1.0).contains(&value) {
        return Err(out_of_range);
    }
    Ok(canonical_zero(value))
}

#[cfg(test)]
mod tests {
    use std::mem::size_of;

    use super::*;

    #[test]
    fn scalars_retain_strict_domains_and_canonical_zero() {
        assert_eq!(LineCount::new(2).unwrap().get(), 2);
        assert!(matches!(LineCount::new(1), Err(Error::LinesOutOfRange)));
        assert!(matches!(LineCount::new(51), Err(Error::LinesOutOfRange)));
        assert!(matches!(CharacterCount::new(0), Err(Error::CharactersZero)));
        assert!(matches!(
            RaisedLines::new(51),
            Err(Error::RaisedLinesOutOfRange)
        ));

        assert!(matches!(
            Padding::from_points(f64::NAN),
            Err(Error::PaddingNotFinite)
        ));
        assert!(matches!(
            Padding::from_points(-1.0),
            Err(Error::PaddingNegative)
        ));
        assert_eq!(
            Padding::from_points(-0.0).unwrap().points().to_bits(),
            0.0_f64.to_bits()
        );

        assert!(matches!(
            Outdent::from_ratio(f64::INFINITY),
            Err(Error::OutdentNotFinite)
        ));
        assert!(matches!(
            Outdent::from_ratio(1.01),
            Err(Error::OutdentOutOfRange)
        ));
        assert_eq!(
            Outdent::from_ratio(-0.0).unwrap().ratio().to_bits(),
            0.0_f64.to_bits()
        );
        assert!(matches!(
            CornerRadius::from_ratio(-0.01),
            Err(Error::CornerRadiusOutOfRange)
        ));
        assert!(matches!(
            CharacterScale::from_ratio(f64::NAN),
            Err(Error::CharacterScaleNotFinite)
        ));
        assert!(matches!(
            CharacterScale::from_ratio(0.0),
            Err(Error::CharacterScaleNonPositive)
        ));
    }

    #[test]
    fn drop_cap_layout_is_fixed_and_compact() {
        assert_eq!(size_of::<LineCount>(), 1);
        assert_eq!(size_of::<RaisedLines>(), 1);
        assert_eq!(size_of::<Wrap>(), 1);
        assert_eq!(size_of::<CharacterCount>(), 4);
        assert_eq!(size_of::<DropCap>(), 40);
        assert_eq!(size_of::<Placement>(), 48);
    }

    #[test]
    fn defaults_and_builders_are_semantic_only() {
        let value = DropCap::default()
            .with_raised_lines(RaisedLines::new(2).unwrap())
            .with_wrap(Wrap::Contour)
            .with_padding(Padding::from_points(6.0).unwrap())
            .with_outdent(Outdent::from_ratio(0.25).unwrap())
            .with_corner_radius(CornerRadius::from_ratio(0.3).unwrap())
            .with_character_scale(CharacterScale::from_ratio(0.9).unwrap());
        assert_eq!(value.lines, LineCount::DEFAULT);
        assert_eq!(value.characters, CharacterCount::DEFAULT);
        assert_eq!(value.raised_lines.get(), 2);
        assert_eq!(value.wrap, Wrap::Contour);
        assert_eq!(value.padding.points(), 6.0);
        assert_eq!(value.outdent.ratio(), 0.25);
        assert_eq!(value.corner_radius.ratio(), 0.3);
        assert_eq!(value.character_scale.ratio(), 0.9);
    }
}
