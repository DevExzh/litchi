//! Text and Paragraph Styling Information
//!
//! iWork documents support rich text with character-level and paragraph-level styling.

pub use litchi_iwa_text::character::{
    Error as CharacterError, TextBaselineShift, TextCapitalization, TextCharacterSpacing,
    TextDecorations, TextLigatures, TextPointSize, TextScript, TextStrikethrough, TextStyle,
    TextUnderline,
};

use super::paragraph_direction::ParagraphWritingDirection;
use super::paragraph_flow::ParagraphFlow;
use super::paragraph_tabs::{
    ParagraphDecimalTabCharacter, ParagraphDefaultTabInterval, ParagraphTabStops,
};
use crate::shapes::{
    Appearance, BlurRadius, Cap, Drop, Join, Offset, Pattern, RgbaColor, Shadow, Stroke, Width,
};
use litchi_iwa_common::shape::shadow::{Angle, Opacity};

/// Effective outline applied to uniformly styled text.
///
/// Current iWork applications store text outlines as the same typed TSD
/// strokes used by drawing shapes. [`TextOutline::standard`] reproduces the
/// one-point outline written by the Outline checkbox in Pages, Numbers, and
/// Keynote.
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub enum TextOutline {
    /// Render text without an outline.
    #[default]
    None,
    /// Render text using a standard native stroke.
    Stroke(Stroke),
}

impl TextOutline {
    /// Construct the exact outline written by current iWork applications.
    pub fn standard() -> Self {
        Self::Stroke(Stroke::new(
            RgbaColor::transparent_black(),
            Width::ONE,
            Pattern::Solid,
        ))
    }
}

/// Effective drop shadow applied to uniformly styled text.
///
/// Text shadows use iWork's native drawing-shadow archive, but the text
/// inspector supports only drop shadows. [`TextShadow::standard`] reproduces
/// the checkbox-authored value in Pages, Numbers, and Keynote.
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub enum TextShadow {
    /// Render text without a shadow.
    #[default]
    None,
    /// Render text with a typed native drop shadow.
    Drop(Drop),
}

impl TextShadow {
    /// Construct the exact shadow written by current iWork applications.
    pub const fn standard() -> Self {
        Self::Drop(Drop::new(
            Appearance::new(
                RgbaColor::black(),
                BlurRadius::ONE_POINT,
                Offset::FIVE_POINTS,
                Opacity::OPAQUE,
            ),
            Angle::FORTY_FIVE_DEGREES,
        ))
    }

    pub(crate) const fn into_shape_shadow(self) -> Shadow {
        match self {
            Self::None => Shadow::Disabled,
            Self::Drop(shadow) => Shadow::Drop(shadow),
        }
    }

    pub(crate) fn from_shape_shadow(shadow: Shadow) -> crate::Result<Self> {
        match shadow {
            Shadow::Disabled => Ok(Self::None),
            Shadow::Drop(shadow) => Ok(Self::Drop(shadow)),
            Shadow::Contact(_) | Shadow::Curved(_) => Err(crate::Error::InvalidFormat(
                "native iWork text uses a non-drop shadow".to_owned(),
            )),
        }
    }
}

/// Effective solid background painted behind uniformly styled text.
///
/// Pages, Numbers, and Keynote store this independently from both the text
/// foreground fill and the enclosing text-box fill.
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub enum TextBackground {
    /// Do not paint a background behind the text glyph run.
    #[default]
    None,
    /// Paint a solid native color behind the text glyph run.
    Color(RgbaColor),
}

/// Effective solid fill painted across a paragraph's layout box.
///
/// This is the “Paragraph Background” control in the iWork Text → Layout
/// inspector. It is independent from character-level [`TextBackground`].
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub enum ParagraphBackground {
    /// Leave the paragraph layout box unfilled.
    #[default]
    None,
    /// Paint the paragraph layout box with a solid native color.
    Color(RgbaColor),
}

const PARAGRAPH_BORDER_TOP_BIT: u8 = 1 << 0;
const PARAGRAPH_BORDER_BOTTOM_BIT: u8 = 1 << 1;
const PARAGRAPH_BORDER_LEFT_BIT: u8 = 1 << 2;
const PARAGRAPH_BORDER_RIGHT_BIT: u8 = 1 << 3;
const PARAGRAPH_BORDER_ALL_BITS: u8 = PARAGRAPH_BORDER_TOP_BIT
    | PARAGRAPH_BORDER_BOTTOM_BIT
    | PARAGRAPH_BORDER_LEFT_BIT
    | PARAGRAPH_BORDER_RIGHT_BIT;
const DEFAULT_PARAGRAPH_BORDER_OFFSET_POINTS: f32 = 6.0;

/// Selected edges of a paragraph layout box.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub struct ParagraphBorderSides(u8);

impl ParagraphBorderSides {
    pub const NONE: Self = Self(0);
    pub const TOP: Self = Self(PARAGRAPH_BORDER_TOP_BIT);
    pub const BOTTOM: Self = Self(PARAGRAPH_BORDER_BOTTOM_BIT);
    pub const LEFT: Self = Self(PARAGRAPH_BORDER_LEFT_BIT);
    pub const RIGHT: Self = Self(PARAGRAPH_BORDER_RIGHT_BIT);
    pub const ALL: Self = Self(PARAGRAPH_BORDER_ALL_BITS);

    /// Construct any combination of paragraph-border edges.
    pub const fn new(top: bool, bottom: bool, left: bool, right: bool) -> Self {
        Self(
            if top { PARAGRAPH_BORDER_TOP_BIT } else { 0 }
                | if bottom {
                    PARAGRAPH_BORDER_BOTTOM_BIT
                } else {
                    0
                }
                | if left { PARAGRAPH_BORDER_LEFT_BIT } else { 0 }
                | if right { PARAGRAPH_BORDER_RIGHT_BIT } else { 0 },
        )
    }

    pub const fn contains(self, side: Self) -> bool {
        self.0 & side.0 == side.0
    }

    pub const fn is_empty(self) -> bool {
        self.0 == 0
    }

    pub const fn is_all(self) -> bool {
        self.0 == PARAGRAPH_BORDER_ALL_BITS
    }

    pub(crate) const fn native_bits(self) -> i32 {
        self.0 as i32
    }

    pub(crate) fn from_native_bits(bits: i32) -> crate::Result<Self> {
        let bits = u8::try_from(bits).map_err(|_| {
            crate::Error::InvalidFormat("native paragraph-border sides are negative".to_owned())
        })?;
        if bits & !PARAGRAPH_BORDER_ALL_BITS != 0 {
            return Err(crate::Error::InvalidFormat(
                "native paragraph-border sides contain unknown bits".to_owned(),
            ));
        }
        Ok(Self(bits))
    }
}

impl std::ops::BitOr for ParagraphBorderSides {
    type Output = Self;

    fn bitor(self, rhs: Self) -> Self::Output {
        Self(self.0 | rhs.0)
    }
}

/// Inspector-visible gap between paragraph text and its border, in points.
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
pub struct ParagraphBorderOffset(f32);

impl ParagraphBorderOffset {
    /// The six-point offset used when an iWork paragraph border is enabled.
    pub const DEFAULT: Self = Self(DEFAULT_PARAGRAPH_BORDER_OFFSET_POINTS);

    /// Construct a finite, nonnegative paragraph-border offset.
    pub fn from_points(points: f32) -> crate::Result<Self> {
        if !points.is_finite() || points < 0.0 {
            return Err(crate::Error::InvalidFormat(
                "paragraph-border offset must be finite and nonnegative".to_owned(),
            ));
        }
        Ok(Self(points))
    }

    pub const fn points(self) -> f32 {
        self.0
    }

    pub(crate) const fn native_inset(self) -> f32 {
        self.0 - DEFAULT_PARAGRAPH_BORDER_OFFSET_POINTS
    }

    pub(crate) fn from_native_inset(inset: f32) -> crate::Result<Self> {
        Self::from_points(inset + DEFAULT_PARAGRAPH_BORDER_OFFSET_POINTS)
    }
}

impl Default for ParagraphBorderOffset {
    fn default() -> Self {
        Self::DEFAULT
    }
}

/// One uniform native paragraph-border appearance.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct ParagraphBorder {
    color: RgbaColor,
    width: Width,
    pattern: Pattern,
    sides: ParagraphBorderSides,
    offset: ParagraphBorderOffset,
    rounded_corners: bool,
}

impl ParagraphBorder {
    /// Construct a paragraph border accepted by the native iWork inspector.
    pub fn new(
        color: RgbaColor,
        width: Width,
        pattern: Pattern,
        sides: ParagraphBorderSides,
        offset: ParagraphBorderOffset,
        rounded_corners: bool,
    ) -> crate::Result<Self> {
        if sides.is_empty() {
            return Err(crate::Error::InvalidFormat(
                "a visible paragraph border must contain at least one side".to_owned(),
            ));
        }
        if rounded_corners && !sides.is_all() {
            return Err(crate::Error::InvalidFormat(
                "rounded paragraph-border corners require all four sides".to_owned(),
            ));
        }
        Ok(Self {
            color,
            width,
            pattern,
            sides,
            offset,
            rounded_corners,
        })
    }

    pub const fn color(self) -> RgbaColor {
        self.color
    }

    pub const fn width(self) -> Width {
        self.width
    }

    pub const fn pattern(self) -> Pattern {
        self.pattern
    }

    pub const fn sides(self) -> ParagraphBorderSides {
        self.sides
    }

    pub const fn offset(self) -> ParagraphBorderOffset {
        self.offset
    }

    pub const fn has_rounded_corners(self) -> bool {
        self.rounded_corners
    }

    pub(crate) fn native_stroke(self) -> Stroke {
        Stroke::new(self.color, self.width, self.pattern)
            .with_cap(Cap::Round)
            .with_join(Join::Round)
    }
}

/// Effective Text → Layout paragraph-border setting.
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub enum ParagraphBorders {
    #[default]
    None,
    Bordered(ParagraphBorder),
}

impl ParagraphBorders {
    pub(crate) const fn native_override_count(self) -> u32 {
        4
    }
}

/// Uniform paragraph properties currently supported by the shared text editor.
#[derive(Debug, Clone, PartialEq, Default)]
pub struct ParagraphStyle {
    /// Solid fill painted across the paragraph's layout box.
    pub background: ParagraphBackground,
    /// Lines painted around the paragraph's layout box.
    pub borders: ParagraphBorders,
    /// Pagination, break, widow/orphan, and hyphenation behavior.
    pub flow: ParagraphFlow,
    /// Base direction used to lay out bidirectional paragraph text.
    pub writing_direction: ParagraphWritingDirection,
    /// Native paragraph alignment.
    pub alignment: TextAlignment,
    /// Native line-spacing mode and amount.
    pub line_spacing: ParagraphLineSpacing,
    /// Space inserted before and after the paragraph.
    pub spacing: ParagraphSpacing,
    /// First-line, left, and right paragraph indentation.
    pub indents: ParagraphIndents,
    /// Character used to align decimal tab stops.
    pub decimal_tab_character: ParagraphDecimalTabCharacter,
    /// Distance between implicit paragraph tab stops.
    pub default_tab_interval: ParagraphDefaultTabInterval,
    /// Explicit ruler tab stops inherited by the paragraph.
    pub tab_stops: ParagraphTabStops,
}

impl ParagraphStyle {
    /// Create a new default paragraph style
    pub fn new() -> Self {
        Self {
            background: ParagraphBackground::None,
            borders: ParagraphBorders::None,
            flow: ParagraphFlow::default(),
            writing_direction: ParagraphWritingDirection::Natural,
            alignment: TextAlignment::Natural,
            line_spacing: ParagraphLineSpacing::default(),
            spacing: ParagraphSpacing::default(),
            indents: ParagraphIndents::default(),
            decimal_tab_character: ParagraphDecimalTabCharacter::default(),
            default_tab_interval: ParagraphDefaultTabInterval::default(),
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
    fn test_paragraph_style() {
        let para = ParagraphStyle::new();
        assert_eq!(para.background, ParagraphBackground::None);
        assert_eq!(para.borders, ParagraphBorders::None);
        assert_eq!(para.flow, ParagraphFlow::default());
        assert_eq!(para.writing_direction, ParagraphWritingDirection::Natural);
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
    fn paragraph_border_types_enforce_native_invariants() {
        assert!(ParagraphBorderOffset::from_points(-0.1).is_err());
        assert!(ParagraphBorderOffset::from_points(f32::NAN).is_err());
        assert_eq!(
            ParagraphBorderOffset::from_points(9.0).unwrap().points(),
            9.0
        );

        let top_left = ParagraphBorderSides::TOP | ParagraphBorderSides::LEFT;
        assert!(top_left.contains(ParagraphBorderSides::TOP));
        assert!(top_left.contains(ParagraphBorderSides::LEFT));
        assert!(!top_left.contains(ParagraphBorderSides::RIGHT));

        let width = Width::new(3.0).unwrap();
        assert!(
            ParagraphBorder::new(
                RgbaColor::black(),
                width,
                Pattern::Solid,
                ParagraphBorderSides::NONE,
                ParagraphBorderOffset::DEFAULT,
                false,
            )
            .is_err()
        );
        assert!(
            ParagraphBorder::new(
                RgbaColor::black(),
                width,
                Pattern::Solid,
                ParagraphBorderSides::TOP,
                ParagraphBorderOffset::DEFAULT,
                true,
            )
            .is_err()
        );
        assert!(
            ParagraphBorder::new(
                RgbaColor::black(),
                width,
                Pattern::Solid,
                ParagraphBorderSides::ALL,
                ParagraphBorderOffset::from_points(9.0).unwrap(),
                true,
            )
            .is_ok()
        );
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
