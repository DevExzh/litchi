//! Strict list presets and nesting levels shared by Pages, Numbers, and Keynote.

use super::super::drop_cap::ParagraphStart;
use super::super::style::TextPointSize;
use crate::shapes::RgbaColor;
use crate::{Error, Result};

const MAX_PARAGRAPH_LIST_LEVEL: u8 = 8;
const MAX_PARAGRAPH_LIST_BULLET_CHARACTERS: usize = 32;
const PERCENT_PER_SCALE_UNIT: f32 = 100.0;
const STANDARD_TOP_LEVEL_BULLET_BASELINE_POINTS: f32 = 0.0;
const STANDARD_NESTED_BULLET_BASELINE_POINTS: f32 = -1.0;
const STANDARD_FONT_EM_POINTS: f32 = 11.0;
const STANDARD_BULLET_INDENT_STEP_POINTS: f32 = 9.0;
const STANDARD_NUMBER_INDENT_STEP_POINTS: f32 = 18.0;

/// A canonical paragraph-list presentation understood by all three iWork apps.
///
/// The presets describe the complete nine-level native list style rather than
/// exposing unvalidated protobuf integers or partial per-level state.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum ParagraphList {
    /// Ordinary paragraphs without labels.
    #[default]
    None,
    /// Apple’s standard bullet preset using the `•` marker.
    Bullet,
    /// Apple’s standard decimal-number preset.
    Numbered,
}

/// One list preset boundary at a validated UTF-16 paragraph start.
///
/// The preset remains effective until the next placement. A complete placement
/// list always begins at [`ParagraphStart::ZERO`].
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ParagraphListPlacement {
    pub paragraph: ParagraphStart,
    pub list: ParagraphList,
}

impl ParagraphListPlacement {
    pub const fn new(paragraph: ParagraphStart, list: ParagraphList) -> Self {
        Self { paragraph, list }
    }
}

impl ParagraphList {
    pub(crate) const fn native_name(self) -> &'static str {
        match self {
            Self::None => "None",
            Self::Bullet => "Bullet",
            Self::Numbered => "Numbered",
        }
    }

    pub(crate) const fn preset_index(self) -> usize {
        match self {
            Self::None => 0,
            Self::Bullet => 1,
            Self::Numbered => 2,
        }
    }
}

/// A zero-based nesting level in iWork's nine-level paragraph-list model.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash, Default)]
pub struct ParagraphListLevel(u8);

impl ParagraphListLevel {
    /// Top-level list item.
    pub const ZERO: Self = Self(0);
    /// First nested list level.
    pub const ONE: Self = Self(1);
    /// Deepest level supported by the native nine-level style model.
    pub const MAX: Self = Self(MAX_PARAGRAPH_LIST_LEVEL);

    /// Construct a validated zero-based nesting level.
    pub fn new(level: u8) -> Result<Self> {
        if level > MAX_PARAGRAPH_LIST_LEVEL {
            return Err(Error::InvalidFormat(format!(
                "paragraph list level must not exceed {MAX_PARAGRAPH_LIST_LEVEL}"
            )));
        }
        Ok(Self(level))
    }

    /// Return the zero-based native nesting level.
    pub const fn get(self) -> u8 {
        self.0
    }

    pub(crate) fn from_native(value: u32) -> Result<Self> {
        u8::try_from(value)
            .map_err(|_| {
                Error::InvalidFormat(format!("native paragraph list level {value} exceeds u8"))
            })
            .and_then(Self::new)
    }
}

/// One effective list-level boundary at a UTF-16 paragraph start.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ParagraphListLevelPlacement {
    pub paragraph: ParagraphStart,
    pub level: ParagraphListLevel,
}

impl ParagraphListLevelPlacement {
    pub const fn new(paragraph: ParagraphStart, level: ParagraphListLevel) -> Self {
        Self { paragraph, level }
    }
}

/// A positive starting number for a restarted numbered list.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct ParagraphListStart(u32);

impl ParagraphListStart {
    /// The first positive list number.
    pub const ONE: Self = Self(1);

    /// Construct a validated positive starting number.
    pub fn new(number: u32) -> Result<Self> {
        if number == 0 {
            return Err(Error::InvalidFormat(
                "paragraph list starting number must be positive".to_owned(),
            ));
        }
        Ok(Self(number))
    }

    /// Return the native positive starting number.
    pub const fn get(self) -> u32 {
        self.0
    }

    pub(crate) fn from_native(value: u32) -> Result<Self> {
        Self::new(value)
    }
}

/// How one numbered-list paragraph participates in the current sequence.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum ParagraphListNumbering {
    /// Continue the preceding numbered sequence.
    #[default]
    Continue,
    /// Restart numbering at the supplied positive number.
    StartAt(ParagraphListStart),
}

impl ParagraphListNumbering {
    pub(crate) const fn native_start(self) -> u32 {
        match self {
            Self::Continue => 0,
            Self::StartAt(start) => start.get(),
        }
    }

    pub(crate) fn from_native(value: u32) -> Result<Self> {
        if value == 0 {
            Ok(Self::Continue)
        } else {
            ParagraphListStart::from_native(value).map(Self::StartAt)
        }
    }
}

/// A validated text marker used by one bullet-list level.
///
/// iWork accepts a short sequence of printable characters, not just a single
/// Unicode scalar. Newlines and control characters are rejected because they
/// cannot be represented by the native bullet inspector.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct ParagraphListBullet(Box<str>);

impl ParagraphListBullet {
    /// Apple's standard text bullet.
    pub const STANDARD: &'static str = "•";

    /// Construct a validated custom text bullet.
    pub fn new(value: impl Into<Box<str>>) -> Result<Self> {
        let value = value.into();
        let character_count = value.chars().count();
        if character_count == 0 {
            return Err(Error::InvalidFormat(
                "paragraph list bullet must not be empty".to_owned(),
            ));
        }
        if character_count > MAX_PARAGRAPH_LIST_BULLET_CHARACTERS {
            return Err(Error::InvalidFormat(format!(
                "paragraph list bullet must not exceed {MAX_PARAGRAPH_LIST_BULLET_CHARACTERS} characters"
            )));
        }
        if value.chars().any(char::is_control) {
            return Err(Error::InvalidFormat(
                "paragraph list bullet must not contain control characters".to_owned(),
            ));
        }
        Ok(Self(value))
    }

    /// Borrow the marker exactly as iWork displays it.
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

impl Default for ParagraphListBullet {
    fn default() -> Self {
        Self(Self::STANDARD.into())
    }
}

impl AsRef<str> for ParagraphListBullet {
    fn as_ref(&self) -> &str {
        self.as_str()
    }
}

/// A positive, finite bullet size relative to its paragraph text.
///
/// iWork's inspector presents this value as a percentage while the native
/// archive stores a scale ratio.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct ParagraphListBulletScale(f32);

impl ParagraphListBulletScale {
    /// The standard 100% bullet size.
    pub const ONE: Self = Self(1.0);

    /// Construct a scale from the percentage displayed by iWork.
    pub fn from_percent(percent: f32) -> Result<Self> {
        Self::from_ratio(percent / PERCENT_PER_SCALE_UNIT)
    }

    /// Construct a scale from its native text-relative ratio.
    pub fn from_ratio(ratio: f32) -> Result<Self> {
        if !ratio.is_finite() || ratio <= 0.0 {
            return Err(Error::InvalidFormat(
                "paragraph list bullet scale must be positive and finite".to_owned(),
            ));
        }
        Ok(Self(ratio))
    }

    /// Return the percentage displayed by iWork.
    pub fn percent(self) -> f32 {
        self.0 * PERCENT_PER_SCALE_UNIT
    }

    /// Return the native text-relative scale ratio.
    pub const fn ratio(self) -> f32 {
        self.0
    }
}

impl Default for ParagraphListBulletScale {
    fn default() -> Self {
        Self::ONE
    }
}

/// A finite vertical bullet offset measured in typographic points.
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub struct ParagraphListBulletBaselineOffset(f32);

impl ParagraphListBulletBaselineOffset {
    /// No vertical offset.
    pub const ZERO: Self = Self(0.0);

    /// Construct a validated vertical offset in points.
    pub fn from_points(points: f32) -> Result<Self> {
        if !points.is_finite() {
            return Err(Error::InvalidFormat(
                "paragraph list bullet baseline offset must be finite".to_owned(),
            ));
        }
        Ok(Self(points))
    }

    /// Return the vertical offset in points.
    pub const fn points(self) -> f32 {
        self.0
    }
}

/// The editable size and vertical position of one bullet-list marker.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct ParagraphListBulletGeometry {
    pub scale: ParagraphListBulletScale,
    pub baseline_offset: ParagraphListBulletBaselineOffset,
}

impl ParagraphListBulletGeometry {
    /// Construct a bullet geometry from validated components.
    pub const fn new(
        scale: ParagraphListBulletScale,
        baseline_offset: ParagraphListBulletBaselineOffset,
    ) -> Self {
        Self {
            scale,
            baseline_offset,
        }
    }

    /// Apple's standard geometry for a particular list nesting level.
    pub fn standard(level: ParagraphListLevel) -> Self {
        let baseline_points = if level == ParagraphListLevel::ZERO {
            STANDARD_TOP_LEVEL_BULLET_BASELINE_POINTS
        } else {
            STANDARD_NESTED_BULLET_BASELINE_POINTS
        };
        Self::new(
            ParagraphListBulletScale::ONE,
            ParagraphListBulletBaselineOffset(baseline_points),
        )
    }
}

/// A finite, nonnegative list-label offset from the paragraph's left margin.
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd, Default)]
pub struct ParagraphListLabelIndent(f32);

impl ParagraphListLabelIndent {
    /// No offset from the left margin.
    pub const ZERO: Self = Self(0.0);

    /// Construct an absolute label indent in typographic points.
    pub fn from_points(points: f32) -> Result<Self> {
        if !points.is_finite() || points < 0.0 {
            return Err(Error::InvalidFormat(
                "paragraph list label indent must be finite and nonnegative".to_owned(),
            ));
        }
        Ok(Self(points))
    }

    /// Return the absolute label indent in typographic points.
    pub const fn points(self) -> f32 {
        self.0
    }
}

/// A finite, nonnegative gap between a list label and its paragraph text.
///
/// iWork stores this gap in em units so it scales with the paragraph's font.
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd, Default)]
pub struct ParagraphListTextGap(f32);

impl ParagraphListTextGap {
    /// No space between the label and text.
    pub const ZERO: Self = Self(0.0);

    /// Construct a font-relative gap in em units.
    pub fn from_em(em: f32) -> Result<Self> {
        if !em.is_finite() || em < 0.0 {
            return Err(Error::InvalidFormat(
                "paragraph list text gap must be finite and nonnegative".to_owned(),
            ));
        }
        Ok(Self(em))
    }

    /// Construct a gap from its displayed point size and paragraph font size.
    pub fn from_points(points: f32, font_size: TextPointSize) -> Result<Self> {
        Self::from_em(points / font_size.points())
    }

    /// Return the native font-relative gap in em units.
    pub const fn em(self) -> f32 {
        self.0
    }

    /// Return the point size iWork displays for the supplied paragraph font.
    pub fn points_at(self, font_size: TextPointSize) -> f32 {
        self.0 * font_size.points()
    }
}

/// Per-level native list indentation shared by bullets and numbering.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct ParagraphListIndentation {
    pub label_from_margin: ParagraphListLabelIndent,
    pub text_from_label: ParagraphListTextGap,
}

/// Color used to draw a bullet or number label.
///
/// `Automatic` follows the paragraph's text color. `Explicit` keeps the list
/// label color independent from later paragraph text-color changes.
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub enum ParagraphListLabelColor {
    #[default]
    Automatic,
    Explicit(RgbaColor),
}

impl ParagraphListIndentation {
    /// Construct list indentation from validated components.
    pub const fn new(
        label_from_margin: ParagraphListLabelIndent,
        text_from_label: ParagraphListTextGap,
    ) -> Self {
        Self {
            label_from_margin,
            text_from_label,
        }
    }

    /// Apple's standard indentation for a list preset and nesting level.
    pub fn standard(list: ParagraphList, level: ParagraphListLevel) -> Result<Self> {
        let step = match list {
            ParagraphList::Bullet => STANDARD_BULLET_INDENT_STEP_POINTS,
            ParagraphList::Numbered => STANDARD_NUMBER_INDENT_STEP_POINTS,
            ParagraphList::None => {
                return Err(Error::InvalidFormat(
                    "ordinary paragraphs do not have list indentation".to_owned(),
                ));
            },
        };
        Ok(Self::new(
            ParagraphListLabelIndent(step * f32::from(level.get())),
            ParagraphListTextGap(step / STANDARD_FONT_EM_POINTS),
        ))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn list_levels_are_bounded_by_the_native_style_model() {
        assert_eq!(
            ParagraphListLevel::new(0).unwrap(),
            ParagraphListLevel::ZERO
        );
        assert_eq!(ParagraphListLevel::new(1).unwrap(), ParagraphListLevel::ONE);
        assert_eq!(ParagraphListLevel::new(8).unwrap(), ParagraphListLevel::MAX);
        assert!(ParagraphListLevel::new(9).is_err());
    }

    #[test]
    fn list_start_numbers_are_positive_and_numbering_is_typed() {
        assert!(ParagraphListStart::new(0).is_err());
        let seven = ParagraphListStart::new(7).unwrap();
        assert_eq!(seven.get(), 7);
        assert_eq!(
            ParagraphListNumbering::from_native(7).unwrap(),
            ParagraphListNumbering::StartAt(seven)
        );
        assert_eq!(
            ParagraphListNumbering::from_native(0).unwrap(),
            ParagraphListNumbering::Continue
        );
    }

    #[test]
    fn text_bullets_are_nonempty_printable_and_bounded() {
        assert_eq!(
            ParagraphListBullet::default().as_str(),
            ParagraphListBullet::STANDARD
        );
        assert_eq!(ParagraphListBullet::new("➡").unwrap().as_str(), "➡");
        assert!(ParagraphListBullet::new("").is_err());
        assert!(ParagraphListBullet::new("a\nb").is_err());
        assert!(ParagraphListBullet::new("x".repeat(33)).is_err());
    }

    #[test]
    fn bullet_geometry_is_typed_finite_and_level_aware() {
        let scale = ParagraphListBulletScale::from_percent(175.0).unwrap();
        assert_eq!(scale.ratio(), 1.75);
        assert_eq!(scale.percent(), 175.0);
        assert!(ParagraphListBulletScale::from_ratio(0.0).is_err());
        assert!(ParagraphListBulletScale::from_ratio(f32::NAN).is_err());
        assert!(ParagraphListBulletBaselineOffset::from_points(f32::INFINITY).is_err());

        let top = ParagraphListBulletGeometry::standard(ParagraphListLevel::ZERO);
        assert_eq!(top.scale, ParagraphListBulletScale::ONE);
        assert_eq!(top.baseline_offset.points(), 0.0);
        let nested = ParagraphListBulletGeometry::standard(ParagraphListLevel::ONE);
        assert_eq!(nested.baseline_offset.points(), -1.0);
    }

    #[test]
    fn list_indentation_distinguishes_points_from_font_relative_em() {
        let label = ParagraphListLabelIndent::from_points(20.0).unwrap();
        let gap = ParagraphListTextGap::from_points(18.0, TextPointSize::TWELVE).unwrap();
        assert_eq!(label.points(), 20.0);
        assert_eq!(gap.em(), 1.5);
        assert_eq!(gap.points_at(TextPointSize::TWELVE), 18.0);
        assert!(ParagraphListLabelIndent::from_points(-1.0).is_err());
        assert!(ParagraphListTextGap::from_em(f32::NAN).is_err());

        let bullet =
            ParagraphListIndentation::standard(ParagraphList::Bullet, ParagraphListLevel::ONE)
                .unwrap();
        assert_eq!(bullet.label_from_margin.points(), 9.0);
        assert_eq!(bullet.text_from_label.em(), 9.0 / 11.0);
        let numbered =
            ParagraphListIndentation::standard(ParagraphList::Numbered, ParagraphListLevel::ONE)
                .unwrap();
        assert_eq!(numbered.label_from_margin.points(), 18.0);
        assert_eq!(numbered.text_from_label.em(), 18.0 / 11.0);
        assert!(
            ParagraphListIndentation::standard(ParagraphList::None, ParagraphListLevel::ZERO)
                .is_err()
        );
    }
}
