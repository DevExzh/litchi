//! Strict list presets and nesting levels shared by Pages, Numbers, and Keynote.

use crate::character::TextPointSize;
use crate::position::TextPosition;
use litchi_iwa_common::color::Rgba;

const MAX_PARAGRAPH_LIST_LEVEL: u8 = 8;
const MAX_PARAGRAPH_LIST_BULLET_CHARACTERS: usize = 32;
const PERCENT_PER_SCALE_UNIT: f32 = 100.0;
const MIN_NUMBER_SCALE_PERCENT: f32 = 1.0;
const MAX_NUMBER_SCALE_PERCENT: f32 = 999.0;
const STANDARD_TOP_LEVEL_BULLET_BASELINE_POINTS: f32 = 0.0;
const STANDARD_NESTED_BULLET_BASELINE_POINTS: f32 = -1.0;
const STANDARD_FONT_EM_POINTS: f32 = 11.0;
const STANDARD_BULLET_INDENT_STEP_POINTS: f32 = 9.0;
const STANDARD_NUMBER_INDENT_STEP_POINTS: f32 = 18.0;

/// Validation failures produced while constructing paragraph-list values.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Error {
    /// A nesting level exceeds the native nine-level model.
    LevelTooHigh { level: u8, maximum: u8 },
    /// A numbered-list restart uses zero rather than a positive number.
    StartNumberZero,
    /// A bullet marker is empty.
    BulletEmpty,
    /// A bullet marker exceeds the native character budget.
    BulletTooLong { characters: usize, maximum: usize },
    /// A bullet marker contains a Unicode control character.
    BulletControlCharacter,
    /// A bullet scale is not finite.
    BulletScaleNonFinite,
    /// A bullet scale is zero or negative.
    BulletScaleNonPositive,
    /// A numbered-list scale is not finite.
    NumberScaleNonFinite,
    /// A numbered-list scale is outside the native inspector range.
    NumberScaleOutOfRange,
    /// A bullet baseline offset is not finite.
    BulletBaselineOffsetNonFinite,
    /// A label indent is not finite.
    LabelIndentNonFinite,
    /// A label indent is negative.
    LabelIndentNegative,
    /// A text gap is not finite.
    TextGapNonFinite,
    /// A text gap is negative.
    TextGapNegative,
    /// An ordinary paragraph has no list indentation.
    IndentationUnavailable,
}

impl std::fmt::Display for Error {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::LevelTooHigh { maximum, .. } => {
                write!(formatter, "paragraph list level must not exceed {maximum}")
            },
            Self::StartNumberZero => {
                formatter.write_str("paragraph list starting number must be positive")
            },
            Self::BulletEmpty => formatter.write_str("paragraph list bullet must not be empty"),
            Self::BulletTooLong { maximum, .. } => write!(
                formatter,
                "paragraph list bullet must not exceed {maximum} characters"
            ),
            Self::BulletControlCharacter => {
                formatter.write_str("paragraph list bullet must not contain control characters")
            },
            Self::BulletScaleNonFinite | Self::BulletScaleNonPositive => {
                formatter.write_str("paragraph list bullet scale must be positive and finite")
            },
            Self::NumberScaleNonFinite | Self::NumberScaleOutOfRange => write!(
                formatter,
                "paragraph list number scale must be between {MIN_NUMBER_SCALE_PERCENT}% and {MAX_NUMBER_SCALE_PERCENT}%"
            ),
            Self::BulletBaselineOffsetNonFinite => {
                formatter.write_str("paragraph list bullet baseline offset must be finite")
            },
            Self::LabelIndentNonFinite | Self::LabelIndentNegative => {
                formatter.write_str("paragraph list label indent must be finite and nonnegative")
            },
            Self::TextGapNonFinite | Self::TextGapNegative => {
                formatter.write_str("paragraph list text gap must be finite and nonnegative")
            },
            Self::IndentationUnavailable => {
                formatter.write_str("ordinary paragraphs do not have list indentation")
            },
        }
    }
}

impl std::error::Error for Error {}

/// Result type for paragraph-list value construction.
pub type Result<T> = std::result::Result<T, Error>;

/// A canonical paragraph-list presentation understood by all three iWork apps.
///
/// The presets describe the complete nine-level native list style rather than
/// exposing unvalidated protobuf integers or partial per-level state.
#[allow(
    clippy::module_name_repetitions,
    reason = "The explicit ParagraphList name remains unambiguous when this value is re-exported by format facades."
)]
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
/// list always begins at [`TextPosition::ZERO`].
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ParagraphListPlacement {
    pub paragraph: TextPosition,
    pub list: ParagraphList,
}

impl ParagraphListPlacement {
    #[must_use]
    pub const fn new(paragraph: TextPosition, list: ParagraphList) -> Self {
        Self { paragraph, list }
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
    ///
    /// # Errors
    ///
    /// Returns [`Error::LevelTooHigh`] when `level` exceeds the native
    /// nine-level style model.
    pub fn new(level: u8) -> Result<Self> {
        if level > MAX_PARAGRAPH_LIST_LEVEL {
            return Err(Error::LevelTooHigh {
                level,
                maximum: MAX_PARAGRAPH_LIST_LEVEL,
            });
        }
        Ok(Self(level))
    }

    /// Return the zero-based native nesting level.
    #[must_use]
    pub const fn get(self) -> u8 {
        self.0
    }
}

/// One effective list-level boundary at a UTF-16 paragraph start.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ParagraphListLevelPlacement {
    pub paragraph: TextPosition,
    pub level: ParagraphListLevel,
}

impl ParagraphListLevelPlacement {
    #[must_use]
    pub const fn new(paragraph: TextPosition, level: ParagraphListLevel) -> Self {
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
    ///
    /// # Errors
    ///
    /// Returns [`Error::StartNumberZero`] when `number` is zero.
    pub fn new(number: u32) -> Result<Self> {
        if number == 0 {
            return Err(Error::StartNumberZero);
        }
        Ok(Self(number))
    }

    /// Return the native positive starting number.
    #[must_use]
    pub const fn get(self) -> u32 {
        self.0
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

impl ParagraphListNumbering {}

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
    ///
    /// # Errors
    ///
    /// Returns a typed validation error when `bullet` is empty, too long, or
    /// contains a control character.
    pub fn new(bullet: impl Into<Box<str>>) -> Result<Self> {
        let value = bullet.into();
        let character_count = value.chars().count();
        if character_count == 0 {
            return Err(Error::BulletEmpty);
        }
        if character_count > MAX_PARAGRAPH_LIST_BULLET_CHARACTERS {
            return Err(Error::BulletTooLong {
                characters: character_count,
                maximum: MAX_PARAGRAPH_LIST_BULLET_CHARACTERS,
            });
        }
        if value.chars().any(char::is_control) {
            return Err(Error::BulletControlCharacter);
        }
        Ok(Self(value))
    }

    /// Borrow the marker exactly as iWork displays it.
    #[must_use]
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
    ///
    /// # Errors
    ///
    /// Returns a typed validation error when `percent` is not finite or is
    /// not positive.
    pub fn from_percent(percent: f32) -> Result<Self> {
        Self::from_ratio(percent / PERCENT_PER_SCALE_UNIT)
    }

    /// Construct a scale from its native text-relative ratio.
    ///
    /// # Errors
    ///
    /// Returns a typed validation error when `ratio` is not finite or is not
    /// positive.
    pub fn from_ratio(ratio: f32) -> Result<Self> {
        if !ratio.is_finite() {
            return Err(Error::BulletScaleNonFinite);
        }
        if ratio <= 0.0 {
            return Err(Error::BulletScaleNonPositive);
        }
        Ok(Self(ratio))
    }

    /// Return the percentage displayed by iWork.
    #[must_use]
    pub fn percent(self) -> f32 {
        self.0 * PERCENT_PER_SCALE_UNIT
    }

    /// Return the native text-relative scale ratio.
    #[must_use]
    pub const fn ratio(self) -> f32 {
        self.0
    }
}

impl Default for ParagraphListBulletScale {
    fn default() -> Self {
        Self::ONE
    }
}

/// A numbered-list label size relative to its paragraph text.
///
/// The range matches the 1%–999% limits enforced by the iWork inspector.
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
pub struct ParagraphListNumberScale(f32);

impl ParagraphListNumberScale {
    /// The standard 100% number-label size.
    pub const ONE: Self = Self(1.0);
    /// Smallest size accepted by iWork's inspector.
    pub const MIN_PERCENT: f32 = MIN_NUMBER_SCALE_PERCENT;
    /// Largest size accepted by iWork's inspector.
    pub const MAX_PERCENT: f32 = MAX_NUMBER_SCALE_PERCENT;

    /// Construct a scale from the percentage displayed by iWork.
    ///
    /// # Errors
    ///
    /// Returns [`Error::NumberScaleNonFinite`] when `percent` is not finite,
    /// or [`Error::NumberScaleOutOfRange`] when it is outside the native
    /// inspector range.
    pub fn from_percent(percent: f32) -> Result<Self> {
        if !percent.is_finite() {
            return Err(Error::NumberScaleNonFinite);
        }
        if !(MIN_NUMBER_SCALE_PERCENT..=MAX_NUMBER_SCALE_PERCENT).contains(&percent) {
            return Err(Error::NumberScaleOutOfRange);
        }
        Ok(Self(percent / PERCENT_PER_SCALE_UNIT))
    }

    /// Construct a scale from its native text-relative ratio.
    ///
    /// # Errors
    ///
    /// Returns a typed validation error when `ratio` is outside the native
    /// inspector range.
    pub fn from_ratio(ratio: f32) -> Result<Self> {
        Self::from_percent(ratio * PERCENT_PER_SCALE_UNIT)
    }

    /// Return the percentage displayed by iWork.
    #[must_use]
    pub fn percent(self) -> f32 {
        self.0 * PERCENT_PER_SCALE_UNIT
    }

    /// Return the native text-relative scale ratio.
    #[must_use]
    pub const fn ratio(self) -> f32 {
        self.0
    }
}

impl Default for ParagraphListNumberScale {
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
    ///
    /// # Errors
    ///
    /// Returns [`Error::BulletBaselineOffsetNonFinite`] when `points` is not
    /// finite.
    pub fn from_points(points: f32) -> Result<Self> {
        if !points.is_finite() {
            return Err(Error::BulletBaselineOffsetNonFinite);
        }
        Ok(Self(points))
    }

    /// Return the vertical offset in points.
    #[must_use]
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
    #[must_use]
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
    #[must_use]
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
    ///
    /// # Errors
    ///
    /// Returns a typed validation error when `points` is not finite or is
    /// negative.
    pub fn from_points(points: f32) -> Result<Self> {
        if !points.is_finite() {
            return Err(Error::LabelIndentNonFinite);
        }
        if points < 0.0 {
            return Err(Error::LabelIndentNegative);
        }
        Ok(Self(points))
    }

    /// Return the absolute label indent in typographic points.
    #[must_use]
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
    ///
    /// # Errors
    ///
    /// Returns a typed validation error when `em` is not finite or is
    /// negative.
    pub fn from_em(em: f32) -> Result<Self> {
        if !em.is_finite() {
            return Err(Error::TextGapNonFinite);
        }
        if em < 0.0 {
            return Err(Error::TextGapNegative);
        }
        Ok(Self(em))
    }

    /// Construct a gap from its displayed point size and paragraph font size.
    ///
    /// # Errors
    ///
    /// Returns a typed validation error when the derived em value is not
    /// finite or is negative.
    pub fn from_points(points: f32, font_size: TextPointSize) -> Result<Self> {
        Self::from_em(points / font_size.points())
    }

    /// Return the native font-relative gap in em units.
    #[must_use]
    pub const fn em(self) -> f32 {
        self.0
    }

    /// Return the point size iWork displays for the supplied paragraph font.
    #[must_use]
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
    Explicit(Rgba),
}

/// Numbering sequence used by an iWork numbered-list label.
///
/// These cover every affix-capable native sequence, including the locale-aware
/// formats that are only shown when the corresponding input language is active.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum ParagraphListNumberSequence {
    #[default]
    Decimal,
    RomanUppercase,
    RomanLowercase,
    LatinUppercase,
    LatinLowercase,
    JapaneseIdeographic,
    JapaneseHiragana,
    JapaneseKatakana,
    JapaneseHiraganaIroha,
    JapaneseKatakanaIroha,
    SimplifiedChineseIdeographic,
    TraditionalChineseIdeographic,
    FormalJapaneseIdeographic,
    FormalSimplifiedChineseIdeographic,
    FormalTraditionalChineseIdeographic,
    KoreanAlphabet,
    ArabicIndic,
    ArabicAlphabet,
    ArabicAbjad,
    HebrewAlphabet,
    HebrewBiblical,
}

impl ParagraphListNumberSequence {
    pub const ALL: [Self; 21] = [
        Self::Decimal,
        Self::RomanUppercase,
        Self::RomanLowercase,
        Self::LatinUppercase,
        Self::LatinLowercase,
        Self::JapaneseIdeographic,
        Self::JapaneseHiragana,
        Self::JapaneseKatakana,
        Self::JapaneseHiraganaIroha,
        Self::JapaneseKatakanaIroha,
        Self::SimplifiedChineseIdeographic,
        Self::TraditionalChineseIdeographic,
        Self::FormalJapaneseIdeographic,
        Self::FormalSimplifiedChineseIdeographic,
        Self::FormalTraditionalChineseIdeographic,
        Self::KoreanAlphabet,
        Self::ArabicIndic,
        Self::ArabicAlphabet,
        Self::ArabicAbjad,
        Self::HebrewAlphabet,
        Self::HebrewBiblical,
    ];
}

/// Punctuation placed around or after an iWork list number.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum ParagraphListNumberPunctuation {
    #[default]
    Period,
    Parentheses,
    RightParenthesis,
}

impl ParagraphListNumberPunctuation {
    pub const ALL: [Self; 3] = [Self::Period, Self::Parentheses, Self::RightParenthesis];
}

/// Complete native number format for a numbered-list label.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ParagraphListNumberFormat {
    /// A number sequence with one of iWork's three standard affix styles.
    Affixed {
        sequence: ParagraphListNumberSequence,
        punctuation: ParagraphListNumberPunctuation,
    },
    /// Circled numbers such as ①, ②, and ③.
    Circled,
    /// Hebrew biblical numbering without additional punctuation.
    HebrewBiblicalStandard,
}

impl Default for ParagraphListNumberFormat {
    fn default() -> Self {
        Self::DECIMAL
    }
}

impl ParagraphListNumberFormat {
    /// Apple's default `1.`, `2.`, `3.` format.
    pub const DECIMAL: Self = Self::Affixed {
        sequence: ParagraphListNumberSequence::Decimal,
        punctuation: ParagraphListNumberPunctuation::Period,
    };

    /// Construct an affixed locale-aware number format.
    #[must_use]
    pub const fn affixed(
        sequence: ParagraphListNumberSequence,
        punctuation: ParagraphListNumberPunctuation,
    ) -> Self {
        Self::Affixed {
            sequence,
            punctuation,
        }
    }
}

/// Whether a numbered-list level displays only its own number or its full hierarchy.
///
/// Tiered numbering renders nested labels such as `1.1` and `1.1.1`; flat
/// numbering renders the current level alone.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum ParagraphListNumberTiering {
    #[default]
    Flat,
    Tiered,
}

impl ParagraphListIndentation {
    /// Construct list indentation from validated components.
    #[must_use]
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
    ///
    /// # Errors
    ///
    /// Returns [`Error::IndentationUnavailable`] for ordinary paragraphs,
    /// which have no list label indentation.
    pub fn standard(list: ParagraphList, level: ParagraphListLevel) -> Result<Self> {
        let step = match list {
            ParagraphList::Bullet => STANDARD_BULLET_INDENT_STEP_POINTS,
            ParagraphList::Numbered => STANDARD_NUMBER_INDENT_STEP_POINTS,
            ParagraphList::None => {
                return Err(Error::IndentationUnavailable);
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
        assert_eq!(ParagraphListStart::new(0), Err(Error::StartNumberZero));
        let seven = ParagraphListStart::new(7).unwrap();
        assert_eq!(seven.get(), 7);
        assert_eq!(
            ParagraphListNumbering::StartAt(seven),
            ParagraphListNumbering::StartAt(seven)
        );
        assert_eq!(
            ParagraphListNumbering::Continue,
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
    fn number_scale_matches_native_inspector_bounds() {
        assert_eq!(
            ParagraphListNumberScale::from_percent(1.0)
                .unwrap()
                .percent(),
            1.0
        );
        assert_eq!(
            ParagraphListNumberScale::from_percent(999.0)
                .unwrap()
                .ratio(),
            9.99
        );
        assert!(ParagraphListNumberScale::from_percent(0.0).is_err());
        assert!(ParagraphListNumberScale::from_percent(1_000.0).is_err());
        assert!(ParagraphListNumberScale::from_ratio(f32::NAN).is_err());
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
