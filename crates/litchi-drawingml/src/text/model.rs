//! Semantic `DrawingML` text domains.
//!
//! These compact, host-neutral values describe the closed schema domains
//! used by text bodies; lexical behavior is implemented by `codec`.

/// Vertical anchoring within a text body (`ST_TextAnchoringType`).
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum Anchor {
    /// Start at the top inset (`t`).
    #[default]
    Top,
    /// Center vertically (`ctr`).
    Center,
    /// End at the bottom inset (`b`).
    Bottom,
    /// Spread lines to fill the body (`just`).
    Justified,
    /// Spread words to fill the body (`dist`).
    Distributed,
}

/// Direction of text within a shape (`ST_TextVerticalType`).
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum Direction {
    /// Horizontal text (`horz`).
    #[default]
    Horizontal,
    /// Lines rotated 90 degrees (`vert`).
    Vertical,
    /// Lines rotated 270 degrees (`vert270`).
    Vertical270,
    /// Upright, stacked `WordArt` letters (`wordArtVert`).
    WordArtVertical,
    /// East Asian vertical text (`eaVert`).
    EastAsianVertical,
    /// Mongolian vertical text (`mongolianVert`).
    MongolianVertical,
    /// Right-to-left `WordArt` vertical text (`wordArtVertRtl`).
    WordArtVerticalRtl,
}

/// Whether text wraps inside the shape extents (`ST_TextWrappingType`).
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum Wrap {
    /// Wrap inside the bounding rectangle (`square`).
    #[default]
    Square,
    /// Do not wrap (`none`).
    None,
}

/// Autofit behavior selected by a text-body child element.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum Autofit {
    /// `a:noAutofit`.
    #[default]
    None,
    /// `a:spAutoFit`.
    Shape,
    /// `a:normAutofit`.
    Normal,
}

/// Lossless underline style shared by `DrawingML` and `WordprocessingML`.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum Underline {
    #[default]
    None,
    Words,
    Single,
    Double,
    Heavy,
    Dotted,
    DottedHeavy,
    Dash,
    DashHeavy,
    DashLong,
    DashLongHeavy,
    DotDash,
    DotDashHeavy,
    DotDotDash,
    DotDotDashHeavy,
    Wavy,
    WavyHeavy,
    WavyDouble,
}

/// Strikethrough applied to a run (`ST_TextStrikeType`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum Strike {
    /// No strikethrough (`noStrike`).
    None,
    /// One line through the text (`sngStrike`).
    Single,
    /// Two lines through the text (`dblStrike`).
    Double,
}

/// Capitalization applied to a run (`ST_TextCapsType`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum Caps {
    /// Text as written (`none`).
    None,
    /// Lowercase letters drawn as reduced capitals (`small`).
    Small,
    /// Every letter drawn as a capital (`all`).
    All,
}

/// Horizontal alignment of a paragraph (`ST_TextAlignType`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum Alignment {
    /// Aligned to the left margin (`l`).
    Left,
    /// Centered (`ctr`).
    Center,
    /// Aligned to the right margin (`r`).
    Right,
    /// Justified between both margins (`just`).
    Justified,
    /// Justified with Kashida-friendly low spacing (`justLow`).
    JustifiedLow,
    /// Characters distributed across the line (`dist`).
    Distributed,
    /// Thai distributed justification (`thaiDist`).
    ThaiDistributed,
}

/// Vertical alignment of characters within a line (`ST_TextFontAlignType`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum FontAlignment {
    /// Chosen by the renderer (`auto`).
    Auto,
    /// Aligned to the top of the line (`t`).
    Top,
    /// Centered on the line (`ctr`).
    Center,
    /// Aligned to the baseline (`base`).
    Baseline,
    /// Aligned to the bottom of the line (`b`).
    Bottom,
}

/// Alignment of text at a custom tab stop (`ST_TextTabAlignType`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum TabAlignment {
    /// Text starts at the stop (`l`).
    Left,
    /// Text is centered on the stop (`ctr`).
    Center,
    /// Text ends at the stop (`r`).
    Right,
    /// The decimal separator sits at the stop (`dec`).
    Decimal,
}

/// Numbering scheme of an automatically numbered bullet
/// (`ST_TextAutonumberScheme`).
///
/// Every one of the 41 schema values is a variant; [`AutoNumber::token`]
/// returns its exact spelling.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum AutoNumber {
    AlphaLowerParenBoth,
    AlphaUpperParenBoth,
    AlphaLowerParenRight,
    AlphaUpperParenRight,
    AlphaLowerPeriod,
    AlphaUpperPeriod,
    ArabicParenBoth,
    ArabicParenRight,
    ArabicPeriod,
    ArabicPlain,
    RomanLowerParenBoth,
    RomanUpperParenBoth,
    RomanLowerParenRight,
    RomanUpperParenRight,
    RomanLowerPeriod,
    RomanUpperPeriod,
    CircleNumberDoubleBytePlain,
    CircleNumberWingdingsBlackPlain,
    CircleNumberWingdingsWhitePlain,
    ArabicDoubleBytePeriod,
    ArabicDoubleBytePlain,
    EastAsianSimplifiedChinesePeriod,
    EastAsianSimplifiedChinesePlain,
    EastAsianTraditionalChinesePeriod,
    EastAsianTraditionalChinesePlain,
    EastAsianJapaneseChineseDoubleBytePeriod,
    EastAsianJapaneseKoreanPlain,
    EastAsianJapaneseKoreanPeriod,
    Arabic1Minus,
    Arabic2Minus,
    Hebrew2Minus,
    ThaiAlphaPeriod,
    ThaiAlphaParenRight,
    ThaiAlphaParenBoth,
    ThaiNumberPeriod,
    ThaiNumberParenRight,
    ThaiNumberParenBoth,
    HindiAlphaPeriod,
    HindiNumberPeriod,
    HindiNumberParenRight,
    HindiAlpha1Period,
}
