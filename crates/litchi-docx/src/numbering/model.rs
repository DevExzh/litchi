//! Package-neutral WordprocessingML numbering definitions.
//!
//! The owner keeps the compact semantic vocabulary here; the OOXML host
//! supplies only package and relationship orchestration around it.

use std::error::Error;
use std::fmt::{Display, Formatter};
use std::str::FromStr;

#[derive(Debug, Clone)]
pub struct Collection {
    pub abstract_nums: Vec<Definition>,
    pub nums: Vec<Instance>,
    pub picture_bullets: Vec<PictureBullet>,
}

/// A picture bullet definition (`w:numPicBullet`) from `numbering.xml`.
///
/// The image itself lives in a package part referenced through a relationship;
/// only the inert relationship ID is captured here, never the image bytes.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PictureBullet {
    pub(crate) id: u32,
    pub(crate) image_relationship_id: Option<String>,
}

impl PictureBullet {
    /// The `w:numPicBulletId` key referenced by `w:lvlPicBulletId` on a level.
    pub fn id(&self) -> u32 {
        self.id
    }

    /// Relationship ID of the bullet image, when the definition carries one.
    pub fn image_relationship_id(&self) -> Option<&str> {
        self.image_relationship_id.as_deref()
    }
}

#[derive(Debug, Clone)]
pub struct Definition {
    pub id: u32,
    pub num_type: Option<MultiLevel>,
    pub num_style_link: Option<String>,
    pub style_link: Option<String>,
    /// Word 2012 policy for restarting this definition at the next section.
    ///
    /// `None` means that the extension attribute was absent. An explicit
    /// `Some(false)` remains distinct from absence, as required by the
    /// `ST_OnOff` attribute in `[MS-DOCX]` §2.5.2.1.
    pub restart_numbering_after_break: Option<bool>,
    pub levels: Vec<Level>,
}

#[derive(Debug, Clone)]
pub struct Instance {
    pub id: u32,
    pub abstract_num_id: u32,
    pub overrides: Vec<Override>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Paragraph {
    pub num_id: u32,
    pub level: u8,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Restart {
    Default,
    Never,
    After(u8),
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Suffix {
    Tab,
    Space,
    Nothing,
}

/// Structure of an abstract numbering definition (`ST_MultiLevelType`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum MultiLevel {
    Single,
    Multi,
    Hybrid,
}

impl MultiLevel {
    /// Return the exact WordprocessingML token.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Single => "singleLevel",
            Self::Multi => "multilevel",
            Self::Hybrid => "hybridMultilevel",
        }
    }
}

/// Error returned for a token outside `ST_MultiLevelType`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ParseMultiLevelError;

impl Display for ParseMultiLevelError {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        formatter.write_str("invalid WordprocessingML multi-level type")
    }
}

impl Error for ParseMultiLevelError {}

impl FromStr for MultiLevel {
    type Err = ParseMultiLevelError;

    fn from_str(value: &str) -> Result<Self, Self::Err> {
        match value {
            "singleLevel" => Ok(Self::Single),
            "multilevel" => Ok(Self::Multi),
            "hybridMultilevel" => Ok(Self::Hybrid),
            _ => Err(ParseMultiLevelError),
        }
    }
}

impl Display for MultiLevel {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(self.as_str())
    }
}

/// Error returned for a token outside `ST_NumberFormat`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ParseFormatError;

impl Display for ParseFormatError {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        formatter.write_str("invalid WordprocessingML number format")
    }
}

impl Error for ParseFormatError {}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum Format {
    Decimal,
    UpperRoman,
    LowerRoman,
    UpperLetter,
    LowerLetter,
    Ordinal,
    CardinalText,
    OrdinalText,
    Hex,
    Chicago,
    IdeographDigital,
    JapaneseCounting,
    Aiueo,
    Iroha,
    DecimalFullWidth,
    DecimalHalfWidth,
    JapaneseLegal,
    JapaneseDigitalTenThousand,
    DecimalEnclosedCircle,
    DecimalFullWidth2,
    AiueoFullWidth,
    IrohaFullWidth,
    DecimalZero,
    Bullet,
    Ganada,
    Chosung,
    DecimalEnclosedFullStop,
    DecimalEnclosedParen,
    DecimalEnclosedCircleChinese,
    IdeographEnclosedCircle,
    IdeographTraditional,
    IdeographZodiac,
    IdeographZodiacTraditional,
    TaiwaneseCounting,
    IdeographLegalTraditional,
    TaiwaneseCountingThousand,
    TaiwaneseDigital,
    ChineseCounting,
    ChineseLegalSimplified,
    ChineseCountingThousand,
    KoreanDigital,
    KoreanCounting,
    KoreanLegal,
    KoreanDigital2,
    VietnameseCounting,
    RussianLower,
    RussianUpper,
    None,
    NumberInDash,
    Hebrew1,
    Hebrew2,
    ArabicAlpha,
    ArabicAbjad,
    HindiVowels,
    HindiConsonants,
    HindiNumbers,
    HindiCounting,
    ThaiLetters,
    ThaiNumbers,
    ThaiCounting,
    BahtText,
    DollarText,
    Custom,
}

impl FromStr for Format {
    type Err = ParseFormatError;

    fn from_str(value: &str) -> Result<Self, Self::Err> {
        Ok(match value {
            "decimal" => Self::Decimal,
            "upperRoman" => Self::UpperRoman,
            "lowerRoman" => Self::LowerRoman,
            "upperLetter" => Self::UpperLetter,
            "lowerLetter" => Self::LowerLetter,
            "ordinal" => Self::Ordinal,
            "cardinalText" => Self::CardinalText,
            "ordinalText" => Self::OrdinalText,
            "hex" => Self::Hex,
            "chicago" => Self::Chicago,
            "ideographDigital" => Self::IdeographDigital,
            "japaneseCounting" => Self::JapaneseCounting,
            "aiueo" => Self::Aiueo,
            "iroha" => Self::Iroha,
            "decimalFullWidth" => Self::DecimalFullWidth,
            "decimalHalfWidth" => Self::DecimalHalfWidth,
            "japaneseLegal" => Self::JapaneseLegal,
            "japaneseDigitalTenThousand" => Self::JapaneseDigitalTenThousand,
            "decimalEnclosedCircle" => Self::DecimalEnclosedCircle,
            "decimalFullWidth2" => Self::DecimalFullWidth2,
            "aiueoFullWidth" => Self::AiueoFullWidth,
            "irohaFullWidth" => Self::IrohaFullWidth,
            "decimalZero" => Self::DecimalZero,
            "bullet" => Self::Bullet,
            "ganada" => Self::Ganada,
            "chosung" => Self::Chosung,
            "decimalEnclosedFullstop" => Self::DecimalEnclosedFullStop,
            "decimalEnclosedParen" => Self::DecimalEnclosedParen,
            "decimalEnclosedCircleChinese" => Self::DecimalEnclosedCircleChinese,
            "ideographEnclosedCircle" => Self::IdeographEnclosedCircle,
            "ideographTraditional" => Self::IdeographTraditional,
            "ideographZodiac" => Self::IdeographZodiac,
            "ideographZodiacTraditional" => Self::IdeographZodiacTraditional,
            "taiwaneseCounting" => Self::TaiwaneseCounting,
            "ideographLegalTraditional" => Self::IdeographLegalTraditional,
            "taiwaneseCountingThousand" => Self::TaiwaneseCountingThousand,
            "taiwaneseDigital" => Self::TaiwaneseDigital,
            "chineseCounting" => Self::ChineseCounting,
            "chineseLegalSimplified" => Self::ChineseLegalSimplified,
            "chineseCountingThousand" => Self::ChineseCountingThousand,
            "koreanDigital" => Self::KoreanDigital,
            "koreanCounting" => Self::KoreanCounting,
            "koreanLegal" => Self::KoreanLegal,
            "koreanDigital2" => Self::KoreanDigital2,
            "vietnameseCounting" => Self::VietnameseCounting,
            "russianLower" => Self::RussianLower,
            "russianUpper" => Self::RussianUpper,
            "none" => Self::None,
            "numberInDash" => Self::NumberInDash,
            "hebrew1" => Self::Hebrew1,
            "hebrew2" => Self::Hebrew2,
            "arabicAlpha" => Self::ArabicAlpha,
            "arabicAbjad" => Self::ArabicAbjad,
            "hindiVowels" => Self::HindiVowels,
            "hindiConsonants" => Self::HindiConsonants,
            "hindiNumbers" => Self::HindiNumbers,
            "hindiCounting" => Self::HindiCounting,
            "thaiLetters" => Self::ThaiLetters,
            "thaiNumbers" => Self::ThaiNumbers,
            "thaiCounting" => Self::ThaiCounting,
            "bahtText" => Self::BahtText,
            "dollarText" => Self::DollarText,
            "custom" => Self::Custom,
            _ => return Err(ParseFormatError),
        })
    }
}

impl Format {
    /// Return the exact WordprocessingML token.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Decimal => "decimal",
            Self::UpperRoman => "upperRoman",
            Self::LowerRoman => "lowerRoman",
            Self::UpperLetter => "upperLetter",
            Self::LowerLetter => "lowerLetter",
            Self::Ordinal => "ordinal",
            Self::CardinalText => "cardinalText",
            Self::OrdinalText => "ordinalText",
            Self::Hex => "hex",
            Self::Chicago => "chicago",
            Self::IdeographDigital => "ideographDigital",
            Self::JapaneseCounting => "japaneseCounting",
            Self::Aiueo => "aiueo",
            Self::Iroha => "iroha",
            Self::DecimalFullWidth => "decimalFullWidth",
            Self::DecimalHalfWidth => "decimalHalfWidth",
            Self::JapaneseLegal => "japaneseLegal",
            Self::JapaneseDigitalTenThousand => "japaneseDigitalTenThousand",
            Self::DecimalEnclosedCircle => "decimalEnclosedCircle",
            Self::DecimalFullWidth2 => "decimalFullWidth2",
            Self::AiueoFullWidth => "aiueoFullWidth",
            Self::IrohaFullWidth => "irohaFullWidth",
            Self::DecimalZero => "decimalZero",
            Self::Bullet => "bullet",
            Self::Ganada => "ganada",
            Self::Chosung => "chosung",
            Self::DecimalEnclosedFullStop => "decimalEnclosedFullstop",
            Self::DecimalEnclosedParen => "decimalEnclosedParen",
            Self::DecimalEnclosedCircleChinese => "decimalEnclosedCircleChinese",
            Self::IdeographEnclosedCircle => "ideographEnclosedCircle",
            Self::IdeographTraditional => "ideographTraditional",
            Self::IdeographZodiac => "ideographZodiac",
            Self::IdeographZodiacTraditional => "ideographZodiacTraditional",
            Self::TaiwaneseCounting => "taiwaneseCounting",
            Self::IdeographLegalTraditional => "ideographLegalTraditional",
            Self::TaiwaneseCountingThousand => "taiwaneseCountingThousand",
            Self::TaiwaneseDigital => "taiwaneseDigital",
            Self::ChineseCounting => "chineseCounting",
            Self::ChineseLegalSimplified => "chineseLegalSimplified",
            Self::ChineseCountingThousand => "chineseCountingThousand",
            Self::KoreanDigital => "koreanDigital",
            Self::KoreanCounting => "koreanCounting",
            Self::KoreanLegal => "koreanLegal",
            Self::KoreanDigital2 => "koreanDigital2",
            Self::VietnameseCounting => "vietnameseCounting",
            Self::RussianLower => "russianLower",
            Self::RussianUpper => "russianUpper",
            Self::None => "none",
            Self::NumberInDash => "numberInDash",
            Self::Hebrew1 => "hebrew1",
            Self::Hebrew2 => "hebrew2",
            Self::ArabicAlpha => "arabicAlpha",
            Self::ArabicAbjad => "arabicAbjad",
            Self::HindiVowels => "hindiVowels",
            Self::HindiConsonants => "hindiConsonants",
            Self::HindiNumbers => "hindiNumbers",
            Self::HindiCounting => "hindiCounting",
            Self::ThaiLetters => "thaiLetters",
            Self::ThaiNumbers => "thaiNumbers",
            Self::ThaiCounting => "thaiCounting",
            Self::BahtText => "bahtText",
            Self::DollarText => "dollarText",
            Self::Custom => "custom",
        }
    }
}

impl Display for Format {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(self.as_str())
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Level {
    pub level: u8,
    pub start: i64,
    pub format: Format,
    pub custom_format: Option<String>,
    pub level_text: Option<String>,
    pub suffix: Suffix,
    pub restart: Restart,
    pub legal: bool,
    pub paragraph_style: Option<String>,
    pub picture_bullet_id: Option<u32>,
}

impl Level {
    pub(super) fn new(level: u8) -> Self {
        Self {
            level,
            start: 0,
            format: Format::Decimal,
            custom_format: None,
            level_text: None,
            suffix: Suffix::Tab,
            restart: Restart::Default,
            legal: false,
            paragraph_style: None,
            picture_bullet_id: None,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Override {
    pub level: u8,
    pub start_override: Option<i64>,
    pub definition: Option<Level>,
}

impl Collection {
    pub fn new() -> Self {
        Self {
            abstract_nums: Vec::new(),
            nums: Vec::new(),
            picture_bullets: Vec::new(),
        }
    }

    pub fn abstract_nums(&self) -> &[Definition] {
        &self.abstract_nums
    }
    pub fn nums(&self) -> &[Instance] {
        &self.nums
    }
    pub fn abstract_num_count(&self) -> usize {
        self.abstract_nums.len()
    }
    pub fn num_count(&self) -> usize {
        self.nums.len()
    }
    pub fn get_abstract_num(&self, id: u32) -> Option<&Definition> {
        self.abstract_nums.iter().find(|value| value.id == id)
    }
    pub fn get_num(&self, id: u32) -> Option<&Instance> {
        self.nums.iter().find(|value| value.id == id)
    }
    pub fn picture_bullets(&self) -> &[PictureBullet] {
        &self.picture_bullets
    }
    pub fn get_picture_bullet(&self, id: u32) -> Option<&PictureBullet> {
        self.picture_bullets.iter().find(|value| value.id == id)
    }
}

impl Default for Collection {
    fn default() -> Self {
        Self::new()
    }
}

impl Definition {
    pub fn id(&self) -> u32 {
        self.id
    }
    pub fn num_type(&self) -> Option<MultiLevel> {
        self.num_type
    }
    pub fn num_style_link(&self) -> Option<&str> {
        self.num_style_link.as_deref()
    }
    pub fn style_link(&self) -> Option<&str> {
        self.style_link.as_deref()
    }
    /// Return the optional Word 2012 section-break restart policy.
    #[must_use]
    pub const fn restart_numbering_after_break(&self) -> Option<bool> {
        self.restart_numbering_after_break
    }
    pub fn levels(&self) -> &[Level] {
        &self.levels
    }
    pub fn level(&self, level: u8) -> Option<&Level> {
        self.levels.iter().find(|value| value.level == level)
    }
}

impl Instance {
    pub fn id(&self) -> u32 {
        self.id
    }
    pub fn abstract_num_id(&self) -> u32 {
        self.abstract_num_id
    }
    pub fn overrides(&self) -> &[Override] {
        &self.overrides
    }
}
