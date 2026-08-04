use crate::Result;

use super::support::invalid;

use std::error::Error as StdError;
use std::fmt::{Display, Formatter};
use std::str::FromStr;

/// Placement of footnote or endnote text (`ST_FtnPos`/`ST_EdnPos`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum NotePosition {
    /// At the bottom of the page.
    PageBottom,
    /// Immediately beneath the page's text.
    BeneathText,
    /// At the end of the section.
    SectionEnd,
    /// At the end of the document.
    DocumentEnd,
}

impl NotePosition {
    /// Return the exact WordprocessingML token.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::PageBottom => "pageBottom",
            Self::BeneathText => "beneathText",
            Self::SectionEnd => "sectEnd",
            Self::DocumentEnd => "docEnd",
        }
    }

    /// Whether this placement is valid for an endnote.
    pub const fn valid_for_endnote(self) -> bool {
        matches!(self, Self::SectionEnd | Self::DocumentEnd)
    }
}

/// Error returned for a token outside `ST_FtnPos`/`ST_EdnPos`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ParseNotePositionError;

impl Display for ParseNotePositionError {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        formatter.write_str("invalid WordprocessingML note position")
    }
}

impl StdError for ParseNotePositionError {}

impl FromStr for NotePosition {
    type Err = ParseNotePositionError;

    fn from_str(value: &str) -> std::result::Result<Self, Self::Err> {
        match value {
            "pageBottom" => Ok(Self::PageBottom),
            "beneathText" => Ok(Self::BeneathText),
            "sectEnd" => Ok(Self::SectionEnd),
            "docEnd" => Ok(Self::DocumentEnd),
            _ => Err(ParseNotePositionError),
        }
    }
}

impl Display for NotePosition {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(self.as_str())
    }
}

/// Numbering restart behavior for footnotes or endnotes (`w:numRestart`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum NoteNumberingRestart {
    /// Numbering continues throughout the document.
    Continuous,
    /// Numbering restarts at each section.
    EachSection,
    /// Numbering restarts at each page.
    EachPage,
}

impl NoteNumberingRestart {
    /// Parse the schema token.
    pub fn from_xml(value: &str) -> Result<Self> {
        match value {
            "continuous" => Ok(Self::Continuous),
            "eachSect" => Ok(Self::EachSection),
            "eachPage" => Ok(Self::EachPage),
            _ => Err(invalid(format!(
                "invalid note numbering restart value '{value}'"
            ))),
        }
    }

    /// Get the XML value for this restart behavior.
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Continuous => "continuous",
            Self::EachSection => "eachSect",
            Self::EachPage => "eachPage",
        }
    }
}

macro_rules! define_note_formats {
    ($($variant:ident => $token:literal),+ $(,)?) => {
        /// Checked `ST_NumberFormat` value used by note numbering properties.
        #[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
        #[repr(u8)]
        pub enum NoteNumberFormat {
            $($variant,)+
        }

        impl NoteNumberFormat {
            /// Return the exact WordprocessingML token.
            pub const fn as_str(self) -> &'static str {
                match self { $(Self::$variant => $token,)+ }
            }
        }

        impl FromStr for NoteNumberFormat {
            type Err = ParseNoteNumberFormatError;

            fn from_str(value: &str) -> std::result::Result<Self, Self::Err> {
                match value { $($token => Ok(Self::$variant),)+ _ => Err(ParseNoteNumberFormatError) }
            }
        }

        impl Display for NoteNumberFormat {
            fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
                formatter.write_str(self.as_str())
            }
        }
    };
}

/// Error returned for a token outside `ST_NumberFormat`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ParseNoteNumberFormatError;

impl Display for ParseNoteNumberFormatError {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        formatter.write_str("invalid WordprocessingML note number format")
    }
}

impl StdError for ParseNoteNumberFormatError {}

define_note_formats! {
    Decimal => "decimal",
    UpperRoman => "upperRoman",
    LowerRoman => "lowerRoman",
    UpperLetter => "upperLetter",
    LowerLetter => "lowerLetter",
    Ordinal => "ordinal",
    CardinalText => "cardinalText",
    OrdinalText => "ordinalText",
    Hex => "hex",
    Chicago => "chicago",
    IdeographDigital => "ideographDigital",
    JapaneseCounting => "japaneseCounting",
    Aiueo => "aiueo",
    Iroha => "iroha",
    DecimalFullWidth => "decimalFullWidth",
    DecimalHalfWidth => "decimalHalfWidth",
    JapaneseLegal => "japaneseLegal",
    JapaneseDigitalTenThousand => "japaneseDigitalTenThousand",
    DecimalEnclosedCircle => "decimalEnclosedCircle",
    DecimalFullWidth2 => "decimalFullWidth2",
    AiueoFullWidth => "aiueoFullWidth",
    IrohaFullWidth => "irohaFullWidth",
    DecimalZero => "decimalZero",
    Bullet => "bullet",
    Ganada => "ganada",
    Chosung => "chosung",
    DecimalEnclosedFullStop => "decimalEnclosedFullstop",
    DecimalEnclosedParen => "decimalEnclosedParen",
    DecimalEnclosedCircleChinese => "decimalEnclosedCircleChinese",
    IdeographEnclosedCircle => "ideographEnclosedCircle",
    IdeographTraditional => "ideographTraditional",
    IdeographZodiac => "ideographZodiac",
    IdeographZodiacTraditional => "ideographZodiacTraditional",
    TaiwaneseCounting => "taiwaneseCounting",
    IdeographLegalTraditional => "ideographLegalTraditional",
    TaiwaneseCountingThousand => "taiwaneseCountingThousand",
    TaiwaneseDigital => "taiwaneseDigital",
    ChineseCounting => "chineseCounting",
    ChineseLegalSimplified => "chineseLegalSimplified",
    ChineseCountingThousand => "chineseCountingThousand",
    KoreanDigital => "koreanDigital",
    KoreanCounting => "koreanCounting",
    KoreanLegal => "koreanLegal",
    KoreanDigital2 => "koreanDigital2",
    VietnameseCounting => "vietnameseCounting",
    RussianLower => "russianLower",
    RussianUpper => "russianUpper",
    None => "none",
    NumberInDash => "numberInDash",
    Hebrew1 => "hebrew1",
    Hebrew2 => "hebrew2",
    ArabicAlpha => "arabicAlpha",
    ArabicAbjad => "arabicAbjad",
    HindiVowels => "hindiVowels",
    HindiConsonants => "hindiConsonants",
    HindiNumbers => "hindiNumbers",
    HindiCounting => "hindiCounting",
    ThaiLetters => "thaiLetters",
    ThaiNumbers => "thaiNumbers",
    ThaiCounting => "thaiCounting",
    BahtText => "bahtText",
    DollarText => "dollarText",
    Custom => "custom",
}

/// Document-level footnote or endnote properties.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct NoteNumberingProperties<F = NoteNumberFormat> {
    pub(crate) position: Option<NotePosition>,
    pub(crate) format: Option<F>,
    pub(crate) start: Option<u32>,
    pub(crate) restart: Option<NoteNumberingRestart>,
}

impl<F> Default for NoteNumberingProperties<F> {
    fn default() -> Self {
        Self {
            position: None,
            format: None,
            start: None,
            restart: None,
        }
    }
}

impl<F> NoteNumberingProperties<F> {
    /// Construct note properties from validated component values.
    pub fn from_parts(
        position: Option<NotePosition>,
        format: Option<F>,
        start: Option<u32>,
        restart: Option<NoteNumberingRestart>,
    ) -> Self {
        Self {
            position,
            format,
            start,
            restart,
        }
    }

    /// Return the note placement, when specified.
    #[inline]
    pub const fn position(&self) -> Option<NotePosition> {
        self.position
    }

    /// Return the numbering format, when specified.
    #[inline]
    pub const fn format(&self) -> Option<F>
    where
        F: Copy,
    {
        self.format
    }

    /// Return the first note number, when specified.
    #[inline]
    pub const fn start(&self) -> Option<u32> {
        self.start
    }

    /// Return the numbering restart behavior, when specified.
    #[inline]
    pub const fn restart(&self) -> Option<NoteNumberingRestart> {
        self.restart
    }
}

impl NoteNumberingProperties<NoteNumberFormat> {
    pub(crate) fn try_map_format<G, E>(
        self,
        map: &mut impl FnMut(NoteNumberFormat) -> std::result::Result<G, E>,
    ) -> std::result::Result<NoteNumberingProperties<G>, E> {
        Ok(NoteNumberingProperties::from_parts(
            self.position,
            self.format.map(map).transpose()?,
            self.start,
            self.restart,
        ))
    }

    pub(crate) fn to_xml(&self, prefix: &str, name: &str) -> String {
        let mut xml = format!("<{prefix}:{name}>");
        if let Some(position) = self.position {
            xml.push_str(&format!(
                "<{prefix}:pos {prefix}:val=\"{}\"/>",
                position.as_str()
            ));
        }
        if let Some(format) = self.format {
            xml.push_str(&format!(
                "<{prefix}:numFmt {prefix}:val=\"{}\"/>",
                format.as_str()
            ));
        }
        if let Some(start) = self.start {
            xml.push_str(&format!("<{prefix}:numStart {prefix}:val=\"{start}\"/>"));
        }
        if let Some(restart) = self.restart {
            xml.push_str(&format!(
                "<{prefix}:numRestart {prefix}:val=\"{}\"/>",
                restart.as_str()
            ));
        }
        xml.push_str(&format!("</{prefix}:{name}>"));
        xml
    }
}
