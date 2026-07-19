//! Inert legacy paragraph-numbering (`pn`) metadata.

use crate::{RtfError, RtfResult};
use std::borrow::Cow;

pub const MAX_LEGACY_PARAGRAPH_NUMBERING_RECORDS: usize = 65_536;
pub const MAX_LEGACY_PARAGRAPH_NUMBERING_TEXT_BYTES: usize = 4_096;
const MAX_LAYOUT_TWIPS: i32 = 10_000_000;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LegacyParagraphNumberingLevel {
    Explicit(u8),
    Bullet,
    Body,
    Continue,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LegacyParagraphNumberingFormat {
    Aiueo,
    AiueoDbChar,
    AiueoExtended,
    AiueoExtendedDbChar,
    Chosung,
    CardinalText,
    Decimal,
    DecimalWithPeriod,
    UpperRoman,
    LowerRoman,
    UpperLetter,
    LowerLetter,
    Ordinal,
    OrdinalText,
    ChineseCounting,
    ChineseCountingDbChar,
    ChineseCountingKorean,
    ChineseCountingLegal,
    ChineseCountingThousand,
    ChineseCountingTraditional,
    Ganada,
    GbCounting,
    GbCountingDbChar,
    GbCountingKorean,
    GbCountingLegal,
    GbLip,
    Iroha,
    IrohaDbChar,
    Zodiac,
    ZodiacDbChar,
    ZodiacLegal,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LegacyParagraphNumberingAlignment {
    Left,
    Center,
    Right,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LegacyParagraphNumberingUnderline {
    None,
    Single,
    Dotted,
    Dashed,
    DashDot,
    DashDotDot,
    Double,
    Hairline,
    Thick,
    Words,
    Wave,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LegacyParagraphNumberingBidi {
    A,
    B,
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct LegacyParagraphNumberingRevision {
    pub author: Option<u16>,
    pub date: Option<i32>,
    pub number_format: Option<i32>,
    pub no_tracking: bool,
    pub paragraph_number: Option<i32>,
    pub rgb: Option<u32>,
    pub start: Option<i32>,
    pub stop: Option<i32>,
    pub text_start: Option<i32>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct LegacyParagraphNumbering<'a> {
    pub level: LegacyParagraphNumberingLevel,
    pub format: Option<LegacyParagraphNumberingFormat>,
    pub alignment: Option<LegacyParagraphNumberingAlignment>,
    pub start_at: Option<i32>,
    pub indent: Option<i32>,
    pub space: Option<i32>,
    pub across: bool,
    pub number_once: bool,
    pub previous: bool,
    pub restart: bool,
    pub hanging: bool,
    pub bidi: Option<LegacyParagraphNumberingBidi>,
    pub font_ref: Option<u16>,
    pub font_size: Option<u16>,
    pub color_ref: Option<u16>,
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub caps: Option<bool>,
    pub small_caps: Option<bool>,
    pub strike: Option<bool>,
    pub underline: Option<LegacyParagraphNumberingUnderline>,
    pub text_before: Option<Cow<'a, str>>,
    pub text_after: Option<Cow<'a, str>>,
    pub revision: LegacyParagraphNumberingRevision,
}

impl<'a> LegacyParagraphNumbering<'a> {
    pub fn new(level: LegacyParagraphNumberingLevel) -> Self {
        Self {
            level,
            format: None,
            alignment: None,
            start_at: None,
            indent: None,
            space: None,
            across: false,
            number_once: false,
            previous: false,
            restart: false,
            hanging: false,
            bidi: None,
            font_ref: None,
            font_size: None,
            color_ref: None,
            bold: None,
            italic: None,
            caps: None,
            small_caps: None,
            strike: None,
            underline: None,
            text_before: None,
            text_after: None,
            revision: Default::default(),
        }
    }

    pub fn validate(&self) -> RtfResult<()> {
        if matches!(
            self.level,
            LegacyParagraphNumberingLevel::Explicit(0 | 10..=u8::MAX)
        ) {
            return Err(RtfError::MalformedDocument(
                "RTF pnlvl value must be in 1..=9".to_string(),
            ));
        }
        if matches!(
            self.level,
            LegacyParagraphNumberingLevel::Body | LegacyParagraphNumberingLevel::Explicit(_)
        ) && self.format.is_none()
        {
            return Err(RtfError::MalformedDocument(
                "RTF pn destination is missing its numbering format".to_string(),
            ));
        }
        if self
            .start_at
            .is_some_and(|value| !(0..=32_767).contains(&value))
        {
            return Err(RtfError::MalformedDocument(
                "RTF pnstart value must be in 0..=32767".to_string(),
            ));
        }
        if self
            .indent
            .into_iter()
            .chain(self.space)
            .any(|value| value.unsigned_abs() > MAX_LAYOUT_TWIPS as u32)
        {
            return Err(RtfError::MalformedDocument(
                "RTF pn layout value exceeds the safety limit".to_string(),
            ));
        }
        if self.font_size == Some(0) {
            return Err(RtfError::MalformedDocument(
                "RTF pnfs value must be in 1..=65535".to_string(),
            ));
        }
        if self
            .revision
            .rgb
            .is_some_and(|value| value > i32::MAX as u32)
        {
            return Err(RtfError::MalformedDocument(
                "RTF pnrrgb exceeds the signed RTF parameter range".to_string(),
            ));
        }
        if self
            .text_before
            .as_ref()
            .is_some_and(|value| value.len() > MAX_LEGACY_PARAGRAPH_NUMBERING_TEXT_BYTES)
            || self
                .text_after
                .as_ref()
                .is_some_and(|value| value.len() > MAX_LEGACY_PARAGRAPH_NUMBERING_TEXT_BYTES)
        {
            return Err(RtfError::MalformedDocument(
                "RTF pn text exceeds the safety limit".to_string(),
            ));
        }
        Ok(())
    }

    pub(crate) fn into_owned(self) -> LegacyParagraphNumbering<'static> {
        LegacyParagraphNumbering {
            level: self.level,
            format: self.format,
            alignment: self.alignment,
            start_at: self.start_at,
            indent: self.indent,
            space: self.space,
            across: self.across,
            number_once: self.number_once,
            previous: self.previous,
            restart: self.restart,
            hanging: self.hanging,
            bidi: self.bidi,
            font_ref: self.font_ref,
            font_size: self.font_size,
            color_ref: self.color_ref,
            bold: self.bold,
            italic: self.italic,
            caps: self.caps,
            small_caps: self.small_caps,
            strike: self.strike,
            underline: self.underline,
            text_before: self.text_before.map(|value| Cow::Owned(value.into_owned())),
            text_after: self.text_after.map(|value| Cow::Owned(value.into_owned())),
            revision: self.revision,
        }
    }
}
