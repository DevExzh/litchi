//! Inert legacy section-numbering defaults (`pnseclvl`).

use crate::{RtfError, RtfResult};
use std::borrow::Cow;
use std::collections::HashSet;

pub(crate) const MAX_LEGACY_NUMBERING_LEVELS: usize = 9;
pub(crate) const MAX_LEGACY_NUMBERING_TEXT_BYTES: usize = 4_096;
const MAX_LEGACY_NUMBERING_TWIPS: i32 = 10_000_000;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LegacyNumberingFormat {
    Decimal,
    UpperRoman,
    LowerRoman,
    UpperLetter,
    LowerLetter,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum LegacyNumberingAlignment {
    #[default]
    Left,
    Center,
    Right,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct LegacySectionNumberingLevel<'a> {
    pub level: u8,
    pub format: LegacyNumberingFormat,
    pub start_at: Option<i32>,
    pub indent: Option<i32>,
    pub space: Option<i32>,
    pub hanging: bool,
    pub previous: bool,
    pub alignment: Option<LegacyNumberingAlignment>,
    pub font_ref: Option<u16>,
    pub text_before: Cow<'a, str>,
    pub text_after: Cow<'a, str>,
}

impl<'a> LegacySectionNumberingLevel<'a> {
    pub fn new(level: u8, format: LegacyNumberingFormat) -> Self {
        Self {
            level,
            format,
            start_at: None,
            indent: None,
            space: None,
            hanging: false,
            previous: false,
            alignment: None,
            font_ref: None,
            text_before: Cow::Borrowed(""),
            text_after: Cow::Borrowed(""),
        }
    }

    pub fn validate(&self) -> RtfResult<()> {
        if !(1..=9).contains(&self.level) {
            return Err(RtfError::MalformedDocument(
                "RTF pnseclvl index must be between 1 and 9".to_string(),
            ));
        }
        if self
            .indent
            .into_iter()
            .chain(self.space)
            .any(|value| value.unsigned_abs() > MAX_LEGACY_NUMBERING_TWIPS as u32)
        {
            return Err(RtfError::MalformedDocument(
                "RTF pnseclvl layout value exceeds the safety limit".to_string(),
            ));
        }
        if self.text_before.len() > MAX_LEGACY_NUMBERING_TEXT_BYTES
            || self.text_after.len() > MAX_LEGACY_NUMBERING_TEXT_BYTES
        {
            return Err(RtfError::MalformedDocument(
                "RTF pnseclvl text exceeds the safety limit".to_string(),
            ));
        }
        Ok(())
    }

    pub(crate) fn into_owned(self) -> LegacySectionNumberingLevel<'static> {
        LegacySectionNumberingLevel {
            level: self.level,
            format: self.format,
            start_at: self.start_at,
            indent: self.indent,
            space: self.space,
            hanging: self.hanging,
            previous: self.previous,
            alignment: self.alignment,
            font_ref: self.font_ref,
            text_before: Cow::Owned(self.text_before.into_owned()),
            text_after: Cow::Owned(self.text_after.into_owned()),
        }
    }
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct LegacySectionNumbering<'a> {
    levels: Vec<LegacySectionNumberingLevel<'a>>,
}

impl<'a> LegacySectionNumbering<'a> {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn levels(&self) -> &[LegacySectionNumberingLevel<'a>] {
        &self.levels
    }

    pub fn get(&self, level: u8) -> Option<&LegacySectionNumberingLevel<'a>> {
        self.levels.iter().find(|entry| entry.level == level)
    }

    pub fn add(&mut self, level: LegacySectionNumberingLevel<'a>) -> RtfResult<()> {
        level.validate()?;
        if self.levels.len() >= MAX_LEGACY_NUMBERING_LEVELS {
            return Err(RtfError::MalformedDocument(
                "RTF contains too many pnseclvl destinations".to_string(),
            ));
        }
        if self
            .levels
            .last()
            .is_some_and(|last| last.level >= level.level)
        {
            return Err(RtfError::MalformedDocument(
                "RTF pnseclvl destinations are duplicated or out of order".to_string(),
            ));
        }
        self.levels.push(level);
        Ok(())
    }

    pub fn validate(&self) -> RtfResult<()> {
        if self.levels.len() > MAX_LEGACY_NUMBERING_LEVELS {
            return Err(RtfError::MalformedDocument(
                "RTF contains too many pnseclvl destinations".to_string(),
            ));
        }
        let mut seen = HashSet::with_capacity(self.levels.len());
        let mut previous = 0;
        for level in &self.levels {
            level.validate()?;
            if !seen.insert(level.level) || level.level <= previous {
                return Err(RtfError::MalformedDocument(
                    "RTF pnseclvl destinations are duplicated or out of order".to_string(),
                ));
            }
            previous = level.level;
        }
        Ok(())
    }

    pub(crate) fn into_owned(self) -> LegacySectionNumbering<'static> {
        LegacySectionNumbering {
            levels: self
                .levels
                .into_iter()
                .map(LegacySectionNumberingLevel::into_owned)
                .collect(),
        }
    }
}
