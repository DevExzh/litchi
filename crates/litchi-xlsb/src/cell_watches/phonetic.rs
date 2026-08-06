//! Typed `BrtPhoneticInfo` defaults.

use super::validation;
use crate::package::error::Result;

/// Phonetic character-set conversion from `[MS-XLSB]` `phType`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Type {
    /// Narrow katakana conversion.
    NarrowKatakana,
    /// Wide katakana conversion.
    WideKatakana,
    /// Hiragana conversion.
    Hiragana,
    /// Display text exactly as entered.
    AsEntered,
}

impl Type {
    pub(crate) const fn wire(self) -> u32 {
        match self {
            Self::NarrowKatakana => 0,
            Self::WideKatakana => 1,
            Self::Hiragana => 2,
            Self::AsEntered => 3,
        }
    }

    pub(crate) fn from_wire(value: u32) -> Result<Self> {
        let result = match value {
            0 => Self::NarrowKatakana,
            1 => Self::WideKatakana,
            2 => Self::Hiragana,
            3 => Self::AsEntered,
            _ => {
                return Err(validation::invalid(
                    "BrtPhoneticInfo phType",
                    value.to_string(),
                ));
            },
        };
        Ok(result)
    }
}

/// Phonetic alignment from `[MS-XLSB]` `phAli`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Alignment {
    /// Left-align all phonetic text over the complete base string.
    AllTextLeft,
    /// Left-align each phonetic run over its base run.
    RunLeft,
    /// Center each phonetic run over its base run.
    Center,
    /// Distribute each phonetic run over its base run.
    Distribute,
}

impl Alignment {
    pub(crate) const fn wire(self) -> u32 {
        match self {
            Self::AllTextLeft => 0,
            Self::RunLeft => 1,
            Self::Center => 2,
            Self::Distribute => 3,
        }
    }

    pub(crate) fn from_wire(value: u32) -> Result<Self> {
        let result = match value {
            0 => Self::AllTextLeft,
            1 => Self::RunLeft,
            2 => Self::Center,
            3 => Self::Distribute,
            _ => {
                return Err(validation::invalid(
                    "BrtPhoneticInfo phAli",
                    value.to_string(),
                ));
            },
        };
        Ok(result)
    }
}

/// Worksheet-wide default phonetic formatting.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Info {
    font_index: u16,
    phonetic_type: Type,
    alignment: Alignment,
}

impl Info {
    /// Construct a validated worksheet phonetic default.
    #[must_use]
    pub const fn new(font_index: u16, phonetic_type: Type, alignment: Alignment) -> Self {
        Self {
            font_index,
            phonetic_type,
            alignment,
        }
    }

    /// Return the zero-based workbook font index.
    #[must_use]
    pub const fn font_index(self) -> u16 {
        self.font_index
    }

    /// Return the phonetic character-set conversion.
    #[must_use]
    pub const fn phonetic_type(self) -> Type {
        self.phonetic_type
    }

    /// Return the phonetic run alignment.
    #[must_use]
    pub const fn alignment(self) -> Alignment {
        self.alignment
    }
}
