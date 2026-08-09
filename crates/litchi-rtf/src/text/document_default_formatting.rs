//! Passive document default-formatting metadata.

use crate::{Formatting, Paragraph, RtfError, RtfResult};

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DefaultFormattingDestination {
    Character,
    Paragraph,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct DocumentDefaultFonts {
    pub primary: Option<u16>,
    pub associated: Option<u16>,
    pub stylesheet_double_byte: Option<u16>,
    pub stylesheet_low_ansi: Option<u16>,
    pub stylesheet_high_ansi: Option<u16>,
    pub stylesheet_bidi: Option<u16>,
}

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct DefaultCharacterProperties {
    pub formatting: Formatting,
    pub low_ansi_font: Option<u16>,
    pub high_ansi_font: Option<u16>,
    pub double_byte_font: Option<u16>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct DefaultParagraphProperties {
    pub paragraph: Paragraph,
    pub table_nesting_level: Option<u8>,
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct DocumentDefaultFormatting {
    pub fonts: DocumentDefaultFonts,
    character: Option<DefaultCharacterProperties>,
    paragraph: Option<DefaultParagraphProperties>,
    destination_order: Vec<DefaultFormattingDestination>,
}

impl DocumentDefaultFormatting {
    #[must_use]
    pub fn character(&self) -> Option<&DefaultCharacterProperties> {
        self.character.as_ref()
    }
    #[must_use]
    pub fn paragraph(&self) -> Option<&DefaultParagraphProperties> {
        self.paragraph.as_ref()
    }
    #[must_use]
    pub fn destination_order(&self) -> &[DefaultFormattingDestination] {
        &self.destination_order
    }

    pub fn set_character(&mut self, value: DefaultCharacterProperties) {
        if self.character.is_none() {
            self.destination_order
                .push(DefaultFormattingDestination::Character);
        }
        self.character = Some(value);
    }
    pub fn set_paragraph(&mut self, value: DefaultParagraphProperties) {
        if self.paragraph.is_none() {
            self.destination_order
                .push(DefaultFormattingDestination::Paragraph);
        }
        self.paragraph = Some(value);
    }
    pub fn clear_character(&mut self) {
        self.character = None;
        self.destination_order
            .retain(|kind| *kind != DefaultFormattingDestination::Character);
    }
    pub fn clear_paragraph(&mut self) {
        self.paragraph = None;
        self.destination_order
            .retain(|kind| *kind != DefaultFormattingDestination::Paragraph);
    }
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn validate(&self) -> RtfResult<()> {
        if self.destination_order.len() > 2
            || self
                .destination_order
                .iter()
                .filter(|kind| **kind == DefaultFormattingDestination::Character)
                .count()
                != usize::from(self.character.is_some())
            || self
                .destination_order
                .iter()
                .filter(|kind| **kind == DefaultFormattingDestination::Paragraph)
                .count()
                != usize::from(self.paragraph.is_some())
        {
            return Err(RtfError::MalformedDocument(
                "RTF default-formatting destination order is inconsistent".to_string(),
            ));
        }
        if let Some(character) = self.character {
            character.formatting.character_positioning.validate()?;
            if let Some(border) = character.formatting.character_border {
                border.validate()?;
            }
            if let Some(shading) = character.formatting.character_shading {
                shading.validate()?;
            }
        }
        if let Some(paragraph) = self.paragraph {
            if paragraph
                .table_nesting_level
                .is_some_and(|value| value > 32)
            {
                return Err(RtfError::MalformedDocument(
                    "RTF defpap itap value must be in 0..=32".to_string(),
                ));
            }
            if paragraph.paragraph.legacy_numbering.is_some() {
                return Err(RtfError::MalformedDocument(
                    "RTF defpap cannot reference a legacy pn destination".to_string(),
                ));
            }
        }
        Ok(())
    }
}
