//! Paragraph, row, and mutable-writer integration for extension attributes.

use crate::error::Result;
use crate::paragraph::Paragraph;
use crate::table::Row;
use crate::writer::{MutableParagraph, MutableRow};

use super::codec::{parse_paragraph, parse_row};
use super::model::{Extensions, Id, Ids};

impl Paragraph {
    /// Read the typed Word 2010 attributes attached directly to this `w:p`.
    ///
    /// The returned value is a detached snapshot; all paragraph content,
    /// unknown markup, and relationship references remain in the source-backed
    /// paragraph and are not rewritten by this query.
    pub fn extensions(&self) -> Result<Extensions> {
        parse_paragraph(self.xml_bytes())
    }
}

impl Row {
    /// Read the typed Word 2010 `paraId`/`textId` pair attached directly to this
    /// `w:tr`.
    pub fn extension_ids(&self) -> Result<Ids> {
        parse_row(self.xml_bytes())
    }
}

impl MutableParagraph {
    /// Borrow the paragraph's detached Word 2010 extension state.
    #[must_use]
    pub fn extensions(&self) -> &Extensions {
        &self.extension_values
    }

    /// Replace the paragraph extension state after validating dependencies.
    pub fn set_extensions(&mut self, value: Extensions) -> Result<&mut Self> {
        value.validate()?;
        self.extension_values = value;
        Ok(self)
    }

    /// Return the paragraph's `paraId`/`textId` pair.
    #[must_use]
    pub fn extension_ids(&self) -> Ids {
        self.extension_values.ids()
    }

    /// Set or remove `paraId`.
    pub fn set_para_id(&mut self, value: Option<Id>) -> Result<&mut Self> {
        self.extension_values.set_para_id(value)?;
        Ok(self)
    }

    /// Set or remove `textId`; a present value requires `paraId`.
    pub fn set_text_id(&mut self, value: Option<Id>) -> Result<&mut Self> {
        self.extension_values.set_text_id(value)?;
        Ok(self)
    }

    /// Set or remove the paragraph's explicit spelling-result marker.
    pub fn set_no_spell_err(&mut self, value: Option<bool>) -> &mut Self {
        self.extension_values.set_no_spell_err(value);
        self
    }

    /// Remove all paragraph extension attributes.
    pub fn clear_extensions(&mut self) -> &mut Self {
        self.extension_values = Extensions::new();
        self
    }
}

impl MutableRow {
    /// Return the row's detached Word 2010 identifier pair.
    #[must_use]
    pub fn extension_ids(&self) -> Ids {
        self.extension_ids
    }

    /// Replace the row identifier pair after validating dependencies.
    pub fn set_extension_ids(&mut self, value: Ids) -> Result<&mut Self> {
        value.validate()?;
        self.extension_ids = value;
        Ok(self)
    }

    /// Set or remove the row's `paraId`.
    pub fn set_para_id(&mut self, value: Option<Id>) -> Result<&mut Self> {
        self.extension_ids.set_para_id(value)?;
        Ok(self)
    }

    /// Set or remove the row's `textId`; a present value requires `paraId`.
    pub fn set_text_id(&mut self, value: Option<Id>) -> Result<&mut Self> {
        self.extension_ids.set_text_id(value)?;
        Ok(self)
    }

    /// Remove all row extension identifiers.
    pub fn clear_extension_ids(&mut self) -> &mut Self {
        self.extension_ids = Ids::new();
        self
    }
}
