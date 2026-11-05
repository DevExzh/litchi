//! Paragraph, row, and mutable-writer integration for extension attributes.

use crate::builder::{Paragraph as DraftParagraph, Row as DraftRow};
use crate::error::Result;
use crate::paragraph::Paragraph;
use crate::paragraph::model::XmlData;
use crate::table::Row;

use super::codec::{parse_paragraph, parse_row};
use super::model::{Extensions, Id, Ids};

impl Paragraph {
    /// Read the typed Word 2010 attributes attached directly to this `w:p`.
    ///
    /// The returned value is a detached snapshot; all paragraph content,
    /// unknown markup, and relationship references remain in the source-backed
    /// paragraph and are not rewritten by this query.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn extensions(&self) -> Result<Extensions> {
        let _parser_admission = self.parser_admission()?;
        // The `w:p` span is parsed standalone and its root has to resolve into
        // the wordprocessing namespace, so the declarations it inherits from
        // `word/document.xml` are re-declared on it first. Change 0653 stopped
        // the markup-compatibility writer from repeating them on every element.
        let xml = self.self_contained_xml()?;
        parse_paragraph(&xml)
    }

    /// Read the Word 2010 `paraId`/`textId` pair from this `w:p`'s start tag.
    ///
    /// Only the start tag is parsed. A paragraph read from a document or
    /// block index resolves its prefixes with the namespace bindings its scan
    /// captured, so the cost is the start tag's, not the paragraph's or its
    /// offset's. A paragraph without captured bindings (a source-backed
    /// managed view) resolves them from its part first, as
    /// [`Self::extensions`] does.
    ///
    /// The identifiers equal those of [`Self::extensions`]; the attributes on
    /// the start tag are checked the same way, and the rest of the paragraph
    /// is not read.
    ///
    /// # Errors
    ///
    /// Returns an error when the start tag is not a `WordprocessingML` `p`,
    /// when an extension attribute is outside the Word 2010 namespace,
    /// repeated or malformed, or when `textId` appears without `paraId`.
    pub(crate) fn extension_ids(&self) -> Result<Ids> {
        let _parser_admission = self.parser_admission()?;
        let namespaces = match &self.xml_data {
            // A range cut from a larger buffer without a scan (the public
            // arena constructors) carries no captured bindings; resolve them
            // from the buffer, as `extensions` does, instead of assuming none.
            XmlData::Shared { slice, namespaces }
                if namespaces.is_empty() && slice.start() != 0 =>
            {
                let reference = self.xml_data.xml_ref()?;
                let lease = reference.namespace_resolver()?;
                let bindings = crate::namespace::resolver_bindings(&lease)?;
                lease.check()?;
                bindings
            },
            data => data.inherited_namespaces()?,
        };
        super::codec::parse_paragraph_start(self.xml_bytes(), &namespaces)
    }

    /// The paragraph's `w14:paraId`, read from its start tag.
    ///
    /// # Errors
    ///
    /// Returns an error for malformed paragraph identifier attributes.
    pub fn para_id(&self) -> Result<Option<Id>> {
        Ok(self.extension_ids()?.para_id())
    }

    /// The paragraph's `w14:textId`, read from its start tag.
    ///
    /// # Errors
    ///
    /// Returns an error for malformed paragraph identifier attributes.
    pub fn text_id(&self) -> Result<Option<Id>> {
        Ok(self.extension_ids()?.text_id())
    }
}

impl Row {
    /// Read the typed Word 2010 `paraId`/`textId` pair attached directly to this
    /// `w:tr`.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn extension_ids(&self) -> Result<Ids> {
        // As for `Paragraph::extensions`: the `w:tr` span is parsed standalone.
        let xml = self.self_contained_xml()?;
        parse_row(&xml)
    }
}

impl DraftParagraph {
    /// Borrow the paragraph's detached Word 2010 extension state.
    #[must_use]
    pub fn extensions(&self) -> &Extensions {
        &self.extension_values
    }

    /// Replace the paragraph extension state after validating dependencies.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
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
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn set_para_id(&mut self, value: Option<Id>) -> Result<&mut Self> {
        self.extension_values.set_para_id(value)?;
        Ok(self)
    }

    /// Set or remove `textId`; a present value requires `paraId`.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
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

impl DraftRow {
    /// Return the row's detached Word 2010 identifier pair.
    #[must_use]
    pub fn extension_ids(&self) -> Ids {
        self.extension_ids
    }

    /// Replace the row identifier pair after validating dependencies.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn set_extension_ids(&mut self, value: Ids) -> Result<&mut Self> {
        value.validate()?;
        self.extension_ids = value;
        Ok(self)
    }

    /// Set or remove the row's `paraId`.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn set_para_id(&mut self, value: Option<Id>) -> Result<&mut Self> {
        self.extension_ids.set_para_id(value)?;
        Ok(self)
    }

    /// Set or remove the row's `textId`; a present value requires `paraId`.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
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
