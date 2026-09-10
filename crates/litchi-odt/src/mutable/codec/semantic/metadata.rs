//! Lossless in-content RDFa and `text:meta` operations.

use super::super::super::model::MutableDocument;
use crate::content_metadata::{self, ContentMetadata, RdfaAttributes, TextMeta};
use litchi_core::{Position, Result};

impl MutableDocument {
    /// Inspect RDFa and inline `text:meta` declarations in the authoritative
    /// content snapshot.
    pub fn in_content_metadata(&self) -> Result<ContentMetadata> {
        self.with_content_xml(|xml| content_metadata::parse_part(xml, crate::MetadataPart::Content))
    }

    /// Set or clear RDFa on one paragraph while retaining unrelated XML.
    pub fn set_paragraph_rdfa(&mut self, position: Position, value: &RdfaAttributes) -> Result<()> {
        let updated = self
            .with_content_xml(|xml| content_metadata::set_paragraph_rdfa(xml, position, value))?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Set or clear RDFa on a uniquely named bookmark start.
    pub fn set_bookmark_rdfa(&mut self, name: &str, value: &RdfaAttributes) -> Result<()> {
        let updated =
            self.with_content_xml(|xml| content_metadata::set_bookmark_rdfa(xml, name, value))?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Set or clear RDFa on one inline `text:meta` occurrence.
    pub fn set_text_meta_rdfa(&mut self, position: Position, value: &RdfaAttributes) -> Result<()> {
        let updated = self
            .with_content_xml(|xml| content_metadata::set_text_meta_rdfa(xml, position, value))?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Insert an inline `text:meta` into a paragraph.
    pub fn insert_text_meta(&mut self, paragraph: Position, value: &TextMeta) -> Result<()> {
        let updated =
            self.with_content_xml(|xml| content_metadata::insert_text_meta(xml, paragraph, value))?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Replace one inline `text:meta` occurrence.
    pub fn replace_text_meta(&mut self, position: Position, value: &TextMeta) -> Result<()> {
        let updated =
            self.with_content_xml(|xml| content_metadata::replace_text_meta(xml, position, value))?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Remove one inline `text:meta` occurrence.
    pub fn remove_text_meta(&mut self, position: Position) -> Result<()> {
        let updated =
            self.with_content_xml(|xml| content_metadata::remove_text_meta(xml, position))?;
        self.content_xml = Some(updated);
        Ok(())
    }
}
