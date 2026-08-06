//! Paragraph and document facade integration for `w12:collapsed`.

use crate::error::{Error, Result};
use crate::paragraph::Paragraph;
use crate::writer::{MutableDocument, MutableParagraph};

use super::codec;
use super::model::Collapsed;

impl Paragraph {
    /// Return the direct Word 2012 collapse state, if present.
    pub fn collapsed(&self) -> Result<Option<Collapsed>> {
        codec::read(self.xml_bytes())
    }

    /// Replace or remove the direct Word 2012 collapse state.
    ///
    /// The paragraph is changed only after the candidate XML has been parsed
    /// and validated. Unknown paragraph and `pPr` content remains untouched.
    pub fn set_collapsed(&mut self, value: Option<Collapsed>) -> Result<&mut Self> {
        let original = self.xml_bytes();
        let rewritten = codec::rewrite(original, value)?;
        if rewritten.as_slice() != original {
            self.xml_data = super::super::model::XmlData::Owned(rewritten.into_boxed_slice());
        }
        Ok(self)
    }
}

impl MutableParagraph {
    /// Return the authored collapse state for this paragraph.
    #[must_use]
    pub fn collapsed(&self) -> Option<Collapsed> {
        self.collapsed
    }

    /// Set or remove the collapse marker on a newly authored paragraph.
    #[must_use]
    pub fn set_collapsed(&mut self, value: Option<Collapsed>) -> &mut Self {
        self.collapsed = value;
        self
    }

    /// Mark this paragraph as collapsing subsequent deeper headings.
    #[must_use]
    pub fn collapse(&mut self) -> &mut Self {
        self.set_collapsed(Some(Collapsed::Enabled))
    }

    /// Explicitly disable collapse behavior for subsequent deeper headings.
    #[must_use]
    pub fn expand(&mut self) -> &mut Self {
        self.set_collapsed(Some(Collapsed::Disabled))
    }

    /// Remove the direct collapse marker.
    #[must_use]
    pub fn clear_collapsed(&mut self) -> &mut Self {
        self.set_collapsed(None)
    }
}

impl MutableDocument {
    /// Read the collapse state of any body paragraph, including preserved
    /// paragraphs from an opened document.
    pub fn paragraph_collapsed(&self, index: usize) -> Result<Option<Collapsed>> {
        self.body.paragraph_collapsed(index)
    }

    /// Atomically set or remove the collapse marker on any body paragraph.
    ///
    /// Existing paragraphs are edited through the bounded snapshot codec, so
    /// opaque runs, relationships, and foreign `pPr` children remain intact.
    pub fn set_paragraph_collapsed(
        &mut self,
        index: usize,
        value: Option<Collapsed>,
    ) -> Result<&mut Self> {
        self.body.set_paragraph_collapsed(index, value)?;
        self.modified = true;
        Ok(self)
    }
}

impl crate::writer::doc::DocumentBody {
    pub(super) fn paragraph_collapsed(&self, index: usize) -> Result<Option<Collapsed>> {
        let position = self
            .paragraph_positions()
            .get(index)
            .copied()
            .ok_or_else(|| Error::OutOfBounds {
                object: "paragraph",
                index,
                len: self.paragraph_count(),
            })?;
        match self.elements.get(position) {
            Some(crate::writer::doc::BodyElement::Paragraph(paragraph)) => {
                Ok(paragraph.collapsed())
            },
            Some(crate::writer::doc::BodyElement::PreservedParagraph(xml)) => {
                codec::read(xml.as_bytes())
            },
            _ => Err(Error::InvalidFormat(
                "paragraph position does not contain a paragraph".to_owned(),
            )),
        }
    }

    pub(super) fn set_paragraph_collapsed(
        &mut self,
        index: usize,
        value: Option<Collapsed>,
    ) -> Result<()> {
        let position = self
            .paragraph_positions()
            .get(index)
            .copied()
            .ok_or_else(|| Error::OutOfBounds {
                object: "paragraph",
                index,
                len: self.paragraph_count(),
            })?;
        match self.elements.get_mut(position) {
            Some(crate::writer::doc::BodyElement::Paragraph(paragraph)) => {
                let _ = paragraph.set_collapsed(value);
                Ok(())
            },
            Some(crate::writer::doc::BodyElement::PreservedParagraph(xml)) => {
                let replacement = codec::rewrite(xml.as_bytes(), value)?;
                let replacement = String::from_utf8(replacement).map_err(|error| {
                    Error::InvalidFormat(format!("collapsed paragraph XML is not UTF-8: {error}"))
                })?;
                *xml = replacement;
                Ok(())
            },
            _ => Err(Error::InvalidFormat(
                "paragraph position does not contain a paragraph".to_owned(),
            )),
        }
    }
}
