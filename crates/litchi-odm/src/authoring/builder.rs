//! Master-document package authoring.

use litchi_core::Result;
use litchi_odf_common::{compact_xml, core::PackageWriter};

/// Detached builder; publication validates through the package facade.
#[derive(Clone, Debug)]
pub struct Builder {
    content_xml: String,
    body_items: Vec<crate::transaction::BodyItemSpec>,
}

impl Builder {
    /// Creates a builder pre-filled with an empty master-document content
    /// document.
    #[must_use]
    pub fn new() -> Self {
        Self {
            content_xml: empty_content().to_owned(),
            body_items: Vec::new(),
        }
    }

    /// Replaces the `content.xml` payload.
    #[must_use]
    pub fn content_xml(mut self, xml: impl Into<String>) -> Self {
        self.content_xml = xml.into();
        self
    }

    /// Appends one typed direct-body item during publication.
    ///
    /// Typed items remain separate from an optional custom `content.xml`
    /// source until the complete compact document is validated.
    #[must_use]
    pub fn body_item(mut self, item: crate::transaction::BodyItemSpec) -> Self {
        self.body_items.push(item);
        self
    }

    /// Validates the document and returns the immutable package.
    ///
    /// Output (bytes, streams, files) lives on the returned
    /// [`crate::Package`].
    ///
    /// # Errors
    ///
    /// Returns an error if content validation or package writing fails.
    pub fn build(self) -> std::result::Result<crate::Package, litchi_odf_common::Error> {
        Ok(crate::Package::from_vec(self.build_bytes()?)?)
    }

    fn build_bytes(self) -> Result<Vec<u8>> {
        compact_xml::validate(self.content_xml.as_bytes(), compact_xml::Limits::default())?;
        crate::codec::validate(&self.content_xml)?;
        let mut content_xml = self.content_xml;
        for item in &self.body_items {
            content_xml = crate::edit_ops::append_body_item(&content_xml, item)?;
        }
        compact_xml::validate(content_xml.as_bytes(), compact_xml::Limits::default())?;
        crate::codec::validate(&content_xml)?;
        let mut writer =
            PackageWriter::new(litchi_odf_common::core::StreamingArchiveLimits::default());
        writer.set_mimetype(crate::package::MIMETYPE)?;
        writer.add_file("content.xml", content_xml.as_bytes())?;
        writer.finish_to_bytes()
    }
}

impl Default for Builder {
    fn default() -> Self {
        Self::new()
    }
}

fn empty_content() -> &'static str {
    r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" office:version="1.3"><office:body><office:text/></office:body></office:document-content>"#
}
