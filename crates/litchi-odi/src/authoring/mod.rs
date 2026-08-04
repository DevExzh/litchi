//! Detached construction for new family packages.

use litchi_core::Result;
use litchi_odf_common::core::PackageWriter;

/// Detached builder; publication validates through the package facade.
#[derive(Clone, Debug)]
pub struct Builder {
    content_xml: String,
}

impl Builder {
    pub fn new() -> Self {
        Self {
            content_xml: empty_content().to_owned(),
        }
    }
    pub fn content_xml(mut self, xml: impl Into<String>) -> Self {
        self.content_xml = xml.into();
        self
    }
    pub fn build(self) -> Result<Vec<u8>> {
        crate::codec::validate_content(&self.content_xml)?;
        let mut writer = PackageWriter::new();
        writer.set_mimetype(crate::package::MIMETYPE)?;
        writer.add_file("content.xml", self.content_xml.as_bytes())?;
        writer.finish_to_bytes()
    }
}

impl Default for Builder {
    fn default() -> Self {
        Self::new()
    }
}

fn empty_content() -> &'static str {
    r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" office:version="1.3"><office:body><office:image/></office:body></office:document-content>"#
}
