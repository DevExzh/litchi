use litchi_core::Result;
use litchi_odf_common::core::PackageWriter;

const MIMETYPE: &str = "application/vnd.oasis.opendocument.spreadsheet";

/// Minimal package builder; richer sheet authoring is migrated independently.
#[derive(Clone, Debug)]
pub struct SpreadsheetBuilder {
    content_xml: String,
}

impl Default for SpreadsheetBuilder {
    fn default() -> Self {
        Self::new()
    }
}

impl SpreadsheetBuilder {
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
        let mut writer = PackageWriter::new();
        writer.set_mimetype(MIMETYPE)?;
        writer.add_file("content.xml", self.content_xml.as_bytes())?;
        writer.finish_to_bytes()
    }
}

fn empty_content() -> &'static str {
    r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" office:version="1.3"><office:body><office:spreadsheet/></office:body></office:document-content>"#
}
