use crate::model::names::{Definition, Expression, Range};
use litchi_core::Result;
use litchi_odf_common::core::PackageWriter;

const MIMETYPE: &str = "application/vnd.oasis.opendocument.spreadsheet";

/// Minimal package builder; richer sheet authoring is migrated independently.
#[derive(Clone, Debug)]
pub struct Builder {
    content_xml: String,
    definitions: Vec<Definition>,
}

impl Default for Builder {
    fn default() -> Self {
        Self::new()
    }
}

impl Builder {
    pub fn new() -> Self {
        Self {
            content_xml: empty_content().to_owned(),
            definitions: Vec::new(),
        }
    }

    pub fn content_xml(mut self, xml: impl Into<String>) -> Self {
        self.content_xml = xml.into();
        self
    }

    /// Return authored named definitions in their insertion order.
    pub fn definitions(&self) -> &[Definition] {
        &self.definitions
    }

    /// Append a validated named range to the builder.
    pub fn add_range(&mut self, range: Range) -> Result<&mut Self> {
        self.add_definition(range.into())
    }

    /// Append a validated named expression to the builder.
    pub fn add_expression(&mut self, expression: Expression) -> Result<&mut Self> {
        self.add_definition(expression.into())
    }

    /// Append a named definition while preserving authored order.
    pub fn add_definition(&mut self, definition: Definition) -> Result<&mut Self> {
        let mut candidate = self.definitions.clone();
        candidate.push(definition);
        crate::model::names::validate_collection(&candidate)?;
        self.definitions = candidate;
        Ok(self)
    }

    pub fn build(self) -> Result<Vec<u8>> {
        let content_xml = if self.definitions.is_empty() {
            self.content_xml
        } else {
            crate::codec::names::replace(&self.content_xml, &self.definitions)?
        };
        let mut writer = PackageWriter::new();
        writer.set_mimetype(MIMETYPE)?;
        writer.add_file("content.xml", content_xml.as_bytes())?;
        writer.finish_to_bytes()
    }
}

fn empty_content() -> &'static str {
    r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" office:version="1.3"><office:body><office:spreadsheet/></office:body></office:document-content>"#
}
