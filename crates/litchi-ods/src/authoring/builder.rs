use crate::model::{NamedDefinition, NamedExpression, NamedRange};
use litchi_core::Result;
use litchi_odf_common::core::PackageWriter;

const MIMETYPE: &str = "application/vnd.oasis.opendocument.spreadsheet";

/// Minimal package builder; richer sheet authoring is migrated independently.
#[derive(Clone, Debug)]
pub struct SpreadsheetBuilder {
    content_xml: String,
    named_definitions: Vec<NamedDefinition>,
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
            named_definitions: Vec::new(),
        }
    }

    pub fn content_xml(mut self, xml: impl Into<String>) -> Self {
        self.content_xml = xml.into();
        self
    }

    /// Return authored named definitions in their insertion order.
    pub fn named_definitions(&self) -> &[NamedDefinition] {
        &self.named_definitions
    }

    /// Append a validated named range to the builder.
    pub fn add_named_range(&mut self, range: NamedRange) -> Result<&mut Self> {
        self.add_named_definition(range.into())
    }

    /// Append a validated named expression to the builder.
    pub fn add_named_expression(&mut self, expression: NamedExpression) -> Result<&mut Self> {
        self.add_named_definition(expression.into())
    }

    /// Append a named definition while preserving authored order.
    pub fn add_named_definition(&mut self, definition: NamedDefinition) -> Result<&mut Self> {
        let mut candidate = self.named_definitions.clone();
        candidate.push(definition);
        crate::model::named_expression::validate_named_definition_collection(&candidate)?;
        self.named_definitions = candidate;
        Ok(self)
    }

    pub fn build(self) -> Result<Vec<u8>> {
        let content_xml = if self.named_definitions.is_empty() {
            self.content_xml
        } else {
            crate::codec::named_expression::replace(&self.content_xml, &self.named_definitions)?
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
