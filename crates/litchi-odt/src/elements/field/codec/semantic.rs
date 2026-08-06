#![allow(
    clippy::wildcard_imports,
    reason = "the semantic field codec consumes the complete typed model"
)]

use super::super::model::*;
use super::super::{FORM_NAMESPACE, STYLE_NAMESPACE, TEXT_DATABASE_NAMESPACE, XLINK_NAMESPACE};
use super::validation::{validate_constructed_database_field, validate_database_field};
use litchi_core::Result;

impl DatabaseField {
    pub fn to_xml_fragment(&self) -> Result<String> {
        let field = validate_database_field(self.clone())?;
        validate_constructed_database_field(&field)?;
        let local = match field.kind {
            DatabaseFieldKind::Display => "database-display",
            DatabaseFieldKind::Next => "database-next",
            DatabaseFieldKind::RowSelect => "database-row-select",
            DatabaseFieldKind::RowNumber => "database-row-number",
            DatabaseFieldKind::Name => "database-name",
        };
        let mut xml = format!(
            "<text:{local} xmlns:text=\"{TEXT_DATABASE_NAMESPACE}\" xmlns:style=\"{STYLE_NAMESPACE}\" xmlns:form=\"{FORM_NAMESPACE}\" xmlns:xlink=\"{XLINK_NAMESPACE}\""
        );
        let mut attribute = |prefix: &str, name: &str, value: &str| {
            xml.push(' ');
            xml.push_str(prefix);
            xml.push(':');
            xml.push_str(name);
            xml.push_str("=\"");
            push_xml_attribute(&mut xml, value);
            xml.push('"');
        };
        if let Some(value) = field.source.database_name.as_deref() {
            attribute("text", "database-name", value);
        }
        attribute("text", "table-name", &field.source.table_name);
        if let Some(value) = field.source.table_type {
            attribute("text", "table-type", value.as_str());
        }
        match field.kind {
            DatabaseFieldKind::Display => {
                attribute(
                    "text",
                    "column-name",
                    field.column_name.as_deref().expect("validated"),
                );
                if let Some(value) = field.data_style_name.as_deref() {
                    attribute("style", "data-style-name", value);
                }
            },
            DatabaseFieldKind::Next => {
                if let Some(value) = field.condition.as_deref() {
                    attribute("text", "condition", value);
                }
            },
            DatabaseFieldKind::RowSelect => {
                if let Some(value) = field.condition.as_deref() {
                    attribute("text", "condition", value);
                }
                if let Some(value) = field.row_number {
                    attribute("text", "row-number", value.as_str());
                }
            },
            DatabaseFieldKind::RowNumber => {
                if let Some(value) = field.value {
                    attribute("text", "value", value.as_str());
                }
                if let Some(value) = field.number_format.as_deref() {
                    attribute("style", "num-format", value);
                }
                if let Some(value) = field.number_letter_sync {
                    attribute(
                        "style",
                        "num-letter-sync",
                        if value { "true" } else { "false" },
                    );
                }
            },
            DatabaseFieldKind::Name => {},
        }
        let _ = attribute;
        if field.source.connection_resource.is_none() && field.display_text.is_empty() {
            xml.push_str("/>");
            return Ok(xml);
        }
        xml.push('>');
        if let Some(resource) = &field.source.connection_resource {
            xml.push_str("<form:connection-resource xlink:href=\"");
            push_xml_attribute(&mut xml, &resource.href);
            xml.push_str("\"/>");
        }
        push_xml_text(&mut xml, &field.display_text);
        xml.push_str("</text:");
        xml.push_str(local);
        xml.push('>');
        Ok(xml)
    }
}
