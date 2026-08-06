//! Resource and semantic validation for the ODS annotation owner.

use super::model::{Cell, Entry};
use litchi_core::{Error, Result};
use litchi_odf_common::annotation::{Annotation, Element, Node};
use std::collections::HashSet;

pub(crate) const MAX_XML_BYTES: usize = 256 * 1024 * 1024;
pub(crate) const MAX_ANNOTATIONS: usize = 65_536;
pub(crate) const MAX_EVENTS: usize = 1_000_000;
pub(crate) const MAX_DEPTH: usize = 1_024;
pub(crate) const MAX_SHEET_NAME_BYTES: usize = 1_024;
pub(crate) const MAX_ANNOTATION_TEXT_BYTES: usize = 16 * 1024 * 1024;
pub(crate) const MAX_ANNOTATION_ATTRIBUTES: usize = 65_536;
pub(crate) const MAX_ANNOTATION_ATTRIBUTE_BYTES: usize = 16 * 1024 * 1024;

pub(crate) fn validate_source(source: &str) -> Result<()> {
    if source.len() > MAX_XML_BYTES {
        return Err(Error::InvalidFormat(format!(
            "ODS annotation content.xml exceeds the {MAX_XML_BYTES}-byte limit"
        )));
    }
    Ok(())
}

pub(crate) fn validate_cell(cell: &Cell) -> Result<()> {
    validate_text(cell.sheet(), "ODS annotation sheet name")?;
    if cell.sheet().is_empty() {
        return Err(Error::InvalidFormat(
            "ODS annotation sheet name cannot be empty".to_string(),
        ));
    }
    if cell.sheet().len() > MAX_SHEET_NAME_BYTES {
        return Err(Error::InvalidFormat(
            "ODS annotation sheet name exceeds the size limit".to_string(),
        ));
    }
    if cell.row() >= crate::worksheet::validation::MAX_LOGICAL_ROWS {
        return Err(Error::InvalidFormat(format!(
            "ODS annotation row {} is outside the supported worksheet grid",
            cell.row()
        )));
    }
    if cell.column() >= crate::worksheet::validation::MAX_LOGICAL_COLUMNS {
        return Err(Error::InvalidFormat(format!(
            "ODS annotation column {} is outside the supported worksheet grid",
            cell.column()
        )));
    }
    Ok(())
}

pub(crate) fn validate_annotation(annotation: &Annotation) -> Result<()> {
    annotation.validate()?;
    let mut attribute_count = annotation.attributes().len();
    let mut attribute_bytes = 0usize;
    for (name, value) in annotation.attributes() {
        validate_text(name, "ODS annotation attribute name")?;
        validate_text(value, "ODS annotation attribute value")?;
        attribute_bytes = attribute_bytes
            .checked_add(name.len())
            .and_then(|size| size.checked_add(value.len()))
            .ok_or_else(|| invalid("ODS annotation attribute size overflow"))?;
    }
    for (prefix, uri) in annotation.namespaces() {
        validate_text(prefix, "ODS annotation namespace prefix")?;
        validate_text(uri, "ODS annotation namespace URI")?;
        attribute_bytes = attribute_bytes
            .checked_add(prefix.len())
            .and_then(|size| size.checked_add(uri.len()))
            .ok_or_else(|| invalid("ODS annotation namespace size overflow"))?;
    }
    for element in annotation.children() {
        validate_element(element, &mut attribute_count, &mut attribute_bytes, 1)?;
    }
    if attribute_count > MAX_ANNOTATION_ATTRIBUTES {
        return Err(invalid(
            "ODS annotation attribute count exceeds the safety limit",
        ));
    }
    if attribute_bytes > MAX_ANNOTATION_ATTRIBUTE_BYTES {
        return Err(invalid("ODS annotation attributes exceed the size limit"));
    }
    Ok(())
}

pub(crate) fn validate_entries(entries: &[Entry]) -> Result<()> {
    if entries.len() > MAX_ANNOTATIONS {
        return Err(invalid("ODS annotation count exceeds the safety limit"));
    }
    let mut cells = HashSet::with_capacity(entries.len());
    let mut names = HashSet::new();
    for (index, entry) in entries.iter().enumerate() {
        if entry.index() != index {
            return Err(invalid("ODS annotation indices are not source ordered"));
        }
        validate_cell(entry.cell())?;
        validate_annotation(entry.annotation())?;
        if !cells.insert(entry.cell()) {
            return Err(invalid("multiple ODS annotations target the same cell"));
        }
        if let Some(name) = entry.annotation().name()
            && !names.insert(name.to_string())
        {
            return Err(Error::InvalidFormat(format!(
                "duplicate ODS annotation name '{name}'"
            )));
        }
    }
    Ok(())
}

fn validate_element(
    element: &Element,
    attribute_count: &mut usize,
    attribute_bytes: &mut usize,
    depth: usize,
) -> Result<()> {
    if depth > MAX_DEPTH {
        return Err(invalid("ODS annotation nesting exceeds the safety limit"));
    }
    validate_text(element.name(), "ODS annotation element name")?;
    for (name, value) in element.attributes() {
        *attribute_count = attribute_count
            .checked_add(1)
            .ok_or_else(|| invalid("ODS annotation attribute count overflow"))?;
        validate_text(name, "ODS annotation attribute name")?;
        validate_text(value, "ODS annotation attribute value")?;
        *attribute_bytes = attribute_bytes
            .checked_add(name.len())
            .and_then(|size| size.checked_add(value.len()))
            .ok_or_else(|| invalid("ODS annotation attribute size overflow"))?;
    }
    for child in element.children() {
        match child {
            Node::Text(value) => {
                validate_text(value, "ODS annotation text")?;
                if value.len() > MAX_ANNOTATION_TEXT_BYTES {
                    return Err(invalid("ODS annotation text exceeds the size limit"));
                }
            },
            Node::Element(child) => {
                validate_element(child, attribute_count, attribute_bytes, depth + 1)?;
            },
        }
    }
    Ok(())
}

pub(crate) fn validate_text(value: &str, label: &str) -> Result<()> {
    if value.len() > MAX_ANNOTATION_TEXT_BYTES {
        return Err(Error::InvalidFormat(format!(
            "{label} exceeds the size limit"
        )));
    }
    if value.chars().any(|character| {
        matches!(
            character,
            '\u{0000}'..='\u{0008}'
                | '\u{000B}'..='\u{000C}'
                | '\u{000E}'..='\u{001F}'
        )
    }) {
        return Err(Error::InvalidFormat(format!(
            "{label} contains an XML-forbidden control character"
        )));
    }
    Ok(())
}

fn invalid(message: &str) -> Error {
    Error::InvalidFormat(message.to_string())
}
