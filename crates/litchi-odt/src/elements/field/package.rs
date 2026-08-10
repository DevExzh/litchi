//! Document-content field parsing and package-facing adapters.

#[allow(
    clippy::wildcard_imports,
    reason = "the package facade exposes the complete field model"
)]
use super::model::*;
#[allow(
    clippy::wildcard_imports,
    reason = "the package facade shares the owner-level field vocabulary"
)]
use super::*;
use crate::elements::element::{Element, ElementBase};
use crate::elements::xml::{
    TEXT_NAMESPACE, append_checked, append_text_control, copy_canonical_attributes,
    decode_reference, is_bound,
};
use litchi_core::{Error, Result};
use quick_xml::XmlVersion;
use quick_xml::events::Event;
use quick_xml::reader::NsReader;

/// Utilities for parsing fields from documents
pub struct FieldParser;

impl FieldParser {
    /// Parse all fields from XML content
    pub fn parse_fields(xml_content: &str) -> Result<Vec<Field>> {
        let mut reader = NsReader::from_str(xml_content);
        let mut buffer = Vec::new();
        let mut document_depth = 0usize;
        let mut active: Vec<ActiveField> = Vec::new();
        let mut fields = Vec::new();
        let mut next_order = 0usize;

        loop {
            let (namespace, event) = reader
                .read_resolved_event_into(&mut buffer)
                .map_err(|error| Error::InvalidFormat(format!("invalid field XML: {error}")))?;
            let text_element = is_bound(&namespace, TEXT_NAMESPACE);
            match event {
                Event::Start(ref source) => {
                    document_depth = codec::checked_field_depth(document_depth)?;
                    for field in &mut active {
                        field.depth += 1;
                    }
                    if active
                        .iter()
                        .any(|field| field.element.tag_name() == "text:script")
                    {
                        return Err(Error::InvalidFormat(
                            "text:script cannot contain child elements".to_string(),
                        ));
                    }
                    if text_element {
                        for field in &mut active {
                            append_text_control(&reader, source, &mut field.text)?;
                        }
                        let tag_name = format!(
                            "text:{}",
                            std::str::from_utf8(source.local_name().as_ref()).map_err(
                                |_error| {
                                    Error::InvalidFormat("non-UTF-8 field element name".to_string())
                                }
                            )?
                        );
                        if Field::is_field_tag(&tag_name) {
                            if next_order >= MAX_FIELDS {
                                return Err(Error::InvalidFormat(format!(
                                    "document exceeds {MAX_FIELDS} fields"
                                )));
                            }
                            let mut element = Element::new(&tag_name);
                            copy_canonical_attributes(&reader, source, &mut element, "field")?;
                            active.push(ActiveField {
                                element,
                                text: String::new(),
                                depth: 1,
                                order: next_order,
                            });
                            next_order += 1;
                        }
                    }
                },
                Event::Empty(_)
                    if active
                        .iter()
                        .any(|field| field.element.tag_name() == "text:script") =>
                {
                    return Err(Error::InvalidFormat(
                        "text:script cannot contain child elements".to_string(),
                    ));
                },
                Event::Empty(ref source) if text_element => {
                    for field in &mut active {
                        append_text_control(&reader, source, &mut field.text)?;
                    }
                    let tag_name = format!(
                        "text:{}",
                        std::str::from_utf8(source.local_name().as_ref()).map_err(|_error| {
                            Error::InvalidFormat("non-UTF-8 field element name".to_string())
                        })?
                    );
                    if Field::is_field_tag(&tag_name) {
                        if next_order >= MAX_FIELDS {
                            return Err(Error::InvalidFormat(format!(
                                "document exceeds {MAX_FIELDS} fields"
                            )));
                        }
                        let mut element = Element::new(&tag_name);
                        copy_canonical_attributes(&reader, source, &mut element, "field")?;
                        fields.push((next_order, Field::from_element(element)?));
                        next_order += 1;
                    }
                },
                Event::Text(ref value) if !active.is_empty() => {
                    let value = value
                        .xml_content(XmlVersion::Explicit1_0)
                        .map_err(|error| {
                            Error::InvalidFormat(format!("invalid field text: {error}"))
                        })?;
                    for field in &mut active {
                        append_checked(&mut field.text, &value)?;
                    }
                },
                Event::CData(ref value) if !active.is_empty() => {
                    let value = value
                        .xml_content(XmlVersion::Explicit1_0)
                        .map_err(|error| {
                            Error::InvalidFormat(format!("invalid field CDATA: {error}"))
                        })?;
                    for field in &mut active {
                        append_checked(&mut field.text, &value)?;
                    }
                },
                Event::GeneralRef(ref reference) if !active.is_empty() => {
                    let value = decode_reference(reference, "field")?;
                    for field in &mut active {
                        append_checked(&mut field.text, &value)?;
                    }
                },
                Event::End(_) => {
                    document_depth = document_depth.checked_sub(1).ok_or_else(|| {
                        Error::InvalidFormat("field XML stack underflow".to_string())
                    })?;
                    for field in &mut active {
                        field.depth = field.depth.checked_sub(1).ok_or_else(|| {
                            Error::InvalidFormat("field element stack underflow".to_string())
                        })?;
                    }
                    if let Some(mut field) = active.pop_if(|field| field.depth == 0) {
                        field.element.set_text(&field.text);
                        fields.push((field.order, Field::from_element(field.element)?));
                    }
                },
                Event::DocType(_)
                    if active
                        .iter()
                        .any(|field| field.element.tag_name() == "text:script") =>
                {
                    return Err(Error::InvalidFormat(
                        "DOCTYPE is not permitted in text:script".to_string(),
                    ));
                },
                Event::PI(_)
                    if active
                        .iter()
                        .any(|field| field.element.tag_name() == "text:script") =>
                {
                    return Err(Error::InvalidFormat(
                        "processing instructions are not permitted in text:script".to_string(),
                    ));
                },
                Event::Eof => break,
                _ => {},
            }
            buffer.clear();
        }
        if document_depth != 0 || !active.is_empty() {
            return Err(Error::InvalidFormat(
                "incomplete field XML structure".to_string(),
            ));
        }
        fields.sort_by_key(|(order, _)| *order);
        Ok(fields.into_iter().map(|(_, field)| field).collect())
    }

    /// Parse database fields without contacting any declared database resource.
    pub fn parse_database_fields(xml_content: &str) -> Result<Vec<DatabaseField>> {
        codec::parse_database_fields(xml_content)
    }

    /// Parse typed dynamic text fields without evaluating them.
    pub fn parse_dynamic_text_fields(xml_content: &str) -> Result<Vec<DynamicTextField>> {
        let mut meta_fields = codec::parse_meta_fields(xml_content)?.into_iter();
        let mut drop_down_fields = codec::parse_drop_down_fields(xml_content)?.into_iter();
        let mut result = Vec::new();
        for field in Self::parse_fields(xml_content)? {
            if field.field_type() == "text:meta-field" {
                result.push(meta_fields.next().ok_or_else(|| {
                    Error::InvalidFormat("missing parsed text:meta-field".to_string())
                })?);
            } else if field.field_type() == "text:drop-down" {
                result.push(drop_down_fields.next().ok_or_else(|| {
                    Error::InvalidFormat("missing parsed text:drop-down".to_string())
                })?);
            } else if let Some(field) = field.dynamic_text_field()? {
                result.push(field);
            }
        }
        if meta_fields.next().is_some() {
            return Err(Error::InvalidFormat(
                "unmatched parsed text:meta-field".to_string(),
            ));
        }
        if drop_down_fields.next().is_some() {
            return Err(Error::InvalidFormat(
                "unmatched parsed text:drop-down".to_string(),
            ));
        }
        Ok(result)
    }
}

struct ActiveField {
    element: Element,
    text: String,
    depth: usize,
    order: usize,
}

pub(crate) fn parse_note_body_contents(xml: &str) -> Result<Vec<NoteBodyContent>> {
    codec::parse_note_body_contents(xml)
}
