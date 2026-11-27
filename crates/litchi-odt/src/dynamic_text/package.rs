//! Content-part mutation facade for ODF dynamic and database fields.

use super::{ParagraphSite, Span, codec};
use crate::elements::field::{DatabaseField, DynamicTextField};
use litchi_core::{Error, Result};

/// Insert an inert database field at the end of the selected paragraph.
pub fn insert_database_field_xml(
    xml: &str,
    paragraph_index: usize,
    field: &DatabaseField,
) -> Result<String> {
    let fragment = field.to_xml_fragment()?;
    let scan = codec::scan(xml, Some(paragraph_index))?;
    let site = scan.paragraph.ok_or_else(|| {
        Error::InvalidFormat(format!(
            "database field paragraph index {paragraph_index} is out of bounds"
        ))
    })?;
    match site {
        ParagraphSite::BeforeEnd(offset) => {
            let mut output = String::with_capacity(xml.len().saturating_add(fragment.len()));
            output.push_str(&xml[..offset]);
            output.push_str(&fragment);
            output.push_str(&xml[offset..]);
            Ok(output)
        },
        ParagraphSite::Empty {
            start,
            end,
            qualified_name,
        } => {
            let empty = &xml[start..end];
            let close = empty.rfind("/>").ok_or_else(|| {
                Error::InvalidFormat("invalid empty ODF paragraph encoding".to_string())
            })?;
            let replacement = format!("{}>{}</{}>", &empty[..close], fragment, qualified_name);
            codec::replace_span(xml, Span { start, end }, &replacement)
        },
    }
}

/// Replace the selected database field in document order.
pub fn replace_database_field_xml(
    xml: &str,
    field_index: usize,
    replacement: &DatabaseField,
) -> Result<String> {
    let fragment = replacement.to_xml_fragment()?;
    let mut scan = codec::scan(xml, None)?;
    let span = codec::take_database_field(&mut scan.database_fields, field_index)?;
    codec::replace_span(xml, span, &fragment)
}

/// Remove the selected database field in document order.
pub fn remove_database_field_xml(xml: &str, field_index: usize) -> Result<String> {
    let mut scan = codec::scan(xml, None)?;
    let span = codec::take_database_field(&mut scan.database_fields, field_index)?;
    codec::replace_span(xml, span, "")
}

/// Insert a field at the end of the selected `text:p` in document order.
pub fn insert_dynamic_text_field_xml(
    xml: &str,
    paragraph_index: usize,
    field: &DynamicTextField,
) -> Result<String> {
    let fragment = field.to_xml_fragment()?;
    let scan = codec::scan(xml, Some(paragraph_index))?;
    let site = scan.paragraph.ok_or_else(|| {
        Error::InvalidFormat(format!(
            "dynamic text paragraph index {paragraph_index} is out of bounds"
        ))
    })?;
    match site {
        ParagraphSite::BeforeEnd(offset) => {
            let mut output = String::with_capacity(xml.len().saturating_add(fragment.len()));
            output.push_str(&xml[..offset]);
            output.push_str(&fragment);
            output.push_str(&xml[offset..]);
            codec::validate_meta_field_mutation(output, field)
        },
        ParagraphSite::Empty {
            start,
            end,
            qualified_name,
        } => {
            let empty = &xml[start..end];
            let close = empty.rfind("/>").ok_or_else(|| {
                Error::InvalidFormat("invalid empty ODF paragraph encoding".to_string())
            })?;
            let mut replacement = String::with_capacity(
                close
                    .saturating_add(fragment.len())
                    .saturating_add(qualified_name.len())
                    .saturating_add(4),
            );
            replacement.push_str(&empty[..close]);
            replacement.push('>');
            replacement.push_str(&fragment);
            replacement.push_str("</");
            replacement.push_str(&qualified_name);
            replacement.push('>');
            let output = codec::replace_span(xml, Span { start, end }, &replacement)?;
            codec::validate_meta_field_mutation(output, field)
        },
    }
}

/// Replace the selected dynamic field in document order.
pub fn replace_dynamic_text_field_xml(
    xml: &str,
    field_index: usize,
    replacement: &DynamicTextField,
) -> Result<String> {
    let fragment = replacement.to_xml_fragment()?;
    let mut scan = codec::scan(xml, None)?;
    let span = codec::take_field(&mut scan.fields, field_index)?;
    let output = codec::replace_span(xml, span, &fragment)?;
    codec::validate_meta_field_mutation(output, replacement)
}

/// Remove the selected dynamic field in document order.
pub fn remove_dynamic_text_field_xml(xml: &str, field_index: usize) -> Result<String> {
    let mut scan = codec::scan(xml, None)?;
    let span = codec::take_field(&mut scan.fields, field_index)?;
    codec::replace_span(xml, span, "")
}
