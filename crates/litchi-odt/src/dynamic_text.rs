//! Namespace-aware mutation of dynamic text fields in ODT `content.xml`.

use crate::elements::field::{DatabaseField, DynamicTextField, FieldParser};
use litchi_core::{Error, Result};
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

const TEXT_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const MAX_CONTENT_XML_BYTES: usize = 64 * 1024 * 1024;
const MAX_DYNAMIC_FIELDS: usize = 1_000_000;
const MAX_PARAGRAPHS: usize = 1_000_000;
const MAX_XML_DEPTH: usize = 4_096;

#[derive(Debug)]
struct Span {
    start: usize,
    end: usize,
}

#[cfg(test)]
mod meta_field_mutation_tests {
    use super::*;
    use crate::elements::field::{DynamicTextField, MetaFieldContent, MetaFieldNode};

    fn content_xml(body: &str) -> String {
        format!(
            r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content
 xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
 xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
 <office:body><office:text><text:p>{body}</text:p></office:text></office:body>
</office:document-content>"#
        )
    }

    #[test]
    fn nested_meta_fields_and_dynamic_children_are_individually_mutable() {
        let xml = content_xml(
            r#"<text:meta-field xml:id="m">before<text:date>2026-07-18</text:date>after</text:meta-field>"#,
        );
        let spans = scan(&xml, None).unwrap().fields;
        assert_eq!(spans.len(), 2);
        assert!(spans[0].end > spans[1].end);

        let xml = replace_dynamic_text_field_xml(
            &xml,
            1,
            &DynamicTextField::PageVariableGet {
                number_format: None,
                display_text: "value".to_string(),
            },
        )
        .unwrap();
        assert!(xml.contains("text:meta-field"));
        assert!(xml.contains("text:page-variable-get"));
        assert!(!xml.contains("text:date"));

        let xml = remove_dynamic_text_field_xml(&xml, 0).unwrap();
        assert!(!xml.contains("text:meta-field"));
        assert!(!xml.contains("text:page-variable-get"));
    }

    #[test]
    fn meta_field_can_be_inserted_and_out_of_bounds_is_rejected() {
        let xml = content_xml("plain");
        let field = DynamicTextField::MetaField {
            xml_id: "inserted".to_string(),
            data_style_name: None,
            content: MetaFieldContent::new(vec![MetaFieldNode::Text("metadata".to_string())])
                .unwrap(),
        };
        let xml = insert_dynamic_text_field_xml(&xml, 0, &field).unwrap();
        assert!(xml.contains("xml:id=\"inserted\""));
        assert!(xml.contains(">metadata</text:meta-field>"));
        assert!(replace_dynamic_text_field_xml(&xml, 99, &field).is_err());
        assert!(remove_dynamic_text_field_xml(&xml, 99).is_err());
    }

    #[test]
    fn meta_field_mutation_rejects_document_wide_xml_id_collisions() {
        let xml = content_xml(r#"<text:span xml:id="existing">plain</text:span>"#);
        let field = DynamicTextField::MetaField {
            xml_id: "existing".to_string(),
            data_style_name: None,
            content: MetaFieldContent::new(vec![MetaFieldNode::Text("metadata".to_string())])
                .unwrap(),
        };
        assert!(insert_dynamic_text_field_xml(&xml, 0, &field).is_err());
    }
}

#[cfg(test)]
mod database_field_mutation_tests {
    use super::*;
    use crate::elements::field::{DatabaseFieldKind, DatabaseSource, NonNegativeInteger};

    const XML: &str = r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><o:body><o:text><t:p>before<t:database-name t:table-name="old">old</t:database-name>after</t:p></o:text></o:body></o:document-content>"#;

    fn field(kind: DatabaseFieldKind, table: &str, text: &str) -> DatabaseField {
        DatabaseField {
            kind,
            source: DatabaseSource {
                database_name: None,
                table_name: table.into(),
                table_type: None,
                connection_resource: None,
            },
            column_name: None,
            condition: None,
            row_number: None,
            value: None,
            data_style_name: None,
            number_format: None,
            number_letter_sync: None,
            display_text: text.into(),
        }
    }

    #[test]
    fn database_fields_insert_replace_remove_and_check_bounds() {
        let replacement = field(DatabaseFieldKind::Name, "new", "new");
        let xml = replace_database_field_xml(XML, 0, &replacement).unwrap();
        assert!(xml.contains("text:table-name=\"new\""));
        let inserted =
            insert_database_field_xml(&xml, 0, &field(DatabaseFieldKind::Next, "next", ""))
                .unwrap();
        assert_eq!(
            FieldParser::parse_database_fields(&inserted).unwrap().len(),
            2
        );
        let removed = remove_database_field_xml(&inserted, 0).unwrap();
        assert_eq!(
            FieldParser::parse_database_fields(&removed).unwrap().len(),
            1
        );
        assert!(replace_database_field_xml(&removed, 9, &replacement).is_err());
        assert!(remove_database_field_xml(&removed, 9).is_err());
    }

    #[test]
    fn database_mutation_preserves_arbitrary_width_row_numbers() {
        let mut row = field(DatabaseFieldKind::RowSelect, "table", "");
        let huge = "18446744073709551616000000000000000000";
        row.row_number = Some(NonNegativeInteger::new(&format!("+000{huge}")).unwrap());
        let xml = insert_database_field_xml(XML, 0, &row).unwrap();
        assert!(xml.contains(&format!("text:row-number=\"{huge}\"")));
        let parsed = FieldParser::parse_database_fields(&xml).unwrap();
        assert_eq!(
            parsed.last().unwrap().row_number.as_ref().unwrap().as_str(),
            huge
        );
    }
}

#[cfg(test)]
mod dde_connection_mutation_tests {
    use super::*;

    const XML: &str = r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><o:body><o:text><t:p>before<t:dde-connection t:connection-name="old">cached</t:dde-connection>after</t:p></o:text></o:body></o:document-content>"#;

    #[test]
    fn dde_cached_fields_insert_replace_remove_and_check_bounds() {
        let replacement = DynamicTextField::DdeConnection {
            connection_name: "new".into(),
            display_text: "new cache".into(),
        };
        let xml = replace_dynamic_text_field_xml(XML, 0, &replacement).unwrap();
        assert!(xml.contains("text:connection-name=\"new\""));
        assert_eq!(
            FieldParser::parse_dynamic_text_fields(&xml).unwrap(),
            vec![replacement.clone()]
        );
        let inserted = insert_dynamic_text_field_xml(&xml, 0, &replacement).unwrap();
        assert_eq!(
            FieldParser::parse_dynamic_text_fields(&inserted)
                .unwrap()
                .len(),
            2
        );
        let removed = remove_dynamic_text_field_xml(&inserted, 0).unwrap();
        assert_eq!(
            FieldParser::parse_dynamic_text_fields(&removed)
                .unwrap()
                .len(),
            1
        );
        assert!(replace_dynamic_text_field_xml(&removed, 9, &replacement).is_err());

        let missing = XML.replace(" t:connection-name=\"old\"", "");
        assert!(FieldParser::parse_dynamic_text_fields(&missing).is_err());
        let spoof = XML
            .replace(
                "<t:dde-connection",
                "<fake:dde-connection xmlns:fake=\"urn:not-text\"",
            )
            .replace("</t:dde-connection>", "</fake:dde-connection>");
        assert!(
            FieldParser::parse_dynamic_text_fields(&spoof)
                .unwrap()
                .is_empty()
        );
        let oversized = DynamicTextField::DdeConnection {
            connection_name: "x".repeat(65_537),
            display_text: String::new(),
        };
        assert!(oversized.to_xml_fragment().is_err());
    }
}

#[derive(Debug)]
enum ParagraphSite {
    BeforeEnd(usize),
    Empty {
        start: usize,
        end: usize,
        qualified_name: String,
    },
}

#[derive(Debug, Default)]
struct Scan {
    fields: Vec<Span>,
    database_fields: Vec<Span>,
    paragraph: Option<ParagraphSite>,
}

/// Insert an inert database field at the end of the selected paragraph.
pub fn insert_database_field_xml(
    xml: &str,
    paragraph_index: usize,
    field: &DatabaseField,
) -> Result<String> {
    let fragment = field.to_xml_fragment()?;
    let scan = scan(xml, Some(paragraph_index))?;
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
            replace_span(xml, Span { start, end }, &replacement)
        },
    }
}

pub fn replace_database_field_xml(
    xml: &str,
    field_index: usize,
    replacement: &DatabaseField,
) -> Result<String> {
    let fragment = replacement.to_xml_fragment()?;
    let mut scan = scan(xml, None)?;
    let span = take_database_field(&mut scan.database_fields, field_index)?;
    replace_span(xml, span, &fragment)
}

pub fn remove_database_field_xml(xml: &str, field_index: usize) -> Result<String> {
    let mut scan = scan(xml, None)?;
    let span = take_database_field(&mut scan.database_fields, field_index)?;
    replace_span(xml, span, "")
}

fn take_database_field(fields: &mut Vec<Span>, index: usize) -> Result<Span> {
    if index >= fields.len() {
        return Err(Error::InvalidFormat(format!(
            "database field index {index} is out of bounds (field count: {})",
            fields.len()
        )));
    }
    Ok(fields.swap_remove(index))
}

/// Insert a field at the end of the selected `text:p` in document order.
pub fn insert_dynamic_text_field_xml(
    xml: &str,
    paragraph_index: usize,
    field: &DynamicTextField,
) -> Result<String> {
    let fragment = field.to_xml_fragment()?;
    let scan = scan(xml, Some(paragraph_index))?;
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
            validate_meta_field_mutation(output, field)
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
            let output = replace_span(xml, Span { start, end }, &replacement)?;
            validate_meta_field_mutation(output, field)
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
    let mut scan = scan(xml, None)?;
    let span = take_field(&mut scan.fields, field_index)?;
    let output = replace_span(xml, span, &fragment)?;
    validate_meta_field_mutation(output, replacement)
}

fn validate_meta_field_mutation(output: String, field: &DynamicTextField) -> Result<String> {
    if matches!(
        field,
        DynamicTextField::MetaField { .. }
            | DynamicTextField::DropDown { .. }
            | DynamicTextField::Script { .. }
    ) {
        FieldParser::parse_dynamic_text_fields(&output)?;
    }
    Ok(output)
}

/// Remove the selected dynamic field in document order.
pub fn remove_dynamic_text_field_xml(xml: &str, field_index: usize) -> Result<String> {
    let mut scan = scan(xml, None)?;
    let span = take_field(&mut scan.fields, field_index)?;
    replace_span(xml, span, "")
}

fn take_field(fields: &mut Vec<Span>, index: usize) -> Result<Span> {
    if index >= fields.len() {
        return Err(Error::InvalidFormat(format!(
            "dynamic text field index {index} is out of bounds (field count: {})",
            fields.len()
        )));
    }
    Ok(fields.swap_remove(index))
}

fn replace_span(xml: &str, span: Span, replacement: &str) -> Result<String> {
    let capacity = xml
        .len()
        .checked_sub(span.end - span.start)
        .and_then(|len| len.checked_add(replacement.len()))
        .ok_or_else(|| Error::InvalidFormat("dynamic text XML size overflow".to_string()))?;
    if capacity > MAX_CONTENT_XML_BYTES {
        return Err(Error::InvalidFormat(format!(
            "mutated content.xml exceeds {MAX_CONTENT_XML_BYTES} bytes"
        )));
    }
    let mut output = String::with_capacity(capacity);
    output.push_str(&xml[..span.start]);
    output.push_str(replacement);
    output.push_str(&xml[span.end..]);
    Ok(output)
}

fn scan(xml: &str, wanted_paragraph: Option<usize>) -> Result<Scan> {
    if xml.len() > MAX_CONTENT_XML_BYTES {
        return Err(Error::InvalidFormat(format!(
            "content.xml exceeds {MAX_CONTENT_XML_BYTES} bytes"
        )));
    }
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut paragraph_count = 0usize;
    let mut active_paragraph_depth = None;
    let mut active_fields: Vec<(usize, usize, bool)> = Vec::new();
    let mut active_database_fields: Vec<(usize, usize)> = Vec::new();
    let mut output = Scan::default();

    loop {
        let event_start = usize::try_from(reader.buffer_position()).map_err(|_| {
            Error::InvalidFormat("content.xml byte offset exceeds platform limits".to_string())
        })?;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid content.xml: {error}")))?;
        let text_namespace = matches!(
            namespace,
            ResolveResult::Bound(Namespace(value)) if value == TEXT_NAMESPACE
        );
        match event {
            Event::Start(ref element) => {
                depth = depth.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("content.xml nesting depth overflow".to_string())
                })?;
                if depth > MAX_XML_DEPTH {
                    return Err(Error::InvalidFormat(format!(
                        "content.xml nesting exceeds {MAX_XML_DEPTH} levels"
                    )));
                }
                let local = element.local_name();
                if text_namespace && local.as_ref() == b"p" {
                    if paragraph_count >= MAX_PARAGRAPHS {
                        return Err(Error::InvalidFormat("too many ODF paragraphs".to_string()));
                    }
                    if wanted_paragraph == Some(paragraph_count) {
                        active_paragraph_depth = Some(depth);
                    }
                    paragraph_count += 1;
                }
                if text_namespace && is_dynamic_local(local.as_ref()) {
                    if active_fields
                        .last()
                        .is_some_and(|(_, _, allows_nested)| !allows_nested)
                    {
                        return Err(Error::InvalidFormat(
                            "nested dynamic text fields are not allowed".to_string(),
                        ));
                    }
                    if output.fields.len() >= MAX_DYNAMIC_FIELDS {
                        return Err(Error::InvalidFormat(
                            "too many dynamic text fields".to_string(),
                        ));
                    }
                    active_fields.push((event_start, depth, local.as_ref() == b"meta-field"));
                }
                if text_namespace && is_database_local(local.as_ref()) {
                    if !active_database_fields.is_empty() {
                        return Err(Error::InvalidFormat(
                            "nested database fields are not allowed".to_string(),
                        ));
                    }
                    active_database_fields.push((event_start, depth));
                }
            },
            Event::Empty(ref element) => {
                let local = element.local_name();
                let event_end = usize::try_from(reader.buffer_position()).map_err(|_| {
                    Error::InvalidFormat(
                        "content.xml byte offset exceeds platform limits".to_string(),
                    )
                })?;
                if text_namespace && local.as_ref() == b"p" {
                    if paragraph_count >= MAX_PARAGRAPHS {
                        return Err(Error::InvalidFormat("too many ODF paragraphs".to_string()));
                    }
                    if wanted_paragraph == Some(paragraph_count) {
                        let qualified_name = std::str::from_utf8(element.name().as_ref())
                            .map_err(|_| {
                                Error::InvalidFormat(
                                    "ODF paragraph name is not valid UTF-8".to_string(),
                                )
                            })?
                            .to_string();
                        output.paragraph = Some(ParagraphSite::Empty {
                            start: event_start,
                            end: event_end,
                            qualified_name,
                        });
                    }
                    paragraph_count += 1;
                }
                if text_namespace && is_dynamic_local(local.as_ref()) {
                    if active_fields
                        .last()
                        .is_some_and(|(_, _, allows_nested)| !allows_nested)
                    {
                        return Err(Error::InvalidFormat(
                            "nested dynamic text fields are not allowed".to_string(),
                        ));
                    }
                    if output.fields.len() >= MAX_DYNAMIC_FIELDS {
                        return Err(Error::InvalidFormat(
                            "too many dynamic text fields".to_string(),
                        ));
                    }
                    output.fields.push(Span {
                        start: event_start,
                        end: event_end,
                    });
                }
                if text_namespace && is_database_local(local.as_ref()) {
                    output.database_fields.push(Span {
                        start: event_start,
                        end: event_end,
                    });
                }
            },
            Event::End(_) => {
                if active_paragraph_depth == Some(depth) {
                    output.paragraph = Some(ParagraphSite::BeforeEnd(event_start));
                    active_paragraph_depth = None;
                }
                if active_fields
                    .last()
                    .is_some_and(|(_, field_depth, _)| *field_depth == depth)
                {
                    let (start, _, _) = active_fields.pop().expect("checked active field");
                    let end = usize::try_from(reader.buffer_position()).map_err(|_| {
                        Error::InvalidFormat(
                            "content.xml byte offset exceeds platform limits".to_string(),
                        )
                    })?;
                    output.fields.push(Span { start, end });
                }
                if active_database_fields
                    .last()
                    .is_some_and(|(_, field_depth)| *field_depth == depth)
                {
                    let (start, _) = active_database_fields
                        .pop()
                        .expect("checked database field");
                    let end = usize::try_from(reader.buffer_position()).map_err(|_| {
                        Error::InvalidFormat(
                            "content.xml byte offset exceeds platform limits".to_string(),
                        )
                    })?;
                    output.database_fields.push(Span { start, end });
                }
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("content.xml nesting depth underflow".to_string())
                })?;
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if !active_fields.is_empty() || !active_database_fields.is_empty() {
        return Err(Error::InvalidFormat(
            "unterminated dynamic text field".to_string(),
        ));
    }
    output.fields.sort_by_key(|span| span.start);
    output.database_fields.sort_by_key(|span| span.start);
    Ok(output)
}

fn is_database_local(local: &[u8]) -> bool {
    matches!(
        local,
        b"database-display"
            | b"database-next"
            | b"database-row-select"
            | b"database-row-number"
            | b"database-name"
    )
}

fn is_dynamic_local(local: &[u8]) -> bool {
    matches!(
        local,
        b"placeholder"
            | b"conditional-text"
            | b"hidden-text"
            | b"hidden-paragraph"
            | b"sequence"
            | b"sequence-ref"
            | b"variable-set"
            | b"variable-get"
            | b"expression"
            | b"variable-input"
            | b"user-field-get"
            | b"user-field-input"
            | b"text-input"
            | b"drop-down"
            | b"script"
            | b"table-formula"
            | b"measure"
            | b"reference-ref"
            | b"bookmark-ref"
            | b"note-ref"
            | b"page-count"
            | b"paragraph-count"
            | b"word-count"
            | b"character-count"
            | b"table-count"
            | b"image-count"
            | b"object-count"
            | b"page-number"
            | b"date"
            | b"time"
            | b"page-continuation"
            | b"page-variable-set"
            | b"page-variable-get"
            | b"creation-date"
            | b"creation-time"
            | b"print-date"
            | b"print-time"
            | b"editing-cycles"
            | b"editing-duration"
            | b"modification-date"
            | b"modification-time"
            | b"initial-creator"
            | b"description"
            | b"printed-by"
            | b"title"
            | b"subject"
            | b"keywords"
            | b"creator"
            | b"user-defined"
            | b"meta-field"
            | b"dde-connection"
    )
}

#[cfg(test)]
mod fixed_page_date_time_mutation_tests {
    use super::*;
    use crate::elements::field::{DynamicTextField, FieldDateValue, PageContinuationSelection};

    const XML: &str = r#"<o:document-content
        xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
        xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
        <o:body><o:text><t:p>before<t:date t:date-value="2024-01-01">old</t:date>after</t:p></o:text></o:body>
    </o:document-content>"#;

    #[test]
    fn fixed_page_date_time_mutation_insert_replace_remove_is_namespace_aware() {
        let replacement = DynamicTextField::PageContinuation {
            select_page: PageContinuationSelection::Next,
            string_value: Some("continued".to_string()),
            display_text: "cached".to_string(),
        };
        let replaced = replace_dynamic_text_field_xml(XML, 0, &replacement).unwrap();
        assert!(replaced.contains("before<text:page-continuation"));
        assert!(replaced.contains("</text:page-continuation>after"));

        let inserted = insert_dynamic_text_field_xml(
            &replaced,
            0,
            &DynamicTextField::Date {
                value: Some(FieldDateValue::new("2024-12-31").unwrap()),
                adjustment: None,
                fixed: Some(true),
                data_style_name: None,
                display_text: "year end".to_string(),
            },
        )
        .unwrap();
        assert!(inserted.contains("<text:date"));
        let removed = remove_dynamic_text_field_xml(&inserted, 0).unwrap();
        assert!(!removed.contains("text:page-continuation"));
        assert!(removed.contains("<text:date"));
        assert!(remove_dynamic_text_field_xml(&removed, 1).is_err());
    }
}

#[cfg(test)]
mod page_variable_family_mutation_tests {
    use super::*;
    use crate::elements::field::{DynamicTextField, SequenceNumberFormat};

    const XML: &str = r#"<o:document-content
        xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
        xmlns:q="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
        <o:body><o:text><q:p>before<q:page-variable-set q:active="true"/>after</q:p></o:text></o:body>
    </o:document-content>"#;

    #[test]
    fn page_variable_family_mutation_inserts_replaces_removes_and_checks_bounds() {
        let getter = DynamicTextField::PageVariableGet {
            number_format: Some(SequenceNumberFormat::new("1", None).unwrap()),
            display_text: "12".to_string(),
        };
        let replaced = replace_dynamic_text_field_xml(XML, 0, &getter).unwrap();
        assert!(replaced.contains("before<text:page-variable-get"));
        assert!(replaced.contains("</text:page-variable-get>after"));

        let setter = DynamicTextField::PageVariableSet {
            active: Some(false),
            page_adjust: Some(-4),
            display_text: String::new(),
        };
        let inserted = insert_dynamic_text_field_xml(&replaced, 0, &setter).unwrap();
        assert!(inserted.contains("<text:page-variable-set"));
        let removed = remove_dynamic_text_field_xml(&inserted, 0).unwrap();
        assert!(!removed.contains("text:page-variable-get"));
        assert!(removed.contains("text:page-variable-set"));
        assert!(remove_dynamic_text_field_xml(&removed, 1).is_err());
    }
}

#[cfg(test)]
mod document_metadata_fixed_field_mutation_tests {
    use super::*;
    use crate::elements::field::{
        DynamicTextField, FieldDuration, MetadataFieldKind, MetadataFieldValue,
    };

    const XML: &str = r#"<o:document-content
        xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
        xmlns:m="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
        <o:body><o:text><m:p>before<m:editing-cycles m:fixed="true">3</m:editing-cycles>after</m:p></o:text></o:body>
    </o:document-content>"#;

    #[test]
    fn document_metadata_fixed_field_mutation_is_namespace_aware_and_bounded() {
        let replacement = DynamicTextField::DocumentMetadata {
            kind: MetadataFieldKind::EditingDuration,
            value: Some(MetadataFieldValue::Duration(
                FieldDuration::new("PT3H").unwrap(),
            )),
            fixed: Some(true),
            data_style_name: Some("Elapsed".to_string()),
            display_text: "three hours".to_string(),
        };
        let replaced = replace_dynamic_text_field_xml(XML, 0, &replacement).unwrap();
        assert!(replaced.contains("before<text:editing-duration"));
        assert!(replaced.contains("</text:editing-duration>after"));

        let inserted = insert_dynamic_text_field_xml(
            &replaced,
            0,
            &DynamicTextField::DocumentMetadata {
                kind: MetadataFieldKind::ModificationTime,
                value: None,
                fixed: None,
                data_style_name: None,
                display_text: "now".to_string(),
            },
        )
        .unwrap();
        assert!(inserted.contains("<text:modification-time"));
        let removed = remove_dynamic_text_field_xml(&inserted, 0).unwrap();
        assert!(!removed.contains("text:editing-duration"));
        assert!(removed.contains("text:modification-time"));
        assert!(remove_dynamic_text_field_xml(&removed, 1).is_err());
    }
}

#[cfg(test)]
mod document_identity_fixed_field_mutation_tests {
    use super::*;
    use crate::elements::field::{DynamicTextField, IdentityFieldKind};

    const XML: &str = r#"<o:document-content
        xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
        xmlns:i="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
        <o:body><o:text><i:p>before<i:title i:fixed="true">old</i:title>after</i:p></o:text></o:body>
    </o:document-content>"#;

    #[test]
    fn document_identity_fixed_field_mutation_is_namespace_aware_and_bounded() {
        let replacement = DynamicTextField::DocumentIdentity {
            kind: IdentityFieldKind::Creator,
            fixed: Some(false),
            display_text: "new creator".to_string(),
        };
        let replaced = replace_dynamic_text_field_xml(XML, 0, &replacement).unwrap();
        assert!(replaced.contains("before<text:creator"));
        assert!(replaced.contains("</text:creator>after"));

        let inserted = insert_dynamic_text_field_xml(
            &replaced,
            0,
            &DynamicTextField::DocumentIdentity {
                kind: IdentityFieldKind::Keywords,
                fixed: None,
                display_text: "one, two".to_string(),
            },
        )
        .unwrap();
        assert!(inserted.contains("<text:keywords"));
        let removed = remove_dynamic_text_field_xml(&inserted, 0).unwrap();
        assert!(!removed.contains("text:creator"));
        assert!(removed.contains("text:keywords"));
        assert!(remove_dynamic_text_field_xml(&removed, 1).is_err());
    }
}

#[cfg(test)]
mod user_defined_metadata_field_mutation_tests {
    use super::*;
    use crate::elements::field::{DynamicTextField, UserDefinedMetadataValues};

    const XML: &str = r#"<o:document-content
        xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
        xmlns:u="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
        <o:body><o:text><u:p>before<u:user-defined u:name="old">old</u:user-defined>after</u:p></o:text></o:body>
    </o:document-content>"#;

    fn field(name: &str, display_text: &str) -> DynamicTextField {
        DynamicTextField::UserDefinedMetadata {
            name: name.to_string(),
            values: UserDefinedMetadataValues {
                string: Some(display_text.to_string()),
                ..UserDefinedMetadataValues::default()
            },
            fixed: Some(true),
            data_style_name: None,
            display_text: display_text.to_string(),
        }
    }

    #[test]
    fn user_defined_metadata_field_mutation_is_namespace_aware_and_bounded() {
        let replaced = replace_dynamic_text_field_xml(XML, 0, &field("new", "cached")).unwrap();
        assert!(replaced.contains("before<text:user-defined"));
        assert!(replaced.contains("</text:user-defined>after"));

        let inserted =
            insert_dynamic_text_field_xml(&replaced, 0, &field("second", "two")).unwrap();
        let removed = remove_dynamic_text_field_xml(&inserted, 0).unwrap();
        assert!(!removed.contains("text:name=\"new\""));
        assert!(removed.contains("text:name=\"second\""));
        assert!(remove_dynamic_text_field_xml(&removed, 1).is_err());
    }
}
