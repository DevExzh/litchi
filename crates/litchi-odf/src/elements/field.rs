//! Field elements for ODF documents.
//!
//! Fields are dynamic content in ODF documents that can be updated automatically,
//! such as page numbers, dates, cross-references, etc.

use super::element::{Element, ElementBase};
use super::xml::{
    TEXT_NAMESPACE, append_checked, append_text_control, copy_canonical_attributes,
    decode_reference, is_bound,
};
use litchi_core::{Error, Result};
use quick_xml::XmlVersion;
use quick_xml::events::Event;
use quick_xml::reader::NsReader;
use std::collections::HashMap;

const MAX_FIELD_DEPTH: usize = 4_096;
const MAX_FIELDS: usize = 1_000_000;
const TEXT_DATABASE_NAMESPACE: &str =
    "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const OFFICE_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const FORM_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:form:1.0";
const STYLE_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const XLINK_NAMESPACE: &str = "http://www.w3.org/1999/xlink";
const MAX_DATABASE_VALUE: usize = 65_536;
const MAX_DATABASE_AGGREGATE: usize = 16 * 1_048_576;

/// One of the five OpenDocument database field elements.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum OdfDatabaseFieldKind {
    Display,
    Next,
    RowSelect,
    RowNumber,
    Name,
}

/// Kind of database object selected by a field.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum OdfDatabaseTableType {
    Table,
    Query,
    Command,
}

impl OdfDatabaseTableType {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "table" => Ok(Self::Table),
            "query" => Ok(Self::Query),
            "command" => Ok(Self::Command),
            _ => Err(Error::InvalidFormat(format!(
                "invalid database table type '{value}'"
            ))),
        }
    }
}

/// An inert `form:connection-resource`. The URI is never resolved or opened.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OdfDatabaseConnectionResource {
    pub href: String,
    pub simple_link: bool,
}

/// Common source identity shared by all database fields.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OdfDatabaseSource {
    pub database_name: Option<String>,
    pub table_name: String,
    pub table_type: Option<OdfDatabaseTableType>,
    pub connection_resource: Option<OdfDatabaseConnectionResource>,
}

impl OdfDatabaseSource {
    /// ODF defaults `text:table-type` to `table`.
    pub fn effective_table_type(&self) -> OdfDatabaseTableType {
        self.table_type.unwrap_or(OdfDatabaseTableType::Table)
    }
}

/// Typed, non-executing database field metadata in document order.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OdfDatabaseField {
    pub kind: OdfDatabaseFieldKind,
    pub source: OdfDatabaseSource,
    pub column_name: Option<String>,
    pub condition: Option<String>,
    pub row_number: Option<u64>,
    pub value: Option<u64>,
    pub data_style_name: Option<String>,
    pub number_format: Option<String>,
    pub number_letter_sync: Option<bool>,
    pub display_text: String,
}

struct ActiveDatabaseField {
    depth: usize,
    field: OdfDatabaseField,
    connection_depth: Option<usize>,
}

type DatabaseAttributes = HashMap<(String, String), String>;

/// Represents a text field in the document
#[derive(Debug, Clone)]
pub struct Field {
    element: Element,
}

impl Field {
    /// Create a new field from an element
    pub fn from_element(element: Element) -> Result<Self> {
        let tag = element.tag_name();
        if !Self::is_field_tag(tag) {
            return Err(Error::InvalidFormat(format!(
                "Element {} is not a field",
                tag
            )));
        }
        Ok(Self { element })
    }

    /// Check if a tag name represents a field
    pub fn is_field_tag(tag: &str) -> bool {
        matches!(
            tag,
            "text:page-number"
                | "text:page-count"
                | "text:page-continuation"
                | "text:page-variable-set"
                | "text:page-variable-get"
                | "text:date"
                | "text:time"
                | "text:file-name"
                | "text:template-name"
                | "text:sheet-name"
                | "text:author-name"
                | "text:author-initials"
                | "text:sender-firstname"
                | "text:sender-lastname"
                | "text:sender-initials"
                | "text:sender-title"
                | "text:sender-position"
                | "text:sender-email"
                | "text:sender-phone-private"
                | "text:sender-fax"
                | "text:sender-company"
                | "text:sender-phone-work"
                | "text:sender-street"
                | "text:sender-city"
                | "text:sender-postal-code"
                | "text:sender-country"
                | "text:sender-state-or-province"
                | "text:chapter"
                | "text:title"
                | "text:subject"
                | "text:keywords"
                | "text:description"
                | "text:user-defined"
                | "text:creator"
                | "text:initial-creator"
                | "text:printed-by"
                | "text:creation-date"
                | "text:creation-time"
                | "text:modification-date"
                | "text:modification-time"
                | "text:print-date"
                | "text:print-time"
                | "text:editing-duration"
                | "text:editing-cycles"
                | "text:reference-ref"
                | "text:sequence-ref"
                | "text:bookmark-ref"
                | "text:note-ref"
                | "text:variable-set"
                | "text:variable-get"
                | "text:variable-input"
                | "text:user-field-get"
                | "text:user-field-input"
                | "text:sequence"
                | "text:expression"
                | "text:text-input"
                | "text:placeholder"
                | "text:conditional-text"
                | "text:hidden-text"
                | "text:hidden-paragraph"
                | "text:measure"
                | "text:table-formula"
                | "text:database-display"
                | "text:database-next"
                | "text:database-row-select"
                | "text:database-row-number"
                | "text:database-name"
                | "text:word-count"
                | "text:paragraph-count"
                | "text:character-count"
                | "text:table-count"
                | "text:image-count"
                | "text:object-count"
        )
    }

    /// Get the field type
    pub fn field_type(&self) -> &str {
        self.element.tag_name()
    }

    /// Get the field value (text content)
    pub fn value(&self) -> String {
        self.element.get_text_recursive()
    }

    /// Get the field display format
    pub fn format(&self) -> Option<&str> {
        self.element
            .get_attribute("style:data-style-name")
            .or_else(|| self.element.get_attribute("number:style"))
    }

    /// Get the field name (for named fields like variables or user fields)
    pub fn name(&self) -> Option<&str> {
        self.element
            .get_attribute("text:name")
            .or_else(|| self.element.get_attribute("text:variable-name"))
    }

    /// Get reference target (for reference fields)
    pub fn reference_target(&self) -> Option<&str> {
        self.element
            .get_attribute("text:ref-name")
            .or_else(|| self.element.get_attribute("text:reference-name"))
    }
}

/// Represents a page number field
#[derive(Debug, Clone)]
#[allow(dead_code)] // Library API for document creation
pub struct PageNumberField {
    element: Element,
}

#[allow(dead_code)] // Library API for document creation
impl PageNumberField {
    /// Create a new page number field
    pub fn new() -> Self {
        Self {
            element: Element::new("text:page-number"),
        }
    }

    /// Create from element
    pub fn from_element(element: Element) -> Result<Self> {
        if element.tag_name() != "text:page-number" {
            return Err(Error::InvalidFormat(
                "Element is not a page number field".to_string(),
            ));
        }
        Ok(Self { element })
    }

    /// Get the current page number value
    pub fn value(&self) -> String {
        self.element.get_text_recursive()
    }
}

impl Default for PageNumberField {
    fn default() -> Self {
        Self::new()
    }
}

/// Represents a date field
#[derive(Debug, Clone)]
#[allow(dead_code)] // Library API for document creation
pub struct DateField {
    element: Element,
}

#[allow(dead_code)] // Library API for document creation
impl DateField {
    /// Create a new date field
    pub fn new() -> Self {
        Self {
            element: Element::new("text:date"),
        }
    }

    /// Create from element
    pub fn from_element(element: Element) -> Result<Self> {
        if element.tag_name() != "text:date" {
            return Err(Error::InvalidFormat(
                "Element is not a date field".to_string(),
            ));
        }
        Ok(Self { element })
    }

    /// Get the date value
    pub fn value(&self) -> String {
        self.element.get_text_recursive()
    }

    /// Get the fixed date (if any)
    pub fn fixed_date(&self) -> Option<&str> {
        self.element.get_attribute("text:date-value")
    }

    /// Get whether this date is fixed
    pub fn is_fixed(&self) -> bool {
        self.element
            .get_bool_attribute("text:fixed")
            .unwrap_or(false)
    }
}

impl Default for DateField {
    fn default() -> Self {
        Self::new()
    }
}

/// Represents a reference field
#[derive(Debug, Clone)]
#[allow(dead_code)] // Library API for document creation
pub struct ReferenceField {
    element: Element,
}

#[allow(dead_code)] // Library API for document creation
impl ReferenceField {
    /// Create a new reference field
    pub fn new(ref_name: &str) -> Self {
        let mut element = Element::new("text:reference-ref");
        element.set_attribute("text:ref-name", ref_name);
        Self { element }
    }

    /// Create from element
    pub fn from_element(element: Element) -> Result<Self> {
        let tag = element.tag_name();
        if !matches!(
            tag,
            "text:reference-ref" | "text:bookmark-ref" | "text:sequence-ref"
        ) {
            return Err(Error::InvalidFormat(format!(
                "Element {} is not a reference field",
                tag
            )));
        }
        Ok(Self { element })
    }

    /// Get the reference name
    pub fn ref_name(&self) -> Option<&str> {
        self.element.get_attribute("text:ref-name")
    }

    /// Get the reference format
    pub fn ref_format(&self) -> Option<&str> {
        self.element.get_attribute("text:reference-format")
    }

    /// Get the reference value
    pub fn value(&self) -> String {
        self.element.get_text_recursive()
    }
}

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
                    document_depth = checked_field_depth(document_depth)?;
                    for field in &mut active {
                        field.depth += 1;
                    }
                    if text_element {
                        for field in &mut active {
                            append_text_control(&reader, source, &mut field.text)?;
                        }
                        let tag_name = format!(
                            "text:{}",
                            std::str::from_utf8(source.local_name().as_ref()).map_err(|_| {
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
                Event::Empty(ref source) if text_element => {
                    for field in &mut active {
                        append_text_control(&reader, source, &mut field.text)?;
                    }
                    let tag_name = format!(
                        "text:{}",
                        std::str::from_utf8(source.local_name().as_ref()).map_err(|_| {
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
                    if active.last().is_some_and(|field| field.depth == 0) {
                        let mut field = active.pop().expect("checked active field");
                        field.element.set_text(&field.text);
                        fields.push((field.order, Field::from_element(field.element)?));
                    }
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
    pub fn parse_database_fields(xml_content: &str) -> Result<Vec<OdfDatabaseField>> {
        parse_database_fields(xml_content)
    }
}

fn parse_database_fields(xml: &str) -> Result<Vec<OdfDatabaseField>> {
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut active: Option<ActiveDatabaseField> = None;
    let mut fields = Vec::new();
    let mut aggregate = 0usize;

    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid database field XML: {error}")))?;
        let namespace_uri = resolved_namespace(&namespace)?;
        match event {
            Event::Start(ref element) => {
                let local = utf8(element.local_name().as_ref(), "database field element")?;
                reject_spoofed_database_name(namespace_uri.as_deref(), &local)?;
                if let Some(field) = active.as_mut() {
                    if namespace_uri.as_deref() != Some(FORM_NAMESPACE)
                        || local != "connection-resource"
                        || depth != field.depth
                        || field.connection_depth.is_some()
                        || field.field.source.connection_resource.is_some()
                    {
                        return Err(Error::InvalidFormat(
                            "database fields may contain only one form:connection-resource"
                                .to_string(),
                        ));
                    }
                    field.field.source.connection_resource = Some(parse_connection_resource(
                        &reader,
                        element,
                        &mut aggregate,
                    )?);
                    field.connection_depth = Some(depth + 1);
                } else if namespace_uri.as_deref() == Some(TEXT_DATABASE_NAMESPACE) {
                    if let Some(kind) = database_field_kind(&local) {
                        if fields.len() >= MAX_FIELDS {
                            return Err(Error::InvalidFormat(format!(
                                "document exceeds {MAX_FIELDS} database fields"
                            )));
                        }
                        active = Some(ActiveDatabaseField {
                            depth: depth + 1,
                            field: parse_database_field(
                                &reader,
                                element,
                                kind,
                                &mut aggregate,
                            )?,
                            connection_depth: None,
                        });
                    }
                }
                depth = checked_field_depth(depth)?;
            }
            Event::Empty(ref element) => {
                let local = utf8(element.local_name().as_ref(), "database field element")?;
                reject_spoofed_database_name(namespace_uri.as_deref(), &local)?;
                if let Some(field) = active.as_mut() {
                    if namespace_uri.as_deref() != Some(FORM_NAMESPACE)
                        || local != "connection-resource"
                        || depth != field.depth
                        || field.field.source.connection_resource.is_some()
                    {
                        return Err(Error::InvalidFormat(
                            "database fields may contain only one form:connection-resource"
                                .to_string(),
                        ));
                    }
                    field.field.source.connection_resource = Some(parse_connection_resource(
                        &reader,
                        element,
                        &mut aggregate,
                    )?);
                } else if namespace_uri.as_deref() == Some(TEXT_DATABASE_NAMESPACE) {
                    if let Some(kind) = database_field_kind(&local) {
                        if fields.len() >= MAX_FIELDS {
                            return Err(Error::InvalidFormat(format!(
                                "document exceeds {MAX_FIELDS} database fields"
                            )));
                        }
                        let field = parse_database_field(
                            &reader,
                            element,
                            kind,
                            &mut aggregate,
                        )?;
                        fields.push(validate_database_field(field)?);
                    }
                }
            }
            Event::End(_) => {
                if let Some(field) = active.as_mut() {
                    if field
                        .connection_depth
                        .is_some_and(|connection_depth| connection_depth == depth)
                    {
                        field.connection_depth = None;
                    }
                }
                if active.as_ref().is_some_and(|field| field.depth == depth) {
                    let field = active.take().expect("checked database field").field;
                    fields.push(validate_database_field(field)?);
                }
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("database field XML depth underflow".to_string())
                })?;
            }
            Event::Text(ref text) if active.is_some() => {
                let value = text.decode().map_err(|error| {
                    Error::InvalidFormat(format!("invalid database field text: {error}"))
                })?;
                append_database_text(active.as_mut().expect("checked field"), &value, &mut aggregate)?;
            }
            Event::CData(ref text) if active.is_some() => {
                let value = text.decode().map_err(|error| {
                    Error::InvalidFormat(format!("invalid database field CDATA: {error}"))
                })?;
                append_database_text(active.as_mut().expect("checked field"), &value, &mut aggregate)?;
            }
            Event::GeneralRef(ref reference) if active.is_some() => {
                let name = std::str::from_utf8(reference.as_ref()).map_err(|_| {
                    Error::InvalidFormat("invalid database field entity reference".to_string())
                })?;
                let value = resolve_database_reference(name)?;
                append_database_text(active.as_mut().expect("checked field"), &value, &mut aggregate)?;
            }
            Event::DocType(_) | Event::PI(_) => {
                return Err(Error::InvalidFormat(
                    "DTDs and processing instructions are prohibited in database field XML"
                        .to_string(),
                ));
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    if depth != 0 || active.is_some() {
        return Err(Error::InvalidFormat(
            "incomplete database field XML structure".to_string(),
        ));
    }
    Ok(fields)
}

fn parse_database_field(
    reader: &NsReader<&[u8]>,
    element: &quick_xml::events::BytesStart<'_>,
    kind: OdfDatabaseFieldKind,
    aggregate: &mut usize,
) -> Result<OdfDatabaseField> {
    let attributes = database_attributes(reader, element, aggregate)?;
    let allowed = match kind {
        OdfDatabaseFieldKind::Display => &[
            (TEXT_DATABASE_NAMESPACE, "database-name"),
            (TEXT_DATABASE_NAMESPACE, "table-name"),
            (TEXT_DATABASE_NAMESPACE, "table-type"),
            (TEXT_DATABASE_NAMESPACE, "column-name"),
            (STYLE_NAMESPACE, "data-style-name"),
        ][..],
        OdfDatabaseFieldKind::Next => &[
            (TEXT_DATABASE_NAMESPACE, "database-name"),
            (TEXT_DATABASE_NAMESPACE, "table-name"),
            (TEXT_DATABASE_NAMESPACE, "table-type"),
            (TEXT_DATABASE_NAMESPACE, "condition"),
        ][..],
        OdfDatabaseFieldKind::RowSelect => &[
            (TEXT_DATABASE_NAMESPACE, "database-name"),
            (TEXT_DATABASE_NAMESPACE, "table-name"),
            (TEXT_DATABASE_NAMESPACE, "table-type"),
            (TEXT_DATABASE_NAMESPACE, "condition"),
            (TEXT_DATABASE_NAMESPACE, "row-number"),
        ][..],
        OdfDatabaseFieldKind::RowNumber => &[
            (TEXT_DATABASE_NAMESPACE, "database-name"),
            (TEXT_DATABASE_NAMESPACE, "table-name"),
            (TEXT_DATABASE_NAMESPACE, "table-type"),
            (TEXT_DATABASE_NAMESPACE, "value"),
            (STYLE_NAMESPACE, "num-format"),
            (STYLE_NAMESPACE, "num-letter-sync"),
        ][..],
        OdfDatabaseFieldKind::Name => &[
            (TEXT_DATABASE_NAMESPACE, "database-name"),
            (TEXT_DATABASE_NAMESPACE, "table-name"),
            (TEXT_DATABASE_NAMESPACE, "table-type"),
        ][..],
    };
    reject_database_attributes(&attributes, allowed)?;
    let table_name =
        required_database_attribute(&attributes, TEXT_DATABASE_NAMESPACE, "table-name")?;
    let table_type = database_attribute(&attributes, TEXT_DATABASE_NAMESPACE, "table-type")
        .map(OdfDatabaseTableType::parse)
        .transpose()?;
    let row_number = database_attribute(&attributes, TEXT_DATABASE_NAMESPACE, "row-number")
        .map(|value| parse_database_u64(value, "row-number"))
        .transpose()?;
    let value = database_attribute(&attributes, TEXT_DATABASE_NAMESPACE, "value")
        .map(|value| parse_database_u64(value, "value"))
        .transpose()?;
    let number_letter_sync = database_attribute(&attributes, STYLE_NAMESPACE, "num-letter-sync")
        .map(parse_database_bool)
        .transpose()?;
    Ok(OdfDatabaseField {
        kind,
        source: OdfDatabaseSource {
            database_name: database_attribute(
                &attributes,
                TEXT_DATABASE_NAMESPACE,
                "database-name",
            )
                .map(str::to_string),
            table_name,
            table_type,
            connection_resource: None,
        },
        column_name: database_attribute(&attributes, TEXT_DATABASE_NAMESPACE, "column-name")
            .map(str::to_string),
        condition: database_attribute(&attributes, TEXT_DATABASE_NAMESPACE, "condition")
            .map(str::to_string),
        row_number,
        value,
        data_style_name: database_attribute(&attributes, STYLE_NAMESPACE, "data-style-name")
            .map(str::to_string),
        number_format: database_attribute(&attributes, STYLE_NAMESPACE, "num-format")
            .map(str::to_string),
        number_letter_sync,
        display_text: String::new(),
    })
}

fn validate_database_field(field: OdfDatabaseField) -> Result<OdfDatabaseField> {
    if field.source.database_name.as_deref().is_none_or(str::is_empty)
        && field.source.connection_resource.is_none()
    {
        return Err(Error::InvalidFormat(
            "database field requires text:database-name or form:connection-resource".to_string(),
        ));
    }
    if field.source.table_name.is_empty() {
        return Err(Error::InvalidFormat(
            "database field requires non-empty text:table-name".to_string(),
        ));
    }
    match field.kind {
        OdfDatabaseFieldKind::Display
            if field.column_name.as_deref().is_none_or(str::is_empty) =>
        {
            return Err(Error::InvalidFormat(
                "text:database-display requires text:column-name".to_string(),
            ));
        }
        OdfDatabaseFieldKind::RowSelect if field.row_number.is_none() => {
            return Err(Error::InvalidFormat(
                "text:database-row-select requires text:row-number".to_string(),
            ));
        }
        OdfDatabaseFieldKind::Next | OdfDatabaseFieldKind::RowSelect
            if !field.display_text.trim().is_empty() =>
        {
            return Err(Error::InvalidFormat(
                "database selection fields cannot contain character data".to_string(),
            ));
        }
        _ => {}
    }
    Ok(field)
}

fn parse_connection_resource(
    reader: &NsReader<&[u8]>,
    element: &quick_xml::events::BytesStart<'_>,
    aggregate: &mut usize,
) -> Result<OdfDatabaseConnectionResource> {
    let attributes = database_attributes(reader, element, aggregate)?;
    reject_database_attributes(
        &attributes,
        &[(XLINK_NAMESPACE, "href"), (XLINK_NAMESPACE, "type")],
    )?;
    let href = required_database_attribute(&attributes, XLINK_NAMESPACE, "href")?;
    if href.is_empty() {
        return Err(Error::InvalidFormat(
            "form:connection-resource requires non-empty xlink:href".to_string(),
        ));
    }
    let simple_link = match database_attribute(&attributes, XLINK_NAMESPACE, "type") {
        None | Some("simple") => true,
        Some(value) => {
            return Err(Error::InvalidFormat(format!(
                "unsupported connection-resource xlink:type '{value}'"
            )));
        }
    };
    Ok(OdfDatabaseConnectionResource { href, simple_link })
}

fn database_attributes(
    reader: &NsReader<&[u8]>,
    element: &quick_xml::events::BytesStart<'_>,
    aggregate: &mut usize,
) -> Result<DatabaseAttributes> {
    let mut attributes = HashMap::new();
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| {
            Error::InvalidFormat(format!("invalid database field attribute: {error}"))
        })?;
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let namespace = resolved_namespace(&namespace)?.unwrap_or_default();
        let local = utf8(local.as_ref(), "database field attribute")?;
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid database field attribute value: {error}"))
            })?
            .into_owned();
        append_database_size(aggregate, value.len())?;
        if attributes.insert((namespace, local), value).is_some() {
            return Err(Error::InvalidFormat(
                "duplicate expanded database field attribute".to_string(),
            ));
        }
    }
    Ok(attributes)
}

fn reject_database_attributes(
    attributes: &DatabaseAttributes,
    allowed: &[(&str, &str)],
) -> Result<()> {
    for (namespace, local) in attributes.keys() {
        if !allowed.iter().any(|(allowed_namespace, allowed_local)| {
            namespace == allowed_namespace && local == allowed_local
        }) && matches!(
            namespace.as_str(),
            TEXT_DATABASE_NAMESPACE
                | FORM_NAMESPACE
                | STYLE_NAMESPACE
                | XLINK_NAMESPACE
                | OFFICE_NAMESPACE
        ) {
            return Err(Error::InvalidFormat(format!(
                "unexpected database field attribute {namespace}:{local}"
            )));
        }
    }
    Ok(())
}

fn append_database_text(
    active: &mut ActiveDatabaseField,
    value: &str,
    aggregate: &mut usize,
) -> Result<()> {
    if active.connection_depth.is_some() {
        if value.is_empty() {
            return Ok(());
        }
        return Err(Error::InvalidFormat(
            "form:connection-resource must be empty".to_string(),
        ));
    }
    if active
        .field
        .display_text
        .len()
        .saturating_add(value.len())
        > MAX_DATABASE_VALUE
    {
        return Err(Error::InvalidFormat(
            "ODF database field display text exceeds the supported limit".to_string(),
        ));
    }
    append_database_size(aggregate, value.len())?;
    active.field.display_text.push_str(value);
    Ok(())
}

fn append_database_size(aggregate: &mut usize, amount: usize) -> Result<()> {
    if amount > MAX_DATABASE_VALUE {
        return Err(Error::InvalidFormat(
            "database field value exceeds 64 KiB".to_string(),
        ));
    }
    *aggregate = aggregate.checked_add(amount).ok_or_else(|| {
        Error::InvalidFormat("database field aggregate size overflow".to_string())
    })?;
    if *aggregate > MAX_DATABASE_AGGREGATE {
        return Err(Error::InvalidFormat(
            "database field metadata exceeds 16 MiB".to_string(),
        ));
    }
    Ok(())
}

fn database_field_kind(local: &str) -> Option<OdfDatabaseFieldKind> {
    match local {
        "database-display" => Some(OdfDatabaseFieldKind::Display),
        "database-next" => Some(OdfDatabaseFieldKind::Next),
        "database-row-select" => Some(OdfDatabaseFieldKind::RowSelect),
        "database-row-number" => Some(OdfDatabaseFieldKind::RowNumber),
        "database-name" => Some(OdfDatabaseFieldKind::Name),
        _ => None,
    }
}

fn database_attribute<'a>(
    attributes: &'a DatabaseAttributes,
    namespace: &str,
    local: &str,
) -> Option<&'a str> {
    attributes
        .get(&(namespace.to_string(), local.to_string()))
        .map(String::as_str)
}

fn required_database_attribute(
    attributes: &DatabaseAttributes,
    namespace: &str,
    local: &str,
) -> Result<String> {
    database_attribute(attributes, namespace, local)
        .map(str::to_string)
        .ok_or_else(|| Error::InvalidFormat(format!("database field requires {local}")))
}

fn parse_database_u64(value: &str, name: &str) -> Result<u64> {
    value
        .parse::<u64>()
        .map_err(|_| Error::InvalidFormat(format!("invalid database field {name} '{value}'")))
}

fn parse_database_bool(value: &str) -> Result<bool> {
    match value {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        _ => Err(Error::InvalidFormat(format!(
            "invalid database field boolean '{value}'"
        ))),
    }
}

fn reject_spoofed_database_name(namespace: Option<&str>, local: &str) -> Result<()> {
    if (database_field_kind(local).is_some() && namespace != Some(TEXT_DATABASE_NAMESPACE))
        || (local == "connection-resource" && namespace != Some(FORM_NAMESPACE))
    {
        return Err(Error::InvalidFormat(
            "database field vocabulary uses the wrong namespace".to_string(),
        ));
    }
    Ok(())
}

fn resolved_namespace(namespace: &quick_xml::name::ResolveResult<'_>) -> Result<Option<String>> {
    match namespace {
        quick_xml::name::ResolveResult::Bound(quick_xml::name::Namespace(value)) => {
            Ok(Some(utf8(value, "namespace URI")?))
        }
        quick_xml::name::ResolveResult::Unbound => Ok(None),
        quick_xml::name::ResolveResult::Unknown(prefix) => Err(Error::InvalidFormat(format!(
            "unbound namespace prefix '{}'",
            String::from_utf8_lossy(prefix)
        ))),
    }
}

fn utf8(value: &[u8], description: &str) -> Result<String> {
    std::str::from_utf8(value)
        .map(str::to_string)
        .map_err(|_| Error::InvalidFormat(format!("invalid UTF-8 {description}")))
}

fn resolve_database_reference(name: &str) -> Result<String> {
    if let Some(value) = quick_xml::escape::resolve_xml_entity(name) {
        return Ok(value.to_string());
    }
    let codepoint = if let Some(value) = name
        .strip_prefix("#x")
        .or_else(|| name.strip_prefix("#X"))
    {
        u32::from_str_radix(value, 16)
    } else if let Some(value) = name.strip_prefix('#') {
        value.parse::<u32>()
    } else {
        return Err(Error::InvalidFormat(
            "undeclared entity in database field".to_string(),
        ));
    }
    .map_err(|_| Error::InvalidFormat("invalid database field character reference".to_string()))?;
    char::from_u32(codepoint)
        .filter(|value| {
            matches!(*value, '\u{9}' | '\u{a}' | '\u{d}')
                || matches!(*value as u32, 0x20..=0xd7ff | 0xe000..=0xfffd | 0x10000..=0x10ffff)
        })
        .map(|value| value.to_string())
        .ok_or_else(|| Error::InvalidFormat("invalid XML character reference".to_string()))
}

#[cfg(test)]
mod database_field_tests {
    use super::*;

    const PREFIX: &str = r#"<o:document-content
        xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
        xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
        xmlns:f="urn:oasis:names:tc:opendocument:xmlns:form:1.0"
        xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
        xmlns:x="http://www.w3.org/1999/xlink"><o:body><o:text><t:p>"#;
    const SUFFIX: &str = "</t:p></o:text></o:body></o:document-content>";

    #[test]
    fn parses_all_database_field_kinds_without_resolving_resources() {
        let xml = format!(
            r#"{PREFIX}<t:database-display t:database-name="Contacts" t:table-name="People"
                t:table-type="query" t:column-name="FullName" s:data-style-name="N1">A&amp;B</t:database-display>
            <t:database-next t:database-name="Contacts" t:table-name="People" t:condition="of:=TRUE()"/>
            <t:database-row-select t:table-name="People" t:row-number="42">
                <f:connection-resource x:href="sdbc:embedded:firebird" x:type="simple"/>
            </t:database-row-select>
            <t:database-row-number t:database-name="Contacts" t:table-name="People"
                t:value="42" s:num-format="1" s:num-letter-sync="false">42</t:database-row-number>
            <t:database-name t:database-name="Contacts" t:table-name="People">Contacts</t:database-name>{SUFFIX}"#
        );
        let fields = FieldParser::parse_database_fields(&xml).unwrap();
        assert_eq!(fields.len(), 5);
        assert_eq!(fields[0].kind, OdfDatabaseFieldKind::Display);
        assert_eq!(fields[0].display_text, "A&B");
        assert_eq!(fields[0].source.effective_table_type(), OdfDatabaseTableType::Query);
        assert_eq!(fields[2].row_number, Some(42));
        assert_eq!(
            fields[2]
                .source
                .connection_resource
                .as_ref()
                .map(|resource| resource.href.as_str()),
            Some("sdbc:embedded:firebird")
        );
        assert_eq!(fields[3].number_letter_sync, Some(false));
    }

    #[test]
    fn rejects_missing_invalid_nested_and_active_database_fields() {
        let bodies = [
            r#"<t:database-display t:database-name="db" t:table-name="t"/>"#,
            r#"<t:database-next t:table-name="t"/>"#,
            r#"<t:database-row-select t:database-name="db" t:table-name="t"/>"#,
            r#"<t:database-row-select t:database-name="db" t:table-name="t" t:row-number="-1"/>"#,
            r#"<t:database-name t:database-name="db" t:table-name="t" t:table-type="view"/>"#,
            r#"<t:database-next t:database-name="db" t:table-name="t">text</t:database-next>"#,
            r#"<t:database-name t:database-name="db" t:table-name="t"><t:span>x</t:span></t:database-name>"#,
            r#"<t:database-name t:table-name="t"><f:connection-resource x:href="https://example.invalid/db"/><f:connection-resource x:href="other"/></t:database-name>"#,
        ];
        for body in bodies {
            let xml = format!("{PREFIX}{body}{SUFFIX}");
            assert!(
                FieldParser::parse_database_fields(&xml).is_err(),
                "accepted {body}"
            );
        }
    }
}

struct ActiveField {
    element: Element,
    text: String,
    depth: usize,
    order: usize,
}

fn checked_field_depth(depth: usize) -> Result<usize> {
    let depth = depth
        .checked_add(1)
        .ok_or_else(|| Error::InvalidFormat("field nesting depth overflow".to_string()))?;
    if depth > MAX_FIELD_DEPTH {
        return Err(Error::InvalidFormat(format!(
            "field nesting exceeds {MAX_FIELD_DEPTH} levels"
        )));
    }
    Ok(depth)
}
