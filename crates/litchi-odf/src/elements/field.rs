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
use std::collections::{HashMap, HashSet};

const MAX_FIELD_DEPTH: usize = 4_096;
const MAX_FIELDS: usize = 1_000_000;
const TEXT_DATABASE_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const OFFICE_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const FORM_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:form:1.0";
const STYLE_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const XLINK_NAMESPACE: &str = "http://www.w3.org/1999/xlink";
const MAX_DATABASE_VALUE: usize = 65_536;
const MAX_DATABASE_AGGREGATE: usize = 16 * 1_048_576;
const MAX_DATABASE_INTEGER_DIGITS: usize = 4_096;
const MAX_DYNAMIC_FIELD_VALUE: usize = 65_536;
const MAX_DYNAMIC_FIELD_AGGREGATE: usize = 1_048_576;
const MAX_META_FIELD_XML_BYTES: usize = 64 * 1_048_576;
const MAX_META_FIELD_DEPTH: usize = 256;
const MAX_META_FIELD_NODES: usize = 100_000;
const MAX_META_FIELD_ATTRIBUTES: usize = 256;
const XML_NAMESPACE: &str = "http://www.w3.org/XML/1998/namespace";
const DRAW_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const TABLE_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const PRESENTATION_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0";
const SVG_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
const FO_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
const NUMBER_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0";
const META_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:meta:1.0";
const DC_NAMESPACE: &str = "http://purl.org/dc/elements/1.1/";
const XHTML_NAMESPACE: &str = "http://www.w3.org/1999/xhtml";
const DR3D_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0";
const SCRIPT_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:script:1.0";

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

    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Table => "table",
            Self::Query => "query",
            Self::Command => "command",
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

/// Canonical, bounded XML Schema `nonNegativeInteger` without arithmetic semantics.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct OdfNonNegativeInteger(String);

impl OdfNonNegativeInteger {
    pub fn new(lexical: &str) -> Result<Self> {
        let lexical = lexical.trim_matches(|ch| matches!(ch, ' ' | '\t' | '\n' | '\r'));
        let (negative, digits) = match lexical.as_bytes().first() {
            Some(b'+') => (false, &lexical[1..]),
            Some(b'-') => (true, &lexical[1..]),
            _ => (false, lexical),
        };
        if digits.is_empty() || !digits.bytes().all(|byte| byte.is_ascii_digit()) {
            return Err(Error::InvalidFormat(format!(
                "invalid XML Schema nonNegativeInteger '{lexical}'"
            )));
        }
        if digits.len() > MAX_DATABASE_INTEGER_DIGITS {
            return Err(Error::InvalidFormat(format!(
                "nonNegativeInteger exceeds {MAX_DATABASE_INTEGER_DIGITS} digits"
            )));
        }
        if negative && digits.bytes().any(|byte| byte != b'0') {
            return Err(Error::InvalidFormat(format!(
                "negative value is not a nonNegativeInteger '{lexical}'"
            )));
        }
        let canonical = digits.trim_start_matches('0');
        Ok(Self(
            if canonical.is_empty() { "0" } else { canonical }.to_string(),
        ))
    }

    pub fn as_str(&self) -> &str {
        &self.0
    }
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
    pub row_number: Option<OdfNonNegativeInteger>,
    pub value: Option<OdfNonNegativeInteger>,
    pub data_style_name: Option<String>,
    pub number_format: Option<String>,
    pub number_letter_sync: Option<bool>,
    pub display_text: String,
}

impl OdfDatabaseField {
    pub fn to_xml_fragment(&self) -> Result<String> {
        let field = validate_database_field(self.clone())?;
        validate_constructed_database_field(&field)?;
        let local = match field.kind {
            OdfDatabaseFieldKind::Display => "database-display",
            OdfDatabaseFieldKind::Next => "database-next",
            OdfDatabaseFieldKind::RowSelect => "database-row-select",
            OdfDatabaseFieldKind::RowNumber => "database-row-number",
            OdfDatabaseFieldKind::Name => "database-name",
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
            OdfDatabaseFieldKind::Display => {
                attribute(
                    "text",
                    "column-name",
                    field.column_name.as_deref().expect("validated"),
                );
                if let Some(value) = field.data_style_name.as_deref() {
                    attribute("style", "data-style-name", value);
                }
            },
            OdfDatabaseFieldKind::Next => {
                if let Some(value) = field.condition.as_deref() {
                    attribute("text", "condition", value);
                }
            },
            OdfDatabaseFieldKind::RowSelect => {
                if let Some(value) = field.condition.as_deref() {
                    attribute("text", "condition", value);
                }
                if let Some(value) = field.row_number {
                    attribute("text", "row-number", value.as_str());
                }
            },
            OdfDatabaseFieldKind::RowNumber => {
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
            OdfDatabaseFieldKind::Name => {},
        }
        drop(attribute);
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

/// The content category requested by a `text:placeholder` field.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum OdfPlaceholderType {
    Text,
    Table,
    TextBox,
    Image,
    Object,
}

/// Numbering metadata for an ODF `text:sequence` field.
///
/// ODF permits `style:num-letter-sync` only for alphabetic formats (`a` and
/// `A`). Other format strings, including producer-defined values and the empty
/// format, remain opaque and are preserved verbatim.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct OdfSequenceNumberFormat {
    format: String,
    letter_sync: Option<bool>,
}

/// Common numbering metadata used by document statistic fields.
pub type OdfStatisticNumberFormat = OdfSequenceNumberFormat;

/// Numbering metadata used by `text:page-number`.
pub type OdfPageNumberFormat = OdfSequenceNumberFormat;

/// Numbering metadata used by `text:page-variable-get`.
pub type OdfPageVariableNumberFormat = OdfSequenceNumberFormat;

/// Page selected by an ODF page-number field.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum OdfPageSelection {
    Previous,
    Current,
    Next,
}

impl OdfPageSelection {
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Previous => "previous",
            Self::Current => "current",
            Self::Next => "next",
        }
    }

    fn parse(value: &str) -> Result<Self> {
        match value {
            "previous" => Ok(Self::Previous),
            "current" => Ok(Self::Current),
            "next" => Ok(Self::Next),
            _ => Err(Error::InvalidFormat(format!(
                "invalid text:select-page value '{value}'"
            ))),
        }
    }
}

/// Page selected by `text:page-continuation`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum OdfPageContinuationSelection {
    Previous,
    Next,
}

impl OdfPageContinuationSelection {
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Previous => "previous",
            Self::Next => "next",
        }
    }

    fn parse(value: &str) -> Result<Self> {
        match value {
            "previous" => Ok(Self::Previous),
            "next" => Ok(Self::Next),
            _ => Err(Error::InvalidFormat(format!(
                "invalid page-continuation text:select-page '{value}'"
            ))),
        }
    }
}

/// Lexical category retained by a typed ODF date value.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum OdfDateValueKind {
    Date,
    DateTime,
}

/// A validated XML Schema `dateOrDateTime` value for `text:date-value`.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct OdfFieldDateValue {
    lexical: String,
    kind: OdfDateValueKind,
}

impl OdfFieldDateValue {
    pub fn new(value: impl Into<String>) -> Result<Self> {
        let lexical = value.into();
        let kind = if lexical.contains('T') {
            OdfDateValueKind::DateTime
        } else {
            OdfDateValueKind::Date
        };
        let value = Self { lexical, kind };
        let mut aggregate = 0usize;
        value.validate(&mut aggregate)?;
        Ok(value)
    }

    pub fn as_str(&self) -> &str {
        &self.lexical
    }

    pub const fn kind(&self) -> OdfDateValueKind {
        self.kind
    }

    fn validate(&self, aggregate: &mut usize) -> Result<()> {
        validate_dynamic_value("text:date-value", Some(&self.lexical), true, aggregate)?;
        match self.kind {
            OdfDateValueKind::Date => validate_xml_schema_date(&self.lexical),
            OdfDateValueKind::DateTime => validate_xml_schema_date_time(&self.lexical),
        }
    }
}

/// Lexical category retained by a typed ODF time value.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum OdfTimeValueKind {
    Time,
    DateTime,
}

/// A validated XML Schema `timeOrDateTime` value for `text:time-value`.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct OdfFieldTimeValue {
    lexical: String,
    kind: OdfTimeValueKind,
}

impl OdfFieldTimeValue {
    pub fn new(value: impl Into<String>) -> Result<Self> {
        let lexical = value.into();
        let kind = if lexical.contains('T') {
            OdfTimeValueKind::DateTime
        } else {
            OdfTimeValueKind::Time
        };
        let value = Self { lexical, kind };
        let mut aggregate = 0usize;
        value.validate(&mut aggregate)?;
        Ok(value)
    }

    pub fn as_str(&self) -> &str {
        &self.lexical
    }

    pub const fn kind(&self) -> OdfTimeValueKind {
        self.kind
    }

    fn validate(&self, aggregate: &mut usize) -> Result<()> {
        validate_dynamic_value("text:time-value", Some(&self.lexical), true, aggregate)?;
        match self.kind {
            OdfTimeValueKind::Time => validate_xml_schema_time(&self.lexical),
            OdfTimeValueKind::DateTime => validate_xml_schema_date_time(&self.lexical),
        }
    }
}

/// A validated, exactly retained XML Schema duration used for field adjustment.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct OdfFieldDuration(String);

impl OdfFieldDuration {
    pub fn new(value: impl Into<String>) -> Result<Self> {
        let value = Self(value.into());
        let mut aggregate = 0usize;
        value.validate("field adjustment", &mut aggregate)?;
        Ok(value)
    }

    pub fn as_str(&self) -> &str {
        &self.0
    }

    fn validate(&self, name: &str, aggregate: &mut usize) -> Result<()> {
        validate_dynamic_value(name, Some(&self.0), true, aggregate)?;
        crate::datatype::DurationOdf::decode_exact(&self.0).map_err(|_| {
            Error::InvalidFormat(format!("invalid XML Schema duration '{}'", self.0))
        })?;
        Ok(())
    }
}

/// Display format for a `text:sequence-ref` field.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum OdfSequenceReferenceFormat {
    Page,
    Chapter,
    Direction,
    Text,
    CategoryAndValue,
    Caption,
    Value,
}

/// Display mode permitted by `text:variable-set`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum OdfVariableSetDisplay {
    Value,
    None,
}

impl OdfVariableSetDisplay {
    const fn as_str(self) -> &'static str {
        match self {
            Self::Value => "value",
            Self::None => "none",
        }
    }

    fn parse(value: &str) -> Result<Self> {
        match value {
            "value" => Ok(Self::Value),
            "none" => Ok(Self::None),
            _ => Err(Error::InvalidFormat(format!(
                "invalid variable-set text:display '{value}'"
            ))),
        }
    }
}

/// Display mode permitted by calculated expressions and variable getters.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum OdfFormulaFieldDisplay {
    Value,
    Formula,
}

impl OdfFormulaFieldDisplay {
    const fn as_str(self) -> &'static str {
        match self {
            Self::Value => "value",
            Self::Formula => "formula",
        }
    }

    fn parse(value: &str) -> Result<Self> {
        match value {
            "value" => Ok(Self::Value),
            "formula" => Ok(Self::Formula),
            _ => Err(Error::InvalidFormat(format!(
                "invalid calculated field text:display '{value}'"
            ))),
        }
    }
}

/// Strict ODF `common-value-and-type-attlist` cached value group.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub enum OdfCalculatedFieldValue {
    Float(String),
    Percentage(String),
    Currency {
        value: String,
        currency: Option<String>,
    },
    Date(String),
    Time(String),
    Boolean(bool),
    String(Option<String>),
}

/// ODF `office:value-type` used by variable input fields.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum OdfFieldValueType {
    Float,
    Time,
    Date,
    Percentage,
    Currency,
    Boolean,
    String,
}

impl OdfFieldValueType {
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Float => "float",
            Self::Time => "time",
            Self::Date => "date",
            Self::Percentage => "percentage",
            Self::Currency => "currency",
            Self::Boolean => "boolean",
            Self::String => "string",
        }
    }

    fn parse(value: &str) -> Result<Self> {
        match value {
            "float" => Ok(Self::Float),
            "time" => Ok(Self::Time),
            "date" => Ok(Self::Date),
            "percentage" => Ok(Self::Percentage),
            "currency" => Ok(Self::Currency),
            "boolean" => Ok(Self::Boolean),
            "string" => Ok(Self::String),
            _ => Err(Error::InvalidFormat(format!(
                "invalid variable input office:value-type '{value}'"
            ))),
        }
    }
}

/// Display mode permitted by `text:user-field-get`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum OdfUserFieldDisplay {
    Value,
    Formula,
    None,
}

/// Component displayed by a `text:measure` field.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum OdfMeasureKind {
    Value,
    Unit,
    Gap,
}

/// Display format shared by `text:reference-ref` and `text:bookmark-ref`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum OdfCrossReferenceFormat {
    Page,
    Chapter,
    Direction,
    Text,
    NumberNoSuperior,
    NumberAllSuperior,
    Number,
}

impl OdfCrossReferenceFormat {
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Page => "page",
            Self::Chapter => "chapter",
            Self::Direction => "direction",
            Self::Text => "text",
            Self::NumberNoSuperior => "number-no-superior",
            Self::NumberAllSuperior => "number-all-superior",
            Self::Number => "number",
        }
    }

    fn parse(value: &str) -> Result<Self> {
        match value {
            "page" => Ok(Self::Page),
            "chapter" => Ok(Self::Chapter),
            "direction" => Ok(Self::Direction),
            "text" => Ok(Self::Text),
            "number-no-superior" => Ok(Self::NumberNoSuperior),
            "number-all-superior" => Ok(Self::NumberAllSuperior),
            "number" => Ok(Self::Number),
            _ => Err(Error::InvalidFormat(format!(
                "invalid cross-reference format '{value}'"
            ))),
        }
    }
}

/// Display format permitted by `text:note-ref`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum OdfNoteReferenceFormat {
    Page,
    Chapter,
    Direction,
    Text,
}

impl OdfNoteReferenceFormat {
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Page => "page",
            Self::Chapter => "chapter",
            Self::Direction => "direction",
            Self::Text => "text",
        }
    }

    fn parse(value: &str) -> Result<Self> {
        match value {
            "page" => Ok(Self::Page),
            "chapter" => Ok(Self::Chapter),
            "direction" => Ok(Self::Direction),
            "text" => Ok(Self::Text),
            _ => Err(Error::InvalidFormat(format!(
                "invalid note reference format '{value}'"
            ))),
        }
    }
}

/// Note class selected by `text:note-ref`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum OdfNoteReferenceClass {
    Footnote,
    Endnote,
}

/// Kind of cached ODF document statistic.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum OdfDocumentStatisticKind {
    Page,
    Paragraph,
    Word,
    Character,
    Table,
    Image,
    Object,
}

impl OdfDocumentStatisticKind {
    pub const fn element_name(self) -> &'static str {
        match self {
            Self::Page => "text:page-count",
            Self::Paragraph => "text:paragraph-count",
            Self::Word => "text:word-count",
            Self::Character => "text:character-count",
            Self::Table => "text:table-count",
            Self::Image => "text:image-count",
            Self::Object => "text:object-count",
        }
    }
}

/// One of the eight temporal/revision ODF document-metadata fields.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum OdfDocumentMetadataFieldKind {
    CreationDate,
    CreationTime,
    PrintDate,
    PrintTime,
    EditingCycles,
    EditingDuration,
    ModificationDate,
    ModificationTime,
}

impl OdfDocumentMetadataFieldKind {
    pub const fn element_name(self) -> &'static str {
        match self {
            Self::CreationDate => "text:creation-date",
            Self::CreationTime => "text:creation-time",
            Self::PrintDate => "text:print-date",
            Self::PrintTime => "text:print-time",
            Self::EditingCycles => "text:editing-cycles",
            Self::EditingDuration => "text:editing-duration",
            Self::ModificationDate => "text:modification-date",
            Self::ModificationTime => "text:modification-time",
        }
    }

    const fn permits_data_style(self) -> bool {
        !matches!(self, Self::EditingCycles)
    }
}

/// Strict typed value attribute for a temporal document-metadata field.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub enum OdfDocumentMetadataFieldValue {
    Date(OdfFieldDateValue),
    Time(OdfFieldTimeValue),
    Duration(OdfFieldDuration),
}

/// One of the nine fixed string/identity document-metadata fields.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum OdfDocumentIdentityFieldKind {
    InitialCreator,
    Description,
    PrintedBy,
    Title,
    Subject,
    Keywords,
    Creator,
    /// The cached full name of the document author from ODF 1.2 §7.3.7.1.
    AuthorName,
    /// The cached initials of the document author from ODF 1.2 §7.3.7.2.
    AuthorInitials,
}

impl OdfDocumentIdentityFieldKind {
    pub const fn element_name(self) -> &'static str {
        match self {
            Self::InitialCreator => "text:initial-creator",
            Self::Description => "text:description",
            Self::PrintedBy => "text:printed-by",
            Self::Title => "text:title",
            Self::Subject => "text:subject",
            Self::Keywords => "text:keywords",
            Self::Creator => "text:creator",
            Self::AuthorName => "text:author-name",
            Self::AuthorInitials => "text:author-initials",
        }
    }
}

/// Independently optional cached values permitted by `text:user-defined`.
///
/// Unlike variable fields, ODF 1.2 does not use `office:value-type` here and
/// its schema permits more than one of these attributes to coexist.
#[derive(Debug, Clone, PartialEq, Eq, Hash, Default)]
pub struct OdfUserDefinedMetadataValues {
    pub number: Option<String>,
    pub date: Option<OdfFieldDateValue>,
    pub time: Option<OdfFieldDuration>,
    pub boolean: Option<bool>,
    pub string: Option<String>,
}

/// A namespace-resolved attribute on inert `text:meta-field` content.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct OdfMetaFieldAttribute {
    pub namespace_uri: String,
    pub local_name: String,
    pub value: String,
}

/// A namespace-resolved inert inline element.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct OdfMetaFieldElement {
    pub namespace_uri: String,
    pub local_name: String,
    pub attributes: Vec<OdfMetaFieldAttribute>,
    pub children: Vec<OdfMetaFieldNode>,
}

/// Ordered mixed content retained by `text:meta-field`.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub enum OdfMetaFieldNode {
    Text(String),
    Element(OdfMetaFieldElement),
}

/// Validated, inert mixed content with a cached plain-text projection.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct OdfMetaFieldContent {
    nodes: Vec<OdfMetaFieldNode>,
    display_text: String,
}

impl OdfMetaFieldContent {
    pub fn new(nodes: Vec<OdfMetaFieldNode>) -> Result<Self> {
        let display_text =
            validated_meta_display_text(&nodes, MetaContentGrammar::ParagraphOrHyperlink)?;
        Ok(Self {
            nodes,
            display_text,
        })
    }

    pub fn nodes(&self) -> &[OdfMetaFieldNode] {
        &self.nodes
    }

    pub fn display_text(&self) -> &str {
        &self.display_text
    }

    fn write_xml(&self, output: &mut String) {
        for node in &self.nodes {
            write_meta_node(node, output);
        }
    }
}

/// Validated, inert structured content for an ODF `text:note-body`.
///
/// The ODF 1.3 schema permits paragraph-like blocks, lists, tables, selected
/// drawing content, and related structured text descendants in a note body.
/// This models ODF 1.3 Part 3, section 6.3.4. Direct character data is
/// deliberately rejected: it belongs inside one of those schema-defined child
/// elements. Links, fields, event listeners, and macro metadata are serialized
/// only as inert XML; this type never follows, evaluates, or executes them.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct OdfNoteBodyContent {
    nodes: Vec<OdfMetaFieldNode>,
    display_text: String,
}

impl OdfNoteBodyContent {
    /// Construct structured note-body content from namespace-resolved nodes.
    pub fn new(nodes: Vec<OdfMetaFieldNode>) -> Result<Self> {
        if nodes
            .iter()
            .any(|node| matches!(node, OdfMetaFieldNode::Text(_)))
        {
            return Err(Error::InvalidFormat(
                "text:note-body cannot contain direct character data".to_string(),
            ));
        }
        validated_meta_display_text(&nodes, MetaContentGrammar::NoteBody)?;
        let display_text = note_body_display_text(&nodes)?;
        Ok(Self {
            nodes,
            display_text,
        })
    }

    /// Return the ordered, namespace-resolved note-body nodes.
    pub fn nodes(&self) -> &[OdfMetaFieldNode] {
        &self.nodes
    }

    /// Return a cached visible-text projection of the structured note body.
    ///
    /// Paragraph and heading descendants are separated by line feeds. Nested
    /// note bodies are omitted from an enclosing note's projection, while a
    /// nested note's citation remains inline. `text:s`, `text:tab`, and
    /// `text:line-break` receive their corresponding text semantics, matching
    /// the bounded semantic note reader.
    pub fn display_text(&self) -> &str {
        &self.display_text
    }

    /// Return whether this body has no schema child elements.
    pub fn is_empty(&self) -> bool {
        self.nodes.is_empty()
    }

    /// Revalidate this value's bounded XML and resource constraints.
    pub fn validate(&self) -> Result<()> {
        Self::new(self.nodes.clone()).map(|_| ())
    }

    pub(crate) fn write_xml(&self, output: &mut String) {
        for node in &self.nodes {
            write_meta_node(node, output);
        }
    }
}

fn validated_meta_display_text(
    nodes: &[OdfMetaFieldNode],
    grammar: MetaContentGrammar,
) -> Result<String> {
    let mut aggregate = 0usize;
    let mut node_count = 0usize;
    let mut display_text = String::new();
    validate_meta_nodes(
        nodes,
        0,
        grammar,
        &mut aggregate,
        &mut node_count,
        &mut display_text,
    )?;
    Ok(display_text)
}

fn note_body_display_text(nodes: &[OdfMetaFieldNode]) -> Result<String> {
    let mut display_text = String::new();
    let mut seen_block = false;
    append_note_body_display_text(nodes, &mut display_text, &mut seen_block, false)?;
    Ok(display_text)
}

fn append_note_body_display_text(
    nodes: &[OdfMetaFieldNode],
    display_text: &mut String,
    seen_block: &mut bool,
    in_paragraph: bool,
) -> Result<()> {
    for node in nodes {
        match node {
            OdfMetaFieldNode::Text(value) if in_paragraph => {
                append_note_body_display_value(display_text, value)?;
            },
            OdfMetaFieldNode::Text(_) => {},
            OdfMetaFieldNode::Element(element) => {
                if element.namespace_uri == TEXT_DATABASE_NAMESPACE && element.local_name == "note"
                {
                    if in_paragraph
                        && let Some(OdfMetaFieldNode::Element(citation)) = element.children.first()
                        && citation.namespace_uri == TEXT_DATABASE_NAMESPACE
                        && citation.local_name == "note-citation"
                    {
                        append_note_body_display_text(
                            &citation.children,
                            display_text,
                            seen_block,
                            true,
                        )?;
                    }
                    continue;
                }
                if element.namespace_uri == TEXT_DATABASE_NAMESPACE
                    && matches!(element.local_name.as_str(), "p" | "h")
                {
                    if !in_paragraph {
                        if *seen_block {
                            append_note_body_display_value(display_text, "\n")?;
                        }
                        *seen_block = true;
                    }
                    append_note_body_display_text(
                        &element.children,
                        display_text,
                        seen_block,
                        true,
                    )?;
                    continue;
                }
                if in_paragraph && element.namespace_uri == TEXT_DATABASE_NAMESPACE {
                    match element.local_name.as_str() {
                        "s" => {
                            append_note_body_spaces(display_text, element)?;
                            continue;
                        },
                        "tab" => {
                            append_note_body_display_value(display_text, "\t")?;
                            continue;
                        },
                        "line-break" => {
                            append_note_body_display_value(display_text, "\n")?;
                            continue;
                        },
                        _ => {},
                    }
                }
                append_note_body_display_text(
                    &element.children,
                    display_text,
                    seen_block,
                    in_paragraph,
                )?;
            },
        }
    }
    Ok(())
}

fn append_note_body_display_value(output: &mut String, value: &str) -> Result<()> {
    let total = output.len().checked_add(value.len()).ok_or_else(|| {
        Error::InvalidFormat("text:note-body display text size overflow".to_string())
    })?;
    if total > MAX_DYNAMIC_FIELD_AGGREGATE {
        return Err(Error::InvalidFormat(format!(
            "text:note-body display text exceeds {MAX_DYNAMIC_FIELD_AGGREGATE} bytes"
        )));
    }
    output.push_str(value);
    Ok(())
}

fn append_note_body_spaces(output: &mut String, element: &OdfMetaFieldElement) -> Result<()> {
    let count = element
        .attributes
        .iter()
        .find(|attribute| {
            attribute.namespace_uri == TEXT_DATABASE_NAMESPACE && attribute.local_name == "c"
        })
        .map(|attribute| {
            attribute.value.parse::<usize>().map_err(|_| {
                Error::InvalidFormat("text:s text:c must be a non-negative integer".to_string())
            })
        })
        .transpose()?
        .unwrap_or(1);
    let total = output.len().checked_add(count).ok_or_else(|| {
        Error::InvalidFormat("text:note-body display text size overflow".to_string())
    })?;
    if total > MAX_DYNAMIC_FIELD_AGGREGATE {
        return Err(Error::InvalidFormat(format!(
            "text:note-body display text exceeds {MAX_DYNAMIC_FIELD_AGGREGATE} bytes"
        )));
    }
    output.extend(std::iter::repeat_n(' ', count));
    Ok(())
}

impl OdfUserDefinedMetadataValues {
    fn validate(&self, aggregate: &mut usize) -> Result<()> {
        if let Some(number) = &self.number {
            validate_double(number)?;
            validate_dynamic_value("office:value", Some(number), true, aggregate)?;
        }
        if let Some(date) = &self.date {
            date.validate(aggregate)?;
        }
        if let Some(time) = &self.time {
            time.validate("office:time-value", aggregate)?;
        }
        validate_dynamic_value(
            "office:string-value",
            self.string.as_deref(),
            false,
            aggregate,
        )
    }

    fn write_attributes(&self, element: &mut Element) {
        if self.number.is_none()
            && self.date.is_none()
            && self.time.is_none()
            && self.boolean.is_none()
            && self.string.is_none()
        {
            return;
        }
        element.set_attribute("xmlns:office", OFFICE_NAMESPACE);
        if let Some(number) = &self.number {
            element.set_attribute("office:value", number);
        }
        if let Some(date) = &self.date {
            element.set_attribute("office:date-value", date.as_str());
        }
        if let Some(time) = &self.time {
            element.set_attribute("office:time-value", time.as_str());
        }
        if let Some(boolean) = self.boolean {
            element.set_attribute(
                "office:boolean-value",
                if boolean { "true" } else { "false" },
            );
        }
        if let Some(string) = &self.string {
            element.set_attribute("office:string-value", string);
        }
    }
}

impl OdfNoteReferenceClass {
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Footnote => "footnote",
            Self::Endnote => "endnote",
        }
    }

    fn parse(value: &str) -> Result<Self> {
        match value {
            "footnote" => Ok(Self::Footnote),
            "endnote" => Ok(Self::Endnote),
            _ => Err(Error::InvalidFormat(format!(
                "invalid text:note-class '{value}'"
            ))),
        }
    }
}

impl OdfMeasureKind {
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Value => "value",
            Self::Unit => "unit",
            Self::Gap => "gap",
        }
    }

    fn parse(value: &str) -> Result<Self> {
        match value {
            "value" => Ok(Self::Value),
            "unit" => Ok(Self::Unit),
            "gap" => Ok(Self::Gap),
            _ => Err(Error::InvalidFormat(format!(
                "invalid text:measure kind '{value}'"
            ))),
        }
    }
}

impl OdfUserFieldDisplay {
    const fn as_str(self) -> &'static str {
        match self {
            Self::Value => "value",
            Self::Formula => "formula",
            Self::None => "none",
        }
    }

    fn parse(value: &str) -> Result<Self> {
        match value {
            "value" => Ok(Self::Value),
            "formula" => Ok(Self::Formula),
            "none" => Ok(Self::None),
            _ => Err(Error::InvalidFormat(format!(
                "invalid user-field-get text:display '{value}'"
            ))),
        }
    }
}

impl OdfCalculatedFieldValue {
    fn validate(&self, aggregate: &mut usize) -> Result<()> {
        match self {
            Self::Float(value) | Self::Percentage(value) => {
                validate_double(value)?;
                validate_dynamic_value("office:value", Some(value), true, aggregate)
            },
            Self::Currency { value, currency } => {
                validate_double(value)?;
                validate_dynamic_value("office:value", Some(value), true, aggregate)?;
                validate_dynamic_value("office:currency", currency.as_deref(), false, aggregate)
            },
            Self::Date(value) => {
                if value.contains('T') {
                    crate::datatype::DateTimeOdf::decode(value).map_err(|_| {
                        Error::InvalidFormat(format!("invalid office:date-value '{value}'"))
                    })?;
                } else {
                    crate::datatype::Date::decode(value).map_err(|_| {
                        Error::InvalidFormat(format!("invalid office:date-value '{value}'"))
                    })?;
                }
                validate_dynamic_value("office:date-value", Some(value), true, aggregate)
            },
            Self::Time(value) => {
                crate::datatype::DurationOdf::decode_exact(value).map_err(|_| {
                    Error::InvalidFormat(format!("invalid office:time-value '{value}'"))
                })?;
                validate_dynamic_value("office:time-value", Some(value), true, aggregate)
            },
            Self::Boolean(_) => Ok(()),
            Self::String(value) => {
                validate_dynamic_value("office:string-value", value.as_deref(), false, aggregate)
            },
        }
    }

    fn write_attributes(&self, element: &mut Element) {
        element.set_attribute("xmlns:office", OFFICE_NAMESPACE);
        match self {
            Self::Float(value) => {
                element.set_attribute("office:value-type", "float");
                element.set_attribute("office:value", value);
            },
            Self::Percentage(value) => {
                element.set_attribute("office:value-type", "percentage");
                element.set_attribute("office:value", value);
            },
            Self::Currency { value, currency } => {
                element.set_attribute("office:value-type", "currency");
                element.set_attribute("office:value", value);
                if let Some(currency) = currency {
                    element.set_attribute("office:currency", currency);
                }
            },
            Self::Date(value) => {
                element.set_attribute("office:value-type", "date");
                element.set_attribute("office:date-value", value);
            },
            Self::Time(value) => {
                element.set_attribute("office:value-type", "time");
                element.set_attribute("office:time-value", value);
            },
            Self::Boolean(value) => {
                element.set_attribute("office:value-type", "boolean");
                element.set_attribute(
                    "office:boolean-value",
                    if *value { "true" } else { "false" },
                );
            },
            Self::String(value) => {
                element.set_attribute("office:value-type", "string");
                if let Some(value) = value {
                    element.set_attribute("office:string-value", value);
                }
            },
        }
    }
}

impl OdfSequenceReferenceFormat {
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Page => "page",
            Self::Chapter => "chapter",
            Self::Direction => "direction",
            Self::Text => "text",
            Self::CategoryAndValue => "category-and-value",
            Self::Caption => "caption",
            Self::Value => "value",
        }
    }

    fn parse(value: &str) -> Result<Self> {
        match value {
            "page" => Ok(Self::Page),
            "chapter" => Ok(Self::Chapter),
            "direction" => Ok(Self::Direction),
            "text" => Ok(Self::Text),
            "category-and-value" => Ok(Self::CategoryAndValue),
            "caption" => Ok(Self::Caption),
            "value" => Ok(Self::Value),
            _ => Err(Error::InvalidFormat(format!(
                "invalid text:sequence-ref reference format '{value}'"
            ))),
        }
    }
}

impl OdfSequenceNumberFormat {
    pub fn new(format: impl Into<String>, letter_sync: Option<bool>) -> Result<Self> {
        let value = Self {
            format: format.into(),
            letter_sync,
        };
        value.validate()?;
        Ok(value)
    }

    pub fn format(&self) -> &str {
        &self.format
    }

    pub const fn letter_sync(&self) -> Option<bool> {
        self.letter_sync
    }

    fn validate(&self) -> Result<()> {
        let mut aggregate = 0usize;
        validate_dynamic_value(
            "style:num-format",
            Some(&self.format),
            false,
            &mut aggregate,
        )?;
        if self.letter_sync.is_some() && !matches!(self.format.as_str(), "a" | "A") {
            return Err(Error::InvalidFormat(
                "style:num-letter-sync requires alphabetic style:num-format 'a' or 'A'".to_string(),
            ));
        }
        Ok(())
    }
}

impl OdfPlaceholderType {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "text" => Ok(Self::Text),
            "table" => Ok(Self::Table),
            "text-box" => Ok(Self::TextBox),
            "image" => Ok(Self::Image),
            "object" => Ok(Self::Object),
            _ => Err(Error::InvalidFormat(format!(
                "invalid text:placeholder-type '{value}'"
            ))),
        }
    }

    /// Return the ODF 1.2 lexical value for `text:placeholder-type`.
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Text => "text",
            Self::Table => "table",
            Self::TextBox => "text-box",
            Self::Image => "image",
            Self::Object => "object",
        }
    }
}

/// Typed conditional and placeholder text metadata in document order.
///
/// Formula strings are retained verbatim and are never evaluated. `display_text`
/// is the cached text stored by the document producer.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum OdfDynamicTextField {
    Placeholder {
        placeholder_type: OdfPlaceholderType,
        description: Option<String>,
        display_text: String,
    },
    ConditionalText {
        condition: String,
        value_if_true: String,
        value_if_false: String,
        current_value: Option<bool>,
        display_text: String,
    },
    HiddenText {
        condition: String,
        string_value: String,
        is_hidden: Option<bool>,
        display_text: String,
    },
    HiddenParagraph {
        condition: String,
        is_hidden: Option<bool>,
        display_text: String,
    },
    /// An inert calculated sequence field from ODF 1.2 section 7.4.11.
    Sequence {
        name: String,
        formula: Option<String>,
        number_format: Option<OdfSequenceNumberFormat>,
        reference_name: Option<String>,
        display_text: String,
    },
    /// A cached reference to a named sequence value.
    SequenceReference {
        reference_name: String,
        reference_format: Option<OdfSequenceReferenceFormat>,
        display_text: String,
    },
    VariableSet {
        name: String,
        formula: Option<String>,
        value: OdfCalculatedFieldValue,
        display: Option<OdfVariableSetDisplay>,
        data_style_name: Option<String>,
        display_text: String,
    },
    VariableGet {
        name: String,
        display: Option<OdfFormulaFieldDisplay>,
        data_style_name: Option<String>,
        display_text: String,
    },
    Expression {
        formula: Option<String>,
        value: Option<OdfCalculatedFieldValue>,
        display: Option<OdfFormulaFieldDisplay>,
        data_style_name: Option<String>,
        display_text: String,
    },
    VariableInput {
        name: String,
        description: Option<String>,
        value_type: OdfFieldValueType,
        display: Option<OdfVariableSetDisplay>,
        data_style_name: Option<String>,
        display_text: String,
    },
    UserFieldGet {
        name: String,
        display: Option<OdfUserFieldDisplay>,
        data_style_name: Option<String>,
        display_text: String,
    },
    UserFieldInput {
        name: String,
        description: Option<String>,
        data_style_name: Option<String>,
        display_text: String,
    },
    TextInput {
        description: Option<String>,
        display_text: String,
    },
    /// An inert table-cell formula display field.
    TableFormula {
        formula: Option<String>,
        display: Option<OdfFormulaFieldDisplay>,
        data_style_name: Option<String>,
        display_text: String,
    },
    /// Cached, non-calculating measurement field text.
    Measure {
        kind: OdfMeasureKind,
        display_text: String,
    },
    Reference {
        reference_name: Option<String>,
        reference_format: Option<OdfCrossReferenceFormat>,
        display_text: String,
    },
    BookmarkReference {
        reference_name: Option<String>,
        reference_format: Option<OdfCrossReferenceFormat>,
        display_text: String,
    },
    NoteReference {
        reference_name: Option<String>,
        note_class: OdfNoteReferenceClass,
        reference_format: Option<OdfNoteReferenceFormat>,
        display_text: String,
    },
    DocumentStatistic {
        kind: OdfDocumentStatisticKind,
        number_format: Option<OdfStatisticNumberFormat>,
        display_text: String,
    },
    /// Current, previous, or next page number with inert cached presentation.
    PageNumber {
        number_format: Option<OdfPageNumberFormat>,
        fixed: Option<bool>,
        page_adjust: Option<i64>,
        select_page: Option<OdfPageSelection>,
        display_text: String,
    },
    /// Current date or an explicitly fixed date/date-time value.
    Date {
        value: Option<OdfFieldDateValue>,
        adjustment: Option<OdfFieldDuration>,
        fixed: Option<bool>,
        data_style_name: Option<String>,
        display_text: String,
    },
    /// Current time or an explicitly fixed time/date-time value.
    Time {
        value: Option<OdfFieldTimeValue>,
        adjustment: Option<OdfFieldDuration>,
        fixed: Option<bool>,
        data_style_name: Option<String>,
        display_text: String,
    },
    /// Previous/next page continuation reminder.
    PageContinuation {
        select_page: OdfPageContinuationSelection,
        string_value: Option<String>,
        display_text: String,
    },
    /// Set or disable the document's single alternative page variable.
    PageVariableSet {
        active: Option<bool>,
        page_adjust: Option<i64>,
        display_text: String,
    },
    /// Display the current alternative page-variable value.
    PageVariableGet {
        number_format: Option<OdfPageVariableNumberFormat>,
        display_text: String,
    },
    /// Cached presentation and optional fixed value of a metadata field.
    DocumentMetadata {
        kind: OdfDocumentMetadataFieldKind,
        value: Option<OdfDocumentMetadataFieldValue>,
        fixed: Option<bool>,
        data_style_name: Option<String>,
        display_text: String,
    },
    /// Fixed or live cached string metadata such as title or creator.
    ///
    /// Author fields retain stored text only and never read or modify host
    /// identity data.
    DocumentIdentity {
        kind: OdfDocumentIdentityFieldKind,
        fixed: Option<bool>,
        display_text: String,
    },
    /// Named custom document metadata with inert cached typed attributes.
    UserDefinedMetadata {
        name: String,
        values: OdfUserDefinedMetadataValues,
        fixed: Option<bool>,
        data_style_name: Option<String>,
        display_text: String,
    },
    /// Cached data from a named DDE declaration; never refreshed or connected.
    DdeConnection {
        connection_name: String,
        display_text: String,
    },
    /// RDF-backed metadata field with namespace-resolved inert inline content.
    MetaField {
        xml_id: String,
        data_style_name: Option<String>,
        content: OdfMetaFieldContent,
    },
}

impl OdfDynamicTextField {
    /// The cached text present in the ODF file, without evaluating any formula.
    pub fn display_text(&self) -> &str {
        match self {
            Self::Placeholder { display_text, .. }
            | Self::ConditionalText { display_text, .. }
            | Self::HiddenText { display_text, .. }
            | Self::HiddenParagraph { display_text, .. }
            | Self::Sequence { display_text, .. }
            | Self::SequenceReference { display_text, .. }
            | Self::VariableSet { display_text, .. }
            | Self::VariableGet { display_text, .. }
            | Self::Expression { display_text, .. }
            | Self::VariableInput { display_text, .. }
            | Self::UserFieldGet { display_text, .. }
            | Self::UserFieldInput { display_text, .. }
            | Self::TextInput { display_text, .. }
            | Self::TableFormula { display_text, .. }
            | Self::Measure { display_text, .. }
            | Self::Reference { display_text, .. }
            | Self::BookmarkReference { display_text, .. }
            | Self::NoteReference { display_text, .. }
            | Self::DocumentStatistic { display_text, .. }
            | Self::DdeConnection { display_text, .. }
            | Self::PageNumber { display_text, .. }
            | Self::Date { display_text, .. }
            | Self::Time { display_text, .. }
            | Self::PageContinuation { display_text, .. }
            | Self::PageVariableSet { display_text, .. }
            | Self::PageVariableGet { display_text, .. }
            | Self::DocumentMetadata { display_text, .. }
            | Self::DocumentIdentity { display_text, .. }
            | Self::UserDefinedMetadata { display_text, .. } => display_text,
            Self::MetaField { content, .. } => content.display_text(),
        }
    }

    /// Effective `text:active` value for a page-variable setter.
    ///
    /// ODF and LibreOffice default the omitted attribute to `true`.
    pub fn effective_page_variable_active(&self) -> Option<bool> {
        match self {
            Self::PageVariableSet { active, .. } => Some(active.unwrap_or(true)),
            _ => None,
        }
    }

    /// Effective page adjustment for a page-variable setter.
    ///
    /// The standard default for an omitted adjustment is zero.
    pub fn effective_page_variable_adjustment(&self) -> Option<i64> {
        match self {
            Self::PageVariableSet { page_adjust, .. } => Some(page_adjust.unwrap_or(0)),
            _ => None,
        }
    }

    /// Validate this field for safe ODF XML serialization.
    ///
    /// Conditions remain opaque strings: validation never parses or evaluates a
    /// formula. It only enforces required values, bounded allocation sizes, and
    /// XML 1.0 character validity.
    pub fn validate(&self) -> Result<()> {
        let mut aggregate = 0usize;
        match self {
            Self::Placeholder {
                description,
                display_text,
                ..
            } => {
                validate_dynamic_value(
                    "text:description",
                    description.as_deref(),
                    false,
                    &mut aggregate,
                )?;
                validate_dynamic_value(
                    "placeholder display text",
                    Some(display_text),
                    false,
                    &mut aggregate,
                )?;
            },
            Self::ConditionalText {
                condition,
                value_if_true,
                value_if_false,
                display_text,
                ..
            } => {
                validate_dynamic_value("text:condition", Some(condition), true, &mut aggregate)?;
                validate_dynamic_value(
                    "text:string-value-if-true",
                    Some(value_if_true),
                    false,
                    &mut aggregate,
                )?;
                validate_dynamic_value(
                    "text:string-value-if-false",
                    Some(value_if_false),
                    false,
                    &mut aggregate,
                )?;
                validate_dynamic_value(
                    "conditional display text",
                    Some(display_text),
                    false,
                    &mut aggregate,
                )?;
            },
            Self::HiddenText {
                condition,
                string_value,
                display_text,
                ..
            } => {
                validate_dynamic_value("text:condition", Some(condition), true, &mut aggregate)?;
                validate_dynamic_value(
                    "text:string-value",
                    Some(string_value),
                    false,
                    &mut aggregate,
                )?;
                validate_dynamic_value(
                    "hidden text display text",
                    Some(display_text),
                    false,
                    &mut aggregate,
                )?;
            },
            Self::HiddenParagraph {
                condition,
                display_text,
                ..
            } => {
                validate_dynamic_value("text:condition", Some(condition), true, &mut aggregate)?;
                validate_dynamic_value(
                    "hidden paragraph display text",
                    Some(display_text),
                    false,
                    &mut aggregate,
                )?;
            },
            Self::DdeConnection {
                connection_name,
                display_text,
            } => {
                validate_dynamic_value(
                    "text:connection-name",
                    Some(connection_name),
                    false,
                    &mut aggregate,
                )?;
                validate_dynamic_value(
                    "DDE cached text",
                    Some(display_text),
                    false,
                    &mut aggregate,
                )?;
            },
            Self::Sequence {
                name,
                formula,
                number_format,
                reference_name,
                display_text,
            } => {
                validate_dynamic_value("text:name", Some(name), false, &mut aggregate)?;
                validate_dynamic_value("text:formula", formula.as_deref(), false, &mut aggregate)?;
                if let Some(number_format) = number_format {
                    number_format.validate()?;
                    aggregate = aggregate
                        .checked_add(number_format.format().len())
                        .ok_or_else(|| {
                            Error::InvalidFormat("dynamic field size overflow".to_string())
                        })?;
                }
                validate_dynamic_value(
                    "text:ref-name",
                    reference_name.as_deref(),
                    false,
                    &mut aggregate,
                )?;
                validate_dynamic_value(
                    "sequence display text",
                    Some(display_text),
                    false,
                    &mut aggregate,
                )?;
                if aggregate > MAX_DYNAMIC_FIELD_AGGREGATE {
                    return Err(Error::InvalidFormat(format!(
                        "dynamic field exceeds {MAX_DYNAMIC_FIELD_AGGREGATE} aggregate bytes"
                    )));
                }
            },
            Self::SequenceReference {
                reference_name,
                display_text,
                ..
            } => {
                validate_dynamic_value(
                    "text:ref-name",
                    Some(reference_name),
                    true,
                    &mut aggregate,
                )?;
                validate_dynamic_value(
                    "sequence reference display text",
                    Some(display_text),
                    false,
                    &mut aggregate,
                )?;
            },
            Self::VariableSet {
                name,
                formula,
                value,
                data_style_name,
                display_text,
                ..
            } => {
                validate_dynamic_value("text:name", Some(name), true, &mut aggregate)?;
                validate_dynamic_value("text:formula", formula.as_deref(), false, &mut aggregate)?;
                value.validate(&mut aggregate)?;
                validate_dynamic_value(
                    "style:data-style-name",
                    data_style_name.as_deref(),
                    false,
                    &mut aggregate,
                )?;
                validate_dynamic_value(
                    "variable-set display text",
                    Some(display_text),
                    false,
                    &mut aggregate,
                )?;
            },
            Self::VariableGet {
                name,
                data_style_name,
                display_text,
                ..
            } => {
                validate_dynamic_value("text:name", Some(name), true, &mut aggregate)?;
                validate_dynamic_value(
                    "style:data-style-name",
                    data_style_name.as_deref(),
                    false,
                    &mut aggregate,
                )?;
                validate_dynamic_value(
                    "variable-get display text",
                    Some(display_text),
                    false,
                    &mut aggregate,
                )?;
            },
            Self::Expression {
                formula,
                value,
                data_style_name,
                display_text,
                ..
            } => {
                validate_dynamic_value("text:formula", formula.as_deref(), false, &mut aggregate)?;
                if let Some(value) = value {
                    value.validate(&mut aggregate)?;
                }
                validate_dynamic_value(
                    "style:data-style-name",
                    data_style_name.as_deref(),
                    false,
                    &mut aggregate,
                )?;
                validate_dynamic_value(
                    "expression display text",
                    Some(display_text),
                    false,
                    &mut aggregate,
                )?;
            },
            Self::VariableInput {
                name,
                description,
                data_style_name,
                display_text,
                ..
            } => {
                validate_dynamic_value("text:name", Some(name), true, &mut aggregate)?;
                validate_dynamic_value(
                    "text:description",
                    description.as_deref(),
                    false,
                    &mut aggregate,
                )?;
                validate_dynamic_value(
                    "style:data-style-name",
                    data_style_name.as_deref(),
                    false,
                    &mut aggregate,
                )?;
                validate_dynamic_value(
                    "variable-input display text",
                    Some(display_text),
                    false,
                    &mut aggregate,
                )?;
            },
            Self::UserFieldGet {
                name,
                data_style_name,
                display_text,
                ..
            } => {
                validate_dynamic_value("text:name", Some(name), true, &mut aggregate)?;
                validate_dynamic_value(
                    "style:data-style-name",
                    data_style_name.as_deref(),
                    false,
                    &mut aggregate,
                )?;
                validate_dynamic_value(
                    "user-field-get display text",
                    Some(display_text),
                    false,
                    &mut aggregate,
                )?;
            },
            Self::UserFieldInput {
                name,
                description,
                data_style_name,
                display_text,
            } => {
                validate_dynamic_value("text:name", Some(name), true, &mut aggregate)?;
                validate_dynamic_value(
                    "text:description",
                    description.as_deref(),
                    false,
                    &mut aggregate,
                )?;
                validate_dynamic_value(
                    "style:data-style-name",
                    data_style_name.as_deref(),
                    false,
                    &mut aggregate,
                )?;
                validate_dynamic_value(
                    "user-field-input display text",
                    Some(display_text),
                    false,
                    &mut aggregate,
                )?;
            },
            Self::TextInput {
                description,
                display_text,
            } => {
                validate_dynamic_value(
                    "text:description",
                    description.as_deref(),
                    false,
                    &mut aggregate,
                )?;
                validate_dynamic_value(
                    "text-input display text",
                    Some(display_text),
                    false,
                    &mut aggregate,
                )?;
            },
            Self::TableFormula {
                formula,
                data_style_name,
                display_text,
                ..
            } => {
                validate_dynamic_value("text:formula", formula.as_deref(), false, &mut aggregate)?;
                validate_dynamic_value(
                    "style:data-style-name",
                    data_style_name.as_deref(),
                    false,
                    &mut aggregate,
                )?;
                validate_dynamic_value(
                    "table-formula display text",
                    Some(display_text),
                    false,
                    &mut aggregate,
                )?;
            },
            Self::Measure { display_text, .. } => {
                validate_dynamic_value(
                    "measure display text",
                    Some(display_text),
                    false,
                    &mut aggregate,
                )?;
            },
            Self::Reference {
                reference_name,
                display_text,
                ..
            }
            | Self::BookmarkReference {
                reference_name,
                display_text,
                ..
            }
            | Self::NoteReference {
                reference_name,
                display_text,
                ..
            } => {
                validate_dynamic_value(
                    "text:ref-name",
                    reference_name.as_deref(),
                    false,
                    &mut aggregate,
                )?;
                validate_dynamic_value(
                    "reference display text",
                    Some(display_text),
                    false,
                    &mut aggregate,
                )?;
            },
            Self::DocumentStatistic {
                number_format,
                display_text,
                ..
            } => {
                if let Some(number_format) = number_format {
                    number_format.validate()?;
                    aggregate = aggregate
                        .checked_add(number_format.format().len())
                        .ok_or_else(|| {
                            Error::InvalidFormat("dynamic field size overflow".to_string())
                        })?;
                }
                validate_dynamic_value(
                    "statistic display text",
                    Some(display_text),
                    false,
                    &mut aggregate,
                )?;
                if aggregate > MAX_DYNAMIC_FIELD_AGGREGATE {
                    return Err(Error::InvalidFormat(format!(
                        "dynamic field exceeds {MAX_DYNAMIC_FIELD_AGGREGATE} aggregate bytes"
                    )));
                }
            },
            Self::PageNumber {
                number_format,
                display_text,
                ..
            } => {
                if let Some(number_format) = number_format {
                    number_format.validate()?;
                    aggregate = aggregate
                        .checked_add(number_format.format().len())
                        .ok_or_else(|| {
                            Error::InvalidFormat("dynamic field size overflow".to_string())
                        })?;
                }
                validate_dynamic_value(
                    "page-number display text",
                    Some(display_text),
                    false,
                    &mut aggregate,
                )?;
            },
            Self::Date {
                value,
                adjustment,
                data_style_name,
                display_text,
                ..
            } => {
                if let Some(value) = value {
                    value.validate(&mut aggregate)?;
                }
                if let Some(adjustment) = adjustment {
                    adjustment.validate("text:date-adjust", &mut aggregate)?;
                }
                validate_dynamic_value(
                    "style:data-style-name",
                    data_style_name.as_deref(),
                    false,
                    &mut aggregate,
                )?;
                validate_dynamic_value(
                    "date display text",
                    Some(display_text),
                    false,
                    &mut aggregate,
                )?;
            },
            Self::Time {
                value,
                adjustment,
                data_style_name,
                display_text,
                ..
            } => {
                if let Some(value) = value {
                    value.validate(&mut aggregate)?;
                }
                if let Some(adjustment) = adjustment {
                    adjustment.validate("text:time-adjust", &mut aggregate)?;
                }
                validate_dynamic_value(
                    "style:data-style-name",
                    data_style_name.as_deref(),
                    false,
                    &mut aggregate,
                )?;
                validate_dynamic_value(
                    "time display text",
                    Some(display_text),
                    false,
                    &mut aggregate,
                )?;
            },
            Self::PageContinuation {
                string_value,
                display_text,
                ..
            } => {
                validate_dynamic_value(
                    "text:string-value",
                    string_value.as_deref(),
                    false,
                    &mut aggregate,
                )?;
                validate_dynamic_value(
                    "page-continuation display text",
                    Some(display_text),
                    false,
                    &mut aggregate,
                )?;
            },
            Self::PageVariableSet { display_text, .. } => {
                validate_dynamic_value(
                    "page-variable-set display text",
                    Some(display_text),
                    false,
                    &mut aggregate,
                )?;
            },
            Self::PageVariableGet {
                number_format,
                display_text,
            } => {
                if let Some(number_format) = number_format {
                    number_format.validate()?;
                    aggregate = aggregate
                        .checked_add(number_format.format().len())
                        .ok_or_else(|| {
                            Error::InvalidFormat("dynamic field size overflow".to_string())
                        })?;
                }
                validate_dynamic_value(
                    "page-variable-get display text",
                    Some(display_text),
                    false,
                    &mut aggregate,
                )?;
            },
            Self::DocumentMetadata {
                kind,
                value,
                data_style_name,
                display_text,
                ..
            } => {
                validate_document_metadata_value(*kind, value.as_ref(), &mut aggregate)?;
                if !kind.permits_data_style() && data_style_name.is_some() {
                    return Err(Error::InvalidFormat(format!(
                        "{} does not permit style:data-style-name",
                        kind.element_name()
                    )));
                }
                validate_dynamic_value(
                    "style:data-style-name",
                    data_style_name.as_deref(),
                    false,
                    &mut aggregate,
                )?;
                validate_dynamic_value(
                    "document metadata display text",
                    Some(display_text),
                    false,
                    &mut aggregate,
                )?;
            },
            Self::DocumentIdentity { display_text, .. } => {
                validate_dynamic_value(
                    "document identity display text",
                    Some(display_text),
                    false,
                    &mut aggregate,
                )?;
            },
            Self::UserDefinedMetadata {
                name,
                values,
                data_style_name,
                display_text,
                ..
            } => {
                validate_dynamic_value("text:name", Some(name), false, &mut aggregate)?;
                values.validate(&mut aggregate)?;
                validate_dynamic_value(
                    "style:data-style-name",
                    data_style_name.as_deref(),
                    false,
                    &mut aggregate,
                )?;
                validate_dynamic_value(
                    "user-defined metadata display text",
                    Some(display_text),
                    false,
                    &mut aggregate,
                )?;
            },
            Self::MetaField {
                xml_id,
                data_style_name,
                content,
            } => {
                validate_xml_id(xml_id)?;
                validate_dynamic_value("xml:id", Some(xml_id), true, &mut aggregate)?;
                validate_dynamic_value(
                    "style:data-style-name",
                    data_style_name.as_deref(),
                    false,
                    &mut aggregate,
                )?;
                let rebuilt = OdfMetaFieldContent::new(content.nodes.clone())?;
                if &rebuilt != content {
                    return Err(Error::InvalidFormat(
                        "inconsistent text:meta-field content cache".to_string(),
                    ));
                }
            },
        }
        Ok(())
    }

    /// Serialize one self-contained ODF field element.
    ///
    /// The returned fragment declares the `text` namespace locally, so it is
    /// namespace-correct regardless of the prefixes used by its destination
    /// document. Formula attributes are emitted verbatim after XML escaping and
    /// are never executed.
    pub fn to_xml_fragment(&self) -> Result<String> {
        if let Self::MetaField {
            xml_id,
            data_style_name,
            content,
        } = self
        {
            self.validate()?;
            let mut xml = String::new();
            xml.push_str("<text:meta-field xmlns:text=\"");
            xml.push_str(TEXT_DATABASE_NAMESPACE);
            xml.push_str("\" xml:id=\"");
            push_xml_attribute(&mut xml, xml_id);
            xml.push('"');
            if let Some(data_style_name) = data_style_name {
                xml.push_str(" xmlns:style=\"");
                xml.push_str(STYLE_NAMESPACE);
                xml.push_str("\" style:data-style-name=\"");
                push_xml_attribute(&mut xml, data_style_name);
                xml.push('"');
            }
            xml.push('>');
            content.write_xml(&mut xml);
            xml.push_str("</text:meta-field>");
            return Ok(xml);
        }
        Ok(self.to_element()?.to_xml_string())
    }

    pub(crate) fn to_element(&self) -> Result<Element> {
        self.validate()?;
        let mut element = match self {
            Self::Placeholder { .. } => Element::new("text:placeholder"),
            Self::ConditionalText { .. } => Element::new("text:conditional-text"),
            Self::HiddenText { .. } => Element::new("text:hidden-text"),
            Self::HiddenParagraph { .. } => Element::new("text:hidden-paragraph"),
            Self::DdeConnection { .. } => Element::new("text:dde-connection"),
            Self::Sequence { .. } => Element::new("text:sequence"),
            Self::SequenceReference { .. } => Element::new("text:sequence-ref"),
            Self::VariableSet { .. } => Element::new("text:variable-set"),
            Self::VariableGet { .. } => Element::new("text:variable-get"),
            Self::Expression { .. } => Element::new("text:expression"),
            Self::VariableInput { .. } => Element::new("text:variable-input"),
            Self::UserFieldGet { .. } => Element::new("text:user-field-get"),
            Self::UserFieldInput { .. } => Element::new("text:user-field-input"),
            Self::TextInput { .. } => Element::new("text:text-input"),
            Self::TableFormula { .. } => Element::new("text:table-formula"),
            Self::Measure { .. } => Element::new("text:measure"),
            Self::Reference { .. } => Element::new("text:reference-ref"),
            Self::BookmarkReference { .. } => Element::new("text:bookmark-ref"),
            Self::NoteReference { .. } => Element::new("text:note-ref"),
            Self::DocumentStatistic { kind, .. } => Element::new(kind.element_name()),
            Self::PageNumber { .. } => Element::new("text:page-number"),
            Self::Date { .. } => Element::new("text:date"),
            Self::Time { .. } => Element::new("text:time"),
            Self::PageContinuation { .. } => Element::new("text:page-continuation"),
            Self::PageVariableSet { .. } => Element::new("text:page-variable-set"),
            Self::PageVariableGet { .. } => Element::new("text:page-variable-get"),
            Self::DocumentMetadata { kind, .. } => Element::new(kind.element_name()),
            Self::DocumentIdentity { kind, .. } => Element::new(kind.element_name()),
            Self::UserDefinedMetadata { .. } => Element::new("text:user-defined"),
            Self::MetaField { .. } => unreachable!("meta-field uses ordered mixed serializer"),
        };
        element.set_attribute("xmlns:text", TEXT_DATABASE_NAMESPACE);

        match self {
            Self::Placeholder {
                placeholder_type,
                description,
                display_text,
            } => {
                element.set_attribute("text:placeholder-type", placeholder_type.as_str());
                if let Some(description) = description {
                    element.set_attribute("text:description", description);
                }
                element.set_text(display_text);
            },
            Self::ConditionalText {
                condition,
                value_if_true,
                value_if_false,
                current_value,
                display_text,
            } => {
                element.set_attribute("text:condition", condition);
                element.set_attribute("text:string-value-if-true", value_if_true);
                element.set_attribute("text:string-value-if-false", value_if_false);
                if let Some(current_value) = current_value {
                    element.set_attribute(
                        "text:current-value",
                        if *current_value { "true" } else { "false" },
                    );
                }
                element.set_text(display_text);
            },
            Self::HiddenText {
                condition,
                string_value,
                is_hidden,
                display_text,
            } => {
                element.set_attribute("text:condition", condition);
                element.set_attribute("text:string-value", string_value);
                if let Some(is_hidden) = is_hidden {
                    element
                        .set_attribute("text:is-hidden", if *is_hidden { "true" } else { "false" });
                }
                element.set_text(display_text);
            },
            Self::HiddenParagraph {
                condition,
                is_hidden,
                display_text,
            } => {
                element.set_attribute("text:condition", condition);
                if let Some(is_hidden) = is_hidden {
                    element
                        .set_attribute("text:is-hidden", if *is_hidden { "true" } else { "false" });
                }
                element.set_text(display_text);
            },
            Self::DdeConnection {
                connection_name,
                display_text,
            } => {
                element.set_attribute("text:connection-name", connection_name);
                element.set_text(display_text);
            },
            Self::Sequence {
                name,
                formula,
                number_format,
                reference_name,
                display_text,
            } => {
                element.set_attribute("text:name", name);
                if let Some(formula) = formula {
                    element.set_attribute("text:formula", formula);
                }
                if let Some(number_format) = number_format {
                    element.set_attribute(
                        "xmlns:style",
                        "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
                    );
                    element.set_attribute("style:num-format", number_format.format());
                    if let Some(letter_sync) = number_format.letter_sync() {
                        element.set_attribute(
                            "style:num-letter-sync",
                            if letter_sync { "true" } else { "false" },
                        );
                    }
                }
                if let Some(reference_name) = reference_name {
                    element.set_attribute("text:ref-name", reference_name);
                }
                element.set_text(display_text);
            },
            Self::SequenceReference {
                reference_name,
                reference_format,
                display_text,
            } => {
                element.set_attribute("text:ref-name", reference_name);
                if let Some(reference_format) = reference_format {
                    element.set_attribute("text:reference-format", reference_format.as_str());
                }
                element.set_text(display_text);
            },
            Self::VariableSet {
                name,
                formula,
                value,
                display,
                data_style_name,
                display_text,
            } => {
                element.set_attribute("text:name", name);
                if let Some(formula) = formula {
                    element.set_attribute("text:formula", formula);
                }
                value.write_attributes(&mut element);
                if let Some(display) = display {
                    element.set_attribute("text:display", display.as_str());
                }
                set_data_style(&mut element, data_style_name.as_deref());
                element.set_text(display_text);
            },
            Self::VariableGet {
                name,
                display,
                data_style_name,
                display_text,
            } => {
                element.set_attribute("text:name", name);
                if let Some(display) = display {
                    element.set_attribute("text:display", display.as_str());
                }
                set_data_style(&mut element, data_style_name.as_deref());
                element.set_text(display_text);
            },
            Self::Expression {
                formula,
                value,
                display,
                data_style_name,
                display_text,
            } => {
                if let Some(formula) = formula {
                    element.set_attribute("text:formula", formula);
                }
                if let Some(value) = value {
                    value.write_attributes(&mut element);
                }
                if let Some(display) = display {
                    element.set_attribute("text:display", display.as_str());
                }
                set_data_style(&mut element, data_style_name.as_deref());
                element.set_text(display_text);
            },
            Self::VariableInput {
                name,
                description,
                value_type,
                display,
                data_style_name,
                display_text,
            } => {
                element.set_attribute("text:name", name);
                if let Some(description) = description {
                    element.set_attribute("text:description", description);
                }
                element.set_attribute("xmlns:office", OFFICE_NAMESPACE);
                element.set_attribute("office:value-type", value_type.as_str());
                if let Some(display) = display {
                    element.set_attribute("text:display", display.as_str());
                }
                set_data_style(&mut element, data_style_name.as_deref());
                element.set_text(display_text);
            },
            Self::UserFieldGet {
                name,
                display,
                data_style_name,
                display_text,
            } => {
                element.set_attribute("text:name", name);
                if let Some(display) = display {
                    element.set_attribute("text:display", display.as_str());
                }
                set_data_style(&mut element, data_style_name.as_deref());
                element.set_text(display_text);
            },
            Self::UserFieldInput {
                name,
                description,
                data_style_name,
                display_text,
            } => {
                element.set_attribute("text:name", name);
                if let Some(description) = description {
                    element.set_attribute("text:description", description);
                }
                set_data_style(&mut element, data_style_name.as_deref());
                element.set_text(display_text);
            },
            Self::TextInput {
                description,
                display_text,
            } => {
                if let Some(description) = description {
                    element.set_attribute("text:description", description);
                }
                element.set_text(display_text);
            },
            Self::TableFormula {
                formula,
                display,
                data_style_name,
                display_text,
            } => {
                if let Some(formula) = formula {
                    element.set_attribute("text:formula", formula);
                }
                if let Some(display) = display {
                    element.set_attribute("text:display", display.as_str());
                }
                set_data_style(&mut element, data_style_name.as_deref());
                element.set_text(display_text);
            },
            Self::Measure { kind, display_text } => {
                element.set_attribute("text:kind", kind.as_str());
                element.set_text(display_text);
            },
            Self::Reference {
                reference_name,
                reference_format,
                display_text,
            }
            | Self::BookmarkReference {
                reference_name,
                reference_format,
                display_text,
            } => {
                if let Some(reference_name) = reference_name {
                    element.set_attribute("text:ref-name", reference_name);
                }
                if let Some(reference_format) = reference_format {
                    element.set_attribute("text:reference-format", reference_format.as_str());
                }
                element.set_text(display_text);
            },
            Self::NoteReference {
                reference_name,
                note_class,
                reference_format,
                display_text,
            } => {
                if let Some(reference_name) = reference_name {
                    element.set_attribute("text:ref-name", reference_name);
                }
                element.set_attribute("text:note-class", note_class.as_str());
                if let Some(reference_format) = reference_format {
                    element.set_attribute("text:reference-format", reference_format.as_str());
                }
                element.set_text(display_text);
            },
            Self::DocumentStatistic {
                number_format,
                display_text,
                ..
            } => {
                if let Some(number_format) = number_format {
                    element.set_attribute("xmlns:style", STYLE_NAMESPACE);
                    element.set_attribute("style:num-format", number_format.format());
                    if let Some(letter_sync) = number_format.letter_sync() {
                        element.set_attribute(
                            "style:num-letter-sync",
                            if letter_sync { "true" } else { "false" },
                        );
                    }
                }
                element.set_text(display_text);
            },
            Self::PageNumber {
                number_format,
                fixed,
                page_adjust,
                select_page,
                display_text,
            } => {
                if let Some(number_format) = number_format {
                    element.set_attribute("xmlns:style", STYLE_NAMESPACE);
                    element.set_attribute("style:num-format", number_format.format());
                    if let Some(letter_sync) = number_format.letter_sync() {
                        element.set_attribute(
                            "style:num-letter-sync",
                            if letter_sync { "true" } else { "false" },
                        );
                    }
                }
                if let Some(fixed) = fixed {
                    element.set_attribute("text:fixed", if *fixed { "true" } else { "false" });
                }
                if let Some(page_adjust) = page_adjust {
                    element.set_attribute("text:page-adjust", &page_adjust.to_string());
                }
                if let Some(select_page) = select_page {
                    element.set_attribute("text:select-page", select_page.as_str());
                }
                element.set_text(display_text);
            },
            Self::Date {
                value,
                adjustment,
                fixed,
                data_style_name,
                display_text,
            } => {
                if let Some(value) = value {
                    element.set_attribute("text:date-value", value.as_str());
                }
                if let Some(adjustment) = adjustment {
                    element.set_attribute("text:date-adjust", adjustment.as_str());
                }
                if let Some(fixed) = fixed {
                    element.set_attribute("text:fixed", if *fixed { "true" } else { "false" });
                }
                set_data_style(&mut element, data_style_name.as_deref());
                element.set_text(display_text);
            },
            Self::Time {
                value,
                adjustment,
                fixed,
                data_style_name,
                display_text,
            } => {
                if let Some(value) = value {
                    element.set_attribute("text:time-value", value.as_str());
                }
                if let Some(adjustment) = adjustment {
                    element.set_attribute("text:time-adjust", adjustment.as_str());
                }
                if let Some(fixed) = fixed {
                    element.set_attribute("text:fixed", if *fixed { "true" } else { "false" });
                }
                set_data_style(&mut element, data_style_name.as_deref());
                element.set_text(display_text);
            },
            Self::PageContinuation {
                select_page,
                string_value,
                display_text,
            } => {
                element.set_attribute("text:select-page", select_page.as_str());
                if let Some(string_value) = string_value {
                    element.set_attribute("text:string-value", string_value);
                }
                element.set_text(display_text);
            },
            Self::PageVariableSet {
                active,
                page_adjust,
                display_text,
            } => {
                if let Some(active) = active {
                    element.set_attribute("text:active", if *active { "true" } else { "false" });
                }
                if let Some(page_adjust) = page_adjust {
                    element.set_attribute("text:page-adjust", &page_adjust.to_string());
                }
                element.set_text(display_text);
            },
            Self::PageVariableGet {
                number_format,
                display_text,
            } => {
                if let Some(number_format) = number_format {
                    element.set_attribute("xmlns:style", STYLE_NAMESPACE);
                    element.set_attribute("style:num-format", number_format.format());
                    if let Some(letter_sync) = number_format.letter_sync() {
                        element.set_attribute(
                            "style:num-letter-sync",
                            if letter_sync { "true" } else { "false" },
                        );
                    }
                }
                element.set_text(display_text);
            },
            Self::DocumentMetadata {
                kind,
                value,
                fixed,
                data_style_name,
                display_text,
            } => {
                if let Some(value) = value {
                    match value {
                        OdfDocumentMetadataFieldValue::Date(value) => {
                            element.set_attribute("text:date-value", value.as_str());
                        },
                        OdfDocumentMetadataFieldValue::Time(value) => {
                            element.set_attribute("text:time-value", value.as_str());
                        },
                        OdfDocumentMetadataFieldValue::Duration(value) => {
                            element.set_attribute("text:duration", value.as_str());
                        },
                    }
                }
                if let Some(fixed) = fixed {
                    element.set_attribute("text:fixed", if *fixed { "true" } else { "false" });
                }
                if kind.permits_data_style() {
                    set_data_style(&mut element, data_style_name.as_deref());
                }
                element.set_text(display_text);
            },
            Self::DocumentIdentity {
                fixed,
                display_text,
                ..
            } => {
                if let Some(fixed) = fixed {
                    element.set_attribute("text:fixed", if *fixed { "true" } else { "false" });
                }
                element.set_text(display_text);
            },
            Self::UserDefinedMetadata {
                name,
                values,
                fixed,
                data_style_name,
                display_text,
            } => {
                element.set_attribute("text:name", name);
                values.write_attributes(&mut element);
                if let Some(fixed) = fixed {
                    element.set_attribute("text:fixed", if *fixed { "true" } else { "false" });
                }
                set_data_style(&mut element, data_style_name.as_deref());
                element.set_text(display_text);
            },
            Self::MetaField { .. } => unreachable!("meta-field uses ordered mixed serializer"),
        }
        Ok(element)
    }
}

fn validate_dynamic_value(
    name: &str,
    value: Option<&str>,
    required_non_empty: bool,
    aggregate: &mut usize,
) -> Result<()> {
    let Some(value) = value else {
        return Ok(());
    };
    if required_non_empty && value.trim().is_empty() {
        return Err(Error::InvalidFormat(format!("{name} must not be empty")));
    }
    if value.len() > MAX_DYNAMIC_FIELD_VALUE {
        return Err(Error::InvalidFormat(format!(
            "{name} exceeds {MAX_DYNAMIC_FIELD_VALUE} bytes"
        )));
    }
    if !value.chars().all(is_xml_1_0_char) {
        return Err(Error::InvalidFormat(format!(
            "{name} contains a character forbidden by XML 1.0"
        )));
    }
    *aggregate = aggregate
        .checked_add(value.len())
        .ok_or_else(|| Error::InvalidFormat("dynamic field size overflow".to_string()))?;
    if *aggregate > MAX_DYNAMIC_FIELD_AGGREGATE {
        return Err(Error::InvalidFormat(format!(
            "dynamic field exceeds {MAX_DYNAMIC_FIELD_AGGREGATE} aggregate bytes"
        )));
    }
    Ok(())
}

const fn is_xml_1_0_char(value: char) -> bool {
    matches!(value, '\u{9}' | '\u{A}' | '\u{D}')
        || (value as u32 >= 0x20 && value as u32 <= 0xD7FF)
        || (value as u32 >= 0xE000 && value as u32 <= 0xFFFD)
        || (value as u32 >= 0x10000 && value as u32 <= 0x10FFFF)
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
                | "text:meta-field"
                | "text:dde-connection"
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

    /// Convert a conditional-content field to its strict typed representation.
    ///
    /// Returns `Ok(None)` for other field kinds. Conditions remain inert strings.
    pub fn dynamic_text_field(&self) -> Result<Option<OdfDynamicTextField>> {
        let text = || self.value();
        let result = match self.field_type() {
            "text:placeholder" => OdfDynamicTextField::Placeholder {
                placeholder_type: OdfPlaceholderType::parse(required_field_attribute(
                    self,
                    "text:placeholder-type",
                )?)?,
                description: self
                    .element
                    .get_attribute("text:description")
                    .map(str::to_owned),
                display_text: text(),
            },
            "text:conditional-text" => OdfDynamicTextField::ConditionalText {
                condition: required_field_attribute(self, "text:condition")?.to_owned(),
                value_if_true: required_field_attribute(self, "text:string-value-if-true")?
                    .to_owned(),
                value_if_false: required_field_attribute(self, "text:string-value-if-false")?
                    .to_owned(),
                current_value: optional_field_bool(self, "text:current-value")?,
                display_text: text(),
            },
            "text:hidden-text" => OdfDynamicTextField::HiddenText {
                condition: required_field_attribute(self, "text:condition")?.to_owned(),
                string_value: required_field_attribute(self, "text:string-value")?.to_owned(),
                is_hidden: optional_field_bool(self, "text:is-hidden")?,
                display_text: text(),
            },
            "text:hidden-paragraph" => OdfDynamicTextField::HiddenParagraph {
                condition: required_field_attribute(self, "text:condition")?.to_owned(),
                is_hidden: optional_field_bool(self, "text:is-hidden")?,
                display_text: text(),
            },
            "text:dde-connection" => OdfDynamicTextField::DdeConnection {
                connection_name: required_field_attribute(self, "text:connection-name")?.to_owned(),
                display_text: text(),
            },
            "text:sequence" => {
                let format = self.element.get_attribute("style:num-format");
                let letter_sync = optional_field_bool(self, "style:num-letter-sync")?;
                let number_format = match (format, letter_sync) {
                    (Some(format), letter_sync) => {
                        Some(OdfSequenceNumberFormat::new(format, letter_sync)?)
                    },
                    (None, Some(_)) => {
                        return Err(Error::InvalidFormat(
                            "style:num-letter-sync requires style:num-format".to_string(),
                        ));
                    },
                    (None, None) => None,
                };
                OdfDynamicTextField::Sequence {
                    name: required_field_attribute(self, "text:name")?.to_owned(),
                    formula: self
                        .element
                        .get_attribute("text:formula")
                        .map(str::to_owned),
                    number_format,
                    reference_name: self
                        .element
                        .get_attribute("text:ref-name")
                        .map(str::to_owned),
                    display_text: text(),
                }
            },
            "text:sequence-ref" => OdfDynamicTextField::SequenceReference {
                reference_name: required_field_attribute(self, "text:ref-name")?.to_owned(),
                reference_format: self
                    .element
                    .get_attribute("text:reference-format")
                    .map(OdfSequenceReferenceFormat::parse)
                    .transpose()?,
                display_text: text(),
            },
            "text:variable-set" => OdfDynamicTextField::VariableSet {
                name: required_field_attribute(self, "text:name")?.to_owned(),
                formula: self
                    .element
                    .get_attribute("text:formula")
                    .map(str::to_owned),
                value: parse_calculated_value(self, true)?.expect("required calculated value"),
                display: self
                    .element
                    .get_attribute("text:display")
                    .map(OdfVariableSetDisplay::parse)
                    .transpose()?,
                data_style_name: self
                    .element
                    .get_attribute("style:data-style-name")
                    .map(str::to_owned),
                display_text: text(),
            },
            "text:variable-get" => {
                reject_calculated_value_attributes(self)?;
                OdfDynamicTextField::VariableGet {
                    name: required_field_attribute(self, "text:name")?.to_owned(),
                    display: self
                        .element
                        .get_attribute("text:display")
                        .map(OdfFormulaFieldDisplay::parse)
                        .transpose()?,
                    data_style_name: self
                        .element
                        .get_attribute("style:data-style-name")
                        .map(str::to_owned),
                    display_text: text(),
                }
            },
            "text:expression" => OdfDynamicTextField::Expression {
                formula: self
                    .element
                    .get_attribute("text:formula")
                    .map(str::to_owned),
                value: parse_calculated_value(self, false)?,
                display: self
                    .element
                    .get_attribute("text:display")
                    .map(OdfFormulaFieldDisplay::parse)
                    .transpose()?,
                data_style_name: self
                    .element
                    .get_attribute("style:data-style-name")
                    .map(str::to_owned),
                display_text: text(),
            },
            "text:variable-input" => OdfDynamicTextField::VariableInput {
                name: required_field_attribute(self, "text:name")?.to_owned(),
                description: self
                    .element
                    .get_attribute("text:description")
                    .map(str::to_owned),
                value_type: parse_value_type_only(self)?,
                display: self
                    .element
                    .get_attribute("text:display")
                    .map(OdfVariableSetDisplay::parse)
                    .transpose()?,
                data_style_name: self
                    .element
                    .get_attribute("style:data-style-name")
                    .map(str::to_owned),
                display_text: text(),
            },
            "text:user-field-get" => {
                reject_calculated_value_attributes(self)?;
                OdfDynamicTextField::UserFieldGet {
                    name: required_field_attribute(self, "text:name")?.to_owned(),
                    display: self
                        .element
                        .get_attribute("text:display")
                        .map(OdfUserFieldDisplay::parse)
                        .transpose()?,
                    data_style_name: self
                        .element
                        .get_attribute("style:data-style-name")
                        .map(str::to_owned),
                    display_text: text(),
                }
            },
            "text:user-field-input" => {
                reject_calculated_value_attributes(self)?;
                OdfDynamicTextField::UserFieldInput {
                    name: required_field_attribute(self, "text:name")?.to_owned(),
                    description: self
                        .element
                        .get_attribute("text:description")
                        .map(str::to_owned),
                    data_style_name: self
                        .element
                        .get_attribute("style:data-style-name")
                        .map(str::to_owned),
                    display_text: text(),
                }
            },
            "text:text-input" => {
                reject_calculated_value_attributes(self)?;
                OdfDynamicTextField::TextInput {
                    description: self
                        .element
                        .get_attribute("text:description")
                        .map(str::to_owned),
                    display_text: text(),
                }
            },
            "text:table-formula" => {
                reject_calculated_value_attributes(self)?;
                OdfDynamicTextField::TableFormula {
                    formula: self
                        .element
                        .get_attribute("text:formula")
                        .map(str::to_owned),
                    display: self
                        .element
                        .get_attribute("text:display")
                        .map(OdfFormulaFieldDisplay::parse)
                        .transpose()?,
                    data_style_name: self
                        .element
                        .get_attribute("style:data-style-name")
                        .map(str::to_owned),
                    display_text: text(),
                }
            },
            "text:measure" => {
                reject_unknown_field_attributes(self, &["text:kind"])?;
                OdfDynamicTextField::Measure {
                    kind: OdfMeasureKind::parse(required_field_attribute(self, "text:kind")?)?,
                    display_text: text(),
                }
            },
            "text:reference-ref" => {
                reject_unknown_field_attributes(self, &["text:ref-name", "text:reference-format"])?;
                OdfDynamicTextField::Reference {
                    reference_name: self
                        .element
                        .get_attribute("text:ref-name")
                        .map(str::to_owned),
                    reference_format: self
                        .element
                        .get_attribute("text:reference-format")
                        .map(OdfCrossReferenceFormat::parse)
                        .transpose()?,
                    display_text: text(),
                }
            },
            "text:bookmark-ref" => {
                reject_unknown_field_attributes(self, &["text:ref-name", "text:reference-format"])?;
                OdfDynamicTextField::BookmarkReference {
                    reference_name: self
                        .element
                        .get_attribute("text:ref-name")
                        .map(str::to_owned),
                    reference_format: self
                        .element
                        .get_attribute("text:reference-format")
                        .map(OdfCrossReferenceFormat::parse)
                        .transpose()?,
                    display_text: text(),
                }
            },
            "text:note-ref" => {
                reject_unknown_field_attributes(
                    self,
                    &["text:ref-name", "text:reference-format", "text:note-class"],
                )?;
                OdfDynamicTextField::NoteReference {
                    reference_name: self
                        .element
                        .get_attribute("text:ref-name")
                        .map(str::to_owned),
                    note_class: OdfNoteReferenceClass::parse(required_field_attribute(
                        self,
                        "text:note-class",
                    )?)?,
                    reference_format: self
                        .element
                        .get_attribute("text:reference-format")
                        .map(OdfNoteReferenceFormat::parse)
                        .transpose()?,
                    display_text: text(),
                }
            },
            "text:page-count"
            | "text:paragraph-count"
            | "text:word-count"
            | "text:character-count"
            | "text:table-count"
            | "text:image-count"
            | "text:object-count" => {
                reject_unknown_field_attributes(
                    self,
                    &["style:num-format", "style:num-letter-sync"],
                )?;
                let kind = match self.field_type() {
                    "text:page-count" => OdfDocumentStatisticKind::Page,
                    "text:paragraph-count" => OdfDocumentStatisticKind::Paragraph,
                    "text:word-count" => OdfDocumentStatisticKind::Word,
                    "text:character-count" => OdfDocumentStatisticKind::Character,
                    "text:table-count" => OdfDocumentStatisticKind::Table,
                    "text:image-count" => OdfDocumentStatisticKind::Image,
                    "text:object-count" => OdfDocumentStatisticKind::Object,
                    _ => unreachable!(),
                };
                OdfDynamicTextField::DocumentStatistic {
                    kind,
                    number_format: parse_common_number_format(self)?,
                    display_text: text(),
                }
            },
            "text:page-number" => {
                reject_unknown_field_attributes(
                    self,
                    &[
                        "style:num-format",
                        "style:num-letter-sync",
                        "text:fixed",
                        "text:page-adjust",
                        "text:select-page",
                    ],
                )?;
                OdfDynamicTextField::PageNumber {
                    number_format: parse_common_number_format(self)?,
                    fixed: optional_field_bool(self, "text:fixed")?,
                    page_adjust: self
                        .element
                        .get_attribute("text:page-adjust")
                        .map(|value| {
                            value.parse::<i64>().map_err(|_| {
                                Error::InvalidFormat(format!(
                                    "invalid text:page-adjust integer '{value}'"
                                ))
                            })
                        })
                        .transpose()?,
                    select_page: self
                        .element
                        .get_attribute("text:select-page")
                        .map(OdfPageSelection::parse)
                        .transpose()?,
                    display_text: text(),
                }
            },
            "text:date" => {
                reject_unknown_field_attributes(
                    self,
                    &[
                        "text:date-value",
                        "text:date-adjust",
                        "text:fixed",
                        "style:data-style-name",
                    ],
                )?;
                OdfDynamicTextField::Date {
                    value: self
                        .element
                        .get_attribute("text:date-value")
                        .map(OdfFieldDateValue::new)
                        .transpose()?,
                    adjustment: self
                        .element
                        .get_attribute("text:date-adjust")
                        .map(OdfFieldDuration::new)
                        .transpose()?,
                    fixed: optional_field_bool(self, "text:fixed")?,
                    data_style_name: self
                        .element
                        .get_attribute("style:data-style-name")
                        .map(str::to_owned),
                    display_text: text(),
                }
            },
            "text:time" => {
                reject_unknown_field_attributes(
                    self,
                    &[
                        "text:time-value",
                        "text:time-adjust",
                        "text:fixed",
                        "style:data-style-name",
                    ],
                )?;
                OdfDynamicTextField::Time {
                    value: self
                        .element
                        .get_attribute("text:time-value")
                        .map(OdfFieldTimeValue::new)
                        .transpose()?,
                    adjustment: self
                        .element
                        .get_attribute("text:time-adjust")
                        .map(OdfFieldDuration::new)
                        .transpose()?,
                    fixed: optional_field_bool(self, "text:fixed")?,
                    data_style_name: self
                        .element
                        .get_attribute("style:data-style-name")
                        .map(str::to_owned),
                    display_text: text(),
                }
            },
            "text:page-continuation" => {
                reject_unknown_field_attributes(self, &["text:select-page", "text:string-value"])?;
                OdfDynamicTextField::PageContinuation {
                    select_page: OdfPageContinuationSelection::parse(required_field_attribute(
                        self,
                        "text:select-page",
                    )?)?,
                    string_value: self
                        .element
                        .get_attribute("text:string-value")
                        .map(str::to_owned),
                    display_text: text(),
                }
            },
            "text:page-variable-set" => {
                reject_unknown_field_attributes(self, &["text:active", "text:page-adjust"])?;
                OdfDynamicTextField::PageVariableSet {
                    active: optional_field_bool(self, "text:active")?,
                    page_adjust: self
                        .element
                        .get_attribute("text:page-adjust")
                        .map(|value| {
                            value.parse::<i64>().map_err(|_| {
                                Error::InvalidFormat(format!(
                                    "invalid page-variable text:page-adjust integer '{value}'"
                                ))
                            })
                        })
                        .transpose()?,
                    display_text: text(),
                }
            },
            "text:page-variable-get" => {
                reject_unknown_field_attributes(
                    self,
                    &["style:num-format", "style:num-letter-sync"],
                )?;
                OdfDynamicTextField::PageVariableGet {
                    number_format: parse_common_number_format(self)?,
                    display_text: text(),
                }
            },
            "text:creation-date"
            | "text:creation-time"
            | "text:print-date"
            | "text:print-time"
            | "text:editing-cycles"
            | "text:editing-duration"
            | "text:modification-date"
            | "text:modification-time" => {
                let kind = match self.field_type() {
                    "text:creation-date" => OdfDocumentMetadataFieldKind::CreationDate,
                    "text:creation-time" => OdfDocumentMetadataFieldKind::CreationTime,
                    "text:print-date" => OdfDocumentMetadataFieldKind::PrintDate,
                    "text:print-time" => OdfDocumentMetadataFieldKind::PrintTime,
                    "text:editing-cycles" => OdfDocumentMetadataFieldKind::EditingCycles,
                    "text:editing-duration" => OdfDocumentMetadataFieldKind::EditingDuration,
                    "text:modification-date" => OdfDocumentMetadataFieldKind::ModificationDate,
                    "text:modification-time" => OdfDocumentMetadataFieldKind::ModificationTime,
                    _ => unreachable!(),
                };
                let allowed = match kind {
                    OdfDocumentMetadataFieldKind::CreationDate
                    | OdfDocumentMetadataFieldKind::PrintDate
                    | OdfDocumentMetadataFieldKind::ModificationDate => {
                        &["text:fixed", "style:data-style-name", "text:date-value"][..]
                    },
                    OdfDocumentMetadataFieldKind::CreationTime
                    | OdfDocumentMetadataFieldKind::PrintTime
                    | OdfDocumentMetadataFieldKind::ModificationTime => {
                        &["text:fixed", "style:data-style-name", "text:time-value"][..]
                    },
                    OdfDocumentMetadataFieldKind::EditingDuration => {
                        &["text:fixed", "style:data-style-name", "text:duration"][..]
                    },
                    OdfDocumentMetadataFieldKind::EditingCycles => &["text:fixed"][..],
                };
                reject_unknown_field_attributes(self, allowed)?;
                let value = match kind {
                    OdfDocumentMetadataFieldKind::CreationDate
                    | OdfDocumentMetadataFieldKind::PrintDate
                    | OdfDocumentMetadataFieldKind::ModificationDate => self
                        .element
                        .get_attribute("text:date-value")
                        .map(OdfFieldDateValue::new)
                        .transpose()?
                        .map(OdfDocumentMetadataFieldValue::Date),
                    OdfDocumentMetadataFieldKind::CreationTime
                    | OdfDocumentMetadataFieldKind::PrintTime
                    | OdfDocumentMetadataFieldKind::ModificationTime => self
                        .element
                        .get_attribute("text:time-value")
                        .map(OdfFieldTimeValue::new)
                        .transpose()?
                        .map(OdfDocumentMetadataFieldValue::Time),
                    OdfDocumentMetadataFieldKind::EditingDuration => self
                        .element
                        .get_attribute("text:duration")
                        .map(OdfFieldDuration::new)
                        .transpose()?
                        .map(OdfDocumentMetadataFieldValue::Duration),
                    OdfDocumentMetadataFieldKind::EditingCycles => None,
                };
                let result = OdfDynamicTextField::DocumentMetadata {
                    kind,
                    value,
                    fixed: optional_field_bool(self, "text:fixed")?,
                    data_style_name: self
                        .element
                        .get_attribute("style:data-style-name")
                        .map(str::to_owned),
                    display_text: text(),
                };
                result.validate()?;
                result
            },
            "text:initial-creator"
            | "text:description"
            | "text:printed-by"
            | "text:title"
            | "text:subject"
            | "text:keywords"
            | "text:creator"
            | "text:author-name"
            | "text:author-initials" => {
                reject_unknown_field_attributes(self, &["text:fixed"])?;
                let kind = match self.field_type() {
                    "text:initial-creator" => OdfDocumentIdentityFieldKind::InitialCreator,
                    "text:description" => OdfDocumentIdentityFieldKind::Description,
                    "text:printed-by" => OdfDocumentIdentityFieldKind::PrintedBy,
                    "text:title" => OdfDocumentIdentityFieldKind::Title,
                    "text:subject" => OdfDocumentIdentityFieldKind::Subject,
                    "text:keywords" => OdfDocumentIdentityFieldKind::Keywords,
                    "text:creator" => OdfDocumentIdentityFieldKind::Creator,
                    "text:author-name" => OdfDocumentIdentityFieldKind::AuthorName,
                    "text:author-initials" => OdfDocumentIdentityFieldKind::AuthorInitials,
                    _ => unreachable!(),
                };
                OdfDynamicTextField::DocumentIdentity {
                    kind,
                    fixed: optional_field_bool(self, "text:fixed")?,
                    display_text: text(),
                }
            },
            "text:user-defined" => {
                reject_unknown_field_attributes(
                    self,
                    &[
                        "text:name",
                        "text:fixed",
                        "style:data-style-name",
                        "office:value",
                        "office:date-value",
                        "office:time-value",
                        "office:boolean-value",
                        "office:string-value",
                    ],
                )?;
                let result = OdfDynamicTextField::UserDefinedMetadata {
                    name: required_field_attribute(self, "text:name")?.to_owned(),
                    values: OdfUserDefinedMetadataValues {
                        number: self
                            .element
                            .get_attribute("office:value")
                            .map(str::to_owned),
                        date: self
                            .element
                            .get_attribute("office:date-value")
                            .map(OdfFieldDateValue::new)
                            .transpose()?,
                        time: self
                            .element
                            .get_attribute("office:time-value")
                            .map(OdfFieldDuration::new)
                            .transpose()?,
                        boolean: optional_field_bool(self, "office:boolean-value")?,
                        string: self
                            .element
                            .get_attribute("office:string-value")
                            .map(str::to_owned),
                    },
                    fixed: optional_field_bool(self, "text:fixed")?,
                    data_style_name: self
                        .element
                        .get_attribute("style:data-style-name")
                        .map(str::to_owned),
                    display_text: text(),
                };
                result.validate()?;
                result
            },
            _ => return Ok(None),
        };
        Ok(Some(result))
    }
}

fn set_data_style(element: &mut Element, value: Option<&str>) {
    if let Some(value) = value {
        element.set_attribute("xmlns:style", STYLE_NAMESPACE);
        element.set_attribute("style:data-style-name", value);
    }
}

fn validate_document_metadata_value(
    kind: OdfDocumentMetadataFieldKind,
    value: Option<&OdfDocumentMetadataFieldValue>,
    aggregate: &mut usize,
) -> Result<()> {
    match (kind, value) {
        (_, None) => Ok(()),
        (
            OdfDocumentMetadataFieldKind::CreationDate,
            Some(OdfDocumentMetadataFieldValue::Date(value)),
        ) => value.validate(aggregate),
        (
            OdfDocumentMetadataFieldKind::CreationTime,
            Some(OdfDocumentMetadataFieldValue::Time(value)),
        ) => value.validate(aggregate),
        (
            OdfDocumentMetadataFieldKind::PrintDate
            | OdfDocumentMetadataFieldKind::ModificationDate,
            Some(OdfDocumentMetadataFieldValue::Date(value)),
        ) if value.kind() == OdfDateValueKind::Date => value.validate(aggregate),
        (
            OdfDocumentMetadataFieldKind::PrintTime
            | OdfDocumentMetadataFieldKind::ModificationTime,
            Some(OdfDocumentMetadataFieldValue::Time(value)),
        ) if value.kind() == OdfTimeValueKind::Time => value.validate(aggregate),
        (
            OdfDocumentMetadataFieldKind::EditingDuration,
            Some(OdfDocumentMetadataFieldValue::Duration(value)),
        ) => value.validate("text:duration", aggregate),
        _ => Err(Error::InvalidFormat(format!(
            "value type is not permitted by {}",
            kind.element_name()
        ))),
    }
}

fn validate_xml_schema_date(value: &str) -> Result<()> {
    let core = strip_xml_schema_timezone(value)?;
    validate_xml_schema_date_core(core)
        .map_err(|_| Error::InvalidFormat(format!("invalid XML Schema date '{value}'")))
}

fn validate_xml_schema_date_time(value: &str) -> Result<()> {
    let (date, time) = value
        .split_once('T')
        .ok_or_else(|| Error::InvalidFormat(format!("invalid XML Schema dateTime '{value}'")))?;
    if time.contains('T')
        || validate_xml_schema_date_core(date).is_err()
        || validate_xml_schema_time(time).is_err()
    {
        return Err(Error::InvalidFormat(format!(
            "invalid XML Schema dateTime '{value}'"
        )));
    }
    Ok(())
}

fn validate_xml_schema_time(value: &str) -> Result<()> {
    let core = strip_xml_schema_timezone(value)?;
    let mut parts = core.split(':');
    let hour = parse_two_digits(parts.next(), "hour")?;
    let minute = parse_two_digits(parts.next(), "minute")?;
    let second_lexical = parts
        .next()
        .ok_or_else(|| Error::InvalidFormat(format!("invalid XML Schema time '{value}'")))?;
    if parts.next().is_some() {
        return Err(Error::InvalidFormat(format!(
            "invalid XML Schema time '{value}'"
        )));
    }
    let (seconds, fraction) = match second_lexical.split_once('.') {
        Some((seconds, fraction))
            if !fraction.is_empty() && fraction.bytes().all(|b| b.is_ascii_digit()) =>
        {
            (seconds, Some(fraction))
        },
        Some(_) => {
            return Err(Error::InvalidFormat(format!(
                "invalid XML Schema time '{value}'"
            )));
        },
        None => (second_lexical, None),
    };
    let second = if seconds.len() == 2 && seconds.bytes().all(|b| b.is_ascii_digit()) {
        seconds.parse::<u8>().unwrap_or(u8::MAX)
    } else {
        u8::MAX
    };
    let midnight_24 = hour == 24
        && minute == 0
        && second == 0
        && fraction.is_none_or(|value| value.bytes().all(|b| b == b'0'));
    if minute > 59 || second > 59 || (hour > 23 && !midnight_24) {
        return Err(Error::InvalidFormat(format!(
            "invalid XML Schema time '{value}'"
        )));
    }
    Ok(())
}

fn validate_xml_schema_date_core(value: &str) -> Result<()> {
    let unsigned = value.strip_prefix('-').unwrap_or(value);
    let mut parts = unsigned.split('-');
    let year = parts.next().unwrap_or_default();
    let month = parse_two_digits(parts.next(), "month")?;
    let day = parse_two_digits(parts.next(), "day")?;
    if parts.next().is_some()
        || year.len() < 4
        || !year.bytes().all(|b| b.is_ascii_digit())
        || year.bytes().all(|b| b == b'0')
    {
        return Err(Error::InvalidFormat("invalid XML Schema date".to_string()));
    }
    let leap =
        decimal_mod(year, 400) == 0 || (decimal_mod(year, 4) == 0 && decimal_mod(year, 100) != 0);
    let max_day = match month {
        1 | 3 | 5 | 7 | 8 | 10 | 12 => 31,
        4 | 6 | 9 | 11 => 30,
        2 if leap => 29,
        2 => 28,
        _ => 0,
    };
    if day == 0 || day > max_day {
        return Err(Error::InvalidFormat("invalid XML Schema date".to_string()));
    }
    Ok(())
}

fn strip_xml_schema_timezone(value: &str) -> Result<&str> {
    if let Some(core) = value.strip_suffix('Z') {
        return if core.is_empty() {
            Err(Error::InvalidFormat(
                "empty XML Schema temporal value".to_string(),
            ))
        } else {
            Ok(core)
        };
    }
    let bytes = value.as_bytes();
    if bytes.len() >= 6
        && matches!(bytes[bytes.len() - 6], b'+' | b'-')
        && bytes[bytes.len() - 3] == b':'
    {
        let timezone = &value[value.len() - 5..];
        let hour = parse_two_digits(Some(&timezone[..2]), "timezone hour")?;
        let minute = parse_two_digits(Some(&timezone[3..]), "timezone minute")?;
        if hour > 14 || minute > 59 || (hour == 14 && minute != 0) {
            return Err(Error::InvalidFormat(format!(
                "invalid XML Schema timezone in '{value}'"
            )));
        }
        return Ok(&value[..value.len() - 6]);
    }
    Ok(value)
}

fn parse_two_digits(value: Option<&str>, component: &str) -> Result<u8> {
    let value = value
        .ok_or_else(|| Error::InvalidFormat(format!("missing XML Schema temporal {component}")))?;
    if value.len() != 2 || !value.bytes().all(|b| b.is_ascii_digit()) {
        return Err(Error::InvalidFormat(format!(
            "invalid XML Schema temporal {component}"
        )));
    }
    value
        .parse::<u8>()
        .map_err(|_| Error::InvalidFormat(format!("invalid XML Schema temporal {component}")))
}

fn decimal_mod(value: &str, modulus: u16) -> u16 {
    value.bytes().fold(0u16, |remainder, digit| {
        (remainder * 10 + u16::from(digit - b'0')) % modulus
    })
}

fn validate_double(value: &str) -> Result<()> {
    if matches!(value, "INF" | "-INF" | "NaN") || value.parse::<f64>().is_ok() {
        Ok(())
    } else {
        Err(Error::InvalidFormat(format!(
            "invalid XML Schema double '{value}'"
        )))
    }
}

const CALCULATED_VALUE_ATTRIBUTES: [&str; 7] = [
    "office:value",
    "office:currency",
    "office:date-value",
    "office:time-value",
    "office:boolean-value",
    "office:string-value",
    "office:value-type",
];

fn reject_calculated_value_attributes(field: &Field) -> Result<()> {
    if let Some(name) = CALCULATED_VALUE_ATTRIBUTES
        .iter()
        .find(|name| field.element.get_attribute(name).is_some())
    {
        return Err(Error::InvalidFormat(format!(
            "{} does not permit {name}",
            field.field_type()
        )));
    }
    Ok(())
}

fn parse_calculated_value(
    field: &Field,
    required: bool,
) -> Result<Option<OdfCalculatedFieldValue>> {
    let Some(value_type) = field.element.get_attribute("office:value-type") else {
        if CALCULATED_VALUE_ATTRIBUTES[..6]
            .iter()
            .any(|name| field.element.get_attribute(name).is_some())
        {
            return Err(Error::InvalidFormat(
                "cached field value attributes require office:value-type".to_string(),
            ));
        }
        return if required {
            Err(Error::InvalidFormat(format!(
                "{} requires office:value-type and its matching value",
                field.field_type()
            )))
        } else {
            Ok(None)
        };
    };
    let attr = |name| field.element.get_attribute(name);
    let required_attr = |name| {
        attr(name).ok_or_else(|| {
            Error::InvalidFormat(format!("office:value-type '{value_type}' requires {name}"))
        })
    };
    let value = match value_type {
        "float" => OdfCalculatedFieldValue::Float(required_attr("office:value")?.to_owned()),
        "percentage" => {
            OdfCalculatedFieldValue::Percentage(required_attr("office:value")?.to_owned())
        },
        "currency" => OdfCalculatedFieldValue::Currency {
            value: required_attr("office:value")?.to_owned(),
            currency: attr("office:currency").map(str::to_owned),
        },
        "date" => OdfCalculatedFieldValue::Date(required_attr("office:date-value")?.to_owned()),
        "time" => OdfCalculatedFieldValue::Time(required_attr("office:time-value")?.to_owned()),
        "boolean" => OdfCalculatedFieldValue::Boolean(
            optional_field_bool(field, "office:boolean-value")?.ok_or_else(|| {
                Error::InvalidFormat(
                    "office:value-type 'boolean' requires office:boolean-value".to_string(),
                )
            })?,
        ),
        "string" => OdfCalculatedFieldValue::String(attr("office:string-value").map(str::to_owned)),
        _ => {
            return Err(Error::InvalidFormat(format!(
                "invalid calculated field office:value-type '{value_type}'"
            )));
        },
    };
    let allowed: &[&str] = match value_type {
        "float" | "percentage" => &["office:value-type", "office:value"],
        "currency" => &["office:value-type", "office:value", "office:currency"],
        "date" => &["office:value-type", "office:date-value"],
        "time" => &["office:value-type", "office:time-value"],
        "boolean" => &["office:value-type", "office:boolean-value"],
        "string" => &["office:value-type", "office:string-value"],
        _ => unreachable!(),
    };
    if let Some(extra) = CALCULATED_VALUE_ATTRIBUTES
        .iter()
        .find(|name| !allowed.contains(name) && attr(name).is_some())
    {
        return Err(Error::InvalidFormat(format!(
            "office:value-type '{value_type}' does not permit {extra}"
        )));
    }
    let mut aggregate = 0usize;
    value.validate(&mut aggregate)?;
    Ok(Some(value))
}

fn parse_value_type_only(field: &Field) -> Result<OdfFieldValueType> {
    let value_type =
        OdfFieldValueType::parse(required_field_attribute(field, "office:value-type")?)?;
    if let Some(extra) = CALCULATED_VALUE_ATTRIBUTES[..6]
        .iter()
        .find(|name| field.element.get_attribute(name).is_some())
    {
        return Err(Error::InvalidFormat(format!(
            "text:variable-input permits office:value-type but not {extra}"
        )));
    }
    Ok(value_type)
}

fn parse_common_number_format(field: &Field) -> Result<Option<OdfSequenceNumberFormat>> {
    let format = field.element.get_attribute("style:num-format");
    let letter_sync = optional_field_bool(field, "style:num-letter-sync")?;
    match (format, letter_sync) {
        (Some(format), letter_sync) => Ok(Some(OdfSequenceNumberFormat::new(format, letter_sync)?)),
        (None, Some(_)) => Err(Error::InvalidFormat(
            "style:num-letter-sync requires style:num-format".to_string(),
        )),
        (None, None) => Ok(None),
    }
}

fn reject_unknown_field_attributes(field: &Field, allowed: &[&str]) -> Result<()> {
    if let Some(name) = field
        .element
        .attributes()
        .keys()
        .find(|name| !allowed.contains(&name.as_str()))
    {
        return Err(Error::InvalidFormat(format!(
            "{} does not permit attribute {name}",
            field.field_type()
        )));
    }
    Ok(())
}

fn required_field_attribute<'a>(field: &'a Field, name: &str) -> Result<&'a str> {
    field
        .element
        .get_attribute(name)
        .ok_or_else(|| Error::InvalidFormat(format!("{} requires {name}", field.field_type())))
}

fn optional_field_bool(field: &Field, name: &str) -> Result<Option<bool>> {
    field
        .element
        .get_attribute(name)
        .map(|value| match value {
            "true" | "1" => Ok(true),
            "false" | "0" => Ok(false),
            _ => Err(Error::InvalidFormat(format!(
                "invalid {name} boolean '{value}'"
            ))),
        })
        .transpose()
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

    /// Parse typed conditional, hidden, and placeholder fields without evaluating them.
    pub fn parse_dynamic_text_fields(xml_content: &str) -> Result<Vec<OdfDynamicTextField>> {
        let mut meta_fields = parse_meta_fields(xml_content)?.into_iter();
        let mut result = Vec::new();
        for field in Self::parse_fields(xml_content)? {
            if field.field_type() == "text:meta-field" {
                result.push(meta_fields.next().ok_or_else(|| {
                    Error::InvalidFormat("missing parsed text:meta-field".to_string())
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
        Ok(result)
    }
}

#[derive(Debug)]
struct ActiveMetaField {
    depth: usize,
    order: usize,
    xml_id: String,
    data_style_name: Option<String>,
    builder: MetaContentBuilder,
}

#[derive(Debug)]
struct ActiveNoteBody {
    depth: usize,
    order: usize,
    builder: MetaContentBuilder,
}

#[derive(Debug)]
struct MetaContentBuilder {
    roots: Vec<OdfMetaFieldNode>,
    stack: Vec<OdfMetaFieldElement>,
    nodes: usize,
    aggregate: usize,
    root_grammar: MetaContentGrammar,
    root_name: &'static str,
}

impl Default for MetaContentBuilder {
    fn default() -> Self {
        Self::new(MetaContentGrammar::ParagraphOrHyperlink, "text:meta-field")
    }
}

impl MetaContentBuilder {
    fn new(root_grammar: MetaContentGrammar, root_name: &'static str) -> Self {
        Self {
            roots: Vec::new(),
            stack: Vec::new(),
            nodes: 0,
            aggregate: 0,
            root_grammar,
            root_name,
        }
    }

    fn note_body() -> Self {
        Self::new(MetaContentGrammar::NoteBody, "text:note-body")
    }

    fn push_text(&mut self, value: &str) -> Result<()> {
        if self.stack.is_empty() && self.root_grammar == MetaContentGrammar::NoteBody {
            if value.chars().all(char::is_whitespace) {
                return Ok(());
            }
            return Err(Error::InvalidFormat(
                "text:note-body cannot contain direct character data".to_string(),
            ));
        }
        add_meta_size(&mut self.aggregate, value.len())?;
        if let Some(OdfMetaFieldNode::Text(text)) = self.current_nodes_mut().last_mut() {
            text.push_str(value);
        } else {
            self.add_node()?;
            self.current_nodes_mut()
                .push(OdfMetaFieldNode::Text(value.to_string()));
        }
        Ok(())
    }

    fn start_element(
        &mut self,
        namespace_uri: String,
        local_name: String,
        attributes: Vec<OdfMetaFieldAttribute>,
    ) -> Result<()> {
        if self.stack.is_empty()
            && meta_child_grammar(self.root_grammar, &namespace_uri, &local_name).is_err()
        {
            return Err(Error::InvalidFormat(format!(
                "{}:{local_name} is not permitted directly in {}",
                namespace_uri, self.root_name
            )));
        }
        validate_meta_element_parts(
            &namespace_uri,
            &local_name,
            &attributes,
            &mut self.aggregate,
        )?;
        if self.stack.len() >= MAX_META_FIELD_DEPTH {
            return Err(Error::InvalidFormat(format!(
                "text:meta-field content exceeds {MAX_META_FIELD_DEPTH} levels"
            )));
        }
        self.add_node()?;
        self.stack.push(OdfMetaFieldElement {
            namespace_uri,
            local_name,
            attributes,
            children: Vec::new(),
        });
        Ok(())
    }

    fn empty_element(
        &mut self,
        namespace_uri: String,
        local_name: String,
        attributes: Vec<OdfMetaFieldAttribute>,
    ) -> Result<()> {
        self.start_element(namespace_uri, local_name, attributes)?;
        self.end_element()
    }

    fn end_element(&mut self) -> Result<()> {
        let element = self.stack.pop().ok_or_else(|| {
            Error::InvalidFormat("text:meta-field content stack underflow".to_string())
        })?;
        self.current_nodes_mut()
            .push(OdfMetaFieldNode::Element(element));
        Ok(())
    }

    fn finish_meta_field(self) -> Result<OdfMetaFieldContent> {
        if !self.stack.is_empty() {
            return Err(Error::InvalidFormat(
                "incomplete text:meta-field content".to_string(),
            ));
        }
        OdfMetaFieldContent::new(self.roots)
    }

    fn finish_note_body(self) -> Result<OdfNoteBodyContent> {
        if !self.stack.is_empty() {
            return Err(Error::InvalidFormat(
                "incomplete text:note-body content".to_string(),
            ));
        }
        OdfNoteBodyContent::new(self.roots)
    }

    fn current_nodes_mut(&mut self) -> &mut Vec<OdfMetaFieldNode> {
        if let Some(element) = self.stack.last_mut() {
            &mut element.children
        } else {
            &mut self.roots
        }
    }

    fn add_node(&mut self) -> Result<()> {
        self.nodes = self.nodes.checked_add(1).ok_or_else(|| {
            Error::InvalidFormat("text:meta-field node count overflow".to_string())
        })?;
        if self.nodes > MAX_META_FIELD_NODES {
            return Err(Error::InvalidFormat(format!(
                "text:meta-field exceeds {MAX_META_FIELD_NODES} content nodes"
            )));
        }
        Ok(())
    }
}

fn parse_meta_fields(xml: &str) -> Result<Vec<OdfDynamicTextField>> {
    if xml.len() > MAX_META_FIELD_XML_BYTES {
        return Err(Error::InvalidFormat(format!(
            "field XML exceeds {MAX_META_FIELD_XML_BYTES} bytes"
        )));
    }
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut stack: Vec<(Option<String>, String)> = Vec::new();
    let mut active: Vec<ActiveMetaField> = Vec::new();
    let mut completed = Vec::new();
    let mut next_order = 0usize;
    let mut document_xml_ids = HashSet::new();

    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid meta-field XML: {error}")))?;
        match event {
            Event::Start(ref source) => {
                let namespace_uri = resolved_namespace(&namespace)?;
                collect_document_xml_id(&reader, source, &mut document_xml_ids)?;
                let local = utf8(source.local_name().as_ref(), "meta-field element name")?;
                if !active.is_empty() {
                    let attributes = parse_meta_node_attributes(&reader, source)?;
                    for field in &mut active {
                        field.builder.start_element(
                            namespace_uri.clone().ok_or_else(|| {
                                Error::InvalidFormat(
                                    "unqualified meta-field child element".to_string(),
                                )
                            })?,
                            local.clone(),
                            attributes.clone(),
                        )?;
                    }
                }
                depth = depth.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("meta-field XML depth overflow".to_string())
                })?;
                if namespace_uri.as_deref() == Some(TEXT_DATABASE_NAMESPACE)
                    && local == "meta-field"
                {
                    validate_meta_field_parent(stack.last())?;
                    let (xml_id, data_style_name) = parse_meta_root_attributes(&reader, source)?;
                    if next_order >= MAX_FIELDS {
                        return Err(Error::InvalidFormat(
                            "too many text:meta-field elements".to_string(),
                        ));
                    }
                    active.push(ActiveMetaField {
                        depth,
                        order: next_order,
                        xml_id,
                        data_style_name,
                        builder: MetaContentBuilder::default(),
                    });
                    next_order += 1;
                }
                stack.push((namespace_uri, local));
            },
            Event::Empty(ref source) => {
                let namespace_uri = resolved_namespace(&namespace)?;
                collect_document_xml_id(&reader, source, &mut document_xml_ids)?;
                let local = utf8(source.local_name().as_ref(), "meta-field element name")?;
                if !active.is_empty() {
                    let attributes = parse_meta_node_attributes(&reader, source)?;
                    for field in &mut active {
                        field.builder.empty_element(
                            namespace_uri.clone().ok_or_else(|| {
                                Error::InvalidFormat(
                                    "unqualified meta-field child element".to_string(),
                                )
                            })?,
                            local.clone(),
                            attributes.clone(),
                        )?;
                    }
                }
                if namespace_uri.as_deref() == Some(TEXT_DATABASE_NAMESPACE)
                    && local == "meta-field"
                {
                    validate_meta_field_parent(stack.last())?;
                    let (xml_id, data_style_name) = parse_meta_root_attributes(&reader, source)?;
                    completed.push((
                        next_order,
                        OdfDynamicTextField::MetaField {
                            xml_id,
                            data_style_name,
                            content: OdfMetaFieldContent::new(Vec::new())?,
                        },
                    ));
                    next_order += 1;
                }
            },
            Event::Text(ref value) => {
                let value = value
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid meta-field text: {error}"))
                    })?;
                for field in &mut active {
                    field.builder.push_text(&value)?;
                }
            },
            Event::CData(ref value) => {
                let value = value
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid meta-field CDATA: {error}"))
                    })?;
                for field in &mut active {
                    field.builder.push_text(&value)?;
                }
            },
            Event::GeneralRef(ref reference) => {
                let value = decode_reference(reference, "meta-field")?;
                for field in &mut active {
                    field.builder.push_text(&value)?;
                }
            },
            Event::End(_) => {
                for field in &mut active {
                    if field.depth < depth {
                        field.builder.end_element()?;
                    }
                }
                if active.last().is_some_and(|field| field.depth == depth) {
                    let field = active.pop().expect("checked active meta-field");
                    completed.push((
                        field.order,
                        OdfDynamicTextField::MetaField {
                            xml_id: field.xml_id,
                            data_style_name: field.data_style_name,
                            content: field.builder.finish_meta_field()?,
                        },
                    ));
                }
                stack.pop().ok_or_else(|| {
                    Error::InvalidFormat("meta-field XML stack underflow".to_string())
                })?;
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("meta-field XML depth underflow".to_string())
                })?;
            },
            Event::DocType(_) => {
                return Err(Error::InvalidFormat(
                    "DOCTYPE is not permitted in ODF field XML".to_string(),
                ));
            },
            Event::PI(_) if !active.is_empty() => {
                return Err(Error::InvalidFormat(
                    "processing instructions are not permitted in text:meta-field".to_string(),
                ));
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if depth != 0 || !stack.is_empty() || !active.is_empty() {
        return Err(Error::InvalidFormat(
            "incomplete meta-field XML".to_string(),
        ));
    }
    completed.sort_by_key(|(order, _)| *order);
    Ok(completed.into_iter().map(|(_, field)| field).collect())
}

/// Parse every direct `text:note-body` child of an ODF `text:note` into the
/// shared inert mixed-content model. This does not evaluate fields, links,
/// event listeners, scripts, or macros represented by the nodes.
pub(crate) fn parse_note_body_contents(xml: &str) -> Result<Vec<OdfNoteBodyContent>> {
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut stack: Vec<(Option<String>, String)> = Vec::new();
    let mut active: Vec<ActiveNoteBody> = Vec::new();
    let mut completed = Vec::new();
    let mut next_order = 0usize;

    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid note-body XML: {error}")))?;
        match event {
            Event::Start(ref source) => {
                let namespace_uri = resolved_namespace(&namespace)?;
                let local = utf8(source.local_name().as_ref(), "note-body element name")?;
                let is_note_body = namespace_uri.as_deref() == Some(TEXT_DATABASE_NAMESPACE)
                    && local == "note-body"
                    && stack
                        .last()
                        .is_some_and(|(parent_namespace, parent_local)| {
                            parent_namespace.as_deref() == Some(TEXT_DATABASE_NAMESPACE)
                                && parent_local == "note"
                        });
                if is_note_body {
                    validate_note_body_attributes(source)?;
                }
                if !active.is_empty() {
                    let attributes = parse_meta_node_attributes(&reader, source)?;
                    let namespace_uri = namespace_uri.clone().ok_or_else(|| {
                        Error::InvalidFormat("unqualified note-body child element".to_string())
                    })?;
                    for body in &mut active {
                        body.builder.start_element(
                            namespace_uri.clone(),
                            local.clone(),
                            attributes.clone(),
                        )?;
                    }
                }
                depth = depth.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("note-body XML depth overflow".to_string())
                })?;
                if depth > MAX_FIELD_DEPTH {
                    return Err(Error::InvalidFormat(format!(
                        "note-body XML exceeds {MAX_FIELD_DEPTH} levels"
                    )));
                }
                if is_note_body {
                    if next_order >= MAX_FIELDS {
                        return Err(Error::InvalidFormat(
                            "document exceeds note-body limit".to_string(),
                        ));
                    }
                    active.push(ActiveNoteBody {
                        depth,
                        order: next_order,
                        builder: MetaContentBuilder::note_body(),
                    });
                    next_order += 1;
                }
                stack.push((namespace_uri, local));
            },
            Event::Empty(ref source) => {
                let namespace_uri = resolved_namespace(&namespace)?;
                let local = utf8(source.local_name().as_ref(), "note-body element name")?;
                let is_note_body = namespace_uri.as_deref() == Some(TEXT_DATABASE_NAMESPACE)
                    && local == "note-body"
                    && stack
                        .last()
                        .is_some_and(|(parent_namespace, parent_local)| {
                            parent_namespace.as_deref() == Some(TEXT_DATABASE_NAMESPACE)
                                && parent_local == "note"
                        });
                if is_note_body {
                    validate_note_body_attributes(source)?;
                }
                if !active.is_empty() {
                    let attributes = parse_meta_node_attributes(&reader, source)?;
                    let namespace_uri = namespace_uri.clone().ok_or_else(|| {
                        Error::InvalidFormat("unqualified note-body child element".to_string())
                    })?;
                    for body in &mut active {
                        body.builder.empty_element(
                            namespace_uri.clone(),
                            local.clone(),
                            attributes.clone(),
                        )?;
                    }
                }
                if is_note_body {
                    if next_order >= MAX_FIELDS {
                        return Err(Error::InvalidFormat(
                            "document exceeds note-body limit".to_string(),
                        ));
                    }
                    completed.push((next_order, OdfNoteBodyContent::new(Vec::new())?));
                    next_order += 1;
                }
            },
            Event::Text(ref value) if !active.is_empty() => {
                let value = value
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid note-body text: {error}"))
                    })?;
                for body in &mut active {
                    body.builder.push_text(&value)?;
                }
            },
            Event::CData(ref value) if !active.is_empty() => {
                let value = value
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid note-body CDATA: {error}"))
                    })?;
                for body in &mut active {
                    body.builder.push_text(&value)?;
                }
            },
            Event::GeneralRef(ref reference) if !active.is_empty() => {
                let value = decode_reference(reference, "note-body")?;
                for body in &mut active {
                    body.builder.push_text(&value)?;
                }
            },
            Event::End(_) => {
                for body in &mut active {
                    if body.depth < depth {
                        body.builder.end_element()?;
                    }
                }
                if active.last().is_some_and(|body| body.depth == depth) {
                    let body = active.pop().expect("checked active note body");
                    completed.push((body.order, body.builder.finish_note_body()?));
                }
                stack.pop().ok_or_else(|| {
                    Error::InvalidFormat("note-body XML stack underflow".to_string())
                })?;
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("note-body XML depth underflow".to_string())
                })?;
            },
            Event::DocType(_) => {
                return Err(Error::InvalidFormat(
                    "DOCTYPE is not permitted in ODF note-body XML".to_string(),
                ));
            },
            Event::PI(_) if !active.is_empty() => {
                return Err(Error::InvalidFormat(
                    "processing instructions are not permitted in text:note-body".to_string(),
                ));
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if depth != 0 || !stack.is_empty() || !active.is_empty() {
        return Err(Error::InvalidFormat("incomplete note-body XML".to_string()));
    }
    completed.sort_by_key(|(order, _)| *order);
    Ok(completed.into_iter().map(|(_, body)| body).collect())
}

fn validate_note_body_attributes(source: &quick_xml::events::BytesStart<'_>) -> Result<()> {
    for attribute in source.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| {
            Error::InvalidFormat(format!("invalid text:note-body attribute: {error}"))
        })?;
        let raw = attribute.key.as_ref();
        if raw == b"xmlns" || raw.starts_with(b"xmlns:") {
            continue;
        }
        return Err(Error::InvalidFormat(
            "text:note-body does not permit attributes".to_string(),
        ));
    }
    Ok(())
}

fn collect_document_xml_id(
    reader: &NsReader<&[u8]>,
    source: &quick_xml::events::BytesStart<'_>,
    ids: &mut HashSet<String>,
) -> Result<()> {
    for attribute in source.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| {
            Error::InvalidFormat(format!(
                "invalid XML attribute while collecting xml:id: {error}"
            ))
        })?;
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        if resolved_namespace(&namespace)?.as_deref() != Some(XML_NAMESPACE)
            || local.as_ref() != b"id"
        {
            continue;
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
            .map_err(|error| Error::InvalidFormat(format!("invalid xml:id value: {error}")))?
            .into_owned();
        validate_xml_id(&value)?;
        if !ids.insert(value.clone()) {
            return Err(Error::InvalidFormat(format!(
                "duplicate document xml:id '{value}'"
            )));
        }
    }
    Ok(())
}

fn parse_meta_root_attributes(
    reader: &NsReader<&[u8]>,
    source: &quick_xml::events::BytesStart<'_>,
) -> Result<(String, Option<String>)> {
    let mut xml_id = None;
    let mut data_style_name = None;
    for attribute in source.attributes() {
        let attribute = attribute.map_err(|error| {
            Error::InvalidFormat(format!("invalid meta-field attribute: {error}"))
        })?;
        let raw = attribute.key.as_ref();
        if raw == b"xmlns" || raw.starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let namespace = resolved_namespace(&namespace)?;
        let local = utf8(local.as_ref(), "meta-field attribute name")?;
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid meta-field attribute: {error}"))
            })?
            .into_owned();
        match (namespace.as_deref(), local.as_str()) {
            (Some(XML_NAMESPACE), "id") => xml_id = Some(value),
            (Some(STYLE_NAMESPACE), "data-style-name") => data_style_name = Some(value),
            _ => {
                return Err(Error::InvalidFormat(format!(
                    "unexpected text:meta-field attribute {}:{local}",
                    namespace.as_deref().unwrap_or("unqualified")
                )));
            },
        }
    }
    let xml_id = xml_id
        .ok_or_else(|| Error::InvalidFormat("text:meta-field requires xml:id".to_string()))?;
    validate_xml_id(&xml_id)?;
    Ok((xml_id, data_style_name))
}

fn parse_meta_node_attributes(
    reader: &NsReader<&[u8]>,
    source: &quick_xml::events::BytesStart<'_>,
) -> Result<Vec<OdfMetaFieldAttribute>> {
    let mut attributes = Vec::new();
    for attribute in source.attributes() {
        let attribute = attribute.map_err(|error| {
            Error::InvalidFormat(format!("invalid meta-field child attribute: {error}"))
        })?;
        let raw = attribute.key.as_ref();
        if raw == b"xmlns" || raw.starts_with(b"xmlns:") {
            continue;
        }
        if attributes.len() >= MAX_META_FIELD_ATTRIBUTES {
            return Err(Error::InvalidFormat(format!(
                "meta-field child exceeds {MAX_META_FIELD_ATTRIBUTES} attributes"
            )));
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let namespace_uri = resolved_namespace(&namespace)?.ok_or_else(|| {
            Error::InvalidFormat("unqualified meta-field child attribute".to_string())
        })?;
        if !is_allowed_meta_namespace(&namespace_uri) {
            return Err(Error::InvalidFormat(format!(
                "foreign meta-field attribute namespace '{namespace_uri}'"
            )));
        }
        attributes.push(OdfMetaFieldAttribute {
            namespace_uri,
            local_name: utf8(local.as_ref(), "meta-field attribute name")?,
            value: attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
                .map_err(|error| {
                    Error::InvalidFormat(format!("invalid meta-field attribute: {error}"))
                })?
                .into_owned(),
        });
    }
    Ok(attributes)
}

fn validate_meta_field_parent(parent: Option<&(Option<String>, String)>) -> Result<()> {
    let valid = parent.is_some_and(|(namespace, local)| {
        namespace.as_deref() == Some(TEXT_DATABASE_NAMESPACE)
            && matches!(
                local.as_str(),
                "a" | "h" | "meta" | "meta-field" | "p" | "ruby-base" | "span"
            )
    });
    if valid {
        Ok(())
    } else {
        Err(Error::InvalidFormat(
            "text:meta-field occurs outside an ODF inline-text host".to_string(),
        ))
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum MetaContentGrammar {
    ParagraphOrHyperlink,
    Paragraph,
    TextOnly,
    Empty,
    Hyperlink,
    Ruby,
    Note,
    NoteBody,
    ExecuteMacro,
    EventListeners,
    PresentationEventListener,
    Annotation,
    Structured,
    ShapeBasic,
    ShapeGroup,
    ShapeFrame,
    ShapeLink,
}

fn validate_meta_nodes(
    nodes: &[OdfMetaFieldNode],
    depth: usize,
    grammar: MetaContentGrammar,
    aggregate: &mut usize,
    node_count: &mut usize,
    display_text: &mut String,
) -> Result<()> {
    if depth > MAX_META_FIELD_DEPTH {
        return Err(Error::InvalidFormat(format!(
            "text:meta-field content exceeds {MAX_META_FIELD_DEPTH} levels"
        )));
    }
    match grammar {
        MetaContentGrammar::Ruby => {
            return validate_meta_exact_pair(
                nodes,
                depth,
                (
                    TEXT_DATABASE_NAMESPACE,
                    "ruby-base",
                    MetaContentGrammar::ParagraphOrHyperlink,
                ),
                (
                    TEXT_DATABASE_NAMESPACE,
                    "ruby-text",
                    MetaContentGrammar::TextOnly,
                ),
                "text:ruby",
                aggregate,
                node_count,
                display_text,
            );
        },
        MetaContentGrammar::Note => {
            return validate_meta_exact_pair(
                nodes,
                depth,
                (
                    TEXT_DATABASE_NAMESPACE,
                    "note-citation",
                    MetaContentGrammar::TextOnly,
                ),
                (
                    TEXT_DATABASE_NAMESPACE,
                    "note-body",
                    MetaContentGrammar::NoteBody,
                ),
                "text:note",
                aggregate,
                node_count,
                display_text,
            );
        },
        MetaContentGrammar::Hyperlink => {
            return validate_meta_optional_listener_then(
                nodes,
                depth,
                MetaContentGrammar::Paragraph,
                "text:a",
                aggregate,
                node_count,
                display_text,
            );
        },
        MetaContentGrammar::ExecuteMacro => {
            return validate_meta_optional_listener_then(
                nodes,
                depth,
                MetaContentGrammar::TextOnly,
                "text:execute-macro",
                aggregate,
                node_count,
                display_text,
            );
        },
        MetaContentGrammar::EventListeners => {
            return validate_meta_event_listeners(
                nodes,
                depth,
                aggregate,
                node_count,
                display_text,
            );
        },
        MetaContentGrammar::PresentationEventListener => {
            if nodes.is_empty() {
                return Ok(());
            }
            if nodes.len() != 1 {
                return Err(Error::InvalidFormat(
                    "presentation:event-listener permits at most one presentation:sound"
                        .to_string(),
                ));
            }
            return validate_meta_required_element(
                &nodes[0],
                depth,
                PRESENTATION_NAMESPACE,
                "sound",
                MetaContentGrammar::Empty,
                aggregate,
                node_count,
                display_text,
            );
        },
        MetaContentGrammar::Annotation => {
            return validate_meta_annotation(nodes, depth, aggregate, node_count, display_text);
        },
        MetaContentGrammar::ShapeLink => {
            if nodes.len() != 1 {
                return Err(Error::InvalidFormat(
                    "draw:a requires exactly one drawing shape".to_string(),
                ));
            }
            let OdfMetaFieldNode::Element(element) = &nodes[0] else {
                return Err(Error::InvalidFormat(
                    "draw:a requires a drawing shape child".to_string(),
                ));
            };
            let grammar = odf_shape_grammar(&element.namespace_uri, &element.local_name)
                .ok_or_else(|| {
                    Error::InvalidFormat("draw:a child is not a drawing shape".to_string())
                })?;
            return validate_meta_required_element(
                &nodes[0],
                depth,
                &element.namespace_uri,
                &element.local_name,
                grammar,
                aggregate,
                node_count,
                display_text,
            );
        },
        _ => {},
    }
    for node in nodes {
        *node_count = node_count.checked_add(1).ok_or_else(|| {
            Error::InvalidFormat("text:meta-field node count overflow".to_string())
        })?;
        if *node_count > MAX_META_FIELD_NODES {
            return Err(Error::InvalidFormat(format!(
                "text:meta-field exceeds {MAX_META_FIELD_NODES} content nodes"
            )));
        }
        match node {
            OdfMetaFieldNode::Text(value) => {
                if matches!(grammar, MetaContentGrammar::Empty) {
                    return Err(Error::InvalidFormat(
                        "ODF empty inline element contains character data".to_string(),
                    ));
                }
                validate_dynamic_value("meta-field text", Some(value), false, aggregate)?;
                display_text.push_str(value);
            },
            OdfMetaFieldNode::Element(element) => {
                validate_meta_element_parts(
                    &element.namespace_uri,
                    &element.local_name,
                    &element.attributes,
                    aggregate,
                )?;
                let child_grammar =
                    meta_child_grammar(grammar, &element.namespace_uri, &element.local_name)?;
                validate_meta_nodes(
                    &element.children,
                    depth + 1,
                    child_grammar,
                    aggregate,
                    node_count,
                    display_text,
                )?;
            },
        }
    }
    Ok(())
}

fn validate_meta_exact_pair(
    nodes: &[OdfMetaFieldNode],
    depth: usize,
    first: (&str, &str, MetaContentGrammar),
    second: (&str, &str, MetaContentGrammar),
    owner: &str,
    aggregate: &mut usize,
    node_count: &mut usize,
    display_text: &mut String,
) -> Result<()> {
    if nodes.len() != 2 {
        return Err(Error::InvalidFormat(format!(
            "{owner} requires exactly two schema-ordered child elements"
        )));
    }
    validate_meta_required_element(
        &nodes[0],
        depth,
        first.0,
        first.1,
        first.2,
        aggregate,
        node_count,
        display_text,
    )?;
    validate_meta_required_element(
        &nodes[1],
        depth,
        second.0,
        second.1,
        second.2,
        aggregate,
        node_count,
        display_text,
    )
}

fn validate_meta_required_element(
    node: &OdfMetaFieldNode,
    depth: usize,
    namespace: &str,
    local: &str,
    child_grammar: MetaContentGrammar,
    aggregate: &mut usize,
    node_count: &mut usize,
    display_text: &mut String,
) -> Result<()> {
    let OdfMetaFieldNode::Element(element) = node else {
        return Err(Error::InvalidFormat(format!(
            "expected {namespace}:{local} element in structured metadata content"
        )));
    };
    if element.namespace_uri != namespace || element.local_name != local {
        return Err(Error::InvalidFormat(format!(
            "expected {namespace}:{local} in structured metadata content"
        )));
    }
    *node_count = node_count
        .checked_add(1)
        .ok_or_else(|| Error::InvalidFormat("text:meta-field node count overflow".to_string()))?;
    if *node_count > MAX_META_FIELD_NODES {
        return Err(Error::InvalidFormat(format!(
            "text:meta-field exceeds {MAX_META_FIELD_NODES} content nodes"
        )));
    }
    validate_meta_element_parts(namespace, local, &element.attributes, aggregate)?;
    validate_meta_nodes(
        &element.children,
        depth + 1,
        child_grammar,
        aggregate,
        node_count,
        display_text,
    )
}

fn validate_meta_optional_listener_then(
    nodes: &[OdfMetaFieldNode],
    depth: usize,
    remaining_grammar: MetaContentGrammar,
    owner: &str,
    aggregate: &mut usize,
    node_count: &mut usize,
    display_text: &mut String,
) -> Result<()> {
    let listener_position = nodes.iter().position(|node| {
        matches!(node, OdfMetaFieldNode::Element(element)
            if element.namespace_uri == OFFICE_NAMESPACE && element.local_name == "event-listeners")
    });
    let start = match listener_position {
        None => 0,
        Some(0) => {
            validate_meta_required_element(
                &nodes[0],
                depth,
                OFFICE_NAMESPACE,
                "event-listeners",
                MetaContentGrammar::EventListeners,
                aggregate,
                node_count,
                display_text,
            )?;
            1
        },
        Some(_) => {
            return Err(Error::InvalidFormat(format!(
                "office:event-listeners must be the first child of {owner}"
            )));
        },
    };
    validate_meta_nodes(
        &nodes[start..],
        depth,
        remaining_grammar,
        aggregate,
        node_count,
        display_text,
    )
}

fn validate_meta_event_listeners(
    nodes: &[OdfMetaFieldNode],
    depth: usize,
    aggregate: &mut usize,
    node_count: &mut usize,
    display_text: &mut String,
) -> Result<()> {
    for node in nodes {
        let OdfMetaFieldNode::Element(element) = node else {
            return Err(Error::InvalidFormat(
                "office:event-listeners cannot contain character data".to_string(),
            ));
        };
        let child_grammar = match (element.namespace_uri.as_str(), element.local_name.as_str()) {
            (SCRIPT_NAMESPACE, "event-listener") => MetaContentGrammar::Empty,
            (PRESENTATION_NAMESPACE, "event-listener") => {
                MetaContentGrammar::PresentationEventListener
            },
            _ => {
                return Err(Error::InvalidFormat(format!(
                    "{}:{} is not an ODF event listener",
                    element.namespace_uri, element.local_name
                )));
            },
        };
        validate_meta_required_element(
            node,
            depth,
            &element.namespace_uri,
            &element.local_name,
            child_grammar,
            aggregate,
            node_count,
            display_text,
        )?;
    }
    Ok(())
}

fn validate_meta_annotation(
    nodes: &[OdfMetaFieldNode],
    depth: usize,
    aggregate: &mut usize,
    node_count: &mut usize,
    display_text: &mut String,
) -> Result<()> {
    let metadata = [
        (DC_NAMESPACE, "creator"),
        (DC_NAMESPACE, "date"),
        (META_NAMESPACE, "date-string"),
    ];
    let mut position = 0usize;
    for (namespace, local) in metadata {
        if matches!(nodes.get(position), Some(OdfMetaFieldNode::Element(element))
            if element.namespace_uri == namespace && element.local_name == local)
        {
            validate_meta_required_element(
                &nodes[position],
                depth,
                namespace,
                local,
                MetaContentGrammar::TextOnly,
                aggregate,
                node_count,
                display_text,
            )?;
            position += 1;
        }
    }
    for node in &nodes[position..] {
        let OdfMetaFieldNode::Element(element) = node else {
            return Err(Error::InvalidFormat(
                "office:annotation only permits metadata followed by text:p or text:list"
                    .to_string(),
            ));
        };
        let grammar = match (element.namespace_uri.as_str(), element.local_name.as_str()) {
            (TEXT_DATABASE_NAMESPACE, "p") => MetaContentGrammar::ParagraphOrHyperlink,
            (TEXT_DATABASE_NAMESPACE, "list") => MetaContentGrammar::Structured,
            _ => {
                return Err(Error::InvalidFormat(
                    "office:annotation only permits metadata followed by text:p or text:list"
                        .to_string(),
                ));
            },
        };
        validate_meta_required_element(
            node,
            depth,
            &element.namespace_uri,
            &element.local_name,
            grammar,
            aggregate,
            node_count,
            display_text,
        )?;
    }
    Ok(())
}

fn meta_child_grammar(
    parent: MetaContentGrammar,
    namespace: &str,
    local: &str,
) -> Result<MetaContentGrammar> {
    if matches!(
        parent,
        MetaContentGrammar::TextOnly | MetaContentGrammar::Empty
    ) {
        return Err(Error::InvalidFormat(format!(
            "{namespace}:{local} is not permitted inside a text-only or empty ODF element"
        )));
    }
    if parent == MetaContentGrammar::NoteBody {
        return note_body_child_grammar(namespace, local);
    }
    if matches!(
        parent,
        MetaContentGrammar::ShapeBasic
            | MetaContentGrammar::ShapeGroup
            | MetaContentGrammar::ShapeFrame
    ) {
        return shape_child_grammar(parent, namespace, local);
    }
    if parent == MetaContentGrammar::Structured {
        return structured_meta_child_grammar(namespace, local);
    }
    let allow_hyperlink = parent == MetaContentGrammar::ParagraphOrHyperlink;
    if namespace == TEXT_DATABASE_NAMESPACE {
        if local == "a" {
            return if allow_hyperlink {
                Ok(MetaContentGrammar::Hyperlink)
            } else {
                Err(Error::InvalidFormat(
                    "text:a cannot be nested in text:a paragraph content".to_string(),
                ))
            };
        }
        return match local {
            "span" | "meta" | "meta-field" => Ok(MetaContentGrammar::ParagraphOrHyperlink),
            "ruby" => Ok(MetaContentGrammar::Ruby),
            "note" => Ok(MetaContentGrammar::Note),
            "execute-macro" => Ok(MetaContentGrammar::ExecuteMacro),
            "s"
            | "tab"
            | "line-break"
            | "soft-page-break"
            | "bookmark"
            | "bookmark-start"
            | "bookmark-end"
            | "reference-mark"
            | "reference-mark-start"
            | "reference-mark-end"
            | "change"
            | "change-start"
            | "change-end"
            | "toc-mark"
            | "toc-mark-start"
            | "toc-mark-end"
            | "user-index-mark"
            | "user-index-mark-start"
            | "user-index-mark-end"
            | "alphabetical-index-mark"
            | "alphabetical-index-mark-start"
            | "alphabetical-index-mark-end" => Ok(MetaContentGrammar::Empty),
            "bibliography-mark" => Ok(MetaContentGrammar::TextOnly),
            "database-next" | "database-row-select" => Ok(MetaContentGrammar::Empty),
            _ if Field::is_field_tag(&format!("text:{local}")) => Ok(MetaContentGrammar::TextOnly),
            _ => Err(Error::InvalidFormat(format!(
                "text:{local} is not paragraph-content in text:meta-field"
            ))),
        };
    }
    match (namespace, local) {
        (OFFICE_NAMESPACE, "annotation") => Ok(MetaContentGrammar::Annotation),
        (OFFICE_NAMESPACE, "annotation-end") => Ok(MetaContentGrammar::Empty),
        (PRESENTATION_NAMESPACE, "header" | "footer" | "date-time") => {
            Ok(MetaContentGrammar::Empty)
        },
        _ if is_odf_shape_root(namespace, local) => {
            Ok(odf_shape_grammar(namespace, local).expect("checked shape root"))
        },
        _ => Err(Error::InvalidFormat(format!(
            "{namespace}:{local} is not paragraph-content in text:meta-field"
        ))),
    }
}

fn note_body_child_grammar(namespace: &str, local: &str) -> Result<MetaContentGrammar> {
    if namespace == TEXT_DATABASE_NAMESPACE {
        return match local {
            "p" | "h" => Ok(MetaContentGrammar::ParagraphOrHyperlink),
            "soft-page-break" | "change" | "change-start" | "change-end" => {
                Ok(MetaContentGrammar::Empty)
            },
            "list" | "numbered-paragraph" | "section" | "table-of-content"
            | "illustration-index" | "table-index" | "object-index" | "user-index"
            | "alphabetical-index" | "bibliography" => Ok(MetaContentGrammar::Structured),
            _ => Err(Error::InvalidFormat(format!(
                "text:{local} is not text-content in text:note-body"
            ))),
        };
    }
    if namespace == TABLE_NAMESPACE && local == "table" {
        return Ok(MetaContentGrammar::Structured);
    }
    if is_odf_shape_root(namespace, local) {
        return Ok(odf_shape_grammar(namespace, local).expect("checked shape root"));
    }
    Err(Error::InvalidFormat(format!(
        "{namespace}:{local} is not text-content in text:note-body"
    )))
}

fn shape_child_grammar(
    owner: MetaContentGrammar,
    namespace: &str,
    local: &str,
) -> Result<MetaContentGrammar> {
    match (namespace, local) {
        (SVG_NAMESPACE, "title" | "desc") => Ok(MetaContentGrammar::TextOnly),
        (OFFICE_NAMESPACE, "event-listeners") => Ok(MetaContentGrammar::EventListeners),
        (TEXT_DATABASE_NAMESPACE, "p") => Ok(MetaContentGrammar::ParagraphOrHyperlink),
        (DRAW_NAMESPACE, "glue-point" | "page-thumbnail" | "control") => {
            Ok(MetaContentGrammar::Empty)
        },
        (
            DRAW_NAMESPACE,
            "text-box" | "image" | "object" | "object-ole" | "applet" | "floating-frame" | "plugin"
            | "image-map" | "enhanced-geometry",
        ) if owner == MetaContentGrammar::ShapeFrame => Ok(MetaContentGrammar::Structured),
        (TABLE_NAMESPACE, "table") if owner == MetaContentGrammar::ShapeFrame => {
            Ok(MetaContentGrammar::Structured)
        },
        _ if owner == MetaContentGrammar::ShapeGroup && is_odf_shape_root(namespace, local) => {
            Ok(odf_shape_grammar(namespace, local).expect("checked shape root"))
        },
        _ => Err(Error::InvalidFormat(format!(
            "{namespace}:{local} is not valid direct drawing-shape content"
        ))),
    }
}

fn structured_meta_child_grammar(namespace: &str, local: &str) -> Result<MetaContentGrammar> {
    match (namespace, local) {
        (TEXT_DATABASE_NAMESPACE, "p" | "h") => Ok(MetaContentGrammar::ParagraphOrHyperlink),
        (TEXT_DATABASE_NAMESPACE, "soft-page-break" | "change" | "change-start" | "change-end") => {
            Ok(MetaContentGrammar::Empty)
        },
        (OFFICE_NAMESPACE, "event-listeners") => Ok(MetaContentGrammar::EventListeners),
        (SVG_NAMESPACE, "title" | "desc") => Ok(MetaContentGrammar::TextOnly),
        (SCRIPT_NAMESPACE, "event-listener") => Ok(MetaContentGrammar::Empty),
        (PRESENTATION_NAMESPACE, "event-listener") => {
            Ok(MetaContentGrammar::PresentationEventListener)
        },
        _ if is_allowed_meta_namespace(namespace) => Ok(MetaContentGrammar::Structured),
        _ => Err(Error::InvalidFormat(format!(
            "foreign structured metadata namespace '{namespace}' for {local}"
        ))),
    }
}

fn is_odf_shape_root(namespace: &str, local: &str) -> bool {
    (namespace == DRAW_NAMESPACE
        && matches!(
            local,
            "rect"
                | "line"
                | "polyline"
                | "polygon"
                | "regular-polygon"
                | "path"
                | "circle"
                | "ellipse"
                | "g"
                | "page-thumbnail"
                | "frame"
                | "measure"
                | "caption"
                | "connector"
                | "control"
                | "custom-shape"
                | "a"
        ))
        || (namespace == DR3D_NAMESPACE && local == "scene")
}

fn odf_shape_grammar(namespace: &str, local: &str) -> Option<MetaContentGrammar> {
    if !is_odf_shape_root(namespace, local) {
        return None;
    }
    Some(match (namespace, local) {
        (DRAW_NAMESPACE, "g") => MetaContentGrammar::ShapeGroup,
        (DRAW_NAMESPACE, "frame") => MetaContentGrammar::ShapeFrame,
        (DRAW_NAMESPACE, "a") => MetaContentGrammar::ShapeLink,
        _ => MetaContentGrammar::ShapeBasic,
    })
}

fn validate_meta_element_parts(
    namespace_uri: &str,
    local_name: &str,
    attributes: &[OdfMetaFieldAttribute],
    aggregate: &mut usize,
) -> Result<()> {
    if !is_allowed_meta_namespace(namespace_uri) || namespace_uri == XLINK_NAMESPACE {
        return Err(Error::InvalidFormat(format!(
            "foreign meta-field element namespace '{namespace_uri}'"
        )));
    }
    validate_xml_ncname(local_name, "meta-field element name")?;
    if attributes.len() > MAX_META_FIELD_ATTRIBUTES {
        return Err(Error::InvalidFormat(format!(
            "meta-field child exceeds {MAX_META_FIELD_ATTRIBUTES} attributes"
        )));
    }
    let mut seen = HashSet::new();
    for attribute in attributes {
        if !is_allowed_meta_namespace(&attribute.namespace_uri) {
            return Err(Error::InvalidFormat(format!(
                "foreign meta-field attribute namespace '{}'",
                attribute.namespace_uri
            )));
        }
        validate_xml_ncname(&attribute.local_name, "meta-field attribute name")?;
        if !seen.insert((&attribute.namespace_uri, &attribute.local_name)) {
            return Err(Error::InvalidFormat(
                "duplicate namespace-resolved meta-field attribute".to_string(),
            ));
        }
        validate_dynamic_value(
            "meta-field attribute value",
            Some(&attribute.value),
            false,
            aggregate,
        )?;
    }
    Ok(())
}

fn validate_xml_id(value: &str) -> Result<()> {
    validate_xml_ncname(value, "xml:id")
}

fn validate_xml_ncname(value: &str, name: &str) -> Result<()> {
    let mut chars = value.chars();
    let Some(first) = chars.next() else {
        return Err(Error::InvalidFormat(format!("{name} must not be empty")));
    };
    let start = first == '_' || first.is_alphabetic() || (first as u32) >= 0x80;
    let rest = chars.all(|ch| {
        ch == '_' || ch == '-' || ch == '.' || ch.is_alphanumeric() || (ch as u32) >= 0x80
    });
    if !start || !rest || value.contains(':') || !value.chars().all(is_xml_1_0_char) {
        return Err(Error::InvalidFormat(format!(
            "invalid XML NCName {name} '{value}'"
        )));
    }
    Ok(())
}

fn is_allowed_meta_namespace(namespace: &str) -> bool {
    matches!(
        namespace,
        TEXT_DATABASE_NAMESPACE
            | OFFICE_NAMESPACE
            | STYLE_NAMESPACE
            | XLINK_NAMESPACE
            | XML_NAMESPACE
            | DRAW_NAMESPACE
            | TABLE_NAMESPACE
            | PRESENTATION_NAMESPACE
            | SVG_NAMESPACE
            | FO_NAMESPACE
            | NUMBER_NAMESPACE
            | META_NAMESPACE
            | DC_NAMESPACE
            | XHTML_NAMESPACE
            | DR3D_NAMESPACE
            | FORM_NAMESPACE
            | SCRIPT_NAMESPACE
    )
}

fn add_meta_size(aggregate: &mut usize, amount: usize) -> Result<()> {
    if amount > MAX_DYNAMIC_FIELD_VALUE {
        return Err(Error::InvalidFormat(format!(
            "meta-field value exceeds {MAX_DYNAMIC_FIELD_VALUE} bytes"
        )));
    }
    *aggregate = aggregate
        .checked_add(amount)
        .ok_or_else(|| Error::InvalidFormat("meta-field aggregate size overflow".to_string()))?;
    if *aggregate > MAX_DYNAMIC_FIELD_AGGREGATE {
        return Err(Error::InvalidFormat(format!(
            "meta-field exceeds {MAX_DYNAMIC_FIELD_AGGREGATE} aggregate bytes"
        )));
    }
    Ok(())
}

fn canonical_meta_prefix(namespace: &str) -> &'static str {
    match namespace {
        TEXT_DATABASE_NAMESPACE => "text",
        OFFICE_NAMESPACE => "office",
        STYLE_NAMESPACE => "style",
        XLINK_NAMESPACE => "xlink",
        XML_NAMESPACE => "xml",
        DRAW_NAMESPACE => "draw",
        TABLE_NAMESPACE => "table",
        PRESENTATION_NAMESPACE => "presentation",
        SVG_NAMESPACE => "svg",
        FO_NAMESPACE => "fo",
        NUMBER_NAMESPACE => "number",
        META_NAMESPACE => "meta",
        DC_NAMESPACE => "dc",
        XHTML_NAMESPACE => "xhtml",
        DR3D_NAMESPACE => "dr3d",
        FORM_NAMESPACE => "form",
        SCRIPT_NAMESPACE => "script",
        _ => unreachable!("validated meta-field namespace"),
    }
}

fn write_meta_node(node: &OdfMetaFieldNode, output: &mut String) {
    match node {
        OdfMetaFieldNode::Text(value) => push_xml_text(output, value),
        OdfMetaFieldNode::Element(element) => {
            let prefix = canonical_meta_prefix(&element.namespace_uri);
            output.push('<');
            output.push_str(prefix);
            output.push(':');
            output.push_str(&element.local_name);
            output.push_str(" xmlns:");
            output.push_str(prefix);
            output.push_str("=\"");
            output.push_str(&element.namespace_uri);
            output.push('"');
            let mut declared = HashSet::new();
            declared.insert(prefix);
            for attribute in &element.attributes {
                let attribute_prefix = canonical_meta_prefix(&attribute.namespace_uri);
                if attribute_prefix != "xml" && declared.insert(attribute_prefix) {
                    output.push_str(" xmlns:");
                    output.push_str(attribute_prefix);
                    output.push_str("=\"");
                    output.push_str(&attribute.namespace_uri);
                    output.push('"');
                }
                output.push(' ');
                output.push_str(attribute_prefix);
                output.push(':');
                output.push_str(&attribute.local_name);
                output.push_str("=\"");
                push_xml_attribute(output, &attribute.value);
                output.push('"');
            }
            if element.children.is_empty() {
                output.push_str("/>");
            } else {
                output.push('>');
                for child in &element.children {
                    write_meta_node(child, output);
                }
                output.push_str("</");
                output.push_str(prefix);
                output.push(':');
                output.push_str(&element.local_name);
                output.push('>');
            }
        },
    }
}

fn push_xml_attribute(output: &mut String, value: &str) {
    for ch in value.chars() {
        match ch {
            '"' => output.push_str("&quot;"),
            '&' => output.push_str("&amp;"),
            '<' => output.push_str("&lt;"),
            '>' => output.push_str("&gt;"),
            _ => output.push(ch),
        }
    }
}

fn push_xml_text(output: &mut String, value: &str) {
    for ch in value.chars() {
        match ch {
            '&' => output.push_str("&amp;"),
            '<' => output.push_str("&lt;"),
            '>' => output.push_str("&gt;"),
            _ => output.push(ch),
        }
    }
}

fn parse_database_fields(xml: &str) -> Result<Vec<OdfDatabaseField>> {
    if xml.len() > MAX_META_FIELD_XML_BYTES {
        return Err(Error::InvalidFormat(
            "database field XML exceeds 64 MiB".to_string(),
        ));
    }
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut active: Option<ActiveDatabaseField> = None;
    let mut fields = Vec::new();
    let mut aggregate = 0usize;
    let mut stack: Vec<(Option<String>, String)> = Vec::new();

    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid database field XML: {error}"))
            })?;
        let namespace_uri = resolved_namespace(&namespace)?;
        match event {
            Event::Start(ref element) => {
                let local = utf8(element.local_name().as_ref(), "database field element")?;
                if let Some(field) = active.as_mut() {
                    if namespace_uri.as_deref() != Some(FORM_NAMESPACE)
                        || local != "connection-resource"
                        || depth != field.depth
                        || field.connection_depth.is_some()
                        || field.field.source.connection_resource.is_some()
                        || !field.field.display_text.is_empty()
                    {
                        return Err(Error::InvalidFormat(
                            "database fields may contain only one form:connection-resource"
                                .to_string(),
                        ));
                    }
                    field.field.source.connection_resource =
                        Some(parse_connection_resource(&reader, element, &mut aggregate)?);
                    field.connection_depth = Some(depth + 1);
                } else if namespace_uri.as_deref() == Some(TEXT_DATABASE_NAMESPACE) {
                    if let Some(kind) = database_field_kind(&local) {
                        validate_database_parent(stack.last())?;
                        if fields.len() >= MAX_FIELDS {
                            return Err(Error::InvalidFormat(format!(
                                "document exceeds {MAX_FIELDS} database fields"
                            )));
                        }
                        active = Some(ActiveDatabaseField {
                            depth: depth + 1,
                            field: parse_database_field(&reader, element, kind, &mut aggregate)?,
                            connection_depth: None,
                        });
                    }
                }
                depth = checked_field_depth(depth)?;
                stack.push((namespace_uri, local));
            },
            Event::Empty(ref element) => {
                let local = utf8(element.local_name().as_ref(), "database field element")?;
                if let Some(field) = active.as_mut() {
                    if namespace_uri.as_deref() != Some(FORM_NAMESPACE)
                        || local != "connection-resource"
                        || depth != field.depth
                        || field.field.source.connection_resource.is_some()
                        || !field.field.display_text.is_empty()
                    {
                        return Err(Error::InvalidFormat(
                            "database fields may contain only one form:connection-resource"
                                .to_string(),
                        ));
                    }
                    field.field.source.connection_resource =
                        Some(parse_connection_resource(&reader, element, &mut aggregate)?);
                } else if namespace_uri.as_deref() == Some(TEXT_DATABASE_NAMESPACE) {
                    if let Some(kind) = database_field_kind(&local) {
                        validate_database_parent(stack.last())?;
                        if fields.len() >= MAX_FIELDS {
                            return Err(Error::InvalidFormat(format!(
                                "document exceeds {MAX_FIELDS} database fields"
                            )));
                        }
                        let field = parse_database_field(&reader, element, kind, &mut aggregate)?;
                        fields.push(validate_database_field(field)?);
                    }
                }
            },
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
                stack.pop().ok_or_else(|| {
                    Error::InvalidFormat("database field XML stack underflow".to_string())
                })?;
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("database field XML depth underflow".to_string())
                })?;
            },
            Event::Text(ref text) if active.is_some() => {
                let value = text.decode().map_err(|error| {
                    Error::InvalidFormat(format!("invalid database field text: {error}"))
                })?;
                append_database_text(
                    active.as_mut().expect("checked field"),
                    &value,
                    &mut aggregate,
                )?;
            },
            Event::CData(ref text) if active.is_some() => {
                let value = text.decode().map_err(|error| {
                    Error::InvalidFormat(format!("invalid database field CDATA: {error}"))
                })?;
                append_database_text(
                    active.as_mut().expect("checked field"),
                    &value,
                    &mut aggregate,
                )?;
            },
            Event::GeneralRef(ref reference) if active.is_some() => {
                let name = std::str::from_utf8(reference.as_ref()).map_err(|_| {
                    Error::InvalidFormat("invalid database field entity reference".to_string())
                })?;
                let value = resolve_database_reference(name)?;
                append_database_text(
                    active.as_mut().expect("checked field"),
                    &value,
                    &mut aggregate,
                )?;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(Error::InvalidFormat(
                    "DTDs and processing instructions are prohibited in database field XML"
                        .to_string(),
                ));
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if depth != 0 || active.is_some() || !stack.is_empty() {
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
        .map(OdfNonNegativeInteger::new)
        .transpose()?;
    let value = database_attribute(&attributes, TEXT_DATABASE_NAMESPACE, "value")
        .map(OdfNonNegativeInteger::new)
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
    match field.kind {
        OdfDatabaseFieldKind::Display if field.column_name.is_none() => {
            return Err(Error::InvalidFormat(
                "text:database-display requires text:column-name".to_string(),
            ));
        },
        OdfDatabaseFieldKind::Next | OdfDatabaseFieldKind::RowSelect
            if !field.display_text.is_empty() =>
        {
            return Err(Error::InvalidFormat(
                "database selection fields cannot contain character data".to_string(),
            ));
        },
        _ => {},
    }
    if field.number_letter_sync.is_some()
        && !matches!(field.number_format.as_deref(), Some("a" | "A"))
    {
        return Err(Error::InvalidFormat(
            "style:num-letter-sync requires style:num-format a or A".to_string(),
        ));
    }
    Ok(field)
}

fn validate_constructed_database_field(field: &OdfDatabaseField) -> Result<()> {
    let mut aggregate = 0usize;
    for value in [
        field.source.database_name.as_deref(),
        Some(field.source.table_name.as_str()),
        field.column_name.as_deref(),
        field.condition.as_deref(),
        field.data_style_name.as_deref(),
        field.number_format.as_deref(),
        Some(field.display_text.as_str()),
        field
            .source
            .connection_resource
            .as_ref()
            .map(|resource| resource.href.as_str()),
    ]
    .into_iter()
    .flatten()
    {
        if !value.chars().all(is_xml_1_0_char) {
            return Err(Error::InvalidFormat(
                "database field contains forbidden XML characters".to_string(),
            ));
        }
        append_database_size(&mut aggregate, value.len())?;
    }
    if field
        .source
        .connection_resource
        .as_ref()
        .is_some_and(|resource| !resource.simple_link)
    {
        return Err(Error::InvalidFormat(
            "ODF form:connection-resource only supports xlink:href".to_string(),
        ));
    }
    let forbidden = match field.kind {
        OdfDatabaseFieldKind::Display => {
            field.condition.is_some()
                || field.row_number.is_some()
                || field.value.is_some()
                || field.number_format.is_some()
                || field.number_letter_sync.is_some()
        },
        OdfDatabaseFieldKind::Next => {
            field.column_name.is_some()
                || field.row_number.is_some()
                || field.value.is_some()
                || field.data_style_name.is_some()
                || field.number_format.is_some()
                || field.number_letter_sync.is_some()
        },
        OdfDatabaseFieldKind::RowSelect => {
            field.column_name.is_some()
                || field.value.is_some()
                || field.data_style_name.is_some()
                || field.number_format.is_some()
                || field.number_letter_sync.is_some()
        },
        OdfDatabaseFieldKind::RowNumber => {
            field.column_name.is_some()
                || field.condition.is_some()
                || field.row_number.is_some()
                || field.data_style_name.is_some()
        },
        OdfDatabaseFieldKind::Name => {
            field.column_name.is_some()
                || field.condition.is_some()
                || field.row_number.is_some()
                || field.value.is_some()
                || field.data_style_name.is_some()
                || field.number_format.is_some()
                || field.number_letter_sync.is_some()
        },
    };
    if forbidden {
        return Err(Error::InvalidFormat(
            "database field contains attributes from another field kind".to_string(),
        ));
    }
    Ok(())
}

fn parse_connection_resource(
    reader: &NsReader<&[u8]>,
    element: &quick_xml::events::BytesStart<'_>,
    aggregate: &mut usize,
) -> Result<OdfDatabaseConnectionResource> {
    let attributes = database_attributes(reader, element, aggregate)?;
    reject_database_attributes(&attributes, &[(XLINK_NAMESPACE, "href")])?;
    let href = required_database_attribute(&attributes, XLINK_NAMESPACE, "href")?;
    Ok(OdfDatabaseConnectionResource {
        href,
        simple_link: true,
    })
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
        let raw = attribute.key.as_ref();
        if raw == b"xmlns" || raw.starts_with(b"xmlns:") {
            continue;
        }
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
        }) {
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
    if active.field.display_text.len().saturating_add(value.len()) > MAX_DATABASE_VALUE {
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

fn parse_database_bool(value: &str) -> Result<bool> {
    match value {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        _ => Err(Error::InvalidFormat(format!(
            "invalid database field boolean '{value}'"
        ))),
    }
}

fn validate_database_parent(parent: Option<&(Option<String>, String)>) -> Result<()> {
    if parent.is_some_and(|(namespace, local)| {
        namespace.as_deref() == Some(TEXT_DATABASE_NAMESPACE)
            && matches!(
                local.as_str(),
                "a" | "h" | "meta" | "meta-field" | "p" | "ruby-base" | "span"
            )
    }) {
        Ok(())
    } else {
        Err(Error::InvalidFormat(
            "database field occurs outside an ODF inline-text host".to_string(),
        ))
    }
}

fn resolved_namespace(namespace: &quick_xml::name::ResolveResult<'_>) -> Result<Option<String>> {
    match namespace {
        quick_xml::name::ResolveResult::Bound(quick_xml::name::Namespace(value)) => {
            Ok(Some(utf8(value, "namespace URI")?))
        },
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
    let codepoint =
        if let Some(value) = name.strip_prefix("#x").or_else(|| name.strip_prefix("#X")) {
            u32::from_str_radix(value, 16)
        } else if let Some(value) = name.strip_prefix('#') {
            value.parse::<u32>()
        } else {
            return Err(Error::InvalidFormat(
                "undeclared entity in database field".to_string(),
            ));
        }
        .map_err(|_| {
            Error::InvalidFormat("invalid database field character reference".to_string())
        })?;
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
            <t:database-row-select t:table-name="People" t:row-number="42"><f:connection-resource x:href="sdbc:embedded:firebird"/></t:database-row-select>
            <t:database-row-number t:database-name="Contacts" t:table-name="People"
                t:value="42" s:num-format="a" s:num-letter-sync="false">42</t:database-row-number>
            <t:database-name t:database-name="Contacts" t:table-name="People">Contacts</t:database-name>{SUFFIX}"#
        );
        let fields = FieldParser::parse_database_fields(&xml).unwrap();
        assert_eq!(fields.len(), 5);
        assert_eq!(fields[0].kind, OdfDatabaseFieldKind::Display);
        assert_eq!(fields[0].display_text, "A&B");
        assert_eq!(
            fields[0].source.effective_table_type(),
            OdfDatabaseTableType::Query
        );
        assert_eq!(
            fields[2]
                .row_number
                .as_ref()
                .map(OdfNonNegativeInteger::as_str),
            Some("42")
        );
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
            r#"<t:database-next t:database-name="db"/>"#,
            r#"<t:database-row-select t:database-name="db" t:table-name="t" t:row-number="-1"/>"#,
            r#"<t:database-name t:database-name="db" t:table-name="t" t:table-type="view"/>"#,
            r#"<t:database-next t:database-name="db" t:table-name="t">text</t:database-next>"#,
            r#"<t:database-name t:database-name="db" t:table-name="t"><t:span>x</t:span></t:database-name>"#,
            r#"<t:database-name t:table-name="t"><f:connection-resource x:href="https://example.invalid/db"/><f:connection-resource x:href="other"/></t:database-name>"#,
            r#"<t:database-name t:table-name="t"><f:connection-resource x:href="db" x:type="simple"/></t:database-name>"#,
            r#"<t:database-row-number t:table-name="t" s:num-format="1" s:num-letter-sync="true">1</t:database-row-number>"#,
            r#"<t:database-name t:table-name="t">text<f:connection-resource x:href="db"/></t:database-name>"#,
            r#"<t:database-name t:table-name="t" xmlns:z="urn:foreign" z:extra="x"/>"#,
        ];
        for body in bodies {
            let xml = format!("{PREFIX}{body}{SUFFIX}");
            assert!(
                FieldParser::parse_database_fields(&xml).is_err(),
                "accepted {body}"
            );
        }
    }

    #[test]
    fn database_fields_roundtrip_all_kinds_and_schema_optional_values() {
        let source = || OdfDatabaseSource {
            database_name: None,
            table_name: String::new(),
            table_type: Some(OdfDatabaseTableType::Command),
            connection_resource: None,
        };
        let fields = vec![
            OdfDatabaseField {
                kind: OdfDatabaseFieldKind::Display,
                source: source(),
                column_name: Some(String::new()),
                condition: None,
                row_number: None,
                value: None,
                data_style_name: Some("N1".into()),
                number_format: None,
                number_letter_sync: None,
                display_text: "A&B".into(),
            },
            OdfDatabaseField {
                kind: OdfDatabaseFieldKind::Next,
                source: source(),
                column_name: None,
                condition: Some("of:=TRUE()".into()),
                row_number: None,
                value: None,
                data_style_name: None,
                number_format: None,
                number_letter_sync: None,
                display_text: String::new(),
            },
            OdfDatabaseField {
                kind: OdfDatabaseFieldKind::RowSelect,
                source: source(),
                column_name: None,
                condition: None,
                row_number: None,
                value: None,
                data_style_name: None,
                number_format: None,
                number_letter_sync: None,
                display_text: String::new(),
            },
            OdfDatabaseField {
                kind: OdfDatabaseFieldKind::RowNumber,
                source: source(),
                column_name: None,
                condition: None,
                row_number: None,
                value: Some(OdfNonNegativeInteger::new("7").unwrap()),
                data_style_name: None,
                number_format: Some("A".into()),
                number_letter_sync: Some(true),
                display_text: "VII".into(),
            },
            OdfDatabaseField {
                kind: OdfDatabaseFieldKind::Name,
                source: OdfDatabaseSource {
                    connection_resource: Some(OdfDatabaseConnectionResource {
                        href: "sdbc:embedded:firebird".into(),
                        simple_link: true,
                    }),
                    ..source()
                },
                column_name: None,
                condition: None,
                row_number: None,
                value: None,
                data_style_name: None,
                number_format: None,
                number_letter_sync: None,
                display_text: "db".into(),
            },
        ];
        for field in fields {
            let fragment = field.to_xml_fragment().unwrap();
            let parsed =
                FieldParser::parse_database_fields(&format!("{PREFIX}{fragment}{SUFFIX}")).unwrap();
            assert_eq!(parsed, vec![field]);
        }
        let optional = format!(
            "{PREFIX}<t:database-next t:table-name=\"\"/><t:database-row-select t:table-name=\"\"/>{SUFFIX}"
        );
        assert_eq!(
            FieldParser::parse_database_fields(&optional).unwrap().len(),
            2
        );
    }

    #[test]
    fn database_parser_ignores_spoofed_names_and_rejects_bad_placement() {
        let spoof =
            format!("{PREFIX}<z:database-display xmlns:z=\"urn:not-text\" z:any=\"x\"/>{SUFFIX}");
        assert!(
            FieldParser::parse_database_fields(&spoof)
                .unwrap()
                .is_empty()
        );
        let misplaced = format!("{PREFIX}{SUFFIX}")
            .replace("<t:p>", "<t:database-name t:table-name=\"t\"/><t:p>");
        assert!(FieldParser::parse_database_fields(&misplaced).is_err());
    }

    #[test]
    fn database_non_negative_integer_is_arbitrary_width_and_canonical() {
        let beyond_u64 = "18446744073709551616000000000000000000";
        for (lexical, canonical) in [
            (beyond_u64, beyond_u64),
            ("+00042", "42"),
            ("-000", "0"),
            ("  0000\t", "0"),
        ] {
            assert_eq!(
                OdfNonNegativeInteger::new(lexical).unwrap().as_str(),
                canonical
            );
        }
        for invalid in ["", "+", "-", "-1", "1.0", "1 2", "１２", "++1"] {
            assert!(
                OdfNonNegativeInteger::new(invalid).is_err(),
                "accepted {invalid}"
            );
        }
        let boundary = "9".repeat(MAX_DATABASE_INTEGER_DIGITS);
        assert_eq!(
            OdfNonNegativeInteger::new(&boundary).unwrap().as_str(),
            boundary
        );
        assert!(OdfNonNegativeInteger::new(&"9".repeat(MAX_DATABASE_INTEGER_DIGITS + 1)).is_err());

        let xml = format!(
            "{PREFIX}<t:database-row-select t:table-name=\"t\" t:row-number=\"+000{beyond_u64}\"/><t:database-row-number t:table-name=\"t\" t:value=\"-000\">0</t:database-row-number>{SUFFIX}"
        );
        let fields = FieldParser::parse_database_fields(&xml).unwrap();
        assert_eq!(fields[0].row_number.as_ref().unwrap().as_str(), beyond_u64);
        assert_eq!(fields[1].value.as_ref().unwrap().as_str(), "0");
        let canonical = fields
            .iter()
            .map(OdfDatabaseField::to_xml_fragment)
            .collect::<Result<Vec<_>>>()
            .unwrap()
            .join("");
        assert!(canonical.contains(&format!("text:row-number=\"{beyond_u64}\"")));
        assert!(canonical.contains("text:value=\"0\""));
        assert!(!canonical.contains("+000"));
    }
}

#[cfg(test)]
mod fixed_page_date_time_tests {
    use super::*;

    fn document(body: &str) -> String {
        format!(
            r#"<o:document-content
                xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
                xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
                xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0">
                <o:body><o:text><t:p>{body}</t:p></o:text></o:body>
            </o:document-content>"#
        )
    }

    #[test]
    fn fixed_page_date_time_round_trips_every_standard_field() {
        let fields = vec![
            OdfDynamicTextField::PageNumber {
                number_format: Some(OdfPageNumberFormat::new("A", Some(true)).unwrap()),
                fixed: Some(false),
                page_adjust: Some(-2),
                select_page: Some(OdfPageSelection::Previous),
                display_text: "IV & cached".to_string(),
            },
            OdfDynamicTextField::Date {
                value: Some(OdfFieldDateValue::new("2024-02-29Z").unwrap()),
                adjustment: Some(OdfFieldDuration::new("-P1Y2M3DT4H5M6.7S").unwrap()),
                fixed: Some(true),
                data_style_name: Some("Date & Time".to_string()),
                display_text: "29 < February".to_string(),
            },
            OdfDynamicTextField::Time {
                value: Some(OdfFieldTimeValue::new("2024-02-29T24:00:00+14:00").unwrap()),
                adjustment: Some(OdfFieldDuration::new("PT15M").unwrap()),
                fixed: Some(false),
                data_style_name: Some("Clock".to_string()),
                display_text: "midnight".to_string(),
            },
            OdfDynamicTextField::PageContinuation {
                select_page: OdfPageContinuationSelection::Next,
                string_value: Some("Continued on & next".to_string()),
                display_text: "continued <cached>".to_string(),
            },
        ];
        let body = fields
            .iter()
            .map(OdfDynamicTextField::to_xml_fragment)
            .collect::<Result<Vec<_>>>()
            .unwrap()
            .join("");
        let parsed = FieldParser::parse_dynamic_text_fields(&document(&body)).unwrap();
        assert_eq!(parsed, fields);
        assert_eq!(parsed[1].display_text(), "29 < February");
    }

    #[test]
    fn fixed_page_date_time_accepts_aliases_and_exact_temporal_lexicals() {
        let xml = document(
            r#"<t:date t:date-value="-12345-01-01+05:30" t:date-adjust="P999999999999Y" s:data-style-name="D">historic</t:date>
               <t:time t:time-value="23:59:59.123456789Z" t:time-adjust="-PT0.5S">clock</t:time>
               <t:page-number t:select-page="next" t:page-adjust="9223372036854775807" s:num-format="a" s:num-letter-sync="0">a</t:page-number>
               <t:page-continuation t:select-page="previous">back</t:page-continuation>"#,
        );
        let fields = FieldParser::parse_dynamic_text_fields(&xml).unwrap();
        assert_eq!(fields.len(), 4);
        assert!(matches!(
            &fields[0],
            OdfDynamicTextField::Date { value: Some(value), .. }
                if value.kind() == OdfDateValueKind::Date
        ));
        assert!(matches!(
            &fields[1],
            OdfDynamicTextField::Time { value: Some(value), .. }
                if value.kind() == OdfTimeValueKind::Time
        ));
    }

    #[test]
    fn fixed_page_date_time_rejects_hostile_invalid_and_extension_inputs() {
        for value in [
            "2023-02-29",
            "0000-01-01",
            "2024-01-01+14:01",
            "+2024-01-01",
        ] {
            assert!(OdfFieldDateValue::new(value).is_err(), "accepted {value}");
        }
        for value in ["24:00:01", "12:60:00", "12:00:60", "12:00:00+15:00"] {
            assert!(OdfFieldTimeValue::new(value).is_err(), "accepted {value}");
        }
        assert!(OdfFieldDuration::new("P").is_err());
        assert!(OdfFieldDateValue::new("2024-01-01\u{0}").is_err());

        let invalid = [
            r#"<t:page-number t:select-page="later">1</t:page-number>"#,
            r#"<t:page-number t:page-adjust="9223372036854775808">1</t:page-number>"#,
            r#"<t:page-number s:num-letter-sync="true">a</t:page-number>"#,
            r#"<t:date t:date-value="2024-01-01" t:extra="x">date</t:date>"#,
            r#"<t:time t:time-value="12:00:00" t:date-adjust="P1D">time</t:time>"#,
            r#"<t:page-continuation t:select-page="current">continued</t:page-continuation>"#,
            r#"<t:page-continuation>continued</t:page-continuation>"#,
        ];
        for body in invalid {
            assert!(
                FieldParser::parse_dynamic_text_fields(&document(body)).is_err(),
                "accepted {body}"
            );
        }

        let wrong_namespace = r#"<o:document-content
            xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
            xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
            xmlns:x="urn:not-style"><o:body><o:text><t:p>
            <t:date t:date-value="2024-01-01" x:data-style-name="spoof">date</t:date>
            </t:p></o:text></o:body></o:document-content>"#;
        assert!(FieldParser::parse_dynamic_text_fields(wrong_namespace).is_err());

        let extension = document(
            r#"<t:page-continuation-string t:select-page="next">extension</t:page-continuation-string>"#,
        );
        assert!(
            FieldParser::parse_dynamic_text_fields(&extension)
                .unwrap()
                .is_empty()
        );

        let oversized = OdfDynamicTextField::PageContinuation {
            select_page: OdfPageContinuationSelection::Next,
            string_value: Some("x".repeat(MAX_DYNAMIC_FIELD_VALUE + 1)),
            display_text: String::new(),
        };
        assert!(oversized.to_xml_fragment().is_err());
    }
}

#[cfg(test)]
mod page_variable_family_tests {
    use super::*;

    fn document(body: &str) -> String {
        format!(
            r#"<o:document-content
                xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
                xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
                xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0">
                <o:body><o:text><t:p>{body}</t:p></o:text></o:body>
            </o:document-content>"#
        )
    }

    #[test]
    fn page_variable_family_round_trips_both_standard_elements() {
        let fields = vec![
            OdfDynamicTextField::PageVariableSet {
                active: Some(false),
                page_adjust: Some(i64::MIN),
                display_text: "inert setter cache & <safe>".to_string(),
            },
            OdfDynamicTextField::PageVariableGet {
                number_format: Some(OdfPageVariableNumberFormat::new("A", Some(true)).unwrap()),
                display_text: "cached A & <not calculated>".to_string(),
            },
        ];
        let body = fields
            .iter()
            .map(OdfDynamicTextField::to_xml_fragment)
            .collect::<Result<Vec<_>>>()
            .unwrap()
            .join("");
        let parsed = FieldParser::parse_dynamic_text_fields(&document(&body)).unwrap();
        assert_eq!(parsed, fields);
        assert_eq!(parsed[1].display_text(), "cached A & <not calculated>");
    }

    #[test]
    fn page_variable_family_preserves_omission_and_exposes_defaults() {
        let xml = document(
            r#"<t:page-variable-set>opaque</t:page-variable-set>
               <t:page-variable-get s:num-format="a" s:num-letter-sync="0">x</t:page-variable-get>"#,
        );
        let fields = FieldParser::parse_dynamic_text_fields(&xml).unwrap();
        assert_eq!(fields.len(), 2);
        assert_eq!(fields[0].effective_page_variable_active(), Some(true));
        assert_eq!(fields[0].effective_page_variable_adjustment(), Some(0));
        assert!(matches!(
            &fields[0],
            OdfDynamicTextField::PageVariableSet {
                active: None,
                page_adjust: None,
                display_text,
            } if display_text == "opaque"
        ));
        assert_eq!(
            fields[0].to_xml_fragment().unwrap(),
            concat!(
                "<text:page-variable-set xmlns:text=\"",
                "urn:oasis:names:tc:opendocument:xmlns:text:1.0\">",
                "opaque</text:page-variable-set>"
            )
        );
    }

    #[test]
    fn page_variable_family_rejects_nonstandard_hostile_and_oversized_input() {
        let invalid = [
            r#"<t:page-variable-set t:active="TRUE"/>"#,
            r#"<t:page-variable-set t:page-adjust="9223372036854775808"/>"#,
            r#"<t:page-variable-set t:value="3"/>"#,
            r#"<t:page-variable-set t:select-page="next"/>"#,
            r#"<t:page-variable-get s:num-letter-sync="true">a</t:page-variable-get>"#,
            r#"<t:page-variable-get s:num-format="I" s:num-letter-sync="true">I</t:page-variable-get>"#,
            r#"<t:page-variable-get t:page-adjust="1">1</t:page-variable-get>"#,
        ];
        for body in invalid {
            assert!(
                FieldParser::parse_dynamic_text_fields(&document(body)).is_err(),
                "accepted {body}"
            );
        }

        let wrong_namespace = r#"<o:document-content
            xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
            xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
            xmlns:x="urn:not-style"><o:body><o:text><t:p>
            <t:page-variable-get x:num-format="1">1</t:page-variable-get>
            </t:p></o:text></o:body></o:document-content>"#;
        assert!(FieldParser::parse_dynamic_text_fields(wrong_namespace).is_err());

        let oversized = OdfDynamicTextField::PageVariableSet {
            active: None,
            page_adjust: None,
            display_text: "x".repeat(MAX_DYNAMIC_FIELD_VALUE + 1),
        };
        assert!(oversized.to_xml_fragment().is_err());
        let forbidden = OdfDynamicTextField::PageVariableGet {
            number_format: None,
            display_text: "bad\u{0}".to_string(),
        };
        assert!(forbidden.to_xml_fragment().is_err());
    }
}

#[cfg(test)]
mod document_metadata_fixed_field_tests {
    use super::*;

    fn document(body: &str) -> String {
        format!(
            r#"<o:document-content
                xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
                xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
                xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0">
                <o:body><o:text><t:p>{body}</t:p></o:text></o:body>
            </o:document-content>"#
        )
    }

    fn metadata_field(
        kind: OdfDocumentMetadataFieldKind,
        value: Option<OdfDocumentMetadataFieldValue>,
        display_text: &str,
    ) -> OdfDynamicTextField {
        OdfDynamicTextField::DocumentMetadata {
            kind,
            value,
            fixed: Some(true),
            data_style_name: kind
                .permits_data_style()
                .then(|| format!("style-{display_text}")),
            display_text: display_text.to_string(),
        }
    }

    #[test]
    fn document_metadata_fixed_fields_round_trip_all_eight_standard_elements() {
        let fields = vec![
            metadata_field(
                OdfDocumentMetadataFieldKind::CreationDate,
                Some(OdfDocumentMetadataFieldValue::Date(
                    OdfFieldDateValue::new("2024-02-29T23:59:59Z").unwrap(),
                )),
                "created date & <cached>",
            ),
            metadata_field(
                OdfDocumentMetadataFieldKind::CreationTime,
                Some(OdfDocumentMetadataFieldValue::Time(
                    OdfFieldTimeValue::new("2024-02-29T24:00:00+14:00").unwrap(),
                )),
                "created time",
            ),
            metadata_field(
                OdfDocumentMetadataFieldKind::PrintDate,
                Some(OdfDocumentMetadataFieldValue::Date(
                    OdfFieldDateValue::new("2025-01-31-05:00").unwrap(),
                )),
                "print date",
            ),
            metadata_field(
                OdfDocumentMetadataFieldKind::PrintTime,
                Some(OdfDocumentMetadataFieldValue::Time(
                    OdfFieldTimeValue::new("12:34:56.789Z").unwrap(),
                )),
                "print time",
            ),
            metadata_field(OdfDocumentMetadataFieldKind::EditingCycles, None, "42"),
            metadata_field(
                OdfDocumentMetadataFieldKind::EditingDuration,
                Some(OdfDocumentMetadataFieldValue::Duration(
                    OdfFieldDuration::new("P999999999999Y11M30DT23H59M59.5S").unwrap(),
                )),
                "edited duration",
            ),
            metadata_field(
                OdfDocumentMetadataFieldKind::ModificationDate,
                Some(OdfDocumentMetadataFieldValue::Date(
                    OdfFieldDateValue::new("-12345-12-31Z").unwrap(),
                )),
                "modified date",
            ),
            metadata_field(
                OdfDocumentMetadataFieldKind::ModificationTime,
                Some(OdfDocumentMetadataFieldValue::Time(
                    OdfFieldTimeValue::new("00:00:00+05:30").unwrap(),
                )),
                "modified time",
            ),
        ];
        let body = fields
            .iter()
            .map(OdfDynamicTextField::to_xml_fragment)
            .collect::<Result<Vec<_>>>()
            .unwrap()
            .join("");
        let parsed = FieldParser::parse_dynamic_text_fields(&document(&body)).unwrap();
        assert_eq!(parsed, fields);
        assert_eq!(parsed[0].display_text(), "created date & <cached>");
    }

    #[test]
    fn document_metadata_fixed_fields_preserve_optional_attributes_and_aliases() {
        let xml = document(
            r#"<t:creation-date>created</t:creation-date>
               <t:creation-time t:fixed="0">time</t:creation-time>
               <t:print-date s:data-style-name="D">date</t:print-date>
               <t:print-time>time</t:print-time>
               <t:editing-cycles t:fixed="1">7</t:editing-cycles>
               <t:editing-duration s:data-style-name="Elapsed">duration</t:editing-duration>
               <t:modification-date>date</t:modification-date>
               <t:modification-time>time</t:modification-time>"#,
        );
        let fields = FieldParser::parse_dynamic_text_fields(&xml).unwrap();
        assert_eq!(fields.len(), 8);
        assert!(matches!(
            &fields[0],
            OdfDynamicTextField::DocumentMetadata {
                kind: OdfDocumentMetadataFieldKind::CreationDate,
                value: None,
                fixed: None,
                data_style_name: None,
                ..
            }
        ));
        assert!(matches!(
            &fields[4],
            OdfDynamicTextField::DocumentMetadata {
                kind: OdfDocumentMetadataFieldKind::EditingCycles,
                fixed: Some(true),
                ..
            }
        ));
    }

    #[test]
    fn document_metadata_fixed_fields_reject_invalid_lexicals_attributes_and_bounds() {
        let invalid = [
            r#"<t:creation-date t:date-value="2023-02-29">bad</t:creation-date>"#,
            r#"<t:creation-time t:time-value="12:60:00">bad</t:creation-time>"#,
            r#"<t:print-date t:date-value="2024-01-01T00:00:00">bad</t:print-date>"#,
            r#"<t:print-time t:time-value="2024-01-01T00:00:00">bad</t:print-time>"#,
            r#"<t:editing-cycles s:data-style-name="N">7</t:editing-cycles>"#,
            r#"<t:editing-cycles t:duration="P1D">7</t:editing-cycles>"#,
            r#"<t:editing-duration t:duration="P">bad</t:editing-duration>"#,
            r#"<t:modification-date t:time-value="12:00:00">bad</t:modification-date>"#,
            r#"<t:modification-time t:fixed="TRUE">bad</t:modification-time>"#,
        ];
        for body in invalid {
            assert!(
                FieldParser::parse_dynamic_text_fields(&document(body)).is_err(),
                "accepted {body}"
            );
        }

        let mismatched = metadata_field(
            OdfDocumentMetadataFieldKind::PrintDate,
            Some(OdfDocumentMetadataFieldValue::Date(
                OdfFieldDateValue::new("2024-01-01T00:00:00Z").unwrap(),
            )),
            "bad",
        );
        assert!(mismatched.to_xml_fragment().is_err());

        let wrong_namespace = r#"<o:document-content
            xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
            xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
            xmlns:x="urn:not-style"><o:body><o:text><t:p>
            <t:print-date x:data-style-name="spoof">date</t:print-date>
            </t:p></o:text></o:body></o:document-content>"#;
        assert!(FieldParser::parse_dynamic_text_fields(wrong_namespace).is_err());

        let oversized = OdfDynamicTextField::DocumentMetadata {
            kind: OdfDocumentMetadataFieldKind::EditingCycles,
            value: None,
            fixed: None,
            data_style_name: None,
            display_text: "x".repeat(MAX_DYNAMIC_FIELD_VALUE + 1),
        };
        assert!(oversized.to_xml_fragment().is_err());
        let forbidden = OdfDynamicTextField::DocumentMetadata {
            kind: OdfDocumentMetadataFieldKind::EditingCycles,
            value: None,
            fixed: None,
            data_style_name: None,
            display_text: "bad\u{0}".to_string(),
        };
        assert!(forbidden.to_xml_fragment().is_err());
    }
}

#[cfg(test)]
mod document_identity_fixed_field_tests {
    use super::*;

    fn document(body: &str) -> String {
        format!(
            r#"<o:document-content
                xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
                xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
                <o:body><o:text><t:p>{body}</t:p></o:text></o:body>
            </o:document-content>"#
        )
    }

    #[test]
    fn document_identity_fixed_fields_round_trip_all_nine_standard_elements() {
        let kinds = [
            OdfDocumentIdentityFieldKind::InitialCreator,
            OdfDocumentIdentityFieldKind::Description,
            OdfDocumentIdentityFieldKind::PrintedBy,
            OdfDocumentIdentityFieldKind::Title,
            OdfDocumentIdentityFieldKind::Subject,
            OdfDocumentIdentityFieldKind::Keywords,
            OdfDocumentIdentityFieldKind::Creator,
            OdfDocumentIdentityFieldKind::AuthorName,
            OdfDocumentIdentityFieldKind::AuthorInitials,
        ];
        let fields = kinds
            .into_iter()
            .enumerate()
            .map(|(index, kind)| OdfDynamicTextField::DocumentIdentity {
                kind,
                fixed: Some(index % 2 == 0),
                display_text: format!("cached {index} & <inert>"),
            })
            .collect::<Vec<_>>();
        let body = fields
            .iter()
            .map(OdfDynamicTextField::to_xml_fragment)
            .collect::<Result<Vec<_>>>()
            .unwrap()
            .join("");
        let parsed = FieldParser::parse_dynamic_text_fields(&document(&body)).unwrap();
        assert_eq!(parsed, fields);
        assert_eq!(parsed[3].display_text(), "cached 3 & <inert>");
    }

    #[test]
    fn document_identity_fixed_fields_preserve_omission_and_namespace_aliases() {
        let xml = document(
            r#"<t:initial-creator>first</t:initial-creator>
               <t:description>description</t:description>
               <t:printed-by>printer</t:printed-by>
               <t:title>title</t:title>
               <t:subject>subject</t:subject>
               <t:keywords>one, two</t:keywords>
               <t:creator>last</t:creator>
               <t:author-name>author</t:author-name>
               <t:author-initials>au</t:author-initials>"#,
        );
        let fields = FieldParser::parse_dynamic_text_fields(&xml).unwrap();
        assert_eq!(fields.len(), 9);
        for field in fields {
            assert!(matches!(
                field,
                OdfDynamicTextField::DocumentIdentity { fixed: None, .. }
            ));
        }
    }

    #[test]
    fn document_identity_fixed_fields_reject_hostile_attributes_and_bounds() {
        let invalid = [
            r#"<t:initial-creator t:fixed="TRUE">first</t:initial-creator>"#,
            r#"<t:description t:name="not-standard">description</t:description>"#,
            r#"<t:printed-by t:fixed="yes">printer</t:printed-by>"#,
            r#"<t:title t:display="value">title</t:title>"#,
            r#"<t:subject t:fixed="2">subject</t:subject>"#,
            r#"<t:keywords t:string-value="one">one</t:keywords>"#,
            r#"<t:creator t:extra="x">creator</t:creator>"#,
            r#"<t:author-name t:display="value">author</t:author-name>"#,
            r#"<t:author-initials t:fixed="yes">au</t:author-initials>"#,
        ];
        for body in invalid {
            assert!(
                FieldParser::parse_dynamic_text_fields(&document(body)).is_err(),
                "accepted {body}"
            );
        }

        let wrong_namespace = r#"<o:document-content
            xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
            xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
            xmlns:x="urn:not-text"><o:body><o:text><t:p>
            <t:title x:fixed="true">spoof</t:title>
            </t:p></o:text></o:body></o:document-content>"#;
        assert!(FieldParser::parse_dynamic_text_fields(wrong_namespace).is_err());

        let oversized = OdfDynamicTextField::DocumentIdentity {
            kind: OdfDocumentIdentityFieldKind::Description,
            fixed: None,
            display_text: "x".repeat(MAX_DYNAMIC_FIELD_VALUE + 1),
        };
        assert!(oversized.to_xml_fragment().is_err());
        let forbidden = OdfDynamicTextField::DocumentIdentity {
            kind: OdfDocumentIdentityFieldKind::Title,
            fixed: Some(true),
            display_text: "bad\u{0}".to_string(),
        };
        assert!(forbidden.to_xml_fragment().is_err());
    }
}

#[cfg(test)]
mod user_defined_metadata_field_tests {
    use super::*;

    fn document(body: &str) -> String {
        format!(
            r#"<o:document-content
                xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
                xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
                xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0">
                <o:body><o:text><t:p>{body}</t:p></o:text></o:body>
            </o:document-content>"#
        )
    }

    #[test]
    fn user_defined_metadata_field_round_trips_every_independent_value_attribute() {
        let field = OdfDynamicTextField::UserDefinedMetadata {
            name: "custom & name".to_string(),
            values: OdfUserDefinedMetadataValues {
                number: Some("-INF".to_string()),
                date: Some(OdfFieldDateValue::new("2024-02-29T23:59:59Z").unwrap()),
                time: Some(OdfFieldDuration::new("P999999999999Y1M2DT3H4M5.6S").unwrap()),
                boolean: Some(false),
                string: Some("cached & <string>".to_string()),
            },
            fixed: Some(true),
            data_style_name: Some("Custom & Style".to_string()),
            display_text: "inert & <presentation>".to_string(),
        };
        let fragment = field.to_xml_fragment().unwrap();
        assert!(fragment.contains("office:value=\"-INF\""));
        assert!(fragment.contains("office:boolean-value=\"false\""));
        let parsed = FieldParser::parse_dynamic_text_fields(&document(&fragment)).unwrap();
        assert_eq!(parsed, vec![field]);
        assert_eq!(parsed[0].display_text(), "inert & <presentation>");
    }

    #[test]
    fn user_defined_metadata_field_preserves_empty_name_values_and_omission() {
        let xml = document(
            r#"<t:user-defined t:name="" o:string-value="" t:fixed="0"
                xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0">cached</t:user-defined>
               <t:user-defined t:name="minimal"/>"#,
        );
        let fields = FieldParser::parse_dynamic_text_fields(&xml).unwrap();
        assert_eq!(fields.len(), 2);
        assert!(matches!(
            &fields[0],
            OdfDynamicTextField::UserDefinedMetadata {
                name,
                values: OdfUserDefinedMetadataValues { string: Some(value), .. },
                fixed: Some(false),
                ..
            } if name.is_empty() && value.is_empty()
        ));
        assert!(matches!(
            &fields[1],
            OdfDynamicTextField::UserDefinedMetadata {
                values: OdfUserDefinedMetadataValues {
                    number: None,
                    date: None,
                    time: None,
                    boolean: None,
                    string: None,
                },
                fixed: None,
                data_style_name: None,
                ..
            }
        ));
    }

    #[test]
    fn user_defined_metadata_field_rejects_nonstandard_invalid_and_hostile_input() {
        let invalid = [
            r#"<t:user-defined>missing name</t:user-defined>"#,
            r#"<t:user-defined t:name="x" o:value-type="float" xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"/>"#,
            r#"<t:user-defined t:name="x" o:currency="USD" xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"/>"#,
            r#"<t:user-defined t:name="x" o:value="1e" xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"/>"#,
            r#"<t:user-defined t:name="x" o:date-value="2023-02-29" xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"/>"#,
            r#"<t:user-defined t:name="x" o:time-value="P" xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"/>"#,
            r#"<t:user-defined t:name="x" o:boolean-value="TRUE" xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"/>"#,
            r#"<t:user-defined t:name="x" t:fixed="yes"/>"#,
        ];
        for body in invalid {
            assert!(
                FieldParser::parse_dynamic_text_fields(&document(body)).is_err(),
                "accepted {body}"
            );
        }

        let wrong_namespace = r#"<o:document-content
            xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
            xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
            xmlns:x="urn:not-office"><o:body><o:text><t:p>
            <t:user-defined t:name="x" x:string-value="spoof">value</t:user-defined>
            </t:p></o:text></o:body></o:document-content>"#;
        assert!(FieldParser::parse_dynamic_text_fields(wrong_namespace).is_err());

        let oversized = OdfDynamicTextField::UserDefinedMetadata {
            name: "x".to_string(),
            values: OdfUserDefinedMetadataValues {
                string: Some("x".repeat(MAX_DYNAMIC_FIELD_VALUE + 1)),
                ..OdfUserDefinedMetadataValues::default()
            },
            fixed: None,
            data_style_name: None,
            display_text: String::new(),
        };
        assert!(oversized.to_xml_fragment().is_err());
        let forbidden = OdfDynamicTextField::UserDefinedMetadata {
            name: "bad\u{0}".to_string(),
            values: OdfUserDefinedMetadataValues::default(),
            fixed: None,
            data_style_name: None,
            display_text: String::new(),
        };
        assert!(forbidden.to_xml_fragment().is_err());
    }
}

#[cfg(test)]
mod meta_field_tests {
    use super::*;

    fn document(body: &str) -> String {
        format!(
            r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content
 xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
 xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
 xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
 xmlns:xlink="http://www.w3.org/1999/xlink"
 xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
 <office:body><office:text>{body}</office:text></office:body>
</office:document-content>"#
        )
    }

    #[test]
    fn meta_field_preserves_ordered_mixed_content_and_roundtrips() {
        let xml = document(
            r#"<text:p><text:meta-field xml:id="meta1" style:data-style-name="N1">before<text:span text:style-name="Em">middle</text:span>after<text:a xlink:href="https://example.invalid" xlink:type="simple">link</text:a>end</text:meta-field></text:p>"#,
        );
        let fields = FieldParser::parse_dynamic_text_fields(&xml).unwrap();
        assert_eq!(fields.len(), 1);
        let OdfDynamicTextField::MetaField {
            xml_id,
            data_style_name,
            content,
        } = &fields[0]
        else {
            panic!("expected metadata field");
        };
        assert_eq!(xml_id, "meta1");
        assert_eq!(data_style_name.as_deref(), Some("N1"));
        assert_eq!(content.display_text(), "beforemiddleafterlinkend");
        assert!(matches!(
            content.nodes(),
            [
                OdfMetaFieldNode::Text(_),
                OdfMetaFieldNode::Element(_),
                OdfMetaFieldNode::Text(_),
                OdfMetaFieldNode::Element(_),
                OdfMetaFieldNode::Text(_),
            ]
        ));

        let fragment = fields[0].to_xml_fragment().unwrap();
        assert!(fragment.contains("xml:id=\"meta1\""));
        assert!(fragment.contains("style:data-style-name=\"N1\""));
        assert!(fragment.contains("xlink:href=\"https://example.invalid\""));
        let reparsed = FieldParser::parse_dynamic_text_fields(&document(&format!(
            "<text:p>{fragment}</text:p>"
        )))
        .unwrap();
        assert_eq!(reparsed, fields);
    }

    #[test]
    fn meta_field_recursion_is_inert_and_fields_remain_in_document_order() {
        let xml = document(
            r#"<text:p><text:meta-field xml:id="outer">A<text:meta-field xml:id="inner">B</text:meta-field>C</text:meta-field></text:p>"#,
        );
        let fields = FieldParser::parse_dynamic_text_fields(&xml).unwrap();
        assert_eq!(fields.len(), 2);
        let OdfDynamicTextField::MetaField {
            xml_id, content, ..
        } = &fields[0]
        else {
            panic!("expected outer metadata field");
        };
        assert_eq!(xml_id, "outer");
        assert_eq!(content.display_text(), "ABC");
        assert!(
            matches!(content.nodes().get(1), Some(OdfMetaFieldNode::Element(element)) if element.local_name == "meta-field")
        );
        let OdfDynamicTextField::MetaField {
            xml_id, content, ..
        } = &fields[1]
        else {
            panic!("expected inner metadata field");
        };
        assert_eq!(xml_id, "inner");
        assert_eq!(content.display_text(), "B");
    }

    #[test]
    fn meta_field_rejects_invalid_identity_placement_and_markup() {
        let invalid = [
            r#"<text:p><text:meta-field>missing</text:meta-field></text:p>"#,
            r#"<text:p><text:meta-field xml:id="">empty</text:meta-field></text:p>"#,
            r#"<text:p><text:meta-field xml:id="1bad">bad</text:meta-field></text:p>"#,
            r#"<text:p><text:meta-field xml:id="bad:id">bad</text:meta-field></text:p>"#,
            r#"<text:p><text:meta-field xml:id="m" text:style-name="bad">bad</text:meta-field></text:p>"#,
            r#"<text:p><text:meta-field xml:id="same">a</text:meta-field><text:meta-field xml:id="same">b</text:meta-field></text:p>"#,
            r#"<text:meta-field xml:id="m">not paragraph content</text:meta-field>"#,
            r#"<text:p><text:meta-field xml:id="m"><table:table/></text:meta-field></text:p>"#,
            r#"<text:p><text:meta-field xml:id="m"><evil:x xmlns:evil="urn:evil"/></text:meta-field></text:p>"#,
            r#"<text:p><text:meta-field xml:id="m"><bad:x/></text:meta-field></text:p>"#,
        ];
        for body in invalid {
            assert!(
                FieldParser::parse_dynamic_text_fields(&document(body)).is_err(),
                "accepted {body}"
            );
        }

        let dtd = document(r#"<text:p><text:meta-field xml:id="m">x</text:meta-field></text:p>"#)
            .replacen(
                "<office:document-content",
                "<!DOCTYPE x><office:document-content",
                1,
            );
        assert!(FieldParser::parse_dynamic_text_fields(&dtd).is_err());
        assert!(FieldParser::parse_dynamic_text_fields(&document(
            r#"<text:p><text:meta-field xml:id="m">a<?unsafe data?>b</text:meta-field></text:p>"#,
        ))
        .is_err());
    }

    #[test]
    fn meta_field_dispatch_ignores_foreign_vocabulary_but_keeps_real_roots_strict() {
        let spoof = document(
            r#"<text:p><fake:meta-field xmlns:fake="urn:not-text" fake:attribute="ignored">spoof</fake:meta-field></text:p>"#,
        );
        assert!(
            FieldParser::parse_dynamic_text_fields(&spoof)
                .unwrap()
                .is_empty()
        );

        let genuine = document(
            r#"<text:p><text:meta-field xmlns:fake="urn:not-text" xml:id="m" fake:attribute="rejected">real</text:meta-field></text:p>"#,
        );
        assert!(FieldParser::parse_dynamic_text_fields(&genuine).is_err());
    }

    #[test]
    fn meta_field_accepts_rng_inline_child_grammars() {
        let allowed = [
            r#"text<text:span><text:a xlink:type="simple" xlink:href="urn:test"><text:date>2026-07-18</text:date></text:a></text:span>"#,
            r#"<text:meta xml:id="nested-meta"><text:meta-field xml:id="nested-field">nested</text:meta-field></text:meta>"#,
            r#"<text:ruby><text:ruby-base>base<text:span>span</text:span></text:ruby-base><text:ruby-text>reading</text:ruby-text></text:ruby>"#,
            r#"<text:note text:note-class="footnote"><text:note-citation>1</text:note-citation><text:note-body><text:p>body</text:p></text:note-body></text:note>"#,
            r#"<text:execute-macro xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0"><office:event-listeners><script:event-listener script:event-name="dom:click" script:language="ooo:script" script:macro-name="M"/></office:event-listeners>cached</text:execute-macro>"#,
            r#"<office:annotation><dc:creator xmlns:dc="http://purl.org/dc/elements/1.1/">A</dc:creator><text:p>comment</text:p></office:annotation>"#,
            r#"<draw:line xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"/>"#,
            r#"<presentation:header xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0"/>"#,
        ];
        for (index, content) in allowed.into_iter().enumerate() {
            let xml = document(&format!(
                r#"<text:p><text:meta-field xml:id="allowed{index}">{content}</text:meta-field></text:p>"#,
            ));
            assert!(
                FieldParser::parse_dynamic_text_fields(&xml).is_ok(),
                "rejected allowed RNG content {content}"
            );
        }
    }

    #[test]
    fn meta_field_rejects_wrong_rng_descendant_elements_and_cardinality() {
        let disallowed = [
            r#"<text:a xlink:type="simple" xlink:href="urn:outer"><text:a xlink:type="simple" xlink:href="urn:inner">nested</text:a></text:a>"#,
            r#"<text:date><text:span>not cached text</text:span></text:date>"#,
            r#"<text:s>not empty</text:s>"#,
            r#"<text:number>heading-only vocabulary</text:number>"#,
            r#"<text:ruby><text:ruby-text>wrong order</text:ruby-text><text:ruby-base>base</text:ruby-base></text:ruby>"#,
            r#"<text:ruby><text:ruby-base>missing reading</text:ruby-base></text:ruby>"#,
            r#"<text:note text:note-class="footnote"><text:note-body/><text:note-citation>1</text:note-citation></text:note>"#,
            r#"<text:note text:note-class="footnote"><text:note-citation>1</text:note-citation><text:note-body><style:style/></text:note-body></text:note>"#,
            r#"<text:execute-macro>text<office:event-listeners/></text:execute-macro>"#,
            r#"<office:event-listeners/>"#,
            r#"<text:span><table:table/></text:span>"#,
            r#"<text:span><style:style/></text:span>"#,
            r#"<draw:line xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"><table:table/></draw:line>"#,
        ];
        for (index, content) in disallowed.into_iter().enumerate() {
            let xml = document(&format!(
                r#"<text:p><text:meta-field xml:id="bad{index}">{content}</text:meta-field></text:p>"#,
            ));
            assert!(
                FieldParser::parse_dynamic_text_fields(&xml).is_err(),
                "accepted disallowed RNG content {content}"
            );
        }
    }

    #[test]
    fn meta_field_scan_enforces_document_wide_xml_id_uniqueness() {
        let duplicate_with_meta = document(
            r#"<text:p xml:id="same">before</text:p><text:p><text:meta-field xml:id="same">meta</text:meta-field></text:p>"#,
        );
        assert!(FieldParser::parse_dynamic_text_fields(&duplicate_with_meta).is_err());

        let duplicate_outside_meta = document(
            r#"<text:p xml:id="same">one</text:p><text:p xml:id="same">two</text:p><text:p><text:meta-field xml:id="unique">meta</text:meta-field></text:p>"#,
        );
        assert!(FieldParser::parse_dynamic_text_fields(&duplicate_outside_meta).is_err());

        let invalid_outside_meta = document(
            r#"<text:p xml:id="1invalid">one</text:p><text:p><text:meta-field xml:id="unique">meta</text:meta-field></text:p>"#,
        );
        assert!(FieldParser::parse_dynamic_text_fields(&invalid_outside_meta).is_err());

        let unique = document(
            r#"<text:p xml:id="paragraph">one</text:p><text:p><text:meta-field xml:id="field">meta</text:meta-field></text:p>"#,
        );
        assert!(FieldParser::parse_dynamic_text_fields(&unique).is_ok());
    }

    #[test]
    fn meta_field_content_constructor_enforces_resource_and_xml_bounds() {
        assert!(
            OdfMetaFieldContent::new(vec![OdfMetaFieldNode::Text(
                "x".repeat(MAX_DYNAMIC_FIELD_VALUE + 1),
            )])
            .is_err()
        );
        assert!(
            OdfMetaFieldContent::new(vec![OdfMetaFieldNode::Text("bad\u{1}control".to_string(),)])
                .is_err()
        );

        let mut nested = OdfMetaFieldNode::Text("leaf".to_string());
        for _ in 0..=MAX_META_FIELD_DEPTH {
            nested = OdfMetaFieldNode::Element(OdfMetaFieldElement {
                namespace_uri: TEXT_DATABASE_NAMESPACE.to_string(),
                local_name: "span".to_string(),
                attributes: Vec::new(),
                children: vec![nested],
            });
        }
        assert!(OdfMetaFieldContent::new(vec![nested]).is_err());

        let oversized_attribute = OdfMetaFieldElement {
            namespace_uri: TEXT_DATABASE_NAMESPACE.to_string(),
            local_name: "span".to_string(),
            attributes: vec![OdfMetaFieldAttribute {
                namespace_uri: TEXT_DATABASE_NAMESPACE.to_string(),
                local_name: "style-name".to_string(),
                value: "x".repeat(MAX_DYNAMIC_FIELD_VALUE + 1),
            }],
            children: Vec::new(),
        };
        assert!(
            OdfMetaFieldContent::new(vec![OdfMetaFieldNode::Element(oversized_attribute,)])
                .is_err()
        );
    }

    #[test]
    fn note_body_content_enforces_block_root_grammar_and_projects_text() {
        let content = OdfNoteBodyContent::new(vec![
            OdfMetaFieldNode::Element(OdfMetaFieldElement {
                namespace_uri: TEXT_DATABASE_NAMESPACE.to_string(),
                local_name: "p".to_string(),
                attributes: Vec::new(),
                children: vec![
                    OdfMetaFieldNode::Text("First ".to_string()),
                    OdfMetaFieldNode::Element(OdfMetaFieldElement {
                        namespace_uri: TEXT_DATABASE_NAMESPACE.to_string(),
                        local_name: "span".to_string(),
                        attributes: vec![OdfMetaFieldAttribute {
                            namespace_uri: TEXT_DATABASE_NAMESPACE.to_string(),
                            local_name: "style-name".to_string(),
                            value: "Emphasis".to_string(),
                        }],
                        children: vec![OdfMetaFieldNode::Text("styled".to_string())],
                    }),
                ],
            }),
            OdfMetaFieldNode::Element(OdfMetaFieldElement {
                namespace_uri: TEXT_DATABASE_NAMESPACE.to_string(),
                local_name: "list".to_string(),
                attributes: Vec::new(),
                children: vec![OdfMetaFieldNode::Element(OdfMetaFieldElement {
                    namespace_uri: TEXT_DATABASE_NAMESPACE.to_string(),
                    local_name: "list-item".to_string(),
                    attributes: Vec::new(),
                    children: vec![OdfMetaFieldNode::Element(OdfMetaFieldElement {
                        namespace_uri: TEXT_DATABASE_NAMESPACE.to_string(),
                        local_name: "p".to_string(),
                        attributes: Vec::new(),
                        children: vec![OdfMetaFieldNode::Text("Second".to_string())],
                    })],
                })],
            }),
        ])
        .unwrap();
        assert_eq!(content.display_text(), "First styled\nSecond");
        assert!(content.validate().is_ok());

        assert!(
            OdfNoteBodyContent::new(vec![OdfMetaFieldNode::Text("not a block".to_string(),)])
                .is_err()
        );
        assert!(
            OdfNoteBodyContent::new(vec![OdfMetaFieldNode::Element(OdfMetaFieldElement {
                namespace_uri: TEXT_DATABASE_NAMESPACE.to_string(),
                local_name: "span".to_string(),
                attributes: Vec::new(),
                children: vec![OdfMetaFieldNode::Text("not a root block".to_string())],
            },)])
            .is_err()
        );
    }

    #[test]
    fn note_body_content_projects_odf_whitespace_controls() {
        let text_control = |local_name: &str, attributes: Vec<OdfMetaFieldAttribute>| {
            OdfMetaFieldNode::Element(OdfMetaFieldElement {
                namespace_uri: TEXT_DATABASE_NAMESPACE.to_string(),
                local_name: local_name.to_string(),
                attributes,
                children: Vec::new(),
            })
        };
        let content =
            OdfNoteBodyContent::new(vec![OdfMetaFieldNode::Element(OdfMetaFieldElement {
                namespace_uri: TEXT_DATABASE_NAMESPACE.to_string(),
                local_name: "p".to_string(),
                attributes: Vec::new(),
                children: vec![
                    OdfMetaFieldNode::Text("A".to_string()),
                    text_control(
                        "s",
                        vec![OdfMetaFieldAttribute {
                            namespace_uri: TEXT_DATABASE_NAMESPACE.to_string(),
                            local_name: "c".to_string(),
                            value: "2".to_string(),
                        }],
                    ),
                    text_control("tab", Vec::new()),
                    text_control("line-break", Vec::new()),
                    OdfMetaFieldNode::Text("B".to_string()),
                ],
            })])
            .unwrap();
        assert_eq!(content.display_text(), "A  \t\nB");

        let invalid =
            OdfNoteBodyContent::new(vec![OdfMetaFieldNode::Element(OdfMetaFieldElement {
                namespace_uri: TEXT_DATABASE_NAMESPACE.to_string(),
                local_name: "p".to_string(),
                attributes: Vec::new(),
                children: vec![text_control(
                    "s",
                    vec![OdfMetaFieldAttribute {
                        namespace_uri: TEXT_DATABASE_NAMESPACE.to_string(),
                        local_name: "c".to_string(),
                        value: "two".to_string(),
                    }],
                )],
            })]);
        assert!(invalid.is_err());
    }

    #[test]
    fn meta_field_serialization_is_canonical_and_escaped() {
        let content = OdfMetaFieldContent::new(vec![
            OdfMetaFieldNode::Text("a<&".to_string()),
            OdfMetaFieldNode::Element(OdfMetaFieldElement {
                namespace_uri: TEXT_DATABASE_NAMESPACE.to_string(),
                local_name: "span".to_string(),
                attributes: vec![OdfMetaFieldAttribute {
                    namespace_uri: TEXT_DATABASE_NAMESPACE.to_string(),
                    local_name: "style-name".to_string(),
                    value: "A&B\"".to_string(),
                }],
                children: vec![OdfMetaFieldNode::Text("z>".to_string())],
            }),
        ])
        .unwrap();
        let field = OdfDynamicTextField::MetaField {
            xml_id: "m1".to_string(),
            data_style_name: None,
            content,
        };
        let xml = field.to_xml_fragment().unwrap();
        assert!(xml.contains("a&lt;&amp;"));
        assert!(xml.contains("text:style-name=\"A&amp;B&quot;\""));
        assert!(xml.contains("z&gt;"));
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
