//! Typed ODF field values and semantic field models.

#[allow(
    clippy::wildcard_imports,
    reason = "the field model shares the owner-level namespace and bounds"
)]
use super::*;
use crate::elements::element::{Element, ElementBase};
use litchi_core::{Error, Result};
use std::collections::HashSet;

/// One of the five OpenDocument database field elements.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum DatabaseFieldKind {
    Display,
    Next,
    RowSelect,
    RowNumber,
    Name,
}

/// Kind of database object selected by a field.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum DatabaseTableType {
    Table,
    Query,
    Command,
}

impl DatabaseTableType {
    pub(super) fn parse(value: &str) -> Result<Self> {
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
pub struct DatabaseConnectionResource {
    pub href: String,
    pub simple_link: bool,
}

/// Common source identity shared by all database fields.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DatabaseSource {
    pub database_name: Option<String>,
    pub table_name: String,
    pub table_type: Option<DatabaseTableType>,
    pub connection_resource: Option<DatabaseConnectionResource>,
}

/// Canonical, bounded XML Schema `nonNegativeInteger` without arithmetic semantics.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct NonNegativeInteger(String);

impl NonNegativeInteger {
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

impl DatabaseSource {
    /// ODF defaults `text:table-type` to `table`.
    pub fn effective_table_type(&self) -> DatabaseTableType {
        self.table_type.unwrap_or(DatabaseTableType::Table)
    }
}

/// Typed, non-executing database field metadata in document order.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DatabaseField {
    pub kind: DatabaseFieldKind,
    pub source: DatabaseSource,
    pub column_name: Option<String>,
    pub condition: Option<String>,
    pub row_number: Option<NonNegativeInteger>,
    pub value: Option<NonNegativeInteger>,
    pub data_style_name: Option<String>,
    pub number_format: Option<String>,
    pub number_letter_sync: Option<bool>,
    pub display_text: String,
}

/// The content category requested by a `text:placeholder` field.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum PlaceholderType {
    Text,
    Table,
    TextBox,
    Image,
    Object,
}

/// One stored option in an ODF `text:drop-down` field.
///
/// Both attributes are optional in the ODF schema. The option itself is inert:
/// this type only retains producer-supplied metadata and never displays a user
/// interface or changes a selection.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct DropDownLabel {
    /// Optional producer-supplied option value.
    pub value: Option<String>,
    /// Optional stored selected-state flag.
    pub current_selected: Option<bool>,
}

/// Numbering metadata for an ODF `text:sequence` field.
///
/// ODF permits `style:num-letter-sync` only for alphabetic formats (`a` and
/// `A`). Other format strings, including producer-defined values and the empty
/// format, remain opaque and are preserved verbatim.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct SequenceNumberFormat {
    format: String,
    letter_sync: Option<bool>,
}

/// Page selected by an ODF page-number field.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum PageSelection {
    Previous,
    Current,
    Next,
}

impl PageSelection {
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
pub enum PageContinuationSelection {
    Previous,
    Next,
}

impl PageContinuationSelection {
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
pub enum DateValueKind {
    Date,
    DateTime,
}

/// A validated XML Schema `dateOrDateTime` value for `text:date-value`.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct FieldDateValue {
    lexical: String,
    kind: DateValueKind,
}

impl FieldDateValue {
    pub fn new(value: impl Into<String>) -> Result<Self> {
        let lexical = value.into();
        let kind = if lexical.contains('T') {
            DateValueKind::DateTime
        } else {
            DateValueKind::Date
        };
        let value = Self { lexical, kind };
        let mut aggregate = 0usize;
        value.validate(&mut aggregate)?;
        Ok(value)
    }

    pub fn as_str(&self) -> &str {
        &self.lexical
    }

    pub const fn kind(&self) -> DateValueKind {
        self.kind
    }

    fn validate(&self, aggregate: &mut usize) -> Result<()> {
        validate_dynamic_value("text:date-value", Some(&self.lexical), true, aggregate)?;
        match self.kind {
            DateValueKind::Date => validate_xml_schema_date(&self.lexical),
            DateValueKind::DateTime => validate_xml_schema_date_time(&self.lexical),
        }
    }
}

/// Lexical category retained by a typed ODF time value.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum TimeValueKind {
    Time,
    DateTime,
}

/// A validated XML Schema `timeOrDateTime` value for `text:time-value`.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct FieldTimeValue {
    lexical: String,
    kind: TimeValueKind,
}

impl FieldTimeValue {
    pub fn new(value: impl Into<String>) -> Result<Self> {
        let lexical = value.into();
        let kind = if lexical.contains('T') {
            TimeValueKind::DateTime
        } else {
            TimeValueKind::Time
        };
        let value = Self { lexical, kind };
        let mut aggregate = 0usize;
        value.validate(&mut aggregate)?;
        Ok(value)
    }

    pub fn as_str(&self) -> &str {
        &self.lexical
    }

    pub const fn kind(&self) -> TimeValueKind {
        self.kind
    }

    fn validate(&self, aggregate: &mut usize) -> Result<()> {
        validate_dynamic_value("text:time-value", Some(&self.lexical), true, aggregate)?;
        match self.kind {
            TimeValueKind::Time => validate_xml_schema_time(&self.lexical),
            TimeValueKind::DateTime => validate_xml_schema_date_time(&self.lexical),
        }
    }
}

/// A validated, exactly retained XML Schema duration used for field adjustment.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct FieldDuration(String);

impl FieldDuration {
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
        crate::datatype::Duration::decode_exact(&self.0).map_err(|_| {
            Error::InvalidFormat(format!("invalid XML Schema duration '{}'", self.0))
        })?;
        Ok(())
    }
}

/// Display format for a `text:sequence-ref` field.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum SequenceReferenceFormat {
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
pub enum VariableSetDisplay {
    Value,
    None,
}

impl VariableSetDisplay {
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
pub enum FormulaFieldDisplay {
    Value,
    Formula,
}

impl FormulaFieldDisplay {
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

/// Display format permitted by ODF 1.2's `text:file-name` field (§19.796.4).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum FileNameDisplay {
    Full,
    Path,
    Name,
    NameAndExtension,
}

impl FileNameDisplay {
    const fn as_str(self) -> &'static str {
        match self {
            Self::Full => "full",
            Self::Path => "path",
            Self::Name => "name",
            Self::NameAndExtension => "name-and-extension",
        }
    }

    fn parse(value: &str) -> Result<Self> {
        match value {
            "full" => Ok(Self::Full),
            "path" => Ok(Self::Path),
            "name" => Ok(Self::Name),
            "name-and-extension" => Ok(Self::NameAndExtension),
            _ => Err(Error::InvalidFormat(format!(
                "invalid file-name text:display '{value}'"
            ))),
        }
    }
}

/// Display format permitted by ODF 1.2's `text:template-name` field (§19.796.8).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum TemplateNameDisplay {
    Area,
    Full,
    Name,
    NameAndExtension,
    Path,
    Title,
}

impl TemplateNameDisplay {
    const fn as_str(self) -> &'static str {
        match self {
            Self::Area => "area",
            Self::Full => "full",
            Self::Name => "name",
            Self::NameAndExtension => "name-and-extension",
            Self::Path => "path",
            Self::Title => "title",
        }
    }

    fn parse(value: &str) -> Result<Self> {
        match value {
            "area" => Ok(Self::Area),
            "full" => Ok(Self::Full),
            "name" => Ok(Self::Name),
            "name-and-extension" => Ok(Self::NameAndExtension),
            "path" => Ok(Self::Path),
            "title" => Ok(Self::Title),
            _ => Err(Error::InvalidFormat(format!(
                "invalid template-name text:display '{value}'"
            ))),
        }
    }
}

/// Display format permitted by ODF 1.2's `text:chapter` field (§19.796.2).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ChapterDisplay {
    Name,
    Number,
    NumberAndName,
    PlainNumber,
    PlainNumberAndName,
}

impl ChapterDisplay {
    const fn as_str(self) -> &'static str {
        match self {
            Self::Name => "name",
            Self::Number => "number",
            Self::NumberAndName => "number-and-name",
            Self::PlainNumber => "plain-number",
            Self::PlainNumberAndName => "plain-number-and-name",
        }
    }

    fn parse(value: &str) -> Result<Self> {
        match value {
            "name" => Ok(Self::Name),
            "number" => Ok(Self::Number),
            "number-and-name" => Ok(Self::NumberAndName),
            "plain-number" => Ok(Self::PlainNumber),
            "plain-number-and-name" => Ok(Self::PlainNumberAndName),
            _ => Err(Error::InvalidFormat(format!(
                "invalid chapter text:display '{value}'"
            ))),
        }
    }
}

/// Strict ODF `common-value-and-type-attlist` cached value group.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub enum CalculatedFieldValue {
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
pub enum FieldValueType {
    Float,
    Time,
    Date,
    Percentage,
    Currency,
    Boolean,
    String,
}

impl FieldValueType {
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
pub enum UserFieldDisplay {
    Value,
    Formula,
    None,
}

/// Component displayed by a `text:measure` field.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum MeasureKind {
    Value,
    Unit,
    Gap,
}

/// Display format shared by `text:reference-ref` and `text:bookmark-ref`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum CrossReferenceFormat {
    Page,
    Chapter,
    Direction,
    Text,
    NumberNoSuperior,
    NumberAllSuperior,
    Number,
}

impl CrossReferenceFormat {
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
pub enum NoteReferenceFormat {
    Page,
    Chapter,
    Direction,
    Text,
}

impl NoteReferenceFormat {
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
pub enum NoteReferenceClass {
    Footnote,
    Endnote,
}

/// Kind of cached ODF document statistic.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum StatisticKind {
    Page,
    Paragraph,
    Word,
    Character,
    Table,
    Image,
    Object,
}

impl StatisticKind {
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
pub enum MetadataFieldKind {
    CreationDate,
    CreationTime,
    PrintDate,
    PrintTime,
    EditingCycles,
    EditingDuration,
    ModificationDate,
    ModificationTime,
}

impl MetadataFieldKind {
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

    pub(super) const fn permits_data_style(self) -> bool {
        !matches!(self, Self::EditingCycles)
    }
}

/// Strict typed value attribute for a temporal document-metadata field.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub enum MetadataFieldValue {
    Date(FieldDateValue),
    Time(FieldTimeValue),
    Duration(FieldDuration),
}

/// One of the nine fixed string/identity document-metadata fields.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum IdentityFieldKind {
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

impl IdentityFieldKind {
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

/// One of the fifteen ODF 1.2 subsequent-author `text:sender-*` field categories.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum SenderFieldKind {
    FirstName,
    LastName,
    Initials,
    Title,
    Position,
    Email,
    PrivatePhone,
    Fax,
    Company,
    WorkPhone,
    Street,
    City,
    PostalCode,
    Country,
    StateOrProvince,
}

impl SenderFieldKind {
    pub const fn element_name(self) -> &'static str {
        match self {
            Self::FirstName => "text:sender-firstname",
            Self::LastName => "text:sender-lastname",
            Self::Initials => "text:sender-initials",
            Self::Title => "text:sender-title",
            Self::Position => "text:sender-position",
            Self::Email => "text:sender-email",
            Self::PrivatePhone => "text:sender-phone-private",
            Self::Fax => "text:sender-fax",
            Self::Company => "text:sender-company",
            Self::WorkPhone => "text:sender-phone-work",
            Self::Street => "text:sender-street",
            Self::City => "text:sender-city",
            Self::PostalCode => "text:sender-postal-code",
            Self::Country => "text:sender-country",
            Self::StateOrProvince => "text:sender-state-or-province",
        }
    }
}

/// Independently optional cached values permitted by `text:user-defined`.
///
/// Unlike variable fields, ODF 1.2 does not use `office:value-type` here and
/// its schema permits more than one of these attributes to coexist.
#[derive(Debug, Clone, PartialEq, Eq, Hash, Default)]
pub struct UserDefinedMetadataValues {
    pub number: Option<String>,
    pub date: Option<FieldDateValue>,
    pub time: Option<FieldDuration>,
    pub boolean: Option<bool>,
    pub string: Option<String>,
}

/// A namespace-resolved attribute on inert `text:meta-field` content.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct MetaFieldAttribute {
    pub namespace_uri: String,
    pub local_name: String,
    pub value: String,
}

/// A namespace-resolved inert inline element.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct MetaFieldElement {
    pub namespace_uri: String,
    pub local_name: String,
    pub attributes: Vec<MetaFieldAttribute>,
    pub children: Vec<MetaFieldNode>,
}

/// Ordered mixed content retained by `text:meta-field`.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub enum MetaFieldNode {
    Text(String),
    Element(MetaFieldElement),
}

/// Validated, inert mixed content with a cached plain-text projection.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct MetaFieldContent {
    nodes: Vec<MetaFieldNode>,
    display_text: String,
}

impl MetaFieldContent {
    pub fn new(nodes: Vec<MetaFieldNode>) -> Result<Self> {
        let display_text =
            validated_meta_display_text(&nodes, MetaContentGrammar::ParagraphOrHyperlink)?;
        Ok(Self {
            nodes,
            display_text,
        })
    }

    pub fn nodes(&self) -> &[MetaFieldNode] {
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
pub struct NoteBodyContent {
    nodes: Vec<MetaFieldNode>,
    display_text: String,
}

impl NoteBodyContent {
    /// Construct structured note-body content from namespace-resolved nodes.
    pub fn new(nodes: Vec<MetaFieldNode>) -> Result<Self> {
        if nodes
            .iter()
            .any(|node| matches!(node, MetaFieldNode::Text(_)))
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
    pub fn nodes(&self) -> &[MetaFieldNode] {
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
    nodes: &[MetaFieldNode],
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

fn note_body_display_text(nodes: &[MetaFieldNode]) -> Result<String> {
    let mut display_text = String::new();
    let mut seen_block = false;
    append_note_body_display_text(nodes, &mut display_text, &mut seen_block, false)?;
    Ok(display_text)
}

fn append_note_body_display_text(
    nodes: &[MetaFieldNode],
    display_text: &mut String,
    seen_block: &mut bool,
    in_paragraph: bool,
) -> Result<()> {
    for node in nodes {
        match node {
            MetaFieldNode::Text(value) if in_paragraph => {
                append_note_body_display_value(display_text, value)?;
            },
            MetaFieldNode::Text(_) => {},
            MetaFieldNode::Element(element) => {
                if element.namespace_uri == TEXT_DATABASE_NAMESPACE && element.local_name == "note"
                {
                    if in_paragraph
                        && let Some(MetaFieldNode::Element(citation)) = element.children.first()
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

fn append_note_body_spaces(output: &mut String, element: &MetaFieldElement) -> Result<()> {
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

impl UserDefinedMetadataValues {
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

impl NoteReferenceClass {
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

impl MeasureKind {
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

impl UserFieldDisplay {
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

impl CalculatedFieldValue {
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
                    crate::datatype::DateTime::decode(value).map_err(|_| {
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
                crate::datatype::Duration::decode_exact(value).map_err(|_| {
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

impl SequenceReferenceFormat {
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

impl SequenceNumberFormat {
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

impl PlaceholderType {
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
pub enum DynamicTextField {
    Placeholder {
        placeholder_type: PlaceholderType,
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
        number_format: Option<SequenceNumberFormat>,
        reference_name: Option<String>,
        display_text: String,
    },
    /// A cached reference to a named sequence value.
    SequenceReference {
        reference_name: String,
        reference_format: Option<SequenceReferenceFormat>,
        display_text: String,
    },
    VariableSet {
        name: String,
        formula: Option<String>,
        value: CalculatedFieldValue,
        display: Option<VariableSetDisplay>,
        data_style_name: Option<String>,
        display_text: String,
    },
    VariableGet {
        name: String,
        display: Option<FormulaFieldDisplay>,
        data_style_name: Option<String>,
        display_text: String,
    },
    Expression {
        formula: Option<String>,
        value: Option<CalculatedFieldValue>,
        display: Option<FormulaFieldDisplay>,
        data_style_name: Option<String>,
        display_text: String,
    },
    VariableInput {
        name: String,
        description: Option<String>,
        value_type: FieldValueType,
        display: Option<VariableSetDisplay>,
        data_style_name: Option<String>,
        display_text: String,
    },
    UserFieldGet {
        name: String,
        display: Option<UserFieldDisplay>,
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
    /// An inert drop-down input field with stored choice metadata.
    ///
    /// The labels and cached selected text are retained exactly as document
    /// metadata. No selection interface is shown and no label is selected,
    /// changed, or resolved by this API.
    DropDown {
        name: String,
        labels: Vec<DropDownLabel>,
        display_text: String,
    },
    /// An inert inline script declaration.
    ///
    /// Linked targets and embedded payloads are retained as document metadata
    /// only. This API never opens, resolves, or executes either form.
    Script {
        /// Optional inert external script reference.
        href: Option<String>,
        /// Optional producer-supplied script-language identifier.
        language: Option<String>,
        /// The stored inline script payload, if any.
        content: String,
    },
    /// An inert table-cell formula display field.
    TableFormula {
        formula: Option<String>,
        display: Option<FormulaFieldDisplay>,
        data_style_name: Option<String>,
        display_text: String,
    },
    /// Cached, non-calculating measurement field text.
    Measure {
        kind: MeasureKind,
        display_text: String,
    },
    Reference {
        reference_name: Option<String>,
        reference_format: Option<CrossReferenceFormat>,
        display_text: String,
    },
    BookmarkReference {
        reference_name: Option<String>,
        reference_format: Option<CrossReferenceFormat>,
        display_text: String,
    },
    NoteReference {
        reference_name: Option<String>,
        note_class: NoteReferenceClass,
        reference_format: Option<NoteReferenceFormat>,
        display_text: String,
    },
    DocumentStatistic {
        kind: StatisticKind,
        number_format: Option<SequenceNumberFormat>,
        display_text: String,
    },
    /// Current, previous, or next page number with inert cached presentation.
    PageNumber {
        number_format: Option<SequenceNumberFormat>,
        fixed: Option<bool>,
        page_adjust: Option<i64>,
        select_page: Option<PageSelection>,
        display_text: String,
    },
    /// Current date or an explicitly fixed date/date-time value.
    Date {
        value: Option<FieldDateValue>,
        adjustment: Option<FieldDuration>,
        fixed: Option<bool>,
        data_style_name: Option<String>,
        display_text: String,
    },
    /// Current time or an explicitly fixed time/date-time value.
    Time {
        value: Option<FieldTimeValue>,
        adjustment: Option<FieldDuration>,
        fixed: Option<bool>,
        data_style_name: Option<String>,
        display_text: String,
    },
    /// Previous/next page continuation reminder.
    PageContinuation {
        select_page: PageContinuationSelection,
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
        number_format: Option<SequenceNumberFormat>,
        display_text: String,
    },
    /// Cached filename presentation; never reads a host path or document location.
    FileName {
        display: Option<FileNameDisplay>,
        fixed: Option<bool>,
        display_text: String,
    },
    /// Cached template presentation; never opens or locates a template resource.
    TemplateName {
        display: Option<TemplateNameDisplay>,
        display_text: String,
    },
    /// Cached active spreadsheet sheet label; never resolves live sheet state.
    SheetName { display_text: String },
    /// Cached chapter presentation; never resolves or updates the document outline.
    Chapter {
        display: Option<ChapterDisplay>,
        outline_level: Option<NonNegativeInteger>,
        display_text: String,
    },
    /// Cached presentation and optional fixed value of a metadata field.
    DocumentMetadata {
        kind: MetadataFieldKind,
        value: Option<MetadataFieldValue>,
        fixed: Option<bool>,
        data_style_name: Option<String>,
        display_text: String,
    },
    /// Fixed or live cached string metadata such as title or creator.
    ///
    /// Author fields retain stored text only and never read or modify host
    /// identity data.
    DocumentIdentity {
        kind: IdentityFieldKind,
        fixed: Option<bool>,
        display_text: String,
    },
    /// Cached subsequent-author identity/contact data.
    ///
    /// These fields never read or modify host identity or contact data, even when
    /// `text:fixed` is omitted or false.
    Sender {
        kind: SenderFieldKind,
        fixed: Option<bool>,
        display_text: String,
    },
    /// Named custom document metadata with inert cached typed attributes.
    UserDefinedMetadata {
        name: String,
        values: UserDefinedMetadataValues,
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
        content: MetaFieldContent,
    },
}

impl DynamicTextField {
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
            | Self::DropDown { display_text, .. }
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
            | Self::FileName { display_text, .. }
            | Self::TemplateName { display_text, .. }
            | Self::SheetName { display_text, .. }
            | Self::Chapter { display_text, .. }
            | Self::DocumentMetadata { display_text, .. }
            | Self::DocumentIdentity { display_text, .. }
            | Self::Sender { display_text, .. }
            | Self::UserDefinedMetadata { display_text, .. } => display_text,
            Self::Script { content, .. } => content,
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
            Self::DropDown {
                name,
                labels,
                display_text,
            } => {
                validate_dynamic_value("text:name", Some(name), false, &mut aggregate)?;
                if labels.len() > MAX_DROP_DOWN_LABELS {
                    return Err(Error::InvalidFormat(format!(
                        "text:drop-down exceeds {MAX_DROP_DOWN_LABELS} labels"
                    )));
                }
                for label in labels {
                    validate_dynamic_value(
                        "text:label text:value",
                        label.value.as_deref(),
                        false,
                        &mut aggregate,
                    )?;
                }
                validate_dynamic_value(
                    "drop-down display text",
                    Some(display_text),
                    false,
                    &mut aggregate,
                )?;
            },
            Self::Script {
                href,
                language,
                content,
            } => {
                validate_dynamic_value("xlink:href", href.as_deref(), false, &mut aggregate)?;
                validate_dynamic_value(
                    "script:language",
                    language.as_deref(),
                    false,
                    &mut aggregate,
                )?;
                validate_dynamic_value(
                    "inline script content",
                    Some(content),
                    false,
                    &mut aggregate,
                )?;
                if href.is_some() && !content.is_empty() {
                    return Err(Error::InvalidFormat(
                        "text:script cannot combine xlink:href with inline content".to_string(),
                    ));
                }
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
            Self::FileName { display_text, .. } => {
                validate_dynamic_value(
                    "file-name display text",
                    Some(display_text),
                    false,
                    &mut aggregate,
                )?;
            },
            Self::TemplateName { display_text, .. } => {
                validate_dynamic_value(
                    "template-name display text",
                    Some(display_text),
                    false,
                    &mut aggregate,
                )?;
            },
            Self::SheetName { display_text } => {
                validate_dynamic_value(
                    "sheet-name display text",
                    Some(display_text),
                    false,
                    &mut aggregate,
                )?;
            },
            Self::Chapter {
                outline_level,
                display_text,
                ..
            } => {
                validate_dynamic_value(
                    "text:outline-level",
                    outline_level.as_ref().map(NonNegativeInteger::as_str),
                    true,
                    &mut aggregate,
                )?;
                validate_dynamic_value(
                    "chapter display text",
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
            Self::Sender { display_text, .. } => {
                validate_dynamic_value(
                    "sender display text",
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
                let rebuilt = MetaFieldContent::new(content.nodes.clone())?;
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
        if let Self::Script {
            href,
            language,
            content,
        } = self
        {
            self.validate()?;
            let mut xml = String::from("<text:script xmlns:text=\"");
            xml.push_str(TEXT_DATABASE_NAMESPACE);
            xml.push('"');
            if let Some(href) = href {
                xml.push_str(" xmlns:xlink=\"");
                xml.push_str(XLINK_NAMESPACE);
                xml.push_str("\" xlink:type=\"simple\" xlink:href=\"");
                push_xml_attribute(&mut xml, href);
                xml.push('"');
            }
            if let Some(language) = language {
                xml.push_str(" xmlns:script=\"");
                xml.push_str(SCRIPT_NAMESPACE);
                xml.push_str("\" script:language=\"");
                push_xml_attribute(&mut xml, language);
                xml.push('"');
            }
            if content.is_empty() {
                xml.push_str("/>");
            } else {
                xml.push('>');
                push_xml_text(&mut xml, content);
                xml.push_str("</text:script>");
            }
            return Ok(xml);
        }
        if let Self::DropDown {
            name,
            labels,
            display_text,
        } = self
        {
            self.validate()?;
            let mut xml = String::from("<text:drop-down xmlns:text=\"");
            xml.push_str(TEXT_DATABASE_NAMESPACE);
            xml.push_str("\" text:name=\"");
            push_xml_attribute(&mut xml, name);
            xml.push('\"');
            if labels.is_empty() && display_text.is_empty() {
                xml.push_str("/>");
                return Ok(xml);
            }
            xml.push('>');
            for label in labels {
                xml.push_str("<text:label");
                if let Some(value) = label.value.as_deref() {
                    xml.push_str(" text:value=\"");
                    push_xml_attribute(&mut xml, value);
                    xml.push('\"');
                }
                if let Some(current_selected) = label.current_selected {
                    xml.push_str(" text:current-selected=\"");
                    xml.push_str(if current_selected { "true" } else { "false" });
                    xml.push('\"');
                }
                xml.push_str("/>");
            }
            push_xml_text(&mut xml, display_text);
            xml.push_str("</text:drop-down>");
            return Ok(xml);
        }
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
            Self::DropDown { .. } => unreachable!("drop-down uses nested-label serializer"),
            Self::Script { .. } => unreachable!("script uses a namespace-aware serializer"),
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
            Self::FileName { .. } => Element::new("text:file-name"),
            Self::TemplateName { .. } => Element::new("text:template-name"),
            Self::SheetName { .. } => Element::new("text:sheet-name"),
            Self::Chapter { .. } => Element::new("text:chapter"),
            Self::DocumentMetadata { kind, .. } => Element::new(kind.element_name()),
            Self::DocumentIdentity { kind, .. } => Element::new(kind.element_name()),
            Self::Sender { kind, .. } => Element::new(kind.element_name()),
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
            Self::DropDown { .. } => unreachable!("drop-down uses nested-label serializer"),
            Self::Script { .. } => unreachable!("script uses a namespace-aware serializer"),
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
            Self::FileName {
                display,
                fixed,
                display_text,
            } => {
                if let Some(display) = display {
                    element.set_attribute("text:display", display.as_str());
                }
                if let Some(fixed) = fixed {
                    element.set_attribute("text:fixed", if *fixed { "true" } else { "false" });
                }
                element.set_text(display_text);
            },
            Self::TemplateName {
                display,
                display_text,
            } => {
                if let Some(display) = display {
                    element.set_attribute("text:display", display.as_str());
                }
                element.set_text(display_text);
            },
            Self::SheetName { display_text } => {
                element.set_text(display_text);
            },
            Self::Chapter {
                display,
                outline_level,
                display_text,
            } => {
                if let Some(display) = display {
                    element.set_attribute("text:display", display.as_str());
                }
                if let Some(outline_level) = outline_level {
                    element.set_attribute("text:outline-level", outline_level.as_str());
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
                        MetadataFieldValue::Date(value) => {
                            element.set_attribute("text:date-value", value.as_str());
                        },
                        MetadataFieldValue::Time(value) => {
                            element.set_attribute("text:time-value", value.as_str());
                        },
                        MetadataFieldValue::Duration(value) => {
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
            Self::Sender {
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

pub(super) fn validate_dynamic_value(
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

pub(super) const fn is_xml_1_0_char(value: char) -> bool {
    matches!(value, '\u{9}' | '\u{A}' | '\u{D}')
        || (value as u32 >= 0x20 && value as u32 <= 0xD7FF)
        || (value as u32 >= 0xE000 && value as u32 <= 0xFFFD)
        || (value as u32 >= 0x10000 && value as u32 <= 0x10FFFF)
}

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
                | "text:drop-down"
                | "text:script"
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
    pub fn dynamic_text_field(&self) -> Result<Option<DynamicTextField>> {
        let text = || self.value();
        let result = match self.field_type() {
            "text:placeholder" => DynamicTextField::Placeholder {
                placeholder_type: PlaceholderType::parse(required_field_attribute(
                    self,
                    "text:placeholder-type",
                )?)?,
                description: self
                    .element
                    .get_attribute("text:description")
                    .map(str::to_owned),
                display_text: text(),
            },
            "text:conditional-text" => DynamicTextField::ConditionalText {
                condition: required_field_attribute(self, "text:condition")?.to_owned(),
                value_if_true: required_field_attribute(self, "text:string-value-if-true")?
                    .to_owned(),
                value_if_false: required_field_attribute(self, "text:string-value-if-false")?
                    .to_owned(),
                current_value: optional_field_bool(self, "text:current-value")?,
                display_text: text(),
            },
            "text:hidden-text" => DynamicTextField::HiddenText {
                condition: required_field_attribute(self, "text:condition")?.to_owned(),
                string_value: required_field_attribute(self, "text:string-value")?.to_owned(),
                is_hidden: optional_field_bool(self, "text:is-hidden")?,
                display_text: text(),
            },
            "text:hidden-paragraph" => DynamicTextField::HiddenParagraph {
                condition: required_field_attribute(self, "text:condition")?.to_owned(),
                is_hidden: optional_field_bool(self, "text:is-hidden")?,
                display_text: text(),
            },
            "text:script" => {
                reject_unknown_field_attributes(
                    self,
                    &["xlink:type", "xlink:href", "script:language"],
                )?;
                let href = match (
                    self.element.get_attribute("xlink:type"),
                    self.element.get_attribute("xlink:href"),
                ) {
                    (None, None) => None,
                    (Some("simple"), Some(href)) => Some(href.to_owned()),
                    (Some("simple"), None) => {
                        return Err(Error::InvalidFormat(
                            "text:script xlink:type requires xlink:href".to_string(),
                        ));
                    },
                    (None, Some(_)) => {
                        return Err(Error::InvalidFormat(
                            "text:script xlink:href requires xlink:type='simple'".to_string(),
                        ));
                    },
                    (Some(kind), Some(_)) => {
                        return Err(Error::InvalidFormat(format!(
                            "text:script xlink:type must be 'simple', got '{kind}'"
                        )));
                    },
                    (Some(kind), None) => {
                        return Err(Error::InvalidFormat(format!(
                            "text:script xlink:type must be 'simple', got '{kind}'"
                        )));
                    },
                };
                let result = DynamicTextField::Script {
                    href,
                    language: self
                        .element
                        .get_attribute("script:language")
                        .map(str::to_owned),
                    content: text(),
                };
                result.validate()?;
                result
            },
            "text:dde-connection" => DynamicTextField::DdeConnection {
                connection_name: required_field_attribute(self, "text:connection-name")?.to_owned(),
                display_text: text(),
            },
            "text:sequence" => {
                let format = self.element.get_attribute("style:num-format");
                let letter_sync = optional_field_bool(self, "style:num-letter-sync")?;
                let number_format = match (format, letter_sync) {
                    (Some(format), letter_sync) => {
                        Some(SequenceNumberFormat::new(format, letter_sync)?)
                    },
                    (None, Some(_)) => {
                        return Err(Error::InvalidFormat(
                            "style:num-letter-sync requires style:num-format".to_string(),
                        ));
                    },
                    (None, None) => None,
                };
                DynamicTextField::Sequence {
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
            "text:sequence-ref" => DynamicTextField::SequenceReference {
                reference_name: required_field_attribute(self, "text:ref-name")?.to_owned(),
                reference_format: self
                    .element
                    .get_attribute("text:reference-format")
                    .map(SequenceReferenceFormat::parse)
                    .transpose()?,
                display_text: text(),
            },
            "text:variable-set" => DynamicTextField::VariableSet {
                name: required_field_attribute(self, "text:name")?.to_owned(),
                formula: self
                    .element
                    .get_attribute("text:formula")
                    .map(str::to_owned),
                value: parse_calculated_value(self, true)?.expect("required calculated value"),
                display: self
                    .element
                    .get_attribute("text:display")
                    .map(VariableSetDisplay::parse)
                    .transpose()?,
                data_style_name: self
                    .element
                    .get_attribute("style:data-style-name")
                    .map(str::to_owned),
                display_text: text(),
            },
            "text:variable-get" => {
                reject_calculated_value_attributes(self)?;
                DynamicTextField::VariableGet {
                    name: required_field_attribute(self, "text:name")?.to_owned(),
                    display: self
                        .element
                        .get_attribute("text:display")
                        .map(FormulaFieldDisplay::parse)
                        .transpose()?,
                    data_style_name: self
                        .element
                        .get_attribute("style:data-style-name")
                        .map(str::to_owned),
                    display_text: text(),
                }
            },
            "text:expression" => DynamicTextField::Expression {
                formula: self
                    .element
                    .get_attribute("text:formula")
                    .map(str::to_owned),
                value: parse_calculated_value(self, false)?,
                display: self
                    .element
                    .get_attribute("text:display")
                    .map(FormulaFieldDisplay::parse)
                    .transpose()?,
                data_style_name: self
                    .element
                    .get_attribute("style:data-style-name")
                    .map(str::to_owned),
                display_text: text(),
            },
            "text:variable-input" => DynamicTextField::VariableInput {
                name: required_field_attribute(self, "text:name")?.to_owned(),
                description: self
                    .element
                    .get_attribute("text:description")
                    .map(str::to_owned),
                value_type: parse_value_type_only(self)?,
                display: self
                    .element
                    .get_attribute("text:display")
                    .map(VariableSetDisplay::parse)
                    .transpose()?,
                data_style_name: self
                    .element
                    .get_attribute("style:data-style-name")
                    .map(str::to_owned),
                display_text: text(),
            },
            "text:user-field-get" => {
                reject_calculated_value_attributes(self)?;
                DynamicTextField::UserFieldGet {
                    name: required_field_attribute(self, "text:name")?.to_owned(),
                    display: self
                        .element
                        .get_attribute("text:display")
                        .map(UserFieldDisplay::parse)
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
                DynamicTextField::UserFieldInput {
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
                DynamicTextField::TextInput {
                    description: self
                        .element
                        .get_attribute("text:description")
                        .map(str::to_owned),
                    display_text: text(),
                }
            },
            "text:table-formula" => {
                reject_calculated_value_attributes(self)?;
                DynamicTextField::TableFormula {
                    formula: self
                        .element
                        .get_attribute("text:formula")
                        .map(str::to_owned),
                    display: self
                        .element
                        .get_attribute("text:display")
                        .map(FormulaFieldDisplay::parse)
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
                DynamicTextField::Measure {
                    kind: MeasureKind::parse(required_field_attribute(self, "text:kind")?)?,
                    display_text: text(),
                }
            },
            "text:reference-ref" => {
                reject_unknown_field_attributes(self, &["text:ref-name", "text:reference-format"])?;
                DynamicTextField::Reference {
                    reference_name: self
                        .element
                        .get_attribute("text:ref-name")
                        .map(str::to_owned),
                    reference_format: self
                        .element
                        .get_attribute("text:reference-format")
                        .map(CrossReferenceFormat::parse)
                        .transpose()?,
                    display_text: text(),
                }
            },
            "text:bookmark-ref" => {
                reject_unknown_field_attributes(self, &["text:ref-name", "text:reference-format"])?;
                DynamicTextField::BookmarkReference {
                    reference_name: self
                        .element
                        .get_attribute("text:ref-name")
                        .map(str::to_owned),
                    reference_format: self
                        .element
                        .get_attribute("text:reference-format")
                        .map(CrossReferenceFormat::parse)
                        .transpose()?,
                    display_text: text(),
                }
            },
            "text:note-ref" => {
                reject_unknown_field_attributes(
                    self,
                    &["text:ref-name", "text:reference-format", "text:note-class"],
                )?;
                DynamicTextField::NoteReference {
                    reference_name: self
                        .element
                        .get_attribute("text:ref-name")
                        .map(str::to_owned),
                    note_class: NoteReferenceClass::parse(required_field_attribute(
                        self,
                        "text:note-class",
                    )?)?,
                    reference_format: self
                        .element
                        .get_attribute("text:reference-format")
                        .map(NoteReferenceFormat::parse)
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
                    "text:page-count" => StatisticKind::Page,
                    "text:paragraph-count" => StatisticKind::Paragraph,
                    "text:word-count" => StatisticKind::Word,
                    "text:character-count" => StatisticKind::Character,
                    "text:table-count" => StatisticKind::Table,
                    "text:image-count" => StatisticKind::Image,
                    "text:object-count" => StatisticKind::Object,
                    _ => unreachable!(),
                };
                DynamicTextField::DocumentStatistic {
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
                DynamicTextField::PageNumber {
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
                        .map(PageSelection::parse)
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
                DynamicTextField::Date {
                    value: self
                        .element
                        .get_attribute("text:date-value")
                        .map(FieldDateValue::new)
                        .transpose()?,
                    adjustment: self
                        .element
                        .get_attribute("text:date-adjust")
                        .map(FieldDuration::new)
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
                DynamicTextField::Time {
                    value: self
                        .element
                        .get_attribute("text:time-value")
                        .map(FieldTimeValue::new)
                        .transpose()?,
                    adjustment: self
                        .element
                        .get_attribute("text:time-adjust")
                        .map(FieldDuration::new)
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
                DynamicTextField::PageContinuation {
                    select_page: PageContinuationSelection::parse(required_field_attribute(
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
                DynamicTextField::PageVariableSet {
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
                DynamicTextField::PageVariableGet {
                    number_format: parse_common_number_format(self)?,
                    display_text: text(),
                }
            },
            "text:file-name" => {
                reject_unknown_field_attributes(self, &["text:display", "text:fixed"])?;
                let result = DynamicTextField::FileName {
                    display: self
                        .element
                        .get_attribute("text:display")
                        .map(FileNameDisplay::parse)
                        .transpose()?,
                    fixed: optional_field_bool(self, "text:fixed")?,
                    display_text: text(),
                };
                result.validate()?;
                result
            },
            "text:template-name" => {
                reject_unknown_field_attributes(self, &["text:display"])?;
                let result = DynamicTextField::TemplateName {
                    display: self
                        .element
                        .get_attribute("text:display")
                        .map(TemplateNameDisplay::parse)
                        .transpose()?,
                    display_text: text(),
                };
                result.validate()?;
                result
            },
            "text:sheet-name" => {
                reject_unknown_field_attributes(self, &[])?;
                let result = DynamicTextField::SheetName {
                    display_text: text(),
                };
                result.validate()?;
                result
            },
            "text:chapter" => {
                reject_unknown_field_attributes(self, &["text:display", "text:outline-level"])?;
                let result = DynamicTextField::Chapter {
                    display: self
                        .element
                        .get_attribute("text:display")
                        .map(ChapterDisplay::parse)
                        .transpose()?,
                    outline_level: self
                        .element
                        .get_attribute("text:outline-level")
                        .map(NonNegativeInteger::new)
                        .transpose()?,
                    display_text: text(),
                };
                result.validate()?;
                result
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
                    "text:creation-date" => MetadataFieldKind::CreationDate,
                    "text:creation-time" => MetadataFieldKind::CreationTime,
                    "text:print-date" => MetadataFieldKind::PrintDate,
                    "text:print-time" => MetadataFieldKind::PrintTime,
                    "text:editing-cycles" => MetadataFieldKind::EditingCycles,
                    "text:editing-duration" => MetadataFieldKind::EditingDuration,
                    "text:modification-date" => MetadataFieldKind::ModificationDate,
                    "text:modification-time" => MetadataFieldKind::ModificationTime,
                    _ => unreachable!(),
                };
                let allowed = match kind {
                    MetadataFieldKind::CreationDate
                    | MetadataFieldKind::PrintDate
                    | MetadataFieldKind::ModificationDate => {
                        &["text:fixed", "style:data-style-name", "text:date-value"][..]
                    },
                    MetadataFieldKind::CreationTime
                    | MetadataFieldKind::PrintTime
                    | MetadataFieldKind::ModificationTime => {
                        &["text:fixed", "style:data-style-name", "text:time-value"][..]
                    },
                    MetadataFieldKind::EditingDuration => {
                        &["text:fixed", "style:data-style-name", "text:duration"][..]
                    },
                    MetadataFieldKind::EditingCycles => &["text:fixed"][..],
                };
                reject_unknown_field_attributes(self, allowed)?;
                let value = match kind {
                    MetadataFieldKind::CreationDate
                    | MetadataFieldKind::PrintDate
                    | MetadataFieldKind::ModificationDate => self
                        .element
                        .get_attribute("text:date-value")
                        .map(FieldDateValue::new)
                        .transpose()?
                        .map(MetadataFieldValue::Date),
                    MetadataFieldKind::CreationTime
                    | MetadataFieldKind::PrintTime
                    | MetadataFieldKind::ModificationTime => self
                        .element
                        .get_attribute("text:time-value")
                        .map(FieldTimeValue::new)
                        .transpose()?
                        .map(MetadataFieldValue::Time),
                    MetadataFieldKind::EditingDuration => self
                        .element
                        .get_attribute("text:duration")
                        .map(FieldDuration::new)
                        .transpose()?
                        .map(MetadataFieldValue::Duration),
                    MetadataFieldKind::EditingCycles => None,
                };
                let result = DynamicTextField::DocumentMetadata {
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
                    "text:initial-creator" => IdentityFieldKind::InitialCreator,
                    "text:description" => IdentityFieldKind::Description,
                    "text:printed-by" => IdentityFieldKind::PrintedBy,
                    "text:title" => IdentityFieldKind::Title,
                    "text:subject" => IdentityFieldKind::Subject,
                    "text:keywords" => IdentityFieldKind::Keywords,
                    "text:creator" => IdentityFieldKind::Creator,
                    "text:author-name" => IdentityFieldKind::AuthorName,
                    "text:author-initials" => IdentityFieldKind::AuthorInitials,
                    _ => unreachable!(),
                };
                DynamicTextField::DocumentIdentity {
                    kind,
                    fixed: optional_field_bool(self, "text:fixed")?,
                    display_text: text(),
                }
            },
            "text:sender-firstname"
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
            | "text:sender-state-or-province" => {
                reject_unknown_field_attributes(self, &["text:fixed"])?;
                let kind = match self.field_type() {
                    "text:sender-firstname" => SenderFieldKind::FirstName,
                    "text:sender-lastname" => SenderFieldKind::LastName,
                    "text:sender-initials" => SenderFieldKind::Initials,
                    "text:sender-title" => SenderFieldKind::Title,
                    "text:sender-position" => SenderFieldKind::Position,
                    "text:sender-email" => SenderFieldKind::Email,
                    "text:sender-phone-private" => SenderFieldKind::PrivatePhone,
                    "text:sender-fax" => SenderFieldKind::Fax,
                    "text:sender-company" => SenderFieldKind::Company,
                    "text:sender-phone-work" => SenderFieldKind::WorkPhone,
                    "text:sender-street" => SenderFieldKind::Street,
                    "text:sender-city" => SenderFieldKind::City,
                    "text:sender-postal-code" => SenderFieldKind::PostalCode,
                    "text:sender-country" => SenderFieldKind::Country,
                    "text:sender-state-or-province" => SenderFieldKind::StateOrProvince,
                    _ => unreachable!(),
                };
                let result = DynamicTextField::Sender {
                    kind,
                    fixed: optional_field_bool(self, "text:fixed")?,
                    display_text: text(),
                };
                result.validate()?;
                result
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
                let result = DynamicTextField::UserDefinedMetadata {
                    name: required_field_attribute(self, "text:name")?.to_owned(),
                    values: UserDefinedMetadataValues {
                        number: self
                            .element
                            .get_attribute("office:value")
                            .map(str::to_owned),
                        date: self
                            .element
                            .get_attribute("office:date-value")
                            .map(FieldDateValue::new)
                            .transpose()?,
                        time: self
                            .element
                            .get_attribute("office:time-value")
                            .map(FieldDuration::new)
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
    kind: MetadataFieldKind,
    value: Option<&MetadataFieldValue>,
    aggregate: &mut usize,
) -> Result<()> {
    match (kind, value) {
        (_, None) => Ok(()),
        (MetadataFieldKind::CreationDate, Some(MetadataFieldValue::Date(value))) => {
            value.validate(aggregate)
        },
        (MetadataFieldKind::CreationTime, Some(MetadataFieldValue::Time(value))) => {
            value.validate(aggregate)
        },
        (
            MetadataFieldKind::PrintDate | MetadataFieldKind::ModificationDate,
            Some(MetadataFieldValue::Date(value)),
        ) if value.kind() == DateValueKind::Date => value.validate(aggregate),
        (
            MetadataFieldKind::PrintTime | MetadataFieldKind::ModificationTime,
            Some(MetadataFieldValue::Time(value)),
        ) if value.kind() == TimeValueKind::Time => value.validate(aggregate),
        (MetadataFieldKind::EditingDuration, Some(MetadataFieldValue::Duration(value))) => {
            value.validate("text:duration", aggregate)
        },
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

fn parse_calculated_value(field: &Field, required: bool) -> Result<Option<CalculatedFieldValue>> {
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
        "float" => CalculatedFieldValue::Float(required_attr("office:value")?.to_owned()),
        "percentage" => CalculatedFieldValue::Percentage(required_attr("office:value")?.to_owned()),
        "currency" => CalculatedFieldValue::Currency {
            value: required_attr("office:value")?.to_owned(),
            currency: attr("office:currency").map(str::to_owned),
        },
        "date" => CalculatedFieldValue::Date(required_attr("office:date-value")?.to_owned()),
        "time" => CalculatedFieldValue::Time(required_attr("office:time-value")?.to_owned()),
        "boolean" => CalculatedFieldValue::Boolean(
            optional_field_bool(field, "office:boolean-value")?.ok_or_else(|| {
                Error::InvalidFormat(
                    "office:value-type 'boolean' requires office:boolean-value".to_string(),
                )
            })?,
        ),
        "string" => CalculatedFieldValue::String(attr("office:string-value").map(str::to_owned)),
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

fn parse_value_type_only(field: &Field) -> Result<FieldValueType> {
    let value_type = FieldValueType::parse(required_field_attribute(field, "office:value-type")?)?;
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

fn parse_common_number_format(field: &Field) -> Result<Option<SequenceNumberFormat>> {
    let format = field.element.get_attribute("style:num-format");
    let letter_sync = optional_field_bool(field, "style:num-letter-sync")?;
    match (format, letter_sync) {
        (Some(format), letter_sync) => Ok(Some(SequenceNumberFormat::new(format, letter_sync)?)),
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

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum MetaContentGrammar {
    ParagraphOrHyperlink,
    Paragraph,
    TextOnly,
    DropDown,
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
    nodes: &[MetaFieldNode],
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
        MetaContentGrammar::DropDown => {
            return validate_meta_drop_down(nodes, depth, aggregate, node_count, display_text);
        },
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
            let MetaFieldNode::Element(element) = &nodes[0] else {
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
            MetaFieldNode::Text(value) => {
                if matches!(grammar, MetaContentGrammar::Empty) {
                    return Err(Error::InvalidFormat(
                        "ODF empty inline element contains character data".to_string(),
                    ));
                }
                validate_dynamic_value("meta-field text", Some(value), false, aggregate)?;
                display_text.push_str(value);
            },
            MetaFieldNode::Element(element) => {
                validate_meta_element_parts(
                    &element.namespace_uri,
                    &element.local_name,
                    &element.attributes,
                    aggregate,
                )?;
                let child_grammar =
                    meta_child_grammar(grammar, &element.namespace_uri, &element.local_name)?;
                validate_meta_element_attributes_for_grammar(element, child_grammar)?;
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

#[allow(clippy::too_many_arguments)]
fn validate_meta_exact_pair(
    nodes: &[MetaFieldNode],
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

#[allow(clippy::too_many_arguments)]
fn validate_meta_required_element(
    node: &MetaFieldNode,
    depth: usize,
    namespace: &str,
    local: &str,
    child_grammar: MetaContentGrammar,
    aggregate: &mut usize,
    node_count: &mut usize,
    display_text: &mut String,
) -> Result<()> {
    let MetaFieldNode::Element(element) = node else {
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
    validate_meta_element_attributes_for_grammar(element, child_grammar)?;
    validate_meta_nodes(
        &element.children,
        depth + 1,
        child_grammar,
        aggregate,
        node_count,
        display_text,
    )
}

fn validate_meta_element_attributes_for_grammar(
    element: &MetaFieldElement,
    grammar: MetaContentGrammar,
) -> Result<()> {
    if grammar == MetaContentGrammar::DropDown {
        validate_meta_drop_down_attributes(&element.attributes)?;
    }
    Ok(())
}

fn validate_meta_drop_down(
    nodes: &[MetaFieldNode],
    depth: usize,
    aggregate: &mut usize,
    node_count: &mut usize,
    display_text: &mut String,
) -> Result<()> {
    let mut display_started = false;
    let mut labels = 0usize;
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
            MetaFieldNode::Text(value) => {
                display_started = true;
                validate_dynamic_value("meta-field drop-down text", Some(value), false, aggregate)?;
                display_text.push_str(value);
            },
            MetaFieldNode::Element(element) => {
                if display_started
                    || element.namespace_uri != TEXT_DATABASE_NAMESPACE
                    || element.local_name != "label"
                {
                    return Err(Error::InvalidFormat(
                        "text:drop-down permits only leading text:label children".to_string(),
                    ));
                }
                if !element.children.is_empty() {
                    return Err(Error::InvalidFormat(
                        "text:label must be empty in text:drop-down".to_string(),
                    ));
                }
                labels = labels.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("text:drop-down label count overflow".to_string())
                })?;
                if labels > MAX_DROP_DOWN_LABELS {
                    return Err(Error::InvalidFormat(format!(
                        "text:drop-down exceeds {MAX_DROP_DOWN_LABELS} labels"
                    )));
                }
                validate_meta_element_parts(
                    &element.namespace_uri,
                    &element.local_name,
                    &element.attributes,
                    aggregate,
                )?;
                validate_meta_drop_down_label_attributes(&element.attributes)?;
            },
        }
    }
    if depth > MAX_META_FIELD_DEPTH {
        return Err(Error::InvalidFormat(format!(
            "text:meta-field content exceeds {MAX_META_FIELD_DEPTH} levels"
        )));
    }
    Ok(())
}

fn validate_meta_drop_down_attributes(attributes: &[MetaFieldAttribute]) -> Result<()> {
    let has_name = attributes.iter().any(|attribute| {
        attribute.namespace_uri == TEXT_DATABASE_NAMESPACE && attribute.local_name == "name"
    });
    if !has_name {
        return Err(Error::InvalidFormat(
            "text:drop-down requires text:name".to_string(),
        ));
    }
    if attributes.iter().any(|attribute| {
        attribute.namespace_uri != TEXT_DATABASE_NAMESPACE || attribute.local_name != "name"
    }) {
        return Err(Error::InvalidFormat(
            "text:drop-down only permits text:name".to_string(),
        ));
    }
    Ok(())
}

fn validate_meta_drop_down_label_attributes(attributes: &[MetaFieldAttribute]) -> Result<()> {
    for attribute in attributes {
        if attribute.namespace_uri != TEXT_DATABASE_NAMESPACE
            || !matches!(attribute.local_name.as_str(), "value" | "current-selected")
        {
            return Err(Error::InvalidFormat(
                "text:label has an unsupported attribute".to_string(),
            ));
        }
        if attribute.local_name == "current-selected" {
            parse_drop_down_boolean(&attribute.value)?;
        }
    }
    Ok(())
}

fn validate_meta_optional_listener_then(
    nodes: &[MetaFieldNode],
    depth: usize,
    remaining_grammar: MetaContentGrammar,
    owner: &str,
    aggregate: &mut usize,
    node_count: &mut usize,
    display_text: &mut String,
) -> Result<()> {
    let listener_position = nodes.iter().position(|node| {
        matches!(node, MetaFieldNode::Element(element)
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
    nodes: &[MetaFieldNode],
    depth: usize,
    aggregate: &mut usize,
    node_count: &mut usize,
    display_text: &mut String,
) -> Result<()> {
    for node in nodes {
        let MetaFieldNode::Element(element) = node else {
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
    nodes: &[MetaFieldNode],
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
        if matches!(nodes.get(position), Some(MetaFieldNode::Element(element))
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
        let MetaFieldNode::Element(element) = node else {
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

pub(super) fn meta_child_grammar(
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
            "drop-down" => Ok(MetaContentGrammar::DropDown),
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

pub(super) fn validate_meta_element_parts(
    namespace_uri: &str,
    local_name: &str,
    attributes: &[MetaFieldAttribute],
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

pub(super) fn validate_xml_id(value: &str) -> Result<()> {
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

pub(super) fn is_allowed_meta_namespace(namespace: &str) -> bool {
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

pub(super) fn add_meta_size(aggregate: &mut usize, amount: usize) -> Result<()> {
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

pub(super) fn write_meta_node(node: &MetaFieldNode, output: &mut String) {
    match node {
        MetaFieldNode::Text(value) => push_xml_text(output, value),
        MetaFieldNode::Element(element) => {
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

pub(super) fn push_xml_attribute(output: &mut String, value: &str) {
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

pub(super) fn push_xml_text(output: &mut String, value: &str) {
    for ch in value.chars() {
        match ch {
            '&' => output.push_str("&amp;"),
            '<' => output.push_str("&lt;"),
            '>' => output.push_str("&gt;"),
            _ => output.push(ch),
        }
    }
}

pub(super) fn parse_drop_down_boolean(value: &str) -> Result<bool> {
    match value {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        _ => Err(Error::InvalidFormat(format!(
            "invalid text:current-selected boolean '{value}'"
        ))),
    }
}
