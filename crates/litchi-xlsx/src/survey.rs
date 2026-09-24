//! Typed, inert `[MS-XLSX]` Survey-part inspection.
//!
//! Survey parts are associated with `SpreadsheetML` tables. They are parsed only
//! when requested and are never rendered, submitted, or otherwise activated.
//! The owning OPC package retains the original part bytes, so an unchanged
//! workbook save preserves survey XML and relationships losslessly.

use std::collections::{HashMap, HashSet};
use std::sync::Arc;

use litchi_ooxml_common::custom_xml::valid_guid;
use litchi_opc::constants::content_type as ct;
use litchi_opc::part::BlobPart;
use litchi_opc::{
    OpcPackage, OwnedContentTypes, OwnedRelationships, PackURI, Part as OpcPart, TargetMode,
};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;

use crate::error::{Error, Result, invalid};

mod source;

/// The `[MS-XLSX]` Survey-part content type.
pub const CONTENT_TYPE: &str = "application/vnd.ms-excel.Survey+xml";
/// The table-to-survey relationship type specified by `[MS-XLSX]` §2.1.9.
pub const RELATIONSHIP_TYPE: &str = "http://schemas.microsoft.com/office/2010/relationships/Survey";

const NAMESPACE: &[u8] = b"http://schemas.microsoft.com/office/spreadsheetml/2010/11/main";
const NAMESPACE_STR: &str = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main";
const MAX_XML_BYTES: usize = 16 * 1024 * 1024;
const MAX_TEXT_BYTES: usize = 1024 * 1024;
const MAX_QUESTIONS: usize = 65_534;
const MAX_DEPTH: usize = 256;
const MAX_SURVEYS: usize = 4_096;
const MAX_EXTENSION_BYTES: usize = 4 * 1024 * 1024;
const MAX_TOTAL_TEXT_BYTES: usize = 16 * 1024 * 1024;

/// Resource limits used while loading and authoring Survey parts.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    max_part_bytes: usize,
    max_surveys: usize,
    max_questions: usize,
    max_extension_bytes: usize,
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            max_part_bytes: MAX_XML_BYTES,
            max_surveys: MAX_SURVEYS,
            max_questions: MAX_QUESTIONS,
            max_extension_bytes: MAX_EXTENSION_BYTES,
        }
    }
}

impl Limits {
    /// Construct the default Survey resource policy.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            max_part_bytes: MAX_XML_BYTES,
            max_surveys: MAX_SURVEYS,
            max_questions: MAX_QUESTIONS,
            max_extension_bytes: MAX_EXTENSION_BYTES,
        }
    }

    /// Set the maximum Survey part size.
    #[must_use]
    pub const fn with_max_part_bytes(mut self, value: usize) -> Self {
        self.max_part_bytes = value;
        self
    }

    /// Set the maximum number of Survey parts.
    #[must_use]
    pub const fn with_max_surveys(mut self, value: usize) -> Self {
        self.max_surveys = value;
        self
    }

    /// Set the maximum number of questions in one Survey.
    #[must_use]
    pub const fn with_max_questions(mut self, value: usize) -> Self {
        self.max_questions = value;
        self
    }

    /// Set the maximum retained extension-list bytes in one Survey.
    #[must_use]
    pub const fn with_max_extension_bytes(mut self, value: usize) -> Self {
        self.max_extension_bytes = value;
        self
    }

    /// Maximum Survey part size.
    #[must_use]
    pub const fn max_part_bytes(self) -> usize {
        if self.max_part_bytes < MAX_XML_BYTES {
            self.max_part_bytes
        } else {
            MAX_XML_BYTES
        }
    }

    /// Maximum Survey count.
    #[must_use]
    pub const fn max_surveys(self) -> usize {
        if self.max_surveys < MAX_SURVEYS {
            self.max_surveys
        } else {
            MAX_SURVEYS
        }
    }

    /// Maximum question count per Survey.
    #[must_use]
    pub const fn max_questions(self) -> usize {
        if self.max_questions < MAX_QUESTIONS {
            self.max_questions
        } else {
            MAX_QUESTIONS
        }
    }

    /// Maximum extension bytes per Survey.
    #[must_use]
    pub const fn max_extension_bytes(self) -> usize {
        if self.max_extension_bytes < MAX_EXTENSION_BYTES {
            self.max_extension_bytes
        } else {
            MAX_EXTENSION_BYTES
        }
    }
}

/// A survey UID. It is an opaque native identifier, not a workbook selector.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct Id(u32);

impl Id {
    /// Create a survey UID from its native value.
    #[must_use]
    pub const fn new(value: u32) -> Self {
        Self(value)
    }

    /// Return the native value for diagnostics and interop.
    #[must_use]
    pub const fn get(self) -> u32 {
        self.0
    }
}

/// A table-column UID referenced by a survey question.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct Binding(u32);

impl Binding {
    /// Create a table-column UID from its native value.
    #[must_use]
    pub const fn new(value: u32) -> Self {
        Self(value)
    }

    /// Return the native value for diagnostics and interop.
    #[must_use]
    pub const fn get(self) -> u32 {
        self.0
    }
}

/// A braced OOXML GUID.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct Guid(Box<str>);

impl Guid {
    /// Parse a braced `ST_Guid` value.
    ///
    /// # Errors
    ///
    /// Returns an error when `value` is not a braced OOXML GUID.
    pub fn new(input: impl Into<Box<str>>) -> Result<Self> {
        let guid = input.into();
        if !valid_guid(&guid) {
            return Err(invalid(format!("invalid survey GUID '{guid}'")));
        }
        Ok(Self(guid))
    }

    /// The braced OOXML lexical representation.
    #[must_use]
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

/// Input control requested by a survey question.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum QuestionType {
    CheckBox,
    Choice,
    Date,
    Time,
    MultipleLinesOfText,
    Number,
    SingleLineOfText,
}

impl QuestionType {
    fn as_str(self) -> &'static str {
        match self {
            Self::CheckBox => "checkBox",
            Self::Choice => "choice",
            Self::Date => "date",
            Self::Time => "time",
            Self::MultipleLinesOfText => "multipleLinesOfText",
            Self::Number => "number",
            Self::SingleLineOfText => "singleLineOfText",
        }
    }

    fn parse(value: &str) -> Result<Self> {
        match value {
            "checkBox" => Ok(Self::CheckBox),
            "choice" => Ok(Self::Choice),
            "date" => Ok(Self::Date),
            "time" => Ok(Self::Time),
            "multipleLinesOfText" => Ok(Self::MultipleLinesOfText),
            "number" => Ok(Self::Number),
            "singleLineOfText" => Ok(Self::SingleLineOfText),
            _ => Err(invalid(format!("invalid survey question type '{value}'"))),
        }
    }
}

/// Display format requested by a survey question.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum QuestionFormat {
    GeneralDate,
    LongDate,
    ShortDate,
    LongTime,
    ShortTime,
    GeneralNumber,
    Standard,
    Fixed,
    Percent,
    Currency,
}

impl QuestionFormat {
    fn as_str(self) -> &'static str {
        match self {
            Self::GeneralDate => "generalDate",
            Self::LongDate => "longDate",
            Self::ShortDate => "shortDate",
            Self::LongTime => "longTime",
            Self::ShortTime => "shortTime",
            Self::GeneralNumber => "generalNumber",
            Self::Standard => "standard",
            Self::Fixed => "fixed",
            Self::Percent => "percent",
            Self::Currency => "currency",
        }
    }

    fn parse(value: &str) -> Result<Self> {
        match value {
            "generalDate" => Ok(Self::GeneralDate),
            "longDate" => Ok(Self::LongDate),
            "shortDate" => Ok(Self::ShortDate),
            "longTime" => Ok(Self::LongTime),
            "shortTime" => Ok(Self::ShortTime),
            "generalNumber" => Ok(Self::GeneralNumber),
            "standard" => Ok(Self::Standard),
            "fixed" => Ok(Self::Fixed),
            "percent" => Ok(Self::Percent),
            "currency" => Ok(Self::Currency),
            _ => Err(invalid(format!("invalid survey question format '{value}'"))),
        }
    }
}

/// CSS-compatible positioning of a survey element.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Position {
    Absolute,
    Fixed,
    Relative,
    Static,
    Inherit,
}

impl Position {
    fn as_str(self) -> &'static str {
        match self {
            Self::Absolute => "absolute",
            Self::Fixed => "fixed",
            Self::Relative => "relative",
            Self::Static => "static",
            Self::Inherit => "inherit",
        }
    }

    fn parse(value: &str) -> Result<Self> {
        match value {
            "absolute" => Ok(Self::Absolute),
            "fixed" => Ok(Self::Fixed),
            "relative" => Ok(Self::Relative),
            "static" => Ok(Self::Static),
            "inherit" => Ok(Self::Inherit),
            _ => Err(invalid(format!("invalid survey position '{value}'"))),
        }
    }
}

/// Optional presentational properties shared by survey elements.
#[derive(Debug, Clone, Default)]
pub struct ElementProperties {
    css_class: Option<Box<str>>,
    bottom: Option<i32>,
    top: Option<i32>,
    left: Option<i32>,
    right: Option<i32>,
    width: Option<u32>,
    height: Option<u32>,
    position: Option<Position>,
    extension_xml: Option<Box<[u8]>>,
    unknown_attributes: Box<[RawAttribute]>,
    namespace_declarations: Box<[RawAttribute]>,
    comments: Box<[Box<[u8]>]>,
    comment_slots: Box<[CommentSlot]>,
    default_namespace: Option<Box<str>>,
    prefix: Option<Box<str>>,
}

impl PartialEq for ElementProperties {
    fn eq(&self, other: &Self) -> bool {
        self.css_class == other.css_class
            && self.bottom == other.bottom
            && self.top == other.top
            && self.left == other.left
            && self.right == other.right
            && self.width == other.width
            && self.height == other.height
            && self.position == other.position
            && self.extension_xml == other.extension_xml
            && self.unknown_attributes == other.unknown_attributes
            && self.namespace_declarations == other.namespace_declarations
            && self.comments == other.comments
            && self.comment_slots == other.comment_slots
            && same_default_namespace(
                self.default_namespace.as_deref(),
                other.default_namespace.as_deref(),
            )
            && self.prefix == other.prefix
    }
}

impl Eq for ElementProperties {}

impl ElementProperties {
    /// Create empty element properties for authoring.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }
    /// Replace the optional CSS class.
    pub fn set_css_class(&mut self, value: Option<String>) {
        self.css_class = value.map(Into::into);
    }
    /// Replace the optional bottom boundary.
    pub const fn set_bottom(&mut self, value: Option<i32>) {
        self.bottom = value;
    }
    /// Replace the optional top boundary.
    pub const fn set_top(&mut self, value: Option<i32>) {
        self.top = value;
    }
    /// Replace the optional left boundary.
    pub const fn set_left(&mut self, value: Option<i32>) {
        self.left = value;
    }
    /// Replace the optional right boundary.
    pub const fn set_right(&mut self, value: Option<i32>) {
        self.right = value;
    }
    /// Replace the optional width.
    pub const fn set_width(&mut self, value: Option<u32>) {
        self.width = value;
    }
    /// Replace the optional height.
    pub const fn set_height(&mut self, value: Option<u32>) {
        self.height = value;
    }
    /// Replace the optional CSS positioning mode.
    pub const fn set_position(&mut self, value: Option<Position>) {
        self.position = value;
    }
    #[must_use]
    pub fn css_class(&self) -> Option<&str> {
        self.css_class.as_deref()
    }
    #[must_use]
    pub const fn bottom(&self) -> Option<i32> {
        self.bottom
    }
    #[must_use]
    pub const fn top(&self) -> Option<i32> {
        self.top
    }
    #[must_use]
    pub const fn left(&self) -> Option<i32> {
        self.left
    }
    #[must_use]
    pub const fn right(&self) -> Option<i32> {
        self.right
    }
    #[must_use]
    pub const fn width(&self) -> Option<u32> {
        self.width
    }
    #[must_use]
    pub const fn height(&self) -> Option<u32> {
        self.height
    }
    #[must_use]
    pub const fn position(&self) -> Option<Position> {
        self.position
    }

    /// Raw future-extension list retained from the source part.
    #[must_use]
    pub fn extension_xml(&self) -> Option<&[u8]> {
        self.extension_xml.as_deref()
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct RawAttribute {
    name: Box<str>,
    value: Box<str>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct CommentSlot {
    slot: usize,
    value: Box<[u8]>,
}

fn same_default_namespace(left: Option<&str>, right: Option<&str>) -> bool {
    left == right
        || (left == Some(NAMESPACE_STR) && right.is_none())
        || (right == Some(NAMESPACE_STR) && left.is_none())
}

fn normalized_survey_comment_slots(survey: &Survey) -> Vec<CommentSlot> {
    let removed = [
        (
            survey.source_properties_present && survey.properties.is_none(),
            0usize,
        ),
        (
            survey.source_title_properties_present && survey.title_properties.is_none(),
            1usize,
        ),
        (
            survey.source_description_properties_present && survey.description_properties.is_none(),
            2usize,
        ),
    ];
    survey
        .comment_slots
        .iter()
        .map(|comment| CommentSlot {
            slot: comment.slot.saturating_sub(
                removed
                    .iter()
                    .filter(|(present, index)| *present && *index < comment.slot)
                    .count(),
            ),
            value: comment.value.clone(),
        })
        .collect()
}

fn normalized_questions_comment_slots(questions: &Questions) -> Vec<CommentSlot> {
    let mut normalized = Vec::new();
    let mut used = vec![false; questions.comment_slots.len()];
    let mut emit = |source_slot: usize, output_slot: usize| {
        for (index, comment) in questions.comment_slots.iter().enumerate() {
            if comment.slot == source_slot {
                used[index] = true;
                normalized.push(CommentSlot {
                    slot: output_slot,
                    value: comment.value.clone(),
                });
            }
        }
    };
    let current_qproperties_present = questions.comment_qproperties_present;
    let source_qproperties_present = questions.source_qproperties_present;
    let source_question_count = questions.source_question_count;
    if source_qproperties_present || current_qproperties_present {
        emit(0, 0);
    }
    let source_slot_offset = usize::from(source_qproperties_present);
    let mut output_question_count = 0usize;
    for index in 0..source_question_count {
        let source_slot = index + source_slot_offset;
        if !(current_qproperties_present && !source_qproperties_present && index == 0) {
            emit(
                source_slot,
                output_question_count + usize::from(current_qproperties_present),
            );
        }
        if index < questions.values.len() {
            output_question_count += 1;
        }
    }
    output_question_count += questions.values.len().saturating_sub(source_question_count);
    emit(
        source_question_count + source_slot_offset,
        output_question_count + usize::from(current_qproperties_present),
    );
    for (index, comment) in questions.comment_slots.iter().enumerate() {
        if !used[index] {
            normalized.push(comment.clone());
        }
    }
    normalized
}

/// One table-bound survey question.
#[derive(Debug, Clone)]
pub struct Question {
    origin: Option<source::QuestionOrigin>,
    binding: Binding,
    column_name: Option<Box<str>>,
    text: Option<Box<str>>,
    kind: Option<QuestionType>,
    format: Option<QuestionFormat>,
    help_text: Option<Box<str>>,
    required: bool,
    default_value: Option<Box<str>>,
    decimal_places: Option<u8>,
    row_source: Option<Box<str>>,
    properties: Option<ElementProperties>,
    extension_xml: Option<Box<[u8]>>,
    unknown_attributes: Box<[RawAttribute]>,
    namespace_declarations: Box<[RawAttribute]>,
    comments: Box<[Box<[u8]>]>,
    comment_slots: Box<[CommentSlot]>,
    default_namespace: Option<Box<str>>,
    prefix: Option<Box<str>>,
}

impl Question {
    /// Create a question bound to one table-column UID.
    #[must_use]
    pub fn new(binding: Binding) -> Self {
        Self {
            origin: None,
            binding,
            column_name: None,
            text: None,
            kind: None,
            format: None,
            help_text: None,
            required: false,
            default_value: None,
            decimal_places: None,
            row_source: None,
            properties: None,
            extension_xml: None,
            unknown_attributes: Box::new([]),
            namespace_declarations: Box::new([]),
            comments: Box::new([]),
            comment_slots: Box::new([]),
            default_namespace: None,
            prefix: None,
        }
    }

    /// Create a question using an owning table-column name.
    ///
    /// The selector is resolved to the table's physical column UID by
    /// [`Transaction::insert_for_table`] or [`Transaction::edit`]. It cannot be serialized directly
    /// until that source-bound resolution has happened.
    pub fn for_column(name: impl Into<String>) -> Result<Self> {
        let name = name.into();
        validate_text(Some(&name), "column name")?;
        if name.is_empty() {
            return Err(invalid("survey column name cannot be empty"));
        }
        let mut question = Self::new(Binding::new(0));
        question.column_name = Some(name.into_boxed_str());
        Ok(question)
    }

    /// Alias for [`Self::for_column`] for constructor-style call sites.
    pub fn new_for_column(name: impl Into<String>) -> Result<Self> {
        Self::for_column(name)
    }
    /// Change the bound table-column UID.
    pub fn set_binding(&mut self, value: Binding) {
        self.binding = value;
        self.column_name = None;
    }
    /// Replace the question text.
    pub fn set_text(&mut self, value: Option<String>) {
        self.text = value.map(Into::into);
    }
    /// Replace the input type.
    pub const fn set_question_type(&mut self, value: Option<QuestionType>) {
        self.kind = value;
    }
    /// Replace the answer format.
    pub const fn set_format(&mut self, value: Option<QuestionFormat>) {
        self.format = value;
    }
    /// Replace the help text.
    pub fn set_help_text(&mut self, value: Option<String>) {
        self.help_text = value.map(Into::into);
    }
    /// Replace the required flag.
    pub const fn set_required(&mut self, value: bool) {
        self.required = value;
    }
    /// Replace the default answer.
    pub fn set_default_value(&mut self, value: Option<String>) {
        self.default_value = value.map(Into::into);
    }
    /// Replace the number of decimal places.
    pub const fn set_decimal_places(&mut self, value: Option<u8>) {
        self.decimal_places = value;
    }
    /// Replace the semicolon-delimited choice source. The MS-XLSX grammar
    /// permits quotes within unquoted values; this retains the string without
    /// inferring a unique choice-list interpretation.
    pub fn set_row_source(&mut self, value: Option<String>) -> Result<()> {
        if let Some(value) = value.as_deref() {
            validate_row_source(value)?;
        }
        self.row_source = value.map(Into::into);
        Ok(())
    }
    /// Replace the optional question properties.
    pub fn set_properties(&mut self, value: Option<ElementProperties>) {
        self.properties = value;
    }
    /// Raw future-extension list retained from the source part.
    #[must_use]
    pub fn extension_xml(&self) -> Option<&[u8]> {
        self.extension_xml.as_deref()
    }

    #[must_use]
    pub const fn binding(&self) -> Binding {
        self.binding
    }
    #[must_use]
    pub fn text(&self) -> Option<&str> {
        self.text.as_deref()
    }
    #[must_use]
    pub const fn question_type(&self) -> Option<QuestionType> {
        self.kind
    }
    #[must_use]
    pub const fn format(&self) -> Option<QuestionFormat> {
        self.format
    }
    #[must_use]
    pub fn help_text(&self) -> Option<&str> {
        self.help_text.as_deref()
    }
    #[must_use]
    pub const fn is_required(&self) -> bool {
        self.required
    }
    #[must_use]
    pub fn default_value(&self) -> Option<&str> {
        self.default_value.as_deref()
    }
    #[must_use]
    pub const fn decimal_places(&self) -> Option<u8> {
        self.decimal_places
    }
    #[must_use]
    pub fn row_source(&self) -> Option<&str> {
        self.row_source.as_deref()
    }
    #[must_use]
    pub fn properties(&self) -> Option<&ElementProperties> {
        self.properties.as_ref()
    }

    fn resolve_column(&mut self, table: &crate::table::Table) -> Result<()> {
        let Some(name) = self.column_name.as_deref() else {
            return Ok(());
        };
        let column = table
            .columns
            .iter()
            .find(|column| {
                column.name.eq_ignore_ascii_case(name)
                    || column
                        .unique_name
                        .as_deref()
                        .is_some_and(|unique_name| unique_name.eq_ignore_ascii_case(name))
            })
            .ok_or_else(|| invalid(format!("survey column '{name}' is absent from the table")))?;
        self.binding = Binding::new(column.id);
        self.column_name = None;
        Ok(())
    }
}

// Source lineage is publication metadata, not part of question semantics.
impl PartialEq for Question {
    fn eq(&self, other: &Self) -> bool {
        self.binding == other.binding
            && self.column_name == other.column_name
            && self.text == other.text
            && self.kind == other.kind
            && self.format == other.format
            && self.help_text == other.help_text
            && self.required == other.required
            && self.default_value == other.default_value
            && self.decimal_places == other.decimal_places
            && self.row_source == other.row_source
            && self.properties == other.properties
            && self.extension_xml == other.extension_xml
            && self.unknown_attributes == other.unknown_attributes
            && self.namespace_declarations == other.namespace_declarations
            && self.comments == other.comments
            && self.comment_slots == other.comment_slots
            && same_default_namespace(
                self.default_namespace.as_deref(),
                other.default_namespace.as_deref(),
            )
            && self.prefix == other.prefix
    }
}

impl Eq for Question {}

/// The ordered survey-question collection.
#[derive(Debug, Clone)]
pub struct Questions {
    properties: Option<ElementProperties>,
    values: Vec<Question>,
    comments: Box<[Box<[u8]>]>,
    comment_slots: Box<[CommentSlot]>,
    comment_qproperties_present: bool,
    source_qproperties_present: bool,
    source_question_count: usize,
    default_namespace: Option<Box<str>>,
    prefix: Option<Box<str>>,
    namespace_declarations: Box<[RawAttribute]>,
    unknown_attributes: Box<[RawAttribute]>,
}

impl PartialEq for Questions {
    fn eq(&self, other: &Self) -> bool {
        self.properties == other.properties
            && self.values == other.values
            && self.comments == other.comments
            && normalized_questions_comment_slots(self) == normalized_questions_comment_slots(other)
            && same_default_namespace(
                self.default_namespace.as_deref(),
                other.default_namespace.as_deref(),
            )
            && self.prefix == other.prefix
            && self.namespace_declarations == other.namespace_declarations
            && self.unknown_attributes == other.unknown_attributes
    }
}

impl Eq for Questions {}

impl Questions {
    /// Create an empty question collection for authoring.
    #[must_use]
    pub fn new() -> Self {
        Self {
            properties: None,
            values: Vec::new(),
            comments: Box::new([]),
            comment_slots: Box::new([]),
            comment_qproperties_present: false,
            source_qproperties_present: false,
            source_question_count: 0,
            default_namespace: None,
            prefix: None,
            namespace_declarations: Box::new([]),
            unknown_attributes: Box::new([]),
        }
    }

    #[must_use]
    pub fn properties(&self) -> Option<&ElementProperties> {
        self.properties.as_ref()
    }
    #[must_use]
    pub fn values(&self) -> &[Question] {
        &self.values
    }
    /// Borrow the staged question collection mutably.
    pub fn values_mut(&mut self) -> &mut [Question] {
        &mut self.values
    }
    /// Append one question, enforcing the protocol cardinality bound.
    pub fn push(&mut self, question: Question) -> Result<()> {
        self.insert(self.values.len(), question)
    }
    /// Insert a question at a checked zero-based position in this staged list.
    pub fn insert(&mut self, index: usize, question: Question) -> Result<()> {
        if index > self.values.len() {
            return Err(invalid(
                "survey question insertion position is out of bounds",
            ));
        }
        if self.values.len() >= MAX_QUESTIONS {
            return Err(invalid("survey has too many questions"));
        }
        self.values
            .try_reserve(1)
            .map_err(|_| invalid("survey question allocation failed"))?;
        self.values.insert(index, question);
        Ok(())
    }
    /// Remove one question by zero-based position.
    pub fn remove(&mut self, index: usize) -> Option<Question> {
        if index >= self.values.len() {
            return None;
        }
        Some(self.values.remove(index))
    }
    /// Replace the optional collection properties.
    pub fn set_properties(&mut self, value: Option<ElementProperties>) {
        self.properties = value;
        self.comment_qproperties_present = self.properties.is_some();
    }
}

impl Default for Questions {
    fn default() -> Self {
        Self::new()
    }
}

/// A parsed survey part. It is intentionally read-only and inert.
#[derive(Debug, Clone)]
pub struct Survey {
    id: Id,
    guid: Guid,
    title: Option<Box<str>>,
    description: Option<Box<str>>,
    properties: Option<ElementProperties>,
    title_properties: Option<ElementProperties>,
    description_properties: Option<ElementProperties>,
    questions: Questions,
    extension_xml: Option<Box<[u8]>>,
    unknown_attributes: Box<[RawAttribute]>,
    namespace_declarations: Box<[RawAttribute]>,
    comments: Box<[Box<[u8]>]>,
    comment_slots: Box<[CommentSlot]>,
    source_properties_present: bool,
    source_title_properties_present: bool,
    source_description_properties_present: bool,
    default_namespace: Option<Box<str>>,
    root_prefix: Option<Box<str>>,
}

impl PartialEq for Survey {
    fn eq(&self, other: &Self) -> bool {
        self.id == other.id
            && self.guid == other.guid
            && self.title == other.title
            && self.description == other.description
            && self.properties == other.properties
            && self.title_properties == other.title_properties
            && self.description_properties == other.description_properties
            && self.questions == other.questions
            && self.extension_xml == other.extension_xml
            && self.unknown_attributes == other.unknown_attributes
            && self.namespace_declarations == other.namespace_declarations
            && self.comments == other.comments
            && normalized_survey_comment_slots(self) == normalized_survey_comment_slots(other)
            && self.default_namespace == other.default_namespace
            && self.root_prefix == other.root_prefix
    }
}

impl Eq for Survey {}

impl Survey {
    /// Create a Survey with its required question collection.
    pub fn new(id: Id, guid: Guid, questions: Vec<Question>) -> Result<Self> {
        let question_count = questions.len();
        let survey = Self {
            id,
            guid,
            title: None,
            description: None,
            properties: None,
            title_properties: None,
            description_properties: None,
            questions: Questions {
                properties: None,
                values: questions,
                comments: Box::new([]),
                comment_slots: Box::new([]),
                comment_qproperties_present: false,
                source_qproperties_present: false,
                source_question_count: question_count,
                default_namespace: None,
                prefix: None,
                namespace_declarations: Box::new([]),
                unknown_attributes: Box::new([]),
            },
            extension_xml: None,
            unknown_attributes: Box::new([]),
            namespace_declarations: Box::new([]),
            comments: Box::new([]),
            comment_slots: Box::new([]),
            source_properties_present: false,
            source_title_properties_present: false,
            source_description_properties_present: false,
            default_namespace: Some(NAMESPACE_STR.into()),
            root_prefix: None,
        };
        validate_survey(&survey)?;
        Ok(survey)
    }
    /// Change the Survey UID.
    pub const fn set_id(&mut self, value: Id) {
        self.id = value;
    }
    /// Change the Survey GUID.
    pub fn set_guid(&mut self, value: Guid) {
        self.guid = value;
    }
    /// Replace the title.
    pub fn set_title(&mut self, value: Option<String>) {
        self.title = value.map(Into::into);
    }
    /// Replace the description.
    pub fn set_description(&mut self, value: Option<String>) {
        self.description = value.map(Into::into);
    }
    /// Replace survey-level properties.
    pub fn set_properties(&mut self, value: Option<ElementProperties>) {
        self.properties = value;
    }
    /// Replace title properties.
    pub fn set_title_properties(&mut self, value: Option<ElementProperties>) {
        self.title_properties = value;
    }
    /// Replace description properties.
    pub fn set_description_properties(&mut self, value: Option<ElementProperties>) {
        self.description_properties = value;
    }
    /// Borrow the question collection mutably.
    pub fn questions_mut(&mut self) -> &mut Questions {
        &mut self.questions
    }
    /// Raw future-extension list retained from the source part.
    #[must_use]
    pub fn extension_xml(&self) -> Option<&[u8]> {
        self.extension_xml.as_deref()
    }

    #[must_use]
    pub const fn id(&self) -> Id {
        self.id
    }
    #[must_use]
    pub fn guid(&self) -> &Guid {
        &self.guid
    }
    #[must_use]
    pub fn title(&self) -> Option<&str> {
        self.title.as_deref()
    }
    #[must_use]
    pub fn description(&self) -> Option<&str> {
        self.description.as_deref()
    }
    #[must_use]
    pub fn properties(&self) -> Option<&ElementProperties> {
        self.properties.as_ref()
    }
    #[must_use]
    pub fn title_properties(&self) -> Option<&ElementProperties> {
        self.title_properties.as_ref()
    }
    #[must_use]
    pub fn description_properties(&self) -> Option<&ElementProperties> {
        self.description_properties.as_ref()
    }
    #[must_use]
    pub fn questions(&self) -> &Questions {
        &self.questions
    }
}

/// A relationship edge captured as part of the Survey source closure.
///
/// The OPC relationship type and lexical target are both retained.  A
/// source-bound transaction must reject a package whose table or Survey
/// relationship graph changed underneath it, even when the typed Survey XML
/// still parses to the same values.
#[derive(Debug, Clone, PartialEq, Eq)]
struct RelationshipState {
    source: PackURI,
    id: String,
    reltype: String,
    target: String,
    mode: TargetMode,
}

/// One survey attached to a table in an immutable workbook snapshot.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Part {
    survey: Arc<Survey>,
    table_part_name: PackURI,
    relationship_id: String,
    relationship_target: String,
    part_name: PackURI,
    source_xml: Arc<Vec<u8>>,
    source_proof: Option<Arc<litchi_opc::OwnedXmlPart>>,
    table_source_xml: Arc<Vec<u8>>,
    table_relationships: Arc<[RelationshipState]>,
    survey_relationships: Arc<[RelationshipState]>,
    /// Exact source relationship member owned by the table, including an
    /// explicitly present empty member.
    table_relationships_source: Option<OwnedRelationships>,
    /// Exact source relationship member owned by the Survey part, including
    /// an explicitly present empty member.
    survey_relationships_source: Option<OwnedRelationships>,
}

impl Part {
    /// Typed, inert survey contents. The associated table remains unchanged.
    #[must_use]
    pub fn survey(&self) -> &Survey {
        &self.survey
    }
    /// Table part owning this Survey relationship.
    #[must_use]
    pub fn table_part_name(&self) -> &PackURI {
        &self.table_part_name
    }
    /// Relationship ID on the owning table part.
    #[must_use]
    pub fn relationship_id(&self) -> &str {
        &self.relationship_id
    }
    /// Lexical target reference stored on the owning table relationship.
    #[must_use]
    pub fn relationship_target(&self) -> &str {
        &self.relationship_target
    }
    /// Physical Survey part name.
    #[must_use]
    pub fn part_name(&self) -> &PackURI {
        &self.part_name
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Scope {
    Survey,
    Properties(&'static [u8]),
    Questions,
    Question,
    Extension,
}

#[derive(Clone, Copy)]
enum ExtensionTarget {
    Survey,
    SurveyProperties,
    QuestionsProperties,
    TitleProperties,
    DescriptionProperties,
    QuestionProperties,
    Question,
}

#[derive(Default)]
struct Parser {
    scopes: Vec<Scope>,
    survey: Option<SurveyBuilder>,
    seen_root: bool,
    closed_root: bool,
    question_limit: usize,
}

#[derive(Default)]
struct SurveyBuilder {
    id: Option<Id>,
    guid: Option<Guid>,
    title: Option<Box<str>>,
    description: Option<Box<str>>,
    properties: Option<ElementProperties>,
    title_properties: Option<ElementProperties>,
    description_properties: Option<ElementProperties>,
    questions_properties: Option<ElementProperties>,
    questions: Vec<Question>,
    questions_comments: Vec<Box<[u8]>>,
    questions_comment_slots: Vec<CommentSlot>,
    questions_default_namespace: Option<Box<str>>,
    questions_prefix: Option<Box<str>>,
    questions_namespaces: Box<[RawAttribute]>,
    questions_attributes: Box<[RawAttribute]>,
    questions_seen: bool,
    child_stage: u8,
    extension_xml: Option<Box<[u8]>>,
    unknown_attributes: Box<[RawAttribute]>,
    namespace_declarations: Box<[RawAttribute]>,
    comments: Vec<Box<[u8]>>,
    comment_slots: Vec<CommentSlot>,
    default_namespace: Option<Box<str>>,
    root_prefix: Option<Box<str>>,
}

impl Parser {
    fn advance_survey(&mut self, stage: u8) -> Result<()> {
        let survey = self.survey_mut()?;
        if stage <= survey.child_stage {
            return Err(invalid("out-of-order Survey child"));
        }
        survey.child_stage = stage;
        Ok(())
    }
    fn in_extension(&self) -> bool {
        self.scopes.contains(&Scope::Extension)
    }
    fn start(&mut self, namespace: &ResolveResult<'_>, element: &BytesStart<'_>) -> Result<()> {
        if self.scopes.len() >= MAX_DEPTH {
            return Err(invalid("survey XML nesting is too deep"));
        }
        let scope = self.begin(namespace, element)?;
        self.scopes.push(scope);
        Ok(())
    }
    fn extension_target(
        &mut self,
        namespace: &ResolveResult<'_>,
        element: &BytesStart<'_>,
    ) -> Result<Option<ExtensionTarget>> {
        if element.local_name().as_ref() != b"extLst" || !is_extension_namespace(namespace) {
            return Ok(None);
        }
        let target = match self.scopes.last().copied() {
            Some(Scope::Survey) => {
                if !self.survey_mut()?.questions_seen {
                    return Err(invalid("Survey extLst must follow questions"));
                }
                self.advance_survey(5)?;
                ExtensionTarget::Survey
            },
            Some(Scope::Question) => ExtensionTarget::Question,
            Some(Scope::Properties(name)) => match name {
                b"surveyPr" => ExtensionTarget::SurveyProperties,
                b"titlePr" => ExtensionTarget::TitleProperties,
                b"descriptionPr" => ExtensionTarget::DescriptionProperties,
                b"questionsPr" => ExtensionTarget::QuestionsProperties,
                b"questionPr" => ExtensionTarget::QuestionProperties,
                _ => return Err(invalid("invalid Survey extension-list owner")),
            },
            Some(Scope::Questions) => {
                return Err(invalid("Survey questions cannot contain extLst"));
            },
            Some(Scope::Extension) | None => {
                return Err(invalid("invalid Survey extension-list location"));
            },
        };
        Ok(Some(target))
    }
    fn attach_extension(&mut self, target: ExtensionTarget, raw: Vec<u8>) -> Result<()> {
        if raw.len() > MAX_EXTENSION_BYTES {
            return Err(invalid("survey extension list exceeds the size limit"));
        }
        match target {
            ExtensionTarget::Survey => {
                set_once(
                    &mut self.survey_mut()?.extension_xml,
                    raw.into_boxed_slice(),
                    "survey extLst",
                )?;
            },
            ExtensionTarget::SurveyProperties => {
                let properties = self
                    .survey_mut()?
                    .properties
                    .as_mut()
                    .ok_or_else(|| invalid("surveyPr extension list is outside surveyPr"))?;
                set_once(
                    &mut properties.extension_xml,
                    raw.into_boxed_slice(),
                    "surveyPr extLst",
                )?;
            },
            ExtensionTarget::TitleProperties => {
                let properties = self
                    .survey_mut()?
                    .title_properties
                    .as_mut()
                    .ok_or_else(|| invalid("titlePr extension list is outside titlePr"))?;
                set_once(
                    &mut properties.extension_xml,
                    raw.into_boxed_slice(),
                    "titlePr extLst",
                )?;
            },
            ExtensionTarget::DescriptionProperties => {
                let properties = self
                    .survey_mut()?
                    .description_properties
                    .as_mut()
                    .ok_or_else(|| {
                        invalid("descriptionPr extension list is outside descriptionPr")
                    })?;
                set_once(
                    &mut properties.extension_xml,
                    raw.into_boxed_slice(),
                    "descriptionPr extLst",
                )?;
            },
            ExtensionTarget::QuestionsProperties => {
                let properties = self
                    .survey_mut()?
                    .questions_properties
                    .as_mut()
                    .ok_or_else(|| invalid("questionsPr extension list is outside questionsPr"))?;
                set_once(
                    &mut properties.extension_xml,
                    raw.into_boxed_slice(),
                    "questionsPr extLst",
                )?;
            },
            ExtensionTarget::QuestionProperties => {
                let question = self
                    .survey_mut()?
                    .questions
                    .last_mut()
                    .ok_or_else(|| invalid("questionPr extension list is outside question"))?;
                let properties = question
                    .properties
                    .as_mut()
                    .ok_or_else(|| invalid("questionPr extension list is outside questionPr"))?;
                set_once(
                    &mut properties.extension_xml,
                    raw.into_boxed_slice(),
                    "questionPr extLst",
                )?;
            },
            ExtensionTarget::Question => {
                let question = self
                    .survey_mut()?
                    .questions
                    .last_mut()
                    .ok_or_else(|| invalid("question extension list is outside question"))?;
                set_once(
                    &mut question.extension_xml,
                    raw.into_boxed_slice(),
                    "question extLst",
                )?;
            },
        }
        Ok(())
    }
    fn empty(&mut self, namespace: &ResolveResult<'_>, element: &BytesStart<'_>) -> Result<()> {
        let scope = self.begin(namespace, element)?;
        match scope {
            Scope::Survey | Scope::Questions => {
                return Err(invalid("survey XML has an empty required container"));
            },
            Scope::Properties(_) | Scope::Question | Scope::Extension => {},
        }
        Ok(())
    }
    fn end(&mut self, namespace: &ResolveResult<'_>, name: &[u8]) -> Result<()> {
        let scope = self
            .scopes
            .pop()
            .ok_or_else(|| invalid("unexpected survey end element"))?;
        let expected = match scope {
            Scope::Survey => b"survey".as_slice(),
            Scope::Properties(expected) => expected,
            Scope::Questions => b"questions".as_slice(),
            Scope::Question => b"question".as_slice(),
            Scope::Extension => return Ok(()),
        };
        if !is_survey_namespace(namespace) || name != expected {
            return Err(invalid("mismatched survey end element"));
        }
        if scope == Scope::Survey {
            self.closed_root = true;
        }
        Ok(())
    }
    fn comment(&mut self, comment: quick_xml::events::BytesText<'_>) -> Result<()> {
        let raw = serialize_event(Event::Comment(comment.into_owned()))?.into_boxed_slice();
        let Some(scope) = self.scopes.last().copied() else {
            return Ok(());
        };
        match scope {
            Scope::Survey => {
                let survey = self.survey_mut()?;
                let slot = survey.child_stage as usize;
                survey.comments.push(raw.clone());
                survey.comment_slots.push(CommentSlot { slot, value: raw });
            },
            Scope::Questions => {
                let survey = self.survey_mut()?;
                let slot =
                    survey.questions.len() + usize::from(survey.questions_properties.is_some());
                survey.questions_comments.push(raw.clone());
                survey
                    .questions_comment_slots
                    .push(CommentSlot { slot, value: raw });
            },
            Scope::Question => {
                let question = self
                    .survey_mut()?
                    .questions
                    .last_mut()
                    .ok_or_else(|| invalid("survey comment is outside question"))?;
                question.comments = question
                    .comments
                    .iter()
                    .cloned()
                    .chain(std::iter::once(raw.clone()))
                    .collect();
                let slot = if question.extension_xml.is_some() {
                    2
                } else {
                    usize::from(question.properties.is_some())
                };
                question.comment_slots = question
                    .comment_slots
                    .iter()
                    .cloned()
                    .chain(std::iter::once(CommentSlot { slot, value: raw }))
                    .collect();
            },
            Scope::Properties(name) => {
                let survey = self.survey_mut()?;
                let properties = match name {
                    b"surveyPr" => survey.properties.as_mut(),
                    b"titlePr" => survey.title_properties.as_mut(),
                    b"descriptionPr" => survey.description_properties.as_mut(),
                    b"questionsPr" => survey.questions_properties.as_mut(),
                    b"questionPr" => survey
                        .questions
                        .last_mut()
                        .and_then(|question| question.properties.as_mut()),
                    _ => None,
                }
                .ok_or_else(|| invalid("survey comment is outside properties"))?;
                properties.comments = properties
                    .comments
                    .iter()
                    .cloned()
                    .chain(std::iter::once(raw.clone()))
                    .collect();
                let slot = usize::from(properties.extension_xml.is_some());
                properties.comment_slots = properties
                    .comment_slots
                    .iter()
                    .cloned()
                    .chain(std::iter::once(CommentSlot { slot, value: raw }))
                    .collect();
            },
            Scope::Extension => {},
        }
        Ok(())
    }
    fn begin(&mut self, namespace: &ResolveResult<'_>, element: &BytesStart<'_>) -> Result<Scope> {
        let local_name = element.local_name();
        let name = local_name.as_ref();
        if self.in_extension() {
            return Ok(Scope::Extension);
        }
        if self.scopes.is_empty() {
            if self.seen_root || name != b"survey" || !is_survey_namespace(namespace) {
                return Err(invalid(
                    "survey part requires a Survey-namespace survey root",
                ));
            }
            self.seen_root = true;
            self.survey = Some(parse_survey(element)?);
            return Ok(Scope::Survey);
        }
        if name == b"extLst" && is_extension_namespace(namespace) {
            return Ok(Scope::Extension);
        }
        if !is_survey_namespace(namespace) {
            return Err(invalid("survey element is outside the Survey namespace"));
        }
        let parent = self
            .scopes
            .last()
            .copied()
            .ok_or_else(|| invalid("survey child has no parent element"))?;
        match (parent, name) {
            (Scope::Survey, b"surveyPr") => {
                self.advance_survey(1)?;
                set_once(
                    &mut self.survey_mut()?.properties,
                    parse_properties(element)?,
                    "surveyPr",
                )?;
                Ok(Scope::Properties(b"surveyPr"))
            },
            (Scope::Survey, b"titlePr") => {
                self.advance_survey(2)?;
                set_once(
                    &mut self.survey_mut()?.title_properties,
                    parse_properties(element)?,
                    "titlePr",
                )?;
                Ok(Scope::Properties(b"titlePr"))
            },
            (Scope::Survey, b"descriptionPr") => {
                self.advance_survey(3)?;
                set_once(
                    &mut self.survey_mut()?.description_properties,
                    parse_properties(element)?,
                    "descriptionPr",
                )?;
                Ok(Scope::Properties(b"descriptionPr"))
            },
            (Scope::Survey, b"questions") => {
                self.advance_survey(4)?;
                let survey = self.survey_mut()?;
                if survey.questions_seen {
                    return Err(invalid("survey has multiple questions elements"));
                }
                survey.questions_seen = true;
                survey.questions_namespaces = namespace_declarations(element)?;
                survey.questions_attributes = unknown_attributes(element, &[])?;
                survey.questions_default_namespace =
                    attr(&attrs(element)?, b"xmlns").map(Into::into);
                survey.questions_prefix = element_prefix(element)?;
                Ok(Scope::Questions)
            },
            (Scope::Questions, b"questionsPr") => {
                if !self.survey_mut()?.questions.is_empty() {
                    return Err(invalid("questionsPr must precede question"));
                }
                set_once(
                    &mut self.survey_mut()?.questions_properties,
                    parse_properties(element)?,
                    "questionsPr",
                )?;
                Ok(Scope::Properties(b"questionsPr"))
            },
            (Scope::Questions, b"question") => {
                let maximum = self.question_limit;
                let survey = self.survey_mut()?;
                if survey.questions.len() >= maximum {
                    return Err(invalid("survey has too many questions"));
                }
                survey.questions.push(parse_question(element)?);
                Ok(Scope::Question)
            },
            (Scope::Question, b"questionPr") => {
                let properties = parse_properties(element)?;
                let question = self
                    .survey_mut()?
                    .questions
                    .last_mut()
                    .ok_or_else(|| invalid("questionPr is outside a question"))?;
                if question.extension_xml.is_some() {
                    return Err(invalid("questionPr must precede extLst"));
                }
                set_once(&mut question.properties, properties, "questionPr")?;
                Ok(Scope::Properties(b"questionPr"))
            },
            _ => Err(invalid(format!(
                "unexpected survey element '{}'",
                String::from_utf8_lossy(name)
            ))),
        }
    }
    fn survey_mut(&mut self) -> Result<&mut SurveyBuilder> {
        self.survey
            .as_mut()
            .ok_or_else(|| invalid("survey root is missing"))
    }
    fn finish(self) -> Result<Survey> {
        if !self.seen_root || !self.scopes.is_empty() || !self.closed_root {
            return Err(invalid("unterminated survey XML"));
        }
        let mut value = self
            .survey
            .ok_or_else(|| invalid("survey root is missing"))?;
        let id = value.id.ok_or_else(|| invalid("survey id is required"))?;
        let guid = value
            .guid
            .ok_or_else(|| invalid("survey guid is required"))?;
        if !value.questions_seen || value.questions.is_empty() {
            return Err(invalid("survey requires at least one question"));
        }
        source::assign_question_origins(&mut value.questions);
        let comment_qproperties_present = value.questions_properties.is_some();
        let source_question_count = value.questions.len();
        let source_properties_present = value.properties.is_some();
        let source_title_properties_present = value.title_properties.is_some();
        let source_description_properties_present = value.description_properties.is_some();
        Ok(Survey {
            id,
            guid,
            title: value.title,
            description: value.description,
            properties: value.properties,
            title_properties: value.title_properties,
            description_properties: value.description_properties,
            questions: Questions {
                properties: value.questions_properties,
                values: value.questions,
                comments: value.questions_comments.into_boxed_slice(),
                comment_slots: value.questions_comment_slots.into_boxed_slice(),
                comment_qproperties_present,
                source_qproperties_present: comment_qproperties_present,
                source_question_count,
                default_namespace: value.questions_default_namespace,
                prefix: value.questions_prefix,
                namespace_declarations: value.questions_namespaces,
                unknown_attributes: value.questions_attributes,
            },
            extension_xml: value.extension_xml,
            unknown_attributes: value.unknown_attributes,
            namespace_declarations: value.namespace_declarations,
            comments: value.comments.into_boxed_slice(),
            comment_slots: value.comment_slots.into_boxed_slice(),
            source_properties_present,
            source_title_properties_present,
            source_description_properties_present,
            default_namespace: value.default_namespace,
            root_prefix: value.root_prefix,
        })
    }
}

/// Parse a standalone Survey part according to `[MS-XLSX]` §§2.4.69,
/// 2.6.142--2.6.145, and 2.7.27--2.7.29.
///
/// # Errors
///
/// Returns an error when the XML is malformed, outside the Survey namespace,
/// exceeds a parser limit, or violates the Survey schema constraints.
pub fn parse(xml: &[u8]) -> Result<Survey> {
    parse_with_limits(xml, &Limits::default())
}

/// Parse a standalone Survey part with an explicit resource policy.
pub fn parse_with_limits(xml: &[u8], limits: &Limits) -> Result<Survey> {
    if xml.len() > limits.max_part_bytes() {
        return Err(invalid("survey XML exceeds the size limit"));
    }
    crate::source_attributes::validate_xml_characters(xml)?;
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut parser = Parser {
        question_limit: limits.max_questions(),
        ..Parser::default()
    };
    let mut extension_total = 0usize;
    loop {
        let start = reader.buffer_position() as usize;
        let event = reader.read_event().map_err(xml_error)?;
        match event {
            Event::Start(element) => {
                validate_element_names(&reader, &element)?;
                let namespace = reader.resolver().resolve_element(element.name()).0;
                if let Some(target) = parser.extension_target(&namespace, &element)? {
                    let remaining = limits
                        .max_extension_bytes()
                        .saturating_sub(extension_total)
                        .min(MAX_EXTENSION_BYTES);
                    let raw = read_extension_list(
                        &mut reader,
                        xml,
                        start,
                        remaining,
                        MAX_DEPTH - parser.scopes.len(),
                    )?;
                    extension_total += raw.len();
                    parser.attach_extension(target, raw)?;
                } else {
                    parser.start(&namespace, &element)?;
                    if parser
                        .survey
                        .as_ref()
                        .is_some_and(|survey| survey.questions.len() > limits.max_questions())
                    {
                        return Err(invalid("survey has too many questions"));
                    }
                }
            },
            Event::Empty(element) => {
                validate_element_names(&reader, &element)?;
                let namespace = reader.resolver().resolve_element(element.name()).0;
                if let Some(target) = parser.extension_target(&namespace, &element)? {
                    let end = reader.buffer_position() as usize;
                    let length = end - start;
                    if length
                        > limits
                            .max_extension_bytes()
                            .saturating_sub(extension_total)
                            .min(MAX_EXTENSION_BYTES)
                    {
                        return Err(invalid("survey extension list exceeds the size limit"));
                    }
                    let raw = xml[start..end].to_vec();
                    extension_total += raw.len();
                    parser.attach_extension(target, raw)?;
                } else {
                    parser.empty(&namespace, &element)?;
                }
            },
            Event::End(element) => {
                let namespace = reader.resolver().resolve_element(element.name()).0;
                parser.end(&namespace, element.local_name().as_ref())?;
            },
            Event::Text(text) => {
                if !parser.in_extension() && !text.decode().map_err(xml_error)?.trim().is_empty() {
                    return Err(invalid("unexpected text in survey XML"));
                }
            },
            Event::CData(_) => {
                if !parser.in_extension() {
                    return Err(invalid("unexpected CDATA in survey XML"));
                }
            },
            Event::Decl(_) => {},
            Event::Comment(comment) => parser.comment(comment)?,
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid(
                    "DTD and processing instructions are rejected in survey XML",
                ));
            },
            Event::GeneralRef(_) => {
                return Err(invalid("entity references are rejected in survey XML"));
            },
            Event::Eof => break,
        }
    }
    let survey = parser.finish()?;
    validate_survey_with_limits(&survey, limits)?;
    let extension_bytes = survey_extension_bytes(&survey);
    if extension_bytes > limits.max_extension_bytes() {
        return Err(invalid("survey extension list exceeds the size limit"));
    }
    Ok(survey)
}

fn survey_extension_bytes(survey: &Survey) -> usize {
    let mut total = survey.extension_xml.as_ref().map_or(0, |value| value.len());
    let mut add = |value: usize| {
        total = total.checked_add(value).unwrap_or(usize::MAX);
    };
    for properties in [
        survey.properties.as_ref(),
        survey.title_properties.as_ref(),
        survey.description_properties.as_ref(),
        survey.questions.properties.as_ref(),
    ]
    .into_iter()
    .flatten()
    {
        add(properties
            .extension_xml
            .as_ref()
            .map_or(0, |value| value.len()));
    }
    for question in &survey.questions.values {
        add(question
            .extension_xml
            .as_ref()
            .map_or(0, |value| value.len()));
        if let Some(properties) = question.properties.as_ref() {
            add(properties
                .extension_xml
                .as_ref()
                .map_or(0, |value| value.len()));
        }
    }
    total
}

fn read_extension_list(
    reader: &mut NsReader<&[u8]>,
    xml: &[u8],
    start: usize,
    maximum: usize,
    max_depth: usize,
) -> Result<Vec<u8>> {
    let mut depth = 1usize;
    loop {
        if reader.buffer_position() as usize - start > maximum {
            return Err(invalid("survey extension list exceeds the size limit"));
        }
        let event = reader.read_event().map_err(xml_error)?;
        let closes_root = matches!(&event, Event::End(_)) && depth == 1;
        match &event {
            Event::Start(element) => {
                validate_element_names(reader, element)?;
                if depth >= max_depth {
                    return Err(invalid("survey extension-list nesting is too deep"));
                }
                depth += 1;
            },
            Event::Empty(element) => {
                if depth >= max_depth {
                    return Err(invalid("survey extension-list nesting is too deep"));
                }
                validate_element_names(reader, element)?;
            },
            Event::End(_) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("invalid Survey extension-list depth"))?;
            },
            Event::Text(text) => {
                let _value = text
                    .xml_content(XmlVersion::Implicit1_0)
                    .map_err(xml_error)?;
            },
            Event::CData(_) | Event::Comment(_) => {},
            Event::Decl(_) | Event::PI(_) | Event::DocType(_) | Event::GeneralRef(_) => {
                return Err(invalid(
                    "invalid declaration or processing instruction in Survey extLst",
                ));
            },
            Event::Eof => return Err(invalid("unterminated Survey extension list")),
        }
        let end = reader.buffer_position() as usize;
        if end - start > maximum {
            return Err(invalid("survey extension list exceeds the size limit"));
        }
        if closes_root {
            let mut raw = Vec::new();
            raw.try_reserve_exact(end - start)
                .map_err(|_| invalid("survey extension allocation failed"))?;
            raw.extend_from_slice(&xml[start..end]);
            return Ok(raw);
        }
    }
}

fn serialize_event(event: Event<'_>) -> Result<Vec<u8>> {
    let mut writer = quick_xml::Writer::new(Vec::new());
    writer.write_event(event).map_err(xml_error)?;
    Ok(writer.into_inner())
}

/// Load every table-owned survey part without changing the package graph.
///
/// # Errors
///
/// Returns an error when a survey part has an invalid table relationship,
/// duplicate survey ID, or malformed Survey XML.
pub fn load(package: &OpcPackage) -> Result<Vec<Part>> {
    load_with_limits(package, &Limits::default())
}

/// Load every table-owned Survey part with an explicit resource policy.
pub fn load_with_limits(package: &OpcPackage, limits: &Limits) -> Result<Vec<Part>> {
    let mut values = Vec::new();
    let mut ids = HashSet::new();
    let survey_parts = package
        .iter_parts()
        .filter(|part| part.content_type() == CONTENT_TYPE)
        .collect::<Vec<_>>();
    if survey_parts.len() > limits.max_surveys() {
        return Err(invalid("survey count exceeds the size limit"));
    }
    for part in survey_parts {
        // The survey payload is read, so it is decoded here (ADR 0030).
        let part = package.get_part(part.partname())?;
        if part.blob().len() > limits.max_part_bytes() {
            return Err(invalid(format!(
                "survey part '{}' exceeds the size limit",
                part.partname()
            )));
        }
        let mut owners = Vec::new();
        for source in package
            .iter_parts()
            .filter(|source| source.content_type() == ct::SML_TABLE)
        {
            for relationship in source
                .rels()
                .iter()
                .filter(|relationship| relationship.reltype() == RELATIONSHIP_TYPE)
            {
                if relationship.is_external() {
                    continue;
                }
                if relationship
                    .target_partname()
                    .ok()
                    .is_some_and(|target| target.is_equivalent_to(part.partname()))
                {
                    owners.push((
                        source.partname().clone(),
                        relationship.r_id().to_owned(),
                        relationship.target_ref().to_owned(),
                    ));
                }
            }
        }
        if owners.len() != 1 {
            return Err(invalid(format!(
                "survey part '{}' must have exactly one table source relationship",
                part.partname()
            )));
        }
        let (table_part_name, relationship_id, relationship_target) = owners
            .pop()
            .ok_or_else(|| invalid("survey owner relationship is absent"))?;
        let table_part = package.get_part(&table_part_name)?;
        let survey = parse_with_limits(part.blob(), limits)?;
        validate_table_bindings(table_part, &survey)?;
        if !ids.insert(survey.id()) {
            return Err(invalid("survey IDs must be unique within a workbook"));
        }
        let table_relationships_source = package.source_relationships(&table_part_name)?;
        let survey_relationships_source = package.source_relationships(part.partname())?;
        values.push(Part {
            survey: Arc::new(survey),
            table_part_name,
            relationship_id,
            relationship_target,
            part_name: part.partname().clone(),
            source_xml: part.blob_arc(),
            source_proof: Some(Arc::new(package.source_xml_part(part.partname())?)),
            table_source_xml: table_part.blob_arc(),
            table_relationships: Arc::from(relationship_states(table_part)),
            survey_relationships: Arc::from(relationship_states(part)),
            table_relationships_source: Some(table_relationships_source),
            survey_relationships_source: Some(survey_relationships_source),
        });
    }
    values.sort_unstable_by_key(|part| part.survey.id());
    Ok(values)
}

fn validate_table_bindings(table: &dyn OpcPart, survey: &Survey) -> Result<()> {
    let table = crate::table::parse_table_xml(table.blob())?
        .ok_or_else(|| invalid("survey owner table is not a supported table part"))?;
    for question in survey.questions().values() {
        if !table
            .columns
            .iter()
            .any(|column| column.id == question.binding().get())
        {
            return Err(invalid(format!(
                "survey question binding {} does not match an owner table column",
                question.binding().get()
            )));
        }
    }
    Ok(())
}

fn relationship_states(part: &dyn OpcPart) -> Vec<RelationshipState> {
    let mut values = part
        .rels()
        .iter()
        .map(|relationship| RelationshipState {
            source: part.partname().clone(),
            id: relationship.r_id().to_owned(),
            reltype: relationship.reltype().to_owned(),
            target: relationship.target_ref().to_owned(),
            mode: relationship.target_mode(),
        })
        .collect::<Vec<_>>();
    values.sort_by(|left, right| {
        left.source
            .as_str()
            .cmp(right.source.as_str())
            .then_with(|| left.id.cmp(&right.id))
            .then_with(|| left.reltype.cmp(&right.reltype))
            .then_with(|| left.target.cmp(&right.target))
            .then_with(|| {
                (left.mode == TargetMode::External).cmp(&(right.mode == TargetMode::External))
            })
    });
    values
}

fn incoming_relationships(
    package: &OpcPackage,
    entries: &[Part],
) -> Result<Vec<RelationshipState>> {
    let targets = entries
        .iter()
        .map(|entry| entry.part_name.clone())
        .collect::<Vec<_>>();
    let mut values = Vec::new();
    let package_source = PackURI::new("/").map_err(invalid)?;
    for relationship in package.rels().iter() {
        if relationship.is_external() {
            continue;
        }
        let target = relationship
            .target_partname()
            .map_err(|error| invalid(format!("invalid Survey relationship closure: {error}")))?;
        if targets
            .iter()
            .any(|candidate| candidate.is_equivalent_to(&target))
        {
            values.push(RelationshipState {
                source: package_source.clone(),
                id: relationship.r_id().to_owned(),
                reltype: relationship.reltype().to_owned(),
                target: relationship.target_ref().to_owned(),
                mode: relationship.target_mode(),
            });
        }
    }
    for source in package.iter_parts() {
        for relationship in source.rels().iter() {
            if relationship.is_external() {
                continue;
            }
            let target = relationship.target_partname().map_err(|error| {
                invalid(format!("invalid Survey relationship closure: {error}"))
            })?;
            if targets
                .iter()
                .any(|candidate| candidate.is_equivalent_to(&target))
            {
                values.push(RelationshipState {
                    source: source.partname().clone(),
                    id: relationship.r_id().to_owned(),
                    reltype: relationship.reltype().to_owned(),
                    target: relationship.target_ref().to_owned(),
                    mode: relationship.target_mode(),
                });
            }
        }
    }
    values.sort_by(|left, right| {
        left.source
            .as_str()
            .cmp(right.source.as_str())
            .then_with(|| left.id.cmp(&right.id))
            .then_with(|| left.reltype.cmp(&right.reltype))
            .then_with(|| left.target.cmp(&right.target))
            .then_with(|| {
                (left.mode == TargetMode::External).cmp(&(right.mode == TargetMode::External))
            })
    });
    Ok(values)
}

fn set_once<T>(slot: &mut Option<T>, value: T, name: &str) -> Result<()> {
    if slot.replace(value).is_some() {
        Err(invalid(format!("duplicate survey {name}")))
    } else {
        Ok(())
    }
}
fn xml_error(error: impl std::fmt::Display) -> Error {
    invalid(format!("invalid survey XML: {error}"))
}
fn text(value: &str, field: &str) -> Result<Box<str>> {
    if value.len() > MAX_TEXT_BYTES {
        Err(invalid(format!("survey {field} exceeds the size limit")))
    } else {
        Ok(value.into())
    }
}

fn validate_element_names(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<()> {
    if matches!(
        reader.resolver().resolve_element(element.name()).0,
        ResolveResult::Unknown(_)
    ) {
        return Err(invalid("unbound Survey element prefix"));
    }
    let mut expanded = HashSet::new();
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        let value = attribute
            .normalized_value(XmlVersion::Implicit1_0)
            .map_err(xml_error)?;
        crate::source_attributes::validate_xml_characters(value.as_bytes())?;
        if attribute.key.as_namespace_binding().is_some() {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let namespace = match namespace {
            ResolveResult::Bound(namespace) => quick_xml::escape::unescape(
                std::str::from_utf8(namespace.as_ref()).map_err(xml_error)?,
            )
            .map_err(xml_error)?
            .into_owned(),
            ResolveResult::Unbound => String::new(),
            ResolveResult::Unknown(_) => return Err(invalid("unbound Survey attribute prefix")),
        };
        if !expanded.insert((namespace, local.as_ref().to_vec())) {
            return Err(invalid("duplicate expanded Survey attribute"));
        }
    }
    Ok(())
}

fn xstring(value: &str, field: &str) -> Result<Box<str>> {
    if value.len() > MAX_TEXT_BYTES * 7 {
        return Err(invalid("survey encoded string exceeds the size limit"));
    }
    let decoded = crate::raw::strings::decode_spreadsheet_text(value)?;
    validate_xstring(Some(&decoded), field)?;
    Ok(decoded.into_boxed_str())
}

fn attrs(element: &BytesStart<'_>) -> Result<Vec<(Vec<u8>, String)>> {
    element
        .attributes()
        .with_checks(true)
        .map(|raw_attribute| {
            let attribute = raw_attribute.map_err(xml_error)?;
            let decoded_value = attribute
                .normalized_value(XmlVersion::Implicit1_0)
                .map_err(xml_error)?
                .into_owned();
            Ok((attribute.key.as_ref().to_vec(), decoded_value))
        })
        .collect()
}

fn raw_attributes(element: &BytesStart<'_>) -> Result<Vec<RawAttribute>> {
    element
        .attributes()
        .with_checks(true)
        .map(|raw_attribute| {
            let attribute = raw_attribute.map_err(xml_error)?;
            let name = std::str::from_utf8(attribute.key.as_ref())
                .map_err(|_| invalid("survey attribute name is not UTF-8"))?;
            let value = attribute
                .normalized_value(XmlVersion::Implicit1_0)
                .map_err(xml_error)?;
            Ok(RawAttribute {
                name: name.into(),
                value: String::from_utf8(value.into_owned().into_bytes())
                    .map_err(|_| invalid("survey attribute value is not UTF-8"))?
                    .into(),
            })
        })
        .collect()
}

fn unknown_attributes(element: &BytesStart<'_>, known: &[&[u8]]) -> Result<Box<[RawAttribute]>> {
    Ok(raw_attributes(element)?
        .into_iter()
        .filter(|attribute| {
            let name = attribute.name.as_bytes();
            name != b"xmlns" && !name.starts_with(b"xmlns:") && !known.contains(&name)
        })
        .collect::<Vec<_>>()
        .into_boxed_slice())
}

fn namespace_declarations(element: &BytesStart<'_>) -> Result<Box<[RawAttribute]>> {
    Ok(raw_attributes(element)?
        .into_iter()
        .filter(|attribute| {
            let name = attribute.name.as_bytes();
            name.starts_with(b"xmlns:")
        })
        .collect::<Vec<_>>()
        .into_boxed_slice())
}
fn attr<'a>(values: &'a [(Vec<u8>, String)], name: &[u8]) -> Option<&'a str> {
    values
        .iter()
        .find(|(key, _)| key.as_slice() == name)
        .map(|(_, value)| value.as_str())
}

fn element_prefix(element: &BytesStart<'_>) -> Result<Option<Box<str>>> {
    element
        .name()
        .prefix()
        .map(|prefix| {
            std::str::from_utf8(prefix.as_ref())
                .map(str::to_owned)
                .map(String::into_boxed_str)
                .map_err(|_| invalid("survey element prefix is not UTF-8"))
        })
        .transpose()
}

fn required<'a>(values: &'a [(Vec<u8>, String)], name: &[u8], field: &str) -> Result<&'a str> {
    attr(values, name).ok_or_else(|| invalid(format!("survey {field} is required")))
}
fn parse_u32(value: &str, field: &str) -> Result<u32> {
    value
        .parse()
        .map_err(|_parse_error| invalid(format!("invalid survey {field}")))
}
fn parse_i32(value: &str, field: &str) -> Result<i32> {
    value
        .parse()
        .map_err(|_parse_error| invalid(format!("invalid survey {field}")))
}
fn parse_bool(value: &str) -> Result<bool> {
    match value {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        _ => Err(invalid("invalid survey required flag")),
    }
}
fn is_survey_namespace(namespace: &ResolveResult<'_>) -> bool {
    matches!(namespace, ResolveResult::Bound(value) if value.as_ref() == NAMESPACE)
}
fn is_extension_namespace(namespace: &ResolveResult<'_>) -> bool {
    is_survey_namespace(namespace)
}

fn parse_survey(element: &BytesStart<'_>) -> Result<SurveyBuilder> {
    let values = attrs(element)?;
    let root_prefix = element_prefix(element)?;
    Ok(SurveyBuilder {
        id: Some(Id::new(parse_u32(required(&values, b"id", "id")?, "id")?)),
        guid: Some(Guid::new(text(
            required(&values, b"guid", "guid")?,
            "guid",
        )?)?),
        title: attr(&values, b"title")
            .map(|value| xstring(value, "title"))
            .transpose()?,
        description: attr(&values, b"description")
            .map(|value| xstring(value, "description"))
            .transpose()?,
        unknown_attributes: unknown_attributes(
            element,
            &[b"id", b"guid", b"title", b"description"],
        )?,
        namespace_declarations: namespace_declarations(element)?,
        default_namespace: attr(&values, b"xmlns").map(Into::into),
        root_prefix,
        ..SurveyBuilder::default()
    })
}
fn parse_properties(element: &BytesStart<'_>) -> Result<ElementProperties> {
    let values = attrs(element)?;
    Ok(ElementProperties {
        css_class: attr(&values, b"cssClass")
            .map(|value| xstring(value, "cssClass"))
            .transpose()?,
        bottom: attr(&values, b"bottom")
            .map(|value| parse_i32(value, "bottom"))
            .transpose()?,
        top: attr(&values, b"top")
            .map(|value| parse_i32(value, "top"))
            .transpose()?,
        left: attr(&values, b"left")
            .map(|value| parse_i32(value, "left"))
            .transpose()?,
        right: attr(&values, b"right")
            .map(|value| parse_i32(value, "right"))
            .transpose()?,
        width: attr(&values, b"width")
            .map(|value| parse_u32(value, "width"))
            .transpose()?,
        height: attr(&values, b"height")
            .map(|value| parse_u32(value, "height"))
            .transpose()?,
        position: attr(&values, b"position")
            .map(Position::parse)
            .transpose()?,
        unknown_attributes: unknown_attributes(
            element,
            &[
                b"cssClass",
                b"bottom",
                b"top",
                b"left",
                b"right",
                b"width",
                b"height",
                b"position",
            ],
        )?,
        namespace_declarations: namespace_declarations(element)?,
        default_namespace: attr(&values, b"xmlns").map(Into::into),
        prefix: element_prefix(element)?,
        ..ElementProperties::default()
    })
}
fn parse_question(element: &BytesStart<'_>) -> Result<Question> {
    let values = attrs(element)?;
    let row_source = attr(&values, b"rowSource")
        .map(|value| {
            let value = xstring(value, "rowSource")?;
            validate_row_source(&value)?;
            Ok::<_, Error>(value)
        })
        .transpose()?;
    let decimal_places = attr(&values, b"decimalPlaces")
        .map(|lexeme| {
            let places = parse_u32(lexeme, "decimalPlaces")?;
            u8::try_from(places)
                .ok()
                .filter(|number| *number <= 15)
                .ok_or_else(|| invalid("survey decimalPlaces must be at most 15"))
        })
        .transpose()?;
    Ok(Question {
        origin: None,
        binding: Binding::new(parse_u32(
            required(&values, b"binding", "question binding")?,
            "question binding",
        )?),
        column_name: None,
        text: attr(&values, b"text")
            .map(|value| xstring(value, "question text"))
            .transpose()?,
        kind: attr(&values, b"type")
            .map(QuestionType::parse)
            .transpose()?,
        format: attr(&values, b"format")
            .map(QuestionFormat::parse)
            .transpose()?,
        help_text: attr(&values, b"helpText")
            .map(|value| xstring(value, "helpText"))
            .transpose()?,
        required: attr(&values, b"required")
            .map(parse_bool)
            .transpose()?
            .unwrap_or(false),
        default_value: attr(&values, b"defaultValue")
            .map(|value| xstring(value, "defaultValue"))
            .transpose()?,
        decimal_places,
        row_source,
        properties: None,
        extension_xml: None,
        unknown_attributes: unknown_attributes(
            element,
            &[
                b"binding",
                b"text",
                b"type",
                b"format",
                b"helpText",
                b"required",
                b"defaultValue",
                b"decimalPlaces",
                b"rowSource",
            ],
        )?,
        namespace_declarations: namespace_declarations(element)?,
        comments: Box::new([]),
        comment_slots: Box::new([]),
        default_namespace: attr(&values, b"xmlns").map(Into::into),
        prefix: element_prefix(element)?,
    })
}
fn validate_row_source(value: &str) -> Result<()> {
    // MS-XLSX 2.6.144 permits quotes in *value-char-with-quote, and
    // terminated-value permits empty values separated by semicolons. Thus
    // every XML Char string has a valid unquoted segmentation. CSV-style
    // quote balancing would reject legal strings; retain the opaque value.
    validate_text(Some(value), "rowSource")
}

/// Serialize one Survey part using the protocol namespace and retained
/// extension XML.
pub fn write(survey: &Survey) -> Result<Vec<u8>> {
    write_with_limits(survey, &Limits::default())
}

/// Validate one Survey's typed semantic state without serializing it.
pub fn validate(survey: &Survey) -> Result<()> {
    validate_survey(survey)
}

/// Validate one Survey with an explicit resource policy.
pub fn validate_with_limits(survey: &Survey, limits: &Limits) -> Result<()> {
    validate_survey_with_limits(survey, limits)
}

#[derive(Clone, Copy)]
struct NamespaceContext<'a> {
    prefix: Option<&'a str>,
}

fn namespace_context(survey: &Survey) -> NamespaceContext<'_> {
    if let Some(prefix) = survey.root_prefix.as_deref() {
        NamespaceContext {
            prefix: Some(prefix),
        }
    } else if survey.default_namespace.as_deref() == Some(NAMESPACE_STR) {
        NamespaceContext { prefix: None }
    } else {
        NamespaceContext {
            prefix: Some("survey"),
        }
    }
}

fn owner_context(prefix: Option<&str>) -> NamespaceContext<'_> {
    NamespaceContext { prefix }
}

fn write_name(output: &mut Vec<u8>, context: NamespaceContext<'_>, name: &str) {
    if let Some(prefix) = context.prefix {
        output.extend_from_slice(prefix.as_bytes());
        output.push(b':');
    }
    output.extend_from_slice(name.as_bytes());
}

fn write_comment_slots(output: &mut Vec<u8>, slots: &[CommentSlot], slot: usize) {
    for comment in slots.iter().filter(|comment| comment.slot == slot) {
        output.extend_from_slice(&comment.value);
    }
}

fn has_namespace_declaration(attributes: &[RawAttribute], prefix: &str) -> bool {
    let expected = format!("xmlns:{prefix}");
    attributes
        .iter()
        .any(|attribute| attribute.name.as_ref() == expected)
}

/// Serialize one Survey part with an explicit resource policy.
pub fn write_with_limits(survey: &Survey, limits: &Limits) -> Result<Vec<u8>> {
    validate_survey_with_limits(survey, limits)?;
    ensure_resolved_questions(survey)?;
    let output_bound = serialized_size_bound(survey)?;
    if output_bound > limits.max_part_bytes() {
        return Err(invalid("survey output exceeds the size limit"));
    }
    let mut output = Vec::with_capacity(output_bound);
    let context = namespace_context(survey);
    output.extend_from_slice(b"<");
    write_name(&mut output, context, "survey");
    if let Some(default_namespace) = survey.default_namespace.as_deref() {
        write_attribute_value(&mut output, "xmlns", default_namespace);
    } else if context.prefix.is_none() {
        write_attribute_value(&mut output, "xmlns", NAMESPACE_STR);
    }
    if let Some(prefix) = context.prefix {
        if !has_namespace_declaration(&survey.namespace_declarations, prefix) {
            let name = format!("xmlns:{prefix}");
            write_attribute_value(
                &mut output,
                &name,
                std::str::from_utf8(NAMESPACE).unwrap_or_default(),
            );
        }
    }
    for attribute in &survey.namespace_declarations {
        write_attribute(&mut output, attribute)?;
    }
    for attribute in &survey.unknown_attributes {
        write_attribute(&mut output, attribute)?;
    }
    write_attribute_value(&mut output, "id", &survey.id.get().to_string());
    write_attribute_value(&mut output, "guid", survey.guid.as_str());
    if let Some(value) = survey.title.as_deref() {
        write_xstring_attribute(&mut output, "title", value);
    }
    if let Some(value) = survey.description.as_deref() {
        write_xstring_attribute(&mut output, "description", value);
    }
    output.extend_from_slice(b">");
    let mut child_slot = 0usize;
    write_comment_slots(&mut output, &survey.comment_slots, child_slot);
    if let Some(properties) = survey.properties.as_ref() {
        write_properties_with_context(
            &mut output,
            "surveyPr",
            properties,
            owner_context(properties.prefix.as_deref()),
        )?;
    }
    child_slot = 1;
    write_comment_slots(&mut output, &survey.comment_slots, child_slot);
    if let Some(properties) = survey.title_properties.as_ref() {
        write_properties_with_context(
            &mut output,
            "titlePr",
            properties,
            owner_context(properties.prefix.as_deref()),
        )?;
    }
    child_slot = 2;
    write_comment_slots(&mut output, &survey.comment_slots, child_slot);
    if let Some(properties) = survey.description_properties.as_ref() {
        write_properties_with_context(
            &mut output,
            "descriptionPr",
            properties,
            owner_context(properties.prefix.as_deref()),
        )?;
    }
    child_slot = 3;
    write_comment_slots(&mut output, &survey.comment_slots, child_slot);
    let questions_context = owner_context(survey.questions.prefix.as_deref());
    output.extend_from_slice(b"<");
    write_name(&mut output, questions_context, "questions");
    if let Some(default_namespace) = survey.questions.default_namespace.as_deref() {
        write_attribute_value(&mut output, "xmlns", default_namespace);
    } else if questions_context.prefix.is_none() {
        write_attribute_value(&mut output, "xmlns", NAMESPACE_STR);
    }
    for attribute in survey
        .questions
        .namespace_declarations
        .iter()
        .chain(survey.questions.unknown_attributes.iter())
    {
        write_attribute(&mut output, attribute)?;
    }
    output.push(b'>');
    let current_qproperties_present = survey.questions.comment_qproperties_present;
    let source_qproperties_present = survey.questions.source_qproperties_present;
    let source_question_count = survey.questions.source_question_count;
    if source_qproperties_present || current_qproperties_present {
        write_comment_slots(&mut output, &survey.questions.comment_slots, 0);
    }
    if let Some(properties) = survey.questions.properties.as_ref() {
        write_properties_with_context(
            &mut output,
            "questionsPr",
            properties,
            owner_context(properties.prefix.as_deref()),
        )?;
    }
    let source_slot_offset = usize::from(source_qproperties_present);
    for index in 0..source_question_count {
        let source_slot = index + source_slot_offset;
        if !(current_qproperties_present && !source_qproperties_present && index == 0) {
            write_comment_slots(&mut output, &survey.questions.comment_slots, source_slot);
        }
        if let Some(question) = survey.questions.values.get(index) {
            write_question_with_context(
                &mut output,
                question,
                owner_context(question.prefix.as_deref()),
            )?;
        }
    }
    for question in survey.questions.values.iter().skip(source_question_count) {
        write_question_with_context(
            &mut output,
            question,
            owner_context(question.prefix.as_deref()),
        )?;
    }
    write_comment_slots(
        &mut output,
        &survey.questions.comment_slots,
        source_question_count + source_slot_offset,
    );
    output.extend_from_slice(b"</");
    write_name(&mut output, questions_context, "questions");
    output.push(b'>');
    child_slot += 1;
    write_comment_slots(&mut output, &survey.comment_slots, child_slot);
    if let Some(extension) = survey.extension_xml.as_deref() {
        output.extend_from_slice(extension);
        child_slot += 1;
        write_comment_slots(&mut output, &survey.comment_slots, child_slot);
    }
    output.extend_from_slice(b"</");
    write_name(&mut output, context, "survey");
    output.push(b'>');
    if output.len() > limits.max_part_bytes() {
        return Err(invalid("survey output exceeds the size limit"));
    }
    Ok(output)
}

fn serialized_size_bound(survey: &Survey) -> Result<usize> {
    let mut size = 0usize;
    // Fixed markup and attribute punctuation.  The per-element allowance is
    // deliberately conservative and is independent of caller-provided text,
    // so a large authored collection is rejected before Vec allocation.
    add_size(
        &mut size,
        1024usize.saturating_add(survey.questions.values.len().saturating_mul(512)),
    )?;
    add_text_size(
        &mut size,
        std::str::from_utf8(NAMESPACE).unwrap_or_default(),
    )?;
    if let Some(default_namespace) = survey.default_namespace.as_deref() {
        add_attribute_bound(
            &mut size,
            &RawAttribute {
                name: "xmlns".into(),
                value: default_namespace.into(),
            },
        )?;
    } else if survey.root_prefix.is_none() {
        add_attribute_bound(
            &mut size,
            &RawAttribute {
                name: "xmlns".into(),
                value: NAMESPACE_STR.into(),
            },
        )?;
    }
    if let Some(prefix) = survey.root_prefix.as_deref() {
        if !has_namespace_declaration(&survey.namespace_declarations, prefix) {
            add_attribute_bound(
                &mut size,
                &RawAttribute {
                    name: format!("xmlns:{prefix}").into_boxed_str(),
                    value: NAMESPACE_STR.into(),
                },
            )?;
        }
    }
    for attribute in &survey.namespace_declarations {
        add_attribute_bound(&mut size, attribute)?;
    }
    for attribute in &survey.unknown_attributes {
        add_attribute_bound(&mut size, attribute)?;
    }
    for attribute in survey
        .questions
        .namespace_declarations
        .iter()
        .chain(survey.questions.unknown_attributes.iter())
    {
        add_attribute_bound(&mut size, attribute)?;
    }
    if let Some(default_namespace) = survey.questions.default_namespace.as_deref() {
        add_attribute_bound(
            &mut size,
            &RawAttribute {
                name: "xmlns".into(),
                value: default_namespace.into(),
            },
        )?;
    } else if survey.questions.prefix.is_none() {
        add_attribute_bound(
            &mut size,
            &RawAttribute {
                name: "xmlns".into(),
                value: NAMESPACE_STR.into(),
            },
        )?;
    }
    if let Some(prefix) = survey.questions.prefix.as_deref() {
        add_text_size(&mut size, prefix)?;
    }
    add_text_size(&mut size, survey.guid.as_str())?;
    if let Some(value) = survey.title.as_deref() {
        add_xstring_size(&mut size, value)?;
    }
    if let Some(value) = survey.description.as_deref() {
        add_xstring_size(&mut size, value)?;
    }
    for comment in &survey.comments {
        add_size(&mut size, comment.len())?;
    }
    for comment in &survey.questions.comments {
        add_size(&mut size, comment.len())?;
    }
    for properties in [
        survey.properties.as_ref(),
        survey.title_properties.as_ref(),
        survey.description_properties.as_ref(),
        survey.questions.properties.as_ref(),
    ]
    .into_iter()
    .flatten()
    {
        add_properties_bound(&mut size, properties)?;
    }
    for question in &survey.questions.values {
        add_xstring_size(&mut size, question.text.as_deref().unwrap_or_default())?;
        add_xstring_size(&mut size, question.help_text.as_deref().unwrap_or_default())?;
        add_xstring_size(
            &mut size,
            question.default_value.as_deref().unwrap_or_default(),
        )?;
        add_xstring_size(
            &mut size,
            question.row_source.as_deref().unwrap_or_default(),
        )?;
        if let Some(default_namespace) = question.default_namespace.as_deref() {
            add_attribute_bound(
                &mut size,
                &RawAttribute {
                    name: "xmlns".into(),
                    value: default_namespace.into(),
                },
            )?;
        } else if question.prefix.is_none() {
            add_attribute_bound(
                &mut size,
                &RawAttribute {
                    name: "xmlns".into(),
                    value: NAMESPACE_STR.into(),
                },
            )?;
        }
        if let Some(prefix) = question.prefix.as_deref() {
            add_text_size(&mut size, prefix)?;
        }
        for attribute in &question.namespace_declarations {
            add_attribute_bound(&mut size, attribute)?;
        }
        for attribute in &question.unknown_attributes {
            add_attribute_bound(&mut size, attribute)?;
        }
        if let Some(properties) = question.properties.as_ref() {
            add_properties_bound(&mut size, properties)?;
        }
        for comment in &question.comments {
            add_size(&mut size, comment.len())?;
        }
    }
    add_size(&mut size, survey_extension_bytes(survey))?;
    Ok(size)
}

fn add_size(size: &mut usize, value: usize) -> Result<()> {
    *size = size
        .checked_add(value)
        .ok_or_else(|| invalid("survey output size overflow"))?;
    Ok(())
}

fn add_text_size(size: &mut usize, value: &str) -> Result<()> {
    add_size(size, escaped_len(value)?)
}

fn add_xstring_size(size: &mut usize, value: &str) -> Result<()> {
    for (at, character) in value.char_indices() {
        let literal_escape = character == '_'
            && value
                .as_bytes()
                .get(at..at.saturating_add(7))
                .is_some_and(|bytes| {
                    bytes[1] == b'x'
                        && bytes[6] == b'_'
                        && bytes[2..6].iter().all(u8::is_ascii_hexdigit)
                });
        let width = if literal_escape {
            7
        } else if character == '>' {
            1
        } else {
            escaped_character_len(character)
        };
        add_size(size, width)?;
    }
    Ok(())
}

fn add_attribute_bound(size: &mut usize, attribute: &RawAttribute) -> Result<()> {
    add_size(size, 4)?;
    add_text_size(size, &attribute.name)?;
    add_text_size(size, &attribute.value)
}

fn add_properties_bound(size: &mut usize, properties: &ElementProperties) -> Result<()> {
    add_size(size, 512)?;
    if let Some(default_namespace) = properties.default_namespace.as_deref() {
        add_attribute_bound(
            size,
            &RawAttribute {
                name: "xmlns".into(),
                value: default_namespace.into(),
            },
        )?;
    } else if properties.prefix.is_none() {
        add_attribute_bound(
            size,
            &RawAttribute {
                name: "xmlns".into(),
                value: NAMESPACE_STR.into(),
            },
        )?;
    }
    if let Some(prefix) = properties.prefix.as_deref() {
        add_text_size(size, prefix)?;
    }
    if let Some(value) = properties.css_class.as_deref() {
        add_xstring_size(size, value)?;
    }
    for comment in &properties.comments {
        add_size(size, comment.len())?;
    }
    for attribute in &properties.namespace_declarations {
        add_attribute_bound(size, attribute)?;
    }
    for attribute in &properties.unknown_attributes {
        add_attribute_bound(size, attribute)?;
    }
    add_size(
        size,
        properties
            .extension_xml
            .as_ref()
            .map_or(0, |value| value.len()),
    )
}

fn escaped_character_len(character: char) -> usize {
    match character {
        '&' => 5,
        '<' | '>' => 4,
        '"' | '\'' => 6,
        '\t' | '\n' | '\r' => 5,
        '\u{0}'..='\u{8}' | '\u{b}' | '\u{c}' | '\u{e}'..='\u{1f}' | '\u{fffe}' | '\u{ffff}' => 7,
        character => character.len_utf8(),
    }
}

fn escaped_len(value: &str) -> Result<usize> {
    value.chars().try_fold(0usize, |total, character| {
        total
            .checked_add(escaped_character_len(character))
            .ok_or_else(|| invalid("survey output size overflow"))
    })
}

fn validate_survey(survey: &Survey) -> Result<()> {
    validate_survey_with_limits(survey, &Limits::default())
}

fn validate_survey_with_limits(survey: &Survey, limits: &Limits) -> Result<()> {
    if survey.questions.values.is_empty() {
        return Err(invalid("survey requires at least one question"));
    }
    if survey.questions.values.len() > limits.max_questions() {
        return Err(invalid("survey has too many questions"));
    }
    if !valid_guid(survey.guid.as_str()) {
        return Err(invalid("invalid survey GUID"));
    }
    validate_xstring(survey.title.as_deref(), "title")?;
    validate_xstring(survey.description.as_deref(), "description")?;
    validate_text(survey.default_namespace.as_deref(), "default namespace")?;
    validate_text(survey.root_prefix.as_deref(), "root prefix")?;
    validate_properties(survey.properties.as_ref())?;
    validate_properties(survey.title_properties.as_ref())?;
    validate_properties(survey.description_properties.as_ref())?;
    validate_properties(survey.questions.properties.as_ref())?;
    for attribute in &survey.unknown_attributes {
        validate_raw_attribute(attribute)?;
    }
    for attribute in &survey.namespace_declarations {
        validate_raw_attribute(attribute)?;
    }
    for attribute in survey
        .questions
        .namespace_declarations
        .iter()
        .chain(survey.questions.unknown_attributes.iter())
    {
        validate_raw_attribute(attribute)?;
    }
    validate_text(
        survey.questions.default_namespace.as_deref(),
        "questions default namespace",
    )?;
    validate_text(survey.questions.prefix.as_deref(), "questions prefix")?;
    validate_extension(survey.extension_xml.as_deref(), limits)?;
    for question in &survey.questions.values {
        validate_question(question, limits)?;
    }
    validate_total_text_bytes(survey, limits)?;
    if survey_extension_bytes(survey) > limits.max_extension_bytes() {
        return Err(invalid("survey extension list exceeds the size limit"));
    }
    Ok(())
}

fn validate_total_text_bytes(survey: &Survey, limits: &Limits) -> Result<()> {
    let maximum = limits.max_part_bytes().min(MAX_TOTAL_TEXT_BYTES);
    let mut total = 0usize;
    let mut add = |value: Option<&str>| -> Result<()> {
        let Some(value) = value else {
            return Ok(());
        };
        total = total
            .checked_add(value.len())
            .ok_or_else(|| invalid("survey text size overflow"))?;
        if total > maximum {
            return Err(invalid("survey text exceeds the aggregate size limit"));
        }
        Ok(())
    };
    for attribute in survey
        .questions
        .namespace_declarations
        .iter()
        .chain(survey.questions.unknown_attributes.iter())
    {
        add(Some(&attribute.name))?;
        add(Some(&attribute.value))?;
    }
    add(Some(survey.guid.as_str()))?;
    add(survey.default_namespace.as_deref())?;
    add(survey.root_prefix.as_deref())?;
    add(survey.title.as_deref())?;
    add(survey.description.as_deref())?;
    for attribute in &survey.unknown_attributes {
        add(Some(&attribute.name))?;
        add(Some(&attribute.value))?;
    }
    for attribute in &survey.namespace_declarations {
        add(Some(&attribute.name))?;
        add(Some(&attribute.value))?;
    }
    for properties in [
        survey.properties.as_ref(),
        survey.title_properties.as_ref(),
        survey.description_properties.as_ref(),
        survey.questions.properties.as_ref(),
    ]
    .into_iter()
    .flatten()
    {
        add(properties.css_class.as_deref())?;
        add(properties.default_namespace.as_deref())?;
        add(properties.prefix.as_deref())?;
        for attribute in &properties.namespace_declarations {
            add(Some(&attribute.name))?;
            add(Some(&attribute.value))?;
        }
        for attribute in &properties.unknown_attributes {
            add(Some(&attribute.name))?;
            add(Some(&attribute.value))?;
        }
    }
    for question in &survey.questions.values {
        add(question.column_name.as_deref())?;
        add(question.text.as_deref())?;
        add(question.help_text.as_deref())?;
        add(question.default_value.as_deref())?;
        add(question.row_source.as_deref())?;
        add(question.default_namespace.as_deref())?;
        add(question.prefix.as_deref())?;
        for attribute in &question.namespace_declarations {
            add(Some(&attribute.name))?;
            add(Some(&attribute.value))?;
        }
        for attribute in &question.unknown_attributes {
            add(Some(&attribute.name))?;
            add(Some(&attribute.value))?;
        }
        if let Some(properties) = question.properties.as_ref() {
            add(properties.css_class.as_deref())?;
            add(properties.default_namespace.as_deref())?;
            add(properties.prefix.as_deref())?;
            for attribute in &properties.namespace_declarations {
                add(Some(&attribute.name))?;
                add(Some(&attribute.value))?;
            }
            for attribute in &properties.unknown_attributes {
                add(Some(&attribute.name))?;
                add(Some(&attribute.value))?;
            }
        }
    }
    Ok(())
}

fn validate_question(question: &Question, limits: &Limits) -> Result<()> {
    validate_text(
        question.default_namespace.as_deref(),
        "question default namespace",
    )?;
    validate_text(question.prefix.as_deref(), "question prefix")?;
    if let Some(column_name) = question.column_name.as_deref() {
        validate_text(Some(column_name), "column name")?;
    }
    validate_xstring(question.text.as_deref(), "question text")?;
    validate_xstring(question.help_text.as_deref(), "helpText")?;
    validate_xstring(question.default_value.as_deref(), "defaultValue")?;
    if let Some(value) = question.row_source.as_deref() {
        validate_row_source(value)?;
        validate_text(Some(value), "rowSource")?;
    }
    if question.decimal_places.is_some_and(|value| value > 15) {
        return Err(invalid("survey decimalPlaces must be at most 15"));
    }
    validate_properties(question.properties.as_ref())?;
    for attribute in &question.unknown_attributes {
        validate_raw_attribute(attribute)?;
    }
    for attribute in &question.namespace_declarations {
        validate_raw_attribute(attribute)?;
    }
    validate_extension(question.extension_xml.as_deref(), limits)
}

fn ensure_resolved_questions(survey: &Survey) -> Result<()> {
    if survey
        .questions
        .values
        .iter()
        .any(|question| question.column_name.is_some())
    {
        return Err(invalid(
            "survey question column selector must be resolved against a table",
        ));
    }
    Ok(())
}

fn validate_properties(properties: Option<&ElementProperties>) -> Result<()> {
    let Some(properties) = properties else {
        return Ok(());
    };
    validate_text(
        properties.default_namespace.as_deref(),
        "property default namespace",
    )?;
    validate_text(properties.prefix.as_deref(), "property prefix")?;
    validate_xstring(properties.css_class.as_deref(), "cssClass")?;
    for attribute in &properties.namespace_declarations {
        validate_raw_attribute(attribute)?;
    }
    for attribute in &properties.unknown_attributes {
        validate_raw_attribute(attribute)?;
    }
    Ok(())
}

fn validate_extension(extension: Option<&[u8]>, limits: &Limits) -> Result<()> {
    if let Some(extension) = extension {
        if extension.len() > limits.max_extension_bytes() {
            return Err(invalid("survey extension list exceeds the size limit"));
        }
        std::str::from_utf8(extension)
            .map_err(|_| invalid("survey extension list is not UTF-8"))?;
    }
    Ok(())
}

fn validate_xstring(value: Option<&str>, field: &str) -> Result<()> {
    if value.is_some_and(|value| value.len() > MAX_TEXT_BYTES) {
        return Err(invalid(format!("survey {field} exceeds the size limit")));
    }
    Ok(())
}

fn validate_text(value: Option<&str>, field: &str) -> Result<()> {
    validate_xstring(value, field)?;
    if let Some(value) = value {
        crate::source_attributes::validate_xml_characters(value.as_bytes())?;
    }
    Ok(())
}

fn validate_raw_attribute(attribute: &RawAttribute) -> Result<()> {
    if attribute.name.is_empty()
        || attribute
            .name
            .chars()
            .any(|character| character.is_whitespace())
    {
        return Err(invalid("survey unknown attribute name is invalid"));
    }
    validate_text(Some(&attribute.value), "unknown attribute")
}

fn write_question_with_namespace(
    output: &mut Vec<u8>,
    question: &Question,
    explicit_namespace: bool,
) -> Result<()> {
    if explicit_namespace {
        output.extend_from_slice(b"<question");
        output.extend_from_slice(b" xmlns=\"");
        output.extend_from_slice(NAMESPACE);
        output.push(b'"');
        return write_question_body(output, question, NamespaceContext { prefix: None });
    }
    write_question_with_context(output, question, NamespaceContext { prefix: None })
}

fn write_question_with_context(
    output: &mut Vec<u8>,
    question: &Question,
    context: NamespaceContext<'_>,
) -> Result<()> {
    output.extend_from_slice(b"<");
    write_name(output, context, "question");
    if let Some(default_namespace) = question.default_namespace.as_deref() {
        write_attribute_value(output, "xmlns", default_namespace);
    } else if context.prefix.is_none() {
        write_attribute_value(output, "xmlns", NAMESPACE_STR);
    }
    write_question_body(output, question, context)
}

fn write_question_body(
    output: &mut Vec<u8>,
    question: &Question,
    context: NamespaceContext<'_>,
) -> Result<()> {
    for attribute in &question.namespace_declarations {
        write_attribute(output, attribute)?;
    }
    write_attribute_value(output, "binding", &question.binding.get().to_string());
    write_optional_xstring(output, "text", question.text.as_deref());
    write_optional_attribute(output, "type", question.kind.map(QuestionType::as_str));
    write_optional_attribute(
        output,
        "format",
        question.format.map(QuestionFormat::as_str),
    );
    write_optional_xstring(output, "helpText", question.help_text.as_deref());
    if question.required {
        write_attribute_value(output, "required", "true");
    }
    write_optional_xstring(output, "defaultValue", question.default_value.as_deref());
    if let Some(value) = question.decimal_places {
        write_attribute_value(output, "decimalPlaces", &value.to_string());
    }
    write_optional_xstring(output, "rowSource", question.row_source.as_deref());
    for attribute in &question.unknown_attributes {
        write_attribute(output, attribute)?;
    }
    if question.properties.is_none()
        && question.extension_xml.is_none()
        && question.comments.is_empty()
        && question.comment_slots.is_empty()
    {
        output.extend_from_slice(b"/>");
        return Ok(());
    }
    output.extend_from_slice(b">");
    let mut child_slot = 0usize;
    write_comment_slots(output, &question.comment_slots, child_slot);
    if let Some(properties) = question.properties.as_ref() {
        write_properties_with_context(
            output,
            "questionPr",
            properties,
            owner_context(properties.prefix.as_deref()),
        )?;
    }
    child_slot = 1;
    write_comment_slots(output, &question.comment_slots, child_slot);
    if let Some(extension) = question.extension_xml.as_deref() {
        output.extend_from_slice(extension);
    }
    child_slot = 2;
    write_comment_slots(output, &question.comment_slots, child_slot);
    output.extend_from_slice(b"</");
    write_name(output, context, "question");
    output.push(b'>');
    Ok(())
}

fn write_properties_with_namespace(
    output: &mut Vec<u8>,
    name: &str,
    properties: &ElementProperties,
    explicit_namespace: bool,
) -> Result<()> {
    output.extend_from_slice(b"<");
    let context = owner_context(properties.prefix.as_deref());
    write_name(output, context, name);
    if explicit_namespace && properties.prefix.is_none() {
        output.extend_from_slice(b" xmlns=\"");
        output.extend_from_slice(NAMESPACE);
        output.push(b'"');
    } else if let Some(prefix) = properties.prefix.as_deref()
        && !has_namespace_declaration(&properties.namespace_declarations, prefix)
    {
        let name = format!("xmlns:{prefix}");
        write_attribute_value(output, &name, NAMESPACE_STR);
    }
    write_properties_body(output, name, properties, context)
}

fn write_properties_with_context(
    output: &mut Vec<u8>,
    name: &str,
    properties: &ElementProperties,
    context: NamespaceContext<'_>,
) -> Result<()> {
    output.extend_from_slice(b"<");
    write_name(output, context, name);
    if let Some(default_namespace) = properties.default_namespace.as_deref() {
        write_attribute_value(output, "xmlns", default_namespace);
    } else if context.prefix.is_none() {
        write_attribute_value(output, "xmlns", NAMESPACE_STR);
    }
    write_properties_body(output, name, properties, context)
}

fn write_properties_body(
    output: &mut Vec<u8>,
    name: &str,
    properties: &ElementProperties,
    context: NamespaceContext<'_>,
) -> Result<()> {
    for attribute in &properties.namespace_declarations {
        write_attribute(output, attribute)?;
    }
    if let Some(value) = properties.css_class.as_deref() {
        write_xstring_attribute(output, "cssClass", value);
    }
    if let Some(value) = properties.bottom {
        write_attribute_value(output, "bottom", &value.to_string());
    }
    if let Some(value) = properties.top {
        write_attribute_value(output, "top", &value.to_string());
    }
    if let Some(value) = properties.left {
        write_attribute_value(output, "left", &value.to_string());
    }
    if let Some(value) = properties.right {
        write_attribute_value(output, "right", &value.to_string());
    }
    if let Some(value) = properties.width {
        write_attribute_value(output, "width", &value.to_string());
    }
    if let Some(value) = properties.height {
        write_attribute_value(output, "height", &value.to_string());
    }
    if let Some(value) = properties.position {
        write_attribute_value(output, "position", value.as_str());
    }
    for attribute in &properties.unknown_attributes {
        write_attribute(output, attribute)?;
    }
    if properties.extension_xml.is_none() {
        if properties.comments.is_empty() {
            output.extend_from_slice(b"/>");
        } else {
            output.extend_from_slice(b">");
            write_comment_slots(output, &properties.comment_slots, 0);
            output.extend_from_slice(b"</");
            write_name(output, context, name);
            output.extend_from_slice(b">");
        }
    } else {
        output.extend_from_slice(b">");
        write_comment_slots(output, &properties.comment_slots, 0);
        output.extend_from_slice(properties.extension_xml.as_deref().unwrap_or_default());
        write_comment_slots(output, &properties.comment_slots, 1);
        output.extend_from_slice(b"</");
        write_name(output, context, name);
        output.extend_from_slice(b">");
    }
    Ok(())
}

fn write_optional_xstring(output: &mut Vec<u8>, name: &str, value: Option<&str>) {
    if let Some(value) = value {
        write_xstring_attribute(output, name, value);
    }
}

fn write_xstring_attribute(output: &mut Vec<u8>, name: &str, value: &str) {
    output.push(b' ');
    output.extend_from_slice(name.as_bytes());
    output.extend_from_slice(b"=\"");
    crate::source_attributes::append_escaped_xstring(output, value);
    output.push(b'"');
}

fn write_optional_attribute(output: &mut Vec<u8>, name: &str, value: Option<&str>) {
    if let Some(value) = value {
        write_attribute_value(output, name, value);
    }
}

fn write_attribute(output: &mut Vec<u8>, attribute: &RawAttribute) -> Result<()> {
    validate_raw_attribute(attribute)?;
    write_attribute_value(output, &attribute.name, &attribute.value);
    Ok(())
}

fn write_attribute_value(output: &mut Vec<u8>, name: &str, value: &str) {
    output.extend_from_slice(b" ");
    output.extend_from_slice(name.as_bytes());
    output.extend_from_slice(b"=\"");
    escape_attribute(output, value);
    output.extend_from_slice(b"\"");
}

fn escape_attribute(output: &mut Vec<u8>, value: &str) {
    for character in value.chars() {
        match character {
            '&' => output.extend_from_slice(b"&amp;"),
            '<' => output.extend_from_slice(b"&lt;"),
            '>' => output.extend_from_slice(b"&gt;"),
            '"' => output.extend_from_slice(b"&quot;"),
            '\'' => output.extend_from_slice(b"&apos;"),
            '\t' => output.extend_from_slice(b"&#x9;"),
            '\n' => output.extend_from_slice(b"&#xA;"),
            '\r' => output.extend_from_slice(b"&#xD;"),
            character => {
                let mut buffer = [0u8; 4];
                output.extend_from_slice(character.encode_utf8(&mut buffer).as_bytes());
            },
        }
    }
}

/// Immutable source-bound Survey catalog.
#[derive(Debug, Clone)]
pub struct Snapshot {
    entries: Arc<[Part]>,
    source_content_types: OwnedContentTypes,
    incoming_relationships: Arc<[RelationshipState]>,
    limits: Limits,
}

impl Snapshot {
    /// Load the Survey catalog with the default limits.
    pub fn load(package: &OpcPackage) -> Result<Self> {
        Self::load_with_limits(package, &Limits::default())
    }

    /// Load the Survey catalog with an explicit resource policy.
    pub fn load_with_limits(package: &OpcPackage, limits: &Limits) -> Result<Self> {
        let entries = Arc::from(load_with_limits(package, limits)?.into_boxed_slice());
        let incoming_relationships =
            Arc::from(incoming_relationships(package, &entries)?.into_boxed_slice());
        Ok(Self {
            entries,
            source_content_types: package.source_content_types()?,
            incoming_relationships,
            limits: *limits,
        })
    }

    /// Alias emphasizing that this result is source-bound.
    pub fn read(package: &OpcPackage) -> Result<Self> {
        Self::load(package)
    }

    /// Survey entries in deterministic UID order.
    #[must_use]
    pub fn entries(&self) -> &[Part] {
        &self.entries
    }

    /// Contextual alias for [`Self::entries`].
    #[must_use]
    pub fn surveys(&self) -> &[Part] {
        self.entries()
    }

    /// Resource policy retained by this snapshot.
    #[must_use]
    pub const fn limits(&self) -> Limits {
        self.limits
    }

    /// Whether no Survey parts are attached to tables.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.entries.is_empty()
    }

    /// Return the number of table-owned Survey parts.
    #[must_use]
    pub fn len(&self) -> usize {
        self.entries.len()
    }

    fn same_source(&self, other: &Self) -> bool {
        self.source_content_types == other.source_content_types
            && self.entries.len() == other.entries.len()
            && self.incoming_relationships == other.incoming_relationships
            && self
                .entries
                .iter()
                .zip(other.entries.iter())
                .all(|(left, right)| {
                    same_identity(left, right)
                        && left.source_xml == right.source_xml
                        && left.table_source_xml == right.table_source_xml
                        && left.table_relationships == right.table_relationships
                        && left.survey_relationships == right.survey_relationships
                        && left.table_relationships_source == right.table_relationships_source
                        && left.survey_relationships_source == right.survey_relationships_source
                })
    }

    fn same_semantics(&self, entries: &[Part]) -> bool {
        self.entries.len() == entries.len()
            && self.entries.iter().all(|left| {
                entries
                    .iter()
                    .find(|right| same_identity(left, right))
                    .is_some_and(|right| left.survey == right.survey)
            })
    }
}

/// A clone-staged transaction over table-owned Survey parts.
pub struct Transaction<'a> {
    target: &'a mut OpcPackage,
    before: Snapshot,
    draft: Vec<Part>,
    limits: Limits,
}

impl<'a> Transaction<'a> {
    /// Start a transaction using the default Survey limits.
    pub fn new(target: &'a mut OpcPackage) -> Result<Self> {
        Self::with_limits(target, &Limits::default())
    }

    /// Start a transaction using an explicit Survey resource policy.
    pub fn with_limits(target: &'a mut OpcPackage, limits: &Limits) -> Result<Self> {
        let before = Snapshot::load_with_limits(target, limits)?;
        Ok(Self {
            target,
            draft: before.entries.to_vec(),
            before,
            limits: *limits,
        })
    }

    /// Immutable source captured at transaction start.
    #[must_use]
    pub fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Currently staged Survey entries.
    #[must_use]
    pub fn entries(&self) -> &[Part] {
        &self.draft
    }

    /// Edit one existing Survey in a clone-staged callback.
    pub fn edit(
        &mut self,
        index: usize,
        edit: impl FnOnce(&mut Survey) -> Result<()>,
    ) -> Result<bool> {
        let mut draft = self
            .draft
            .get(index)
            .cloned()
            .ok_or_else(|| invalid(format!("survey index {index} is absent")))?;
        edit(Arc::make_mut(&mut draft.survey))?;
        if draft
            .survey
            .questions
            .values
            .iter()
            .any(|question| question.column_name.is_some())
        {
            let table = crate::table::parse_table_xml(
                self.target.get_part(&draft.table_part_name)?.blob(),
            )?
            .ok_or_else(|| invalid("Survey owner table is not a supported table part"))?;
            for question in Arc::make_mut(&mut draft.survey)
                .questions_mut()
                .values_mut()
            {
                question.resolve_column(&table)?;
            }
        }
        validate_survey_with_limits(&draft.survey, &self.limits)?;
        if draft.survey == self.draft[index].survey {
            return Ok(false);
        }
        // Existing entries were validated at ingress or their last successful
        // edit. Only this Survey can change here; ownership and other entries
        // remain immutable. Validate its bindings and catalog-wide uniqueness
        // before replacing one slot, so failures retain the exact staged state.
        if self
            .draft
            .iter()
            .enumerate()
            .any(|(other, entry)| other != index && entry.survey.id() == draft.survey.id())
        {
            return Err(invalid("survey IDs must be unique within a workbook"));
        }
        self.validate_entry_bindings(index, &draft)?;
        self.draft[index] = draft;
        Ok(true)
    }

    /// Replace one existing Survey while retaining its table relationship and
    /// physical part identity.
    pub fn set(&mut self, index: usize, survey: Survey) -> Result<bool> {
        self.edit(index, |value| {
            *value = survey;
            Ok(())
        })
    }

    /// Attach a new Survey to an existing table part.
    pub fn insert(&mut self, table_part_name: PackURI, survey: Survey) -> Result<usize> {
        validate_survey_with_limits(&survey, &self.limits)?;
        let table = self.target.get_part(&table_part_name)?;
        if table.content_type() != ct::SML_TABLE {
            return Err(invalid(format!(
                "Survey owner '{}' is not a table part",
                table_part_name
            )));
        }
        validate_table_bindings(table, &survey)?;
        if self
            .draft
            .iter()
            .any(|entry| entry.survey.id() == survey.id())
        {
            return Err(invalid("survey IDs must be unique within a workbook"));
        }
        let relationship_id = allocate_relationship_id(table, &self.draft, &table_part_name);
        let part_name = allocate_part_name(self.target, &self.draft)?;
        let relationship_target = part_name.relative_ref(table_part_name.base_uri());
        self.draft.push(Part {
            survey: Arc::new(survey),
            table_part_name,
            relationship_id,
            relationship_target,
            part_name,
            source_xml: Arc::new(Vec::new()),
            source_proof: None,
            table_source_xml: Arc::new(Vec::new()),
            table_relationships: Arc::from([]),
            survey_relationships: Arc::from([]),
            table_relationships_source: None,
            survey_relationships_source: None,
        });
        if let Err(error) = self.validate_draft() {
            self.draft.pop();
            return Err(error);
        }
        Ok(self.draft.len() - 1)
    }

    /// Attach a new Survey by semantic table name.
    ///
    /// Questions created with [`Question::for_column`] are resolved against
    /// the selected table's column names before the package relationship is
    /// authored. This keeps ordinary authoring independent of physical table
    /// part URIs and numeric column UIDs.
    pub fn insert_for_table(&mut self, table_name: &str, mut survey: Survey) -> Result<usize> {
        let table_part_name = find_table_part(self.target, table_name)?;
        let table = crate::table::parse_table_xml(self.target.get_part(&table_part_name)?.blob())?
            .ok_or_else(|| invalid("Survey owner table is not a supported table part"))?;
        for question in survey.questions_mut().values_mut() {
            question.resolve_column(&table)?;
        }
        self.insert(table_part_name, survey)
    }

    /// Remove one staged Survey and return its physical entry.
    pub fn remove(&mut self, index: usize) -> Result<Option<Part>> {
        if index >= self.draft.len() {
            return Ok(None);
        }
        let removed = self.draft.remove(index);
        if let Err(error) = self.validate_draft() {
            self.draft.insert(index, removed);
            return Err(error);
        }
        Ok(Some(removed))
    }

    /// Whether staged Survey semantics or ownership identities differ.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        !self.before.same_semantics(&self.draft)
    }

    /// Validate and atomically publish the staged Survey catalog.
    pub fn commit(self) -> Result<Commit> {
        if !self.is_changed() {
            let patch = Patch::new(self.before.clone(), self.before.clone());
            return Ok(Commit::new(self.before, patch, false));
        }
        if self.target.is_signed() || self.target.requires_signature_edit_policy() {
            return Err(Error::Signed);
        }
        let current = Snapshot::load_with_limits(self.target, &self.limits)?;
        if !current.same_source(&self.before) {
            return Err(Error::PatchConflict {
                part: "Survey catalog".into(),
            });
        }
        let mut candidate = self.target.clone();
        let content_types = transition_content_types(
            &self.before.source_content_types,
            self.before.entries(),
            &self.draft,
        )?;
        apply_entries(
            &mut candidate,
            self.before.entries(),
            &self.draft,
            &self.limits,
            &content_types,
        )?;
        let snapshot = Snapshot::load_with_limits(&candidate, &self.limits)?;
        if !snapshot.same_semantics(&self.draft) {
            return Err(invalid("Survey publication changed the staged semantics"));
        }
        let patch = Patch::new(self.before, snapshot.clone());
        *self.target = candidate;
        Ok(Commit::new(snapshot, patch, true))
    }

    fn validate_draft(&self) -> Result<()> {
        self.validate_draft_entries(&self.draft)
    }

    fn validate_draft_entries(&self, entries: &[Part]) -> Result<()> {
        if entries.len() > self.limits.max_surveys() {
            return Err(invalid("survey count exceeds the size limit"));
        }
        let mut ids = HashSet::new();
        for (index, entry) in entries.iter().enumerate() {
            validate_survey_with_limits(&entry.survey, &self.limits)?;
            if !ids.insert(entry.survey.id()) {
                return Err(invalid("survey IDs must be unique within a workbook"));
            }
            self.validate_entry_bindings(index, entry)?;
        }
        Ok(())
    }

    fn validate_entry_bindings(&self, index: usize, entry: &Part) -> Result<()> {
        let table = self.target.get_part(&entry.table_part_name)?;
        if table.content_type() != ct::SML_TABLE {
            return Err(invalid(format!(
                "Survey owner '{}' is not a table part",
                entry.table_part_name
            )));
        }
        validate_table_bindings(table, &entry.survey)
            .map_err(|error| invalid(format!("survey {index} has invalid table binding: {error}")))
    }
}

/// An exact source-checked Survey catalog replacement.
#[derive(Clone, Debug)]
pub struct Patch {
    before: Snapshot,
    after: Snapshot,
}

impl Patch {
    fn new(before: Snapshot, after: Snapshot) -> Self {
        Self { before, after }
    }

    /// Source state required before application.
    #[must_use]
    pub fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Exact state produced by application.
    #[must_use]
    pub fn after(&self) -> &Snapshot {
        &self.after
    }

    /// Whether this patch preserves the exact source graph and bytes.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.before.same_source(&self.after)
    }

    /// Return an exact source-bound inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }

    /// Apply this patch atomically after validating its source closure.
    pub fn apply(&self, target: &mut OpcPackage) -> Result<()> {
        let current = Snapshot::load_with_limits(target, &self.before.limits)?;
        if !current.same_source(&self.before) {
            return Err(Error::PatchConflict {
                part: "Survey catalog".into(),
            });
        }
        if self.is_empty() {
            return Ok(());
        }
        if target.is_signed() || target.requires_signature_edit_policy() {
            return Err(Error::Signed);
        }
        let mut candidate = target.clone();
        apply_entries(
            &mut candidate,
            self.before.entries(),
            self.after.entries(),
            &self.after.limits,
            &self.after.source_content_types,
        )?;
        restore_relationship_provenance(&mut candidate, &self.after)?;
        let resulting = Snapshot::load_with_limits(&candidate, &self.after.limits)?;
        if !resulting.same_semantics(self.after.entries()) {
            return Err(invalid("Survey patch publication changed its target state"));
        }
        *target = candidate;
        Ok(())
    }
}

/// Successful Survey transaction publication.
#[derive(Debug)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
    changed: bool,
}

impl Commit {
    fn new(snapshot: Snapshot, patch: Patch, changed: bool) -> Self {
        Self {
            snapshot,
            patch,
            changed,
        }
    }

    /// Whether Survey semantics or package ownership changed.
    #[must_use]
    pub fn changed(&self) -> bool {
        self.changed
    }

    /// Resulting source-bound catalog.
    #[must_use]
    pub fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Exact reversible patch produced by the transaction.
    #[must_use]
    pub fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Consume the commit into its snapshot and patch.
    #[must_use]
    pub fn into_parts(self) -> (Snapshot, Patch) {
        (self.snapshot, self.patch)
    }
}

fn same_identity(left: &Part, right: &Part) -> bool {
    left.table_part_name
        .is_equivalent_to(&right.table_part_name)
        && left.relationship_id == right.relationship_id
        && left.relationship_target == right.relationship_target
        && left.part_name.is_equivalent_to(&right.part_name)
}

fn allocate_relationship_id(
    table: &dyn OpcPart,
    entries: &[Part],
    table_part_name: &PackURI,
) -> String {
    let mut used = table
        .rels()
        .iter()
        .map(|relationship| relationship.r_id().to_owned())
        .collect::<HashSet<_>>();
    used.extend(
        entries
            .iter()
            .filter(|entry| entry.table_part_name.is_equivalent_to(table_part_name))
            .map(|entry| entry.relationship_id.clone()),
    );
    let mut candidate = "rIdSurvey".to_owned();
    let mut suffix = 2u32;
    while !used.insert(candidate.clone()) {
        candidate = format!("rIdSurvey{suffix}");
        suffix = suffix.saturating_add(1);
    }
    candidate
}

/// Whether the package holds a part named `name`.
///
/// Only `PartNotFound` proves absence; any other refusal from an existing
/// part propagates instead of reading as a free name (ADR 0030).
fn part_exists(package: &OpcPackage, name: &PackURI) -> Result<bool> {
    match package.get_part(name) {
        Ok(_) => Ok(true),
        Err(litchi_opc::OpcError::PartNotFound(_)) => Ok(false),
        Err(error) => Err(error.into()),
    }
}

fn allocate_part_name(package: &OpcPackage, entries: &[Part]) -> Result<PackURI> {
    let mut index = 1u32;
    loop {
        let candidate = PackURI::new(format!("/xl/surveys/survey{index}.xml"))
            .map_err(|error| invalid(error.to_string()))?;
        if !part_exists(package, &candidate)?
            && !entries
                .iter()
                .any(|entry| entry.part_name.is_equivalent_to(&candidate))
        {
            return Ok(candidate);
        }
        index = index
            .checked_add(1)
            .ok_or_else(|| invalid("Survey part-name space exhausted"))?;
    }
}

fn find_table_part(package: &OpcPackage, table_name: &str) -> Result<PackURI> {
    if table_name.is_empty() {
        return Err(invalid("Survey table name cannot be empty"));
    }
    let mut match_name = None;
    for part in package
        .iter_parts()
        .filter(|part| part.content_type() == ct::SML_TABLE)
    {
        // Only a table part's payload is parsed, so only table parts are
        // decoded (ADR 0030).
        let Some(table) = crate::table::parse_table_xml(package.get_part(part.partname())?.blob())?
        else {
            continue;
        };
        if table.name.eq_ignore_ascii_case(table_name)
            || table.display_name.eq_ignore_ascii_case(table_name)
        {
            if match_name.is_some() {
                return Err(invalid(format!(
                    "Survey table selector '{table_name}' is ambiguous"
                )));
            }
            match_name = Some(part.partname().clone());
        }
    }
    match_name.ok_or_else(|| invalid(format!("Survey table '{table_name}' is absent")))
}

fn apply_entries(
    package: &mut OpcPackage,
    before: &[Part],
    after: &[Part],
    limits: &Limits,
    content_types: &OwnedContentTypes,
) -> Result<()> {
    for entry in before {
        if after
            .iter()
            .any(|candidate| same_identity(entry, candidate))
        {
            continue;
        }
        remove_entry(package, entry)?;
    }

    for entry in after {
        if let Some(previous) = before
            .iter()
            .find(|candidate| same_identity(candidate, entry))
        {
            if previous.survey != entry.survey {
                replace_entry(package, previous, entry, limits)?;
            }
            continue;
        }
        add_entry(package, entry, limits)?;
    }
    let current_content_types = package.source_content_types()?;
    package.try_replace_content_types(current_content_types.bytes(), content_types)?;
    Ok(())
}

fn transition_content_types(
    before: &OwnedContentTypes,
    before_entries: &[Part],
    after_entries: &[Part],
) -> Result<OwnedContentTypes> {
    let mut removed = Vec::new();
    for entry in before_entries {
        if !after_entries
            .iter()
            .any(|candidate| same_identity(entry, candidate))
        {
            removed.push(entry.part_name.clone());
        }
    }
    let mut additions = Vec::new();
    for entry in after_entries {
        if !before_entries
            .iter()
            .any(|candidate| same_identity(candidate, entry))
        {
            additions.push((&entry.part_name, CONTENT_TYPE));
        }
    }
    let maximum = litchi_opc::ReadLimits::default().max_content_types_bytes();
    let mut result = before.without_parts(&removed, maximum)?;
    if !additions.is_empty() {
        result = result.with_part_overrides(&additions, maximum)?;
    }
    Ok(result)
}

fn replace_entry(
    package: &mut OpcPackage,
    previous: &Part,
    entry: &Part,
    limits: &Limits,
) -> Result<()> {
    if !entry.source_xml.is_empty() && entry.source_xml != previous.source_xml {
        let proof = entry
            .source_proof
            .as_ref()
            .ok_or_else(|| invalid("Survey patch lacks source provenance"))?;
        package
            .try_replace_owned_xml_part(previous.source_xml.as_slice(), proof.as_ref().clone())?;
        return Ok(());
    }
    let proof = previous
        .source_proof
        .as_ref()
        .ok_or_else(|| invalid("Survey source provenance is missing"))?;
    if let Some(updated) = source::rewrite(proof, &previous.survey, &entry.survey, limits)? {
        package.try_replace_owned_xml_part(previous.source_xml.as_slice(), updated)?;
    } else {
        let xml = write_with_limits(&entry.survey, limits)?;
        package.get_part_mut(&entry.part_name)?.set_blob(xml);
    }
    Ok(())
}

fn add_entry(package: &mut OpcPackage, entry: &Part, limits: &Limits) -> Result<()> {
    let table = package.get_part(&entry.table_part_name)?;
    if table.content_type() != ct::SML_TABLE {
        return Err(invalid("Survey owner is not a table part"));
    }
    if table.rels().get(&entry.relationship_id).is_some() {
        return Err(invalid(format!(
            "Survey relationship ID '{}' already exists",
            entry.relationship_id
        )));
    }
    if part_exists(package, &entry.part_name)? {
        return Err(invalid(format!(
            "Survey part '{}' already exists",
            entry.part_name
        )));
    }
    // Capture the complete owner member before adding the new edge. The
    // relationship token keeps comments, PIs, prefixes, attribute order and
    // explicit empty-member presence when the edge is spliced in.
    let table_relationships = package.source_relationships(&entry.table_part_name)?;
    if let Some(proof) = &entry.source_proof {
        package.try_add_owned_xml_part(proof.as_ref().clone())?;
    } else {
        let xml = write_with_limits(&entry.survey, limits)?;
        package.try_add_part(Box::new(BlobPart::new(
            entry.part_name.clone(),
            CONTENT_TYPE.into(),
            xml,
        )))?;
    }
    let table_replacement = table_relationships.with_relationship(
        RELATIONSHIP_TYPE,
        &entry.relationship_target,
        &entry.relationship_id,
        TargetMode::Internal,
        usize::MAX,
    )?;
    package.try_replace_relationships(&table_relationships, &table_replacement)?;
    if let Some(source) = entry.survey_relationships_source.as_ref() {
        let current = package.source_relationships(&entry.part_name)?;
        package.try_replace_relationships(&current, source)?;
    }
    Ok(())
}

fn remove_entry(package: &mut OpcPackage, entry: &Part) -> Result<()> {
    let table = package.get_part(&entry.table_part_name)?;
    let relationship = table
        .rels()
        .get(&entry.relationship_id)
        .ok_or_else(|| invalid("Survey owner relationship is absent"))?;
    let target = relationship.target_partname()?;
    if !target.is_equivalent_to(&entry.part_name) {
        return Err(invalid("Survey owner relationship target changed"));
    }
    if relationship.target_ref() != entry.relationship_target {
        return Err(invalid("Survey owner relationship lexical target changed"));
    }
    let table_relationships = package.source_relationships(&entry.table_part_name)?;
    let table_replacement =
        table_relationships.without_relationship(&entry.relationship_id, usize::MAX)?;
    package.try_replace_relationships(&table_relationships, &table_replacement)?;
    if part_is_referenced(package, &entry.part_name) {
        return Err(invalid("Survey part has an unexpected relationship owner"));
    }
    if !package.remove_part(&entry.part_name) {
        return Err(invalid("Survey part is absent"));
    }
    Ok(())
}

fn part_is_referenced(package: &OpcPackage, part_name: &PackURI) -> bool {
    package.rels().iter().any(|relationship| {
        !relationship.is_external()
            && relationship
                .target_partname()
                .ok()
                .is_some_and(|target| target.is_equivalent_to(part_name))
    }) || package.iter_parts().any(|part| {
        part.rels().iter().any(|relationship| {
            !relationship.is_external()
                && relationship
                    .target_partname()
                    .ok()
                    .is_some_and(|target| target.is_equivalent_to(part_name))
        })
    })
}

fn restore_relationship_provenance(package: &mut OpcPackage, snapshot: &Snapshot) -> Result<()> {
    let restore = |package: &mut OpcPackage, source: &OwnedRelationships| -> Result<()> {
        let current = package.source_relationships(source.owner())?;
        package.try_replace_relationships(&current, source)?;
        Ok(())
    };
    for entry in snapshot.entries.iter() {
        if let Some(source) = entry.table_relationships_source.as_ref() {
            restore(package, source)?;
        }
        if let Some(source) = entry.survey_relationships_source.as_ref() {
            restore(package, source)?;
        }
    }
    Ok(())
}

#[cfg(test)]
#[allow(
    clippy::expect_used,
    reason = "test fixture setup uses explicit expectations to keep failure diagnostics local"
)]
mod tests {
    use super::*;
    use litchi_opc::{BlobPart, PackURI, TargetMode};

    const XML: &[u8] = br#"<survey xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" id="7" guid="{01234567-89ab-cdef-0123-456789abcdef}" title="Survey"><surveyPr cssClass="card" top="-2" width="320" position="relative"/><questions><questionsPr height="20"/><question binding="1" text="Pick one" type="choice" format="standard" required="true" decimalPlaces="15" rowSource="one;&quot;two;three&quot;;"><questionPr left="1"/></question></questions></survey>"#;

    fn as_source(package: &OpcPackage) -> OpcPackage {
        let mut writer = soapberry_zip::office::StreamingArchiveWriter::new();
        let mut parts = package
            .try_iter_parts()
            .map(|part| part.expect("part payload decodes"))
            .collect::<Vec<_>>();
        parts.sort_by(|a, b| a.partname().as_str().cmp(b.partname().as_str()));
        let mut types = String::from(
            r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>"#,
        );
        for part in &parts {
            types.push_str(&format!(
                r#"<Override PartName="{}" ContentType="{}"/>"#,
                part.partname(),
                part.content_type()
            ));
            writer
                .write_stored(
                    part.partname().as_str().trim_start_matches('/'),
                    part.blob(),
                )
                .unwrap();
            if !part.rels().is_empty() {
                writer
                    .write_stored(
                        part.partname()
                            .rels_uri()
                            .unwrap()
                            .as_str()
                            .trim_start_matches('/'),
                        part.rels().to_xml().as_bytes(),
                    )
                    .unwrap();
            }
        }
        types.push_str("</Types>");
        writer
            .write_stored("[Content_Types].xml", types.as_bytes())
            .unwrap();
        writer
            .write_stored("_rels/.rels", package.rels().to_xml().as_bytes())
            .unwrap();
        OpcPackage::from_bytes(&writer.finish_to_bytes().unwrap()).unwrap()
    }

    fn schema_xml(body: &str) -> Vec<u8> {
        format!(r#"<survey xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" id="7" guid="{{01234567-89ab-cdef-0123-456789abcdef}}">{body}</survey>"#).into_bytes()
    }

    #[test]
    fn catalog_staging_shares_untouched_surveys_and_preserves_failed_edit_identity() {
        let mut package = package_with_survey(XML);
        let mut second = authored_survey(8);
        second.questions_mut().values_mut()[0].set_text(Some("payload".repeat(8_192)));
        let mut setup = Transaction::new(&mut package).unwrap();
        setup.insert_for_table("T", second).unwrap();
        setup.commit().unwrap();
        let mut transaction = Transaction::new(&mut package).unwrap();
        let before = transaction.before().clone();
        for (source, draft) in before.entries().iter().zip(transaction.entries()) {
            assert!(Arc::ptr_eq(&source.survey, &draft.survey));
            assert!(Arc::ptr_eq(&source.survey, &source.clone().survey));
        }
        let payload = before.entries()[1].survey().questions().values()[0]
            .text()
            .unwrap()
            .as_ptr();
        transaction
            .edit(0, |survey| {
                survey.set_title(Some("changed".into()));
                Ok(())
            })
            .unwrap();
        assert_eq!(before.entries()[0].survey().title(), Some("Survey"));
        assert!(!Arc::ptr_eq(
            &before.entries()[0].survey,
            &transaction.entries()[0].survey
        ));
        assert!(Arc::ptr_eq(
            &before.entries()[1].survey,
            &transaction.entries()[1].survey
        ));
        assert_eq!(
            transaction.entries()[1].survey().questions().values()[0]
                .text()
                .unwrap()
                .as_ptr(),
            payload
        );
        let accepted = transaction.entries()[0].survey.clone();
        assert!(
            transaction
                .edit(0, |survey| {
                    survey.set_id(Id::new(8));
                    Ok(())
                })
                .is_err()
        );
        assert!(Arc::ptr_eq(&accepted, &transaction.entries()[0].survey));
        assert!(
            transaction
                .edit(0, |survey| {
                    survey.questions_mut().values_mut()[0].set_binding(Binding::new(99));
                    Ok(())
                })
                .is_err()
        );
        assert!(Arc::ptr_eq(&accepted, &transaction.entries()[0].survey));
        assert!(
            transaction
                .edit(0, |survey| {
                    survey.set_title(Some("failed callback".into()));
                    Err(invalid("injected callback failure"))
                })
                .is_err()
        );
        assert!(Arc::ptr_eq(&accepted, &transaction.entries()[0].survey));
        assert!(!transaction.edit(0, |_| Ok(())).unwrap());
        assert!(Arc::ptr_eq(&accepted, &transaction.entries()[0].survey));
        assert!(Arc::ptr_eq(
            &before.entries()[1].survey,
            &transaction.entries()[1].survey
        ));
        let patch = transaction.commit().unwrap().patch().clone();
        patch.inverse().apply(&mut package).unwrap();
        assert!(Snapshot::load(&package).unwrap().same_source(&before));
    }

    #[test]
    fn question_sequence_edits_preserve_source_context_and_exact_inverse() {
        let namespace = std::str::from_utf8(NAMESPACE).unwrap();
        let first = "<s:question binding = '1' text='first' xmlns:q='urn:one' q:tag='keep'><s:questionPr width = '9'/><s:extLst><node type='q:Only'/></s:extLst><!--owned0--></s:question>";
        let second = "<s:question binding='1' text='second'><!--owned1--><s:extLst><node/></s:extLst></s:question>";
        let xml = format!(
            "<s:survey xmlns:s=\"{namespace}\" xmlns=\"urn:vendor\" id = '7' guid='{{01234567-89ab-cdef-0123-456789abcdef}}' title='old'><!--root--><s:questions xml:lang='en' xml:base='relative/'><!--a-->{first}<!--b-->{second}<!--tail--></s:questions></s:survey>"
        );
        let mut package =
            crate::package::Package::from_opc(as_source(&package_with_survey(xml.as_bytes())))
                .unwrap();
        let before = Snapshot::load(&package.clone().into_plain_opc()).unwrap();
        let mut transaction = package.edit_surveys().unwrap();
        transaction
            .edit(0, |survey| {
                survey.set_title(Some("new".into()));
                let questions = survey.questions_mut();
                let mut first = questions.remove(0).unwrap();
                first.set_text(Some("first moved".into()));
                first.set_properties(None);
                let mut copy = questions.values()[0].clone();
                copy.set_text(Some("second copy".into()));
                let mut new = Question::for_column("Answer")?;
                new.set_text(Some("new&_x0041_".into()));
                questions.insert(1, new)?;
                questions.push(first)?;
                questions.push(copy)?;
                Ok(())
            })
            .unwrap();
        let patch = transaction.commit().unwrap().patch().clone();
        let first_moved = first
            .replace("text='first'", "text='first moved'")
            .replace("<s:questionPr width = '9'/>", "");
        let second_copy = second.replace("text='second'", "text='second copy'");
        let new = format!(
            "<question xmlns=\"{namespace}\" binding=\"1\" text=\"new&amp;_x005F_x0041_\"/>"
        );
        let expected = xml.replace("title='old'", "title='new'").replace(
            &format!("{first}<!--b-->{second}"),
            &format!("{second}<!--b-->{new}{first_moved}{second_copy}"),
        );
        let bytes = litchi_opc::PackageWriter::to_bytes(&package.into_plain_opc()).unwrap();
        let mut reopened = OpcPackage::from_bytes(&bytes).unwrap();
        let name = PackURI::new("/xl/surveys/survey1.xml").unwrap();
        assert_eq!(
            reopened.get_part(&name).unwrap().blob(),
            expected.as_bytes()
        );
        let after = Snapshot::load(&reopened).unwrap();
        assert_eq!(
            after.entries()[0]
                .survey()
                .questions()
                .values()
                .iter()
                .map(Question::text)
                .collect::<Vec<_>>(),
            [
                Some("second"),
                Some("new&_x0041_"),
                Some("first moved"),
                Some("second copy")
            ]
        );
        patch.inverse().apply(&mut reopened).unwrap();
        let bytes = litchi_opc::PackageWriter::to_bytes(&reopened).unwrap();
        let restored = OpcPackage::from_bytes(&bytes).unwrap();
        assert!(Snapshot::load(&restored).unwrap().same_source(&before));
        assert_eq!(restored.get_part(&name).unwrap().blob(), xml.as_bytes());
    }

    #[test]
    fn question_collection_edits_move_allocations_and_cancel_exactly() {
        let mut survey = parse(XML).unwrap();
        let original = survey.clone();
        let questions = survey.questions_mut();
        let text = questions.values()[0].text().unwrap().as_ptr();
        questions.push(Question::new(Binding::new(1))).unwrap();
        assert_eq!(questions.values()[0].text().unwrap().as_ptr(), text);
        questions.remove(1).unwrap();
        let first = questions.remove(0).unwrap();
        questions.insert(0, first).unwrap();
        assert_eq!(questions.values()[0].text().unwrap().as_ptr(), text);
        assert_eq!(survey, original);
        assert!(
            survey
                .questions_mut()
                .insert(usize::MAX, Question::new(Binding::new(1)))
                .is_err()
        );
        assert_eq!(survey, original);
        let mut package =
            crate::package::Package::from_opc(as_source(&package_with_survey(XML))).unwrap();
        let before = Snapshot::load(&package.clone().into_plain_opc()).unwrap();
        let mut transaction = package.edit_surveys().unwrap();
        assert!(
            transaction
                .edit(0, |survey| {
                    survey
                        .questions_mut()
                        .push(Question::for_column("missing column")?)
                })
                .is_err()
        );
        assert!(!transaction.is_changed());
        transaction
            .edit(0, |survey| {
                let question = survey.questions_mut().remove(0).unwrap();
                survey.questions_mut().insert(0, question)?;
                Ok(())
            })
            .unwrap();
        let commit = transaction.commit().unwrap();
        assert!(!commit.changed());
        assert!(commit.patch().is_empty());
        assert!(
            Snapshot::load(&package.into_plain_opc())
                .unwrap()
                .same_source(&before)
        );
    }

    #[test]
    fn foreign_opaque_question_transplant_refusal_is_atomic() {
        let mut package =
            crate::package::Package::from_opc(as_source(&package_with_survey(XML))).unwrap();
        let before = Snapshot::load(&package.clone().into_plain_opc()).unwrap();
        let foreign = parse(&schema_xml("<questions><question binding=\"1\"><extLst><opaque xmlns=\"urn:other\"/></extLst></question></questions>")).unwrap();
        let mut transaction = package.edit_surveys().unwrap();
        transaction
            .edit(0, |survey| {
                survey
                    .questions_mut()
                    .push(foreign.questions.values[0].clone())
            })
            .unwrap();
        assert!(matches!(
            transaction.commit(),
            Err(Error::Unsupported {
                feature: "Survey question source transplantation"
            })
        ));
        assert!(
            Snapshot::load(&package.into_plain_opc())
                .unwrap()
                .same_source(&before)
        );
    }

    #[test]
    fn property_insertions_preserve_foreign_default_context_and_inverse() {
        let namespace = std::str::from_utf8(NAMESPACE).unwrap();
        let xml = r#"<s:survey xmlns:s="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" xmlns="urn:vendor" id = '7' guid='{01234567-89ab-cdef-0123-456789abcdef}'>
<!--root--><s:questions><!--list--><s:question binding = '1' /><!--tail--></s:questions></s:survey>"#;
        let mut package =
            crate::package::Package::from_opc(as_source(&package_with_survey(xml.as_bytes())))
                .unwrap();
        let before = Snapshot::load(&package.clone().into_plain_opc()).unwrap();
        let mut transaction = package.edit_surveys().unwrap();
        transaction
            .edit(0, |survey| {
                survey.set_properties(Some(ElementProperties::new()));
                survey.set_title_properties(Some(ElementProperties::new()));
                survey.set_description_properties(Some(ElementProperties::new()));
                survey
                    .questions
                    .set_properties(Some(ElementProperties::new()));
                survey.questions.values_mut()[0].set_properties(Some(ElementProperties::new()));
                Ok(())
            })
            .unwrap();
        let patch = transaction.commit().unwrap().patch().clone();
        let element = |name: &str| format!("<{name} xmlns=\"{namespace}\"/>");
        let expected = xml
            .replace(
                "<s:questions>",
                &format!(
                    "{}{}{}<s:questions>",
                    element("surveyPr"),
                    element("titlePr"),
                    element("descriptionPr")
                ),
            )
            .replace(
                "<s:question binding = '1' />",
                &format!(
                    "{}<s:question binding = '1' >{}</s:question>",
                    element("questionsPr"),
                    element("questionPr")
                ),
            );
        let bytes = litchi_opc::PackageWriter::to_bytes(&package.into_plain_opc()).unwrap();
        let mut reopened = OpcPackage::from_bytes(&bytes).unwrap();
        let name = PackURI::new("/xl/surveys/survey1.xml").unwrap();
        assert_eq!(
            reopened.get_part(&name).unwrap().blob(),
            expected.as_bytes()
        );
        let snapshot = Snapshot::load(&reopened).unwrap();
        let survey = snapshot.entries()[0].survey();
        assert!(survey.properties.is_some());
        assert!(survey.title_properties.is_some());
        assert!(survey.description_properties.is_some());
        assert!(survey.questions.properties.is_some());
        assert!(survey.questions.values[0].properties.is_some());
        patch.inverse().apply(&mut reopened).unwrap();
        let bytes = litchi_opc::PackageWriter::to_bytes(&reopened).unwrap();
        let restored = OpcPackage::from_bytes(&bytes).unwrap();
        assert!(Snapshot::load(&restored).unwrap().same_source(&before));
        assert_eq!(restored.get_part(&name).unwrap().blob(), xml.as_bytes());
    }

    #[test]
    fn property_removal_replacement_and_insertion_keep_adjacent_source_xml() {
        let namespace = std::str::from_utf8(NAMESPACE).unwrap();
        let xml = r#"<s:survey xmlns:s="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" xmlns="urn:vendor" id = '7' guid='{01234567-89ab-cdef-0123-456789abcdef}' title='old'>
<s:surveyPr><!--owned--><s:extLst><opaque/></s:extLst></s:surveyPr><!--a--><s:titlePr cssClass='old'/><!--b-->
<s:questions><s:questionsPr/><!--list--><s:question binding='1'><!--q--><s:extLst><opaque/></s:extLst></s:question></s:questions></s:survey>"#;
        let mut package =
            crate::package::Package::from_opc(as_source(&package_with_survey(xml.as_bytes())))
                .unwrap();
        let before = Snapshot::load(&package.clone().into_plain_opc()).unwrap();
        let mut transaction = package.edit_surveys().unwrap();
        transaction
            .edit(0, |survey| {
                survey.set_title(Some("new".into()));
                let mut fresh = ElementProperties::new();
                fresh.set_css_class(Some("_x0041_&".into()));
                survey.set_properties(Some(fresh));
                survey.set_title_properties(None);
                survey.set_description_properties(Some(ElementProperties::new()));
                survey.questions.set_properties(None);
                survey.questions.values_mut()[0].set_properties(Some(ElementProperties::new()));
                Ok(())
            })
            .unwrap();
        let patch = transaction.commit().unwrap().patch().clone();
        let expected = xml
            .replace("title='old'", "title='new'")
            .replace(
                "<s:surveyPr><!--owned--><s:extLst><opaque/></s:extLst></s:surveyPr>",
                &format!("<surveyPr xmlns=\"{namespace}\" cssClass=\"_x005F_x0041_&amp;\"/>"),
            )
            .replace("<s:titlePr cssClass='old'/>", "")
            .replace(
                "<s:questions>",
                &format!("<descriptionPr xmlns=\"{namespace}\"/><s:questions>"),
            )
            .replace("<s:questionsPr/>", "")
            .replace(
                "<!--q--><s:extLst>",
                &format!("<!--q--><questionPr xmlns=\"{namespace}\"/><s:extLst>"),
            );
        let bytes = litchi_opc::PackageWriter::to_bytes(&package.into_plain_opc()).unwrap();
        let mut reopened = OpcPackage::from_bytes(&bytes).unwrap();
        let name = PackURI::new("/xl/surveys/survey1.xml").unwrap();
        assert_eq!(
            reopened.get_part(&name).unwrap().blob(),
            expected.as_bytes()
        );
        patch.inverse().apply(&mut reopened).unwrap();
        let bytes = litchi_opc::PackageWriter::to_bytes(&reopened).unwrap();
        assert!(
            Snapshot::load(&OpcPackage::from_bytes(&bytes).unwrap())
                .unwrap()
                .same_source(&before)
        );
    }

    #[test]
    fn opaque_property_transplant_refusal_is_atomic() {
        let xml = schema_xml(
            "<surveyPr><extLst><opaque xmlns=\"urn:vendor\"/></extLst></surveyPr><questions><question binding=\"1\"/></questions>",
        );
        let mut package =
            crate::package::Package::from_opc(as_source(&package_with_survey(&xml))).unwrap();
        let before = Snapshot::load(&package.clone().into_plain_opc()).unwrap();
        let mut transaction = package.edit_surveys().unwrap();
        transaction
            .edit(0, |survey| {
                survey.set_title_properties(survey.properties.clone());
                Ok(())
            })
            .unwrap();
        assert!(matches!(
            transaction.commit(),
            Err(Error::Unsupported {
                feature: "Survey property source transplantation"
            })
        ));
        assert!(
            Snapshot::load(&package.into_plain_opc())
                .unwrap()
                .same_source(&before)
        );
    }

    #[test]
    fn scalar_edits_preserve_noncompact_source_context_and_inverse_publication() {
        let xml = r#"<s:survey xmlns:s="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" xmlns="urn:default" xmlns:v="urn:vendor" id = '7' guid='{01234567-89ab-cdef-0123-456789abcdef}' title = 'root-old'>
<!--before--><s:surveyPr cssClass='css-old' bottom = '1' top='2' left='3' right='4' width='5' height='6' position='absolute'><s:extLst><vendor type='v:QName'/></s:extLst><!--property-end--></s:surveyPr>
<!--between--><s:titlePr cssClass='heading'/><s:descriptionPr height='8'/>
<s:questions xmlns:q='urn:question'><!--questions-start--><s:questionsPr height='9'/><s:question binding='1' text='q-old' type='choice' format='standard' helpText='help-old' required='1' defaultValue='default-old' decimalPlaces='4' rowSource='a;b' q:opaque='keep'>
<s:questionPr left='10'/><!--question-middle--><s:extLst><s:question binding='999' text='opaque'/><vendor type='q:OnlyInValue'/></s:extLst><!--question-end--></s:question></s:questions><!--after--></s:survey>"#;
        let mut opc = package_with_survey(xml.as_bytes());
        let table_name = PackURI::new("/xl/tables/table1.xml").unwrap();
        let table = String::from_utf8(opc.get_part(&table_name).unwrap().blob().to_vec())
            .unwrap()
            .replace("A1:A2", "A1:B2")
            .replace("count=\"1\"", "count=\"2\"")
            .replace(
                "</tableColumns>",
                "<tableColumn id=\"2\" name=\"Second\"/></tableColumns>",
            );
        opc.get_part_mut(&table_name)
            .unwrap()
            .set_blob(table.into_bytes());
        let mut package = crate::package::Package::from_opc(as_source(&opc)).unwrap();
        let before = Snapshot::load(&package.clone().into_plain_opc()).unwrap();
        let mut transaction = package.edit_surveys().unwrap();
        transaction
            .edit(0, |survey| {
                survey.set_id(Id::new(8));
                survey.set_guid(Guid::new("{11234567-89ab-cdef-0123-456789abcdef}")?);
                survey.set_title(Some("root-new".into()));
                survey.set_description(Some("added _x0041_".into()));
                let properties = survey.properties.as_mut().unwrap();
                properties.set_css_class(Some("new&css".into()));
                properties.set_bottom(None);
                properties.set_top(Some(12));
                properties.set_left(Some(13));
                properties.set_right(Some(14));
                properties.set_width(Some(15));
                properties.set_height(Some(16));
                properties.set_position(Some(Position::Fixed));
                survey
                    .title_properties
                    .as_mut()
                    .unwrap()
                    .set_css_class(None);
                survey
                    .description_properties
                    .as_mut()
                    .unwrap()
                    .set_height(Some(18));
                survey
                    .questions
                    .properties
                    .as_mut()
                    .unwrap()
                    .set_height(Some(19));
                let question = &mut survey.questions.values_mut()[0];
                question.set_binding(Binding::new(2));
                question.set_text(Some("q-new".into()));
                question.set_question_type(Some(QuestionType::Number));
                question.set_format(Some(QuestionFormat::Fixed));
                question.set_help_text(Some("help-new".into()));
                question.set_required(false);
                question.set_default_value(None);
                question.set_decimal_places(Some(5));
                question.set_row_source(Some("c\"q;_x0000_".into()))?;
                question.properties.as_mut().unwrap().set_left(Some(20));
                Ok(())
            })
            .unwrap();
        let patch = transaction.commit().unwrap().patch().clone();
        let mut expected = xml.to_owned();
        for (from, to) in [
            ("id = '7'", "id = '8'"),
            (
                "{01234567-89ab-cdef-0123-456789abcdef}",
                "{11234567-89ab-cdef-0123-456789abcdef}",
            ),
            (
                "title = 'root-old'>",
                "title = 'root-new' description=\"added _x005F_x0041_\">",
            ),
            ("cssClass='css-old'", "cssClass='new&amp;css'"),
            ("bottom = '1'", ""),
            ("top='2'", "top='12'"),
            ("left='3'", "left='13'"),
            ("right='4'", "right='14'"),
            ("width='5'", "width='15'"),
            ("height='6'", "height='16'"),
            ("position='absolute'", "position='fixed'"),
            ("cssClass='heading'", ""),
            ("height='8'", "height='18'"),
            ("height='9'", "height='19'"),
            ("binding='1'", "binding='2'"),
            ("text='q-old'", "text='q-new'"),
            ("type='choice'", "type='number'"),
            ("format='standard'", "format='fixed'"),
            ("helpText='help-old'", "helpText='help-new'"),
            ("required='1'", "required='false'"),
            ("defaultValue='default-old'", ""),
            ("decimalPlaces='4'", "decimalPlaces='5'"),
            ("rowSource='a;b'", "rowSource='c&quot;q;_x005F_x0000_'"),
            ("left='10'", "left='20'"),
        ] {
            expected = expected.replace(from, to);
        }
        let bytes = litchi_opc::PackageWriter::to_bytes(&package.into_plain_opc()).unwrap();
        let mut reopened = OpcPackage::from_bytes(&bytes).unwrap();
        let name = PackURI::new("/xl/surveys/survey1.xml").unwrap();
        assert_eq!(
            reopened.get_part(&name).unwrap().blob(),
            expected.as_bytes()
        );
        patch.inverse().apply(&mut reopened).unwrap();
        let bytes = litchi_opc::PackageWriter::to_bytes(&reopened).unwrap();
        let restored = OpcPackage::from_bytes(&bytes).unwrap();
        assert!(Snapshot::load(&restored).unwrap().same_source(&before));
        assert_eq!(restored.get_part(&name).unwrap().blob(), xml.as_bytes());
    }

    #[test]
    fn removing_a_survey_restores_its_source_proof_for_inverse_save() {
        let xml = String::from_utf8(XML.to_vec())
            .unwrap()
            .replace("title=\"Survey\"", "title = 'Survey'");
        let mut package =
            crate::package::Package::from_opc(as_source(&package_with_survey(xml.as_bytes())))
                .unwrap();
        let before = Snapshot::load(&package.clone().into_plain_opc()).unwrap();
        let mut transaction = package.edit_surveys().unwrap();
        transaction.remove(0).unwrap();
        let patch = transaction.commit().unwrap().patch().clone();
        let bytes = litchi_opc::PackageWriter::to_bytes(&package.into_plain_opc()).unwrap();
        let mut reopened = OpcPackage::from_bytes(&bytes).unwrap();
        patch.inverse().apply(&mut reopened).unwrap();
        let bytes = litchi_opc::PackageWriter::to_bytes(&reopened).unwrap();
        let restored = OpcPackage::from_bytes(&bytes).unwrap();
        assert!(Snapshot::load(&restored).unwrap().same_source(&before));
    }

    #[test]
    fn xstring_output_budget_charges_encoded_growth_before_allocation() {
        let mut survey = parse(XML).unwrap();
        survey.set_title(Some("\u{1}".repeat(1024)));
        let bound = serialized_size_bound(&survey).unwrap();
        let output =
            write_with_limits(&survey, &Limits::default().with_max_part_bytes(bound)).unwrap();
        assert!(output.len() <= bound);
        assert!(
            write_with_limits(&survey, &Limits::default().with_max_part_bytes(bound - 1)).is_err()
        );
        assert_eq!(parse(&output).unwrap().title(), survey.title());
        assert!(xstring(&"a".repeat(MAX_TEXT_BYTES + 1), "title").is_err());
        assert!(xstring(&"a".repeat(MAX_TEXT_BYTES * 7 + 1), "title").is_err());

        for (value, expected) in [
            ("_plain_", 7),
            ("_x0041_", 13),
            ("_x005F_x0041_", 25),
            ("\0", 7),
            ("\u{fffe}", 7),
            ("\t\n\r", 15),
            ("😀", 4),
        ] {
            let mut size = 0;
            add_xstring_size(&mut size, value).unwrap();
            assert_eq!(size, expected, "{value:?}");
            assert_eq!(crate::source_attributes::escaped_xstring(value).len(), size);
        }
    }

    #[test]
    fn row_source_accepts_the_normative_unquoted_quote_production() {
        for value in [
            "",
            ";",
            ";;",
            "a\"b",
            "\"",
            "\"unterminated",
            "\"a\"suffix",
            "one;\"two;three\";",
            "\";\"",
            "😀;\t\n\r",
            "_x0000_",
        ] {
            let mut survey = parse(XML).unwrap();
            survey.questions.values[0]
                .set_row_source(Some(value.into()))
                .unwrap();
            let output = write(&survey).unwrap();
            assert_eq!(
                parse(&output).unwrap().questions.values[0].row_source(),
                Some(value)
            );
        }
        for value in ["\0", "\u{1}", "\u{fffe}", "\u{ffff}"] {
            assert!(
                Question::new(Binding::new(1))
                    .set_row_source(Some(value.into()))
                    .is_err()
            );
        }
        for encoded in ["_x0000_", "_xFFFE_", "_xFFFF_"] {
            let xml = schema_xml(&format!(
                r#"<questions><question binding="1" rowSource="{encoded}"/></questions>"#
            ));
            assert!(parse(&xml).is_err());
        }
    }

    #[test]
    fn known_xstrings_decode_once_and_foreign_attributes_remain_opaque() {
        let encoded = "_x0041__x005F_x0042__xD83D__xDE00__x0001__xFFFE_";
        let expected = "A_x0042_😀\u{1}\u{fffe}";
        let xml = format!(
            r#"<survey xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" xmlns:q="urn:vendor" id="7" guid="{{01234567-89ab-cdef-0123-456789abcdef}}" title="{encoded}" description="{encoded}" q:title="_x0041_"><surveyPr cssClass="{encoded}"/><questions><question binding="1" text="{encoded}" helpText="{encoded}" defaultValue="{encoded}" rowSource="_x003B__x0022__x005F_x0000_"/></questions></survey>"#
        );
        let survey = parse(xml.as_bytes()).unwrap();
        for value in [
            survey.title(),
            survey.description(),
            survey.properties.as_ref().unwrap().css_class.as_deref(),
            survey.questions.values[0].text(),
            survey.questions.values[0].help_text(),
            survey.questions.values[0].default_value(),
        ] {
            assert_eq!(value, Some(expected));
        }
        assert_eq!(survey.questions.values[0].row_source(), Some(";\"_x0000_"));
        assert_eq!(survey.unknown_attributes[0].value.as_ref(), "_x0041_");
        let output = write(&survey).unwrap();
        assert_eq!(parse(&output).unwrap(), survey);
        assert!(String::from_utf8_lossy(&output).contains(r#"q:title="_x0041_""#));
    }

    #[test]
    fn xstring_authoring_preserves_controls_literal_escapes_and_attribute_whitespace() {
        let mut survey = parse(XML).unwrap();
        let text = "_x0041_ _x005F_ _xD800_ & < > \" ' \t\n\r\0\u{1}\u{fffe}\u{ffff}😀";
        survey.set_title(Some(text.into()));
        survey.set_description(Some(text.into()));
        survey
            .properties
            .as_mut()
            .unwrap()
            .set_css_class(Some(text.into()));
        survey.questions.values[0].set_text(Some(text.into()));
        survey.questions.values[0].set_help_text(Some(text.into()));
        survey.questions.values[0].set_default_value(Some(text.into()));
        survey.unknown_attributes = vec![RawAttribute {
            name: "opaque".into(),
            value: "_x0041_\t\n\r".into(),
        }]
        .into_boxed_slice();
        let output = write(&survey).unwrap();
        assert!(output.len() <= serialized_size_bound(&survey).unwrap());
        assert_eq!(parse(&output).unwrap(), survey);
        assert!(
            !output
                .iter()
                .any(|byte| matches!(byte, 0 | 1 | b'\t' | b'\n' | b'\r'))
        );
        assert!(String::from_utf8_lossy(&output).contains("_x005F_x0041_"));
    }

    #[test]
    fn malformed_xml_and_xstring_surrogates_are_rejected_at_the_correct_layer() {
        for encoded in ["_xD800_", "_xDC00_", "_xD800__x0041_", "&#x0;", "&#xFFFE;"] {
            let xml = schema_xml(&format!(
                r#"<questions><question binding="1" text="{encoded}"/></questions>"#
            ));
            assert!(parse(&xml).is_err(), "{encoded}");
        }
        assert!(
            parse(&schema_xml(
                r#"<questions><question binding="_x0031_"/></questions>"#
            ))
            .is_err()
        );
        assert!(parse(&schema_xml(r#"<questions><question binding="1" xmlns:q="urn:vendor" q:text="_xD800_"/></questions>"#)).is_ok());
        assert!(
            parse(&schema_xml(
                "<questions><question binding='1' text='\u{fffe}'/></questions>"
            ))
            .is_err()
        );
    }

    #[test]
    fn xstring_changes_save_reopen_and_restore_the_original_source() {
        let mut package = crate::package::Package::from_opc(package_with_survey(XML)).unwrap();
        let mut transaction = package.edit_surveys().unwrap();
        let title = "literal _x0041_\t\n\r\0😀";
        transaction
            .edit(0, |survey| {
                survey.set_title(Some(title.into()));
                survey.questions.values_mut()[0]
                    .set_row_source(Some("a\"b;\"unterminated".into()))?;
                Ok(())
            })
            .unwrap();
        let patch = transaction.commit().unwrap().patch().clone();
        let bytes = litchi_opc::PackageWriter::to_bytes(&package.into_plain_opc()).unwrap();
        let mut reopened =
            crate::package::Package::from_opc(OpcPackage::from_bytes(&bytes).unwrap()).unwrap();
        assert_eq!(
            reopened.surveys().unwrap().entries()[0].survey().title(),
            Some(title)
        );
        reopened.apply_surveys_patch(&patch.inverse()).unwrap();
        let bytes = litchi_opc::PackageWriter::to_bytes(&reopened.into_plain_opc()).unwrap();
        let restored = OpcPackage::from_bytes(&bytes).unwrap();
        assert_eq!(
            restored
                .get_part(&PackURI::new("/xl/surveys/survey1.xml").unwrap())
                .unwrap()
                .blob(),
            XML
        );
    }

    #[test]
    fn extension_lists_use_the_survey_namespace_at_every_owner() {
        let xml = schema_xml(
            r#"<surveyPr><extLst/></surveyPr><titlePr><extLst/></titlePr><descriptionPr><extLst/></descriptionPr><questions><questionsPr><extLst/></questionsPr><question binding="1"><questionPr><extLst/></questionPr><extLst/></question></questions><extLst/>"#,
        );
        let survey = parse(&xml).unwrap();
        assert_eq!(parse(&write(&survey).unwrap()).unwrap(), survey);
        for namespace in [
            "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
            "http://purl.oclc.org/ooxml/spreadsheetml/main",
            "urn:foreign",
            "",
        ] {
            let body = format!(
                r#"<questions><question binding="1"/></questions><extLst xmlns="{namespace}"/>"#
            );
            assert!(parse(&schema_xml(&body)).is_err(), "{namespace}");
        }
        let aliased = schema_xml(
            r#"<questions><question binding="1"/></questions><s:extLst xmlns:s="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"/>"#,
        );
        assert!(parse(&aliased).is_ok());
    }

    #[test]
    fn survey_children_follow_the_schema_sequence() {
        for body in [
            r#"<questions><question binding="1"/></questions><surveyPr/>"#,
            r#"<titlePr/><surveyPr/><questions><question binding="1"/></questions>"#,
            r#"<descriptionPr/><titlePr/><questions><question binding="1"/></questions>"#,
            r#"<extLst/><questions><question binding="1"/></questions>"#,
            r#"<questions><question binding="1"/><questionsPr/></questions>"#,
            r#"<questions><question binding="1"><extLst/><questionPr/></question></questions>"#,
            r#"<questions><question binding="1"/></questions><extLst/><titlePr/>"#,
            r#"<questions><question binding="1"/></questions><extLst/><extLst/>"#,
        ] {
            assert!(parse(&schema_xml(body)).is_err(), "{body}");
        }
    }

    #[test]
    fn qualified_attributes_cannot_supply_unqualified_protocol_fields() {
        assert!(
            parse(&schema_xml(
                r#"<questions><question xmlns:q="urn:vendor" q:binding="1"/></questions>"#
            ))
            .is_err()
        );
        let xml = schema_xml(
            r#"<surveyPr xmlns:q="urn:vendor" q:width="invalid"/><questions><question xmlns:q="urn:vendor" q:binding="999" binding="1" q:type="invalid"/></questions>"#,
        );
        let survey = parse(&xml).unwrap();
        assert_eq!(survey.questions.values[0].binding(), Binding::new(1));
        assert_eq!(survey.questions.values[0].kind, None);
        assert_eq!(survey.properties.as_ref().unwrap().width, None);
        let output = write(&survey).unwrap();
        assert!(String::from_utf8_lossy(&output).contains(r#"q:binding="999""#));
        assert_eq!(parse(&output).unwrap(), survey);
        let wrong_id = String::from_utf8(schema_xml(
            r#"<questions><question binding="1"/></questions>"#,
        ))
        .unwrap()
        .replace(r#"id="7""#, r#"xmlns:q="urn:vendor" q:id="7""#);
        assert!(parse(wrong_id.as_bytes()).is_err());
    }

    #[test]
    fn unknown_attribute_prefixes_and_expanded_duplicates_are_validated() {
        for body in [
            r#"<questions q:flag="1"><question binding="1"/></questions>"#,
            r#"<questions><question binding="1" q:flag="1"/></questions>"#,
            r#"<questions><question binding="1" xmlns:a="urn:v" xmlns:b="urn:v" a:flag="1" b:flag="2"/></questions>"#,
            r#"<questions><question binding="1" xmlns:a="urn:v" xmlns:b="urn:&#118;" a:flag="1" b:flag="2"/></questions>"#,
            r#"<questions><question binding="1"/></questions><extLst><q:unknown/></extLst>"#,
            r#"<questions><question binding="1"/></questions><extLst><unknown xmlns:a="urn:v" xmlns:b="urn:v" a:flag="1" b:flag="2"/></extLst>"#,
        ] {
            assert!(parse(&schema_xml(body)).is_err(), "{body}");
        }
    }

    #[test]
    fn question_container_namespace_bindings_survive_edit_save_and_inverse() {
        let xml = schema_xml(
            r#"<questions xmlns:q="urn:questions" q:flag="container"><question binding="1" q:flag="child"><questionPr q:flag="properties"/></question></questions>"#,
        );
        let mut package = crate::package::Package::from_opc(package_with_survey(&xml)).unwrap();
        let before = package.surveys().unwrap().entries()[0].survey().clone();
        let mut transaction = package.edit_surveys().unwrap();
        transaction
            .edit(0, |survey| {
                survey.set_title(Some("edited".into()));
                Ok(())
            })
            .unwrap();
        let patch = transaction.commit().unwrap().patch().clone();
        let bytes = litchi_opc::PackageWriter::to_bytes(&package.clone().into_plain_opc()).unwrap();
        let mut reopened =
            crate::package::Package::from_opc(OpcPackage::from_bytes(&bytes).unwrap()).unwrap();
        let after = reopened.surveys().unwrap().entries()[0].survey().clone();
        assert_eq!(after.questions, before.questions);
        assert_eq!(after.title(), Some("edited"));
        reopened.apply_surveys_patch(&patch.inverse()).unwrap();
        let bytes = litchi_opc::PackageWriter::to_bytes(&reopened.into_plain_opc()).unwrap();
        let restored = OpcPackage::from_bytes(&bytes).unwrap();
        assert_eq!(
            restored
                .get_part(&PackURI::new("/xl/surveys/survey1.xml").unwrap())
                .unwrap()
                .blob(),
            xml.as_slice()
        );
    }

    #[test]
    fn custom_question_limit_is_admitted_before_parsing_another_question() {
        for question in [
            r#"<question binding="invalid"/>"#,
            r#"<question binding="invalid"></question>"#,
        ] {
            let xml = schema_xml(&format!(
                r#"<questions><question binding="1"/>{question}</questions>"#
            ));
            assert!(
                matches!(parse_with_limits(&xml, &Limits::default().with_max_questions(1)),
                Err(Error::Invalid(message)) if message.contains("too many questions"))
            );
        }
    }

    #[test]
    fn extension_capture_preserves_source_tokens_and_precharges_aggregate_bytes() {
        let extension = "<extLst ><!-- one --><future a = '1'/></extLst>";
        let body = format!(
            r#"<surveyPr>{extension}</surveyPr><questions><question binding="1">{extension}</question></questions>"#
        );
        let xml = schema_xml(&body);
        let limits = Limits::default().with_max_extension_bytes(extension.len() * 2);
        let survey = parse_with_limits(&xml, &limits).unwrap();
        assert_eq!(
            survey
                .properties
                .as_ref()
                .unwrap()
                .extension_xml
                .as_deref()
                .unwrap(),
            extension.as_bytes()
        );
        assert_eq!(
            survey.questions.values[0].extension_xml.as_deref().unwrap(),
            extension.as_bytes()
        );
        assert!(
            parse_with_limits(
                &xml,
                &limits.with_max_extension_bytes(extension.len() * 2 - 1)
            )
            .is_err()
        );
        let empty = schema_xml(
            r#"<surveyPr><extLst/></surveyPr><questions><question binding="1"><extLst/></question></questions>"#,
        );
        assert!(parse_with_limits(&empty, &Limits::default().with_max_extension_bytes(18)).is_ok());
        assert!(
            parse_with_limits(&empty, &Limits::default().with_max_extension_bytes(17)).is_err()
        );
    }

    #[test]
    fn parses_complete_ct_survey_family() {
        let survey = parse(XML).expect("survey");
        assert_eq!(survey.id(), Id::new(7));
        assert_eq!(survey.title(), Some("Survey"));
        assert_eq!(
            survey.properties().and_then(ElementProperties::top),
            Some(-2)
        );
        let question = &survey.questions().values()[0];
        assert_eq!(question.binding(), Binding::new(1));
        assert_eq!(question.question_type(), Some(QuestionType::Choice));
        assert_eq!(question.decimal_places(), Some(15));
        assert!(question.is_required());
    }

    #[test]
    fn enforces_survey_constraints() {
        assert!(parse(br#"<survey xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" id="1" guid="{01234567-89ab-cdef-0123-456789abcdef}"><questions><question binding="1" decimalPlaces="16"/></questions></survey>"#).is_err());
        assert!(parse(br#"<survey xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" id="1" guid="{01234567-89ab-cdef-0123-456789abcdef}"><questions><question binding="1" rowSource="&quot;unterminated"/></questions></survey>"#).is_ok());
        assert!(parse(br#"<survey xmlns:survey="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" id="1" guid="{01234567-89ab-cdef-0123-456789abcdef}"><questions><question binding="1"/></questions></survey>"#).is_err());
        assert!(parse(br#"<survey xmlns="urn:not-survey" id="1" guid="{01234567-89ab-cdef-0123-456789abcdef}"><questions><question binding="1"/></questions></survey>"#).is_err());
    }

    #[test]
    fn accepts_the_survey_namespace_when_it_is_prefixed() {
        let survey = parse(br#"<s:survey xmlns:s="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" id="1" guid="{01234567-89ab-cdef-0123-456789abcdef}"><s:questions><s:question binding="1"/></s:questions></s:survey>"#).expect("prefixed survey");
        assert_eq!(survey.id(), Id::new(1));
    }

    #[test]
    fn prefixed_survey_retains_foreign_default_namespace_for_opaque_extensions() {
        let xml = format!(
            r#"<s:survey xmlns:s="{NAMESPACE_STR}" xmlns="urn:foreign" id="7" guid="{{01234567-89ab-cdef-0123-456789abcdef}}"><s:questions><s:question xmlns="urn:child-foreign" binding="1"><s:extLst><opaque><nested/></opaque></s:extLst></s:question></s:questions></s:survey>"#
        );
        let survey = parse(xml.as_bytes()).expect("foreign default namespace");
        let output = write(&survey).expect("write foreign default namespace");
        let text = String::from_utf8(output.clone()).expect("UTF-8 Survey XML");
        assert!(text.starts_with("<s:survey xmlns=\"urn:foreign\""));
        assert!(text.contains("<s:question xmlns=\"urn:child-foreign\""));
        assert!(text.contains("<opaque><nested/></opaque>"));
        let reopened = parse(&output).expect("reparse foreign default namespace");
        assert_eq!(reopened, survey);

        let mut detached = parse(xml.as_bytes()).expect("detached source survey");
        detached.set_id(Id::new(8));
        let mut package = package_with_survey(XML);
        let part_name = {
            let mut transaction = Transaction::new(&mut package).expect("transaction");
            let index = transaction
                .insert(PackURI::new("/xl/tables/table1.xml").unwrap(), detached)
                .expect("detached insert");
            let part_name = transaction.entries()[index].part_name().clone();
            transaction.commit().expect("detached commit");
            part_name
        };
        let inserted = package
            .get_part(&part_name)
            .expect("inserted Survey part")
            .blob();
        let inserted_text = String::from_utf8(inserted.to_vec()).expect("UTF-8 inserted Survey");
        assert!(inserted_text.starts_with("<s:survey xmlns=\"urn:foreign\""));
        assert!(inserted_text.contains("<s:question xmlns=\"urn:child-foreign\""));
        assert!(inserted_text.contains("<opaque><nested/></opaque>"));
        assert_eq!(
            parse(inserted).unwrap(),
            survey_with_id(&survey, Id::new(8))
        );
    }

    #[test]
    fn full_writer_preserves_comment_slots_across_semantic_edits_and_detached_insert() {
        let xml = format!(
            r#"<s:survey xmlns:s="{NAMESPACE_STR}" id="7" guid="{{01234567-89ab-cdef-0123-456789abcdef}}"><!--root-before--><s:surveyPr/><!--between-survey-and-title--><s:titlePr/><!--between-title-and-questions--><s:questions><!--before-questions-properties--><s:questionsPr><!--before-questions-properties-extension--><s:extLst><s:ext/></s:extLst><!--after-questions-properties-extension--></s:questionsPr><!--between-properties-and-first--><s:question binding="1"/><!--between-first-and-second--><s:question binding="1"/><!--after-second--></s:questions><!--between-questions-and-extension--><s:extLst><s:ext/></s:extLst><!--root-after-extension--></s:survey>"#
        );
        let survey = parse(xml.as_bytes()).expect("commented Survey");
        let output = write(&survey).expect("write commented Survey");
        assert_comment_order(
            &output,
            &[
                "root-before",
                "between-survey-and-title",
                "between-title-and-questions",
                "before-questions-properties",
                "before-questions-properties-extension",
                "after-questions-properties-extension",
                "between-properties-and-first",
                "between-first-and-second",
                "after-second",
                "between-questions-and-extension",
                "root-after-extension",
            ],
        );

        let mut edited = survey.clone();
        edited.set_title(Some("edited".into()));
        let edited_output = write(&edited).expect("write edited commented Survey");
        assert_comment_order(
            &edited_output,
            &[
                "root-before",
                "between-survey-and-title",
                "between-title-and-questions",
                "before-questions-properties",
                "before-questions-properties-extension",
                "after-questions-properties-extension",
                "between-properties-and-first",
                "between-first-and-second",
                "after-second",
                "between-questions-and-extension",
                "root-after-extension",
            ],
        );

        let mut package = package_with_survey(XML);
        edited.set_id(Id::new(8));
        let part_name = {
            let mut transaction = Transaction::new(&mut package).expect("transaction");
            let index = transaction
                .insert(PackURI::new("/xl/tables/table1.xml").unwrap(), edited)
                .expect("detached commented insert");
            let part_name = transaction.entries()[index].part_name().clone();
            transaction.commit().expect("detached commented commit");
            part_name
        };
        assert_comment_order(
            package.get_part(&part_name).unwrap().blob(),
            &[
                "root-before",
                "between-survey-and-title",
                "between-title-and-questions",
                "before-questions-properties",
                "before-questions-properties-extension",
                "after-questions-properties-extension",
                "between-properties-and-first",
                "between-first-and-second",
                "after-second",
                "between-questions-and-extension",
                "root-after-extension",
            ],
        );
    }

    #[test]
    fn nested_modeled_prefixes_preserve_foreign_default_context_everywhere() {
        let xml = format!(
            r#"<s:survey xmlns:s="{NAMESPACE_STR}" xmlns="urn:foreign" id="7" guid="{{01234567-89ab-cdef-0123-456789abcdef}}"><q:surveyPr xmlns:q="{NAMESPACE_STR}"><q:extLst><opaque><nested/></opaque></q:extLst></q:surveyPr><q:questions xmlns:q="{NAMESPACE_STR}"><q:questionsPr/><q:question binding="1"><q:questionPr><q:extLst><opaque/></q:extLst></q:questionPr><q:extLst><opaque><nested/></opaque></q:extLst></q:question></q:questions></s:survey>"#
        );
        let survey = parse(xml.as_bytes()).expect("nested modeled prefixes");
        let output = write(&survey).expect("standalone write");
        let output_text = String::from_utf8(output.clone()).expect("UTF-8 Survey XML");
        assert!(output_text.starts_with("<s:survey xmlns=\"urn:foreign\""));
        assert!(output_text.contains("<q:surveyPr xmlns:q=\""));
        assert!(output_text.contains("<q:questions xmlns:q=\""));
        assert!(output_text.contains("<q:question binding=\"1\""));
        assert!(output_text.contains("<opaque><nested/></opaque>"));
        assert_eq!(parse(&output).expect("reparse standalone output"), survey);

        let mut detached = survey.clone();
        detached.set_id(Id::new(8));
        detached
            .questions_mut()
            .insert(0, Question::new(Binding::new(1)))
            .expect("detached question");
        detached
            .properties
            .as_mut()
            .expect("surveyPr")
            .set_css_class(Some("detached".into()));
        let mut package = package_with_survey(XML);
        let part_name = {
            let mut transaction = Transaction::new(&mut package).expect("transaction");
            let index = transaction
                .insert(PackURI::new("/xl/tables/table1.xml").unwrap(), detached)
                .expect("detached insert");
            let part_name = transaction.entries()[index].part_name().clone();
            transaction.commit().expect("detached commit");
            part_name
        };
        let inserted = package
            .get_part(&part_name)
            .expect("inserted Survey")
            .blob();
        let inserted_text = String::from_utf8(inserted.to_vec()).expect("UTF-8 inserted Survey");
        assert!(inserted_text.starts_with("<s:survey xmlns=\"urn:foreign\""));
        assert!(inserted_text.contains("<q:surveyPr xmlns:q=\""));
        assert!(inserted_text.contains("cssClass=\"detached\""));
        assert!(inserted_text.contains("<q:questions xmlns:q=\""));
        assert!(inserted_text.contains("<question xmlns=\""));
        assert!(inserted_text.contains("<q:question binding=\"1\""));
        assert_eq!(parse(inserted).unwrap().id(), Id::new(8));

        let mut source_package =
            crate::package::Package::from_opc(as_source(&package_with_survey(xml.as_bytes())))
                .unwrap();
        let before = Snapshot::load(&source_package.clone().into_plain_opc()).unwrap();
        let mut transaction = source_package.edit_surveys().unwrap();
        transaction
            .edit(0, |survey| {
                survey
                    .properties
                    .as_mut()
                    .expect("surveyPr")
                    .set_css_class(Some("edited".into()));
                Ok(())
            })
            .unwrap();
        let patch = transaction.commit().unwrap().patch().clone();
        let bytes = litchi_opc::PackageWriter::to_bytes(&source_package.into_plain_opc()).unwrap();
        let reopened = OpcPackage::from_bytes(&bytes).unwrap();
        let name = PackURI::new("/xl/surveys/survey1.xml").unwrap();
        let changed = reopened.get_part(&name).unwrap().blob();
        let changed_text = String::from_utf8(changed.to_vec()).unwrap();
        assert!(changed_text.contains("xmlns=\"urn:foreign\""));
        assert!(changed_text.contains("<q:surveyPr xmlns:q=\""));
        assert!(changed_text.contains("cssClass=\"edited\""));
        assert!(changed_text.contains("<q:questions xmlns:q=\""));
        assert!(changed_text.contains("<q:question binding=\"1\""));
        let mut restored = reopened;
        patch.inverse().apply(&mut restored).unwrap();
        assert_eq!(
            restored.get_part(&name).unwrap().blob(),
            before.entries()[0].source_xml.as_slice()
        );
    }

    #[test]
    fn equality_tracks_opaque_namespace_context_and_comment_anchors() {
        let namespace_xml = |suffix: &str| {
            format!(
                r#"<s:survey xmlns:s="{NAMESPACE_STR}" xmlns="urn:root-{suffix}" id="7" guid="{{01234567-89ab-cdef-0123-456789abcdef}}"><s:questions xmlns="urn:questions-{suffix}"><s:questionsPr xmlns="urn:questions-properties-{suffix}"><s:extLst><opaque/></s:extLst></s:questionsPr><s:question xmlns="urn:question-{suffix}" binding="1"><s:questionPr xmlns="urn:question-properties-{suffix}"><s:extLst><opaque/></s:extLst></s:questionPr><s:extLst><opaque/></s:extLst></s:question></s:questions></s:survey>"#
            )
        };
        let first = parse(namespace_xml("a").as_bytes()).unwrap();
        let second = parse(namespace_xml("b").as_bytes()).unwrap();
        assert_ne!(first.questions, second.questions);
        assert_ne!(first.questions.values[0], second.questions.values[0]);
        assert_ne!(first.questions.properties, second.questions.properties);
        assert_ne!(
            first.questions.values[0].properties,
            second.questions.values[0].properties
        );

        let commented = parse(
            br#"<survey xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" id="7" guid="{01234567-89ab-cdef-0123-456789abcdef}"><!--root-before--><surveyPr><!--property--></surveyPr><!--root-after--><questions><!--questions--><questionsPr/><question binding="1"><!--question--><questionPr/></question></questions></survey>"#,
        )
        .unwrap();
        let mut root = commented.clone();
        root.comment_slots.swap(0, 1);
        assert_ne!(commented, root);
        let mut questions = commented.questions.clone();
        questions.comment_slots[0].slot = 1;
        assert_ne!(commented.questions, questions);
        let mut question = commented.questions.values[0].clone();
        question.comment_slots[0].slot = 1;
        assert_ne!(commented.questions.values[0], question);
        let mut properties = commented.properties.clone().unwrap();
        properties.comment_slots[0].slot = 1;
        assert_ne!(commented.properties, Some(properties));

        let mut package = as_source(&package_with_survey(namespace_xml("a").as_bytes()));
        let replacement = parse(namespace_xml("b").as_bytes()).unwrap();
        let mut transaction = Transaction::new(&mut package).unwrap();
        assert!(transaction.set(0, replacement.clone()).unwrap());
        assert!(transaction.is_changed());
        assert!(matches!(
            transaction.commit(),
            Err(Error::Unsupported {
                feature: "Survey question source transplantation"
            })
        ));
    }

    #[test]
    fn questions_property_presence_keeps_parent_comment_slots_single_and_ordered() {
        let xml = schema_xml(
            r#"<questions><!--before--><question binding="1"/><!--after--></questions>"#,
        );
        let mut survey = parse(&xml).unwrap();
        survey
            .questions_mut()
            .set_properties(Some(ElementProperties::new()));
        let output = write(&survey).unwrap();
        let text = String::from_utf8(output.clone()).unwrap();
        assert_eq!(text.matches("<!--before-->").count(), 1);
        assert_eq!(text.matches("<!--after-->").count(), 1);
        assert!(text.find("<!--before-->").unwrap() < text.find("<questionsPr").unwrap());
        assert!(text.find("<questionsPr").unwrap() < text.find(" binding=\"1\"").unwrap());
        assert_eq!(parse(&output).unwrap(), survey);

        let mut package =
            crate::package::Package::from_opc(as_source(&package_with_survey(&xml))).unwrap();
        let before = Snapshot::load(&package.clone().into_plain_opc()).unwrap();
        let mut transaction = package.edit_surveys().unwrap();
        transaction
            .edit(0, |survey| {
                survey
                    .questions_mut()
                    .set_properties(Some(ElementProperties::new()));
                Ok(())
            })
            .unwrap();
        let patch = transaction.commit().unwrap().patch().clone();
        let bytes = litchi_opc::PackageWriter::to_bytes(&package.into_plain_opc()).unwrap();
        let mut reopened = OpcPackage::from_bytes(&bytes).unwrap();
        let name = PackURI::new("/xl/surveys/survey1.xml").unwrap();
        let changed = reopened.get_part(&name).unwrap().blob();
        let changed_text = String::from_utf8(changed.to_vec()).unwrap();
        assert_eq!(changed_text.matches("<!--before-->").count(), 1);
        assert!(
            changed_text.find("<!--before-->").unwrap()
                < changed_text.find("<questionsPr").unwrap()
        );
        assert!(
            changed_text.find("<questionsPr").unwrap()
                < changed_text.find(" binding=\"1\"").unwrap()
        );
        patch.inverse().apply(&mut reopened).unwrap();
        assert_eq!(
            reopened.get_part(&name).unwrap().blob(),
            before.entries()[0].source_xml.as_slice()
        );

        let remove_xml = schema_xml(
            r#"<questions><!--before--><questionsPr/><!--between--><question binding="1"/><!--after--></questions>"#,
        );
        let mut removed = parse(&remove_xml).unwrap();
        removed.questions_mut().set_properties(None);
        let removed_output = write(&removed).unwrap();
        let removed_text = String::from_utf8(removed_output.clone()).unwrap();
        for marker in ["<!--before-->", "<!--between-->", "<!--after-->"] {
            assert_eq!(removed_text.matches(marker).count(), 1);
        }
        assert!(
            removed_text.find("<!--before-->").unwrap()
                < removed_text.find(" binding=\"1\"").unwrap()
        );
        assert!(
            removed_text.find("<!--between-->").unwrap()
                < removed_text.find(" binding=\"1\"").unwrap()
        );
        assert_eq!(parse(&removed_output).unwrap(), removed);

        let mut remove_package =
            crate::package::Package::from_opc(as_source(&package_with_survey(&remove_xml)))
                .unwrap();
        let remove_before = Snapshot::load(&remove_package.clone().into_plain_opc()).unwrap();
        let mut remove_transaction = remove_package.edit_surveys().unwrap();
        remove_transaction
            .edit(0, |survey| {
                survey.questions_mut().set_properties(None);
                Ok(())
            })
            .unwrap();
        let remove_patch = remove_transaction.commit().unwrap().patch().clone();
        let remove_bytes =
            litchi_opc::PackageWriter::to_bytes(&remove_package.into_plain_opc()).unwrap();
        let mut remove_reopened = OpcPackage::from_bytes(&remove_bytes).unwrap();
        let remove_name = PackURI::new("/xl/surveys/survey1.xml").unwrap();
        let remove_changed = remove_reopened.get_part(&remove_name).unwrap().blob();
        let remove_text = String::from_utf8(remove_changed.to_vec()).unwrap();
        for marker in ["<!--before-->", "<!--between-->", "<!--after-->"] {
            assert_eq!(remove_text.matches(marker).count(), 1);
        }
        assert!(
            remove_text.find("<!--before-->").unwrap()
                < remove_text.find(" binding=\"1\"").unwrap()
        );
        assert!(
            remove_text.find("<!--between-->").unwrap()
                < remove_text.find(" binding=\"1\"").unwrap()
        );
        remove_patch.inverse().apply(&mut remove_reopened).unwrap();
        assert_eq!(
            remove_reopened.get_part(&remove_name).unwrap().blob(),
            remove_before.entries()[0].source_xml.as_slice()
        );
    }

    #[test]
    fn question_insert_remove_preserve_parent_comment_slots_in_full_and_source_writers() {
        let xml = schema_xml(
            r#"<questions><!--slot0--><question binding="1"/><!--slot1--><question binding="1"/><!--slot2--></questions>"#,
        );
        let mut inserted = parse(&xml).unwrap();
        inserted
            .questions_mut()
            .insert(0, Question::new(Binding::new(1)))
            .unwrap();
        let inserted_output = write(&inserted).unwrap();
        assert_comment_order(&inserted_output, &["slot0", "slot1", "slot2"]);
        let inserted_text = String::from_utf8(inserted_output.clone()).unwrap();
        assert!(
            inserted_text.find("<!--slot0-->").unwrap()
                < inserted_text.find("<!--slot1-->").unwrap()
        );
        assert_eq!(parse(&inserted_output).unwrap(), inserted);

        let mut removed = parse(&xml).unwrap();
        removed.questions_mut().remove(0).unwrap();
        let removed_output = write(&removed).unwrap();
        assert_comment_order(&removed_output, &["slot0", "slot1", "slot2"]);
        assert_eq!(parse(&removed_output).unwrap(), removed);

        let mut package =
            crate::package::Package::from_opc(as_source(&package_with_survey(&xml))).unwrap();
        let before = Snapshot::load(&package.clone().into_plain_opc()).unwrap();
        let mut transaction = package.edit_surveys().unwrap();
        transaction
            .edit(0, |survey| {
                survey
                    .questions_mut()
                    .insert(0, Question::new(Binding::new(1)))?;
                Ok(())
            })
            .unwrap();
        let patch = transaction.commit().unwrap().patch().clone();
        let bytes = litchi_opc::PackageWriter::to_bytes(&package.into_plain_opc()).unwrap();
        let mut reopened = OpcPackage::from_bytes(&bytes).unwrap();
        let name = PackURI::new("/xl/surveys/survey1.xml").unwrap();
        let changed = reopened.get_part(&name).unwrap().blob();
        assert_comment_order(changed, &["slot0", "slot1", "slot2"]);
        patch.inverse().apply(&mut reopened).unwrap();
        assert_eq!(
            reopened.get_part(&name).unwrap().blob(),
            before.entries()[0].source_xml.as_slice()
        );

        let mut remove_package =
            crate::package::Package::from_opc(as_source(&package_with_survey(&xml))).unwrap();
        let remove_before = Snapshot::load(&remove_package.clone().into_plain_opc()).unwrap();
        let mut remove_transaction = remove_package.edit_surveys().unwrap();
        remove_transaction
            .edit(0, |survey| {
                survey.questions_mut().remove(0);
                Ok(())
            })
            .unwrap();
        let remove_patch = remove_transaction.commit().unwrap().patch().clone();
        let remove_bytes =
            litchi_opc::PackageWriter::to_bytes(&remove_package.into_plain_opc()).unwrap();
        let mut remove_reopened = OpcPackage::from_bytes(&remove_bytes).unwrap();
        let remove_changed = remove_reopened.get_part(&name).unwrap().blob();
        assert_comment_order(remove_changed, &["slot0", "slot1", "slot2"]);
        remove_patch.inverse().apply(&mut remove_reopened).unwrap();
        assert_eq!(
            remove_reopened.get_part(&name).unwrap().blob(),
            remove_before.entries()[0].source_xml.as_slice()
        );
    }

    #[test]
    fn relationship_part_name_aliases_are_equivalent_but_lexical_targets_remain_checked() {
        let mut package = package_with_survey(XML);
        let table_uri = PackURI::new("/xl/tables/table1.xml").unwrap();
        package
            .get_part_mut(&table_uri)
            .unwrap()
            .rels_mut()
            .remove("rIdSurvey");
        package
            .get_part_mut(&table_uri)
            .unwrap()
            .rels_mut()
            .try_add_relationship(
                RELATIONSHIP_TYPE.into(),
                "../SURVEYS/SURVEY1.XML".into(),
                "rIdSurvey".into(),
                TargetMode::Internal,
            )
            .unwrap();
        let snapshot = Snapshot::load(&package).expect("equivalent owner target");
        assert_eq!(snapshot.len(), 1);
        assert_eq!(
            snapshot.entries()[0].relationship_target(),
            "../SURVEYS/SURVEY1.XML"
        );

        package.rels_mut().add_relationship(
            "urn:foreign-survey-owner".into(),
            "xl/SURVEYS/SURVEY1.XML".into(),
            "rIdForeignSurvey".into(),
            false,
        );
        let mut transaction = Transaction::new(&mut package).expect("transaction");
        transaction.remove(0).expect("remove");
        assert!(transaction.commit().is_err());
        assert!(
            package
                .get_part(&PackURI::new("/xl/surveys/survey1.xml").unwrap())
                .is_ok()
        );
    }

    #[test]
    fn workbook_view_and_save_preserve_the_original_survey_part() {
        let mut package = crate::package::build_minimal_package().expect("minimal package");
        let table_uri = PackURI::new("/xl/tables/table1.xml").expect("table URI");
        let survey_uri = PackURI::new("/xl/surveys/survey1.xml").expect("survey URI");
        package
            .try_add_part(Box::new(BlobPart::new(
                table_uri.clone(),
                ct::SML_TABLE.into(),
                br#"<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="T" displayName="T" ref="A1:A2"><tableColumns count="1"><tableColumn id="1" name="Answer"/></tableColumns></table>"#.to_vec(),
            )))
            .expect("table part");
        package
            .try_add_part(Box::new(BlobPart::new(
                survey_uri.clone(),
                CONTENT_TYPE.into(),
                XML.to_vec(),
            )))
            .expect("survey part");
        let target = survey_uri.relative_ref(table_uri.base_uri());
        package
            .get_part_mut(&table_uri)
            .expect("table")
            .rels_mut()
            .try_add_relationship(
                RELATIONSHIP_TYPE.into(),
                target,
                "rIdSurvey".into(),
                TargetMode::Internal,
            )
            .expect("survey relationship");

        let workbook = crate::Workbook::from_package(package).expect("workbook");
        assert_eq!(workbook.surveys().expect("surveys").len(), 1);
        let bytes = workbook.to_bytes().expect("save");
        let reopened = crate::Workbook::from_bytes(bytes).expect("reopen");
        assert_eq!(
            reopened.surveys().expect("surveys")[0].survey().id(),
            Id::new(7)
        );
    }

    #[test]
    fn package_transaction_supports_edit_insert_remove_and_inverse() {
        let mut package = crate::package::Package::from_opc(package_with_survey(XML)).unwrap();
        let mut transaction = package.edit_surveys().expect("survey transaction");
        assert!(
            transaction
                .edit(0, |survey| {
                    survey.set_title(Some("Changed".to_owned()));
                    Ok(())
                })
                .expect("edit")
        );
        let inserted = transaction
            .insert(
                PackURI::new("/xl/tables/table1.xml").unwrap(),
                authored_survey(1),
            )
            .expect("insert");
        assert_eq!(inserted, 1);
        let commit = transaction.commit().expect("commit");
        assert!(commit.changed());
        assert_eq!(package.surveys().unwrap().len(), 2);

        let inverse = commit.patch().inverse();
        package
            .apply_surveys_patch(&inverse)
            .expect("inverse patch");
        assert_eq!(package.surveys().unwrap().len(), 1);
        assert_eq!(
            package.surveys().unwrap().entries()[0].survey().title(),
            Some("Survey")
        );

        let mut transaction = package.edit_surveys().expect("second transaction");
        assert!(transaction.remove(0).expect("remove").is_some());
        let commit = transaction.commit().expect("remove commit");
        assert!(commit.changed());
        assert!(package.surveys().unwrap().is_empty());
    }

    #[test]
    fn semantic_table_and_column_selectors_resolve_during_insert() {
        let mut package = crate::package::Package::from_opc(package_with_survey(XML)).unwrap();
        let survey = Survey::new(
            Id::new(8),
            Guid::new("{01234567-89ab-cdef-0123-456789abcdef}").unwrap(),
            vec![Question::for_column("Answer").unwrap()],
        )
        .unwrap();
        let mut transaction = package.edit_surveys().unwrap();
        let index = transaction.insert_for_table("T", survey).unwrap();
        assert_eq!(
            transaction.entries()[index].survey().questions().values()[0].binding(),
            Binding::new(1)
        );
        transaction.commit().unwrap();
        assert_eq!(package.surveys().unwrap().len(), 2);
    }

    #[test]
    fn no_op_transaction_preserves_exact_source_bytes() {
        let mut package = crate::package::Package::from_opc(package_with_survey(XML)).unwrap();
        let part_name = package.surveys().unwrap().entries()[0].part_name().clone();
        let before = package
            .clone()
            .into_plain_opc()
            .get_part(&part_name)
            .unwrap()
            .blob()
            .to_vec();
        let mut transaction = package.edit_surveys().unwrap();
        assert!(
            !transaction
                .edit(0, |survey| {
                    survey.set_title(Some("Survey".to_owned()));
                    Ok(())
                })
                .unwrap()
        );
        let commit = transaction.commit().unwrap();
        assert!(!commit.changed());
        assert!(commit.patch().is_empty());
        let after = package
            .clone()
            .into_plain_opc()
            .get_part(&part_name)
            .unwrap()
            .blob()
            .to_vec();
        assert_eq!(before, after);
    }

    #[test]
    fn inverse_patch_restores_survey_part_and_relationship_bytes_exactly() {
        let mut package = crate::package::Package::from_opc(package_with_survey(XML)).unwrap();
        let survey_part = PackURI::new("/xl/surveys/survey1.xml").unwrap();
        let table_part = PackURI::new("/xl/tables/table1.xml").unwrap();
        let before_survey = package
            .clone()
            .into_plain_opc()
            .get_part(&survey_part)
            .unwrap()
            .blob()
            .to_vec();
        let before_table_rels = package
            .clone()
            .into_plain_opc()
            .get_part(&table_part)
            .unwrap()
            .rels()
            .to_xml();
        let mut transaction = package.edit_surveys().unwrap();
        transaction
            .edit(0, |survey| {
                survey.set_description(Some("edited".to_owned()));
                Ok(())
            })
            .unwrap();
        let patch = transaction.commit().unwrap().patch().clone();
        package.apply_surveys_patch(&patch.inverse()).unwrap();
        assert_eq!(
            package
                .clone()
                .into_plain_opc()
                .get_part(&survey_part)
                .unwrap()
                .blob(),
            before_survey
        );
        assert_eq!(
            package
                .clone()
                .into_plain_opc()
                .get_part(&table_part)
                .unwrap()
                .rels()
                .to_xml(),
            before_table_rels
        );
    }

    #[test]
    fn graph_inverse_restores_noncanonical_content_types_and_relationship_members_after_reopen() {
        let mut package = noncanonical_graph_source(package_with_survey(XML));
        let table = PackURI::new("/xl/tables/table1.xml").unwrap();
        let survey = PackURI::new("/xl/surveys/survey1.xml").unwrap();
        let content_types = package.source_content_types().unwrap().bytes().to_vec();
        let table_relationships = package
            .source_relationships(&table)
            .unwrap()
            .bytes()
            .to_vec();
        let survey_relationships = package
            .source_relationships(&survey)
            .unwrap()
            .bytes()
            .to_vec();

        let mut transaction = Transaction::new(&mut package).unwrap();
        transaction.remove(0).unwrap();
        let patch = transaction.commit().unwrap().patch().clone();
        let removed = litchi_opc::PackageWriter::to_bytes(&package).unwrap();
        let mut reopened = OpcPackage::from_bytes(&removed).unwrap();
        assert!(Snapshot::load(&reopened).unwrap().is_empty());

        patch.inverse().apply(&mut reopened).unwrap();
        let restored = litchi_opc::PackageWriter::to_bytes(&reopened).unwrap();
        let restored = OpcPackage::from_bytes(&restored).unwrap();
        assert_eq!(
            restored.source_content_types().unwrap().bytes(),
            content_types.as_slice()
        );
        assert_eq!(
            restored.source_relationships(&table).unwrap().bytes(),
            table_relationships.as_slice()
        );
        assert_eq!(
            restored.source_relationships(&survey).unwrap().bytes(),
            survey_relationships.as_slice()
        );
        assert_eq!(Snapshot::load(&restored).unwrap().len(), 1);
    }

    #[test]
    fn graph_insert_inverse_restores_noncanonical_owner_members_after_reopen() {
        let mut package = noncanonical_graph_source(package_with_survey(XML));
        let table = PackURI::new("/xl/tables/table1.xml").unwrap();
        let survey = PackURI::new("/xl/surveys/survey1.xml").unwrap();
        let content_types = package.source_content_types().unwrap().bytes().to_vec();
        let table_relationships = package
            .source_relationships(&table)
            .unwrap()
            .bytes()
            .to_vec();
        let survey_relationships = package
            .source_relationships(&survey)
            .unwrap()
            .bytes()
            .to_vec();

        let mut transaction = Transaction::new(&mut package).unwrap();
        transaction
            .insert(table.clone(), authored_survey(8))
            .unwrap();
        let patch = transaction.commit().unwrap().patch().clone();
        let inserted = litchi_opc::PackageWriter::to_bytes(&package).unwrap();
        let mut reopened = OpcPackage::from_bytes(&inserted).unwrap();
        assert_eq!(Snapshot::load(&reopened).unwrap().len(), 2);

        patch.inverse().apply(&mut reopened).unwrap();
        let restored = litchi_opc::PackageWriter::to_bytes(&reopened).unwrap();
        let restored = OpcPackage::from_bytes(&restored).unwrap();
        assert_eq!(
            restored.source_content_types().unwrap().bytes(),
            content_types.as_slice()
        );
        assert_eq!(
            restored.source_relationships(&table).unwrap().bytes(),
            table_relationships.as_slice()
        );
        assert_eq!(
            restored.source_relationships(&survey).unwrap().bytes(),
            survey_relationships.as_slice()
        );
        assert_eq!(Snapshot::load(&restored).unwrap().len(), 1);
    }

    #[test]
    fn edits_retain_unknown_extension_xml_and_namespaces() {
        let xml = br#"<survey xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" xmlns:x="urn:survey-future" id="7" guid="{01234567-89ab-cdef-0123-456789abcdef}"><questions><question binding="1"/></questions><extLst><ext xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" uri="urn:future"><x:future x:flag="1"/></ext></extLst></survey>"#;
        let survey = parse(xml).expect("survey with extension");
        assert!(survey.extension_xml().unwrap().contains(&b'f'));
        let serialized = write(&survey).expect("write extension");
        assert!(
            serialized
                .windows(b"x:future".len())
                .any(|window| window == b"x:future")
        );

        let mut package = crate::package::Package::from_opc(package_with_survey(xml)).unwrap();
        let mut transaction = package.edit_surveys().unwrap();
        transaction
            .edit(0, |survey| {
                survey.set_description(Some("description".to_owned()));
                Ok(())
            })
            .unwrap();
        transaction.commit().unwrap();
        let part_name = package.surveys().unwrap().entries()[0].part_name().clone();
        let bytes = package
            .clone()
            .into_plain_opc()
            .get_part(&part_name)
            .unwrap()
            .blob()
            .to_vec();
        assert!(
            bytes
                .windows(b"x:future".len())
                .any(|window| window == b"x:future")
        );
        assert!(
            bytes
                .windows(b"urn:survey-future".len())
                .any(|window| window == b"urn:survey-future")
        );
    }

    #[test]
    fn edits_retain_nested_namespaces_comments_and_unknown_attributes() {
        let xml = br#"<survey xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" id="7" guid="{01234567-89ab-cdef-0123-456789abcdef}"><questions><!-- between --><question xmlns:q="urn:question-future" q:unknown="question" binding="1"><questionPr xmlns:p="urn:properties-future" p:unknown="properties"><!-- property comment --></questionPr><!-- question comment --></question></questions></survey>"#;
        let mut package = crate::package::Package::from_opc(package_with_survey(xml)).unwrap();
        let mut transaction = package.edit_surveys().unwrap();
        transaction
            .edit(0, |survey| {
                survey.set_description(Some("edited".to_owned()));
                Ok(())
            })
            .unwrap();
        transaction.commit().unwrap();
        let part_name = package.surveys().unwrap().entries()[0].part_name().clone();
        let bytes = package
            .clone()
            .into_plain_opc()
            .get_part(&part_name)
            .unwrap()
            .blob()
            .to_vec();
        for marker in [
            b"q:unknown".as_slice(),
            b"urn:question-future".as_slice(),
            b"p:unknown".as_slice(),
            b"urn:properties-future".as_slice(),
            b"property comment".as_slice(),
            b"question comment".as_slice(),
            b"between".as_slice(),
        ] {
            assert!(bytes.windows(marker.len()).any(|window| window == marker));
        }
    }

    #[test]
    fn failed_edit_does_not_mutate_the_transaction_draft() {
        let mut package = crate::package::Package::from_opc(package_with_survey(XML)).unwrap();
        let mut transaction = package.edit_surveys().unwrap();
        let result = transaction.edit(0, |survey| {
            survey.questions_mut().values_mut()[0].set_binding(Binding::new(999));
            Ok(())
        });
        assert!(result.is_err());
        assert_eq!(
            transaction.entries()[0].survey().questions().values()[0].binding(),
            Binding::new(1)
        );
    }

    #[test]
    fn stale_patch_and_explicit_limits_are_rejected_atomically() {
        let mut package = crate::package::Package::from_opc(package_with_survey(XML)).unwrap();
        let mut transaction = package.edit_surveys().unwrap();
        transaction
            .edit(0, |survey| {
                survey.set_title(Some("Changed".to_owned()));
                Ok(())
            })
            .unwrap();
        let patch = transaction.commit().unwrap().patch().clone();

        let mut stale = package_with_survey(XML);
        let survey_part = PackURI::new("/xl/surveys/survey1.xml").unwrap();
        stale.get_part_mut(&survey_part).unwrap().set_blob(
            String::from_utf8(XML.to_vec())
                .unwrap()
                .replace("title=\"Survey\"", "title=\"Stale\"")
                .into_bytes(),
        );
        let before = stale.get_part(&survey_part).unwrap().blob().to_vec();
        assert!(matches!(
            patch.apply(&mut stale),
            Err(Error::PatchConflict { .. })
        ));
        assert_eq!(stale.get_part(&survey_part).unwrap().blob(), before);

        let limits = Limits::new().with_max_part_bytes(XML.len() - 1);
        assert!(parse_with_limits(XML, &limits).is_err());
    }

    #[test]
    fn signature_policy_guards_changed_survey_publications_atomically() {
        let mut unchanged = signature_protected_source();
        let before = survey_package_state(&unchanged);
        let commit = Transaction::new(&mut unchanged).unwrap().commit().unwrap();
        assert!(!commit.changed());
        assert_eq!(survey_package_state(&unchanged), before);

        let mut scalar = signature_protected_source();
        let before = survey_package_state(&scalar);
        let mut transaction = Transaction::new(&mut scalar).unwrap();
        transaction
            .edit(0, |survey| {
                survey.set_title(Some("changed".into()));
                Ok(())
            })
            .unwrap();
        assert!(matches!(transaction.commit(), Err(Error::Signed)));
        assert_eq!(survey_package_state(&scalar), before);

        let mut added_question = signature_protected_source();
        let before = survey_package_state(&added_question);
        let mut transaction = Transaction::new(&mut added_question).unwrap();
        transaction
            .edit(0, |survey| {
                survey
                    .questions_mut()
                    .push(Question::new(Binding::new(1)))?;
                Ok(())
            })
            .unwrap();
        assert!(matches!(transaction.commit(), Err(Error::Signed)));
        assert_eq!(survey_package_state(&added_question), before);

        let mut removed_survey = signature_protected_source();
        let before = survey_package_state(&removed_survey);
        let mut transaction = Transaction::new(&mut removed_survey).unwrap();
        transaction.remove(0).unwrap();
        assert!(matches!(transaction.commit(), Err(Error::Signed)));
        assert_eq!(survey_package_state(&removed_survey), before);

        let mut inserted_survey = signature_protected_source();
        let before = survey_package_state(&inserted_survey);
        let mut transaction = Transaction::new(&mut inserted_survey).unwrap();
        transaction
            .insert(
                PackURI::new("/xl/tables/table1.xml").unwrap(),
                authored_survey(8),
            )
            .unwrap();
        assert!(matches!(transaction.commit(), Err(Error::Signed)));
        assert_eq!(survey_package_state(&inserted_survey), before);

        let protected_source = signature_protected_source();
        let mut unsigned = protected_source.clone();
        unsigned.unsign();
        let mut transaction = Transaction::new(&mut unsigned).unwrap();
        transaction
            .insert(
                PackURI::new("/xl/tables/table1.xml").unwrap(),
                authored_survey(8),
            )
            .unwrap();
        let patch = transaction.commit().unwrap().patch().clone();
        let mut protected = protected_source;
        let protected_before = Snapshot::load(&protected).unwrap();
        assert!(patch.before().same_source(&protected_before));
        let before = survey_package_state(&protected);
        assert!(matches!(patch.apply(&mut protected), Err(Error::Signed)));
        assert_eq!(survey_package_state(&protected), before);

        let mut explicitly_unsigned = signature_protected_source();
        explicitly_unsigned.unsign();
        assert!(!explicitly_unsigned.requires_signature_edit_policy());
        let mut transaction = Transaction::new(&mut explicitly_unsigned).unwrap();
        transaction
            .edit(0, |survey| {
                survey.set_title(Some("after unsign".into()));
                Ok(())
            })
            .unwrap();
        assert!(transaction.commit().is_ok());
    }

    #[test]
    fn foreign_package_relationship_keeps_survey_removal_graph_atomic() {
        let mut package = package_with_survey(XML);
        package.rels_mut().add_relationship(
            "urn:foreign-survey-owner".into(),
            "xl/surveys/survey1.xml".into(),
            "rIdForeignSurvey".into(),
            false,
        );
        let before = package
            .get_part(&PackURI::new("/xl/surveys/survey1.xml").unwrap())
            .unwrap()
            .blob()
            .to_vec();
        let mut transaction = Transaction::new(&mut package).unwrap();
        transaction.remove(0).unwrap();
        assert!(transaction.commit().is_err());
        assert_eq!(
            package
                .get_part(&PackURI::new("/xl/surveys/survey1.xml").unwrap())
                .unwrap()
                .blob(),
            before
        );
    }

    fn authored_survey(id: u32) -> Survey {
        Survey::new(
            Id::new(id),
            Guid::new("{01234567-89ab-cdef-0123-456789abcdef}").unwrap(),
            vec![Question::new(Binding::new(1))],
        )
        .unwrap()
    }

    fn survey_with_id(survey: &Survey, id: Id) -> Survey {
        let mut value = survey.clone();
        value.set_id(id);
        value
    }

    fn assert_comment_order(xml: &[u8], markers: &[&str]) {
        let text = String::from_utf8_lossy(xml);
        let mut previous = 0;
        for marker in markers {
            let needle = format!("<!--{marker}-->");
            let position = text
                .find(&needle)
                .unwrap_or_else(|| panic!("missing comment marker {marker}"));
            assert!(
                position >= previous,
                "comment marker {marker} moved before a prior slot"
            );
            previous = position;
        }
    }

    fn package_with_survey(xml: &[u8]) -> OpcPackage {
        let mut package = crate::package::build_minimal_package().unwrap();
        let table_uri = PackURI::new("/xl/tables/table1.xml").unwrap();
        let survey_uri = PackURI::new("/xl/surveys/survey1.xml").unwrap();
        package
            .try_add_part(Box::new(BlobPart::new(
                table_uri.clone(),
                ct::SML_TABLE.into(),
                br#"<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="T" displayName="T" ref="A1:A2"><tableColumns count="1"><tableColumn id="1" name="Answer"/></tableColumns></table>"#.to_vec(),
            )))
            .unwrap();
        package
            .try_add_part(Box::new(BlobPart::new(
                survey_uri.clone(),
                CONTENT_TYPE.into(),
                xml.to_vec(),
            )))
            .unwrap();
        package
            .get_part_mut(&table_uri)
            .unwrap()
            .rels_mut()
            .try_add_relationship(
                RELATIONSHIP_TYPE.into(),
                survey_uri.relative_ref(table_uri.base_uri()),
                "rIdSurvey".into(),
                TargetMode::Internal,
            )
            .unwrap();
        package
    }

    fn noncanonical_graph_source(package: OpcPackage) -> OpcPackage {
        let table = PackURI::new("/xl/tables/table1.xml").unwrap();
        let survey = PackURI::new("/xl/surveys/survey1.xml").unwrap();
        let table_target = survey.relative_ref(table.base_uri());
        let table_relationships = format!(
            r#"<?xml version="1.0"?><r:Relationships xmlns:r="http://schemas.openxmlformats.org/package/2006/relationships">
<!-- table-before --> <r:Relationship Target="{table_target}" Type="{RELATIONSHIP_TYPE}" Id="rIdSurvey"/>
<!-- table-after --> </r:Relationships>"#
        );
        let survey_relationships = br#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<!-- explicitly empty survey member -->
</Relationships>"#;
        let mut writer = soapberry_zip::office::StreamingArchiveWriter::new();
        let mut parts = package
            .try_iter_parts()
            .map(|part| part.expect("part payload decodes"))
            .collect::<Vec<_>>();
        parts.sort_by(|left, right| left.partname().as_str().cmp(right.partname().as_str()));
        let mut content_types = String::from(
            r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<!-- noncanonical manifest --> <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>"#,
        );
        for part in &parts {
            content_types.push_str(&format!(
                r#"<Override PartName="{}" ContentType="{}"/>"#,
                part.partname(),
                part.content_type()
            ));
            writer
                .write_stored(
                    part.partname().as_str().trim_start_matches('/'),
                    part.blob(),
                )
                .unwrap();
            if part.partname().is_equivalent_to(&table) {
                writer
                    .write_stored(
                        table.rels_uri().unwrap().as_str().trim_start_matches('/'),
                        table_relationships.as_bytes(),
                    )
                    .unwrap();
            } else if !part.rels().is_empty() {
                writer
                    .write_stored(
                        part.partname()
                            .rels_uri()
                            .unwrap()
                            .as_str()
                            .trim_start_matches('/'),
                        part.rels().to_xml().as_bytes(),
                    )
                    .unwrap();
            }
        }
        writer
            .write_stored(
                survey.rels_uri().unwrap().as_str().trim_start_matches('/'),
                survey_relationships,
            )
            .unwrap();
        content_types.push_str(&format!(
            r#"<Override PartName="{}" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="{}" ContentType="application/vnd.openxmlformats-package.relationships+xml"/></Types>"#,
            table.rels_uri().unwrap(),
            survey.rels_uri().unwrap()
        ));
        writer
            .write_stored("[Content_Types].xml", content_types.as_bytes())
            .unwrap();
        writer
            .write_stored("_rels/.rels", package.rels().to_xml().as_bytes())
            .unwrap();
        OpcPackage::from_bytes(&writer.finish_to_bytes().unwrap()).unwrap()
    }

    fn signature_protected_source() -> OpcPackage {
        let mut package = package_with_survey(XML);
        let origin = PackURI::new("/_xmlsignatures/origin.sigs").unwrap();
        package
            .try_add_part(Box::new(BlobPart::new(
                origin,
                ct::OPC_DIGITAL_SIGNATURE_ORIGIN.into(),
                Vec::new(),
            )))
            .unwrap();
        package.relate_to(
            "_xmlsignatures/origin.sigs",
            litchi_opc::constants::relationship_type::DIGITAL_SIGNATURE_ORIGIN,
        );
        let mut package = as_source(&package);
        let origin = PackURI::new("/_xmlsignatures/origin.sigs").unwrap();
        let signature_relationship = package
            .rels()
            .iter()
            .find(|relationship| {
                relationship.reltype()
                    == litchi_opc::constants::relationship_type::DIGITAL_SIGNATURE_ORIGIN
            })
            .unwrap()
            .r_id()
            .to_owned();
        package.rels_mut().remove(&signature_relationship);
        assert!(package.remove_part(&origin));
        assert!(!package.is_signed());
        assert!(package.requires_signature_edit_policy());
        package
    }

    fn survey_package_state(package: &OpcPackage) -> (Vec<u8>, String, String) {
        let survey = PackURI::new("/xl/surveys/survey1.xml").unwrap();
        let table = PackURI::new("/xl/tables/table1.xml").unwrap();
        (
            package.get_part(&survey).unwrap().blob().to_vec(),
            package.get_part(&table).unwrap().rels().to_xml(),
            package.rels().to_xml(),
        )
    }
}
