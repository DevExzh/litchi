//! Typed, inert `[MS-XLSX]` Survey-part inspection.
//!
//! Survey parts are associated with `SpreadsheetML` tables. They are parsed only
//! when requested and are never rendered, submitted, or otherwise activated.
//! The owning OPC package retains the original part bytes, so an unchanged
//! workbook save preserves survey XML and relationships losslessly.

use std::collections::HashSet;

use litchi_ooxml_common::custom_xml::valid_guid;
use litchi_opc::OpcPackage;
use litchi_opc::constants::content_type as ct;
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;

use crate::error::{Result, invalid};

/// The `[MS-XLSX]` Survey-part content type.
pub const CONTENT_TYPE: &str = "application/vnd.ms-excel.Survey+xml";
/// The table-to-survey relationship type specified by `[MS-XLSX]` §2.1.9.
pub const RELATIONSHIP_TYPE: &str = "http://schemas.microsoft.com/office/2010/relationships/Survey";

const NAMESPACE: &[u8] = b"http://schemas.microsoft.com/office/spreadsheetml/2010/11/main";
const TRANSITIONAL_MAIN_NAMESPACE: &[u8] =
    b"http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT_MAIN_NAMESPACE: &[u8] = b"http://purl.oclc.org/ooxml/spreadsheetml/main";
const MAX_XML_BYTES: usize = 16 * 1024 * 1024;
const MAX_TEXT_BYTES: usize = 1024 * 1024;
const MAX_QUESTIONS: usize = 65_534;
const MAX_DEPTH: usize = 256;

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
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ElementProperties {
    css_class: Option<Box<str>>,
    bottom: Option<i32>,
    top: Option<i32>,
    left: Option<i32>,
    right: Option<i32>,
    width: Option<u32>,
    height: Option<u32>,
    position: Option<Position>,
}

impl ElementProperties {
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
}

/// One table-bound survey question.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Question {
    binding: Binding,
    text: Option<Box<str>>,
    kind: Option<QuestionType>,
    format: Option<QuestionFormat>,
    help_text: Option<Box<str>>,
    required: bool,
    default_value: Option<Box<str>>,
    decimal_places: Option<u8>,
    row_source: Option<Box<str>>,
    properties: Option<ElementProperties>,
}

impl Question {
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
}

/// The ordered survey-question collection.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Questions {
    properties: Option<ElementProperties>,
    values: Box<[Question]>,
}

impl Questions {
    #[must_use]
    pub fn properties(&self) -> Option<&ElementProperties> {
        self.properties.as_ref()
    }
    #[must_use]
    pub fn values(&self) -> &[Question] {
        &self.values
    }
}

/// A parsed survey part. It is intentionally read-only and inert.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Survey {
    id: Id,
    guid: Guid,
    title: Option<Box<str>>,
    description: Option<Box<str>>,
    properties: Option<ElementProperties>,
    title_properties: Option<ElementProperties>,
    description_properties: Option<ElementProperties>,
    questions: Questions,
}

impl Survey {
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

/// One survey attached to a table in an immutable workbook snapshot.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Part {
    survey: Survey,
}

impl Part {
    /// Typed, inert survey contents. The associated table remains unchanged.
    #[must_use]
    pub fn survey(&self) -> &Survey {
        &self.survey
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

#[derive(Default)]
struct Parser {
    scopes: Vec<Scope>,
    survey: Option<SurveyBuilder>,
    seen_root: bool,
    closed_root: bool,
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
    questions_seen: bool,
}

impl Parser {
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
                set_once(
                    &mut self.survey_mut()?.properties,
                    parse_properties(element)?,
                    "surveyPr",
                )?;
                Ok(Scope::Properties(b"surveyPr"))
            },
            (Scope::Survey, b"titlePr") => {
                set_once(
                    &mut self.survey_mut()?.title_properties,
                    parse_properties(element)?,
                    "titlePr",
                )?;
                Ok(Scope::Properties(b"titlePr"))
            },
            (Scope::Survey, b"descriptionPr") => {
                set_once(
                    &mut self.survey_mut()?.description_properties,
                    parse_properties(element)?,
                    "descriptionPr",
                )?;
                Ok(Scope::Properties(b"descriptionPr"))
            },
            (Scope::Survey, b"questions") => {
                let survey = self.survey_mut()?;
                if survey.questions_seen {
                    return Err(invalid("survey has multiple questions elements"));
                }
                survey.questions_seen = true;
                Ok(Scope::Questions)
            },
            (Scope::Questions, b"questionsPr") => {
                set_once(
                    &mut self.survey_mut()?.questions_properties,
                    parse_properties(element)?,
                    "questionsPr",
                )?;
                Ok(Scope::Properties(b"questionsPr"))
            },
            (Scope::Questions, b"question") => {
                let survey = self.survey_mut()?;
                if survey.questions.len() >= MAX_QUESTIONS {
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
        let value = self
            .survey
            .ok_or_else(|| invalid("survey root is missing"))?;
        let id = value.id.ok_or_else(|| invalid("survey id is required"))?;
        let guid = value
            .guid
            .ok_or_else(|| invalid("survey guid is required"))?;
        if !value.questions_seen || value.questions.is_empty() {
            return Err(invalid("survey requires at least one question"));
        }
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
                values: value.questions.into_boxed_slice(),
            },
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
    if xml.len() > MAX_XML_BYTES {
        return Err(invalid("survey XML exceeds the size limit"));
    }
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut parser = Parser::default();
    loop {
        let (namespace, event) = reader.read_resolved_event().map_err(xml_error)?;
        match event {
            Event::Start(element) => parser.start(&namespace, &element)?,
            Event::Empty(element) => parser.empty(&namespace, &element)?,
            Event::End(element) => parser.end(&namespace, element.local_name().as_ref())?,
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
            Event::Decl(_) | Event::Comment(_) => {},
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
    parser.finish()
}

/// Load every table-owned survey part without changing the package graph.
///
/// # Errors
///
/// Returns an error when a survey part has an invalid table relationship,
/// duplicate survey ID, or malformed Survey XML.
pub fn load(package: &OpcPackage) -> Result<Vec<Part>> {
    let mut values = Vec::new();
    let mut ids = HashSet::new();
    for part in package
        .iter_parts()
        .filter(|part| part.content_type() == CONTENT_TYPE)
    {
        let sources = package
            .iter_parts()
            .filter(|source| source.content_type() == ct::SML_TABLE)
            .filter(|source| {
                source.rels().iter().any(|relationship| {
                    !relationship.is_external()
                        && relationship.reltype() == RELATIONSHIP_TYPE
                        && relationship.target_partname().ok().as_ref() == Some(part.partname())
                })
            })
            .count();
        if sources != 1 {
            return Err(invalid(format!(
                "survey part '{}' must have exactly one table source relationship",
                part.partname()
            )));
        }
        let survey = parse(part.blob())?;
        if !ids.insert(survey.id()) {
            return Err(invalid("survey IDs must be unique within a workbook"));
        }
        values.push(Part { survey });
    }
    values.sort_unstable_by_key(|part| part.survey.id());
    Ok(values)
}

fn set_once<T>(slot: &mut Option<T>, value: T, name: &str) -> Result<()> {
    if slot.replace(value).is_some() {
        Err(invalid(format!("duplicate survey {name}")))
    } else {
        Ok(())
    }
}
fn xml_error(error: impl std::fmt::Display) -> crate::Error {
    invalid(format!("invalid survey XML: {error}"))
}
fn text(value: &str, field: &str) -> Result<Box<str>> {
    if value.len() > MAX_TEXT_BYTES {
        Err(invalid(format!("survey {field} exceeds the size limit")))
    } else {
        Ok(value.into())
    }
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
            Ok((attribute.key.local_name().as_ref().to_vec(), decoded_value))
        })
        .collect()
}
fn attr<'a>(values: &'a [(Vec<u8>, String)], name: &[u8]) -> Option<&'a str> {
    values
        .iter()
        .find(|(key, _)| key.as_slice() == name)
        .map(|(_, value)| value.as_str())
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
    matches!(namespace, ResolveResult::Bound(value) if {
        matches!(value.as_ref(), TRANSITIONAL_MAIN_NAMESPACE | STRICT_MAIN_NAMESPACE)
    })
}

fn parse_survey(element: &BytesStart<'_>) -> Result<SurveyBuilder> {
    let values = attrs(element)?;
    Ok(SurveyBuilder {
        id: Some(Id::new(parse_u32(required(&values, b"id", "id")?, "id")?)),
        guid: Some(Guid::new(text(
            required(&values, b"guid", "guid")?,
            "guid",
        )?)?),
        title: attr(&values, b"title")
            .map(|value| text(value, "title"))
            .transpose()?,
        description: attr(&values, b"description")
            .map(|value| text(value, "description"))
            .transpose()?,
        ..SurveyBuilder::default()
    })
}
fn parse_properties(element: &BytesStart<'_>) -> Result<ElementProperties> {
    let values = attrs(element)?;
    Ok(ElementProperties {
        css_class: attr(&values, b"cssClass")
            .map(|value| text(value, "cssClass"))
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
    })
}
fn parse_question(element: &BytesStart<'_>) -> Result<Question> {
    let values = attrs(element)?;
    let row_source = attr(&values, b"rowSource")
        .map(|value| {
            validate_row_source(value)?;
            text(value, "rowSource")
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
        binding: Binding::new(parse_u32(
            required(&values, b"binding", "question binding")?,
            "question binding",
        )?),
        text: attr(&values, b"text")
            .map(|value| text(value, "question text"))
            .transpose()?,
        kind: attr(&values, b"type")
            .map(QuestionType::parse)
            .transpose()?,
        format: attr(&values, b"format")
            .map(QuestionFormat::parse)
            .transpose()?,
        help_text: attr(&values, b"helpText")
            .map(|value| text(value, "helpText"))
            .transpose()?,
        required: attr(&values, b"required")
            .map(parse_bool)
            .transpose()?
            .unwrap_or(false),
        default_value: attr(&values, b"defaultValue")
            .map(|value| text(value, "defaultValue"))
            .transpose()?,
        decimal_places,
        row_source,
        properties: None,
    })
}
fn validate_row_source(value: &str) -> Result<()> {
    let bytes = value.as_bytes();
    let mut position = 0;
    while position < bytes.len() {
        if bytes[position] == b'"' {
            position += 1;
            while position < bytes.len() && bytes[position] != b'"' {
                position += 1;
            }
            if position == bytes.len() {
                return Err(invalid("unterminated survey rowSource quote"));
            }
            position += 1;
            if position < bytes.len() && bytes[position] != b';' {
                return Err(invalid("invalid survey rowSource quoted value"));
            }
        } else {
            while position < bytes.len() && bytes[position] != b';' {
                if bytes[position] == b'"' {
                    return Err(invalid("invalid survey rowSource quote"));
                }
                position += 1;
            }
        }
        if position < bytes.len() {
            position += 1;
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
        assert!(parse(br#"<survey xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" id="1" guid="{01234567-89ab-cdef-0123-456789abcdef}"><questions><question binding="1" rowSource="&quot;unterminated"/></questions></survey>"#).is_err());
        assert!(parse(br#"<survey xmlns:survey="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" id="1" guid="{01234567-89ab-cdef-0123-456789abcdef}"><questions><question binding="1"/></questions></survey>"#).is_err());
        assert!(parse(br#"<survey xmlns="urn:not-survey" id="1" guid="{01234567-89ab-cdef-0123-456789abcdef}"><questions><question binding="1"/></questions></survey>"#).is_err());
    }

    #[test]
    fn accepts_the_survey_namespace_when_it_is_prefixed() {
        let survey = parse(br#"<s:survey xmlns:s="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" id="1" guid="{01234567-89ab-cdef-0123-456789abcdef}"><s:questions><s:question binding="1"/></s:questions></s:survey>"#).expect("prefixed survey");
        assert_eq!(survey.id(), Id::new(1));
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
}
