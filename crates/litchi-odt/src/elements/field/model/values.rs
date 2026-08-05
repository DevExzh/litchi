//! Reusable typed field values and display options.

#![allow(
    clippy::wildcard_imports,
    reason = "semantic field owners share the stable model facade namespace"
)]
use super::*;
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

    pub(super) fn parse(value: &str) -> Result<Self> {
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

    pub(super) fn parse(value: &str) -> Result<Self> {
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

    pub(super) fn validate(&self, aggregate: &mut usize) -> Result<()> {
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

    pub(super) fn validate(&self, aggregate: &mut usize) -> Result<()> {
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

    pub(super) fn validate(&self, name: &str, aggregate: &mut usize) -> Result<()> {
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
    pub(super) const fn as_str(self) -> &'static str {
        match self {
            Self::Value => "value",
            Self::None => "none",
        }
    }

    pub(super) fn parse(value: &str) -> Result<Self> {
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
    pub(super) const fn as_str(self) -> &'static str {
        match self {
            Self::Value => "value",
            Self::Formula => "formula",
        }
    }

    pub(super) fn parse(value: &str) -> Result<Self> {
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
    pub(super) const fn as_str(self) -> &'static str {
        match self {
            Self::Full => "full",
            Self::Path => "path",
            Self::Name => "name",
            Self::NameAndExtension => "name-and-extension",
        }
    }

    pub(super) fn parse(value: &str) -> Result<Self> {
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
    pub(super) const fn as_str(self) -> &'static str {
        match self {
            Self::Area => "area",
            Self::Full => "full",
            Self::Name => "name",
            Self::NameAndExtension => "name-and-extension",
            Self::Path => "path",
            Self::Title => "title",
        }
    }

    pub(super) fn parse(value: &str) -> Result<Self> {
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
    pub(super) const fn as_str(self) -> &'static str {
        match self {
            Self::Name => "name",
            Self::Number => "number",
            Self::NumberAndName => "number-and-name",
            Self::PlainNumber => "plain-number",
            Self::PlainNumberAndName => "plain-number-and-name",
        }
    }

    pub(super) fn parse(value: &str) -> Result<Self> {
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

    pub(super) fn parse(value: &str) -> Result<Self> {
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

impl CalculatedFieldValue {
    pub(super) fn validate(&self, aggregate: &mut usize) -> Result<()> {
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

    pub(super) fn write_attributes(&self, element: &mut Element) {
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

    pub(super) fn parse(value: &str) -> Result<Self> {
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

    pub(super) fn validate(&self) -> Result<()> {
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
    pub(super) fn parse(value: &str) -> Result<Self> {
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
