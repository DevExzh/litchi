//! Semantic typed object model for ODF spreadsheet content validations.

use litchi_core::{Error, Result};
use std::collections::HashSet;

pub(super) const MAX_VALIDATIONS: usize = 65_536;
pub(super) const MAX_PARAGRAPHS: usize = 262_144;
pub(super) const MAX_VALUE_BYTES: usize = 65_536;
pub(super) const MAX_CAPTURE_BYTES: usize = 1_048_576;
const MAX_AGGREGATE_BYTES: usize = 16 * 1_048_576;

#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum ContentValidationPart {
    Content,
    FlatDocument,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum ValidationDisplayList {
    None,
    Unsorted,
    SortAscending,
}

impl ValidationDisplayList {
    pub(super) fn parse(value: &str) -> Result<Self> {
        match value {
            "none" => Ok(Self::None),
            "unsorted" => Ok(Self::Unsorted),
            "sort-ascending" => Ok(Self::SortAscending),
            _ => invalid(format!("unsupported table:display-list value '{value}'")),
        }
    }
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::None => "none",
            Self::Unsorted => "unsorted",
            Self::SortAscending => "sort-ascending",
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum ValidationMessageType {
    Stop,
    Warning,
    Information,
}

impl ValidationMessageType {
    pub(super) fn parse(value: &str) -> Result<Self> {
        match value {
            "stop" => Ok(Self::Stop),
            "warning" => Ok(Self::Warning),
            "information" => Ok(Self::Information),
            _ => invalid(format!("unsupported table:message-type value '{value}'")),
        }
    }
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Stop => "stop",
            Self::Warning => "warning",
            Self::Information => "information",
        }
    }
}

macro_rules! lexical_type {
    ($name:ident, $label:literal, $allow_empty:expr) => {
        #[derive(Clone, Debug, PartialEq, Eq, Hash)]
        pub struct $name(String);
        impl $name {
            pub fn new(value: impl Into<String>) -> Result<Self> {
                let value = value.into();
                validate_text(&value, $label, $allow_empty)?;
                Ok(Self(value))
            }
            pub fn as_str(&self) -> &str {
                &self.0
            }
        }
    };
}

lexical_type!(ValidationCondition, "table:condition", true);
lexical_type!(ValidationCellAddress, "table:base-cell-address", false);

/// One message paragraph with its inert XML and flattened text.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ValidationParagraph {
    pub(super) xml: String,
    pub(super) text: String,
}

impl ValidationParagraph {
    pub fn from_text(value: impl Into<String>) -> Result<Self> {
        let text = value.into();
        validate_text(&text, "validation paragraph", true)?;
        let mut xml = String::with_capacity(text.len() + 32);
        xml.push_str("<text:p>");
        escape_text(&mut xml, &text);
        xml.push_str("</text:p>");
        Ok(Self { xml, text })
    }
    pub fn text(&self) -> &str {
        &self.text
    }
    pub fn as_xml(&self) -> &str {
        &self.xml
    }
}

/// A help or ordinary error message. `message_type` is valid only for errors.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct ValidationMessage {
    pub title: Option<String>,
    pub display: Option<bool>,
    pub message_type: Option<ValidationMessageType>,
    pub paragraphs: Vec<ValidationParagraph>,
}

impl ValidationMessage {
    pub fn effective_display(&self) -> bool {
        self.display.unwrap_or(false)
    }
    pub fn effective_message_type(&self) -> ValidationMessageType {
        self.message_type.unwrap_or(ValidationMessageType::Stop)
    }
}

/// Bounded `office:event-listeners` XML retained without dispatching it.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ValidationEventListeners(pub(super) String);

impl ValidationEventListeners {
    pub fn as_xml(&self) -> &str {
        &self.0
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum ValidationFailure {
    Message(ValidationMessage),
    Macro {
        execute: Option<bool>,
        event_listeners: Option<ValidationEventListeners>,
    },
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ContentValidation {
    pub name: String,
    pub condition: Option<ValidationCondition>,
    pub base_cell_address: Option<ValidationCellAddress>,
    pub allow_empty_cell: Option<bool>,
    pub display_list: Option<ValidationDisplayList>,
    pub help_message: Option<ValidationMessage>,
    pub failure: Option<ValidationFailure>,
}

impl ContentValidation {
    pub fn effective_allow_empty_cell(&self) -> bool {
        self.allow_empty_cell.unwrap_or(true)
    }
    pub fn effective_display_list(&self) -> ValidationDisplayList {
        self.display_list.unwrap_or(ValidationDisplayList::Unsorted)
    }
    fn validate(&self) -> Result<()> {
        validate_text(&self.name, "table:name", false)?;
        if let Some(value) = &self.condition {
            validate_text(value.as_str(), "table:condition", true)?;
        }
        if let Some(value) = &self.base_cell_address {
            validate_text(value.as_str(), "table:base-cell-address", false)?;
        }
        if let Some(help) = &self.help_message {
            if help.message_type.is_some() {
                return invalid("help messages cannot have table:message-type");
            }
            validate_message(help, "help message")?;
        }
        if let Some(failure) = &self.failure {
            match failure {
                ValidationFailure::Message(message) => validate_message(message, "error message")?,
                ValidationFailure::Macro {
                    event_listeners, ..
                } => {
                    if let Some(value) = event_listeners
                        && value.0.len() > MAX_CAPTURE_BYTES
                    {
                        return invalid("validation event listeners exceed 1 MiB");
                    }
                },
            }
        }
        Ok(())
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ContentValidations {
    pub part: ContentValidationPart,
    pub validations: Vec<ContentValidation>,
}

impl ContentValidations {
    pub fn get(&self, name: &str) -> Option<&ContentValidation> {
        self.validations.iter().find(|value| value.name == name)
    }
    pub fn validate(&self) -> Result<()> {
        if self.validations.len() > MAX_VALIDATIONS {
            return invalid(format!(
                "document exceeds {MAX_VALIDATIONS} content validations"
            ));
        }
        let mut names = HashSet::with_capacity(self.validations.len());
        let mut paragraphs = 0usize;
        let mut aggregate = 0usize;
        for validation in &self.validations {
            validation.validate()?;
            if !names.insert(validation.name.as_str()) {
                return invalid(format!(
                    "duplicate content-validation name '{}'",
                    validation.name
                ));
            }
            aggregate = aggregate
                .checked_add(validation.name.len())
                .ok_or_else(|| make_error("content-validation size overflow"))?;
            if let Some(value) = &validation.condition {
                aggregate = aggregate
                    .checked_add(value.as_str().len())
                    .ok_or_else(|| make_error("content-validation size overflow"))?;
            }
            if let Some(value) = &validation.base_cell_address {
                aggregate = aggregate
                    .checked_add(value.as_str().len())
                    .ok_or_else(|| make_error("content-validation size overflow"))?;
            }
            for message in validation
                .help_message
                .iter()
                .chain(match &validation.failure {
                    Some(ValidationFailure::Message(value)) => Some(value),
                    _ => None,
                })
            {
                paragraphs = paragraphs
                    .checked_add(message.paragraphs.len())
                    .ok_or_else(|| make_error("validation paragraph count overflow"))?;
                if paragraphs > MAX_PARAGRAPHS {
                    return invalid(format!(
                        "document exceeds {MAX_PARAGRAPHS} validation paragraphs"
                    ));
                }
                aggregate = aggregate
                    .checked_add(message.title.as_ref().map_or(0, String::len))
                    .ok_or_else(|| make_error("content-validation size overflow"))?;
                for paragraph in &message.paragraphs {
                    aggregate = aggregate
                        .checked_add(paragraph.xml.len())
                        .ok_or_else(|| make_error("content-validation size overflow"))?;
                }
            }
            if let Some(ValidationFailure::Macro {
                event_listeners: Some(value),
                ..
            }) = &validation.failure
            {
                aggregate = aggregate
                    .checked_add(value.0.len())
                    .ok_or_else(|| make_error("content-validation size overflow"))?;
            }
            if aggregate > MAX_AGGREGATE_BYTES {
                return invalid("content-validation metadata exceeds 16 MiB");
            }
        }
        Ok(())
    }
}

fn validate_message(message: &ValidationMessage, context: &str) -> Result<()> {
    if let Some(value) = &message.title {
        validate_text(value, "table:title", true)?;
    }
    if message.paragraphs.len() > MAX_PARAGRAPHS {
        return invalid(format!("{context} exceeds {MAX_PARAGRAPHS} paragraphs"));
    }
    for paragraph in &message.paragraphs {
        if paragraph.xml.len() > MAX_CAPTURE_BYTES {
            return invalid(format!("{context} paragraph exceeds 1 MiB"));
        }
        validate_text(&paragraph.text, "validation paragraph text", true)?;
    }
    Ok(())
}

pub(super) fn make_error(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}
pub(super) fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(make_error(message))
}

pub(super) fn validate_text(value: &str, name: &str, allow_empty: bool) -> Result<()> {
    if !allow_empty && value.is_empty() {
        return invalid(format!("{name} cannot be empty"));
    }
    if value.len() > MAX_VALUE_BYTES {
        return invalid(format!("{name} exceeds {MAX_VALUE_BYTES} bytes"));
    }
    if value.chars().any(
        |character| matches!(character, '\0'..='\u{8}' | '\u{b}' | '\u{c}' | '\u{e}'..='\u{1f}'),
    ) {
        return invalid(format!("{name} contains invalid XML characters"));
    }
    Ok(())
}

fn escape_text(output: &mut String, value: &str) {
    for character in value.chars() {
        match character {
            '&' => output.push_str("&amp;"),
            '<' => output.push_str("&lt;"),
            '>' => output.push_str("&gt;"),
            '\r' => output.push_str("&#13;"),
            _ => output.push(character),
        }
    }
}
