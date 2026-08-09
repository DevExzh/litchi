//! ODF spreadsheet content-validation definitions.

use litchi_core::{Error, Result, xml::escape_xml};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{
    Namespace as XmlNamespace, NamespaceResolver, PrefixDeclaration, ResolveResult,
};
use quick_xml::reader::NsReader;
use std::collections::HashSet;

use super::names::formula;

const TABLE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const TEXT_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const OFFICE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const SCRIPT_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:script:1.0";
const PRESENTATION_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:presentation:1.0";
const XLINK_NAMESPACE: &[u8] = b"http://www.w3.org/1999/xlink";
const XML_NAMESPACE: &[u8] = b"http://www.w3.org/XML/1998/namespace";
const OPENFORMULA_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:of:1.2";
const OPENOFFICE_CALC_NAMESPACE: &str = "http://openoffice.org/2004/calc";

/// How a validation list is displayed in the spreadsheet UI.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum DisplayList {
    #[default]
    None,
    Unsorted,
    SortAscending,
}

impl DisplayList {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "none" => Ok(Self::None),
            "unsorted" => Ok(Self::Unsorted),
            "sort-ascending" => Ok(Self::SortAscending),
            _ => Err(Error::InvalidFormat(format!(
                "invalid table:display-list value '{value}'"
            ))),
        }
    }

    fn as_str(self) -> &'static str {
        match self {
            Self::None => "none",
            Self::Unsorted => "unsorted",
            Self::SortAscending => "sort-ascending",
        }
    }
}

/// Severity associated with an invalid cell value.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MessageType {
    Stop,
    Warning,
    Information,
}

impl MessageType {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "stop" => Ok(Self::Stop),
            "warning" => Ok(Self::Warning),
            "information" => Ok(Self::Information),
            _ => Err(Error::InvalidFormat(format!(
                "invalid table:message-type value '{value}'"
            ))),
        }
    }

    fn as_str(self) -> &'static str {
        match self {
            Self::Stop => "stop",
            Self::Warning => "warning",
            Self::Information => "information",
        }
    }
}

/// Help or error text associated with a validation definition.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct Message {
    pub title: Option<String>,
    /// `None` preserves the ODF default rather than forcing an explicit value.
    pub display: Option<bool>,
    /// Text paragraphs in document order.
    pub paragraphs: Vec<String>,
}

/// Error behavior and text for a validation definition.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct ErrorMessage {
    pub title: Option<String>,
    pub display: Option<bool>,
    pub message_type: Option<MessageType>,
    pub paragraphs: Vec<String>,
}

/// A script event listener preserved as inert metadata.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ScriptEventListener {
    pub event_name: String,
    pub language: String,
    /// Exactly one of `macro_name` and `href` is required by ODF.
    pub macro_name: Option<String>,
    pub href: Option<String>,
    pub actuate: Option<String>,
}

/// Sound metadata attached to a presentation event listener.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PresentationSound {
    pub href: String,
    pub actuate: Option<String>,
    pub show: Option<String>,
    pub play_full: Option<bool>,
    pub xml_id: Option<String>,
}

/// A presentation event listener preserved as inert metadata.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PresentationEventListener {
    pub event_name: String,
    pub action: String,
    pub effect: Option<String>,
    pub direction: Option<String>,
    pub speed: Option<String>,
    pub start_scale: Option<String>,
    pub href: Option<String>,
    pub show: Option<String>,
    pub actuate: Option<String>,
    pub verb: Option<u64>,
    pub sound: Option<PresentationSound>,
}

/// An event listener following a validation error macro.
///
/// Litchi reads and writes this metadata but never invokes it.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum EventListener {
    Script(ScriptEventListener),
    Presentation(Box<PresentationEventListener>),
}

/// Macro metadata attached to invalid input. Litchi preserves it but never executes it.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct ErrorMacro {
    pub execute: Option<bool>,
    pub event_listeners: Vec<EventListener>,
}

impl ErrorMacro {
    fn validate(&self) -> Result<()> {
        for listener in &self.event_listeners {
            match listener {
                EventListener::Script(listener) => {
                    if listener.event_name.is_empty() || listener.language.is_empty() {
                        return Err(Error::InvalidFormat(
                            "validation script event listener requires an event name and language"
                                .to_string(),
                        ));
                    }
                    if listener.macro_name.is_some() == listener.href.is_some() {
                        return Err(Error::InvalidFormat(
                            "validation script event listener requires exactly one macro name or href"
                                .to_string(),
                        ));
                    }
                    if listener.actuate.is_some() && listener.href.is_none() {
                        return Err(Error::InvalidFormat(
                            "validation script event listener actuate requires an href".to_string(),
                        ));
                    }
                },
                EventListener::Presentation(listener) => {
                    if listener.event_name.is_empty() || listener.action.is_empty() {
                        return Err(Error::InvalidFormat(
                            "validation presentation event listener requires an event name and action"
                                .to_string(),
                        ));
                    }
                    if let Some(sound) = &listener.sound
                        && sound.href.is_empty()
                    {
                        return Err(Error::InvalidFormat(
                            "validation presentation sound href must not be empty".to_string(),
                        ));
                    }
                    if listener.href.is_none()
                        && (listener.show.is_some() || listener.actuate.is_some())
                    {
                        return Err(Error::InvalidFormat(
                            "validation presentation listener link behavior requires an href"
                                .to_string(),
                        ));
                    }
                },
            }
        }
        Ok(())
    }
}

/// A document-level `table:content-validation` definition.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ContentValidation {
    pub name: String,
    pub condition: Option<String>,
    /// Namespace binding used by a qualified validation condition.
    pub formula_namespace: Option<formula::Namespace>,
    pub base_cell_address: Option<String>,
    pub allow_empty_cell: Option<bool>,
    pub display_list: Option<DisplayList>,
    pub help_message: Option<Message>,
    pub error_message: Option<ErrorMessage>,
    pub error_macro: Option<ErrorMacro>,
}

impl ContentValidation {
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn new(name: impl Into<String>) -> Result<Self> {
        let validation = Self {
            name: name.into(),
            condition: None,
            formula_namespace: None,
            base_cell_address: None,
            allow_empty_cell: None,
            display_list: None,
            help_message: None,
            error_message: None,
            error_macro: None,
        };
        validation.validate()?;
        Ok(validation)
    }

    /// Set a validation condition, inferring standard `of` and `oooc` bindings.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn set_condition(&mut self, condition: impl Into<String>) -> Result<&mut Self> {
        let condition = condition.into();
        let namespace = default_formula_namespace(&condition);
        let previous_condition = self.condition.replace(condition);
        let previous_namespace = std::mem::replace(&mut self.formula_namespace, namespace);
        if let Err(error) = self.validate() {
            self.condition = previous_condition;
            self.formula_namespace = previous_namespace;
            return Err(error);
        }
        Ok(self)
    }

    /// Set a validation condition with an explicit formula namespace URI.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn set_condition_with_namespace(
        &mut self,
        condition: impl Into<String>,
        namespace_uri: impl Into<String>,
    ) -> Result<&mut Self> {
        let condition = condition.into();
        let prefix = condition_prefix(&condition).ok_or_else(|| {
            Error::InvalidFormat(
                "an explicit formula namespace requires a qualified validation condition"
                    .to_string(),
            )
        })?;
        let namespace = formula::Namespace {
            prefix: prefix.to_string(),
            uri: namespace_uri.into(),
        };
        let previous_condition = self.condition.replace(condition);
        let previous_namespace = self.formula_namespace.replace(namespace);
        if let Err(error) = self.validate() {
            self.condition = previous_condition;
            self.formula_namespace = previous_namespace;
            return Err(error);
        }
        Ok(self)
    }

    /// Remove the validation condition and its namespace binding.
    pub fn clear_condition(&mut self) {
        self.condition = None;
        self.formula_namespace = None;
    }

    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn validate(&self) -> Result<()> {
        if self.name.is_empty() {
            return Err(Error::InvalidFormat(
                "content validation name must not be empty".to_string(),
            ));
        }
        if self.error_message.is_some() && self.error_macro.is_some() {
            return Err(Error::InvalidFormat(format!(
                "content validation '{}' contains both an error message and error macro",
                self.name
            )));
        }
        match (&self.condition, &self.formula_namespace) {
            (Some(condition), Some(namespace)) => {
                if condition.trim().is_empty() || namespace.uri.trim().is_empty() {
                    return Err(Error::InvalidFormat(
                        "validation condition and formula namespace must not be empty".to_string(),
                    ));
                }
                validate_xml_prefix(&namespace.prefix)?;
                if condition_prefix(condition) != Some(namespace.prefix.as_str()) {
                    return Err(Error::InvalidFormat(format!(
                        "formula namespace prefix '{}' does not match validation condition",
                        namespace.prefix
                    )));
                }
            },
            (Some(condition), None) => {
                if condition.trim().is_empty() {
                    return Err(Error::InvalidFormat(
                        "validation condition must not be empty".to_string(),
                    ));
                }
                if let Some(prefix) = condition_prefix(condition) {
                    return Err(Error::InvalidFormat(format!(
                        "validation condition prefix '{prefix}' has no namespace binding"
                    )));
                }
            },
            (None, Some(_)) => {
                return Err(Error::InvalidFormat(
                    "formula namespace supplied without a validation condition".to_string(),
                ));
            },
            (None, None) => {},
        }
        if let Some(error_macro) = &self.error_macro {
            error_macro.validate()?;
        }
        Ok(())
    }
}

/// # Errors
///
/// Returns an error when a value violates the format or resource constraints.
pub fn validate_collection(validations: &[ContentValidation]) -> Result<()> {
    let mut names = HashSet::with_capacity(validations.len());
    for validation in validations {
        validation.validate()?;
        if !names.insert(validation.name.as_str()) {
            return Err(Error::InvalidFormat(format!(
                "duplicate content validation '{}'",
                validation.name
            )));
        }
    }
    Ok(())
}

pub fn write(out: &mut String, validations: &[ContentValidation]) {
    if validations.is_empty() {
        return;
    }
    out.push_str("<table:content-validations>");
    for validation in validations {
        out.push_str("<table:content-validation table:name=\"");
        out.push_str(&escape_xml(&validation.name));
        out.push('"');
        if let Some(namespace) = &validation.formula_namespace {
            out.push_str(" xmlns:");
            out.push_str(&namespace.prefix);
            out.push_str("=\"");
            out.push_str(&escape_xml(&namespace.uri));
            out.push('"');
        }
        write_optional_attribute(out, "table:condition", validation.condition.as_deref());
        write_optional_attribute(
            out,
            "table:base-cell-address",
            validation.base_cell_address.as_deref(),
        );
        if let Some(value) = validation.allow_empty_cell {
            write_bool_attribute(out, "table:allow-empty-cell", value);
        }
        if let Some(value) = validation.display_list {
            write_optional_attribute(out, "table:display-list", Some(value.as_str()));
        }
        if validation.help_message.is_none()
            && validation.error_message.is_none()
            && validation.error_macro.is_none()
        {
            out.push_str("/>");
            continue;
        }
        out.push('>');
        if let Some(message) = &validation.help_message {
            write_message(
                out,
                "help-message",
                message.title.as_deref(),
                message.display,
                None,
                &message.paragraphs,
            );
        }
        if let Some(message) = &validation.error_message {
            write_message(
                out,
                "error-message",
                message.title.as_deref(),
                message.display,
                message.message_type,
                &message.paragraphs,
            );
        } else if let Some(macro_metadata) = &validation.error_macro {
            out.push_str("<table:error-macro");
            if let Some(execute) = macro_metadata.execute {
                write_bool_attribute(out, "table:execute", execute);
            }
            out.push_str("/>");
            write_event_listeners(out, &macro_metadata.event_listeners);
        }
        out.push_str("</table:content-validation>");
    }
    out.push_str("</table:content-validations>");
}

fn write_event_listeners(out: &mut String, listeners: &[EventListener]) {
    if listeners.is_empty() {
        return;
    }
    out.push_str("<office:event-listeners>");
    for listener in listeners {
        match listener {
            EventListener::Script(listener) => {
                out.push_str("<script:event-listener");
                write_optional_attribute(out, "script:event-name", Some(&listener.event_name));
                write_optional_attribute(out, "script:language", Some(&listener.language));
                write_optional_attribute(out, "script:macro-name", listener.macro_name.as_deref());
                if let Some(href) = &listener.href {
                    write_optional_attribute(out, "xlink:type", Some("simple"));
                    write_optional_attribute(out, "xlink:href", Some(href));
                    write_optional_attribute(out, "xlink:actuate", listener.actuate.as_deref());
                }
                out.push_str("/>");
            },
            EventListener::Presentation(listener) => {
                out.push_str("<presentation:event-listener");
                write_optional_attribute(out, "script:event-name", Some(&listener.event_name));
                write_optional_attribute(out, "presentation:action", Some(&listener.action));
                write_optional_attribute(out, "presentation:effect", listener.effect.as_deref());
                write_optional_attribute(
                    out,
                    "presentation:direction",
                    listener.direction.as_deref(),
                );
                write_optional_attribute(out, "presentation:speed", listener.speed.as_deref());
                write_optional_attribute(
                    out,
                    "presentation:start-scale",
                    listener.start_scale.as_deref(),
                );
                if let Some(href) = &listener.href {
                    write_optional_attribute(out, "xlink:type", Some("simple"));
                    write_optional_attribute(out, "xlink:href", Some(href));
                    write_optional_attribute(out, "xlink:show", listener.show.as_deref());
                    write_optional_attribute(out, "xlink:actuate", listener.actuate.as_deref());
                }
                if let Some(verb) = listener.verb {
                    write_optional_attribute(out, "presentation:verb", Some(&verb.to_string()));
                }
                if let Some(sound) = &listener.sound {
                    out.push('>');
                    out.push_str("<presentation:sound");
                    write_optional_attribute(out, "xlink:type", Some("simple"));
                    write_optional_attribute(out, "xlink:href", Some(&sound.href));
                    write_optional_attribute(out, "xlink:actuate", sound.actuate.as_deref());
                    write_optional_attribute(out, "xlink:show", sound.show.as_deref());
                    if let Some(play_full) = sound.play_full {
                        write_bool_attribute(out, "presentation:play-full", play_full);
                    }
                    write_optional_attribute(out, "xml:id", sound.xml_id.as_deref());
                    out.push_str("/></presentation:event-listener>");
                } else {
                    out.push_str("/>");
                }
            },
        }
    }
    out.push_str("</office:event-listeners>");
}

fn write_message(
    out: &mut String,
    element: &str,
    title: Option<&str>,
    display: Option<bool>,
    message_type: Option<MessageType>,
    paragraphs: &[String],
) {
    out.push_str("<table:");
    out.push_str(element);
    write_optional_attribute(out, "table:title", title);
    if let Some(display) = display {
        write_bool_attribute(out, "table:display", display);
    }
    if let Some(message_type) = message_type {
        write_optional_attribute(out, "table:message-type", Some(message_type.as_str()));
    }
    if paragraphs.is_empty() {
        out.push_str("/>");
        return;
    }
    out.push('>');
    for paragraph in paragraphs {
        out.push_str("<text:p>");
        write_paragraph_text(out, paragraph);
        out.push_str("</text:p>");
    }
    out.push_str("</table:");
    out.push_str(element);
    out.push('>');
}

fn write_paragraph_text(out: &mut String, paragraph: &str) {
    let mut text = String::new();
    let flush_text = |out: &mut String, text: &mut String| {
        if !text.is_empty() {
            out.push_str(&escape_xml(text));
            text.clear();
        }
    };
    let mut characters = paragraph.chars().peekable();
    while let Some(character) = characters.next() {
        match character {
            ' ' => {
                flush_text(out, &mut text);
                let mut count = 1usize;
                while characters.next_if_eq(&' ').is_some() {
                    count += 1;
                }
                out.push_str("<text:s");
                if count > 1 {
                    out.push_str(" text:c=\"");
                    out.push_str(&count.to_string());
                    out.push('"');
                }
                out.push_str("/>");
            },
            '\t' => {
                flush_text(out, &mut text);
                out.push_str("<text:tab/>");
            },
            '\n' => {
                flush_text(out, &mut text);
                out.push_str("<text:line-break/>");
            },
            _ => text.push(character),
        }
    }
    flush_text(out, &mut text);
}

fn write_optional_attribute(out: &mut String, name: &str, value: Option<&str>) {
    if let Some(value) = value {
        out.push(' ');
        out.push_str(name);
        out.push_str("=\"");
        out.push_str(&escape_xml(value));
        out.push('"');
    }
}

fn write_bool_attribute(out: &mut String, name: &str, value: bool) {
    write_optional_attribute(out, name, Some(if value { "true" } else { "false" }));
}

/// # Errors
///
/// Returns an error when the input is malformed or exceeds the parser's resource limits.
///
/// # Panics
///
/// Panics if the parser's internal state becomes inconsistent; every `expect` is guarded by a preceding state check.
pub fn parse(xml: &str) -> Result<Vec<ContentValidation>> {
    let mut reader = NsReader::from_str(xml);
    let mut buf = Vec::new();
    let mut validations = Vec::new();
    let mut inside_collection = false;
    let mut inside_event_listeners = false;
    let mut event_listeners_seen = false;
    let mut current: Option<ContentValidation> = None;
    let mut message: Option<(bool, MessageBuilder)> = None;
    let mut presentation_listener: Option<PresentationEventListener> = None;

    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buf)
            .map_err(|error| Error::InvalidFormat(format!("XML parsing error: {error}")))?;
        match event {
            Event::Start(element) if is_namespace(&namespace, TABLE_NAMESPACE) => {
                match element.local_name().as_ref() {
                    b"content-validations" => {
                        if inside_collection {
                            return Err(Error::InvalidFormat(
                                "nested table:content-validations element".to_string(),
                            ));
                        }
                        inside_collection = true;
                    },
                    b"content-validation" if inside_collection => {
                        if current.is_some() {
                            return Err(Error::InvalidFormat(
                                "nested table:content-validation element".to_string(),
                            ));
                        }
                        event_listeners_seen = false;
                        current = Some(parse_attributes(&reader, &element)?);
                    },
                    b"help-message" | b"error-message" if current.is_some() => {
                        if message.is_some() {
                            return Err(Error::InvalidFormat(
                                "nested validation message".to_string(),
                            ));
                        }
                        message = Some((
                            element.local_name().as_ref() == b"error-message",
                            MessageBuilder::new(&reader, &element)?,
                        ));
                    },
                    b"error-macro" if current.is_some() => {
                        let execute = optional_bool_attribute(
                            reader.resolver(),
                            reader.decoder(),
                            &element,
                            b"execute",
                        )?;
                        set_error_macro(current.as_mut().expect("checked validation"), execute)?;
                    },
                    _ => {},
                }
            },
            Event::Start(element)
                if is_namespace(&namespace, OFFICE_NAMESPACE)
                    && element.local_name().as_ref() == b"event-listeners"
                    && current.is_some() =>
            {
                let error_macro = current
                    .as_ref()
                    .and_then(|validation| validation.error_macro.as_ref())
                    .ok_or_else(|| {
                        Error::InvalidFormat(
                            "validation event listeners require an error macro".to_string(),
                        )
                    })?;
                if inside_event_listeners
                    || event_listeners_seen
                    || !error_macro.event_listeners.is_empty()
                {
                    return Err(Error::InvalidFormat(
                        "duplicate validation event-listeners element".to_string(),
                    ));
                }
                inside_event_listeners = true;
                event_listeners_seen = true;
            },
            Event::Empty(element)
                if is_namespace(&namespace, OFFICE_NAMESPACE)
                    && element.local_name().as_ref() == b"event-listeners"
                    && current.is_some() =>
            {
                if current
                    .as_ref()
                    .and_then(|validation| validation.error_macro.as_ref())
                    .is_none()
                {
                    return Err(Error::InvalidFormat(
                        "validation event listeners require an error macro".to_string(),
                    ));
                }
                if event_listeners_seen {
                    return Err(Error::InvalidFormat(
                        "duplicate validation event-listeners element".to_string(),
                    ));
                }
                event_listeners_seen = true;
            },
            Event::Start(element) | Event::Empty(element)
                if is_namespace(&namespace, SCRIPT_NAMESPACE)
                    && element.local_name().as_ref() == b"event-listener"
                    && inside_event_listeners =>
            {
                let listener = parse_script_event_listener(&reader, &element)?;
                current
                    .as_mut()
                    .and_then(|validation| validation.error_macro.as_mut())
                    .expect("event-listeners require an error macro")
                    .event_listeners
                    .push(EventListener::Script(listener));
            },
            Event::Start(element)
                if is_namespace(&namespace, PRESENTATION_NAMESPACE)
                    && element.local_name().as_ref() == b"event-listener"
                    && inside_event_listeners =>
            {
                if presentation_listener.is_some() {
                    return Err(Error::InvalidFormat(
                        "nested presentation event listener".to_string(),
                    ));
                }
                presentation_listener = Some(parse_presentation_event_listener(&reader, &element)?);
            },
            Event::Empty(element)
                if is_namespace(&namespace, PRESENTATION_NAMESPACE)
                    && element.local_name().as_ref() == b"event-listener"
                    && inside_event_listeners =>
            {
                let listener = parse_presentation_event_listener(&reader, &element)?;
                current
                    .as_mut()
                    .and_then(|validation| validation.error_macro.as_mut())
                    .expect("event-listeners require an error macro")
                    .event_listeners
                    .push(EventListener::Presentation(Box::new(listener)));
            },
            Event::Start(element) | Event::Empty(element)
                if is_namespace(&namespace, PRESENTATION_NAMESPACE)
                    && element.local_name().as_ref() == b"sound"
                    && presentation_listener.is_some() =>
            {
                let listener = presentation_listener.as_mut().expect("checked listener");
                if listener.sound.is_some() {
                    return Err(Error::InvalidFormat(
                        "duplicate presentation event sound".to_string(),
                    ));
                }
                listener.sound = Some(parse_presentation_sound(&reader, &element)?);
            },
            Event::Empty(element) if is_namespace(&namespace, TABLE_NAMESPACE) => {
                match element.local_name().as_ref() {
                    b"content-validation" if inside_collection => {
                        let validation = parse_attributes(&reader, &element)?;
                        validation.validate()?;
                        validations.push(validation);
                    },
                    b"help-message" | b"error-message" if current.is_some() => {
                        let is_error = element.local_name().as_ref() == b"error-message";
                        let builder = MessageBuilder::new(&reader, &element)?;
                        finish_message(
                            current.as_mut().expect("checked validation"),
                            is_error,
                            builder,
                        )?;
                    },
                    b"error-macro" if current.is_some() => {
                        let execute = optional_bool_attribute(
                            reader.resolver(),
                            reader.decoder(),
                            &element,
                            b"execute",
                        )?;
                        set_error_macro(current.as_mut().expect("checked validation"), execute)?;
                    },
                    _ => {},
                }
            },
            Event::Start(element)
                if is_namespace(&namespace, TEXT_NAMESPACE)
                    && element.local_name().as_ref() == b"p"
                    && message.is_some() =>
            {
                message
                    .as_mut()
                    .expect("checked message")
                    .1
                    .start_paragraph()?;
            },
            Event::Empty(element)
                if is_namespace(&namespace, TEXT_NAMESPACE)
                    && element.local_name().as_ref() == b"p"
                    && message.is_some() =>
            {
                let builder = &mut message.as_mut().expect("checked message").1;
                builder.start_paragraph()?;
                builder.end_paragraph()?;
            },
            Event::Empty(element)
                if is_namespace(&namespace, TEXT_NAMESPACE) && message.is_some() =>
            {
                let builder = &mut message.as_mut().expect("checked message").1;
                match element.local_name().as_ref() {
                    b"line-break" => builder.push_text("\n"),
                    b"tab" => builder.push_text("\t"),
                    b"s" => {
                        let count = optional_text_space_count(
                            reader.resolver(),
                            reader.decoder(),
                            &element,
                        )?;
                        if count > 1_000_000 {
                            return Err(Error::InvalidFormat(
                                "text:s count exceeds the supported safety limit".to_string(),
                            ));
                        }
                        builder.push_spaces(count);
                    },
                    _ => {},
                }
            },
            Event::Text(text) if message.is_some() => {
                let decoded = text.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                    Error::InvalidFormat(format!("invalid validation message text: {error}"))
                })?;
                let decoded = quick_xml::escape::unescape(&decoded).map_err(|error| {
                    Error::InvalidFormat(format!("invalid validation character reference: {error}"))
                })?;
                message
                    .as_mut()
                    .expect("checked message")
                    .1
                    .push_text(&decoded);
            },
            Event::CData(text) if message.is_some() => {
                let decoded = text.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                    Error::InvalidFormat(format!("invalid validation CDATA: {error}"))
                })?;
                message
                    .as_mut()
                    .expect("checked message")
                    .1
                    .push_text(&decoded);
            },
            Event::End(element)
                if is_namespace(&namespace, TEXT_NAMESPACE)
                    && element.local_name().as_ref() == b"p"
                    && message.is_some() =>
            {
                message
                    .as_mut()
                    .expect("checked message")
                    .1
                    .end_paragraph()?;
            },
            Event::End(element) if is_namespace(&namespace, TABLE_NAMESPACE) => {
                match element.local_name().as_ref() {
                    b"help-message" | b"error-message" if message.is_some() => {
                        let (is_error, builder) = message.take().expect("checked message");
                        finish_message(
                            current.as_mut().expect("message requires validation"),
                            is_error,
                            builder,
                        )?;
                    },
                    b"content-validation" => {
                        let validation = current.take().ok_or_else(|| {
                            Error::InvalidFormat("unexpected content-validation end".to_string())
                        })?;
                        validation.validate()?;
                        validations.push(validation);
                    },
                    b"content-validations" => inside_collection = false,
                    _ => {},
                }
            },
            Event::End(element)
                if is_namespace(&namespace, PRESENTATION_NAMESPACE)
                    && element.local_name().as_ref() == b"event-listener"
                    && inside_event_listeners =>
            {
                let listener = presentation_listener.take().ok_or_else(|| {
                    Error::InvalidFormat("unexpected presentation event-listener end".to_string())
                })?;
                current
                    .as_mut()
                    .and_then(|validation| validation.error_macro.as_mut())
                    .expect("event-listeners require an error macro")
                    .event_listeners
                    .push(EventListener::Presentation(Box::new(listener)));
            },
            Event::End(element)
                if is_namespace(&namespace, OFFICE_NAMESPACE)
                    && element.local_name().as_ref() == b"event-listeners"
                    && inside_event_listeners =>
            {
                if presentation_listener.is_some() {
                    return Err(Error::InvalidFormat(
                        "unexpected validation event-listeners end".to_string(),
                    ));
                }
                inside_event_listeners = false;
            },
            Event::Eof => break,
            Event::Start(_)
            | Event::End(_)
            | Event::Empty(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }
        buf.clear();
    }
    if current.is_some()
        || message.is_some()
        || presentation_listener.is_some()
        || inside_event_listeners
        || inside_collection
    {
        return Err(Error::InvalidFormat(
            "unterminated content-validation collection".to_string(),
        ));
    }
    validate_collection(&validations)?;
    Ok(validations)
}

fn parse_attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<ContentValidation> {
    let condition = optional_attribute(reader.resolver(), reader.decoder(), element, b"condition")?;
    let formula_namespace = condition
        .as_deref()
        .map(|condition| formula_namespace(reader.resolver(), condition))
        .transpose()?
        .flatten();
    Ok(ContentValidation {
        name: required_attribute(reader.resolver(), reader.decoder(), element, b"name")?,
        condition,
        formula_namespace,
        base_cell_address: optional_attribute(
            reader.resolver(),
            reader.decoder(),
            element,
            b"base-cell-address",
        )?,
        allow_empty_cell: optional_bool_attribute(
            reader.resolver(),
            reader.decoder(),
            element,
            b"allow-empty-cell",
        )?,
        display_list: optional_attribute(
            reader.resolver(),
            reader.decoder(),
            element,
            b"display-list",
        )?
        .map(|value| DisplayList::parse(&value))
        .transpose()?,
        help_message: None,
        error_message: None,
        error_macro: None,
    })
}

fn set_error_macro(validation: &mut ContentValidation, execute: Option<bool>) -> Result<()> {
    if validation.error_message.is_some() || validation.error_macro.is_some() {
        return Err(Error::InvalidFormat(
            "duplicate validation error behavior".to_string(),
        ));
    }
    validation.error_macro = Some(ErrorMacro {
        execute,
        event_listeners: Vec::new(),
    });
    Ok(())
}

fn parse_script_event_listener(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<ScriptEventListener> {
    Ok(ScriptEventListener {
        event_name: required_attribute_ns(
            reader.resolver(),
            reader.decoder(),
            element,
            SCRIPT_NAMESPACE,
            b"event-name",
        )?,
        language: required_attribute_ns(
            reader.resolver(),
            reader.decoder(),
            element,
            SCRIPT_NAMESPACE,
            b"language",
        )?,
        macro_name: optional_attribute_ns(
            reader.resolver(),
            reader.decoder(),
            element,
            SCRIPT_NAMESPACE,
            b"macro-name",
        )?,
        href: optional_attribute_ns(
            reader.resolver(),
            reader.decoder(),
            element,
            XLINK_NAMESPACE,
            b"href",
        )?,
        actuate: optional_attribute_ns(
            reader.resolver(),
            reader.decoder(),
            element,
            XLINK_NAMESPACE,
            b"actuate",
        )?,
    })
}

fn parse_presentation_event_listener(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<PresentationEventListener> {
    let verb = optional_attribute_ns(
        reader.resolver(),
        reader.decoder(),
        element,
        PRESENTATION_NAMESPACE,
        b"verb",
    )?
    .map(|value| {
        value
            .parse::<u64>()
            .map_err(|_error| Error::InvalidFormat(format!("invalid presentation:verb '{value}'")))
    })
    .transpose()?;
    Ok(PresentationEventListener {
        event_name: required_attribute_ns(
            reader.resolver(),
            reader.decoder(),
            element,
            SCRIPT_NAMESPACE,
            b"event-name",
        )?,
        action: required_attribute_ns(
            reader.resolver(),
            reader.decoder(),
            element,
            PRESENTATION_NAMESPACE,
            b"action",
        )?,
        effect: optional_attribute_ns(
            reader.resolver(),
            reader.decoder(),
            element,
            PRESENTATION_NAMESPACE,
            b"effect",
        )?,
        direction: optional_attribute_ns(
            reader.resolver(),
            reader.decoder(),
            element,
            PRESENTATION_NAMESPACE,
            b"direction",
        )?,
        speed: optional_attribute_ns(
            reader.resolver(),
            reader.decoder(),
            element,
            PRESENTATION_NAMESPACE,
            b"speed",
        )?,
        start_scale: optional_attribute_ns(
            reader.resolver(),
            reader.decoder(),
            element,
            PRESENTATION_NAMESPACE,
            b"start-scale",
        )?,
        href: optional_attribute_ns(
            reader.resolver(),
            reader.decoder(),
            element,
            XLINK_NAMESPACE,
            b"href",
        )?,
        show: optional_attribute_ns(
            reader.resolver(),
            reader.decoder(),
            element,
            XLINK_NAMESPACE,
            b"show",
        )?,
        actuate: optional_attribute_ns(
            reader.resolver(),
            reader.decoder(),
            element,
            XLINK_NAMESPACE,
            b"actuate",
        )?,
        verb,
        sound: None,
    })
}

fn parse_presentation_sound(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<PresentationSound> {
    Ok(PresentationSound {
        href: required_attribute_ns(
            reader.resolver(),
            reader.decoder(),
            element,
            XLINK_NAMESPACE,
            b"href",
        )?,
        actuate: optional_attribute_ns(
            reader.resolver(),
            reader.decoder(),
            element,
            XLINK_NAMESPACE,
            b"actuate",
        )?,
        show: optional_attribute_ns(
            reader.resolver(),
            reader.decoder(),
            element,
            XLINK_NAMESPACE,
            b"show",
        )?,
        play_full: optional_bool_attribute_ns(
            reader.resolver(),
            reader.decoder(),
            element,
            PRESENTATION_NAMESPACE,
            b"play-full",
        )?,
        xml_id: optional_attribute_ns(
            reader.resolver(),
            reader.decoder(),
            element,
            XML_NAMESPACE,
            b"id",
        )?,
    })
}

fn finish_message(
    validation: &mut ContentValidation,
    is_error: bool,
    builder: MessageBuilder,
) -> Result<()> {
    if is_error {
        if validation.error_message.is_some() || validation.error_macro.is_some() {
            return Err(Error::InvalidFormat(
                "duplicate validation error behavior".to_string(),
            ));
        }
        validation.error_message = Some(ErrorMessage {
            title: builder.title,
            display: builder.display,
            message_type: builder.message_type,
            paragraphs: builder.paragraphs,
        });
    } else {
        if validation.help_message.is_some() {
            return Err(Error::InvalidFormat(
                "duplicate validation help message".to_string(),
            ));
        }
        validation.help_message = Some(Message {
            title: builder.title,
            display: builder.display,
            paragraphs: builder.paragraphs,
        });
    }
    Ok(())
}

struct MessageBuilder {
    title: Option<String>,
    display: Option<bool>,
    message_type: Option<MessageType>,
    paragraphs: Vec<String>,
    paragraph: Option<String>,
}

impl MessageBuilder {
    fn new(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<Self> {
        Ok(Self {
            title: optional_attribute(reader.resolver(), reader.decoder(), element, b"title")?,
            display: optional_bool_attribute(
                reader.resolver(),
                reader.decoder(),
                element,
                b"display",
            )?,
            message_type: optional_attribute(
                reader.resolver(),
                reader.decoder(),
                element,
                b"message-type",
            )?
            .map(|value| MessageType::parse(&value))
            .transpose()?,
            paragraphs: Vec::new(),
            paragraph: None,
        })
    }

    fn start_paragraph(&mut self) -> Result<()> {
        if self.paragraph.is_some() {
            return Err(Error::InvalidFormat(
                "nested validation paragraph".to_string(),
            ));
        }
        self.paragraph = Some(String::new());
        Ok(())
    }

    fn push_text(&mut self, text: &str) {
        if let Some(paragraph) = &mut self.paragraph {
            paragraph.push_str(text);
        }
    }

    fn push_spaces(&mut self, count: usize) {
        if let Some(paragraph) = &mut self.paragraph {
            paragraph.extend(std::iter::repeat_n(' ', count));
        }
    }

    fn end_paragraph(&mut self) -> Result<()> {
        self.paragraph
            .take()
            .map(|value| self.paragraphs.push(value))
            .ok_or_else(|| {
                Error::InvalidFormat("validation paragraph end without start".to_string())
            })
    }
}

fn is_namespace(namespace: &ResolveResult<'_>, expected: &[u8]) -> bool {
    matches!(namespace, ResolveResult::Bound(XmlNamespace(value)) if *value == expected)
}

fn required_attribute(
    resolver: &NamespaceResolver,
    decoder: quick_xml::encoding::Decoder,
    element: &BytesStart<'_>,
    local_name: &[u8],
) -> Result<String> {
    required_attribute_ns(resolver, decoder, element, TABLE_NAMESPACE, local_name)
}

fn required_attribute_ns(
    resolver: &NamespaceResolver,
    decoder: quick_xml::encoding::Decoder,
    element: &BytesStart<'_>,
    namespace: &[u8],
    local_name: &[u8],
) -> Result<String> {
    optional_attribute_ns(resolver, decoder, element, namespace, local_name)?.ok_or_else(|| {
        Error::InvalidFormat(format!(
            "{} is missing required {} attribute",
            String::from_utf8_lossy(element.local_name().as_ref()),
            String::from_utf8_lossy(local_name)
        ))
    })
}

fn optional_attribute(
    resolver: &NamespaceResolver,
    decoder: quick_xml::encoding::Decoder,
    element: &BytesStart<'_>,
    local_name: &[u8],
) -> Result<Option<String>> {
    optional_attribute_ns(resolver, decoder, element, TABLE_NAMESPACE, local_name)
}

fn optional_attribute_ns(
    resolver: &NamespaceResolver,
    decoder: quick_xml::encoding::Decoder,
    element: &BytesStart<'_>,
    namespace_uri: &[u8],
    local_name: &[u8],
) -> Result<Option<String>> {
    for attribute in element.attributes() {
        let attribute = attribute
            .map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")))?;
        let (namespace, local) = resolver.resolve_attribute(attribute.key);
        if is_namespace(&namespace, namespace_uri) && local.as_ref() == local_name {
            return attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                .map(|value| Some(value.into_owned()))
                .map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")));
        }
    }
    Ok(None)
}

fn optional_bool_attribute(
    resolver: &NamespaceResolver,
    decoder: quick_xml::encoding::Decoder,
    element: &BytesStart<'_>,
    local_name: &[u8],
) -> Result<Option<bool>> {
    optional_bool_attribute_ns(resolver, decoder, element, TABLE_NAMESPACE, local_name)
}

fn optional_bool_attribute_ns(
    resolver: &NamespaceResolver,
    decoder: quick_xml::encoding::Decoder,
    element: &BytesStart<'_>,
    namespace: &[u8],
    local_name: &[u8],
) -> Result<Option<bool>> {
    optional_attribute_ns(resolver, decoder, element, namespace, local_name)?
        .map(|value| match value.as_str() {
            "true" | "1" => Ok(true),
            "false" | "0" => Ok(false),
            _ => Err(Error::InvalidFormat(format!(
                "invalid Boolean value '{value}'"
            ))),
        })
        .transpose()
}

fn optional_text_space_count(
    resolver: &NamespaceResolver,
    decoder: quick_xml::encoding::Decoder,
    element: &BytesStart<'_>,
) -> Result<usize> {
    for attribute in element.attributes() {
        let attribute = attribute
            .map_err(|error| Error::InvalidFormat(format!("invalid text:s attribute: {error}")))?;
        let (namespace, local) = resolver.resolve_attribute(attribute.key);
        if is_namespace(&namespace, TEXT_NAMESPACE) && local.as_ref() == b"c" {
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                .map_err(|error| Error::InvalidFormat(format!("invalid text:s count: {error}")))?;
            return value
                .parse::<usize>()
                .map_err(|_error| Error::InvalidFormat(format!("invalid text:s count '{value}'")));
        }
    }
    Ok(1)
}

fn condition_prefix(condition: &str) -> Option<&str> {
    let (prefix, remainder) = condition.split_once(':')?;
    if prefix.is_empty() || remainder.is_empty() {
        None
    } else {
        Some(prefix)
    }
}

fn default_formula_namespace(condition: &str) -> Option<formula::Namespace> {
    match condition_prefix(condition) {
        Some("of") => Some(formula::Namespace {
            prefix: "of".to_string(),
            uri: OPENFORMULA_NAMESPACE.to_string(),
        }),
        Some("oooc") => Some(formula::Namespace {
            prefix: "oooc".to_string(),
            uri: OPENOFFICE_CALC_NAMESPACE.to_string(),
        }),
        _ => None,
    }
}

fn formula_namespace(
    resolver: &NamespaceResolver,
    condition: &str,
) -> Result<Option<formula::Namespace>> {
    let Some(prefix) = condition_prefix(condition) else {
        return Ok(None);
    };
    for (declaration, namespace) in resolver.bindings() {
        if let PrefixDeclaration::Named(candidate) = declaration
            && candidate == prefix.as_bytes()
        {
            let uri = String::from_utf8(namespace.as_ref().to_vec()).map_err(|_error| {
                Error::InvalidFormat(format!(
                    "validation formula namespace for prefix '{prefix}' is not UTF-8"
                ))
            })?;
            return Ok(Some(formula::Namespace {
                prefix: prefix.to_string(),
                uri,
            }));
        }
    }
    Err(Error::InvalidFormat(format!(
        "validation condition prefix '{prefix}' is not bound to a namespace"
    )))
}

fn validate_xml_prefix(prefix: &str) -> Result<()> {
    let mut bytes = prefix.bytes();
    let Some(first) = bytes.next() else {
        return Err(Error::InvalidFormat(
            "formula namespace prefix must not be empty".to_string(),
        ));
    };
    if !(first.is_ascii_alphabetic() || first == b'_')
        || !bytes.all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'_' | b'-' | b'.'))
    {
        return Err(Error::InvalidFormat(format!(
            "invalid formula namespace prefix '{prefix}'"
        )));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parses_and_writes_complete_validation_collection() {
        let xml = r#"<office:spreadsheet
            xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
            xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
            xmlns:x="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
            xmlns:f="urn:oasis:names:tc:opendocument:xmlns:of:1.2">
          <t:content-validations><t:content-validation t:name="whole" t:condition="f:cell-content-is-whole-number()" t:base-cell-address="$Sheet1.$A$1" t:allow-empty-cell="true" t:display-list="unsorted">
            <t:help-message t:title="Input" t:display="true"><x:p>Enter a whole number</x:p></t:help-message>
            <t:error-message t:title="Invalid" t:display="true" t:message-type="stop"><x:p>Try again</x:p></t:error-message>
          </t:content-validation></t:content-validations>
        </office:spreadsheet>"#;
        let parsed = parse(xml).expect("test fixture or operation should succeed");
        assert_eq!(parsed.len(), 1);
        assert_eq!(parsed[0].name, "whole");
        assert_eq!(
            parsed[0].formula_namespace,
            Some(formula::Namespace {
                prefix: "f".to_string(),
                uri: OPENFORMULA_NAMESPACE.to_string(),
            })
        );
        assert_eq!(parsed[0].display_list, Some(DisplayList::Unsorted));
        assert_eq!(
            parsed[0]
                .help_message
                .as_ref()
                .expect("test fixture or operation should succeed")
                .paragraphs,
            ["Enter a whole number"]
        );
        assert_eq!(
            parsed[0]
                .error_message
                .as_ref()
                .expect("test fixture or operation should succeed")
                .message_type,
            Some(MessageType::Stop)
        );
        let mut encoded = String::new();
        write(&mut encoded, &parsed);
        assert_eq!(parse(&format!(
            r#"<office:spreadsheet xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">{encoded}</office:spreadsheet>"#
        )).expect("test fixture or operation should succeed"), parsed);
    }

    #[test]
    fn rejects_duplicates_invalid_enums_and_conflicting_error_actions() {
        let duplicate = r#"<t:content-validations xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><t:content-validation t:name="x"/><t:content-validation t:name="x"/></t:content-validations>"#;
        assert!(parse(duplicate).is_err());
        let invalid = r#"<t:content-validations xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><t:content-validation t:name="x" t:display-list="bad"/></t:content-validations>"#;
        assert!(parse(invalid).is_err());
        let mut validation =
            ContentValidation::new("x").expect("test fixture or operation should succeed");
        validation.error_message = Some(ErrorMessage::default());
        validation.error_macro = Some(ErrorMacro::default());
        assert!(validation.validate().is_err());
    }

    #[test]
    fn preserves_odf_whitespace_elements_in_messages() {
        let xml = r#"<t:content-validations xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:x="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><t:content-validation t:name="x"><t:help-message><x:p>two<x:s x:c="2"/>spaces<x:tab/>tab<x:line-break/>line</x:p></t:help-message></t:content-validation></t:content-validations>"#;
        let parsed = parse(xml).expect("test fixture or operation should succeed");
        assert_eq!(
            parsed[0]
                .help_message
                .as_ref()
                .expect("test fixture or operation should succeed")
                .paragraphs,
            ["two  spaces\ttab\nline"]
        );

        let mut encoded = String::new();
        write(&mut encoded, &parsed);
        assert!(encoded.contains("<text:s text:c=\"2\"/>"));
        assert!(encoded.contains("<text:tab/>"));
        assert!(encoded.contains("<text:line-break/>"));
    }

    #[test]
    fn preserves_inert_validation_event_listener_metadata() {
        let xml = r#"<t:content-validations
            xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
            xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
            xmlns:s="urn:oasis:names:tc:opendocument:xmlns:script:1.0"
            xmlns:p="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0"
            xmlns:l="http://www.w3.org/1999/xlink">
          <t:content-validation t:name="macro"><t:error-macro t:execute="false"/>
            <o:event-listeners>
              <s:event-listener s:event-name="dom:click" s:language="ooo:script" s:macro-name="Standard.Module1.Main"/>
              <p:event-listener s:event-name="dom:activate" p:action="sound" p:speed="medium" p:verb="2">
                <p:sound l:type="simple" l:href="Sounds/alert.wav" l:show="new" p:play-full="true" xml:id="sound1"/>
              </p:event-listener>
            </o:event-listeners>
          </t:content-validation>
        </t:content-validations>"#;
        let parsed = parse(xml).expect("test fixture or operation should succeed");
        let error_macro = parsed[0]
            .error_macro
            .as_ref()
            .expect("test fixture or operation should succeed");
        assert_eq!(error_macro.execute, Some(false));
        assert_eq!(error_macro.event_listeners.len(), 2);
        let EventListener::Presentation(listener) = &error_macro.event_listeners[1] else {
            panic!("expected presentation listener");
        };
        assert_eq!(listener.verb, Some(2));
        assert_eq!(
            listener
                .sound
                .as_ref()
                .expect("test fixture or operation should succeed")
                .xml_id
                .as_deref(),
            Some("sound1")
        );

        let mut encoded = String::new();
        write(&mut encoded, &parsed);
        let document = format!(
            r#"<office:spreadsheet xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:xlink="http://www.w3.org/1999/xlink">{encoded}</office:spreadsheet>"#
        );
        assert_eq!(
            parse(&document).expect("test fixture or operation should succeed"),
            parsed
        );
    }
}
