//! Bounded, lossless XML codec for ODF spreadsheet content validations.

use super::model::{
    ContentValidation, ContentValidationPart, ContentValidations, MAX_CAPTURE_BYTES,
    MAX_PARAGRAPHS, MAX_VALIDATIONS, MAX_VALUE_BYTES, ValidationCellAddress, ValidationCondition,
    ValidationDisplayList, ValidationEventListeners, ValidationFailure, ValidationMessage,
    ValidationMessageType, ValidationParagraph, invalid, make_error, validate_text,
};
use litchi_core::Result;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;
use quick_xml::{Writer, XmlVersion};
use std::collections::HashMap;
use std::io::BufRead;

const OFFICE_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TABLE_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const TEXT_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const MAX_XML_BYTES: usize = 64 * 1_048_576;
const MAX_DEPTH: usize = 256;

impl ContentValidations {
    pub fn to_xml_fragment(&self) -> Result<String> {
        self.validate()?;
        if self.validations.is_empty() {
            return invalid("table:content-validations requires at least one validation");
        }
        let mut output = String::with_capacity(512 + self.validations.len() * 256);
        output.push_str(r#"<table:content-validations xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0">"#);
        for validation in &self.validations {
            write_validation(&mut output, validation);
        }
        output.push_str("</table:content-validations>");
        Ok(output)
    }
}

fn write_validation(output: &mut String, value: &ContentValidation) {
    output.push_str("<table:content-validation table:name=\"");
    escape_attribute(output, &value.name);
    output.push('"');
    if let Some(condition) = &value.condition {
        output.push_str(" table:condition=\"");
        escape_attribute(output, condition.as_str());
        output.push('"');
    }
    if let Some(address) = &value.base_cell_address {
        output.push_str(" table:base-cell-address=\"");
        escape_attribute(output, address.as_str());
        output.push('"');
    }
    if let Some(allow) = value.allow_empty_cell {
        output.push_str(" table:allow-empty-cell=\"");
        output.push_str(if allow { "true" } else { "false" });
        output.push('"');
    }
    if let Some(display) = value.display_list {
        output.push_str(" table:display-list=\"");
        output.push_str(display.as_str());
        output.push('"');
    }
    if value.help_message.is_none() && value.failure.is_none() {
        output.push_str("/>");
        return;
    }
    output.push('>');
    if let Some(help) = &value.help_message {
        write_message(output, "help-message", help, false);
    }
    if let Some(failure) = &value.failure {
        match failure {
            ValidationFailure::Message(message) => {
                write_message(output, "error-message", message, true);
            },
            ValidationFailure::Macro {
                execute,
                event_listeners,
            } => {
                output.push_str("<table:error-macro");
                if let Some(execute) = execute {
                    output.push_str(" table:execute=\"");
                    output.push_str(if *execute { "true" } else { "false" });
                    output.push('"');
                }
                output.push_str("/>");
                if let Some(value) = event_listeners {
                    output.push_str(&value.0);
                }
            },
        }
    }
    output.push_str("</table:content-validation>");
}

fn write_message(output: &mut String, element: &str, message: &ValidationMessage, error: bool) {
    output.push_str("<table:");
    output.push_str(element);
    if let Some(title) = &message.title {
        output.push_str(" table:title=\"");
        escape_attribute(output, title);
        output.push('"');
    }
    if let Some(display) = message.display {
        output.push_str(" table:display=\"");
        output.push_str(if display { "true" } else { "false" });
        output.push('"');
    }
    if error && let Some(kind) = message.message_type {
        output.push_str(" table:message-type=\"");
        output.push_str(kind.as_str());
        output.push('"');
    }
    if message.paragraphs.is_empty() {
        output.push_str("/>");
        return;
    }
    output.push('>');
    for paragraph in &message.paragraphs {
        output.push_str(&paragraph.xml);
    }
    output.push_str("</table:");
    output.push_str(element);
    output.push('>');
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
enum NamespaceKind {
    None,
    Office,
    Table,
    Text,
    Other,
}
#[derive(Clone, Debug)]
struct Frame {
    namespace: NamespaceKind,
    local: String,
}
struct ActiveContainer {
    parent_depth: usize,
    count: usize,
}
struct ActiveValidation {
    parent_depth: usize,
    child_order: u8,
    value: ContentValidation,
}
#[derive(Clone, Copy)]
enum MessageKind {
    Help,
    Error,
}
struct ActiveMessage {
    parent_depth: usize,
    kind: MessageKind,
    value: ValidationMessage,
}
struct ActiveMacro {
    parent_depth: usize,
}
#[derive(Clone, Copy)]
enum CaptureKind {
    Paragraph,
    EventListeners,
}
struct Capture {
    parent_depth: usize,
    kind: CaptureKind,
    writer: Writer<Vec<u8>>,
    text: String,
}
type Attributes = HashMap<(NamespaceKind, String), String>;

pub fn parse_content_validations(xml: &str) -> Result<ContentValidations> {
    parse_part(xml, ContentValidationPart::FlatDocument)
}

pub(super) fn parse_part(xml: &str, part: ContentValidationPart) -> Result<ContentValidations> {
    if xml.len() > MAX_XML_BYTES {
        return invalid("content-validation XML exceeds 64 MiB");
    }
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    let mut stack = Vec::<Frame>::new();
    let mut container: Option<ActiveContainer> = None;
    let mut validation: Option<ActiveValidation> = None;
    let mut message: Option<ActiveMessage> = None;
    let mut active_macro: Option<ActiveMacro> = None;
    let mut capture: Option<Capture> = None;
    let mut seen_container = false;
    let mut result = ContentValidations {
        part,
        validations: Vec::new(),
    };
    loop {
        let (resolved, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| make_error(format!("invalid content-validation XML: {error}")))?;
        let namespace = namespace_kind(&resolved)?;
        if capture.is_some() {
            let mut completed = false;
            match event {
                Event::Start(ref element) => {
                    write_capture(
                        capture.as_mut().expect("checked"),
                        Event::Start(element.to_owned()),
                    )?;
                    let local = decode(element.local_name().as_ref(), "element name")?;
                    stack.push(Frame { namespace, local });
                    if stack.len() > MAX_DEPTH {
                        return invalid(format!(
                            "content-validation XML exceeds depth {MAX_DEPTH}"
                        ));
                    }
                },
                Event::Empty(ref element) => write_capture(
                    capture.as_mut().expect("checked"),
                    Event::Empty(element.to_owned()),
                )?,
                Event::End(ref element) => {
                    write_capture(
                        capture.as_mut().expect("checked"),
                        Event::End(element.to_owned()),
                    )?;
                    let local = decode(element.local_name().as_ref(), "element name")?;
                    let frame = stack
                        .pop()
                        .ok_or_else(|| make_error("unexpected captured XML end element"))?;
                    if frame.namespace != namespace || frame.local != local {
                        return invalid("captured XML end element mismatch");
                    }
                    completed = capture
                        .as_ref()
                        .is_some_and(|value| value.parent_depth == stack.len());
                },
                Event::Text(ref value) => {
                    append_capture_text(
                        capture.as_mut().expect("checked"),
                        &value.decode().map_err(|error| {
                            make_error(format!("invalid validation paragraph text: {error}"))
                        })?,
                    )?;
                    write_capture(
                        capture.as_mut().expect("checked"),
                        Event::Text(value.to_owned()),
                    )?;
                },
                Event::CData(ref value) => {
                    append_capture_text(
                        capture.as_mut().expect("checked"),
                        &value.decode().map_err(|error| {
                            make_error(format!("invalid validation paragraph CDATA: {error}"))
                        })?,
                    )?;
                    write_capture(
                        capture.as_mut().expect("checked"),
                        Event::CData(value.to_owned()),
                    )?;
                },
                Event::GeneralRef(ref value) => {
                    let text = resolve_reference(value)?;
                    append_capture_text(capture.as_mut().expect("checked"), &text)?;
                    write_capture(
                        capture.as_mut().expect("checked"),
                        Event::GeneralRef(value.to_owned()),
                    )?;
                },
                Event::Comment(ref value) => write_capture(
                    capture.as_mut().expect("checked"),
                    Event::Comment(value.to_owned()),
                )?,
                Event::PI(_) | Event::DocType(_) => {
                    return invalid("active XML is not allowed in validation metadata");
                },
                Event::Eof => return invalid("truncated captured validation XML"),
                _ => {},
            }
            if completed {
                finish_capture(
                    capture.take().expect("checked"),
                    message.as_mut(),
                    validation.as_mut(),
                )?;
            }
            buffer.clear();
            continue;
        }
        match event {
            Event::Start(ref element) => {
                let local = decode(element.local_name().as_ref(), "element name")?;
                reject_spoofed_name(namespace, &local)?;
                let relevant = container.is_some()
                    || (namespace == NamespaceKind::Table
                        && matches!(
                            local.as_str(),
                            "content-validations"
                                | "content-validation"
                                | "help-message"
                                | "error-message"
                                | "error-macro"
                        ));
                let attributes = if relevant {
                    read_attributes(&mut reader, element)?
                } else {
                    Attributes::new()
                };
                let was_capturing = capture.is_some();
                handle_start(
                    namespace,
                    &local,
                    attributes,
                    part,
                    stack.len(),
                    stack.last(),
                    &mut container,
                    &mut validation,
                    &mut message,
                    &mut active_macro,
                    &mut capture,
                    seen_container,
                )?;
                if !was_capturing && let Some(active) = capture.as_mut() {
                    write_capture(active, Event::Start(element.to_owned()))?;
                }
                stack.push(Frame { namespace, local });
                if stack.len() > MAX_DEPTH {
                    return invalid(format!("content-validation XML exceeds depth {MAX_DEPTH}"));
                }
            },
            Event::Empty(ref element) => {
                let local = decode(element.local_name().as_ref(), "element name")?;
                reject_spoofed_name(namespace, &local)?;
                let relevant = container.is_some()
                    || (namespace == NamespaceKind::Table
                        && matches!(
                            local.as_str(),
                            "content-validations"
                                | "content-validation"
                                | "help-message"
                                | "error-message"
                                | "error-macro"
                        ));
                let attributes = if relevant {
                    read_attributes(&mut reader, element)?
                } else {
                    Attributes::new()
                };
                handle_empty(
                    namespace,
                    &local,
                    attributes,
                    part,
                    stack.len(),
                    stack.last(),
                    &mut container,
                    &mut validation,
                    &mut message,
                    &mut active_macro,
                    element,
                    &mut result,
                    seen_container,
                )?;
            },
            Event::End(ref element) => {
                let local = decode(element.local_name().as_ref(), "element name")?;
                let frame = stack
                    .pop()
                    .ok_or_else(|| make_error("unexpected content-validation end element"))?;
                if frame.namespace != namespace || frame.local != local {
                    return invalid("content-validation end element mismatch");
                }
                if active_macro
                    .as_ref()
                    .is_some_and(|value| value.parent_depth == stack.len())
                {
                    active_macro = None;
                }
                if message
                    .as_ref()
                    .is_some_and(|value| value.parent_depth == stack.len())
                {
                    let value = message.take().expect("checked");
                    assign_message(
                        validation
                            .as_mut()
                            .ok_or_else(|| make_error("validation message has no parent"))?,
                        value,
                    )?;
                }
                if validation
                    .as_ref()
                    .is_some_and(|value| value.parent_depth == stack.len())
                {
                    result
                        .validations
                        .push(validation.take().expect("checked").value);
                    if result.validations.len() > MAX_VALIDATIONS {
                        return invalid(format!(
                            "document exceeds {MAX_VALIDATIONS} content validations"
                        ));
                    }
                    container
                        .as_mut()
                        .ok_or_else(|| make_error("content validation has no container"))?
                        .count += 1;
                }
                if container
                    .as_ref()
                    .is_some_and(|value| value.parent_depth == stack.len())
                {
                    if container.take().expect("checked").count == 0 {
                        return invalid(
                            "table:content-validations requires at least one validation",
                        );
                    }
                    seen_container = true;
                }
            },
            Event::Text(ref value) if container.is_some() => {
                let value = value.decode().map_err(|error| {
                    make_error(format!("invalid content-validation whitespace: {error}"))
                })?;
                if !value.trim().is_empty() {
                    return invalid("unexpected character data in content-validation structure");
                }
            },
            Event::CData(_) | Event::GeneralRef(_) if container.is_some() => {
                return invalid("unexpected text construct in content-validation structure");
            },
            Event::DocType(_) => return invalid("DTDs are not allowed in content-validation XML"),
            Event::PI(_) => {
                return invalid(
                    "processing instructions are not allowed in content-validation XML",
                );
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if !stack.is_empty()
        || container.is_some()
        || validation.is_some()
        || message.is_some()
        || active_macro.is_some()
        || capture.is_some()
    {
        return invalid("truncated content-validation XML");
    }
    result.validate()?;
    Ok(result)
}

#[allow(clippy::too_many_arguments)]
fn handle_start(
    namespace: NamespaceKind,
    local: &str,
    attributes: Attributes,
    part: ContentValidationPart,
    depth: usize,
    parent: Option<&Frame>,
    container: &mut Option<ActiveContainer>,
    validation: &mut Option<ActiveValidation>,
    message: &mut Option<ActiveMessage>,
    active_macro: &mut Option<ActiveMacro>,
    capture: &mut Option<Capture>,
    seen: bool,
) -> Result<()> {
    if active_macro.is_some() {
        return invalid("table:error-macro must have empty content");
    }
    if let Some(active) = message.as_mut() {
        if namespace != NamespaceKind::Text || local != "p" || depth != active.parent_depth + 1 {
            return invalid("validation messages may contain only direct text:p children");
        }
        *capture = Some(start_capture(
            CaptureKind::Paragraph,
            depth,
            namespace,
            local,
        )?);
        return Ok(());
    }
    if let Some(active) = validation.as_mut() {
        if depth != active.parent_depth + 1 {
            return invalid("content-validation children must be direct and ordered");
        }
        start_validation_child(
            namespace,
            local,
            attributes,
            depth,
            active,
            message,
            active_macro,
            capture,
        )?;
        return Ok(());
    }
    if let Some(active) = container.as_ref() {
        if namespace != NamespaceKind::Table
            || local != "content-validation"
            || depth != active.parent_depth + 1
        {
            return invalid(
                "table:content-validations may contain only direct table:content-validation children",
            );
        }
        *validation = Some(ActiveValidation {
            parent_depth: depth,
            child_order: 0,
            value: parse_validation(attributes)?,
        });
        return Ok(());
    }
    if namespace == NamespaceKind::Table && local == "content-validations" {
        if seen {
            return invalid("a spreadsheet may contain only one table:content-validations element");
        }
        if !attributes.is_empty() {
            return invalid("table:content-validations does not allow attributes");
        }
        if !matches!(parent, Some(Frame { namespace: NamespaceKind::Office, local, .. }) if local == "spreadsheet")
        {
            return invalid(
                "table:content-validations must be a direct child of office:spreadsheet",
            );
        }
        *container = Some(ActiveContainer {
            parent_depth: depth,
            count: 0,
        });
        return Ok(());
    }
    if is_validation_element(namespace, local) {
        return invalid(format!(
            "table:{local} is outside table:content-validations"
        ));
    }
    let _ = part;
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn handle_empty(
    namespace: NamespaceKind,
    local: &str,
    attributes: Attributes,
    part: ContentValidationPart,
    depth: usize,
    parent: Option<&Frame>,
    container: &mut Option<ActiveContainer>,
    validation: &mut Option<ActiveValidation>,
    message: &mut Option<ActiveMessage>,
    active_macro: &mut Option<ActiveMacro>,
    element: &BytesStart<'_>,
    result: &mut ContentValidations,
    seen: bool,
) -> Result<()> {
    if active_macro.is_some() {
        return invalid("table:error-macro must have empty content");
    }
    if let Some(active) = message.as_mut() {
        if namespace != NamespaceKind::Text || local != "p" || depth != active.parent_depth + 1 {
            return invalid("validation messages may contain only direct text:p children");
        }
        let mut capture = start_capture(CaptureKind::Paragraph, depth, namespace, local)?;
        write_capture(&mut capture, Event::Empty(element.to_owned()))?;
        finish_capture(capture, Some(active), validation.as_mut())?;
        return Ok(());
    }
    if let Some(active) = validation.as_mut() {
        if depth != active.parent_depth + 1 {
            return invalid("content-validation children must be direct and ordered");
        }
        if namespace == NamespaceKind::Office && local == "event-listeners" {
            require_macro_events(active)?;
            let mut capture = start_capture(CaptureKind::EventListeners, depth, namespace, local)?;
            write_capture(&mut capture, Event::Empty(element.to_owned()))?;
            finish_capture(capture, None, Some(active))?;
            active.child_order = 3;
            return Ok(());
        }
        let mut local_message = None;
        let mut local_macro = None;
        let mut local_capture = None;
        start_validation_child(
            namespace,
            local,
            attributes,
            depth,
            active,
            &mut local_message,
            &mut local_macro,
            &mut local_capture,
        )?;
        if let Some(value) = local_message {
            assign_message(active, value)?;
        }
        return Ok(());
    }
    if let Some(active) = container.as_mut() {
        if namespace != NamespaceKind::Table
            || local != "content-validation"
            || depth != active.parent_depth + 1
        {
            return invalid(
                "table:content-validations may contain only direct table:content-validation children",
            );
        }
        result.validations.push(parse_validation(attributes)?);
        active.count += 1;
        return Ok(());
    }
    if namespace == NamespaceKind::Table && local == "content-validations" {
        if seen {
            return invalid("a spreadsheet may contain only one table:content-validations element");
        }
        if !attributes.is_empty() {
            return invalid("table:content-validations does not allow attributes");
        }
        if !matches!(parent, Some(Frame { namespace: NamespaceKind::Office, local, .. }) if local == "spreadsheet")
        {
            return invalid(
                "table:content-validations must be a direct child of office:spreadsheet",
            );
        }
        return invalid("table:content-validations requires at least one validation");
    }
    if is_validation_element(namespace, local) {
        return invalid(format!(
            "table:{local} is outside table:content-validations"
        ));
    }
    let _ = (part, active_macro);
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn start_validation_child(
    namespace: NamespaceKind,
    local: &str,
    attributes: Attributes,
    depth: usize,
    active: &mut ActiveValidation,
    message: &mut Option<ActiveMessage>,
    active_macro: &mut Option<ActiveMacro>,
    capture: &mut Option<Capture>,
) -> Result<()> {
    match (namespace, local) {
        (NamespaceKind::Table, "help-message") if active.child_order < 1 => {
            active.child_order = 1;
            *message = Some(ActiveMessage {
                parent_depth: depth,
                kind: MessageKind::Help,
                value: parse_message(attributes, false)?,
            });
        },
        (NamespaceKind::Table, "error-message") if active.child_order < 2 => {
            active.child_order = 2;
            *message = Some(ActiveMessage {
                parent_depth: depth,
                kind: MessageKind::Error,
                value: parse_message(attributes, true)?,
            });
        },
        (NamespaceKind::Table, "error-macro") if active.child_order < 2 => {
            active.child_order = 2;
            let execute = parse_macro(attributes)?;
            active.value.failure = Some(ValidationFailure::Macro {
                execute,
                event_listeners: None,
            });
            *active_macro = Some(ActiveMacro {
                parent_depth: depth,
            });
        },
        (NamespaceKind::Office, "event-listeners") if active.child_order == 2 => {
            require_macro_events(active)?;
            active.child_order = 3;
            *capture = Some(start_capture(
                CaptureKind::EventListeners,
                depth,
                namespace,
                local,
            )?);
        },
        _ => return invalid("invalid, duplicate, or out-of-order content-validation child"),
    }
    Ok(())
}

fn require_macro_events(active: &ActiveValidation) -> Result<()> {
    match &active.value.failure {
        Some(ValidationFailure::Macro {
            event_listeners: None,
            ..
        }) => Ok(()),
        _ => invalid("office:event-listeners must follow exactly one table:error-macro"),
    }
}

fn assign_message(active: &mut ActiveValidation, message: ActiveMessage) -> Result<()> {
    match message.kind {
        MessageKind::Help => {
            if active.value.help_message.replace(message.value).is_some() {
                return invalid("duplicate validation help message");
            }
        },
        MessageKind::Error => {
            if active
                .value
                .failure
                .replace(ValidationFailure::Message(message.value))
                .is_some()
            {
                return invalid("duplicate validation failure action");
            }
        },
    }
    Ok(())
}

fn finish_capture(
    capture: Capture,
    message: Option<&mut ActiveMessage>,
    validation: Option<&mut ActiveValidation>,
) -> Result<()> {
    let xml = String::from_utf8(capture.writer.into_inner())
        .map_err(|error| make_error(format!("captured validation XML is not UTF-8: {error}")))?;
    if xml.len() > MAX_CAPTURE_BYTES {
        return invalid("captured validation XML exceeds 1 MiB");
    }
    match capture.kind {
        CaptureKind::Paragraph => {
            let message = message
                .ok_or_else(|| make_error("captured paragraph has no validation message"))?;
            if message.value.paragraphs.len() >= MAX_PARAGRAPHS {
                return invalid(format!(
                    "validation message exceeds {MAX_PARAGRAPHS} paragraphs"
                ));
            }
            message.value.paragraphs.push(ValidationParagraph {
                xml,
                text: capture.text,
            });
        },
        CaptureKind::EventListeners => {
            let validation = validation
                .ok_or_else(|| make_error("captured event listeners have no validation"))?;
            match validation.value.failure.as_mut() {
                Some(ValidationFailure::Macro {
                    event_listeners, ..
                }) if event_listeners.is_none() => {
                    *event_listeners = Some(ValidationEventListeners(xml));
                },
                _ => return invalid("captured event listeners do not follow a macro"),
            }
        },
    }
    Ok(())
}

fn start_capture(
    kind: CaptureKind,
    parent_depth: usize,
    _namespace: NamespaceKind,
    _local: &str,
) -> Result<Capture> {
    Ok(Capture {
        parent_depth,
        kind,
        writer: Writer::new(Vec::new()),
        text: String::new(),
    })
}

fn write_capture(capture: &mut Capture, event: Event<'_>) -> Result<()> {
    capture
        .writer
        .write_event(event)
        .map_err(|error| make_error(format!("cannot preserve validation XML: {error}")))?;
    if capture.writer.get_ref().len() > MAX_CAPTURE_BYTES {
        return invalid("captured validation XML exceeds 1 MiB");
    }
    Ok(())
}

fn append_capture_text(capture: &mut Capture, value: &str) -> Result<()> {
    if matches!(capture.kind, CaptureKind::Paragraph) {
        if capture.text.len().saturating_add(value.len()) > MAX_VALUE_BYTES {
            return invalid("validation paragraph text exceeds 64 KiB");
        }
        capture.text.push_str(value);
    }
    Ok(())
}

fn parse_validation(mut attributes: Attributes) -> Result<ContentValidation> {
    let name = required(&mut attributes, "name")?;
    validate_text(&name, "table:name", false)?;
    let condition = attributes
        .remove(&(NamespaceKind::Table, "condition".to_owned()))
        .map(ValidationCondition::new)
        .transpose()?;
    let base_cell_address = attributes
        .remove(&(NamespaceKind::Table, "base-cell-address".to_owned()))
        .map(ValidationCellAddress::new)
        .transpose()?;
    let allow_empty_cell = attributes
        .remove(&(NamespaceKind::Table, "allow-empty-cell".to_owned()))
        .map(|value| parse_bool(&value))
        .transpose()?;
    let display_list = attributes
        .remove(&(NamespaceKind::Table, "display-list".to_owned()))
        .map(|value| ValidationDisplayList::parse(&value))
        .transpose()?;
    reject_remaining(attributes, "content-validation")?;
    Ok(ContentValidation {
        name,
        condition,
        base_cell_address,
        allow_empty_cell,
        display_list,
        help_message: None,
        failure: None,
    })
}

fn parse_message(mut attributes: Attributes, error: bool) -> Result<ValidationMessage> {
    let title = attributes.remove(&(NamespaceKind::Table, "title".to_owned()));
    let display = attributes
        .remove(&(NamespaceKind::Table, "display".to_owned()))
        .map(|value| parse_bool(&value))
        .transpose()?;
    let message_type = attributes
        .remove(&(NamespaceKind::Table, "message-type".to_owned()))
        .map(|value| ValidationMessageType::parse(&value))
        .transpose()?;
    if !error && message_type.is_some() {
        return invalid("table:help-message cannot have table:message-type");
    }
    reject_remaining(
        attributes,
        if error {
            "error-message"
        } else {
            "help-message"
        },
    )?;
    Ok(ValidationMessage {
        title,
        display,
        message_type,
        paragraphs: Vec::new(),
    })
}

fn parse_macro(mut attributes: Attributes) -> Result<Option<bool>> {
    let execute = attributes
        .remove(&(NamespaceKind::Table, "execute".to_owned()))
        .map(|value| parse_bool(&value))
        .transpose()?;
    reject_remaining(attributes, "error-macro")?;
    Ok(execute)
}

fn required(attributes: &mut Attributes, local: &str) -> Result<String> {
    attributes
        .remove(&(NamespaceKind::Table, local.to_owned()))
        .ok_or_else(|| make_error(format!("content validation requires table:{local}")))
}
fn reject_remaining(attributes: Attributes, context: &str) -> Result<()> {
    if let Some(((namespace, local), _)) = attributes.into_iter().next() {
        return invalid(format!(
            "unsupported {namespace:?} {context} attribute '{local}'"
        ));
    }
    Ok(())
}
fn is_validation_element(namespace: NamespaceKind, local: &str) -> bool {
    (namespace == NamespaceKind::Table
        && matches!(
            local,
            "content-validations"
                | "content-validation"
                | "help-message"
                | "error-message"
                | "error-macro"
        ))
        || (namespace == NamespaceKind::Office && local == "event-listeners")
}

fn read_attributes<R: BufRead>(
    reader: &mut NsReader<R>,
    element: &BytesStart<'_>,
) -> Result<Attributes> {
    let mut result = HashMap::new();
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| {
            make_error(format!("invalid content-validation attribute: {error}"))
        })?;
        let raw = attribute.key.as_ref();
        if raw == b"xmlns" || raw.starts_with(b"xmlns:") {
            continue;
        }
        let (resolved, local) = reader.resolver_mut().resolve_attribute(attribute.key);
        let namespace = namespace_kind(&resolved)?;
        let local = decode(local.as_ref(), "attribute name")?;
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| {
                make_error(format!(
                    "invalid content-validation attribute value: {error}"
                ))
            })?
            .into_owned();
        validate_text(&value, &local, true)?;
        if result.insert((namespace, local.clone()), value).is_some() {
            return invalid(format!("duplicate content-validation attribute '{local}'"));
        }
    }
    Ok(result)
}

fn namespace_kind(resolved: &ResolveResult<'_>) -> Result<NamespaceKind> {
    match resolved {
        ResolveResult::Unbound => Ok(NamespaceKind::None),
        ResolveResult::Bound(namespace) => match namespace.as_ref() {
            OFFICE_NS => Ok(NamespaceKind::Office),
            TABLE_NS => Ok(NamespaceKind::Table),
            TEXT_NS => Ok(NamespaceKind::Text),
            _ => Ok(NamespaceKind::Other),
        },
        ResolveResult::Unknown(prefix) => invalid(format!(
            "unknown XML namespace prefix '{}'",
            String::from_utf8_lossy(prefix.as_ref())
        )),
    }
}

fn reject_spoofed_name(namespace: NamespaceKind, local: &str) -> Result<()> {
    if matches!(
        local,
        "content-validations"
            | "content-validation"
            | "help-message"
            | "error-message"
            | "error-macro"
    ) && namespace != NamespaceKind::Table
    {
        return invalid(format!("spoofed table:{local} element namespace"));
    }
    Ok(())
}

fn resolve_reference(value: &quick_xml::events::BytesRef<'_>) -> Result<String> {
    if let Some(character) = value
        .resolve_char_ref()
        .map_err(|error| make_error(format!("invalid character reference: {error}")))?
    {
        return Ok(character.to_string());
    }
    match value
        .decode()
        .map_err(|error| make_error(format!("invalid entity reference: {error}")))?
        .as_ref()
    {
        "amp" => Ok("&".into()),
        "lt" => Ok("<".into()),
        "gt" => Ok(">".into()),
        "apos" => Ok("'".into()),
        "quot" => Ok("\"".into()),
        name => invalid(format!("unsupported entity reference '&{name};'")),
    }
}

fn parse_bool(value: &str) -> Result<bool> {
    match value {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        _ => invalid(format!("invalid ODF boolean '{value}'")),
    }
}
fn decode(value: &[u8], name: &str) -> Result<String> {
    std::str::from_utf8(value)
        .map(str::to_owned)
        .map_err(|error| make_error(format!("invalid UTF-8 {name}: {error}")))
}

fn escape_attribute(output: &mut String, value: &str) {
    for character in value.chars() {
        match character {
            '&' => output.push_str("&amp;"),
            '<' => output.push_str("&lt;"),
            '"' => output.push_str("&quot;"),
            '\r' => output.push_str("&#13;"),
            '\n' => output.push_str("&#10;"),
            '\t' => output.push_str("&#9;"),
            _ => output.push(character),
        }
    }
}
