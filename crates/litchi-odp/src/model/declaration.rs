//! Inert ODF presentation header, footer, and date-time declarations.

use litchi_core::{Error, Result, xml::escape_xml};
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::{Namespace, ResolveResult},
    reader::NsReader,
};
use std::collections::HashSet;

const OFFICE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const PRESENTATION_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:presentation:1.0";
const DRAW_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const STYLE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const MAX_XML_BYTES: usize = 8 * 1024 * 1024;
const MAX_TEXT_BYTES: usize = 1024 * 1024;
const MAX_DECLARATIONS: usize = 65_536;
const MAX_BINDINGS: usize = 131_072;

/// A named static header or footer declaration.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PresentationTextDeclaration {
    pub name: String,
    pub text: String,
}

impl PresentationTextDeclaration {
    /// Create a validated header/footer declaration.
    pub fn new(name: impl Into<String>, text: impl Into<String>) -> Result<Self> {
        let value = Self {
            name: name.into(),
            text: text.into(),
        };
        validate_name(&value.name, "presentation declaration name")?;
        validate_text(&value.text, "presentation declaration text", true)?;
        Ok(value)
    }
}

/// Source behavior for a presentation date/time declaration.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PresentationDateTimeSource {
    Fixed,
    CurrentDate,
}

impl PresentationDateTimeSource {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "fixed" => Ok(Self::Fixed),
            "current-date" => Ok(Self::CurrentDate),
            _ => Err(invalid(format!(
                "unsupported presentation:source '{value}'"
            ))),
        }
    }

    fn as_str(self) -> &'static str {
        match self {
            Self::Fixed => "fixed",
            Self::CurrentDate => "current-date",
        }
    }
}

/// A named presentation date/time declaration.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PresentationDateTimeDeclaration {
    pub name: String,
    /// Omission is retained independently from either schema value.
    pub source: Option<PresentationDateTimeSource>,
    pub data_style_name: Option<String>,
    pub text: String,
}

impl PresentationDateTimeDeclaration {
    /// Create a validated date/time declaration.
    pub fn new(name: impl Into<String>) -> Result<Self> {
        let value = Self {
            name: name.into(),
            source: None,
            data_style_name: None,
            text: String::new(),
        };
        validate_date_time(&value)?;
        Ok(value)
    }
}

/// Whether declaration references belong to a slide or its notes page.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum PresentationDeclarationTarget {
    Slide,
    Notes,
}

/// Header/footer/date-time references attached to one slide or notes page.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PresentationDeclarationBinding {
    pub slide_index: usize,
    pub target: PresentationDeclarationTarget,
    pub header_name: Option<String>,
    pub footer_name: Option<String>,
    pub date_time_name: Option<String>,
}

impl PresentationDeclarationBinding {
    /// Create an empty binding for a zero-based slide index.
    pub fn new(slide_index: usize, target: PresentationDeclarationTarget) -> Self {
        Self {
            slide_index,
            target,
            header_name: None,
            footer_name: None,
            date_time_name: None,
        }
    }

    fn is_empty(&self) -> bool {
        self.header_name.is_none() && self.footer_name.is_none() && self.date_time_name.is_none()
    }
}

/// Complete declaration and page-binding metadata in document order.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct PresentationDeclarations {
    pub headers: Vec<PresentationTextDeclaration>,
    pub footers: Vec<PresentationTextDeclaration>,
    pub date_times: Vec<PresentationDateTimeDeclaration>,
    pub bindings: Vec<PresentationDeclarationBinding>,
}

impl PresentationDeclarations {
    /// Validate declaration uniqueness and reference integrity.
    pub fn validate(&self) -> Result<()> {
        self.validate_for_slide_count(None)
    }

    /// Validate declaration metadata against a concrete slide count.
    pub fn validate_for_slides(&self, slide_count: usize) -> Result<()> {
        self.validate_for_slide_count(Some(slide_count))
    }

    fn validate_for_slide_count(&self, slide_count: Option<usize>) -> Result<()> {
        if self.headers.len() > MAX_DECLARATIONS
            || self.footers.len() > MAX_DECLARATIONS
            || self.date_times.len() > MAX_DECLARATIONS
        {
            return Err(invalid(
                "presentation declaration collection exceeds 65536 entries",
            ));
        }
        if self.bindings.len() > MAX_BINDINGS {
            return Err(invalid(
                "presentation declaration bindings exceed 131072 entries",
            ));
        }
        let headers = validate_text_declarations(&self.headers, "header")?;
        let footers = validate_text_declarations(&self.footers, "footer")?;
        let mut date_times = HashSet::with_capacity(self.date_times.len());
        for value in &self.date_times {
            validate_date_time(value)?;
            if !date_times.insert(value.name.as_str()) {
                return Err(invalid(format!(
                    "duplicate presentation date-time declaration '{}'",
                    value.name
                )));
            }
        }
        let mut targets = HashSet::with_capacity(self.bindings.len());
        for binding in &self.bindings {
            if binding.is_empty() {
                return Err(invalid("empty presentation declaration binding"));
            }
            if let Some(count) = slide_count
                && binding.slide_index >= count
            {
                return Err(invalid(format!(
                    "presentation declaration binding slide index {} exceeds slide count {count}",
                    binding.slide_index
                )));
            }
            if !targets.insert((binding.slide_index, binding.target)) {
                return Err(invalid(format!(
                    "duplicate {:?} declaration binding for slide {}",
                    binding.target, binding.slide_index
                )));
            }
            validate_reference(binding.header_name.as_deref(), "header", &headers)?;
            validate_reference(binding.footer_name.as_deref(), "footer", &footers)?;
            validate_reference(binding.date_time_name.as_deref(), "date-time", &date_times)?;
        }
        Ok(())
    }

    /// Find a binding for a slide or its notes page.
    pub fn binding(
        &self,
        slide_index: usize,
        target: PresentationDeclarationTarget,
    ) -> Option<&PresentationDeclarationBinding> {
        self.bindings
            .iter()
            .find(|value| value.slide_index == slide_index && value.target == target)
    }

    /// Return whether no declarations or bindings are present.
    pub fn is_empty(&self) -> bool {
        self.headers.is_empty()
            && self.footers.is_empty()
            && self.date_times.is_empty()
            && self.bindings.is_empty()
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum DeclarationKind {
    Header,
    Footer,
    DateTime,
}

struct OpenDeclaration {
    depth: usize,
    kind: DeclarationKind,
    name: String,
    source: Option<PresentationDateTimeSource>,
    data_style_name: Option<String>,
    text: String,
}

/// Parse all presentation declarations and their slide/notes bindings.
pub fn parse_presentation_declarations(xml: &str) -> Result<PresentationDeclarations> {
    if xml.len() > MAX_XML_BYTES {
        return Err(invalid("presentation declaration XML exceeds 8 MiB"));
    }
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut presentation_depth = None;
    let mut page_depth = None;
    let mut notes_depth = None;
    let mut page_count = 0usize;
    let mut found_presentation = false;
    let mut open_declaration: Option<OpenDeclaration> = None;
    let mut result = PresentationDeclarations::default();

    loop {
        match reader.read_event_into(&mut buffer).map_err(xml_error)? {
            Event::Start(element) => {
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("XML nesting overflow"))?;
                if element_is(&reader, &element, OFFICE_NAMESPACE, b"presentation") {
                    if found_presentation {
                        return Err(invalid("duplicate office:presentation element"));
                    }
                    found_presentation = true;
                    presentation_depth = Some(depth);
                } else if let Some(kind) = declaration_kind(&reader, &element) {
                    if presentation_depth != Some(depth - 1) || open_declaration.is_some() {
                        return Err(invalid(
                            "presentation declarations must be direct office:presentation children",
                        ));
                    }
                    open_declaration =
                        Some(parse_declaration_start(&reader, &element, kind, depth)?);
                } else if open_declaration.is_some() {
                    return Err(invalid("presentation declarations may contain text only"));
                } else if element_is(&reader, &element, DRAW_NAMESPACE, b"page")
                    && presentation_depth == Some(depth - 1)
                {
                    let slide_index = page_count;
                    page_count = page_count
                        .checked_add(1)
                        .ok_or_else(|| invalid("slide count overflow"))?;
                    if let Some(binding) = parse_binding(
                        &reader,
                        &element,
                        slide_index,
                        PresentationDeclarationTarget::Slide,
                    )? {
                        result.bindings.push(binding);
                    }
                    page_depth = Some(depth);
                } else if element_is(&reader, &element, PRESENTATION_NAMESPACE, b"notes")
                    && page_depth.is_some()
                {
                    if notes_depth.is_some() {
                        return Err(invalid("nested presentation:notes element"));
                    }
                    if let Some(binding) = parse_binding(
                        &reader,
                        &element,
                        page_count - 1,
                        PresentationDeclarationTarget::Notes,
                    )? {
                        result.bindings.push(binding);
                    }
                    notes_depth = Some(depth);
                }
            },
            Event::Empty(element) => {
                if let Some(kind) = declaration_kind(&reader, &element) {
                    if presentation_depth != Some(depth) || open_declaration.is_some() {
                        return Err(invalid(
                            "presentation declarations must be direct office:presentation children",
                        ));
                    }
                    let value = parse_declaration_start(&reader, &element, kind, depth + 1)?;
                    finish_declaration(value, &mut result)?;
                } else if element_is(&reader, &element, PRESENTATION_NAMESPACE, b"notes")
                    && page_depth.is_some()
                {
                    if let Some(binding) = parse_binding(
                        &reader,
                        &element,
                        page_count - 1,
                        PresentationDeclarationTarget::Notes,
                    )? {
                        result.bindings.push(binding);
                    }
                } else if element_is(&reader, &element, DRAW_NAMESPACE, b"page")
                    && presentation_depth == Some(depth)
                {
                    let slide_index = page_count;
                    page_count = page_count
                        .checked_add(1)
                        .ok_or_else(|| invalid("slide count overflow"))?;
                    if let Some(binding) = parse_binding(
                        &reader,
                        &element,
                        slide_index,
                        PresentationDeclarationTarget::Slide,
                    )? {
                        result.bindings.push(binding);
                    }
                } else if open_declaration.is_some() {
                    return Err(invalid("presentation declarations may contain text only"));
                }
            },
            Event::Text(text) if open_declaration.is_some() => {
                let decoded = text
                    .xml_content(XmlVersion::Implicit1_0)
                    .map_err(xml_error)?;
                let declaration = open_declaration.as_mut().expect("checked above");
                if declaration.text.len().saturating_add(decoded.len()) > MAX_TEXT_BYTES {
                    return Err(invalid("presentation declaration text exceeds 1 MiB"));
                }
                declaration.text.push_str(&decoded);
            },
            Event::GeneralRef(reference) if open_declaration.is_some() => {
                let name: &[u8] = reference.as_ref();
                let replacement = match name {
                    b"amp" => "&",
                    b"lt" => "<",
                    b"gt" => ">",
                    b"apos" => "'",
                    b"quot" => "\"",
                    _ => {
                        return Err(invalid(
                            "unsupported entity in presentation declaration text",
                        ));
                    },
                };
                let declaration = open_declaration.as_mut().expect("checked above");
                declaration.text.push_str(replacement);
            },
            Event::CData(text) if open_declaration.is_some() => {
                let decoded = reader.decoder().decode(text.as_ref()).map_err(xml_error)?;
                let declaration = open_declaration.as_mut().expect("checked above");
                if declaration.text.len().saturating_add(decoded.len()) > MAX_TEXT_BYTES {
                    return Err(invalid("presentation declaration text exceeds 1 MiB"));
                }
                declaration.text.push_str(&decoded);
            },
            Event::End(element) => {
                if open_declaration
                    .as_ref()
                    .is_some_and(|value| value.depth == depth)
                {
                    let declaration = open_declaration.take().expect("checked above");
                    let expected = match declaration.kind {
                        DeclarationKind::Header => b"header-decl".as_slice(),
                        DeclarationKind::Footer => b"footer-decl".as_slice(),
                        DeclarationKind::DateTime => b"date-time-decl".as_slice(),
                    };
                    if !end_is(&reader, &element, PRESENTATION_NAMESPACE, expected) {
                        return Err(invalid("unexpected presentation declaration end element"));
                    }
                    finish_declaration(declaration, &mut result)?;
                } else if notes_depth == Some(depth)
                    && end_is(&reader, &element, PRESENTATION_NAMESPACE, b"notes")
                {
                    notes_depth = None;
                } else if page_depth == Some(depth)
                    && end_is(&reader, &element, DRAW_NAMESPACE, b"page")
                {
                    page_depth = None;
                    notes_depth = None;
                } else if presentation_depth == Some(depth)
                    && end_is(&reader, &element, OFFICE_NAMESPACE, b"presentation")
                {
                    presentation_depth = None;
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("unbalanced XML end element"))?;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("active XML declarations are prohibited"));
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if open_declaration.is_some() {
        return Err(invalid("unterminated presentation declaration"));
    }
    result.validate_for_slides(page_count)?;
    Ok(result)
}

pub(crate) fn write_declaration_elements(
    declarations: Option<&PresentationDeclarations>,
    slide_count: usize,
) -> Result<String> {
    let Some(declarations) = declarations else {
        return Ok(String::new());
    };
    declarations.validate_for_slides(slide_count)?;
    let mut output = String::with_capacity(256);
    for value in &declarations.headers {
        write_text_declaration(&mut output, "header-decl", value);
    }
    for value in &declarations.footers {
        write_text_declaration(&mut output, "footer-decl", value);
    }
    for value in &declarations.date_times {
        output.push_str("<presentation:date-time-decl presentation:name=\"");
        output.push_str(&escape_xml(&value.name));
        output.push('"');
        if let Some(source) = value.source {
            write_attribute(&mut output, "presentation:source", source.as_str());
        }
        if let Some(style) = &value.data_style_name {
            write_attribute(&mut output, "style:data-style-name", style);
        }
        if value.text.is_empty() {
            output.push_str("/>");
        } else {
            output.push('>');
            output.push_str(&escape_xml(&value.text));
            output.push_str("</presentation:date-time-decl>");
        }
    }
    Ok(output)
}

pub(crate) fn write_binding_attributes(
    declarations: Option<&PresentationDeclarations>,
    slide_index: usize,
    target: PresentationDeclarationTarget,
) -> String {
    let Some(binding) = declarations.and_then(|value| value.binding(slide_index, target)) else {
        return String::new();
    };
    let mut output = String::new();
    if let Some(value) = &binding.header_name {
        write_attribute(&mut output, "presentation:use-header-name", value);
    }
    if let Some(value) = &binding.footer_name {
        write_attribute(&mut output, "presentation:use-footer-name", value);
    }
    if let Some(value) = &binding.date_time_name {
        write_attribute(&mut output, "presentation:use-date-time-name", value);
    }
    output
}

pub(crate) fn apply_notes_binding(base: String, attributes: &str) -> Result<String> {
    if attributes.is_empty() {
        return Ok(base);
    }
    if base.is_empty() {
        return Ok(format!("<presentation:notes{attributes}/>"));
    }
    const PREFIX: &str = "<presentation:notes";
    if !base.starts_with(PREFIX) {
        return Err(invalid("unexpected generated presentation notes fragment"));
    }
    let mut output = String::with_capacity(base.len() + attributes.len());
    output.push_str(PREFIX);
    output.push_str(attributes);
    output.push_str(&base[PREFIX.len()..]);
    Ok(output)
}

fn declaration_kind(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Option<DeclarationKind> {
    if element_is(reader, element, PRESENTATION_NAMESPACE, b"header-decl") {
        Some(DeclarationKind::Header)
    } else if element_is(reader, element, PRESENTATION_NAMESPACE, b"footer-decl") {
        Some(DeclarationKind::Footer)
    } else if element_is(reader, element, PRESENTATION_NAMESPACE, b"date-time-decl") {
        Some(DeclarationKind::DateTime)
    } else {
        None
    }
}

fn parse_declaration_start(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    kind: DeclarationKind,
    depth: usize,
) -> Result<OpenDeclaration> {
    let mut name = None;
    let mut source = None;
    let mut data_style_name = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(xml_error)?;
        if is_namespace_declaration(attribute.key.as_ref()) {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(xml_error)?
            .into_owned();
        match (namespace, local.as_ref()) {
            (ResolveResult::Bound(found), b"name")
                if found == Namespace(PRESENTATION_NAMESPACE) && name.is_none() =>
            {
                name = Some(value)
            },
            (ResolveResult::Bound(found), b"source")
                if found == Namespace(PRESENTATION_NAMESPACE)
                    && kind == DeclarationKind::DateTime
                    && source.is_none() =>
            {
                source = Some(PresentationDateTimeSource::parse(&value)?)
            },
            (ResolveResult::Bound(found), b"data-style-name")
                if found == Namespace(STYLE_NAMESPACE)
                    && kind == DeclarationKind::DateTime
                    && data_style_name.is_none() =>
            {
                validate_name_reference(&value, "style:data-style-name")?;
                data_style_name = Some(value);
            },
            _ => {
                return Err(invalid(
                    "unsupported or duplicate presentation declaration attribute",
                ));
            },
        }
    }
    let name =
        name.ok_or_else(|| invalid("presentation declaration requires presentation:name"))?;
    validate_name(&name, "presentation declaration name")?;
    Ok(OpenDeclaration {
        depth,
        kind,
        name,
        source,
        data_style_name,
        text: String::new(),
    })
}

fn parse_binding(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    slide_index: usize,
    target: PresentationDeclarationTarget,
) -> Result<Option<PresentationDeclarationBinding>> {
    let mut value = PresentationDeclarationBinding::new(slide_index, target);
    for attribute in element.attributes() {
        let attribute = attribute.map_err(xml_error)?;
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        if !matches!(namespace, ResolveResult::Bound(found) if found == Namespace(PRESENTATION_NAMESPACE))
        {
            continue;
        }
        let decoded = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(xml_error)?
            .into_owned();
        match local.as_ref() {
            b"use-header-name" if value.header_name.is_none() => value.header_name = Some(decoded),
            b"use-footer-name" if value.footer_name.is_none() => value.footer_name = Some(decoded),
            b"use-date-time-name" if value.date_time_name.is_none() => {
                value.date_time_name = Some(decoded)
            },
            b"use-header-name" | b"use-footer-name" | b"use-date-time-name" => {
                return Err(invalid(
                    "duplicate presentation declaration binding attribute",
                ));
            },
            _ => continue,
        }
    }
    for (name, description) in [
        (value.header_name.as_deref(), "presentation:use-header-name"),
        (value.footer_name.as_deref(), "presentation:use-footer-name"),
        (
            value.date_time_name.as_deref(),
            "presentation:use-date-time-name",
        ),
    ] {
        if let Some(name) = name {
            validate_name(name, description)?;
        }
    }
    Ok((!value.is_empty()).then_some(value))
}

fn finish_declaration(value: OpenDeclaration, result: &mut PresentationDeclarations) -> Result<()> {
    match value.kind {
        DeclarationKind::Header => result
            .headers
            .push(PresentationTextDeclaration::new(value.name, value.text)?),
        DeclarationKind::Footer => result
            .footers
            .push(PresentationTextDeclaration::new(value.name, value.text)?),
        DeclarationKind::DateTime => {
            let value = PresentationDateTimeDeclaration {
                name: value.name,
                source: value.source,
                data_style_name: value.data_style_name,
                text: value.text,
            };
            validate_date_time(&value)?;
            result.date_times.push(value);
        },
    }
    Ok(())
}

fn validate_text_declarations<'a>(
    values: &'a [PresentationTextDeclaration],
    kind: &str,
) -> Result<HashSet<&'a str>> {
    let mut names = HashSet::with_capacity(values.len());
    for value in values {
        validate_name(&value.name, "presentation declaration name")?;
        validate_text(&value.text, "presentation declaration text", true)?;
        if !names.insert(value.name.as_str()) {
            return Err(invalid(format!(
                "duplicate presentation {kind} declaration '{}'",
                value.name
            )));
        }
    }
    Ok(names)
}

fn validate_date_time(value: &PresentationDateTimeDeclaration) -> Result<()> {
    validate_name(&value.name, "presentation date-time declaration name")?;
    validate_text(&value.text, "presentation date-time declaration text", true)?;
    if let Some(style) = &value.data_style_name {
        validate_name_reference(style, "style:data-style-name")?;
    }
    Ok(())
}

fn validate_reference(name: Option<&str>, kind: &str, names: &HashSet<&str>) -> Result<()> {
    let Some(name) = name else {
        return Ok(());
    };
    validate_name(name, "presentation declaration reference")?;
    if !names.contains(name) {
        return Err(invalid(format!(
            "presentation binding references missing {kind} declaration '{name}'"
        )));
    }
    Ok(())
}

/// Validate a `styleNameRef` attribute value.
///
/// ODF types style references as `styleNameRef`, which is either an `NCName` or
/// the empty string meaning "no referenced style". Empty values are therefore
/// preserved verbatim instead of being rejected.
fn validate_name_reference(value: &str, description: &str) -> Result<()> {
    if value.is_empty() {
        return Ok(());
    }
    validate_name(value, description)
}

fn validate_name(value: &str, description: &str) -> Result<()> {
    validate_text(value, description, false)?;
    if value.chars().any(char::is_whitespace) {
        return Err(invalid(format!("{description} cannot contain whitespace")));
    }
    Ok(())
}

fn validate_text(value: &str, description: &str, allow_empty: bool) -> Result<()> {
    if value.len() > MAX_TEXT_BYTES {
        return Err(invalid(format!("{description} exceeds 1 MiB")));
    }
    if !allow_empty && value.is_empty() {
        return Err(invalid(format!("{description} cannot be empty")));
    }
    if value.chars().any(|character| {
        character == '\0' || (character.is_control() && !matches!(character, '\t' | '\n' | '\r'))
    }) {
        return Err(invalid(format!(
            "{description} contains invalid XML characters"
        )));
    }
    Ok(())
}

fn write_text_declaration(output: &mut String, element: &str, value: &PresentationTextDeclaration) {
    output.push_str("<presentation:");
    output.push_str(element);
    output.push_str(" presentation:name=\"");
    output.push_str(&escape_xml(&value.name));
    if value.text.is_empty() {
        output.push_str("\"/>");
    } else {
        output.push_str("\">");
        output.push_str(&escape_xml(&value.text));
        output.push_str("</presentation:");
        output.push_str(element);
        output.push('>');
    }
}

fn write_attribute(output: &mut String, name: &str, value: &str) {
    output.push(' ');
    output.push_str(name);
    output.push_str("=\"");
    output.push_str(&escape_xml(value));
    output.push('"');
}

fn element_is(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    namespace: &[u8],
    local: &[u8],
) -> bool {
    let (resolved, local_name) = reader.resolver().resolve_element(element.name());
    matches!(resolved, ResolveResult::Bound(found) if found == Namespace(namespace))
        && local_name.as_ref() == local
}

fn end_is(
    reader: &NsReader<&[u8]>,
    element: &quick_xml::events::BytesEnd<'_>,
    namespace: &[u8],
    local: &[u8],
) -> bool {
    let (resolved, local_name) = reader.resolver().resolve_element(element.name());
    matches!(resolved, ResolveResult::Bound(found) if found == Namespace(namespace))
        && local_name.as_ref() == local
}

fn is_namespace_declaration(name: &[u8]) -> bool {
    name == b"xmlns" || name.starts_with(b"xmlns:")
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}

fn xml_error(error: impl std::fmt::Display) -> Error {
    invalid(format!(
        "presentation declaration XML parsing error: {error}"
    ))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{Presentation, PresentationBuilder};

    const PREFIX: &str = r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:p="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0"><o:body><o:presentation>"#;
    const SUFFIX: &str = "</o:presentation></o:body></o:document-content>";

    fn declarations() -> PresentationDeclarations {
        PresentationDeclarations {
            headers: vec![PresentationTextDeclaration::new("h1", "Quarterly & Review").unwrap()],
            footers: vec![PresentationTextDeclaration::new("f1", "Confidential").unwrap()],
            date_times: vec![PresentationDateTimeDeclaration {
                name: "d1".to_string(),
                source: Some(PresentationDateTimeSource::CurrentDate),
                data_style_name: Some("N2".to_string()),
                text: String::new(),
            }],
            bindings: vec![
                PresentationDeclarationBinding {
                    slide_index: 0,
                    target: PresentationDeclarationTarget::Slide,
                    header_name: Some("h1".to_string()),
                    footer_name: Some("f1".to_string()),
                    date_time_name: Some("d1".to_string()),
                },
                PresentationDeclarationBinding {
                    slide_index: 0,
                    target: PresentationDeclarationTarget::Notes,
                    header_name: None,
                    footer_name: Some("f1".to_string()),
                    date_time_name: None,
                },
            ],
        }
    }

    #[test]
    fn parses_all_declarations_and_page_bindings() {
        let xml = format!(
            r#"{PREFIX}<p:header-decl p:name="h1">Quarterly &amp; Review</p:header-decl><p:footer-decl p:name="f1">Confidential</p:footer-decl><p:date-time-decl p:name="d1" p:source="current-date" s:data-style-name="N2"/><d:page p:use-header-name="h1" p:use-footer-name="f1" p:use-date-time-name="d1"><p:notes p:use-footer-name="f1"/></d:page>{SUFFIX}"#
        );
        assert_eq!(
            parse_presentation_declarations(&xml).unwrap(),
            declarations()
        );
    }

    #[test]
    fn builder_round_trips_declarations() {
        let declarations = declarations();
        let mut builder = PresentationBuilder::new();
        builder.add_slide_with_title("Title", "Body").unwrap();
        builder
            .set_declarations(Some(declarations.clone()))
            .unwrap();
        let presentation = Presentation::from_bytes(builder.build().unwrap()).unwrap();
        assert_eq!(presentation.declarations().unwrap(), declarations);
    }

    #[test]
    fn rejects_duplicates_orphans_bad_values_and_active_xml() {
        for body in [
            r#"<p:header-decl p:name="h"/><p:header-decl p:name="h"/>"#,
            r#"<p:date-time-decl p:name="d" p:source="clock"/>"#,
            r#"<d:page p:use-header-name="missing"/>"#,
            r#"<p:header-decl p:name="h"><p:footer-decl p:name="f"/></p:header-decl>"#,
            r#"<p:header-decl p:name="bad name"/>"#,
        ] {
            let xml = format!("{PREFIX}{body}{SUFFIX}");
            assert!(
                parse_presentation_declarations(&xml).is_err(),
                "accepted {xml}"
            );
        }
        let active = format!("{PREFIX}<!DOCTYPE x><p:header-decl p:name=\"h\"/>{SUFFIX}");
        assert!(parse_presentation_declarations(&active).is_err());
    }
}
