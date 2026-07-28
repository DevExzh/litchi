//! ODF text line-numbering configuration.

use std::collections::HashSet;

use litchi_core::{Error, Result};
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::ResolveResult,
    reader::NsReader,
};

use crate::{FlatOpenDocument, OpenDocumentPackage};

const OFFICE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TEXT_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const STYLE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const MAX_DOCUMENT_XML_BYTES: usize = 32 * 1024 * 1024;
const MAX_VALUE_BYTES: usize = 64 * 1024;
const MAX_XML_DEPTH: usize = 128;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
enum NamespaceKind {
    Office,
    Text,
    Style,
    Other,
}

/// Numbering format for line numbers.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum OdfLineNumberFormat {
    Empty,
    Arabic,
    LowerRoman,
    UpperRoman,
    LowerAlpha,
    UpperAlpha,
    Custom(String),
}

impl OdfLineNumberFormat {
    pub fn parse(value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        if value.len() > MAX_VALUE_BYTES {
            return invalid(format!(
                "style:num-format exceeds the {MAX_VALUE_BYTES} byte limit"
            ));
        }
        Ok(match value.as_str() {
            "" => Self::Empty,
            "1" => Self::Arabic,
            "i" => Self::LowerRoman,
            "I" => Self::UpperRoman,
            "a" => Self::LowerAlpha,
            "A" => Self::UpperAlpha,
            _ => Self::Custom(value),
        })
    }

    pub fn as_str(&self) -> &str {
        match self {
            Self::Empty => "",
            Self::Arabic => "1",
            Self::LowerRoman => "i",
            Self::UpperRoman => "I",
            Self::LowerAlpha => "a",
            Self::UpperAlpha => "A",
            Self::Custom(value) => value,
        }
    }

    fn permits_letter_sync(&self) -> bool {
        matches!(self, Self::LowerAlpha | Self::UpperAlpha)
    }
}

/// Placement of line numbers relative to the text area.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum OdfLineNumberPosition {
    Left,
    Right,
    Inner,
    Outer,
}

impl OdfLineNumberPosition {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "left" => Ok(Self::Left),
            "right" => Ok(Self::Right),
            "inner" => Ok(Self::Inner),
            "outer" => Ok(Self::Outer),
            _ => invalid(format!("unsupported text:number-position '{value}'")),
        }
    }

    fn as_str(self) -> &'static str {
        match self {
            Self::Left => "left",
            Self::Right => "right",
            Self::Inner => "inner",
            Self::Outer => "outer",
        }
    }
}

/// A validated ODF nonnegative length.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OdfNonNegativeLength(String);

impl OdfNonNegativeLength {
    pub fn new(value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        validate_nonnegative_length(&value)?;
        Ok(Self(value))
    }

    pub fn as_str(&self) -> &str {
        &self.0
    }
}

/// Optional separator emitted after every configured number of lines.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct OdfLineNumberingSeparator {
    pub increment: Option<u64>,
    pub text: String,
}

/// One standard `text:linenumbering-configuration` declaration.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct OdfLineNumberingConfiguration {
    pub number_lines: Option<bool>,
    pub number_format: Option<OdfLineNumberFormat>,
    pub letter_sync: Option<bool>,
    pub style_name: Option<String>,
    pub increment: Option<u64>,
    pub number_position: Option<OdfLineNumberPosition>,
    pub offset: Option<OdfNonNegativeLength>,
    pub count_empty_lines: Option<bool>,
    pub count_in_text_boxes: Option<bool>,
    pub restart_on_page: Option<bool>,
    pub separator: Option<OdfLineNumberingSeparator>,
}

impl OdfLineNumberingConfiguration {
    pub fn validate(&self) -> Result<()> {
        if self.letter_sync.is_some()
            && !self
                .number_format
                .as_ref()
                .is_some_and(OdfLineNumberFormat::permits_letter_sync)
        {
            return invalid("style:num-letter-sync requires style:num-format 'a' or 'A'");
        }
        if let Some(format) = &self.number_format {
            validate_value(format.as_str(), "style:num-format", true)?;
        }
        if let Some(style_name) = &self.style_name {
            validate_value(style_name, "text:style-name", false)?;
        }
        if let Some(offset) = &self.offset {
            validate_nonnegative_length(offset.as_str())?;
        }
        if let Some(separator) = &self.separator {
            validate_value(&separator.text, "text:linenumbering-separator", true)?;
        }
        Ok(())
    }

    /// Serialize a namespace-complete line-numbering configuration element.
    pub fn to_xml(&self) -> Result<String> {
        self.validate()?;
        let mut output = String::with_capacity(384);
        output.push_str("<text:linenumbering-configuration xmlns:text=\"");
        output.push_str(std::str::from_utf8(TEXT_NAMESPACE).expect("namespace is UTF-8"));
        output.push_str("\" xmlns:style=\"");
        output.push_str(std::str::from_utf8(STYLE_NAMESPACE).expect("namespace is UTF-8"));
        output.push('"');
        write_bool_attr(&mut output, "text:number-lines", self.number_lines);
        write_attr(
            &mut output,
            "style:num-format",
            self.number_format.as_ref().map(OdfLineNumberFormat::as_str),
        );
        write_bool_attr(&mut output, "style:num-letter-sync", self.letter_sync);
        write_attr(&mut output, "text:style-name", self.style_name.as_deref());
        write_u64_attr(&mut output, "text:increment", self.increment);
        write_attr(
            &mut output,
            "text:number-position",
            self.number_position.map(OdfLineNumberPosition::as_str),
        );
        write_attr(
            &mut output,
            "text:offset",
            self.offset.as_ref().map(OdfNonNegativeLength::as_str),
        );
        write_bool_attr(
            &mut output,
            "text:count-empty-lines",
            self.count_empty_lines,
        );
        write_bool_attr(
            &mut output,
            "text:count-in-text-boxes",
            self.count_in_text_boxes,
        );
        write_bool_attr(&mut output, "text:restart-on-page", self.restart_on_page);
        let Some(separator) = &self.separator else {
            output.push_str("/>");
            return Ok(output);
        };
        output.push_str("><text:linenumbering-separator");
        write_u64_attr(&mut output, "text:increment", separator.increment);
        if separator.text.is_empty() {
            output.push_str("/>");
        } else {
            output.push('>');
            escape_text(&mut output, &separator.text);
            output.push_str("</text:linenumbering-separator>");
        }
        output.push_str("</text:linenumbering-configuration>");
        Ok(output)
    }
}

/// Parse the optional line-numbering declaration from an ODF styles document.
pub fn parse_line_numbering_configuration(
    xml: &str,
) -> Result<Option<OdfLineNumberingConfiguration>> {
    if xml.len() > MAX_DOCUMENT_XML_BYTES {
        return invalid(format!(
            "ODF XML exceeds the {MAX_DOCUMENT_XML_BYTES} byte line-numbering limit"
        ));
    }
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut styles_content_depth = None;
    let mut result = None;
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(xml_error)?;
        let namespace = namespace_kind(&namespace);
        match event {
            Event::Start(element)
                if namespace == NamespaceKind::Text
                    && element.local_name().as_ref() == b"linenumbering-configuration" =>
            {
                if styles_content_depth != Some(depth) {
                    return invalid(
                        "text:linenumbering-configuration must be a direct office:styles child",
                    );
                }
                if result.is_some() {
                    return invalid("ODF styles contain duplicate line-numbering configurations");
                }
                let mut configuration = parse_attributes(&reader, &element)?;
                configuration.separator = parse_configuration_body(&mut reader)?;
                configuration.validate()?;
                result = Some(configuration);
            },
            Event::Empty(element)
                if namespace == NamespaceKind::Text
                    && element.local_name().as_ref() == b"linenumbering-configuration" =>
            {
                if styles_content_depth != Some(depth) {
                    return invalid(
                        "text:linenumbering-configuration must be a direct office:styles child",
                    );
                }
                if result.is_some() {
                    return invalid("ODF styles contain duplicate line-numbering configurations");
                }
                let configuration = parse_attributes(&reader, &element)?;
                configuration.validate()?;
                result = Some(configuration);
            },
            Event::Start(element) => {
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| Error::InvalidFormat("ODF XML depth overflow".to_string()))?;
                if depth > MAX_XML_DEPTH {
                    return invalid(format!("ODF XML exceeds the {MAX_XML_DEPTH} depth limit"));
                }
                if namespace == NamespaceKind::Office && element.local_name().as_ref() == b"styles"
                    && styles_content_depth.replace(depth).is_some() {
                        return invalid("ODF XML contains nested office:styles elements");
                    }
            },
            Event::End(element) => {
                if namespace == NamespaceKind::Office
                    && element.local_name().as_ref() == b"styles"
                    && styles_content_depth == Some(depth)
                {
                    styles_content_depth = None;
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| Error::InvalidFormat("invalid ODF XML depth".to_string()))?;
            },
            Event::DocType(_) => return invalid("DOCTYPE is not allowed in ODF line metadata"),
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    Ok(result)
}

#[derive(Clone)]
struct XmlSpan {
    start: usize,
    end: usize,
}

enum StylesSite {
    Content(usize),
    Empty(XmlSpan, String),
}

fn event_start(xml: &str, end: usize) -> Result<usize> {
    xml[..end].rfind('<').ok_or_else(|| {
        Error::InvalidFormat("invalid line-numbering XML event boundary".to_string())
    })
}

fn locate_configuration(xml: &str) -> Result<(Option<XmlSpan>, StylesSite)> {
    parse_line_numbering_configuration(xml)?;

    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut stack: Vec<(NamespaceKind, Vec<u8>)> = Vec::new();
    let mut target = None;
    let mut open_target = None::<(usize, usize)>;
    let mut styles_site = None;
    loop {
        let (resolved, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(xml_error)?;
        let namespace = namespace_kind(&resolved);
        match event {
            Event::Start(ref element) => {
                let end = reader.buffer_position() as usize;
                let start = event_start(xml, end)?;
                let local = element.local_name().as_ref().to_vec();
                let depth = stack.len() + 1;
                if namespace == NamespaceKind::Text
                    && local == b"linenumbering-configuration"
                    && matches!(stack.last(), Some((NamespaceKind::Office, parent)) if parent == b"styles")
                {
                    open_target = Some((depth, start));
                }
                stack.push((namespace, local));
            },
            Event::Empty(ref element) => {
                let end = reader.buffer_position() as usize;
                let start = event_start(xml, end)?;
                let local = element.local_name().as_ref().to_vec();
                if namespace == NamespaceKind::Text
                    && local == b"linenumbering-configuration"
                    && matches!(stack.last(), Some((NamespaceKind::Office, parent)) if parent == b"styles")
                {
                    target = Some(XmlSpan { start, end });
                }
                if namespace == NamespaceKind::Office && local == b"styles" {
                    if styles_site.is_some() {
                        return invalid("multiple office:styles elements are not supported");
                    }
                    styles_site = Some(StylesSite::Empty(
                        XmlSpan { start, end },
                        std::str::from_utf8(element.name().as_ref())
                            .map_err(|_| {
                                Error::InvalidFormat("invalid office:styles QName".to_string())
                            })?
                            .to_string(),
                    ));
                }
            },
            Event::End(_) => {
                let end = reader.buffer_position() as usize;
                let start = event_start(xml, end)?;
                let depth = stack.len();
                if open_target.is_some_and(|(target_depth, _)| target_depth == depth) {
                    let (_, target_start) = open_target.take().expect("target depth checked");
                    target = Some(XmlSpan {
                        start: target_start,
                        end,
                    });
                }
                if matches!(stack.last(), Some((NamespaceKind::Office, local)) if local == b"styles")
                {
                    if styles_site.is_some() {
                        return invalid("multiple office:styles elements are not supported");
                    }
                    styles_site = Some(StylesSite::Content(start));
                }
                stack.pop();
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }

    Ok((
        target,
        styles_site
            .ok_or_else(|| Error::InvalidFormat("document has no office:styles".to_string()))?,
    ))
}

/// Insert or replace the document line-numbering declaration without rewriting
/// unrelated style XML.
pub(crate) fn set_line_numbering_configuration_xml(
    xml: &str,
    configuration: &OdfLineNumberingConfiguration,
) -> Result<String> {
    configuration.validate()?;
    let (target, site) = locate_configuration(xml)?;
    let fragment = configuration.to_xml()?;
    if let Some(span) = target {
        return Ok(format!(
            "{}{}{}",
            &xml[..span.start],
            fragment,
            &xml[span.end..]
        ));
    }
    match site {
        StylesSite::Content(insertion) => Ok(format!(
            "{}{}{}",
            &xml[..insertion],
            fragment,
            &xml[insertion..]
        )),
        StylesSite::Empty(span, qname) => {
            let raw = &xml[span.start..span.end];
            let slash = raw
                .rfind("/>")
                .ok_or_else(|| Error::InvalidFormat("invalid empty office:styles".to_string()))?;
            Ok(format!(
                "{}{}>{}</{}>{}",
                &xml[..span.start],
                &raw[..slash],
                fragment,
                qname,
                &xml[span.end..]
            ))
        },
    }
}

/// Remove the document line-numbering declaration without rewriting unrelated
/// style XML.
pub(crate) fn remove_line_numbering_configuration_xml(xml: &str) -> Result<String> {
    let (target, _) = locate_configuration(xml)?;
    let Some(span) = target else {
        return Ok(xml.to_string());
    };
    Ok(format!("{}{}", &xml[..span.start], &xml[span.end..]))
}

impl OpenDocumentPackage {
    /// Return stored document line-numbering configuration from styles XML.
    ///
    /// The declaration is presentation metadata only. It is never used to
    /// paginate a document or generate line numbers.
    pub fn line_numbering_configuration(&self) -> Result<Option<OdfLineNumberingConfiguration>> {
        self.styles_xml()?
            .map_or_else(|| Ok(None), |xml| parse_line_numbering_configuration(&xml))
    }
}

impl FlatOpenDocument {
    /// Return stored document line-numbering configuration from flat ODF XML.
    ///
    /// The declaration is presentation metadata only. It is never used to
    /// paginate a document or generate line numbers.
    pub fn line_numbering_configuration(&self) -> Result<Option<OdfLineNumberingConfiguration>> {
        parse_line_numbering_configuration(self.xml())
    }
}

fn parse_attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<OdfLineNumberingConfiguration> {
    let mut configuration = OdfLineNumberingConfiguration::default();
    let mut seen = HashSet::new();
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        if attribute.key.as_ref() == b"xmlns" || attribute.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let namespace = namespace_kind(&namespace);
        if !seen.insert((namespace, local.as_ref().to_vec())) {
            return invalid("duplicate line-numbering configuration attribute");
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(xml_error)?
            .into_owned();
        validate_value(&value, "line-numbering attribute", true)?;
        match (namespace, local.as_ref()) {
            (NamespaceKind::Text, b"number-lines") => {
                configuration.number_lines = Some(parse_bool(&value, "text:number-lines")?)
            },
            (NamespaceKind::Style, b"num-format") => {
                configuration.number_format = Some(OdfLineNumberFormat::parse(value)?)
            },
            (NamespaceKind::Style, b"num-letter-sync") => {
                configuration.letter_sync = Some(parse_bool(&value, "style:num-letter-sync")?)
            },
            (NamespaceKind::Text, b"style-name") => {
                validate_value(&value, "text:style-name", false)?;
                configuration.style_name = Some(value);
            },
            (NamespaceKind::Text, b"increment") => {
                configuration.increment = Some(parse_nonnegative_integer(&value, "text:increment")?)
            },
            (NamespaceKind::Text, b"number-position") => {
                configuration.number_position = Some(OdfLineNumberPosition::parse(&value)?)
            },
            (NamespaceKind::Text, b"offset") => {
                configuration.offset = Some(OdfNonNegativeLength::new(value)?)
            },
            (NamespaceKind::Text, b"count-empty-lines") => {
                configuration.count_empty_lines =
                    Some(parse_bool(&value, "text:count-empty-lines")?)
            },
            (NamespaceKind::Text, b"count-in-text-boxes") => {
                configuration.count_in_text_boxes =
                    Some(parse_bool(&value, "text:count-in-text-boxes")?)
            },
            (NamespaceKind::Text, b"restart-on-page") => {
                configuration.restart_on_page = Some(parse_bool(&value, "text:restart-on-page")?)
            },
            _ => return invalid("unsupported line-numbering configuration attribute"),
        }
    }
    Ok(configuration)
}

fn parse_configuration_body(
    reader: &mut NsReader<&[u8]>,
) -> Result<Option<OdfLineNumberingSeparator>> {
    let mut separator = None;
    let mut buffer = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(xml_error)?;
        let namespace = namespace_kind(&namespace);
        match event {
            Event::Start(element)
                if namespace == NamespaceKind::Text
                    && element.local_name().as_ref() == b"linenumbering-separator" =>
            {
                if separator.is_some() {
                    return invalid("duplicate text:linenumbering-separator");
                }
                separator = Some(parse_separator(reader, &element)?);
            },
            Event::Empty(element)
                if namespace == NamespaceKind::Text
                    && element.local_name().as_ref() == b"linenumbering-separator" =>
            {
                if separator.is_some() {
                    return invalid("duplicate text:linenumbering-separator");
                }
                separator = Some(OdfLineNumberingSeparator {
                    increment: parse_separator_increment(reader, &element)?,
                    text: String::new(),
                });
            },
            Event::End(element)
                if namespace == NamespaceKind::Text
                    && element.local_name().as_ref() == b"linenumbering-configuration" =>
            {
                break;
            },
            Event::Text(text) => require_whitespace(
                &text.decode().map_err(xml_error)?,
                "linenumbering-configuration",
            )?,
            Event::Comment(_) | Event::PI(_) => {},
            Event::DocType(_) => return invalid("DOCTYPE is not allowed in line numbering"),
            Event::Eof => return invalid("unterminated text:linenumbering-configuration"),
            _ => return invalid("unsupported child in linenumbering-configuration"),
        }
        buffer.clear();
    }
    Ok(separator)
}

fn parse_separator(
    reader: &mut NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<OdfLineNumberingSeparator> {
    let increment = parse_separator_increment(reader, element)?;
    let mut text = String::new();
    let mut buffer = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(xml_error)?;
        let namespace = namespace_kind(&namespace);
        match event {
            Event::Text(value) => {
                let value = value.decode().map_err(xml_error)?;
                append_separator_text(&mut text, &value)?;
            },
            Event::CData(value) => {
                let value = value.decode().map_err(xml_error)?;
                append_separator_text(&mut text, &value)?;
            },
            Event::GeneralRef(reference) => {
                if let Some(character) = reference.resolve_char_ref().map_err(xml_error)? {
                    let mut encoded = [0u8; 4];
                    append_separator_text(&mut text, character.encode_utf8(&mut encoded))?;
                } else {
                    let name = reference.decode().map_err(xml_error)?;
                    let value = match name.as_ref() {
                        "amp" => "&",
                        "lt" => "<",
                        "gt" => ">",
                        "apos" => "'",
                        "quot" => "\"",
                        _ => return invalid("unsupported entity in line-numbering separator"),
                    };
                    append_separator_text(&mut text, value)?;
                }
            },
            Event::End(element)
                if namespace == NamespaceKind::Text
                    && element.local_name().as_ref() == b"linenumbering-separator" =>
            {
                break;
            },
            Event::Comment(_) | Event::PI(_) => {},
            Event::Eof => return invalid("unterminated text:linenumbering-separator"),
            _ => return invalid("text:linenumbering-separator may contain text only"),
        }
        buffer.clear();
    }
    Ok(OdfLineNumberingSeparator { increment, text })
}

fn parse_separator_increment(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<Option<u64>> {
    let mut increment = None;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        if attribute.key.as_ref() == b"xmlns" || attribute.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        if namespace_kind(&namespace) != NamespaceKind::Text || local.as_ref() != b"increment" {
            return invalid("unsupported linenumbering-separator attribute");
        }
        if increment.is_some() {
            return invalid("duplicate linenumbering-separator increment");
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(xml_error)?;
        increment = Some(parse_nonnegative_integer(
            &value,
            "linenumbering-separator increment",
        )?);
    }
    Ok(increment)
}

fn namespace_kind(namespace: &ResolveResult<'_>) -> NamespaceKind {
    match namespace {
        ResolveResult::Bound(value) if value.as_ref() == OFFICE_NAMESPACE => NamespaceKind::Office,
        ResolveResult::Bound(value) if value.as_ref() == TEXT_NAMESPACE => NamespaceKind::Text,
        ResolveResult::Bound(value) if value.as_ref() == STYLE_NAMESPACE => NamespaceKind::Style,
        _ => NamespaceKind::Other,
    }
}

fn validate_nonnegative_length(value: &str) -> Result<()> {
    let Some(number) = ["cm", "mm", "in", "pt", "pc", "px"]
        .iter()
        .find_map(|unit| value.strip_suffix(unit))
    else {
        return invalid(format!("invalid nonnegative ODF length '{value}'"));
    };
    let mut dots = 0usize;
    let mut digits = 0usize;
    for byte in number.bytes() {
        match byte {
            b'.' => dots += 1,
            b'0'..=b'9' => digits += 1,
            _ => return invalid(format!("invalid nonnegative ODF length '{value}'")),
        }
    }
    if number.is_empty() || number == "." || dots > 1 || digits == 0 {
        return invalid(format!("invalid nonnegative ODF length '{value}'"));
    }
    validate_value(value, "text:offset", false)
}

fn parse_bool(value: &str, name: &str) -> Result<bool> {
    match value {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        _ => invalid(format!("{name} must be true, false, 1, or 0")),
    }
}

fn parse_nonnegative_integer(value: &str, name: &str) -> Result<u64> {
    if value.is_empty() || !value.bytes().all(|byte| byte.is_ascii_digit()) {
        return invalid(format!("{name} must be a nonnegative integer"));
    }
    value
        .parse::<u64>()
        .map_err(|_| Error::InvalidFormat(format!("{name} exceeds the u64 range")))
}

fn append_separator_text(output: &mut String, value: &str) -> Result<()> {
    let size = output
        .len()
        .checked_add(value.len())
        .ok_or_else(|| Error::InvalidFormat("separator text size overflow".to_string()))?;
    if size > MAX_VALUE_BYTES {
        return invalid(format!(
            "line-numbering separator exceeds the {MAX_VALUE_BYTES} byte limit"
        ));
    }
    output.push_str(value);
    Ok(())
}

fn validate_value(value: &str, name: &str, allow_empty: bool) -> Result<()> {
    if !allow_empty && value.is_empty() {
        return invalid(format!("{name} must not be empty"));
    }
    if value.len() > MAX_VALUE_BYTES {
        return invalid(format!("{name} exceeds the {MAX_VALUE_BYTES} byte limit"));
    }
    Ok(())
}

fn require_whitespace(value: &str, context: &str) -> Result<()> {
    if value.trim().is_empty() {
        Ok(())
    } else {
        invalid(format!("text:{context} cannot contain text"))
    }
}

fn write_attr(output: &mut String, name: &str, value: Option<&str>) {
    let Some(value) = value else { return };
    output.push(' ');
    output.push_str(name);
    output.push_str("=\"");
    escape_attribute(output, value);
    output.push('"');
}

fn write_bool_attr(output: &mut String, name: &str, value: Option<bool>) {
    write_attr(
        output,
        name,
        value.map(|value| if value { "true" } else { "false" }),
    );
}

fn write_u64_attr(output: &mut String, name: &str, value: Option<u64>) {
    let Some(value) = value else { return };
    output.push(' ');
    output.push_str(name);
    output.push_str("=\"");
    output.push_str(&value.to_string());
    output.push('"');
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

fn escape_text(output: &mut String, value: &str) {
    for character in value.chars() {
        match character {
            '&' => output.push_str("&amp;"),
            '<' => output.push_str("&lt;"),
            '>' => output.push_str("&gt;"),
            _ => output.push(character),
        }
    }
}

fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::InvalidFormat(format!("invalid ODF line-numbering XML: {error}"))
}

fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::InvalidFormat(message.into()))
}

#[cfg(test)]
mod tests {
    use super::*;

    const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
    const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
    const STYLE: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";

    fn styles(body: &str) -> String {
        format!(
            r#"<o:document-styles xmlns:o="{OFFICE}" xmlns:t="{TEXT}" xmlns:s="{STYLE}"><o:styles>{body}</o:styles><o:automatic-styles/><o:master-styles/></o:document-styles>"#
        )
    }

    #[test]
    fn parses_and_round_trips_complete_line_numbering() {
        // ODF 1.2/1.3 text:linenumbering-configuration grammar. LibreOffice
        // writes the same configuration in office:styles.
        let xml = styles(
            r#"<t:linenumbering-configuration t:number-lines="1" s:num-format="A" s:num-letter-sync="false" t:style-name="Line &amp; Number" t:increment="5" t:number-position="outer" t:offset="0.25in" t:count-empty-lines="true" t:count-in-text-boxes="0" t:restart-on-page="false"><t:linenumbering-separator t:increment="10"> / &amp; </t:linenumbering-separator></t:linenumbering-configuration>"#,
        );
        let configuration = parse_line_numbering_configuration(&xml).unwrap().unwrap();
        assert_eq!(configuration.number_lines, Some(true));
        assert_eq!(
            configuration.number_format,
            Some(OdfLineNumberFormat::UpperAlpha)
        );
        assert_eq!(
            configuration.number_position,
            Some(OdfLineNumberPosition::Outer)
        );
        assert_eq!(configuration.offset.as_ref().unwrap().as_str(), "0.25in");
        assert_eq!(configuration.separator.as_ref().unwrap().text, " / & ");

        let serialized = configuration.to_xml().unwrap();
        let reparsed = parse_line_numbering_configuration(&styles(&serialized))
            .unwrap()
            .unwrap();
        assert_eq!(reparsed, configuration);
    }

    #[test]
    fn preserves_empty_and_custom_number_formats() {
        for format in [
            OdfLineNumberFormat::Empty,
            OdfLineNumberFormat::Custom("一, 二, 三".to_string()),
        ] {
            let configuration = OdfLineNumberingConfiguration {
                number_format: Some(format.clone()),
                ..OdfLineNumberingConfiguration::default()
            };
            let parsed =
                parse_line_numbering_configuration(&styles(&configuration.to_xml().unwrap()))
                    .unwrap()
                    .unwrap();
            assert_eq!(parsed.number_format, Some(format));
        }
    }

    #[test]
    fn rejects_malformed_or_misplaced_configurations() {
        for body in [
            r#"<t:linenumbering-configuration t:number-lines="yes"/>"#,
            r#"<t:linenumbering-configuration s:num-format="1" s:num-letter-sync="true"/>"#,
            r#"<t:linenumbering-configuration t:number-position="center"/>"#,
            r#"<t:linenumbering-configuration t:offset="-1cm"/>"#,
            r#"<t:linenumbering-configuration t:increment="-1"/>"#,
            r#"<t:linenumbering-configuration><t:linenumbering-separator/><t:linenumbering-separator/></t:linenumbering-configuration>"#,
            r#"<t:linenumbering-configuration><t:linenumbering-separator><t:span/></t:linenumbering-separator></t:linenumbering-configuration>"#,
        ] {
            assert!(
                parse_line_numbering_configuration(&styles(body)).is_err(),
                "{body}"
            );
        }
        let misplaced = format!(
            r#"<o:document-content xmlns:o="{OFFICE}" xmlns:t="{TEXT}"><o:body><t:linenumbering-configuration/></o:body></o:document-content>"#
        );
        assert!(parse_line_numbering_configuration(&misplaced).is_err());
        assert!(
            parse_line_numbering_configuration(&styles(
                r#"<t:linenumbering-configuration/><t:linenumbering-configuration/>"#
            ))
            .is_err()
        );
    }

    #[test]
    fn replaces_inserts_and_removes_configuration_without_rewriting_other_styles() {
        let original = styles(
            r#"<s:style s:name="Preserved"/><t:linenumbering-configuration t:number-lines="false" s:num-format="1"/>"#,
        );
        let configuration = OdfLineNumberingConfiguration {
            number_lines: Some(true),
            number_format: Some(OdfLineNumberFormat::LowerAlpha),
            letter_sync: Some(true),
            style_name: Some("LineNumbers".to_string()),
            increment: Some(3),
            number_position: Some(OdfLineNumberPosition::Outer),
            offset: Some(OdfNonNegativeLength::new("0.25in").unwrap()),
            count_empty_lines: Some(true),
            count_in_text_boxes: Some(false),
            restart_on_page: Some(true),
            separator: Some(OdfLineNumberingSeparator {
                increment: Some(6),
                text: " · ".to_string(),
            }),
        };

        let replaced = set_line_numbering_configuration_xml(&original, &configuration).unwrap();
        assert!(replaced.contains(r#"<s:style s:name="Preserved"/>"#));
        assert_eq!(
            parse_line_numbering_configuration(&replaced).unwrap(),
            Some(configuration.clone())
        );

        let removed = remove_line_numbering_configuration_xml(&replaced).unwrap();
        assert!(removed.contains(r#"<s:style s:name="Preserved"/>"#));
        assert_eq!(parse_line_numbering_configuration(&removed).unwrap(), None);

        let empty_styles = format!(
            r#"<o:document-styles xmlns:o="{OFFICE}" xmlns:t="{TEXT}" xmlns:s="{STYLE}"><o:styles/><o:automatic-styles/><o:master-styles/></o:document-styles>"#
        );
        let inserted = set_line_numbering_configuration_xml(&empty_styles, &configuration).unwrap();
        assert!(inserted.contains("<o:styles>"));
        assert_eq!(
            parse_line_numbering_configuration(&inserted).unwrap(),
            Some(configuration.clone())
        );

        let flat_xml = format!(
            r#"<o:document xmlns:o="{OFFICE}" xmlns:t="{TEXT}" xmlns:s="{STYLE}" o:mimetype="application/vnd.oasis.opendocument.text" o:version="1.3"><o:styles>{}</o:styles><o:body><o:text/></o:body></o:document>"#,
            configuration.to_xml().unwrap()
        );
        let flat = FlatOpenDocument::from_bytes(flat_xml.into_bytes()).unwrap();
        assert_eq!(
            flat.line_numbering_configuration().unwrap(),
            Some(configuration)
        );
    }
}
