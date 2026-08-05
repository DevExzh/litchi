//! ODF footnote and endnote configuration metadata.

use std::collections::HashSet;

use litchi_core::{Error, Result};
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::ResolveResult,
    reader::NsReader,
};

use crate::line_numbering::LineNumberFormat;
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

/// Note class selected by `text:note-class`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Class {
    Footnote,
    Endnote,
}

impl Class {
    pub const ALL: [Self; 2] = [Self::Footnote, Self::Endnote];
    fn parse(value: &str) -> Result<Self> {
        match value {
            "footnote" => Ok(Self::Footnote),
            "endnote" => Ok(Self::Endnote),
            _ => invalid(format!("unsupported text:note-class '{value}'")),
        }
    }

    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Footnote => "footnote",
            Self::Endnote => "endnote",
        }
    }
}

/// Scope at which note numbering restarts.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum NumberingScope {
    Document,
    Chapter,
    Page,
}

impl NumberingScope {
    pub const ALL: [Self; 3] = [Self::Document, Self::Chapter, Self::Page];
    fn parse(value: &str) -> Result<Self> {
        match value {
            "document" => Ok(Self::Document),
            "chapter" => Ok(Self::Chapter),
            "page" => Ok(Self::Page),
            _ => invalid(format!("unsupported text:start-numbering-at '{value}'")),
        }
    }

    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Document => "document",
            Self::Chapter => "chapter",
            Self::Page => "page",
        }
    }
}

/// Placement of footnotes in the document.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Position {
    Text,
    Page,
    Section,
    Document,
}

impl Position {
    pub const ALL: [Self; 4] = [Self::Text, Self::Page, Self::Section, Self::Document];
    fn parse(value: &str) -> Result<Self> {
        match value {
            "text" => Ok(Self::Text),
            "page" => Ok(Self::Page),
            "section" => Ok(Self::Section),
            "document" => Ok(Self::Document),
            _ => invalid(format!("unsupported text:footnotes-position '{value}'")),
        }
    }

    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Text => "text",
            Self::Page => "page",
            Self::Section => "section",
            Self::Document => "document",
        }
    }
}

/// One `text:notes-configuration` declaration.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Configuration {
    pub note_class: Class,
    pub citation_style_name: Option<String>,
    pub citation_body_style_name: Option<String>,
    pub default_style_name: Option<String>,
    pub master_page_name: Option<String>,
    pub start_value: Option<u64>,
    pub number_prefix: Option<String>,
    pub number_suffix: Option<String>,
    pub number_format: Option<LineNumberFormat>,
    pub letter_sync: Option<bool>,
    pub start_numbering_at: Option<NumberingScope>,
    pub footnotes_position: Option<Position>,
    pub continuation_notice_forward: Option<String>,
    pub continuation_notice_backward: Option<String>,
}

impl Configuration {
    pub fn new(note_class: Class) -> Self {
        Self {
            note_class,
            citation_style_name: None,
            citation_body_style_name: None,
            default_style_name: None,
            master_page_name: None,
            start_value: None,
            number_prefix: None,
            number_suffix: None,
            number_format: None,
            letter_sync: None,
            start_numbering_at: None,
            footnotes_position: None,
            continuation_notice_forward: None,
            continuation_notice_backward: None,
        }
    }

    pub fn validate(&self) -> Result<()> {
        if self.letter_sync.is_some()
            && !matches!(
                self.number_format,
                Some(LineNumberFormat::LowerAlpha | LineNumberFormat::UpperAlpha)
            )
        {
            return invalid("style:num-letter-sync requires style:num-format 'a' or 'A'");
        }
        for (value, name) in [
            (
                self.citation_style_name.as_deref(),
                "text:citation-style-name",
            ),
            (
                self.citation_body_style_name.as_deref(),
                "text:citation-body-style-name",
            ),
            (
                self.default_style_name.as_deref(),
                "text:default-style-name",
            ),
            (self.master_page_name.as_deref(), "text:master-page-name"),
        ] {
            if let Some(value) = value {
                validate_style_name_ref(value, name)?;
            }
        }
        for (value, name, allow_empty) in [
            (self.number_prefix.as_deref(), "style:num-prefix", true),
            (self.number_suffix.as_deref(), "style:num-suffix", true),
            (
                self.continuation_notice_forward.as_deref(),
                "text:note-continuation-notice-forward",
                true,
            ),
            (
                self.continuation_notice_backward.as_deref(),
                "text:note-continuation-notice-backward",
                true,
            ),
        ] {
            if let Some(value) = value {
                validate_value(value, name, allow_empty)?;
            }
        }
        if let Some(format) = &self.number_format {
            validate_value(format.as_str(), "style:num-format", true)?;
        }
        Ok(())
    }

    /// Serialize a namespace-complete notes configuration element.
    pub fn to_xml(&self) -> Result<String> {
        self.validate()?;
        let mut output = String::with_capacity(512);
        output.push_str("<text:notes-configuration xmlns:text=\"");
        output.push_str(std::str::from_utf8(TEXT_NAMESPACE).expect("namespace is UTF-8"));
        output.push_str("\" xmlns:style=\"");
        output.push_str(std::str::from_utf8(STYLE_NAMESPACE).expect("namespace is UTF-8"));
        output.push('"');
        write_attr(
            &mut output,
            "text:note-class",
            Some(self.note_class.as_str()),
        );
        write_attr(
            &mut output,
            "text:citation-style-name",
            self.citation_style_name.as_deref(),
        );
        write_attr(
            &mut output,
            "text:citation-body-style-name",
            self.citation_body_style_name.as_deref(),
        );
        write_attr(
            &mut output,
            "text:default-style-name",
            self.default_style_name.as_deref(),
        );
        write_attr(
            &mut output,
            "text:master-page-name",
            self.master_page_name.as_deref(),
        );
        write_u64_attr(&mut output, "text:start-value", self.start_value);
        write_attr(
            &mut output,
            "style:num-prefix",
            self.number_prefix.as_deref(),
        );
        write_attr(
            &mut output,
            "style:num-suffix",
            self.number_suffix.as_deref(),
        );
        write_attr(
            &mut output,
            "style:num-format",
            self.number_format.as_ref().map(LineNumberFormat::as_str),
        );
        write_bool_attr(&mut output, "style:num-letter-sync", self.letter_sync);
        write_attr(
            &mut output,
            "text:start-numbering-at",
            self.start_numbering_at.map(NumberingScope::as_str),
        );
        write_attr(
            &mut output,
            "text:footnotes-position",
            self.footnotes_position.map(Position::as_str),
        );
        if self.continuation_notice_forward.is_none() && self.continuation_notice_backward.is_none()
        {
            output.push_str("/>");
            return Ok(output);
        }
        output.push('>');
        write_notice(
            &mut output,
            "text:note-continuation-notice-forward",
            self.continuation_notice_forward.as_deref(),
        );
        write_notice(
            &mut output,
            "text:note-continuation-notice-backward",
            self.continuation_notice_backward.as_deref(),
        );
        output.push_str("</text:notes-configuration>");
        Ok(output)
    }
}

/// The at-most-one configuration for each standard note class.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Configurations {
    pub footnote: Option<Configuration>,
    pub endnote: Option<Configuration>,
}

impl Configurations {
    pub fn get(&self, note_class: Class) -> Option<&Configuration> {
        match note_class {
            Class::Footnote => self.footnote.as_ref(),
            Class::Endnote => self.endnote.as_ref(),
        }
    }

    pub fn validate(&self) -> Result<()> {
        if let Some(configuration) = &self.footnote {
            if configuration.note_class != Class::Footnote {
                return invalid("footnote slot contains an endnote configuration");
            }
            configuration.validate()?;
        }
        if let Some(configuration) = &self.endnote {
            if configuration.note_class != Class::Endnote {
                return invalid("endnote slot contains a footnote configuration");
            }
            configuration.validate()?;
        }
        Ok(())
    }

    /// Serialize both configurations in canonical footnote/endnote order.
    pub fn to_xml_fragment(&self) -> Result<String> {
        self.validate()?;
        let mut output = String::new();
        if let Some(configuration) = &self.footnote {
            output.push_str(&configuration.to_xml()?);
        }
        if let Some(configuration) = &self.endnote {
            output.push_str(&configuration.to_xml()?);
        }
        Ok(output)
    }
}

/// Parse footnote/endnote configurations from an ODF styles document.
pub fn parse(xml: &str) -> Result<Configurations> {
    if xml.len() > MAX_DOCUMENT_XML_BYTES {
        return invalid(format!(
            "ODF XML exceeds the {MAX_DOCUMENT_XML_BYTES} byte notes limit"
        ));
    }
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut styles_content_depth = None;
    let mut section_properties_depth = None;
    let mut result = Configurations::default();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(xml_error)?;
        let namespace = namespace_kind(&namespace);
        match event {
            Event::Start(element)
                if namespace == NamespaceKind::Text
                    && element.local_name().as_ref() == b"notes-configuration" =>
            {
                let collect = if styles_content_depth == Some(depth) {
                    true
                } else if section_properties_depth == Some(depth) {
                    false
                } else {
                    require_styles_scope(styles_content_depth, depth)?;
                    unreachable!("scope validation always fails outside supported parents")
                };
                let mut configuration = parse_attributes(&reader, &element)?;
                parse_notices(&mut reader, &mut configuration)?;
                configuration.validate()?;
                if collect {
                    insert_configuration(&mut result, configuration)?;
                }
            },
            Event::Empty(element)
                if namespace == NamespaceKind::Text
                    && element.local_name().as_ref() == b"notes-configuration" =>
            {
                let collect = if styles_content_depth == Some(depth) {
                    true
                } else if section_properties_depth == Some(depth) {
                    false
                } else {
                    require_styles_scope(styles_content_depth, depth)?;
                    unreachable!("scope validation always fails outside supported parents")
                };
                let configuration = parse_attributes(&reader, &element)?;
                configuration.validate()?;
                if collect {
                    insert_configuration(&mut result, configuration)?;
                }
            },
            Event::Start(element) if element.local_name().as_ref() == b"notes-configuration" => {
                return invalid("notes-configuration uses the wrong namespace");
            },
            Event::Empty(element) if element.local_name().as_ref() == b"notes-configuration" => {
                return invalid("notes-configuration uses the wrong namespace");
            },
            Event::Start(element) => {
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| Error::InvalidFormat("ODF XML depth overflow".to_string()))?;
                if depth > MAX_XML_DEPTH {
                    return invalid(format!("ODF XML exceeds the {MAX_XML_DEPTH} depth limit"));
                }
                if namespace == NamespaceKind::Office
                    && element.local_name().as_ref() == b"styles"
                    && styles_content_depth.replace(depth).is_some()
                {
                    return invalid("ODF XML contains nested office:styles elements");
                }
                if namespace == NamespaceKind::Style
                    && element.local_name().as_ref() == b"section-properties"
                    && section_properties_depth.replace(depth).is_some()
                {
                    return invalid("ODF XML contains nested style:section-properties elements");
                }
            },
            Event::End(element) => {
                if namespace == NamespaceKind::Office
                    && element.local_name().as_ref() == b"styles"
                    && styles_content_depth == Some(depth)
                {
                    styles_content_depth = None;
                }
                if namespace == NamespaceKind::Style
                    && element.local_name().as_ref() == b"section-properties"
                    && section_properties_depth == Some(depth)
                {
                    section_properties_depth = None;
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| Error::InvalidFormat("invalid ODF XML depth".to_string()))?;
            },
            Event::DocType(_) => return invalid("DOCTYPE is not allowed in ODF note metadata"),
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    Ok(result)
}

fn parse_attributes(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<Configuration> {
    let mut note_class = None;
    let mut citation_style_name = None;
    let mut citation_body_style_name = None;
    let mut default_style_name = None;
    let mut master_page_name = None;
    let mut start_value = None;
    let mut number_prefix = None;
    let mut number_suffix = None;
    let mut number_format = None;
    let mut letter_sync = None;
    let mut start_numbering_at = None;
    let mut footnotes_position = None;
    let mut seen = HashSet::new();

    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        if attribute.key.as_ref() == b"xmlns" || attribute.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let namespace = namespace_kind(&namespace);
        if !seen.insert((namespace, local.as_ref().to_vec())) {
            return invalid("duplicate text:notes-configuration attribute");
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(xml_error)?
            .into_owned();
        validate_value(&value, "notes configuration attribute", true)?;
        match (namespace, local.as_ref()) {
            (NamespaceKind::Text, b"note-class") => note_class = Some(Class::parse(&value)?),
            (NamespaceKind::Text, b"citation-style-name") => citation_style_name = Some(value),
            (NamespaceKind::Text, b"citation-body-style-name") => {
                citation_body_style_name = Some(value)
            },
            (NamespaceKind::Text, b"default-style-name") => default_style_name = Some(value),
            (NamespaceKind::Text, b"master-page-name") => master_page_name = Some(value),
            (NamespaceKind::Text, b"start-value") => {
                start_value = Some(parse_nonnegative_integer(&value, "text:start-value")?)
            },
            (NamespaceKind::Style, b"num-prefix") => number_prefix = Some(value),
            (NamespaceKind::Style, b"num-suffix") => number_suffix = Some(value),
            (NamespaceKind::Style, b"num-format") => {
                number_format = Some(LineNumberFormat::parse(value)?)
            },
            (NamespaceKind::Style, b"num-letter-sync") => {
                letter_sync = Some(parse_bool(&value, "style:num-letter-sync")?)
            },
            (NamespaceKind::Text, b"start-numbering-at") => {
                start_numbering_at = Some(NumberingScope::parse(&value)?)
            },
            (NamespaceKind::Text, b"footnotes-position") => {
                footnotes_position = Some(Position::parse(&value)?)
            },
            _ => return invalid("unsupported text:notes-configuration attribute"),
        }
    }
    let note_class = note_class.ok_or_else(|| {
        Error::InvalidFormat("notes configuration requires text:note-class".to_string())
    })?;
    Ok(Configuration {
        note_class,
        citation_style_name,
        citation_body_style_name,
        default_style_name,
        master_page_name,
        start_value,
        number_prefix,
        number_suffix,
        number_format,
        letter_sync,
        start_numbering_at,
        footnotes_position,
        continuation_notice_forward: None,
        continuation_notice_backward: None,
    })
}

fn parse_notices(reader: &mut NsReader<&[u8]>, configuration: &mut Configuration) -> Result<()> {
    let mut buffer = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(xml_error)?;
        let namespace = namespace_kind(&namespace);
        match event {
            Event::Start(element)
                if namespace == NamespaceKind::Text
                    && element.local_name().as_ref() == b"note-continuation-notice-forward" =>
            {
                reject_attributes(&element)?;
                if configuration.continuation_notice_forward.is_some() {
                    return invalid("duplicate forward note continuation notice");
                }
                configuration.continuation_notice_forward =
                    Some(parse_notice(reader, b"note-continuation-notice-forward")?);
            },
            Event::Empty(element)
                if namespace == NamespaceKind::Text
                    && element.local_name().as_ref() == b"note-continuation-notice-forward" =>
            {
                if configuration
                    .continuation_notice_forward
                    .replace(String::new())
                    .is_some()
                {
                    return invalid("duplicate forward note continuation notice");
                }
                reject_attributes(&element)?;
            },
            Event::Start(element)
                if namespace == NamespaceKind::Text
                    && element.local_name().as_ref() == b"note-continuation-notice-backward" =>
            {
                reject_attributes(&element)?;
                if configuration.continuation_notice_backward.is_some() {
                    return invalid("duplicate backward note continuation notice");
                }
                configuration.continuation_notice_backward =
                    Some(parse_notice(reader, b"note-continuation-notice-backward")?);
            },
            Event::Empty(element)
                if namespace == NamespaceKind::Text
                    && element.local_name().as_ref() == b"note-continuation-notice-backward" =>
            {
                if configuration
                    .continuation_notice_backward
                    .replace(String::new())
                    .is_some()
                {
                    return invalid("duplicate backward note continuation notice");
                }
                reject_attributes(&element)?;
            },
            Event::End(element)
                if namespace == NamespaceKind::Text
                    && element.local_name().as_ref() == b"notes-configuration" =>
            {
                break;
            },
            Event::Text(text) => {
                require_whitespace(&text.decode().map_err(xml_error)?, "notes-configuration")?
            },
            Event::Comment(_) | Event::PI(_) => {},
            Event::DocType(_) => return invalid("DOCTYPE is not allowed in notes configuration"),
            Event::Eof => return invalid("unterminated text:notes-configuration"),
            _ => return invalid("unsupported child in text:notes-configuration"),
        }
        buffer.clear();
    }
    Ok(())
}

fn parse_notice(reader: &mut NsReader<&[u8]>, expected_local: &[u8]) -> Result<String> {
    let mut output = String::new();
    let mut buffer = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(xml_error)?;
        let namespace = namespace_kind(&namespace);
        match event {
            Event::Text(value) => append_text(&mut output, &value.decode().map_err(xml_error)?)?,
            Event::CData(value) => append_text(&mut output, &value.decode().map_err(xml_error)?)?,
            Event::GeneralRef(reference) => append_reference(&mut output, &reference)?,
            Event::End(element)
                if namespace == NamespaceKind::Text
                    && element.local_name().as_ref() == expected_local =>
            {
                break;
            },
            Event::Comment(_) | Event::PI(_) => {},
            Event::Eof => return invalid("unterminated note continuation notice"),
            _ => return invalid("note continuation notices may contain text only"),
        }
        buffer.clear();
    }
    Ok(output)
}

fn reject_attributes(element: &BytesStart<'_>) -> Result<()> {
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        if attribute.key.as_ref() != b"xmlns" && !attribute.key.as_ref().starts_with(b"xmlns:") {
            return invalid("note continuation notice cannot contain attributes");
        }
    }
    Ok(())
}

fn insert_configuration(result: &mut Configurations, configuration: Configuration) -> Result<()> {
    let slot = match configuration.note_class {
        Class::Footnote => &mut result.footnote,
        Class::Endnote => &mut result.endnote,
    };
    if slot.replace(configuration).is_some() {
        return invalid("duplicate notes configuration for the same note class");
    }
    Ok(())
}

fn require_styles_scope(styles_depth: Option<usize>, depth: usize) -> Result<()> {
    if styles_depth == Some(depth) {
        Ok(())
    } else {
        invalid("text:notes-configuration must be a direct office:styles child")
    }
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
    xml[..end]
        .rfind('<')
        .ok_or_else(|| Error::InvalidFormat("invalid notes XML event boundary".to_string()))
}

fn locate_configuration(xml: &str, note_class: Class) -> Result<(Option<XmlSpan>, StylesSite)> {
    parse(xml)?;
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
                    && local == b"notes-configuration"
                    && matches!(stack.last(), Some((NamespaceKind::Office, parent)) if parent == b"styles")
                    && parse_attributes(&reader, element)?.note_class == note_class
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
                    && local == b"notes-configuration"
                    && matches!(stack.last(), Some((NamespaceKind::Office, parent)) if parent == b"styles")
                    && parse_attributes(&reader, element)?.note_class == note_class
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
                            .map_err(|_| Error::InvalidFormat("invalid styles QName".to_string()))?
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

/// Insert or replace one note-class configuration without rewriting unrelated XML.
pub fn set_xml(xml: &str, configuration: &Configuration) -> Result<String> {
    configuration.validate()?;
    let (target, site) = locate_configuration(xml, configuration.note_class)?;
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

/// Remove one note-class configuration without rewriting unrelated XML.
pub fn remove_xml(xml: &str, note_class: Class) -> Result<String> {
    let (target, _) = locate_configuration(xml, note_class)?;
    let Some(span) = target else {
        return Ok(xml.to_string());
    };
    Ok(format!("{}{}", &xml[..span.start], &xml[span.end..]))
}

impl OpenDocumentPackage {
    pub fn notes_configurations(&self) -> Result<Configurations> {
        self.styles_xml()?
            .map_or_else(|| Ok(Configurations::default()), |xml| parse(&xml))
    }
}

impl FlatOpenDocument {
    pub fn notes_configurations(&self) -> Result<Configurations> {
        parse(self.xml())
    }
}

fn namespace_kind(namespace: &ResolveResult<'_>) -> NamespaceKind {
    match namespace {
        ResolveResult::Bound(value) if value.as_ref() == OFFICE_NAMESPACE => NamespaceKind::Office,
        ResolveResult::Bound(value) if value.as_ref() == TEXT_NAMESPACE => NamespaceKind::Text,
        ResolveResult::Bound(value) if value.as_ref() == STYLE_NAMESPACE => NamespaceKind::Style,
        _ => NamespaceKind::Other,
    }
}

fn append_reference(
    output: &mut String,
    reference: &quick_xml::events::BytesRef<'_>,
) -> Result<()> {
    if let Some(character) = reference.resolve_char_ref().map_err(xml_error)? {
        let mut encoded = [0u8; 4];
        append_text(output, character.encode_utf8(&mut encoded))
    } else {
        let name = reference.decode().map_err(xml_error)?;
        append_text(
            output,
            match name.as_ref() {
                "amp" => "&",
                "lt" => "<",
                "gt" => ">",
                "apos" => "'",
                "quot" => "\"",
                _ => return invalid("unsupported entity in note continuation notice"),
            },
        )
    }
}

fn append_text(output: &mut String, value: &str) -> Result<()> {
    let size = output
        .len()
        .checked_add(value.len())
        .ok_or_else(|| Error::InvalidFormat("note notice size overflow".to_string()))?;
    if size > MAX_VALUE_BYTES {
        return invalid(format!(
            "note continuation notice exceeds the {MAX_VALUE_BYTES} byte limit"
        ));
    }
    output.push_str(value);
    Ok(())
}

fn parse_bool(value: &str, name: &str) -> Result<bool> {
    match value {
        "true" => Ok(true),
        "false" => Ok(false),
        _ => invalid(format!("{name} must be true or false")),
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

fn validate_value(value: &str, name: &str, allow_empty: bool) -> Result<()> {
    if !allow_empty && value.is_empty() {
        return invalid(format!("{name} must not be empty"));
    }
    if value.len() > MAX_VALUE_BYTES {
        return invalid(format!("{name} exceeds the {MAX_VALUE_BYTES} byte limit"));
    }
    Ok(())
}

fn ncname_start(character: char) -> bool {
    matches!(character,
        'A'..='Z' | '_' | 'a'..='z'
        | '\u{c0}'..='\u{d6}' | '\u{d8}'..='\u{f6}' | '\u{f8}'..='\u{2ff}'
        | '\u{370}'..='\u{37d}' | '\u{37f}'..='\u{1fff}' | '\u{200c}'..='\u{200d}'
        | '\u{2070}'..='\u{218f}' | '\u{2c00}'..='\u{2fef}' | '\u{3001}'..='\u{d7ff}'
        | '\u{f900}'..='\u{fdcf}' | '\u{fdf0}'..='\u{fffd}' | '\u{10000}'..='\u{effff}'
    )
}

fn ncname_continue(character: char) -> bool {
    ncname_start(character)
        || matches!(character, '-' | '.' | '0'..='9' | '\u{b7}' | '\u{300}'..='\u{36f}' | '\u{203f}'..='\u{2040}')
}

fn validate_style_name_ref(value: &str, name: &str) -> Result<()> {
    validate_value(value, name, true)?;
    if value.is_empty() {
        return Ok(());
    }
    let mut characters = value.chars();
    if !characters.next().is_some_and(ncname_start) || !characters.all(ncname_continue) {
        return invalid(format!("{name} must be an NCName or empty"));
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

fn write_notice(output: &mut String, name: &str, value: Option<&str>) {
    let Some(value) = value else { return };
    output.push('<');
    output.push_str(name);
    if value.is_empty() {
        output.push_str("/>");
    } else {
        output.push('>');
        escape_text(output, value);
        output.push_str("</");
        output.push_str(name);
        output.push('>');
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
    Error::InvalidFormat(format!("invalid ODF notes XML: {error}"))
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
    fn parses_and_round_trips_footnote_and_endnote_configurations() {
        // ODF 1.2/1.3 text:notes-configuration grammar; values mirror
        // LibreOffice styles.xml footnote/endnote declarations.
        let xml = styles(
            r#"<t:notes-configuration t:note-class="footnote" t:citation-style-name="Footnote_20_Symbol" t:citation-body-style-name="Footnote_20_anchor" t:default-style-name="Footnote" t:master-page-name="Standard" t:start-value="2" s:num-prefix="[" s:num-suffix="]" s:num-format="a" s:num-letter-sync="true" t:start-numbering-at="chapter" t:footnotes-position="page"><t:note-continuation-notice-forward>Continued &amp; next</t:note-continuation-notice-forward><t:note-continuation-notice-backward><![CDATA[From <previous>]]></t:note-continuation-notice-backward></t:notes-configuration><t:notes-configuration t:note-class="endnote" s:num-format="I" t:start-numbering-at="document"/>"#,
        );
        let configurations = parse(&xml).unwrap();
        let footnote = configurations.get(Class::Footnote).unwrap();
        assert_eq!(footnote.start_value, Some(2));
        assert_eq!(footnote.number_format, Some(LineNumberFormat::LowerAlpha));
        assert_eq!(footnote.letter_sync, Some(true));
        assert_eq!(
            footnote.continuation_notice_forward.as_deref(),
            Some("Continued & next")
        );
        assert_eq!(
            footnote.continuation_notice_backward.as_deref(),
            Some("From <previous>")
        );
        assert!(configurations.get(Class::Endnote).is_some());

        let serialized = configurations.to_xml_fragment().unwrap();
        let reparsed = parse(&styles(&serialized)).unwrap();
        assert_eq!(reparsed, configurations);
    }

    #[test]
    fn rejects_malformed_or_duplicate_note_configuration() {
        for body in [
            r#"<t:notes-configuration/>"#,
            r#"<t:notes-configuration t:note-class="margin"/>"#,
            r#"<t:notes-configuration t:note-class="footnote" s:num-format="1" s:num-letter-sync="true"/>"#,
            r#"<t:notes-configuration t:note-class="footnote" t:start-value="-1"/>"#,
            r#"<t:notes-configuration t:note-class="footnote" t:start-numbering-at="section"/>"#,
            r#"<t:notes-configuration t:note-class="footnote"><t:note-continuation-notice-forward/><t:note-continuation-notice-forward/></t:notes-configuration>"#,
            r#"<t:notes-configuration t:note-class="footnote"><t:note-continuation-notice-forward><t:span/></t:note-continuation-notice-forward></t:notes-configuration>"#,
            r#"<t:notes-configuration t:note-class="footnote"/><t:notes-configuration t:note-class="footnote"/>"#,
        ] {
            assert!(parse(&styles(body)).is_err(), "{body}");
        }
    }

    #[test]
    fn rejects_misplaced_note_configuration() {
        let xml = format!(
            r#"<o:document-content xmlns:o="{OFFICE}" xmlns:t="{TEXT}"><o:body><t:notes-configuration t:note-class="footnote"/></o:body></o:document-content>"#
        );
        assert!(parse(&xml).is_err());
    }

    #[test]
    fn exhaustive_enums_and_number_formats_round_trip() {
        for note_class in Class::ALL {
            for scope in NumberingScope::ALL {
                for position in Position::ALL {
                    let mut value = Configuration::new(note_class);
                    value.citation_style_name = Some(String::new());
                    value.default_style_name = Some("Style_1".to_string());
                    value.number_format = Some(LineNumberFormat::LowerAlpha);
                    value.letter_sync = Some(false);
                    value.start_numbering_at = Some(scope);
                    value.footnotes_position = Some(position);
                    let parsed = parse(&styles(&value.to_xml().unwrap())).unwrap();
                    assert_eq!(parsed.get(note_class), Some(&value));
                }
            }
        }
        for format in [
            LineNumberFormat::Empty,
            LineNumberFormat::Arabic,
            LineNumberFormat::LowerRoman,
            LineNumberFormat::UpperRoman,
            LineNumberFormat::LowerAlpha,
            LineNumberFormat::UpperAlpha,
            LineNumberFormat::Custom("①".to_string()),
        ] {
            let mut value = Configuration::new(Class::Footnote);
            value.letter_sync = matches!(
                format,
                LineNumberFormat::LowerAlpha | LineNumberFormat::UpperAlpha
            )
            .then_some(true);
            value.number_format = Some(format);
            assert_eq!(
                parse(&styles(&value.to_xml().unwrap())).unwrap().footnote,
                Some(value)
            );
        }
    }

    #[test]
    fn accepts_interleaved_notice_order_and_real_libreoffice() {
        let reverse = styles(
            r#"<t:notes-configuration t:note-class="footnote"><t:note-continuation-notice-backward>back</t:note-continuation-notice-backward><t:note-continuation-notice-forward>forward</t:note-continuation-notice-forward></t:notes-configuration>"#,
        );
        let parsed = parse(&reverse).unwrap();
        assert_eq!(
            parsed
                .footnote
                .as_ref()
                .unwrap()
                .continuation_notice_forward
                .as_deref(),
            Some("forward")
        );
        assert_eq!(
            parsed
                .footnote
                .as_ref()
                .unwrap()
                .continuation_notice_backward
                .as_deref(),
            Some("back")
        );

        let fixture =
            include_str!("../../../test-data/libreoffice-core/sw/qa/uitest/data/tdf145178.fodt");
        let real = parse(fixture).unwrap();
        assert_eq!(
            real.footnote.as_ref().unwrap().number_format,
            Some(LineNumberFormat::Arabic)
        );
        assert_eq!(
            real.endnote.as_ref().unwrap().number_format,
            Some(LineNumberFormat::LowerRoman)
        );
        let flat = FlatOpenDocument::from_bytes(fixture.as_bytes().to_vec()).unwrap();
        assert_eq!(flat.notes_configurations().unwrap(), real);
    }

    #[test]
    fn rejects_style_refs_namespaces_booleans_and_caps() {
        for body in [
            r#"<t:notes-configuration t:note-class="footnote" t:citation-style-name="bad name"/>"#,
            r#"<t:notes-configuration t:note-class="footnote" t:default-style-name="1bad"/>"#,
            r#"<t:notes-configuration t:note-class="footnote" s:num-format="a" s:num-letter-sync="1"/>"#,
            r#"<t:notes-configuration t:note-class="footnote" s:num-format="custom" s:num-letter-sync="true"/>"#,
            r#"<t:notes-configuration t:note-class="footnote"><t:note-continuation-notice-forward t:bad="1"/></t:notes-configuration>"#,
            r#"<t:notes-configuration xmlns:x="urn:wrong" t:note-class="footnote" x:note-class="endnote"/>"#,
        ] {
            assert!(parse(&styles(body)).is_err(), "accepted {body}");
        }
        let wrong_namespace = format!(
            r#"<o:document-styles xmlns:o="{OFFICE}" xmlns:t="{TEXT}" xmlns:x="urn:wrong"><o:styles><x:notes-configuration t:note-class="footnote"/></o:styles></o:document-styles>"#
        );
        assert!(parse(&wrong_namespace).is_err());
        let mut capped = Configuration::new(Class::Footnote);
        capped.continuation_notice_forward = Some("x".repeat(MAX_VALUE_BYTES + 1));
        assert!(capped.validate().is_err());
    }

    #[test]
    fn lossless_mutation_and_builder_package_access() {
        use crate::Builder;

        let original = styles(
            r#"<!--keep--><style:list-style xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" style:name="L"/>"#,
        );
        let mut value = Configuration::new(Class::Footnote);
        value.start_value = Some(2);
        let inserted = set_xml(&original, &value).unwrap();
        assert!(inserted.contains("<!--keep--><style:list-style"));
        value.start_value = Some(3);
        let replaced = set_xml(&inserted, &value).unwrap();
        assert!(replaced.contains("text:start-value=\"3\""));
        assert!(!replaced.contains("text:start-value=\"2\""));
        assert_eq!(remove_xml(&replaced, Class::Footnote).unwrap(), original);

        let mut builder = Builder::new();
        builder.set_notes_configuration(value.clone()).unwrap();
        let package = OpenDocumentPackage::from_bytes(builder.build().unwrap()).unwrap();
        assert_eq!(
            package.notes_configurations().unwrap().footnote,
            Some(value)
        );
    }
}
