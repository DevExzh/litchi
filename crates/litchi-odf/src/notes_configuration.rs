//! ODF footnote and endnote configuration metadata.

use std::collections::HashSet;

use litchi_core::{Error, Result};
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::ResolveResult,
    reader::NsReader,
};

use crate::line_numbering::OdfLineNumberFormat;

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
pub enum OdfNoteClass {
    Footnote,
    Endnote,
}

impl OdfNoteClass {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "footnote" => Ok(Self::Footnote),
            "endnote" => Ok(Self::Endnote),
            _ => invalid(format!("unsupported text:note-class '{value}'")),
        }
    }

    fn as_str(self) -> &'static str {
        match self {
            Self::Footnote => "footnote",
            Self::Endnote => "endnote",
        }
    }
}

/// Scope at which note numbering restarts.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum OdfNoteNumberingScope {
    Document,
    Chapter,
    Page,
}

impl OdfNoteNumberingScope {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "document" => Ok(Self::Document),
            "chapter" => Ok(Self::Chapter),
            "page" => Ok(Self::Page),
            _ => invalid(format!("unsupported text:start-numbering-at '{value}'")),
        }
    }

    fn as_str(self) -> &'static str {
        match self {
            Self::Document => "document",
            Self::Chapter => "chapter",
            Self::Page => "page",
        }
    }
}

/// Placement of footnotes in the document.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum OdfFootnotePosition {
    Text,
    Page,
    Section,
    Document,
}

impl OdfFootnotePosition {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "text" => Ok(Self::Text),
            "page" => Ok(Self::Page),
            "section" => Ok(Self::Section),
            "document" => Ok(Self::Document),
            _ => invalid(format!("unsupported text:footnotes-position '{value}'")),
        }
    }

    fn as_str(self) -> &'static str {
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
pub struct OdfNotesConfiguration {
    pub note_class: OdfNoteClass,
    pub citation_style_name: Option<String>,
    pub citation_body_style_name: Option<String>,
    pub default_style_name: Option<String>,
    pub master_page_name: Option<String>,
    pub start_value: Option<u64>,
    pub number_prefix: Option<String>,
    pub number_suffix: Option<String>,
    pub number_format: Option<OdfLineNumberFormat>,
    pub letter_sync: Option<bool>,
    pub start_numbering_at: Option<OdfNoteNumberingScope>,
    pub footnotes_position: Option<OdfFootnotePosition>,
    pub continuation_notice_forward: Option<String>,
    pub continuation_notice_backward: Option<String>,
}

impl OdfNotesConfiguration {
    pub fn validate(&self) -> Result<()> {
        if self.letter_sync.is_some()
            && !matches!(
                self.number_format,
                Some(OdfLineNumberFormat::LowerAlpha | OdfLineNumberFormat::UpperAlpha)
            )
        {
            return invalid(
                "style:num-letter-sync requires style:num-format 'a' or 'A'",
            );
        }
        for (value, name, allow_empty) in [
            (
                self.citation_style_name.as_deref(),
                "text:citation-style-name",
                false,
            ),
            (
                self.citation_body_style_name.as_deref(),
                "text:citation-body-style-name",
                false,
            ),
            (
                self.default_style_name.as_deref(),
                "text:default-style-name",
                false,
            ),
            (
                self.master_page_name.as_deref(),
                "text:master-page-name",
                false,
            ),
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
            self.number_format.as_ref().map(OdfLineNumberFormat::as_str),
        );
        write_bool_attr(&mut output, "style:num-letter-sync", self.letter_sync);
        write_attr(
            &mut output,
            "text:start-numbering-at",
            self.start_numbering_at.map(OdfNoteNumberingScope::as_str),
        );
        write_attr(
            &mut output,
            "text:footnotes-position",
            self.footnotes_position.map(OdfFootnotePosition::as_str),
        );
        if self.continuation_notice_forward.is_none()
            && self.continuation_notice_backward.is_none()
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
pub struct OdfNotesConfigurations {
    pub footnote: Option<OdfNotesConfiguration>,
    pub endnote: Option<OdfNotesConfiguration>,
}

impl OdfNotesConfigurations {
    pub fn get(&self, note_class: OdfNoteClass) -> Option<&OdfNotesConfiguration> {
        match note_class {
            OdfNoteClass::Footnote => self.footnote.as_ref(),
            OdfNoteClass::Endnote => self.endnote.as_ref(),
        }
    }

    pub fn validate(&self) -> Result<()> {
        if let Some(configuration) = &self.footnote {
            if configuration.note_class != OdfNoteClass::Footnote {
                return invalid("footnote slot contains an endnote configuration");
            }
            configuration.validate()?;
        }
        if let Some(configuration) = &self.endnote {
            if configuration.note_class != OdfNoteClass::Endnote {
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
pub fn parse_notes_configurations(xml: &str) -> Result<OdfNotesConfigurations> {
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
    let mut result = OdfNotesConfigurations::default();
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
                require_styles_scope(styles_content_depth, depth)?;
                let mut configuration = parse_attributes(&reader, &element)?;
                parse_notices(&mut reader, &mut configuration)?;
                configuration.validate()?;
                insert_configuration(&mut result, configuration)?;
            }
            Event::Empty(element)
                if namespace == NamespaceKind::Text
                    && element.local_name().as_ref() == b"notes-configuration" =>
            {
                require_styles_scope(styles_content_depth, depth)?;
                let configuration = parse_attributes(&reader, &element)?;
                configuration.validate()?;
                insert_configuration(&mut result, configuration)?;
            }
            Event::Start(element) => {
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| Error::InvalidFormat("ODF XML depth overflow".to_string()))?;
                if depth > MAX_XML_DEPTH {
                    return invalid(format!("ODF XML exceeds the {MAX_XML_DEPTH} depth limit"));
                }
                if namespace == NamespaceKind::Office
                    && element.local_name().as_ref() == b"styles"
                {
                    if styles_content_depth.replace(depth).is_some() {
                        return invalid("ODF XML contains nested office:styles elements");
                    }
                }
            }
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
            }
            Event::DocType(_) => return invalid("DOCTYPE is not allowed in ODF note metadata"),
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    Ok(result)
}

fn parse_attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<OdfNotesConfiguration> {
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
            (NamespaceKind::Text, b"note-class") => {
                note_class = Some(OdfNoteClass::parse(&value)?)
            }
            (NamespaceKind::Text, b"citation-style-name") => citation_style_name = Some(value),
            (NamespaceKind::Text, b"citation-body-style-name") => {
                citation_body_style_name = Some(value)
            }
            (NamespaceKind::Text, b"default-style-name") => default_style_name = Some(value),
            (NamespaceKind::Text, b"master-page-name") => master_page_name = Some(value),
            (NamespaceKind::Text, b"start-value") => {
                start_value = Some(parse_nonnegative_integer(&value, "text:start-value")?)
            }
            (NamespaceKind::Style, b"num-prefix") => number_prefix = Some(value),
            (NamespaceKind::Style, b"num-suffix") => number_suffix = Some(value),
            (NamespaceKind::Style, b"num-format") => {
                number_format = Some(OdfLineNumberFormat::parse(value)?)
            }
            (NamespaceKind::Style, b"num-letter-sync") => {
                letter_sync = Some(parse_bool(&value, "style:num-letter-sync")?)
            }
            (NamespaceKind::Text, b"start-numbering-at") => {
                start_numbering_at = Some(OdfNoteNumberingScope::parse(&value)?)
            }
            (NamespaceKind::Text, b"footnotes-position") => {
                footnotes_position = Some(OdfFootnotePosition::parse(&value)?)
            }
            _ => return invalid("unsupported text:notes-configuration attribute"),
        }
    }
    let note_class = note_class
        .ok_or_else(|| Error::InvalidFormat("notes configuration requires text:note-class".to_string()))?;
    Ok(OdfNotesConfiguration {
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

fn parse_notices(
    reader: &mut NsReader<&[u8]>,
    configuration: &mut OdfNotesConfiguration,
) -> Result<()> {
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
                if configuration.continuation_notice_forward.is_some() {
                    return invalid("duplicate forward note continuation notice");
                }
                configuration.continuation_notice_forward =
                    Some(parse_notice(reader, b"note-continuation-notice-forward")?);
            }
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
            }
            Event::Start(element)
                if namespace == NamespaceKind::Text
                    && element.local_name().as_ref() == b"note-continuation-notice-backward" =>
            {
                if configuration.continuation_notice_backward.is_some() {
                    return invalid("duplicate backward note continuation notice");
                }
                configuration.continuation_notice_backward =
                    Some(parse_notice(reader, b"note-continuation-notice-backward")?);
            }
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
            }
            Event::End(element)
                if namespace == NamespaceKind::Text
                    && element.local_name().as_ref() == b"notes-configuration" =>
            {
                break;
            }
            Event::Text(text) => require_whitespace(
                &text.decode().map_err(xml_error)?,
                "notes-configuration",
            )?,
            Event::Comment(_) | Event::PI(_) => {}
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
            }
            Event::Comment(_) | Event::PI(_) => {}
            Event::Eof => return invalid("unterminated note continuation notice"),
            _ => return invalid("note continuation notices may contain text only"),
        }
        buffer.clear();
    }
    Ok(output)
}

fn reject_attributes(element: &BytesStart<'_>) -> Result<()> {
    if element.attributes().with_checks(true).next().is_some() {
        invalid("note continuation notice cannot contain attributes")
    } else {
        Ok(())
    }
}

fn insert_configuration(
    result: &mut OdfNotesConfigurations,
    configuration: OdfNotesConfiguration,
) -> Result<()> {
    let slot = match configuration.note_class {
        OdfNoteClass::Footnote => &mut result.footnote,
        OdfNoteClass::Endnote => &mut result.endnote,
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

fn namespace_kind(namespace: &ResolveResult<'_>) -> NamespaceKind {
    match namespace {
        ResolveResult::Bound(value) if value.as_ref() == OFFICE_NAMESPACE => NamespaceKind::Office,
        ResolveResult::Bound(value) if value.as_ref() == TEXT_NAMESPACE => NamespaceKind::Text,
        ResolveResult::Bound(value) if value.as_ref() == STYLE_NAMESPACE => NamespaceKind::Style,
        _ => NamespaceKind::Other,
    }
}

fn append_reference(output: &mut String, reference: &quick_xml::events::BytesRef<'_>) -> Result<()> {
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
    write_attr(output, name, value.map(|value| if value { "true" } else { "false" }));
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
            r#"<t:notes-configuration t:note-class="footnote" t:citation-style-name="Footnote &amp; Symbol" t:citation-body-style-name="Footnote anchor" t:default-style-name="Footnote" t:master-page-name="Standard" t:start-value="2" s:num-prefix="[" s:num-suffix="]" s:num-format="a" s:num-letter-sync="1" t:start-numbering-at="chapter" t:footnotes-position="page"><t:note-continuation-notice-forward>Continued &amp; next</t:note-continuation-notice-forward><t:note-continuation-notice-backward><![CDATA[From <previous>]]></t:note-continuation-notice-backward></t:notes-configuration><t:notes-configuration t:note-class="endnote" s:num-format="I" t:start-numbering-at="document"/>"#,
        );
        let configurations = parse_notes_configurations(&xml).unwrap();
        let footnote = configurations.get(OdfNoteClass::Footnote).unwrap();
        assert_eq!(footnote.start_value, Some(2));
        assert_eq!(footnote.number_format, Some(OdfLineNumberFormat::LowerAlpha));
        assert_eq!(footnote.letter_sync, Some(true));
        assert_eq!(
            footnote.continuation_notice_forward.as_deref(),
            Some("Continued & next")
        );
        assert_eq!(
            footnote.continuation_notice_backward.as_deref(),
            Some("From <previous>")
        );
        assert!(configurations.get(OdfNoteClass::Endnote).is_some());

        let serialized = configurations.to_xml_fragment().unwrap();
        let reparsed = parse_notes_configurations(&styles(&serialized)).unwrap();
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
            assert!(parse_notes_configurations(&styles(body)).is_err(), "{body}");
        }
    }

    #[test]
    fn rejects_misplaced_note_configuration() {
        let xml = format!(
            r#"<o:document-content xmlns:o="{OFFICE}" xmlns:t="{TEXT}"><o:body><t:notes-configuration t:note-class="footnote"/></o:body></o:document-content>"#
        );
        assert!(parse_notes_configurations(&xml).is_err());
    }
}
