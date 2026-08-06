//! Streaming text extraction for paragraph and run content.

use crate::error::{Error, Result};
use litchi_ooxml_common::xml::decode_xml_reference;
use quick_xml::XmlVersion;
use quick_xml::events::Event;
use quick_xml::name::{QName, ResolveResult};
use quick_xml::reader::NsReader;

use super::super::model::Paragraph;
use super::xml::is_fragment_word_name;

/// Maximum nesting depth accepted when extracting paragraph text.
const MAX_TEXT_SCAN_DEPTH: usize = 128;
/// Maximum number of elements scanned while extracting paragraph text.
const MAX_TEXT_SCAN_NODES: usize = 1_000_000;

pub(crate) fn extract_word_text(xml_bytes: &[u8]) -> Result<String> {
    let mut reader = NsReader::from_reader(xml_bytes);
    let mut result = String::with_capacity(xml_bytes.len() / 8);
    let mut fragment_prefix: Option<Option<Vec<u8>>> = None;
    let mut depth = 0usize;
    let mut nodes = 0usize;
    let mut text_depth = None;

    loop {
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;

        if fragment_prefix.is_none()
            && let Event::Start(element) = &event
            && !matches!(namespace, ResolveResult::Bound(_))
        {
            fragment_prefix = Some(
                element
                    .name()
                    .prefix()
                    .map(|prefix| prefix.into_inner().to_vec()),
            );
        }

        match event {
            Event::Start(element) => {
                nodes = nodes.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("Word XML element counter overflow".to_string())
                })?;
                if nodes > MAX_TEXT_SCAN_NODES {
                    return Err(Error::InvalidFormat(format!(
                        "Word XML exceeds {MAX_TEXT_SCAN_NODES} elements"
                    )));
                }
                depth = depth.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("Word XML nesting is too deep".to_string())
                })?;
                if depth > MAX_TEXT_SCAN_DEPTH {
                    return Err(Error::InvalidFormat(format!(
                        "Word XML nesting exceeds the {MAX_TEXT_SCAN_DEPTH} depth limit"
                    )));
                }
                if text_depth.is_none()
                    && is_fragment_word_name(&namespace, element.name(), b"t", &fragment_prefix)
                {
                    text_depth = Some(depth);
                } else if let Some(character) =
                    word_special_character(&namespace, element.name(), &fragment_prefix)
                {
                    result.push(character);
                }
            },
            Event::Empty(element) => {
                nodes = nodes.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("Word XML element counter overflow".to_string())
                })?;
                if nodes > MAX_TEXT_SCAN_NODES {
                    return Err(Error::InvalidFormat(format!(
                        "Word XML exceeds {MAX_TEXT_SCAN_NODES} elements"
                    )));
                }
                if let Some(character) =
                    word_special_character(&namespace, element.name(), &fragment_prefix)
                {
                    result.push(character);
                }
            },
            Event::Text(text) if text_depth.is_some() => {
                let decoded = text
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| Error::Xml(error.to_string()))?;
                let unescaped = quick_xml::escape::unescape(&decoded)
                    .map_err(|error| Error::Xml(error.to_string()))?;
                result.push_str(&unescaped);
            },
            Event::CData(text) if text_depth.is_some() => {
                let decoded = text
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| Error::Xml(error.to_string()))?;
                result.push_str(&decoded);
            },
            Event::GeneralRef(reference) if text_depth.is_some() => {
                result.push_str(&decode_xml_reference(&reference)?);
            },
            Event::End(element) => {
                if text_depth == Some(depth)
                    && is_fragment_word_name(&namespace, element.name(), b"t", &fragment_prefix)
                {
                    text_depth = None;
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| Error::InvalidFormat("invalid Word XML nesting".to_string()))?;
            },
            Event::Eof if depth != 0 || text_depth.is_some() => {
                return Err(Error::InvalidFormat("unterminated Word XML".to_string()));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    result.shrink_to_fit();
    Ok(result)
}

fn word_special_character(
    namespace: &ResolveResult<'_>,
    name: QName<'_>,
    fragment_prefix: &Option<Option<Vec<u8>>>,
) -> Option<char> {
    if is_fragment_word_name(namespace, name, b"tab", fragment_prefix) {
        Some('\t')
    } else if is_fragment_word_name(namespace, name, b"br", fragment_prefix)
        || is_fragment_word_name(namespace, name, b"cr", fragment_prefix)
    {
        Some('\n')
    } else if is_fragment_word_name(namespace, name, b"noBreakHyphen", fragment_prefix) {
        Some('\u{2011}')
    } else if is_fragment_word_name(namespace, name, b"softHyphen", fragment_prefix) {
        Some('\u{00ad}')
    } else {
        None
    }
}

impl Paragraph {
    /// Get the text content of this paragraph.
    ///
    /// Concatenates all text from all runs in the paragraph.
    ///
    /// # Performance
    ///
    /// Uses streaming XML parsing with pre-allocated buffer to extract text efficiently.
    pub fn text(&self) -> Result<String> {
        extract_word_text(self.xml_bytes())
    }
}
