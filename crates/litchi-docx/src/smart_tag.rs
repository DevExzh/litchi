//! WordprocessingML smart-tag metadata.

use crate::error::{Error, Result};
use crate::paragraph::extract_word_text;
use litchi_core::XmlSlice;
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::reader::Reader;

/// One custom attribute attached to a smart tag.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SmartTagAttribute {
    /// Optional namespace URI for the attribute.
    pub uri: Option<String>,
    /// Attribute name.
    pub name: String,
    /// Attribute value.
    pub value: String,
}

/// A run-level Word smart tag and its custom metadata.
#[derive(Debug, Clone)]
pub struct SmartTag {
    /// Optional namespace URI identifying the smart-tag vocabulary.
    pub uri: Option<String>,
    /// Required smart-tag element name.
    pub element: String,
    /// Custom attributes from the optional `smartTagPr` container.
    pub attributes: Vec<SmartTagAttribute>,
    xml: XmlSlice,
}

impl SmartTag {
    pub(crate) fn parse(xml: XmlSlice) -> Result<Self> {
        let mut reader = Reader::from_reader(xml.as_bytes());
        let mut uri = None;
        let mut element = None;
        let mut attributes = Vec::new();
        let mut saw_root = false;
        let mut properties_depth = None;
        let mut depth = 0usize;

        loop {
            let decoder = reader.decoder();
            match reader
                .read_event()
                .map_err(|error| Error::Xml(error.to_string()))?
            {
                Event::Start(start) => {
                    depth = depth.checked_add(1).ok_or_else(|| {
                        Error::InvalidFormat("smart-tag XML nesting is too deep".into())
                    })?;
                    match start.local_name().as_ref() {
                        b"smartTag" if depth == 1 && !saw_root => {
                            saw_root = true;
                            uri = optional_attribute(&start, b"uri", decoder)?;
                            element = optional_attribute(&start, b"element", decoder)?;
                        },
                        b"smartTagPr" if depth == 2 && saw_root && properties_depth.is_none() => {
                            properties_depth = Some(depth);
                        },
                        b"attr" if properties_depth.is_some_and(|value| depth == value + 1) => {
                            attributes.push(parse_attribute(&start, decoder)?);
                        },
                        _ => {},
                    }
                },
                Event::Empty(start)
                    if start.local_name().as_ref() == b"attr"
                        && properties_depth.is_some_and(|value| depth == value) =>
                {
                    attributes.push(parse_attribute(&start, decoder)?);
                },
                Event::Empty(start)
                    if depth == 0 && start.local_name().as_ref() == b"smartTag" && !saw_root =>
                {
                    saw_root = true;
                    uri = optional_attribute(&start, b"uri", decoder)?;
                    element = optional_attribute(&start, b"element", decoder)?;
                },
                Event::End(end) => {
                    if properties_depth == Some(depth) && end.local_name().as_ref() == b"smartTagPr"
                    {
                        properties_depth = None;
                    }
                    depth = depth.checked_sub(1).ok_or_else(|| {
                        Error::InvalidFormat("invalid smart-tag XML nesting".into())
                    })?;
                },
                Event::Eof if depth != 0 => {
                    return Err(Error::InvalidFormat("unterminated smart-tag XML".into()));
                },
                Event::Eof => break,
                _ => {},
            }
        }

        if !saw_root {
            return Err(Error::InvalidFormat(
                "smart-tag XML has no smartTag root".into(),
            ));
        }
        Ok(Self {
            uri,
            element: element.ok_or_else(|| {
                Error::InvalidFormat("Word smart tag is missing its element attribute".into())
            })?,
            attributes,
            xml,
        })
    }

    /// Extract the visible text contained by this smart tag.
    pub fn text(&self) -> Result<String> {
        extract_word_text(self.xml.as_bytes())
    }

    /// Return the original smart-tag XML fragment.
    #[inline]
    pub fn raw_xml(&self) -> &[u8] {
        self.xml.as_bytes()
    }
}

impl PartialEq for SmartTag {
    fn eq(&self, other: &Self) -> bool {
        self.uri == other.uri
            && self.element == other.element
            && self.attributes == other.attributes
            && self.raw_xml() == other.raw_xml()
    }
}

impl Eq for SmartTag {}

fn parse_attribute(
    start: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
) -> Result<SmartTagAttribute> {
    Ok(SmartTagAttribute {
        uri: optional_attribute(start, b"uri", decoder)?,
        name: optional_attribute(start, b"name", decoder)?.ok_or_else(|| {
            Error::InvalidFormat("Word smart-tag property is missing its name".into())
        })?,
        value: optional_attribute(start, b"val", decoder)?.ok_or_else(|| {
            Error::InvalidFormat("Word smart-tag property is missing its value".into())
        })?,
    })
}

fn optional_attribute(
    start: &BytesStart<'_>,
    name: &[u8],
    decoder: quick_xml::encoding::Decoder,
) -> Result<Option<String>> {
    let mut value = None;
    for attribute in start.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.key.local_name().as_ref() != name {
            continue;
        }
        if value.is_some() {
            return Err(Error::InvalidFormat(format!(
                "Word smart tag contains duplicate '{}' attributes",
                String::from_utf8_lossy(name)
            )));
        }
        value = Some(
            attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                .map_err(|error| Error::Xml(error.to_string()))?
                .into_owned(),
        );
    }
    Ok(value)
}
