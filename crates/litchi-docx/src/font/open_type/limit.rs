//! Typed admission limits for detached Word 2010 OpenType XML.

use super::validation::{
    MAX_STYLE_SETS, MAX_XML_ATTRIBUTE_BYTES, MAX_XML_ATTRIBUTES, MAX_XML_BYTES, MAX_XML_DEPTH,
    MAX_XML_NODES,
};
use crate::{Error, Result};
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
use quick_xml::Reader;
use quick_xml::events::Event;

/// Caller-selected bounds for the detached OpenType run-property codec.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    xml_bytes: usize,
    depth: usize,
    nodes: usize,
    attributes: usize,
    attribute_bytes: usize,
    style_sets: usize,
}

impl Limits {
    /// The standard bounded OpenType policy.
    pub const DEFAULT: Self = Self {
        xml_bytes: MAX_XML_BYTES,
        depth: MAX_XML_DEPTH,
        nodes: MAX_XML_NODES,
        attributes: MAX_XML_ATTRIBUTES,
        attribute_bytes: MAX_XML_ATTRIBUTE_BYTES,
        style_sets: MAX_STYLE_SETS,
    };

    /// Construct a finite detached OpenType policy.
    #[must_use]
    pub const fn new(
        xml_bytes: usize,
        depth: usize,
        nodes: usize,
        attributes: usize,
        attribute_bytes: usize,
        style_sets: usize,
    ) -> Option<Self> {
        if xml_bytes == 0
            || xml_bytes > MAX_XML_BYTES
            || depth == 0
            || depth > MAX_XML_DEPTH
            || nodes == 0
            || nodes > MAX_XML_NODES
            || attributes == 0
            || attributes > MAX_XML_ATTRIBUTES
            || attribute_bytes == 0
            || attribute_bytes > MAX_XML_ATTRIBUTE_BYTES
            || style_sets == 0
            || style_sets > MAX_STYLE_SETS
        {
            None
        } else {
            Some(Self {
                xml_bytes,
                depth,
                nodes,
                attributes,
                attribute_bytes,
                style_sets,
            })
        }
    }

    #[must_use]
    pub const fn xml_bytes(self) -> usize {
        self.xml_bytes
    }

    #[must_use]
    pub const fn depth(self) -> usize {
        self.depth
    }

    #[must_use]
    pub const fn nodes(self) -> usize {
        self.nodes
    }

    #[must_use]
    pub const fn attributes(self) -> usize {
        self.attributes
    }

    #[must_use]
    pub const fn attribute_bytes(self) -> usize {
        self.attribute_bytes
    }

    #[must_use]
    pub const fn style_sets(self) -> usize {
        self.style_sets
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self::DEFAULT
    }
}

/// Check XML shape and retained scalar budgets before the codec allocates its
/// semantic model or a source-preserving `Arc`.
pub(crate) fn admit_xml(xml: &[u8], limits: &Limits) -> Result<()> {
    if xml.len() > limits.xml_bytes() {
        return Err(limit(
            crate::font::LimitKind::OpenTypeXmlBytes,
            xml.len(),
            limits.xml_bytes(),
        ));
    }
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut nodes = 0usize;
    let mut attributes = 0usize;
    let mut attribute_bytes = 0usize;
    let mut style_sets = 0usize;

    loop {
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(|error| Error::Xml(error.to_string()))?;
        match event {
            Event::Start(element) => {
                depth = depth.saturating_add(1);
                nodes = nodes.saturating_add(1);
                check_shape(depth, nodes, attributes, limits)?;
                count_attributes(&element, &mut attributes, &mut attribute_bytes, limits)?;
                if element.local_name().as_ref() == b"styleSet" {
                    style_sets = style_sets.saturating_add(1);
                    if style_sets > limits.style_sets() {
                        return Err(limit(
                            crate::font::LimitKind::OpenTypeStyleSets,
                            style_sets,
                            limits.style_sets(),
                        ));
                    }
                }
            },
            Event::Empty(element) => {
                nodes = nodes.saturating_add(1);
                check_shape(depth.saturating_add(1), nodes, attributes, limits)?;
                count_attributes(&element, &mut attributes, &mut attribute_bytes, limits)?;
                if element.local_name().as_ref() == b"styleSet" {
                    style_sets = style_sets.saturating_add(1);
                    if style_sets > limits.style_sets() {
                        return Err(limit(
                            crate::font::LimitKind::OpenTypeStyleSets,
                            style_sets,
                            limits.style_sets(),
                        ));
                    }
                }
            },
            Event::End(_) => {
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::Invalid("OpenType XML has an unexpected closing element".to_owned())
                })?;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(Error::Invalid(
                    "OpenType XML contains a forbidden declaration".to_owned(),
                ));
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if depth != 0 {
        return Err(Error::Invalid("OpenType XML is unterminated".to_owned()));
    }
    Ok(())
}

fn count_attributes(
    element: &quick_xml::events::BytesStart<'_>,
    attributes: &mut usize,
    attribute_bytes: &mut usize,
    limits: &Limits,
) -> Result<()> {
    for attribute in element.unchecked_attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        *attributes = attributes.saturating_add(1);
        if *attributes > limits.attributes() {
            return Err(limit(
                crate::font::LimitKind::OpenTypeXmlAttributes,
                *attributes,
                limits.attributes(),
            ));
        }
        *attribute_bytes = attribute_bytes
            .checked_add(
                attribute
                    .key
                    .as_ref()
                    .len()
                    .saturating_add(attribute.value.len()),
            )
            .ok_or_else(|| {
                limit(
                    crate::font::LimitKind::OpenTypeXmlAttributeBytes,
                    usize::MAX,
                    limits.attribute_bytes(),
                )
            })?;
        if *attribute_bytes > limits.attribute_bytes() {
            return Err(limit(
                crate::font::LimitKind::OpenTypeXmlAttributeBytes,
                *attribute_bytes,
                limits.attribute_bytes(),
            ));
        }
    }
    Ok(())
}

fn check_shape(depth: usize, nodes: usize, attributes: usize, limits: &Limits) -> Result<()> {
    if depth > limits.depth() {
        return Err(limit(
            crate::font::LimitKind::OpenTypeXmlDepth,
            depth,
            limits.depth(),
        ));
    }
    if nodes > limits.nodes() {
        return Err(limit(
            crate::font::LimitKind::OpenTypeXmlNodes,
            nodes,
            limits.nodes(),
        ));
    }
    if attributes > limits.attributes() {
        return Err(limit(
            crate::font::LimitKind::OpenTypeXmlAttributes,
            attributes,
            limits.attributes(),
        ));
    }
    Ok(())
}

fn limit(kind: crate::font::LimitKind, actual: usize, maximum: usize) -> Error {
    Error::FontLimit {
        kind,
        actual,
        maximum,
    }
}
