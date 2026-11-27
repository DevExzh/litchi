//! Mutable, package-neutral WordprocessingML document codec model.
//!
//! [`Tree`] owns one complete main-document XML payload and stages the
//! source-preserving document operations in memory. It is deliberately below
//! the package owner: output is bytes, while package publication goes through
//! `Package::edit`.

use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
use quick_xml::Reader;
use quick_xml::events::Event;

use super::transaction;

/// Caller-selected admission and output policy for the raw document codec.
///
/// The bounds are checked before the document layout is retained. Their hard
/// ceilings are format safety bounds; a caller may choose any smaller policy,
/// but output and replay cannot widen the policy retained by a tree.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    max_xml_bytes: usize,
    max_depth: usize,
    max_nodes: usize,
    max_attributes: usize,
    max_attribute_bytes: usize,
    max_operations: usize,
    max_replacement_text_bytes: usize,
}

impl Limits {
    /// Construct a bounded raw document policy.
    pub fn new(
        max_xml_bytes: usize,
        max_depth: usize,
        max_nodes: usize,
        max_attributes: usize,
        max_attribute_bytes: usize,
        max_operations: usize,
        max_replacement_text_bytes: usize,
    ) -> crate::document::Result<Self> {
        let limits = Self {
            max_xml_bytes,
            max_depth,
            max_nodes,
            max_attributes,
            max_attribute_bytes,
            max_operations,
            max_replacement_text_bytes,
        };
        limits.validate()?;
        Ok(limits)
    }

    /// Maximum accepted XML input or output bytes.
    #[must_use]
    pub const fn max_xml_bytes(self) -> usize {
        self.max_xml_bytes
    }

    /// Maximum accepted XML nesting depth.
    #[must_use]
    pub const fn max_depth(self) -> usize {
        self.max_depth
    }

    /// Maximum number of XML elements retained by the tree.
    #[must_use]
    pub const fn max_nodes(self) -> usize {
        self.max_nodes
    }

    /// Maximum number of attributes on one XML element.
    #[must_use]
    pub const fn max_attributes(self) -> usize {
        self.max_attributes
    }

    /// Maximum raw attribute bytes on one XML element.
    #[must_use]
    pub const fn max_attribute_bytes(self) -> usize {
        self.max_attribute_bytes
    }

    /// Maximum semantic operations staged by one tree.
    #[must_use]
    pub const fn max_operations(self) -> usize {
        self.max_operations
    }

    /// Maximum replacement text bytes staged by one tree.
    #[must_use]
    pub const fn max_replacement_text_bytes(self) -> usize {
        self.max_replacement_text_bytes
    }

    pub(crate) fn validate(self) -> transaction::Result<()> {
        if self.max_xml_bytes == 0 || self.max_xml_bytes > transaction::MAX_DOCUMENT_XML_BYTES {
            return Err(limit(
                crate::document::LimitKind::XmlBytes,
                self.max_xml_bytes,
                transaction::MAX_DOCUMENT_XML_BYTES,
            ));
        }
        if self.max_depth == 0 || self.max_depth > transaction::MAX_DOCUMENT_DEPTH {
            return Err(limit(
                crate::document::LimitKind::XmlDepth,
                self.max_depth,
                transaction::MAX_DOCUMENT_DEPTH,
            ));
        }
        if self.max_nodes == 0 || self.max_nodes > transaction::MAX_DOCUMENT_NODES {
            return Err(limit(
                crate::document::LimitKind::XmlElements,
                self.max_nodes,
                transaction::MAX_DOCUMENT_NODES,
            ));
        }
        if self.max_attributes == 0 || self.max_attributes > transaction::MAX_DOCUMENT_ATTRIBUTES {
            return Err(limit(
                crate::document::LimitKind::XmlAttributes,
                self.max_attributes,
                transaction::MAX_DOCUMENT_ATTRIBUTES,
            ));
        }
        if self.max_attribute_bytes == 0
            || self.max_attribute_bytes > transaction::MAX_DOCUMENT_ATTRIBUTE_BYTES
        {
            return Err(limit(
                crate::document::LimitKind::XmlAttributeBytes,
                self.max_attribute_bytes,
                transaction::MAX_DOCUMENT_ATTRIBUTE_BYTES,
            ));
        }
        if self.max_operations > transaction::MAX_OPERATIONS {
            return Err(limit(
                crate::document::LimitKind::Operations,
                self.max_operations,
                transaction::MAX_OPERATIONS,
            ));
        }
        if self.max_replacement_text_bytes > transaction::MAX_REPLACEMENT_TEXT_BYTES {
            return Err(limit(
                crate::document::LimitKind::ReplacementTextBytes,
                self.max_replacement_text_bytes,
                transaction::MAX_REPLACEMENT_TEXT_BYTES,
            ));
        }
        Ok(())
    }

    pub(crate) const fn allows(self, requested: Self) -> bool {
        requested.max_xml_bytes <= self.max_xml_bytes
            && requested.max_depth <= self.max_depth
            && requested.max_nodes <= self.max_nodes
            && requested.max_attributes <= self.max_attributes
            && requested.max_attribute_bytes <= self.max_attribute_bytes
            && requested.max_operations <= self.max_operations
            && requested.max_replacement_text_bytes <= self.max_replacement_text_bytes
    }

    pub(crate) const fn first_widened(
        self,
        requested: Self,
    ) -> Option<(crate::document::LimitKind, usize, usize)> {
        if self.allows(requested) {
            return None;
        }
        if requested.max_xml_bytes > self.max_xml_bytes {
            return Some((
                crate::document::LimitKind::XmlBytes,
                requested.max_xml_bytes,
                self.max_xml_bytes,
            ));
        }
        if requested.max_depth > self.max_depth {
            return Some((
                crate::document::LimitKind::XmlDepth,
                requested.max_depth,
                self.max_depth,
            ));
        }
        if requested.max_nodes > self.max_nodes {
            return Some((
                crate::document::LimitKind::XmlElements,
                requested.max_nodes,
                self.max_nodes,
            ));
        }
        if requested.max_attributes > self.max_attributes {
            return Some((
                crate::document::LimitKind::XmlAttributes,
                requested.max_attributes,
                self.max_attributes,
            ));
        }
        if requested.max_attribute_bytes > self.max_attribute_bytes {
            return Some((
                crate::document::LimitKind::XmlAttributeBytes,
                requested.max_attribute_bytes,
                self.max_attribute_bytes,
            ));
        }
        if requested.max_operations > self.max_operations {
            return Some((
                crate::document::LimitKind::Operations,
                requested.max_operations,
                self.max_operations,
            ));
        }
        if requested.max_replacement_text_bytes > self.max_replacement_text_bytes {
            return Some((
                crate::document::LimitKind::ReplacementTextBytes,
                requested.max_replacement_text_bytes,
                self.max_replacement_text_bytes,
            ));
        }
        None
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            max_xml_bytes: transaction::MAX_DOCUMENT_XML_BYTES,
            max_depth: transaction::MAX_DOCUMENT_DEPTH,
            max_nodes: transaction::MAX_DOCUMENT_NODES,
            max_attributes: transaction::MAX_DOCUMENT_ATTRIBUTES,
            max_attribute_bytes: transaction::MAX_DOCUMENT_ATTRIBUTE_BYTES,
            max_operations: transaction::MAX_OPERATIONS,
            max_replacement_text_bytes: transaction::MAX_REPLACEMENT_TEXT_BYTES,
        }
    }
}

/// A mutable, package-neutral WordprocessingML codec tree.
pub use transaction::Edit as Tree;

/// Find a paragraph block by its native Word paragraph identifier.
pub fn by_para_id(
    blocks: &super::Blocks,
    id: crate::paragraph::Id,
) -> crate::Result<Option<super::IndexedBlock>> {
    blocks.by_para_id(id)
}

/// Read a paragraph's native extension identifiers from its start tag.
pub fn paragraph_ids(paragraph: &crate::Paragraph) -> crate::Result<crate::paragraph::Ids> {
    paragraph.extension_ids()
}

/// Read a content control's native identifier.
pub fn content_control_id(control: &super::BlockContentControl) -> crate::Result<Option<i32>> {
    control.id()
}

/// Check raw XML limits before the transaction parser retains layout state.
pub(crate) fn validate_xml(xml: &[u8], limits: Limits) -> transaction::Result<()> {
    limits.validate()?;
    if xml.len() > limits.max_xml_bytes {
        return Err(limit(
            crate::document::LimitKind::XmlBytes,
            xml.len(),
            limits.max_xml_bytes,
        ));
    }
    let mut reader = Reader::from_reader(xml);
    let mut depth = 0usize;
    let mut nodes = 0usize;
    loop {
        let event = reader
            .read_event()
            .map_err(|error| transaction::Error::Document(crate::Error::Xml(error.to_string())))?;
        match event {
            Event::Start(element) => {
                check_attributes(&element, limits)?;
                nodes = nodes.saturating_add(1);
                if nodes > limits.max_nodes {
                    return Err(limit(
                        crate::document::LimitKind::XmlElements,
                        nodes,
                        limits.max_nodes,
                    ));
                }
                depth = depth.saturating_add(1);
                if depth > limits.max_depth {
                    return Err(limit(
                        crate::document::LimitKind::XmlDepth,
                        depth,
                        limits.max_depth,
                    ));
                }
            },
            Event::Empty(element) => {
                check_attributes(&element, limits)?;
                nodes = nodes.saturating_add(1);
                if nodes > limits.max_nodes {
                    return Err(limit(
                        crate::document::LimitKind::XmlElements,
                        nodes,
                        limits.max_nodes,
                    ));
                }
                if depth.saturating_add(1) > limits.max_depth {
                    return Err(limit(
                        crate::document::LimitKind::XmlDepth,
                        depth.saturating_add(1),
                        limits.max_depth,
                    ));
                }
            },
            Event::End(_) => {
                depth = depth.checked_sub(1).ok_or_else(|| {
                    transaction::Error::Document(crate::Error::InvalidFormat(
                        "document XML nesting is unbalanced".into(),
                    ))
                })?;
            },
            Event::Eof => break,
            Event::Decl(_)
            | Event::DocType(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::PI(_)
            | Event::GeneralRef(_) => {},
        }
    }
    if depth != 0 {
        return Err(transaction::Error::Document(crate::Error::InvalidFormat(
            "document XML nesting is unbalanced".into(),
        )));
    }
    Ok(())
}

fn check_attributes(
    element: &quick_xml::events::BytesStart<'_>,
    limits: Limits,
) -> transaction::Result<()> {
    let mut count = 0usize;
    for attribute in element.checked_attributes() {
        attribute
            .map_err(|error| transaction::Error::Document(crate::Error::Xml(error.to_string())))?;
        count = count.saturating_add(1);
        if count > limits.max_attributes {
            return Err(limit(
                crate::document::LimitKind::XmlAttributes,
                count,
                limits.max_attributes,
            ));
        }
    }
    let bytes = element.attributes_raw().len();
    if bytes > limits.max_attribute_bytes {
        return Err(limit(
            crate::document::LimitKind::XmlAttributeBytes,
            bytes,
            limits.max_attribute_bytes,
        ));
    }
    Ok(())
}

fn limit(kind: crate::document::LimitKind, observed: usize, maximum: usize) -> transaction::Error {
    transaction::Error::Limit {
        kind,
        observed,
        maximum,
    }
}
