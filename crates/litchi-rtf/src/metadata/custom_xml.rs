//! Inert RTF custom XML markup tags.
//!
//! The RTF 1.9.1 specification (Word 2003 custom XML markup) defines the
//! `\xmlopen`, `\xmlclose`, `\xmlattrname`, and `\xmlattrvalue` destinations,
//! which wrap body text in custom XML tags. Together with the `\xmlnstbl`
//! namespace table (see [`crate::XmlNamespace`]) they let producers annotate
//! document ranges with namespaced XML element names and attributes.
//!
//! The markup is parsed and stored as passive metadata only: tag and
//! attribute names are never resolved against a schema, and no XML content is
//! fetched, validated, or executed.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "items stay grouped by RTF feature area rather than by item kind"
)]
use crate::{RtfError, RtfResult};
use std::borrow::Cow;
use std::collections::HashSet;

pub(crate) const MAX_CUSTOM_XML_TAGS: usize = 65_536;
pub(crate) const MAX_CUSTOM_XML_DEPTH: usize = 64;
pub(crate) const MAX_CUSTOM_XML_NAME_BYTES: usize = 1_024;
pub(crate) const MAX_CUSTOM_XML_ATTRIBUTES_PER_TAG: usize = 1_024;
pub(crate) const MAX_CUSTOM_XML_ATTRIBUTE_NAME_BYTES: usize = 1_024;
pub(crate) const MAX_CUSTOM_XML_ATTRIBUTE_VALUE_BYTES: usize = 65_536;
pub(crate) const MAX_CUSTOM_XML_TOTAL_BYTES: usize = 16 * 1_048_576;

fn validate_name(kind: &str, name: &str, max_bytes: usize) -> RtfResult<()> {
    if name.trim().is_empty() {
        return Err(RtfError::MalformedDocument(format!(
            "RTF custom XML {kind} cannot be empty"
        )));
    }
    if name.len() > max_bytes {
        return Err(RtfError::MalformedDocument(format!(
            "RTF custom XML {kind} exceeds the safety limit"
        )));
    }
    if name.contains(['\0', '\r', '\n']) {
        return Err(RtfError::MalformedDocument(format!(
            "RTF custom XML {kind} contains a forbidden control character"
        )));
    }
    Ok(())
}

/// One inert attribute of a custom XML markup tag.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CustomXmlAttribute<'a> {
    /// Attribute name from the `\xmlattrname` destination.
    pub name: Cow<'a, str>,
    /// Attribute value from the `\xmlattrvalue` destination.
    pub value: Cow<'a, str>,
}

impl<'a> CustomXmlAttribute<'a> {
    /// Create a validated custom XML attribute.
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn new(name: Cow<'a, str>, value: Cow<'a, str>) -> RtfResult<Self> {
        let attribute = Self { name, value };
        attribute.validate()?;
        Ok(attribute)
    }

    pub(crate) fn validate(&self) -> RtfResult<()> {
        validate_name(
            "attribute name",
            self.name.as_ref(),
            MAX_CUSTOM_XML_ATTRIBUTE_NAME_BYTES,
        )?;
        if self.value.len() > MAX_CUSTOM_XML_ATTRIBUTE_VALUE_BYTES {
            return Err(RtfError::MalformedDocument(
                "RTF custom XML attribute value exceeds the safety limit".to_string(),
            ));
        }
        if self.value.contains(['\0', '\r', '\n']) {
            return Err(RtfError::MalformedDocument(
                "RTF custom XML attribute value contains a forbidden control character".to_string(),
            ));
        }
        Ok(())
    }

    pub(crate) fn into_owned(self) -> CustomXmlAttribute<'static> {
        CustomXmlAttribute {
            name: Cow::Owned(self.name.into_owned()),
            value: Cow::Owned(self.value.into_owned()),
        }
    }
}

/// One inert custom XML markup tag spanning a range of body text.
///
/// The tag opens at `position` (a UTF-8 byte offset into the document body
/// text) and covers `content`, mirroring how [`crate::Bookmark`] ranges are
/// recorded. Tags parsed from a document are ordered by source order and are
/// properly nested: a tag's range never partially overlaps another tag's
/// range.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CustomXmlTag<'a> {
    /// Element name from the `\xmlopen`/`\xmlclose` destination text.
    pub name: Cow<'a, str>,
    /// Optional `\xmlnsN` namespace-table reference selected for the tag.
    pub namespace: Option<u32>,
    /// Ordered inert attributes attached to the tag.
    pub attributes: Vec<CustomXmlAttribute<'a>>,
    /// UTF-8 byte offset in the document body text where the tag opens.
    pub position: usize,
    /// Body text covered by the tag.
    pub content: Cow<'a, str>,
}

impl<'a> CustomXmlTag<'a> {
    /// Create a validated custom XML markup tag.
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn new(
        name: Cow<'a, str>,
        namespace: Option<u32>,
        attributes: Vec<CustomXmlAttribute<'a>>,
        position: usize,
        content: Cow<'a, str>,
    ) -> RtfResult<Self> {
        let tag = Self {
            name,
            namespace,
            attributes,
            position,
            content,
        };
        tag.validate()?;
        Ok(tag)
    }

    pub(crate) fn validate(&self) -> RtfResult<()> {
        validate_name("tag name", self.name.as_ref(), MAX_CUSTOM_XML_NAME_BYTES)?;
        if let Some(id) = self.namespace
            && (id == 0 || id > i32::MAX as u32)
        {
            return Err(RtfError::MalformedDocument(
                "RTF custom XML namespace references must be in 1..=2147483647".to_string(),
            ));
        }
        if self.attributes.len() > MAX_CUSTOM_XML_ATTRIBUTES_PER_TAG {
            return Err(RtfError::MalformedDocument(
                "RTF custom XML attribute count exceeds the safety limit".to_string(),
            ));
        }
        let mut names = HashSet::new();
        crate::error::try_reserve_set(
            &mut names,
            self.attributes.len(),
            "custom XML attribute names",
        )?;
        for attribute in &self.attributes {
            attribute.validate()?;
            if !names.insert(attribute.name.as_ref()) {
                return Err(RtfError::MalformedDocument(
                    "RTF custom XML attribute names must be unique within a tag".to_string(),
                ));
            }
        }
        Ok(())
    }

    pub(crate) fn into_owned(self) -> CustomXmlTag<'static> {
        CustomXmlTag {
            name: Cow::Owned(self.name.into_owned()),
            namespace: self.namespace,
            attributes: self
                .attributes
                .into_iter()
                .map(CustomXmlAttribute::into_owned)
                .collect(),
            position: self.position,
            content: Cow::Owned(self.content.into_owned()),
        }
    }
}
