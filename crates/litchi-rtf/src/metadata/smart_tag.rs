//! Inert RTF SmartTag/factoid markup.
//!
//! RTF 1.9.1 represents a SmartTag as a starred `\xmlopen` destination with
//! required `\xmlnsN` and `\factoidname` controls, zero or more `\xmlattr`
//! groups whose required namespace control is `\xmlattrnsN`, followed by a
//! starred `\xmlclose` destination.  Attribute names and values may use the
//! direct `#PCDATA` controls from the normative grammar or the nested groups
//! emitted by the canonical writer.  The values are producer-defined
//! metadata.  Litchi stores them as bounded main-body ranges and never
//! resolves a namespace, follows a factoid, calls an action provider, or
//! executes code.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "items stay grouped by RTF SmartTag vocabulary"
)]

use crate::{RtfError, RtfResult};
use std::borrow::Cow;

pub(crate) const MAX_SMART_TAGS: usize = 65_536;
pub(crate) const MAX_SMART_TAG_DEPTH: usize = 64;
pub(crate) const MAX_SMART_TAG_NAME_BYTES: usize = 1_024;
pub(crate) const MAX_SMART_TAG_ATTRIBUTES_PER_TAG: usize = 1_024;
pub(crate) const MAX_SMART_TAG_ATTRIBUTE_NAME_BYTES: usize = 1_024;
pub(crate) const MAX_SMART_TAG_ATTRIBUTE_VALUE_BYTES: usize = 65_536;
pub(crate) const MAX_SMART_TAG_TOTAL_BYTES: usize = 16 * 1_048_576;

fn validate_name(kind: &str, value: &str, limit: usize) -> RtfResult<()> {
    if value.trim().is_empty() {
        return Err(RtfError::MalformedDocument(format!(
            "RTF SmartTag {kind} cannot be empty"
        )));
    }
    if value.len() > limit {
        return Err(RtfError::MalformedDocument(format!(
            "RTF SmartTag {kind} exceeds the safety limit"
        )));
    }
    if value.contains(['\0', '\r', '\n']) {
        return Err(RtfError::MalformedDocument(format!(
            "RTF SmartTag {kind} contains a forbidden control character"
        )));
    }
    Ok(())
}

/// One inert `\xmlattr` property attached to a SmartTag.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SmartTagAttribute<'a> {
    /// XML namespace-table reference from the required `\xmlattrnsN` control.
    pub namespace: Option<u32>,
    /// Attribute name from direct `\xmlattrname` PCDATA or a nested destination.
    pub name: Cow<'a, str>,
    /// Attribute value from direct `\xmlattrvalue` PCDATA or a nested destination.
    pub value: Cow<'a, str>,
}

impl<'a> SmartTagAttribute<'a> {
    /// Construct a validated inert SmartTag attribute.
    pub fn new(namespace: Option<u32>, name: Cow<'a, str>, value: Cow<'a, str>) -> RtfResult<Self> {
        let attribute = Self {
            namespace,
            name,
            value,
        };
        attribute.validate()?;
        Ok(attribute)
    }

    pub(crate) fn validate(&self) -> RtfResult<()> {
        let Some(namespace) = self.namespace else {
            return Err(RtfError::MalformedDocument(
                "RTF SmartTag attributes require xmlattrnsN".to_string(),
            ));
        };
        if namespace > i32::MAX as u32 {
            return Err(RtfError::MalformedDocument(
                "RTF SmartTag attribute namespace exceeds the signed RTF range".to_string(),
            ));
        }
        validate_name(
            "attribute name",
            self.name.as_ref(),
            MAX_SMART_TAG_ATTRIBUTE_NAME_BYTES,
        )?;
        if self.value.len() > MAX_SMART_TAG_ATTRIBUTE_VALUE_BYTES {
            return Err(RtfError::MalformedDocument(
                "RTF SmartTag attribute value exceeds the safety limit".to_string(),
            ));
        }
        if self.value.contains(['\0', '\r', '\n']) {
            return Err(RtfError::MalformedDocument(
                "RTF SmartTag attribute value contains a forbidden control character".to_string(),
            ));
        }
        Ok(())
    }

    pub(crate) fn into_owned(self) -> SmartTagAttribute<'static> {
        SmartTagAttribute {
            namespace: self.namespace,
            name: Cow::Owned(self.name.into_owned()),
            value: Cow::Owned(self.value.into_owned()),
        }
    }
}

/// One inert SmartTag/factoid range in the main body story.
///
/// Parsed values always carry the required non-negative `\xmlnsN` reference and
/// each attribute carries its required `\xmlattrnsN` reference.  The `Option`
/// fields preserve the existing public representation, while validation
/// rejects missing controls before a value can be written.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SmartTag<'a> {
    /// Factoid name from the nested `\factoidname` destination.
    pub name: Cow<'a, str>,
    /// XML namespace-table reference selected by the required `\xmlnsN` control.
    pub namespace: Option<u32>,
    /// Ordered `\xmlattr` values attached to the opening marker.
    pub attributes: Vec<SmartTagAttribute<'a>>,
    /// UTF-8 byte offset at which the SmartTag opens in body text.
    pub position: usize,
    /// Body text covered by the SmartTag.
    pub content: Cow<'a, str>,
}

impl<'a> SmartTag<'a> {
    /// Construct a validated inert SmartTag range.
    pub fn new(
        name: Cow<'a, str>,
        namespace: Option<u32>,
        attributes: Vec<SmartTagAttribute<'a>>,
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
        validate_name("factoid name", self.name.as_ref(), MAX_SMART_TAG_NAME_BYTES)?;
        let Some(_) = self.namespace else {
            return Err(RtfError::MalformedDocument(
                "RTF SmartTag xmlopen requires xmlnsN".to_string(),
            ));
        };
        if self.namespace.is_some_and(|id| id > i32::MAX as u32) {
            return Err(RtfError::MalformedDocument(
                "RTF SmartTag XML namespace exceeds the signed RTF range".to_string(),
            ));
        }
        if self.attributes.len() > MAX_SMART_TAG_ATTRIBUTES_PER_TAG {
            return Err(RtfError::MalformedDocument(
                "RTF SmartTag attribute count exceeds the safety limit".to_string(),
            ));
        }
        for attribute in &self.attributes {
            attribute.validate()?;
        }
        if self.content.len() > MAX_SMART_TAG_TOTAL_BYTES {
            return Err(RtfError::MalformedDocument(
                "RTF SmartTag content exceeds the safety limit".to_string(),
            ));
        }
        Ok(())
    }

    pub(crate) fn into_owned(self) -> SmartTag<'static> {
        SmartTag {
            name: Cow::Owned(self.name.into_owned()),
            namespace: self.namespace,
            attributes: self
                .attributes
                .into_iter()
                .map(SmartTagAttribute::into_owned)
                .collect(),
            position: self.position,
            content: Cow::Owned(self.content.into_owned()),
        }
    }
}
