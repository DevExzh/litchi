//! Inert RTF XML namespace-table entries.

use crate::{RtfError, RtfResult};
use std::borrow::Cow;

pub(crate) const MAX_XML_NAMESPACES: usize = 65_536;
pub(crate) const MAX_XML_NAMESPACE_BYTES: usize = 65_536;
pub(crate) const MAX_XML_NAMESPACE_TOTAL_BYTES: usize = 16 * 1_048_576;

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XmlNamespace<'a> {
    pub id: u32,
    pub namespace: Cow<'a, str>,
}

impl<'a> XmlNamespace<'a> {
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn new(id: u32, namespace: Cow<'a, str>) -> RtfResult<Self> {
        let entry = Self { id, namespace };
        entry.validate()?;
        Ok(entry)
    }

    pub(crate) fn validate(&self) -> RtfResult<()> {
        if self.id == 0 || self.id > i32::MAX as u32 {
            return Err(RtfError::MalformedDocument(
                "RTF XML namespace IDs must be in 1..=2147483647".to_string(),
            ));
        }
        if self.namespace.trim().is_empty() {
            return Err(RtfError::MalformedDocument(
                "RTF XML namespace identifier cannot be empty".to_string(),
            ));
        }
        if self.namespace.len() > MAX_XML_NAMESPACE_BYTES {
            return Err(RtfError::MalformedDocument(
                "RTF XML namespace identifier exceeds the safety limit".to_string(),
            ));
        }
        if self.namespace.contains(['\0', '\r', '\n']) {
            return Err(RtfError::MalformedDocument(
                "RTF XML namespace identifier contains a forbidden control character".to_string(),
            ));
        }
        Ok(())
    }

    pub(crate) fn into_owned(self) -> XmlNamespace<'static> {
        XmlNamespace {
            id: self.id,
            namespace: Cow::Owned(self.namespace.into_owned()),
        }
    }
}
