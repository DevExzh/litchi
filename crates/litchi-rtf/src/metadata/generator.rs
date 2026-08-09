//! Inert producer provenance from the RTF `generator` destination.

use crate::{RtfError, RtfResult};
use std::borrow::Cow;

pub(crate) const MAX_GENERATOR_BYTES: usize = 65_536;

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DocumentGenerator<'a> {
    pub value: Cow<'a, str>,
}

impl<'a> DocumentGenerator<'a> {
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn new(value: Cow<'a, str>) -> RtfResult<Self> {
        let generator = Self { value };
        generator.validate()?;
        Ok(generator)
    }

    pub(crate) fn validate(&self) -> RtfResult<()> {
        if self.value.trim().is_empty() {
            return Err(RtfError::MalformedDocument(
                "RTF generator value cannot be empty".to_string(),
            ));
        }
        if self.value.len() > MAX_GENERATOR_BYTES {
            return Err(RtfError::MalformedDocument(
                "RTF generator value exceeds the safety limit".to_string(),
            ));
        }
        if self.value.contains(['\0', '\r', '\n']) {
            return Err(RtfError::MalformedDocument(
                "RTF generator value contains a forbidden control character".to_string(),
            ));
        }
        Ok(())
    }

    pub(crate) fn into_owned(self) -> DocumentGenerator<'static> {
        DocumentGenerator {
            value: Cow::Owned(self.value.into_owned()),
        }
    }
}
