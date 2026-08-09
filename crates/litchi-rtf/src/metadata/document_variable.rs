//! Inert RTF document-variable metadata.

use crate::{RtfError, RtfResult};
use std::borrow::Cow;

pub(crate) const MAX_DOCUMENT_VARIABLES: usize = 65_536;
pub(crate) const MAX_DOCUMENT_VARIABLE_NAME_BYTES: usize = 65_536;
pub(crate) const MAX_DOCUMENT_VARIABLE_VALUE_BYTES: usize = 4 * 1_048_576;
pub(crate) const MAX_DOCUMENT_VARIABLE_TEXT_BYTES: usize = 16 * 1_048_576;

/// One ordered `\docvar` name/value pair.
///
/// Values are inert metadata. Litchi never evaluates macros or refreshes
/// `DOCVARIABLE` fields from this value.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DocumentVariable<'a> {
    pub name: Cow<'a, str>,
    pub value: Cow<'a, str>,
}

impl<'a> DocumentVariable<'a> {
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn new(name: impl Into<Cow<'a, str>>, value: impl Into<Cow<'a, str>>) -> RtfResult<Self> {
        let variable = Self {
            name: name.into(),
            value: value.into(),
        };
        variable.validate()?;
        Ok(variable)
    }

    pub(crate) fn validate(&self) -> RtfResult<()> {
        if self.name.is_empty() {
            return Err(RtfError::MalformedDocument(
                "RTF document-variable name cannot be empty".to_string(),
            ));
        }
        if self.name.len() > MAX_DOCUMENT_VARIABLE_NAME_BYTES {
            return Err(RtfError::MalformedDocument(format!(
                "RTF document-variable name exceeds {MAX_DOCUMENT_VARIABLE_NAME_BYTES} bytes"
            )));
        }
        if self.value.len() > MAX_DOCUMENT_VARIABLE_VALUE_BYTES {
            return Err(RtfError::MalformedDocument(format!(
                "RTF document-variable value exceeds {MAX_DOCUMENT_VARIABLE_VALUE_BYTES} bytes"
            )));
        }
        Ok(())
    }

    #[must_use]
    pub fn into_owned(self) -> DocumentVariable<'static> {
        DocumentVariable {
            name: Cow::Owned(self.name.into_owned()),
            value: Cow::Owned(self.value.into_owned()),
        }
    }
}
