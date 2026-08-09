//! Inert RTF document-theme payloads.

use crate::{RtfError, RtfResult};
use std::borrow::Cow;

pub(crate) const MAX_THEME_DATA_BYTES: usize = 16 * 1_048_576;
pub(crate) const MAX_COLOR_SCHEME_MAPPING_BYTES: usize = 1_048_576;

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DocumentTheme<'a> {
    pub data: Cow<'a, [u8]>,
    pub color_scheme_mapping: Option<Cow<'a, [u8]>>,
}

impl<'a> DocumentTheme<'a> {
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn new(
        data: Cow<'a, [u8]>,
        color_scheme_mapping: Option<Cow<'a, [u8]>>,
    ) -> RtfResult<Self> {
        let theme = Self {
            data,
            color_scheme_mapping,
        };
        theme.validate()?;
        Ok(theme)
    }

    pub(crate) fn validate(&self) -> RtfResult<()> {
        if self.data.is_empty() {
            return Err(RtfError::MalformedDocument(
                "RTF theme data cannot be empty".to_string(),
            ));
        }
        if self.data.len() > MAX_THEME_DATA_BYTES {
            return Err(RtfError::MalformedDocument(
                "RTF theme data exceeds the safety limit".to_string(),
            ));
        }
        if self
            .color_scheme_mapping
            .as_ref()
            .is_some_and(|mapping| mapping.is_empty())
        {
            return Err(RtfError::MalformedDocument(
                "RTF color-scheme mapping cannot be empty".to_string(),
            ));
        }
        if self
            .color_scheme_mapping
            .as_ref()
            .is_some_and(|mapping| mapping.len() > MAX_COLOR_SCHEME_MAPPING_BYTES)
        {
            return Err(RtfError::MalformedDocument(
                "RTF color-scheme mapping exceeds the safety limit".to_string(),
            ));
        }
        Ok(())
    }

    pub(crate) fn into_owned(self) -> DocumentTheme<'static> {
        DocumentTheme {
            data: Cow::Owned(self.data.into_owned()),
            color_scheme_mapping: self
                .color_scheme_mapping
                .map(|mapping| Cow::Owned(mapping.into_owned())),
        }
    }
}
