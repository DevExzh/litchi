use std::borrow::Cow;

/// Maximum encoded UTF-8 size accepted for a document-window caption.
pub const MAX_WINDOW_CAPTION_BYTES: usize = 65_536;

/// Passive caption metadata for the document window.
///
/// This value is inert: parsing or setting it never creates, locates, or
/// modifies an application window.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DocumentWindowCaption<'a> {
    pub text: Cow<'a, str>,
}

impl<'a> DocumentWindowCaption<'a> {
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn new(text: Cow<'a, str>) -> crate::Result<Self> {
        let caption = Self { text };
        caption.validate()?;
        Ok(caption)
    }
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn validate(&self) -> crate::Result<()> {
        if self.text.is_empty() {
            return Err(crate::Error::MalformedDocument(
                "RTF window caption must not be empty".to_string(),
            ));
        }
        if self.text.len() > MAX_WINDOW_CAPTION_BYTES {
            return Err(crate::Error::MalformedDocument(
                "RTF window caption exceeds the resource limit".to_string(),
            ));
        }
        if self
            .text
            .chars()
            .any(|character| matches!(character, '\0' | '\r' | '\n'))
        {
            return Err(crate::Error::MalformedDocument(
                "RTF window caption contains a forbidden character".to_string(),
            ));
        }
        Ok(())
    }

    pub(crate) fn into_owned(self) -> DocumentWindowCaption<'static> {
        DocumentWindowCaption {
            text: Cow::Owned(self.text.into_owned()),
        }
    }
}
