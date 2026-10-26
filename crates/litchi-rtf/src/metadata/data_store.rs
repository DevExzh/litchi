//! Inert RTF XML data-store payload preservation.

use std::borrow::Cow;

pub(crate) const MAX_DATA_STORE_BYTES: usize = 32 * 1_048_576;

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DocumentDataStore<'a> {
    pub data: Cow<'a, [u8]>,
}

impl<'a> DocumentDataStore<'a> {
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn new(data: Cow<'a, [u8]>) -> crate::Result<Self> {
        let store = Self { data };
        store.validate()?;
        Ok(store)
    }

    pub(crate) fn validate(&self) -> crate::Result<()> {
        if self.data.is_empty() {
            return Err(crate::Error::MalformedDocument(
                "RTF data-store payload cannot be empty".to_string(),
            ));
        }
        if self.data.len() > MAX_DATA_STORE_BYTES {
            return Err(crate::Error::MalformedDocument(
                "RTF data-store payload exceeds the safety limit".to_string(),
            ));
        }
        Ok(())
    }

    pub(crate) fn into_owned(self) -> DocumentDataStore<'static> {
        DocumentDataStore {
            data: Cow::Owned(self.data.into_owned()),
        }
    }
}
