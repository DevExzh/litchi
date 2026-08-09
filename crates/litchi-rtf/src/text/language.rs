//! RTF document and character language metadata.

use crate::{RtfError, RtfResult};

/// An RTF language identifier from the specification's standard language table.
///
/// Unknown identifiers remain representable so long as they fit the unsigned
/// 16-bit range used by RTF and Windows language identifiers.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct LanguageId(u16);

impl LanguageId {
    /// The RTF undefined/no-language identifier used with `\noproof`.
    pub const UNDEFINED: Self = Self(1024);
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn new(value: u32) -> RtfResult<Self> {
        u16::try_from(value).map(Self).map_err(|_err| {
            RtfError::MalformedDocument("RTF language ID must be in 0..=65535".to_string())
        })
    }
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn from_rtf(value: i32) -> RtfResult<Self> {
        let raw = u32::try_from(value).map_err(|_err| {
            RtfError::MalformedDocument("RTF language ID must be in 0..=65535".to_string())
        })?;
        Self::new(raw)
    }

    #[must_use]
    pub const fn value(self) -> u16 {
        self.0
    }

    #[must_use]
    pub const fn rtf_value(self) -> i32 {
        self.0 as i32
    }
}

/// Default languages declared in the RTF header.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct DocumentLanguageDefaults {
    pub primary: Option<LanguageId>,
    pub east_asian: Option<LanguageId>,
    pub complex_script: Option<LanguageId>,
}

impl DocumentLanguageDefaults {
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    #[allow(
        clippy::unused_self,
        clippy::unnecessary_wraps,
        reason = "keeps the metadata validate() call symmetry; language defaults currently have no fallible checks"
    )]
    pub(crate) fn validate(&self) -> RtfResult<()> {
        Ok(())
    }
}
